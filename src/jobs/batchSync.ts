import 'dotenv/config';
import pLimit from 'p-limit';
import { prisma } from '../store/db.js';
import { getAuthedClientFromStoredToken, getMissingGmailScopes } from '../auth/google.js';
import { ingestInboxWithClient } from '../gmail/fetch.js';

const ACTIVE_WINDOW_DAYS = Number(process.env.BATCH_SYNC_ACTIVE_DAYS ?? 30);
const MIN_INTERVAL_MINUTES = Number(process.env.BATCH_SYNC_MIN_INTERVAL_MINUTES ?? 30);
const MAX_USERS = Number(process.env.BATCH_SYNC_MAX_USERS ?? 200);
const CONCURRENCY = Number(process.env.BATCH_SYNC_CONCURRENCY ?? 3);
const MAX_PAGES = Number(process.env.BATCH_SYNC_MAX_PAGES ?? 2);
const MIN_NEW = Number(process.env.BATCH_SYNC_MIN_NEW ?? 20);

type SyncUser = {
  id: string;
  email: string;
  googleToken: {
    refreshToken: string;
    accessToken: string | null;
    scope: string | null;
    tokenType: string | null;
    expiryDate: Date | null;
  } | null;
};

async function runBatchSync() {
  const now = new Date();
  const activeSince = new Date(now.getTime() - ACTIVE_WINDOW_DAYS * 24 * 60 * 60 * 1000);
  const minSyncTime = new Date(now.getTime() - MIN_INTERVAL_MINUTES * 60 * 1000);

  console.log('[batch-sync] start', {
    activeWindowDays: ACTIVE_WINDOW_DAYS,
    minIntervalMinutes: MIN_INTERVAL_MINUTES,
    maxUsers: MAX_USERS,
    concurrency: CONCURRENCY,
    maxPages: MAX_PAGES,
    minNew: MIN_NEW
  });

  const users = await prisma.user.findMany({
    where: {
      lastActiveAt: { gte: activeSince },
      googleToken: { isNot: null },
      OR: [
        { lastBatchSyncAt: null },
        { lastBatchSyncAt: { lt: minSyncTime } }
      ]
    },
    include: { googleToken: true },
    take: MAX_USERS,
    orderBy: [{ lastBatchSyncAt: 'asc' }]
  }) as SyncUser[];

  if (!users.length) {
    console.log('[batch-sync] no eligible users');
    return;
  }

  const limit = pLimit(CONCURRENCY);
  let synced = 0;
  let skipped = 0;
  let failed = 0;

  await Promise.allSettled(
    users.map(user => limit(async () => {
      const result = await syncUser(user);
      if (result === 'ok') synced += 1;
      if (result === 'skipped') skipped += 1;
      if (result === 'error') failed += 1;
    }))
  );

  console.log('[batch-sync] done', { attempted: users.length, synced, skipped, failed });
}

async function syncUser(user: SyncUser): Promise<'ok' | 'skipped' | 'error'> {
  const token = user.googleToken;
  if (!token?.refreshToken) {
    await markSyncResult(user.id, 'missing_refresh', 'Missing refresh token');
    console.warn('[batch-sync] missing refresh token', { userId: user.id });
    return 'skipped';
  }

  const missingScopes = getMissingGmailScopes({ scope: token.scope ?? '' });
  if (missingScopes.length) {
    await markSyncResult(user.id, 'missing_scopes', `Missing scopes: ${missingScopes.join(', ')}`);
    console.warn('[batch-sync] missing scopes', { userId: user.id, missingScopes });
    return 'skipped';
  }

  try {
    const auth = getAuthedClientFromStoredToken(user.id, {
      refreshToken: token.refreshToken,
      accessToken: token.accessToken,
      scope: token.scope,
      tokenType: token.tokenType,
      expiryDate: token.expiryDate
    });
    const result = await ingestInboxWithClient(auth, user.id, { maxPages: MAX_PAGES, minNew: MIN_NEW });
    await markSyncResult(user.id, 'ok', null);
    console.log('[batch-sync] synced', { userId: user.id, created: result.created, hasMore: result.hasMore });
    return 'ok';
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    await markSyncResult(user.id, 'error', message);
    console.warn('[batch-sync] sync failed', { userId: user.id, error: message });
    return 'error';
  }
}

async function markSyncResult(userId: string, status: string, error: string | null) {
  await prisma.user.update({
    where: { id: userId },
    data: {
      lastBatchSyncAt: new Date(),
      lastBatchSyncStatus: status,
      lastBatchSyncError: error
    }
  });
}

runBatchSync()
  .catch(err => {
    console.error('[batch-sync] fatal', err);
    process.exitCode = 1;
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
