import pLimit from 'p-limit';
import { GaxiosError } from 'gaxios';
import type { gmail_v1 } from 'googleapis';
import type { OAuth2Client } from 'google-auth-library';
import { gmailClient, INBOX_QUERY } from './client.js';
import { extractUnsubscribeMetadata, type UnsubscribeMetadata } from './unsubscribe.js';
import { prisma } from '../store/db.js';
import { enqueuePrioritization } from '../prioritization/queue.js';

const INITIAL_SYNC_DAYS = Number(process.env.INITIAL_SYNC_DAYS ?? 30);
const INITIAL_SYNC_MAX_THREADS = Number(process.env.INITIAL_SYNC_MAX_THREADS ?? 1000);
const METADATA_CONCURRENCY = Number(process.env.SYNC_METADATA_CONCURRENCY ?? 6);
const HISTORY_PAGE_LIMIT = Number(process.env.SYNC_HISTORY_PAGE_LIMIT ?? 500);

const METADATA_HEADERS = [
  'Subject',
  'From',
  'To',
  'Cc',
  'Date',
  'List-Unsubscribe',
  'List-Unsubscribe-Post',
  'List-Id',
  'Precedence'
];

export type SyncOutcome = {
  mode: 'initial' | 'history';
  fetched: number;
  updated: number;
  removed: number;
  affectedThreadIds: string[];
  historyCursor: string | null;
};

type ThreadMetadata = {
  threadId: string;
  subject: string | null;
  participants: string[];
  lastMessageDate: Date;
  lastMessageId: string | null;
  fromName: string | null;
  fromEmail: string | null;
  unreadCount: number;
  snippet: string | null;
  gmailLabelIds: string[];
  inPrimaryInbox: boolean;
  lastGmailHistoryIdSeen: string | null;
  contentVersion: string | null;
  unsubscribe: UnsubscribeMetadata | null;
};

function buildInitialQuery() {
  if (INITIAL_SYNC_DAYS > 0) {
    return `${INBOX_QUERY} newer_than:${INITIAL_SYNC_DAYS}d`;
  }
  return INBOX_QUERY;
}

function parseFrom(headerValue: string | null | undefined): { name?: string; email?: string } {
  if (!headerValue) return {};
  const emailMatch = headerValue.match(/<([^>]+)>/);
  const email = emailMatch ? emailMatch[1].trim() : (/@/.test(headerValue) ? headerValue.trim() : undefined);
  let name: string | undefined = undefined;
  if (emailMatch) {
    name = headerValue.replace(emailMatch[0], '').trim().replace(/^"|"$/g, '');
  } else if (email) {
    name = undefined;
  } else {
    name = headerValue.trim();
  }
  return { name, email };
}

function headerValue(headers: gmail_v1.Schema$MessagePartHeader[] | undefined, name: string) {
  if (!headers?.length) return null;
  const match = headers.find(h => h?.name === name);
  return match?.value ?? null;
}

function collectParticipants(messages: gmail_v1.Schema$Message[]): string[] {
  const participants = new Set<string>();
  for (const msg of messages) {
    const headers = msg.payload?.headers || [];
    for (const name of ['From', 'To', 'Cc']) {
      const raw = headerValue(headers, name);
      if (!raw) continue;
      raw.split(',')
        .map(item => item.trim())
        .filter(Boolean)
        .forEach(value => participants.add(value));
    }
  }
  return Array.from(participants);
}

function extractThreadMetadata(thread: gmail_v1.Schema$Thread): ThreadMetadata | null {
  const messages = (thread.messages || []).filter(
    (msg: gmail_v1.Schema$Message | null | undefined): msg is gmail_v1.Schema$Message => Boolean(msg)
  );
  if (!messages.length || !thread.id) return null;
  const sorted = messages.slice().sort((a, b) => {
    const aDate = a.internalDate ? Number(a.internalDate) : 0;
    const bDate = b.internalDate ? Number(b.internalDate) : 0;
    return aDate - bDate;
  });
  const latest = sorted[sorted.length - 1];
  const headers = latest.payload?.headers || [];
  const subject = headerValue(headers, 'Subject');
  const fromRaw = headerValue(headers, 'From');
  const { name: fromName, email: fromEmail } = parseFrom(fromRaw);
  const lastMessageDate = new Date(latest.internalDate ? Number(latest.internalDate) : Date.now());
  const lastMessageId = latest.id ?? null;
  const participants = collectParticipants(sorted);
  const unreadCount = sorted.filter(msg => msg.labelIds?.includes('UNREAD')).length;
  const labelIds = latest.labelIds || [];
  const categoryLabels = labelIds.filter(label => label.startsWith('CATEGORY_'));
  const isPrimaryCategory = categoryLabels.includes('CATEGORY_PRIMARY') || categoryLabels.includes('CATEGORY_PERSONAL');
  const isNonPrimaryCategory = categoryLabels.some(label => (
    label === 'CATEGORY_PROMOTIONS'
    || label === 'CATEGORY_SOCIAL'
    || label === 'CATEGORY_FORUMS'
    || label === 'CATEGORY_UPDATES'
  ));
  let inPrimaryInbox = labelIds.includes('INBOX') && !labelIds.includes('CHAT') && !labelIds.includes('SPAM') && !labelIds.includes('TRASH');
  if (categoryLabels.length) {
    inPrimaryInbox = inPrimaryInbox && (isPrimaryCategory || !isNonPrimaryCategory);
  }
  const snippet = thread.snippet ?? null;
  const contentVersion = buildContentVersion(lastMessageDate, lastMessageId);
  const unsubscribe = extractUnsubscribeMetadata(headers);

  return {
    threadId: thread.id,
    subject: subject || null,
    participants,
    lastMessageDate,
    lastMessageId,
    fromName: fromName ?? null,
    fromEmail: fromEmail ?? null,
    unreadCount,
    snippet,
    gmailLabelIds: labelIds,
    inPrimaryInbox,
    lastGmailHistoryIdSeen: thread.historyId ?? null,
    contentVersion,
    unsubscribe
  };
}

function buildContentVersion(lastMessageDate: Date | null, lastMessageId: string | null) {
  if (!lastMessageDate) return null;
  const stamp = lastMessageDate.toISOString();
  return lastMessageId ? `${stamp}:${lastMessageId}` : stamp;
}

async function upsertThreadIndex(userId: string, metadata: ThreadMetadata) {
  const participantsJson = metadata.participants.length ? JSON.stringify(metadata.participants) : null;
  return prisma.threadIndex.upsert({
    where: { threadId_userId: { threadId: metadata.threadId, userId } },
    update: {
      subject: metadata.subject ?? undefined,
      participants: participantsJson,
      lastMessageDate: metadata.lastMessageDate,
      lastMessageId: metadata.lastMessageId,
      fromName: metadata.fromName,
      fromEmail: metadata.fromEmail,
      inPrimaryInbox: metadata.inPrimaryInbox,
      unreadCount: metadata.unreadCount,
      snippet: metadata.snippet,
      gmailLabelIds: metadata.gmailLabelIds,
      lastGmailHistoryIdSeen: metadata.lastGmailHistoryIdSeen,
      contentVersion: metadata.contentVersion,
      unsubscribe: metadata.unsubscribe ?? undefined
    },
    create: {
      threadId: metadata.threadId,
      userId,
      subject: metadata.subject ?? undefined,
      participants: participantsJson,
      lastMessageDate: metadata.lastMessageDate,
      lastMessageId: metadata.lastMessageId,
      fromName: metadata.fromName,
      fromEmail: metadata.fromEmail,
      inPrimaryInbox: metadata.inPrimaryInbox,
      unreadCount: metadata.unreadCount,
      snippet: metadata.snippet,
      gmailLabelIds: metadata.gmailLabelIds,
      lastGmailHistoryIdSeen: metadata.lastGmailHistoryIdSeen,
      contentVersion: metadata.contentVersion,
      unsubscribe: metadata.unsubscribe ?? undefined
    }
  });
}

async function upsertGmailAccount(userId: string, emailAddress: string, historyCursor?: string | null) {
  const existing = await prisma.gmailAccount.findUnique({ where: { userId } });
  if (existing) {
    return prisma.gmailAccount.update({
      where: { userId },
      data: {
        emailAddress,
        ...(historyCursor ? { historyCursor } : {})
      }
    });
  }
  return prisma.gmailAccount.create({
    data: {
      userId,
      emailAddress,
      historyCursor: historyCursor ?? null
    }
  });
}

export function computeReconcileTargets(currentIds: string[], fetchedThreadIds: Set<string>) {
  return currentIds.filter(id => !fetchedThreadIds.has(id));
}

export async function reconcilePrimaryInboxSet(userId: string, fetchedThreadIds: Set<string>) {
  const ids = Array.from(fetchedThreadIds);
  const result = await prisma.threadIndex.updateMany({
    where: {
      userId,
      inPrimaryInbox: true,
      ...(ids.length ? { threadId: { notIn: ids } } : {})
    },
    data: { inPrimaryInbox: false }
  });
  return result.count;
}

export async function initialIndexBuildPrimaryInbox(
  auth: OAuth2Client,
  userId: string,
  opts: { skipPriorityEnqueue?: boolean } = {}
): Promise<SyncOutcome> {
  const gmail: gmail_v1.Gmail = gmailClient(auth);
  const query = buildInitialQuery();
  const limit = pLimit(METADATA_CONCURRENCY);

  const fetchedThreadIds = new Set<string>();
  const affectedThreadIds = new Set<string>();
  let updated = 0;
  let pageToken: string | undefined = undefined;

  while (fetchedThreadIds.size < INITIAL_SYNC_MAX_THREADS) {
    const list: gmail_v1.Schema$ListThreadsResponse = (await gmail.users.threads.list({
      userId: 'me',
      q: query,
      pageToken,
      maxResults: Math.min(100, INITIAL_SYNC_MAX_THREADS - fetchedThreadIds.size)
    })).data;

    const threads = (list.threads || []).filter(
      (thread: gmail_v1.Schema$Thread | null | undefined): thread is gmail_v1.Schema$Thread => Boolean(thread?.id)
    );

    await Promise.all(
      threads.map((thread: gmail_v1.Schema$Thread) => limit(async () => {
        if (!thread.id) return;
        fetchedThreadIds.add(thread.id);
        try {
          const full = await gmail.users.threads.get({
            userId: 'me',
            id: thread.id,
            format: 'metadata',
            metadataHeaders: METADATA_HEADERS
          });
          const meta = extractThreadMetadata(full.data);
          if (!meta) return;
          meta.inPrimaryInbox = true;
          await upsertThreadIndex(userId, meta);
          updated += 1;
          if (meta.inPrimaryInbox) {
            affectedThreadIds.add(meta.threadId);
          }
        } catch (err) {
          if (isNotFound(err)) {
            console.warn('[sync] thread not found during initial sync', { userId, threadId: thread.id });
            return;
          }
          throw err;
        }
      }))
    );

    pageToken = list.nextPageToken || undefined;
    if (!pageToken) break;
  }

  const removed = await reconcilePrimaryInboxSet(userId, fetchedThreadIds);
  const profile = await gmail.users.getProfile({ userId: 'me' });
  const historyCursor = profile.data.historyId ?? null;
  const emailAddress = profile.data.emailAddress || '';
  await upsertGmailAccount(userId, emailAddress, historyCursor);
  await prisma.gmailAccount.update({
    where: { userId },
    data: { lastInitialSyncAt: new Date(), lastSyncAt: new Date(), historyCursor }
  });

  console.log('[sync] initial index', {
    userId,
    fetched: fetchedThreadIds.size,
    updated,
    removed,
    historyCursor
  });

  if (!opts.skipPriorityEnqueue) {
    enqueuePrioritization(userId, Array.from(affectedThreadIds), 'initial_sync');
  }

  return {
    mode: 'initial',
    fetched: fetchedThreadIds.size,
    updated,
    removed,
    affectedThreadIds: Array.from(affectedThreadIds),
    historyCursor
  };
}

function isHistoryTooOld(err: unknown) {
  const gaxios = err instanceof GaxiosError ? err : null;
  return gaxios?.response?.status === 404;
}

function isNotFound(err: unknown) {
  const gaxios = err instanceof GaxiosError ? err : null;
  return gaxios?.response?.status === 404;
}

export function extractHistoryThreadIds(history: gmail_v1.Schema$History[]): Set<string> {
  const ids = new Set<string>();
  history.forEach(entry => {
    const messages = entry.messages || [];
    messages.forEach(message => {
      if (message?.threadId) ids.add(message.threadId);
    });
    const added = entry.labelsAdded || [];
    added.forEach(item => {
      if (item?.message?.threadId) ids.add(item.message.threadId);
    });
    const removed = entry.labelsRemoved || [];
    removed.forEach(item => {
      if (item?.message?.threadId) ids.add(item.message.threadId);
    });
  });
  return ids;
}

export async function incrementalSyncFromHistoryCursor(
  auth: OAuth2Client,
  userId: string,
  opts: { skipPriorityEnqueue?: boolean } = {}
): Promise<SyncOutcome> {
  const gmail: gmail_v1.Gmail = gmailClient(auth);
  const account = await prisma.gmailAccount.findUnique({ where: { userId } });
  if (!account?.historyCursor) {
    return initialIndexBuildPrimaryInbox(auth, userId);
  }

  const limit = pLimit(METADATA_CONCURRENCY);
  const affectedThreadIds = new Set<string>();
  let updated = 0;
  let removed = 0;
  let pageToken: string | undefined = undefined;
  let historyCursor: string | null = account.historyCursor;

  try {
    do {
      const response: gmail_v1.Schema$ListHistoryResponse = (await gmail.users.history.list({
        userId: 'me',
        startHistoryId: account.historyCursor,
        maxResults: HISTORY_PAGE_LIMIT,
        pageToken,
        historyTypes: ['messageAdded', 'messageDeleted', 'labelAdded', 'labelRemoved']
      })).data;

      if (response.historyId) {
        historyCursor = response.historyId;
      }

      const history = response.history || [];
      const threadIds = extractHistoryThreadIds(history);
      threadIds.forEach(id => affectedThreadIds.add(id));

      pageToken = response.nextPageToken || undefined;
    } while (pageToken);
  } catch (err) {
    if (isHistoryTooOld(err)) {
      console.warn('[sync] history cursor too old; falling back to initial sync', { userId });
      return initialIndexBuildPrimaryInbox(auth, userId);
    }
    throw err;
  }

  if (affectedThreadIds.size) {
    await Promise.all(
      Array.from(affectedThreadIds).map(threadId => limit(async () => {
        try {
          const full = await gmail.users.threads.get({
            userId: 'me',
            id: threadId,
            format: 'metadata',
            metadataHeaders: METADATA_HEADERS
          });
          const meta = extractThreadMetadata(full.data);
          if (!meta) return;
          await upsertThreadIndex(userId, meta);
          updated += 1;
          if (!meta.inPrimaryInbox) {
            removed += 1;
          }
        } catch (err) {
          if (isNotFound(err)) {
            console.warn('[sync] thread not found during history sync', { userId, threadId });
            return;
          }
          throw err;
        }
      }))
    );
  }

  await prisma.gmailAccount.update({
    where: { userId },
    data: { historyCursor, lastSyncAt: new Date() }
  });

  console.log('[sync] history delta', {
    userId,
    affected: affectedThreadIds.size,
    updated,
    removed,
    historyCursor
  });

  const primaryIds = Array.from(affectedThreadIds);
  if (!opts.skipPriorityEnqueue) {
    enqueuePrioritization(userId, primaryIds, 'history_delta');
  }

  return {
    mode: 'history',
    fetched: affectedThreadIds.size,
    updated,
    removed,
    affectedThreadIds: primaryIds,
    historyCursor
  };
}

export async function syncPrimaryInbox(
  auth: OAuth2Client,
  userId: string,
  opts: { skipPriorityEnqueue?: boolean } = {}
) {
  const account = await prisma.gmailAccount.findUnique({ where: { userId } });
  if (!account?.historyCursor) {
    return initialIndexBuildPrimaryInbox(auth, userId, opts);
  }
  const existingCount = await prisma.threadIndex.count({ where: { userId } });
  if (existingCount === 0) {
    return initialIndexBuildPrimaryInbox(auth, userId, opts);
  }
  return incrementalSyncFromHistoryCursor(auth, userId, opts);
}
