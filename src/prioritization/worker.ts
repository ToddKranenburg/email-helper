import type { gmail_v1 } from 'googleapis';
import type { OAuth2Client } from 'google-auth-library';
import { prisma } from '../store/db.js';
import { gmailClient } from '../gmail/client.js';
import { normalizeBody } from '../gmail/normalize.js';
import { getAuthedClientFromStoredToken } from '../auth/google.js';
import { scoreThreadPriority } from '../llm/priorityScore.js';

const SCORE_VERSION = process.env.PRIORITY_SCORE_VERSION ?? 'priority-v1';

export const MAX_THREADS_PER_BATCH = 200;
export const MAX_TOTAL_CHARS_PER_BATCH = 1_200_000;
export const MAX_CHARS_PER_THREAD = 25_000;
export const MAX_MESSAGES_PER_THREAD = 10;
export const BATCH_TIME_BUDGET_SECONDS = 60;

export type PrioritizationCandidate = {
  threadId: string;
  priorityScore: number | null;
  lastMessageDate: Date | null;
  unreadCount: number;
};

export type GuardrailPlan = {
  selected: string[];
  deferred: string[];
  totalChars: number;
};

export function sortCandidatesForScoring(candidates: PrioritizationCandidate[]) {
  return candidates.slice().sort((a, b) => {
    const aNoScore = a.priorityScore == null ? 1 : 0;
    const bNoScore = b.priorityScore == null ? 1 : 0;
    if (aNoScore !== bNoScore) return bNoScore - aNoScore;
    const aDate = a.lastMessageDate?.getTime() ?? 0;
    const bDate = b.lastMessageDate?.getTime() ?? 0;
    if (aDate !== bDate) return bDate - aDate;
    const aUnread = a.unreadCount > 0 ? 1 : 0;
    const bUnread = b.unreadCount > 0 ? 1 : 0;
    return bUnread - aUnread;
  });
}

export function applyGuardrails(order: { threadId: string; contentLength: number }[]): GuardrailPlan {
  const selected: string[] = [];
  const deferred: string[] = [];
  let totalChars = 0;

  for (const item of order) {
    if (selected.length >= MAX_THREADS_PER_BATCH) {
      deferred.push(item.threadId);
      continue;
    }
    if (totalChars + item.contentLength > MAX_TOTAL_CHARS_PER_BATCH) {
      deferred.push(item.threadId);
      continue;
    }
    totalChars += item.contentLength;
    selected.push(item.threadId);
  }

  return { selected, deferred, totalChars };
}

function parseScopes(raw: string | null | undefined): string[] {
  if (!raw) return [];
  return raw.split(/\s+/).map(s => s.trim()).filter(Boolean);
}

function hasGmailScope(scopes: string[]) {
  const scopeSet = new Set(scopes);
  return scopeSet.has('https://www.googleapis.com/auth/gmail.modify')
    || scopeSet.has('https://www.googleapis.com/auth/gmail.readonly')
    || scopeSet.has('https://mail.google.com/');
}

async function fetchGmailAuth(userId: string): Promise<OAuth2Client | null> {
  const user = await prisma.user.findUnique({
    where: { id: userId },
    include: { googleToken: true }
  });
  const token = user?.googleToken;
  if (!token?.refreshToken) {
    console.warn('[priority] missing refresh token', { userId });
    return null;
  }
  const scopes = parseScopes(token.scope ?? '');
  if (!hasGmailScope(scopes)) {
    console.warn('[priority] missing gmail scope', { userId, scopes });
    return null;
  }
  return getAuthedClientFromStoredToken(userId, {
    refreshToken: token.refreshToken,
    accessToken: token.accessToken,
    scope: token.scope,
    tokenType: token.tokenType,
    expiryDate: token.expiryDate
  });
}

function buildContentVersion(lastMessageDate: Date | null, lastMessageId: string | null) {
  if (!lastMessageDate) return null;
  const stamp = lastMessageDate.toISOString();
  return lastMessageId ? `${stamp}:${lastMessageId}` : stamp;
}

function parseParticipants(raw: string | null) {
  if (!raw) return [] as string[];
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.map(value => String(value || '').trim()).filter(Boolean);
    }
  } catch {
    return [];
  }
  return [];
}

async function fetchThreadContent(
  gmail: gmail_v1.Gmail,
  threadId: string
): Promise<{ contentText: string } | null> {
  const thread = await gmail.users.threads.get({ userId: 'me', id: threadId, format: 'full' });
  const messages = (thread.data.messages || []).filter(
    (msg: gmail_v1.Schema$Message | null | undefined): msg is gmail_v1.Schema$Message => Boolean(msg)
  );
  if (!messages.length) return null;
  const sorted = messages.slice().sort((a, b) => {
    const aDate = a.internalDate ? Number(a.internalDate) : 0;
    const bDate = b.internalDate ? Number(b.internalDate) : 0;
    return aDate - bDate;
  });
  const selected = sorted.slice(-MAX_MESSAGES_PER_THREAD);
  const parts = selected
    .map(msg => normalizeBody(msg.payload))
    .filter(Boolean);
  const combined = parts.join('\n\n---\n\n');
  const truncated = combined.length > MAX_CHARS_PER_THREAD ? combined.slice(0, MAX_CHARS_PER_THREAD) : combined;
  return { contentText: truncated };
}

export async function runPrioritizationBatch(userId: string, threadIds: string[], trigger: string) {
  const auth = await fetchGmailAuth(userId);
  if (!auth) {
    console.warn('[priority] missing gmail auth; skipping batch', { userId });
    return;
  }
  const gmail = gmailClient(auth);
  const unique = new Set(threadIds);

  const deferred = await prisma.deferredPrioritization.findMany({ where: { userId } });
  deferred.forEach(item => unique.add(item.threadId));
  if (!unique.size) return;
  const uniqueList = Array.from(unique);

  const threadRecords = await prisma.threadIndex.findMany({
    where: {
      userId,
      threadId: { in: uniqueList },
      inPrimaryInbox: true
    }
  });

  const ordered = sortCandidatesForScoring(
    threadRecords.map(record => ({
      threadId: record.threadId,
      priorityScore: record.priorityScore,
      lastMessageDate: record.lastMessageDate,
      unreadCount: record.unreadCount
    }))
  );

  const batch = await prisma.prioritizationBatch.create({
    data: {
      userId,
      status: 'running',
      totalThreadsPlanned: ordered.length,
      processedThreads: 0,
      deferredThreads: 0,
      trigger
    }
  });

  const cache = await prisma.threadContentCache.findMany({
    where: { userId, threadId: { in: ordered.map(item => item.threadId) } }
  });
  const cacheMap = new Map(cache.map(item => [item.threadId, item]));
  const recordMap = new Map(threadRecords.map(record => [record.threadId, record]));

  const start = Date.now();
  let totalChars = 0;
  let processed = 0;
  const deferredThreadIds: string[] = [];
  const processedThreadIds: string[] = [];

  try {
    for (const item of ordered) {
    if (processed >= MAX_THREADS_PER_BATCH) {
      deferredThreadIds.push(item.threadId);
      continue;
    }
    if ((Date.now() - start) / 1000 > BATCH_TIME_BUDGET_SECONDS) {
      deferredThreadIds.push(item.threadId);
      continue;
    }

    const record = recordMap.get(item.threadId);
    if (!record) continue;

    const contentVersion = record.contentVersion || buildContentVersion(record.lastMessageDate, record.lastMessageId);
    if (record.priorityScore != null && record.scoreVersion === SCORE_VERSION && record.contentVersion === contentVersion && record.lastScoredAt) {
      processedThreadIds.push(record.threadId);
      continue;
    }

    const cached = cacheMap.get(record.threadId);
    let contentText = cached && cached.contentVersion === contentVersion ? cached.contentText : null;

    if (!contentText) {
      const fetched = await fetchThreadContent(gmail, record.threadId);
      if (!fetched) {
        deferredThreadIds.push(record.threadId);
        continue;
      }
      contentText = fetched.contentText;
      await prisma.threadContentCache.upsert({
        where: { userId_threadId: { userId, threadId: record.threadId } },
        update: {
          contentText,
          contentVersion: contentVersion ?? buildContentVersion(record.lastMessageDate, record.lastMessageId) ?? ''
        },
        create: {
          userId,
          threadId: record.threadId,
          contentText,
          contentVersion: contentVersion ?? buildContentVersion(record.lastMessageDate, record.lastMessageId) ?? ''
        }
      });
    }

    if (totalChars + contentText.length > MAX_TOTAL_CHARS_PER_BATCH) {
      deferredThreadIds.push(record.threadId);
      continue;
    }

      const score = await scoreThreadPriority({
        threadId: record.threadId,
        subject: record.subject || '',
        participants: parseParticipants(record.participants),
        snippet: record.snippet || '',
        content: contentText
      });

    await prisma.threadIndex.update({
      where: { threadId_userId: { threadId: record.threadId, userId } },
      data: {
        priorityScore: score.priorityScore,
        priorityReason: score.priorityReason,
        suggestedActionType: score.suggestedActionType,
        extracted: score.extracted ?? undefined,
        lastScoredAt: new Date(),
        scoreVersion: SCORE_VERSION,
        contentVersion: contentVersion ?? buildContentVersion(record.lastMessageDate, record.lastMessageId)
      }
    });

    processed += 1;
    totalChars += contentText.length;
    processedThreadIds.push(record.threadId);
  }

    if (processedThreadIds.length) {
      await prisma.deferredPrioritization.deleteMany({
        where: { userId, threadId: { in: processedThreadIds } }
      });
    }

    if (deferredThreadIds.length) {
      await Promise.all(
        deferredThreadIds.map(threadId => prisma.deferredPrioritization.upsert({
          where: { userId_threadId: { userId, threadId } },
          update: { reason: 'guardrail' },
          create: { userId, threadId, reason: 'guardrail' }
        }))
      );
    }

    await prisma.prioritizationBatch.update({
      where: { id: batch.id },
      data: {
        status: 'completed',
        processedThreads: processed,
        deferredThreads: deferredThreadIds.length,
        finishedAt: new Date()
      }
    });

    console.log('[priority] batch completed', {
      userId,
      planned: ordered.length,
      processed,
      deferred: deferredThreadIds.length,
      totalChars
    });
  } catch (err) {
    await prisma.prioritizationBatch.update({
      where: { id: batch.id },
      data: {
        status: 'failed',
        processedThreads: processed,
        deferredThreads: deferredThreadIds.length,
        finishedAt: new Date()
      }
    });
    console.error('[priority] batch failed', { userId, error: err instanceof Error ? err.message : err });
    throw err;
  }
}
