import { prisma } from '../store/db.js';
import type { Summary, ThreadIndex } from '@prisma/client';

export type PriorityEntry = {
  threadId: string;
  score: number;
  reason: string;
  reasonWeight: number;
};

export type PriorityItem = {
  thread: ThreadIndex;
  summary: Summary | null;
};

export async function loadPriorityQueue(userId: string, opts: { limit: number }) {
  const limit = Math.max(opts.limit || 0, 0);
  if (!limit) {
    return { priority: [] as PriorityEntry[], items: [] as PriorityItem[] };
  }

  const threads = await prisma.threadIndex.findMany({
    where: {
      userId,
      inPrimaryInbox: true,
      priorityScore: { not: null }
    },
    orderBy: [{ priorityScore: 'desc' }, { lastMessageDate: 'desc' }],
    take: limit
  });

  if (!threads.length) {
    return { priority: [] as PriorityEntry[], items: [] as PriorityItem[] };
  }

  const priority = threads.map(thread => ({
    threadId: thread.threadId,
    score: thread.priorityScore ?? 0,
    reason: thread.priorityReason || 'Needs attention',
    reasonWeight: 0
  }));

  const threadIds = threads.map(thread => thread.threadId);
  const summaries = await prisma.summary.findMany({
    where: { userId, threadId: { in: threadIds } }
  });
  const summaryMap = new Map(summaries.map(item => [item.threadId, item]));

  const items = threads.map(thread => ({
    thread,
    summary: summaryMap.get(thread.threadId) || null
  }));

  return { priority, items };
}
