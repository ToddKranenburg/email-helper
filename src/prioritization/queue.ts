import { prisma } from '../store/db.js';
import { runPrioritizationBatch } from './worker.js';

const DEBOUNCE_MS = Number(process.env.PRIORITY_DEBOUNCE_MS ?? 90_000);
const FOLLOWUP_MS = Number(process.env.PRIORITY_FOLLOWUP_MS ?? 30_000);

type PendingBatch = {
  threadIds: Set<string>;
  trigger: string;
  mode: 'first' | 'normal';
  timer: NodeJS.Timeout | null;
};

const pending = new Map<string, PendingBatch>();
const running = new Set<string>();
const followUps = new Map<string, NodeJS.Timeout>();

function enqueueWithDelay(
  userId: string,
  threadIds: string[],
  trigger: string,
  mode: 'first' | 'normal',
  delayMs: number,
  allowEmpty: boolean = false
) {
  if (!threadIds.length && !allowEmpty) return;
  const entry = pending.get(userId) || { threadIds: new Set<string>(), trigger, mode, timer: null };
  threadIds.forEach(id => entry.threadIds.add(id));
  entry.trigger = trigger || entry.trigger;
  entry.mode = mode;
  if (!entry.timer) {
    entry.timer = setTimeout(() => {
      void flushBatch(userId);
    }, delayMs);
  }
  pending.set(userId, entry);
}

export function enqueuePrioritization(userId: string, threadIds: string[], trigger: string) {
  enqueueWithDelay(userId, threadIds, trigger, 'normal', DEBOUNCE_MS);
}

export function enqueueDeferredPrioritization(userId: string, trigger: string, delayMs: number = FOLLOWUP_MS) {
  enqueueWithDelay(userId, [], trigger, 'normal', delayMs, true);
}

export function enqueuePrioritizationNow(
  userId: string,
  threadIds: string[],
  trigger: string,
  mode: 'first' | 'normal' = 'normal'
) {
  if (!threadIds.length) return;
  const entry = pending.get(userId) || { threadIds: new Set<string>(), trigger, mode, timer: null };
  threadIds.forEach(id => entry.threadIds.add(id));
  entry.trigger = trigger || entry.trigger;
  entry.mode = mode;
  pending.set(userId, entry);
  void flushBatch(userId);
}

async function flushBatch(userId: string) {
  const entry = pending.get(userId);
  if (!entry) return;
  if (entry.timer) clearTimeout(entry.timer);
  pending.delete(userId);

  if (running.has(userId)) {
    enqueueWithDelay(userId, Array.from(entry.threadIds), entry.trigger, entry.mode, DEBOUNCE_MS);
    return;
  }

  running.add(userId);
  try {
    await runPrioritizationBatch(userId, Array.from(entry.threadIds), entry.trigger, { mode: entry.mode });
  } catch (err) {
    console.error('[priority] batch error', { userId, error: err instanceof Error ? err.message : err });
  } finally {
    running.delete(userId);
  }

  const remaining = await prisma.deferredPrioritization.count({ where: { userId } });
  if (remaining > 0 && !followUps.has(userId)) {
    const timer = setTimeout(() => {
      followUps.delete(userId);
      enqueueDeferredPrioritization(userId, 'history_delta', 0);
    }, FOLLOWUP_MS);
    followUps.set(userId, timer);
  }
}
