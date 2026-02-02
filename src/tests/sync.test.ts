import assert from 'node:assert/strict';
import { applyGuardrails, sortCandidatesForScoring, MAX_THREADS_PER_BATCH, MAX_TOTAL_CHARS_PER_BATCH } from '../prioritization/worker.js';
import { computeReconcileTargets, extractHistoryThreadIds } from '../gmail/sync.js';

function testReconcileTargets() {
  const current = ['t1', 't2', 't3'];
  const fetched = new Set(['t2', 't3']);
  const removed = computeReconcileTargets(current, fetched);
  assert.deepStrictEqual(removed, ['t1']);
}

function testExtractHistoryThreadIds() {
  const history = [
    {
      messages: [{ threadId: 'a' }, { threadId: 'b' }],
      labelsAdded: [{ message: { threadId: 'c' } }],
      labelsRemoved: [{ message: { threadId: 'd' } }]
    },
    {
      messages: [{ threadId: 'b' }],
      labelsAdded: [{ message: { threadId: 'e' } }]
    }
  ] as any;
  const ids = Array.from(extractHistoryThreadIds(history)).sort();
  assert.deepStrictEqual(ids, ['a', 'b', 'c', 'd', 'e']);
}

function testSortCandidates() {
  const now = new Date('2026-02-02T00:00:00Z');
  const candidates = [
    { threadId: 'scored-old', priorityScore: 50, lastMessageDate: new Date('2026-01-01T00:00:00Z'), unreadCount: 0 },
    { threadId: 'no-score-new', priorityScore: null, lastMessageDate: new Date('2026-02-01T00:00:00Z'), unreadCount: 0 },
    { threadId: 'no-score-unread', priorityScore: null, lastMessageDate: new Date('2026-01-31T00:00:00Z'), unreadCount: 2 },
    { threadId: 'scored-new', priorityScore: 10, lastMessageDate: now, unreadCount: 3 }
  ];
  const ordered = sortCandidatesForScoring(candidates).map(item => item.threadId);
  assert.deepStrictEqual(ordered.slice(0, 2), ['no-score-new', 'no-score-unread']);
}

function testGuardrails() {
  const items = Array.from({ length: MAX_THREADS_PER_BATCH + 1 }, (_, idx) => ({
    threadId: `t${idx}`,
    contentLength: 1
  }));
  const plan = applyGuardrails(items);
  assert.equal(plan.selected.length, MAX_THREADS_PER_BATCH);
  assert.equal(plan.deferred.length, 1);

  const planChars = applyGuardrails([
    { threadId: 'big', contentLength: MAX_TOTAL_CHARS_PER_BATCH },
    { threadId: 'small', contentLength: 10 }
  ]);
  assert.deepStrictEqual(planChars.selected, ['big']);
  assert.deepStrictEqual(planChars.deferred, ['small']);
}

function run() {
  testReconcileTargets();
  testExtractHistoryThreadIds();
  testSortCandidates();
  testGuardrails();
  console.log('sync tests passed');
}

run();
