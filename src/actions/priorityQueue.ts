import { prisma } from '../store/db.js';
import type { ActionFlow, Summary, Thread } from '@prisma/client';

export type PriorityEntry = {
  threadId: string;
  score: number;
  reason: string;
  reasonWeight: number;
};

type SummaryWithThread = Summary & { Thread: Thread | null };

type PriorityCandidate = {
  threadId: string;
  headline: string;
  summary: string;
  nextStep: string;
  category: string;
  subject: string;
  receivedAt: Date;
  actionType: string;
};

type PrioritySignal = { label: string; weight: number };

const PRIORITY_PATTERNS = {
  urgent: /\b(urgent|asap|immediately|time[-\s]?sensitive|deadline|final notice|action required|response required|reply needed|respond by|past due|overdue|expir(?:e|es|ing))\b/i,
  security: /\b(security|verify|verification|password|2fa|unauthorized|suspicious|fraud|breach|locked?|login|sign[-\s]?in|account alert)\b/i,
  payment: /\b(payment|invoice|receipt|billing|charge|charged|refund|past due|overdue|card)\b/i,
  approval: /\b(approve|approval|sign[-\s]?off|contract|legal|compliance|policy)\b/i,
  scheduling: /\b(meeting|call|calendar|schedule|reschedule|availability|zoom|appointment|rsvp|invite)\b/i
};

const PRIORITY_MIN_SCORE = 4;

export async function loadPriorityQueue(userId: string, opts: { limit: number; minScore?: number }) {
  const summaries = await prisma.summary.findMany({
    where: { userId },
    select: {
      threadId: true,
      headline: true,
      tldr: true,
      nextStep: true,
      category: true,
      createdAt: true,
      Thread: {
        select: {
          subject: true,
          lastMessageTs: true
        }
      }
    }
  });

  if (!summaries.length) {
    return { priority: [] as PriorityEntry[], items: [] as SummaryWithThread[] };
  }

  const flows = await prisma.actionFlow.findMany({
    where: { userId },
    select: { threadId: true, actionType: true }
  });
  const flowMap = new Map<string, ActionFlow['actionType']>();
  flows.forEach(flow => {
    if (!flow?.threadId) return;
    flowMap.set(flow.threadId, flow.actionType || '');
  });

  const minScore = typeof opts.minScore === 'number' ? opts.minScore : PRIORITY_MIN_SCORE;
  const scored = summaries
    .map(summary => {
      const candidate: PriorityCandidate = {
        threadId: summary.threadId,
        headline: summary.headline || '',
        summary: summary.tldr || '',
        nextStep: summary.nextStep || '',
        category: summary.category || '',
        subject: summary.Thread?.subject || '',
        receivedAt: summary.Thread?.lastMessageTs || summary.createdAt,
        actionType: flowMap.get(summary.threadId) || ''
      };
      const evaluation = scoreThreadPriority(candidate);
      return {
        threadId: candidate.threadId,
        receivedAt: candidate.receivedAt,
        ...evaluation
      };
    })
    .filter(item => item.score >= minScore);

  scored.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    if (b.reasonWeight !== a.reasonWeight) return b.reasonWeight - a.reasonWeight;
    return b.receivedAt.getTime() - a.receivedAt.getTime();
  });

  const priority = scored.slice(0, Math.max(opts.limit || 0, 0)).map(item => ({
    threadId: item.threadId,
    score: item.score,
    reason: item.reason,
    reasonWeight: item.reasonWeight
  }));

  const priorityIds = priority.map(item => item.threadId);
  if (!priorityIds.length) {
    return { priority, items: [] as SummaryWithThread[] };
  }

  const items = await prisma.summary.findMany({
    where: { userId, threadId: { in: priorityIds } },
    include: { Thread: true }
  }) as SummaryWithThread[];

  const itemMap = new Map(items.map(item => [item.threadId, item]));
  const orderedItems = priorityIds.map(id => itemMap.get(id)).filter(Boolean) as SummaryWithThread[];

  return { priority, items: orderedItems };
}

function scoreThreadPriority(candidate: PriorityCandidate) {
  const signals: PrioritySignal[] = [];
  let score = 0;
  const text = buildPriorityText(candidate);

  if (requiresAction(candidate.nextStep)) {
    score += addPrioritySignal(signals, 'Action needed', 2);
  }

  if (candidate.actionType === 'external_action') {
    score += addPrioritySignal(signals, 'Action required', 4);
  }

  score += applyDueSignal(text, signals);

  if (PRIORITY_PATTERNS.urgent.test(text)) {
    score += addPrioritySignal(signals, 'Time-sensitive', 2);
  }
  if (PRIORITY_PATTERNS.security.test(text)) {
    score += addPrioritySignal(signals, 'Account alert', 4);
  }
  if (PRIORITY_PATTERNS.payment.test(text)) {
    score += addPrioritySignal(signals, 'Payment issue', 2);
  }
  if (PRIORITY_PATTERNS.approval.test(text)) {
    score += addPrioritySignal(signals, 'Needs approval', 2);
  }
  if (PRIORITY_PATTERNS.scheduling.test(text)) {
    score += addPrioritySignal(signals, 'Scheduling', 1);
  }

  const categoryBoost = categoryPriorityBoost(candidate.category);
  score += categoryBoost.score;
  if (categoryBoost.label) {
    addPrioritySignal(signals, categoryBoost.label, categoryBoost.score);
  }

  const normalizedScore = Math.max(0, score);
  const reason = pickPriorityReason(signals);
  return {
    score: normalizedScore,
    reason: reason?.label || 'Needs attention',
    reasonWeight: reason?.weight || 0
  };
}

function buildPriorityText(candidate: PriorityCandidate) {
  return [candidate.nextStep, candidate.summary, candidate.headline, candidate.subject].filter(Boolean).join(' ');
}

function requiresAction(nextStep: string) {
  const text = (nextStep || '').toLowerCase();
  if (!text) return false;
  return !/(no action|fyi|none|no need|no response needed)/i.test(text);
}

function addPrioritySignal(signals: PrioritySignal[], label: string, weight: number) {
  if (!label || !weight) return 0;
  signals.push({ label, weight });
  return weight;
}

function applyDueSignal(text: string, signals: PrioritySignal[]) {
  const due = extractDueDate(text);
  if (!due) return 0;
  const days = diffCalendarDays(new Date(), due);
  if (days < 0) return addPrioritySignal(signals, 'Overdue', 4);
  if (days === 0) return addPrioritySignal(signals, 'Due today', 3);
  if (days === 1) return addPrioritySignal(signals, 'Due tomorrow', 3);
  if (days <= 3) return addPrioritySignal(signals, 'Due soon', 2);
  if (days <= 7) return addPrioritySignal(signals, 'Due this week', 1);
  return 0;
}

function extractDueDate(text: string) {
  if (!text) return null;
  const lower = text.toLowerCase();
  const inDays = lower.match(/\bin\s+(\d+)\s+days?\b/);
  if (inDays) {
    const days = Number(inDays[1]);
    if (Number.isFinite(days)) return addDays(new Date(), days);
  }
  if (lower.includes('end of day') || lower.includes('eod')) {
    return new Date();
  }
  if (lower.includes('end of week') || lower.includes('by end of week') || lower.includes('this week')) {
    return nextWeekdayDate(5);
  }
  if (lower.includes('tomorrow')) {
    return addDays(new Date(), 1);
  }
  if (lower.includes('today')) {
    return new Date();
  }
  const weekday = detectWeekday(lower);
  if (weekday !== null) {
    return nextWeekdayDate(weekday);
  }
  const isoMatch = text.match(/\b(\d{4}-\d{2}-\d{2})\b/);
  if (isoMatch) {
    const parsed = new Date(isoMatch[1]);
    return isValidDate(parsed) ? parsed : null;
  }
  const slash = text.match(/\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/);
  if (slash) {
    const month = Number(slash[1]) - 1;
    const day = Number(slash[2]);
    const year = slash[3]
      ? Number(slash[3].length === 2 ? `20${slash[3]}` : slash[3])
      : new Date().getFullYear();
    const parsed = new Date(year, month, day);
    return isValidDate(parsed) ? parsed : null;
  }
  return null;
}

function detectWeekday(text: string) {
  const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
  for (let i = 0; i < days.length; i += 1) {
    const needle = days[i];
    const pattern = new RegExp(`\\b(?:by|on|this|next)?\\s*${needle}\\b`, 'i');
    if (pattern.test(text)) return i;
  }
  return null;
}

function nextWeekdayDate(targetDay: number) {
  const today = new Date();
  const result = new Date(today);
  const delta = (targetDay - today.getDay() + 7) % 7 || 7;
  result.setDate(today.getDate() + delta);
  return result;
}

function addDays(date: Date, days: number) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function isValidDate(date: Date) {
  return date instanceof Date && !Number.isNaN(date.getTime());
}

function diffCalendarDays(from: Date, to: Date) {
  const start = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  const end = new Date(to.getFullYear(), to.getMonth(), to.getDate());
  return Math.round((end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000));
}

function categoryPriorityBoost(category: string) {
  const label = typeof category === 'string' ? category.toLowerCase() : '';
  if (label.includes('billing')) return { score: 2, label: 'Billing' };
  if (label.includes('personal request')) return { score: 2, label: 'Personal request' };
  if (label.includes('introduction')) return { score: 1, label: 'Introduction' };
  if (label.includes('personal event')) return { score: 1, label: 'Event planning' };
  if (label.includes('catch up')) return { score: 0, label: '' };
  if (label.includes('marketing') || label.includes('promotion')) return { score: -3, label: '' };
  if (label.includes('editorial') || label.includes('writing')) return { score: -2, label: '' };
  if (label.includes('fyi')) return { score: -1, label: '' };
  return { score: 0, label: '' };
}

function pickPriorityReason(signals: PrioritySignal[]) {
  if (!signals.length) return null;
  const sorted = signals.slice().sort((a, b) => b.weight - a.weight);
  return sorted[0] || null;
}
