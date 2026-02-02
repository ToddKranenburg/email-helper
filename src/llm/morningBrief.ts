import OpenAI from 'openai';

export type InboxBrief = {
  title: string;
  overview: string;
  highlights: string[];
  highlightTargets: string[];
};

type SummaryRecord = {
  headline: string | null;
  category: string | null;
  tldr: string | null;
  nextStep: string | null;
  createdAt: Date;
  threadId: string;
  threadIndex?: {
    subject?: string | null;
    fromName?: string | null;
    fromEmail?: string | null;
    lastMessageDate?: Date | null;
  } | null;
};

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

const SYSTEM_PROMPT = `You are an executive assistant writing a crisp morning briefing.
GOAL: give the user an at-a-glance note covering what needs attention in their inbox.

OUTPUT STRICT JSON with keys: title, overview, highlights.
- title: 3-6 word headline (no emojis).
- overview: 2 short sentences (<=60 words total) summarizing workload & priorities.
- highlights: array of all actionable reminders for the day, ordered by priority-level / urgency ("Reply to Sam re: contract by noon"). Mention essential information you'd need for context like proper nouns. Keep  reminder to < 150 characters for quick skim -- no need for full sentences

Tone: calm, decisive, proactive. Mention specific categories/people whenever helpful. Look for patterns across multiple emails to synthesize higher-level insights.
Avoid repeating the same verbs. Focus on what's actionable today.`;

const NO_DATA_BRIEF: InboxBrief = {
  title: 'Inbox is clear',
  overview: 'No active summaries yet. Ingest your inbox to generate today\'s briefing.',
  highlights: [],
  highlightTargets: []
};

function sanitizeText(value: unknown): string {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim();
}

function titleCaseCategory(value: string | null): string {
  return sanitizeText(value)
    .toLowerCase()
    .replace(/\b([a-z])/g, (_, c: string) => c.toUpperCase());
}

function extractJSONObject(payload: string) {
  const fenced = payload.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced ? fenced[1] : payload;
  const brace = candidate.match(/\{[\s\S]*\}/);
  const jsonStr = brace ? brace[0] : candidate;
  return JSON.parse(jsonStr);
}

function categoryCounts(summaries: SummaryRecord[]): string {
  const counts = summaries.reduce<Record<string, number>>((acc, summary) => {
    const key = titleCaseCategory(summary.category || 'FYI') || 'FYI';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});

  const parts = Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 4)
    .map(([cat, count]) => `${cat}: ${count}`);

  return parts.join(', ');
}

function hasAction(summary: SummaryRecord): boolean {
  const next = sanitizeText(summary.nextStep || '');
  return next.length > 0 && !/^no action$/i.test(next);
}

function fallbackBrief(summaries: SummaryRecord[]): InboxBrief {
  if (!summaries.length) return NO_DATA_BRIEF;

  const total = summaries.length;
  const needAction = summaries.filter(hasAction).length;
  const cats = categoryCounts(summaries) || 'FYI: 0';

  const actionable = summaries.filter(hasAction);
  const topHighlights = actionable.map(summary => {
    const subject = sanitizeText(summary.headline || summary.threadIndex?.subject || 'Inbox item');
    const next = sanitizeText(summary.nextStep || 'Follow up');
    return `${subject}: ${next}`;
  });
  const highlightTargets = actionable.map(summary => summary.threadId);

  const highlights = topHighlights.length
    ? topHighlights
    : ['Nothing requires action right now — enjoy the empty queue.'];

  return {
    title: 'Inbox snapshot',
    overview: `${total} active threads, ${needAction} need attention. Top categories: ${cats}.`,
    highlights,
    highlightTargets: topHighlights.length ? highlightTargets : []
  };
}

function summarizeEntries(summaries: SummaryRecord[]) {
  return summaries.map((summary, index) => {
    const subject = sanitizeText(summary.threadIndex?.subject || summary.headline || '');
    const tldr = sanitizeText(summary.tldr);
    const next = sanitizeText(summary.nextStep || 'No action');
    const category = titleCaseCategory(summary.category || 'FYI');
    const from = sanitizeText(summary.threadIndex?.fromName || summary.threadIndex?.fromEmail || '');
    const when = summary.threadIndex?.lastMessageDate || summary.createdAt;
    return `${index + 1}. Subject: ${subject} | Category: ${category} | From: ${from || 'Unknown'} | TLDR: ${tldr} | Next: ${next} | Last Activity: ${when.toISOString()}`;
  }).join('\n');
}

function normalizeHighlights(value: unknown): string[] {
  if (Array.isArray(value)) return value.map(sanitizeText).filter(Boolean);
  if (typeof value === 'string') {
    return value
      .split(/\n|\||;|•/)
      .map(sanitizeText)
      .filter(Boolean);
  }
  return [];
}

export async function buildInboxBrief(summaries: SummaryRecord[]): Promise<InboxBrief> {
  if (!summaries.length) return NO_DATA_BRIEF;

  const fallback = fallbackBrief(summaries);
  if (!openai) return fallback;

  try {
    const actionable = summaries.filter(hasAction).length;
    const counts = categoryCounts(summaries);
    const subset = summaries.slice(0, 12);
    const dataset = summarizeEntries(subset);

    const userPrompt = `You have ${summaries.length} active threads; ${actionable} need action.
Category mix: ${counts || 'None'}.

Threads:
${dataset}

Write the JSON briefing.`;

    const response = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      temperature: 0.2,
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: userPrompt }
      ],
      response_format: { type: 'json_object' as const }
    });

    const raw = response.choices[0]?.message?.content ?? '';
    const parsed = extractJSONObject(raw);

    const brief: InboxBrief = {
      title: sanitizeText(parsed.title) || fallback.title,
      overview: sanitizeText(parsed.overview) || fallback.overview,
      highlights: normalizeHighlights(parsed.highlights),
      highlightTargets: []
    };

    if (!brief.highlights.length) {
      brief.highlights = fallback.highlights;
      brief.highlightTargets = fallback.highlightTargets;
      return brief;
    }

    const mappedTargets = mapHighlightsToTargets(brief.highlights, summaries);
    const finalTargets = brief.highlights.map((_, index) => mappedTargets[index] || fallback.highlightTargets[index] || '');
    brief.highlightTargets = finalTargets;
    return brief;
  } catch (err) {
    console.error('Failed to build inbox brief', err);
    return fallback;
  }
}

function tokenizeForMatch(text: string): string[] {
  return sanitizeText(text)
    .toLowerCase()
    .match(/[a-z0-9]{3,}/g) ?? [];
}

function overlapCount(highlightTokens: string[], targetText: string): number {
  if (!targetText) return 0;
  const tokens = tokenizeForMatch(targetText);
  if (!tokens.length || !highlightTokens.length) return 0;
  const highlightSet = new Set(highlightTokens);
  return tokens.reduce((acc, token) => acc + (highlightSet.has(token) ? 1 : 0), 0);
}

function mapHighlightsToTargets(highlights: string[], summaries: SummaryRecord[]): string[] {
  if (!highlights.length) return [];
  const actionable = summaries.filter(hasAction);
  if (!actionable.length) return [];

  const assigned = new Set<string>();
  return highlights.map(highlight => {
    const match = findBestSummaryForHighlight(highlight, actionable, assigned);
    if (match) {
      assigned.add(match.threadId);
      return match.threadId;
    }
    const fallback = actionable.find(summary => !assigned.has(summary.threadId));
    if (fallback) {
      assigned.add(fallback.threadId);
      return fallback.threadId;
    }
    return '';
  });
}

function findBestSummaryForHighlight(
  highlight: string,
  summaries: SummaryRecord[],
  used: Set<string>
): SummaryRecord | null {
  const tokens = tokenizeForMatch(highlight);
  if (!tokens.length) return null;
  const normalized = sanitizeText(highlight).toLowerCase();
  let best: { summary: SummaryRecord; score: number } | null = null;

  for (const summary of summaries) {
    if (used.has(summary.threadId)) continue;
    const subject = sanitizeText(summary.threadIndex?.subject || summary.headline || '');
    const next = sanitizeText(summary.nextStep || '');
    const from = sanitizeText(summary.threadIndex?.fromName || summary.threadIndex?.fromEmail || '');
    const category = titleCaseCategory(summary.category || '');
    let score = 0;

    if (subject && normalized.includes(subject.toLowerCase())) score += 6;
    if (next && normalized.includes(next.toLowerCase())) score += 4;
    if (from && normalized.includes(from.toLowerCase())) score += 3;
    if (category && normalized.includes(category.toLowerCase())) score += 1;

    score += overlapCount(tokens, subject) * 3;
    score += overlapCount(tokens, next) * 2;
    score += overlapCount(tokens, from);

    if (score > 0 && (!best || score > best.score)) {
      best = { summary, score };
    }
  }

  return best?.summary ?? null;
}
