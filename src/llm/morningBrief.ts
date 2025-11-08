import OpenAI from 'openai';

export type InboxBrief = {
  title: string;
  overview: string;
  highlights: string[];
};

type SummaryRecord = {
  headline: string | null;
  category: string | null;
  tldr: string | null;
  nextStep: string | null;
  createdAt: Date;
  Thread?: {
    subject?: string | null;
    fromName?: string | null;
    fromEmail?: string | null;
    lastMessageTs?: Date | null;
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
- highlights: array of all actionable reminders for the day, ordered by priority-level / urgency ("Reply to Sam re: contract by noon"). Keep the reminder to < 200 characters for quick skim

Tone: calm, decisive, proactive. Mention specific categories/people whenever helpful. Look for patterns across multiple emails to synthesize higher-level insights.
Avoid repeating the same verbs. Focus on what's actionable today.`;

const NO_DATA_BRIEF: InboxBrief = {
  title: 'Inbox is clear',
  overview: 'No active summaries yet. Ingest your inbox to generate today\'s briefing.',
  highlights: []
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

  const topHighlights = summaries
    .filter(hasAction)
    .slice(0, 3)
    .map(summary => {
      const subject = sanitizeText(summary.headline || summary.Thread?.subject || 'Inbox item');
      const next = sanitizeText(summary.nextStep || 'Follow up');
      return `${subject}: ${next}`;
    });

  const highlights = topHighlights.length
    ? topHighlights
    : ['Nothing requires action right now — enjoy the empty queue.'];

  return {
    title: 'Inbox snapshot',
    overview: `${total} active threads, ${needAction} need attention. Top categories: ${cats}.`,
    highlights
  };
}

function summarizeEntries(summaries: SummaryRecord[]) {
  return summaries.map((summary, index) => {
    const subject = sanitizeText(summary.Thread?.subject || summary.headline || '');
    const tldr = sanitizeText(summary.tldr);
    const next = sanitizeText(summary.nextStep || 'No action');
    const category = titleCaseCategory(summary.category || 'FYI');
    const from = sanitizeText(summary.Thread?.fromName || summary.Thread?.fromEmail || '');
    const when = summary.Thread?.lastMessageTs || summary.createdAt;
    return `${index + 1}. Subject: ${subject} | Category: ${category} | From: ${from || 'Unknown'} | TLDR: ${tldr} | Next: ${next} | Last Activity: ${when.toISOString()}`;
  }).join('\n');
}

function normalizeHighlights(value: unknown): string[] {
  if (Array.isArray(value)) return value.map(sanitizeText).filter(Boolean).slice(0, 4);
  if (typeof value === 'string') {
    return value
      .split(/\n|\||;|•/)
      .map(sanitizeText)
      .filter(Boolean)
      .slice(0, 4);
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
      highlights: normalizeHighlights(parsed.highlights)
    };

    if (!brief.highlights.length) brief.highlights = fallback.highlights;
    return brief;
  } catch (err) {
    console.error('Failed to build inbox brief', err);
    return fallback;
  }
}
