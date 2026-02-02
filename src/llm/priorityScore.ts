import OpenAI from 'openai';

export type PriorityScoreInput = {
  threadId: string;
  subject: string;
  participants: string[];
  snippet: string;
  content: string;
};

export type PriorityScoreResult = {
  priorityScore: number;
  priorityReason: string;
  suggestedActionType: string;
  extracted?: {
    deadlines?: string[];
    asks?: string[];
    people?: string[];
  } | null;
};

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

const SYSTEM_PROMPT = `You are an email triage assistant. Your job is to assign a single priority score and a short reason.
OUTPUT STRICT JSON with keys: priority_score, priority_reason, suggested_action_type, extracted.
- priority_score: number from 0-100 (100 = most urgent/important).
- priority_reason: 1-2 sentences, plain text, explain the urgency.
- suggested_action_type: one of reply | schedule | task | wait | archive | label.
- extracted: optional object with arrays: deadlines, asks, people.
Rules:
- Treat security alerts, account lock risk, payment issues, legal/contract deadlines, and time-sensitive requests as high priority.
- Marketing/newsletters/announcements without action should be low priority (0-20) and usually archive or label.
- If there is a clear ask or decision needed, prefer reply.
- If there is a time-bound meeting request, prefer schedule.
- If it requires future follow-up with no immediate response, prefer task.
- If no action is needed, prefer wait or archive.
Return ONLY JSON.`;

const PRIORITY_PATTERNS = {
  urgent: /\b(urgent|asap|immediately|time[-\s]?sensitive|deadline|final notice|action required|response required|reply needed|respond by|past due|overdue|expir(?:e|es|ing))\b/i,
  security: /\b(security|verify|verification|password|2fa|unauthorized|suspicious|fraud|breach|locked?|login|sign[-\s]?in|account alert)\b/i,
  payment: /\b(payment|invoice|receipt|billing|charge|charged|refund|past due|overdue|card)\b/i,
  approval: /\b(approve|approval|sign[-\s]?off|contract|legal|compliance|policy)\b/i,
  scheduling: /\b(meeting|call|calendar|schedule|reschedule|availability|zoom|appointment|rsvp|invite)\b/i,
  marketing: /\b(unsubscribe|newsletter|promotion|sale|discount|marketing)\b/i
};

function extractJson(payload: string) {
  if (!payload) return null;
  const match = payload.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch {
    return null;
  }
}

function clampScore(value: unknown) {
  const num = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.min(100, Math.max(0, Math.round(num)));
}

function normalizeActionType(value: unknown) {
  const raw = typeof value === 'string' ? value.trim().toLowerCase() : '';
  if (['reply', 'schedule', 'task', 'wait', 'archive', 'label'].includes(raw)) return raw;
  return 'wait';
}

function normalizeReason(value: unknown) {
  const text = typeof value === 'string' ? value.trim() : '';
  if (!text) return 'Needs attention.';
  return text.replace(/\s+/g, ' ');
}

function normalizeExtracted(raw: any) {
  if (!raw || typeof raw !== 'object') return null;
  const pickArray = (value: unknown) => Array.isArray(value) ? value.map(item => String(item || '').trim()).filter(Boolean) : [];
  return {
    deadlines: pickArray(raw.deadlines),
    asks: pickArray(raw.asks),
    people: pickArray(raw.people)
  };
}

function fallbackPriority(input: PriorityScoreInput): PriorityScoreResult {
  const text = `${input.subject}\n${input.snippet}\n${input.content}`.toLowerCase();
  let score = 10;
  let reason = 'Low urgency.';
  if (PRIORITY_PATTERNS.marketing.test(text)) {
    return { priorityScore: 10, priorityReason: 'Promotional or informational email.', suggestedActionType: 'archive', extracted: null };
  }
  if (PRIORITY_PATTERNS.security.test(text)) {
    score += 50;
    reason = 'Account or security alert requires attention.';
  }
  if (PRIORITY_PATTERNS.urgent.test(text)) {
    score += 25;
    reason = 'Time-sensitive request.';
  }
  if (PRIORITY_PATTERNS.payment.test(text)) {
    score += 20;
    reason = 'Billing or payment related.';
  }
  if (PRIORITY_PATTERNS.approval.test(text)) {
    score += 15;
    reason = 'Approval or legal decision needed.';
  }
  if (PRIORITY_PATTERNS.scheduling.test(text)) {
    score += 10;
    reason = 'Scheduling request.';
  }
  const suggestedActionType = score >= 60 ? 'reply' : score >= 30 ? 'task' : 'wait';
  return { priorityScore: Math.min(score, 100), priorityReason: reason, suggestedActionType, extracted: null };
}

export async function scoreThreadPriority(input: PriorityScoreInput): Promise<PriorityScoreResult> {
  if (!openai) return fallbackPriority(input);
  const content = input.content.length > 25_000 ? input.content.slice(0, 25_000) : input.content;
  const userPrompt = `Subject: ${input.subject || '(no subject)'}
Participants: ${(input.participants || []).join(', ') || 'Unknown'}
Snippet: ${input.snippet || '(none)'}

Thread (oldest→newest):
${content}`;

  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      temperature: 0,
      response_format: { type: 'json_object' as const },
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: userPrompt }
      ]
    });

    const raw = response.choices[0]?.message?.content ?? '';
    const parsed = extractJson(raw);
    if (!parsed) return fallbackPriority(input);

    const priorityScore = clampScore(parsed.priority_score ?? parsed.priorityScore);
    const priorityReason = normalizeReason(parsed.priority_reason ?? parsed.priorityReason);
    const suggestedActionType = normalizeActionType(parsed.suggested_action_type ?? parsed.suggestedActionType);
    const extracted = normalizeExtracted(parsed.extracted);

    return { priorityScore, priorityReason, suggestedActionType, extracted };
  } catch (err) {
    console.error('priority scoring failed', err);
    return fallbackPriority(input);
  }
}
