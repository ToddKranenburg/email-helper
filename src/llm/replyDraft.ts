import OpenAI from 'openai';

export type ReplyDraftResult = {
  body: string;
  confidence: number;
  safeToDraft: boolean;
  reason: string;
};

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

const SYSTEM_PROMPT = `You draft short email replies using ONLY the provided transcript as ground truth.
Output STRICT JSON: {"body":"...","confidence":0-1,"safe_to_draft":true|false,"reason":"short phrase"}.
Rules:
- Never invent names, dates, numbers, commitments, or facts not explicitly stated in the transcript.
- If any detail needed to reply is missing or ambiguous, return body="" and safe_to_draft=false with confidence <= 0.3.
- Do not include a greeting or signature. The reply body should be ready to paste after a greeting.
- Keep it under 120 words. Use plain text only (no markdown).`;

const GUIDED_SYSTEM_PROMPT = `You draft short, warm, and terse email replies using the thread context plus the user's desired reply instructions.
Output STRICT JSON: {"body":"...","confidence":0-1,"safe_to_draft":true|false,"reason":"short phrase"}.
Rules:
- Use the user's instruction as authoritative. Use the transcript for facts and context.
- Be creative in phrasing (do not copy the user's words verbatim), while preserving meaning.
- Never invent names, dates, numbers, commitments, or facts not in the transcript or user instruction.
- If information is missing or ambiguous, ask one concise clarifying question instead of guessing.
- Do not include a greeting or signature. The reply body should be ready to paste after a greeting.
- Keep it under 120 words. Use plain text only (no markdown).`;

const FALLBACK: ReplyDraftResult = {
  body: '',
  confidence: 0,
  safeToDraft: false,
  reason: 'OpenAI not configured'
};

export async function generateReplyDraft(input: {
  subject: string;
  headline?: string;
  summary?: string;
  nextStep?: string;
  participants: string[];
  transcript: string;
  fromLine?: string;
}): Promise<ReplyDraftResult> {
  const context = buildContext(input);
  if (!openai) return FALLBACK;
  try {
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      temperature: 0,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: SYSTEM_PROMPT },
        { role: 'user', content: context }
      ]
    });
    const raw = completion.choices[0]?.message?.content || '';
    const parsed = parseReplyDraft(raw);
    if (parsed) return parsed;
  } catch (err) {
    console.error('reply draft generation failed', err);
  }
  return { ...FALLBACK, reason: 'Reply draft unavailable' };
}

export async function generateGuidedReplyDraft(input: {
  subject: string;
  headline?: string;
  summary?: string;
  nextStep?: string;
  participants: string[];
  transcript: string;
  fromLine?: string;
  userInstruction: string;
}): Promise<ReplyDraftResult> {
  const context = buildGuidedContext(input);
  if (!openai) return FALLBACK;
  try {
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      temperature: 0.4,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: GUIDED_SYSTEM_PROMPT },
        { role: 'user', content: context }
      ]
    });
    const raw = completion.choices[0]?.message?.content || '';
    const parsed = parseReplyDraft(raw);
    if (parsed) return parsed;
  } catch (err) {
    console.error('guided reply draft generation failed', err);
  }
  return { ...FALLBACK, reason: 'Guided reply draft unavailable' };
}

function buildContext(input: {
  subject: string;
  headline?: string;
  summary?: string;
  nextStep?: string;
  participants: string[];
  transcript: string;
  fromLine?: string;
}) {
  const participants = input.participants?.length ? input.participants.join(', ') : 'Unknown participants';
  const trimmedTranscript = input.transcript?.length > 9000
    ? input.transcript.slice(-9000)
    : input.transcript || '(no transcript provided)';
  return `Subject: ${input.subject || '(no subject)'}
Participants: ${participants}
Sender line: ${input.fromLine || '(unknown sender)'}

Thread (oldest to newest):
${trimmedTranscript}`;
}

function buildGuidedContext(input: {
  subject: string;
  headline?: string;
  summary?: string;
  nextStep?: string;
  participants: string[];
  transcript: string;
  fromLine?: string;
  userInstruction: string;
}) {
  const participants = input.participants?.length ? input.participants.join(', ') : 'Unknown participants';
  const trimmedTranscript = input.transcript?.length > 9000
    ? input.transcript.slice(-9000)
    : input.transcript || '(no transcript provided)';
  const instruction = input.userInstruction?.trim() || '(no instruction provided)';
  const headline = input.headline?.trim() || '';
  const summary = input.summary?.trim() || '';
  const nextStep = input.nextStep?.trim() || '';
  const extraContext = [headline && `Headline: ${headline}`, summary && `Summary: ${summary}`, nextStep && `Next step: ${nextStep}`]
    .filter(Boolean)
    .join('\n');
  return `Subject: ${input.subject || '(no subject)'}
Participants: ${participants}
Sender line: ${input.fromLine || '(unknown sender)'}
User instruction: ${instruction}
${extraContext ? `\n${extraContext}` : ''}

Thread (oldest to newest):
${trimmedTranscript}`;
}

function parseReplyDraft(payload: string): ReplyDraftResult | null {
  const safe = extractJson(payload);
  if (!safe) return null;
  const body = normalizeBody(safe.body);
  const confidenceRaw = typeof safe.confidence === 'number' ? safe.confidence : 0;
  const confidence = clamp(confidenceRaw, 0, 1);
  const safeToDraft = Boolean(safe.safe_to_draft ?? safe.safeToDraft);
  const reason = typeof safe.reason === 'string' && safe.reason.trim()
    ? safe.reason.trim()
    : 'Model decision';
  if (!body || !safeToDraft) {
    return { body: '', confidence, safeToDraft: false, reason };
  }
  return { body, confidence, safeToDraft, reason };
}

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

function normalizeBody(value: unknown) {
  const text = typeof value === 'string' ? value.trim() : '';
  if (!text) return '';
  return text.replace(/\r\n/g, '\n').trim();
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}
