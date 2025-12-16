import OpenAI from 'openai';

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

export type ArchiveIntentDecision = {
  archive: boolean;
  confidence: number;
  reason: string;
  source: 'model' | 'heuristic';
};

export type SecretaryIntent = 'archive' | 'skip' | 'none';

export type SecretaryIntentDecision = {
  intent: SecretaryIntent;
  confidence: number;
  reason: string;
  source: 'model' | 'heuristic';
};

const INTENT_PROMPT = `Classify the user's ask for an email assistant.
- "archive": They want the thread cleared from the inbox / archived.
- "skip": They want to skip handling this thread for now (do not archive).
- "none": Anything else (questions, drafts, summaries, etc.).
Return JSON: {"intent":"archive|skip|none","confidence":0-1,"reason":"short phrase"}. Keep it concise.`;

export async function detectArchiveIntent(userText: string): Promise<ArchiveIntentDecision> {
  const decision = await classifyIntent(userText);
  return {
    archive: decision.intent === 'archive',
    confidence: decision.confidence,
    reason: decision.reason,
    source: decision.source
  };
}

export async function classifyIntent(userText: string): Promise<SecretaryIntentDecision> {
  const text = (userText || '').trim();
  if (!text) {
    return { intent: 'none', confidence: 0, reason: 'Empty input', source: 'heuristic' };
  }

  const archiveGuess = heuristicArchive(text);
  const skipGuess = heuristicSkip(text);
  const fallbackIntent: SecretaryIntent = archiveGuess ? 'archive' : skipGuess ? 'skip' : 'none';

  if (!openai) {
    return {
      intent: fallbackIntent,
      confidence: fallbackIntent === 'none' ? 0.05 : 0.35,
      reason: 'OpenAI not configured; heuristic applied',
      source: 'heuristic'
    };
  }

  try {
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      temperature: 0,
      messages: [
        { role: 'system', content: INTENT_PROMPT },
        {
          role: 'user',
          content:
            `User message: """${text.slice(0, 1200)}"""\nClassify intent.`
        }
      ]
    });
    const raw = completion.choices[0]?.message?.content || '';
    const parsed = parseIntentDecision(raw);
    if (parsed) return parsed;
    throw new Error('Unable to parse model response');
  } catch (err) {
    console.error('Secretary intent model failed', err);
    return {
      intent: fallbackIntent,
      confidence: fallbackIntent === 'none' ? 0.05 : 0.3,
      reason: 'Model unavailable; heuristic fallback used',
      source: 'heuristic'
    };
  }
}

function parseIntentDecision(payload: string): SecretaryIntentDecision | null {
  if (!payload) return null;
  const match = payload.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    const json = JSON.parse(match[0]);
    const intent = normalizeIntent(json?.intent);
    if (!intent) return null;
    const confidence = clamp(typeof json.confidence === 'number' ? json.confidence : 0.5, 0, 1);
    const reason = typeof json.reason === 'string' && json.reason.trim()
      ? json.reason.trim()
      : 'Model decision';
    return {
      intent,
      confidence,
      reason,
      source: 'model'
    };
  } catch {
    return null;
  }
}

function normalizeIntent(value: unknown): SecretaryIntent | null {
  const val = typeof value === 'string' ? value.toLowerCase().trim() : '';
  if (val === 'archive' || val === 'skip' || val === 'none') return val;
  return null;
}

function heuristicArchive(text: string): boolean {
  const cleaned = text.replace(/[.!?]/g, ' ').toLowerCase().trim();
  if (!cleaned) return false;

  const exactPhrases = new Set([
    'archive',
    'archive it',
    'archive this',
    'archive email',
    'archive message',
    'archive thread',
    'archive please'
  ]);
  if (exactPhrases.has(cleaned)) return true;

  const keywords = [
    'file this away',
    'file away',
    'put this away',
    'remove from inbox',
    'clear this out',
    'clear out my inbox',
    'clear it out',
    'done with this email',
    'done with this thread',
    'we are done here',
    'we are done with this',
    'no further action needed',
    'close this out',
    'mark as done and archive',
    'clean up my inbox',
    'can be archived',
    'go ahead and archive',
    'please archive'
  ];
  return keywords.some(phrase => cleaned.includes(phrase));
}

function heuristicSkip(text: string): boolean {
  const cleaned = text.replace(/[.!?]/g, ' ').toLowerCase().trim();
  if (!cleaned) return false;
  const skipPhrases = [
    'skip',
    'skip it',
    'skip this',
    'skip for now',
    'skip this one',
    'skip this email',
    'skip this thread',
    'not now',
    'later',
    'deal with later',
    'come back later',
    'remind me later',
    'remind me about this later',
    'move on',
    'move to the next',
    'next one',
    'next email',
    'next thread',
    'pass on this',
    'pass for now',
    'hold off',
    'park this',
    'leave it for now',
    'leave this for now'
  ];
  return skipPhrases.some(phrase => cleaned.includes(phrase));
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max);
}
