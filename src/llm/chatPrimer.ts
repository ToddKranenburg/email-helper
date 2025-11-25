import OpenAI from 'openai';

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

export type ChatPrimerInput = {
  threadId: string;
  subject: string;
  summary: string;
  nextStep: string;
  headline: string;
  fromLine: string;
};

const SYSTEM_PROMPT = `You are the texting voice of a hyper-capable Gen Z executive assistant who sounds like a modern, confident secretary.
Your job: craft the first message the assistant sends after scanning an email thread.
Rules:
- Output JSON object: {"primers":[{ "threadId": "...", "prompt": "..." }]}
- One entry per thread ID provided.
- Start by summarizing what the user just received (sender, org, subject, or summary) so they instantly know the situation.
- After the context, propose the most logical next step (use NextStep if provided) and end by explicitly confirming if they want you to handle it.
- Tone: casual-but-professional texting style, crisp, no fillers, always helpful.
- Keep it punchy while staying descriptive.`;

const FALLBACK_SUGGESTIONS = [
  'Want me to draft a reply and get it queued up?',
  'Should I send a quick follow-up so this keeps moving?',
  'Want me to set a reminder so it doesn’t fall through?'
];

export async function generateChatPrimers(entries: ChatPrimerInput[]): Promise<Record<string, string>> {
  const result: Record<string, string> = {};
  if (!entries.length) return result;
  if (!openai) {
    for (const entry of entries) {
      result[entry.threadId] = fallbackPrimer(entry);
    }
    return result;
  }

  const chunks = chunk(entries, 6);
  for (const batch of chunks) {
    const payload = batch.map(item => {
      return `threadId: ${item.threadId}
Subject: ${item.subject || '(no subject)'}
Headline: ${item.headline || '(no headline)'}
From: ${item.fromLine || '(unknown sender)'}
Summary: ${item.summary || '(no summary)'}
NextStep: ${item.nextStep || 'No action'}`;
    }).join('\n---\n');

    try {
      const completion = await openai.chat.completions.create({
        model: 'gpt-4o-mini',
        temperature: 0.5,
        messages: [
          { role: 'system', content: SYSTEM_PROMPT },
          {
            role: 'user',
            content:
              `Create prompts for these threads:\n${payload}\n\nRemember: respond with JSON using the schema described in the system prompt.`
          }
        ]
      });
      const text = completion.choices[0]?.message?.content || '';
      const parsed = parsePrimerResponse(text);
      for (const entry of parsed) {
        if (entry.threadId && entry.prompt) {
          result[entry.threadId] = entry.prompt.trim();
        }
      }
    } catch (err) {
      console.error('Failed to generate chat primers', err);
    }
  }

  for (const entry of entries) {
    if (!result[entry.threadId]) {
      result[entry.threadId] = fallbackPrimer(entry);
    }
  }
  return result;
}

function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

function parsePrimerResponse(payload: string): { threadId: string; prompt: string }[] {
  if (!payload) return [];
  const match = payload.match(/\{[\s\S]*\}/);
  if (!match) return [];
  try {
    const json = JSON.parse(match[0]);
    if (Array.isArray(json)) {
      return json.filter(item => typeof item?.threadId === 'string' && typeof item?.prompt === 'string');
    }
    if (Array.isArray(json?.primers)) {
      return json.primers.filter((item: any) => typeof item?.threadId === 'string' && typeof item?.prompt === 'string');
    }
    return [];
  } catch {
    return [];
  }
}

function fallbackPrimer(entry: ChatPrimerInput): string {
  const summaryBase = (entry.summary || entry.headline || entry.subject || 'an email that needs your call').trim();
  const normalizedSummary = summaryBase.replace(/\s+/g, ' ');
  const sender = entry.fromLine ? entry.fromLine.trim() : '';
  const context = sender ? `${normalizedSummary} from ${sender}` : normalizedSummary;
  const hasNextStep = entry.nextStep && entry.nextStep.toLowerCase() !== 'no action';
  const action = hasNextStep ? entry.nextStep!.trim() : '';
  const next = hasNextStep ? `Want me to ${action}?` : randomSuggestion();
  return `Heads up: looks like ${context}. ${next}`;
}

function randomSuggestion() {
  return FALLBACK_SUGGESTIONS[Math.floor(Math.random() * FALLBACK_SUGGESTIONS.length)];
}
