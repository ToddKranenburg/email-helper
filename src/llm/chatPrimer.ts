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

const SYSTEM_PROMPT = `You craft proactive opening prompts for an email follow-up assistant.
Each prompt should help the user ask a focused question about the specific thread.
Rules:
- Output JSON object: {"primers":[{ "threadId": "...", "prompt": "..." }]}
- One entry per thread ID provided.
- <= 30 words per prompt.
- Mention the most relevant detail (subject, sender, or next action).
- If a next step exists, point toward it or ask if they want to move forward.
- Tone: concise, energetic, no pleasantries, end with a question or clear suggestion.`;

const FALLBACK_SUGGESTIONS = [
  'Want to clarify the next action?',
  'Need help drafting the reply?',
  'Should we double-check the details?'
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
  const subject = entry.subject ? `**${entry.subject}**` : 'this email';
  const next = entry.nextStep && entry.nextStep.toLowerCase() !== 'no action'
    ? `Need to move on "${entry.nextStep}"?`
    : randomSuggestion();
  return `Need more detail on ${subject}? ${next}`;
}

function randomSuggestion() {
  return FALLBACK_SUGGESTIONS[Math.floor(Math.random() * FALLBACK_SUGGESTIONS.length)];
}
