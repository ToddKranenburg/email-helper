import OpenAI from 'openai';
import { performance } from 'node:perf_hooks';

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

export type SuggestedAction = 'archive' | 'more_info' | 'create_task' | 'skip';

export type PrimerOutput = {
  prompt: string;
  suggestedAction: SuggestedAction;
};

const SYSTEM_PROMPT = `You are the texting voice of a hyper-capable Gen Z AI executive assistant who sounds like a modern, confident secretary and wants to help the user get through their inbox, accepting your limitations as an AI agent.
Your job: craft the very first message after scanning an email thread.
Rules:
- Output JSON: {"primers":[{ "threadId": "...", "prompt": "...", "suggestedAction": "archive|more_info|create_task|skip" }]}
- One entry per thread ID provided.
- Start by grounding: briefly name what just arrived (sender/org + subject/summary) so the user instantly knows which email this is.
- Pick ONE most likely next action YOU would take as the recipient (archive | more_info | create_task | skip). Use NextStep if given; otherwise infer.
- The prompt must naturally state that action, give a quick why, and end by explicitly asking permission to do it. No separate “Suggested action” label, no extra lines.
- If you cannot justify an action, choose "more_info" and ask to dig deeper.
- Keep it tight, conversational, and decisive (texting length).
- Never claim you can directly RSVP, accept invites, click links, or send anything. If a response is needed (e.g., invite or approval), use only the allowed actions (archive|more_info|create_task|skip) and phrase it as drafting a reply, logging a reminder, or asking for more info before acting.`;

const FALLBACK_SUGGESTIONS = [
  'Want a tighter rundown? I can skim for deadlines and asks.',
  'Need more context? I can dig up the key decisions for you.',
  'Want me to pull the main points and next steps?'
];

type PrimerOptions = {
  traceId?: string;
};

export async function generateChatPrimers(
  entries: ChatPrimerInput[],
  options: PrimerOptions = {}
): Promise<Record<string, PrimerOutput>> {
  const result: Record<string, PrimerOutput> = {};
  const log = createPrimerLogger(options.traceId);
  const totalStart = performance.now();
  log('start', { count: entries.length, openai: Boolean(openai) });
  const inputLookup = new Map(entries.map(item => [item.threadId, item]));
  if (!entries.length) {
    log('no entries supplied');
    return result;
  }
  if (!openai) {
    log('openai not configured, using fallback primers only');
    for (const entry of entries) {
      result[entry.threadId] = fallbackPrimer(entry);
    }
    log('complete', { durationMs: elapsedMs(totalStart) });
    return result;
  }

  const chunks = chunk(entries, 6);
  let processed = 0;
  for (const batch of chunks) {
    const chunkStart = performance.now();
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
          const input = inputLookup.get(entry.threadId);
          result[entry.threadId] = {
            prompt: entry.prompt.trim(),
            suggestedAction: normalizeAction(entry.suggestedAction) || guessAction(input || {})
          };
        }
      }
      processed += batch.length;
      log('chunk complete', {
        batchSize: batch.length,
        processed,
        durationMs: elapsedMs(chunkStart)
      });
    } catch (err) {
      log('chunk failed', { error: (err as Error).message || err });
      console.error('Failed to generate chat primers', err);
    }
  }

  for (const entry of entries) {
    if (!result[entry.threadId]) {
      result[entry.threadId] = fallbackPrimer(entry);
    }
  }
  log('complete', { durationMs: elapsedMs(totalStart) });
  return result;
}

function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

function parsePrimerResponse(payload: string): { threadId: string; prompt: string; suggestedAction?: SuggestedAction }[] {
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

function fallbackPrimer(entry: ChatPrimerInput): PrimerOutput {
  const summaryBase = (entry.summary || entry.headline || entry.subject || 'an email that needs your call').trim();
  const normalizedSummary = summaryBase.replace(/\s+/g, ' ');
  const sender = entry.fromLine ? entry.fromLine.trim() : '';
  const context = sender ? `${normalizedSummary} from ${sender}` : normalizedSummary;
  const hasNextStep = entry.nextStep && entry.nextStep.toLowerCase() !== 'no action';
  const action = hasNextStep ? entry.nextStep!.trim() : '';
  const suggestedAction = normalizeAction(actionFromNextStep(entry.nextStep)) || 'more_info';
  const next = hasNextStep
    ? `I’d log a quick task for "${action}" so it doesn’t slip. Want me to do that?`
    : randomSuggestion(suggestedAction);
  return { prompt: `Heads up: looks like ${context}. ${next}`, suggestedAction };
}

function randomSuggestion(action: SuggestedAction) {
  if (action === 'archive') return 'Looks wrapped up—want me to archive it and keep you moving?';
  if (action === 'skip') return 'We can park this and move to the next email if you want.';
  if (action === 'create_task') return 'Want me to turn this into a task so it’s on your list?';
  return FALLBACK_SUGGESTIONS[Math.floor(Math.random() * FALLBACK_SUGGESTIONS.length)];
}

function createPrimerLogger(traceId?: string) {
  const prefix = traceId ? `[chatPrimer:${traceId}]` : '[chatPrimer]';
  return (message: string, extra?: Record<string, unknown>) => {
    if (extra && Object.keys(extra).length) {
      console.log(`${prefix} ${message}`, extra);
    } else {
      console.log(`${prefix} ${message}`);
    }
  };
}

function elapsedMs(start: number) {
  return Math.round((performance.now() - start) * 10) / 10;
}

function normalizeAction(value: unknown): SuggestedAction | null {
  const val = typeof value === 'string' ? value.trim().toLowerCase() : '';
  if (val === 'archive' || val === 'more_info' || val === 'create_task' || val === 'skip') return val;
  return null;
}

function guessAction(entry: { summary?: string; nextStep?: string }): SuggestedAction {
  const inferred = normalizeAction(actionFromNextStep(entry.nextStep));
  if (inferred) return inferred;
  const text = `${entry.summary || ''}`.toLowerCase();
  if (text.includes('schedule') || text.includes('follow up') || text.includes('deadline') || text.includes('due')) {
    return 'create_task';
  }
  if (text.includes('fyi') || text.includes('newsletter') || text.includes('update')) {
    return 'skip';
  }
  return 'more_info';
}

function actionFromNextStep(nextStep?: string | null): SuggestedAction | null {
  const normalized = (nextStep || '').trim().toLowerCase();
  if (!normalized) return null;
  if (normalized.includes('archive')) return 'archive';
  if (normalized.includes('task') || normalized.includes('remind') || normalized.includes('follow up') || normalized.includes('todo')) {
    return 'create_task';
  }
  return null;
}
