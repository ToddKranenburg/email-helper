import OpenAI from 'openai';
import { type ActionType } from '../actions/persistence.js';

export type SuggestedActionPayload = {
  actionType: ActionType;
  userFacingPrompt: string;
};

export type AutoSummaryResult = {
  mustKnow: string;
  suggestedAction: SuggestedActionPayload;
};

export type TaskDraft = {
  title: string;
  notes: string;
  dueDate: string | null;
};

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

const AUTO_SUMMARY_PROMPT = `You summarize email threads for a busy professional and propose ONE next action.
Output STRICT JSON:
{"must_know":"<plain text essentials only>","suggested_action":{"actionType":"archive|create_task|more_info|skip","userFacingPrompt":"<concise action prompt>"}}
Rules:
- "must_know": concise, essential content only (no actions, no emojis, no markdown).
- "suggested_action": exactly one of archive | create_task | more_info | skip. Keep "userFacingPrompt" short, direct, and grounded in the thread. No markdown, no quotes, no emojis.`;

const TASK_DRAFT_PROMPT = `You draft Google Tasks from an email thread.
Output STRICT JSON: {"title":"...","notes":"...","dueDate":"YYYY-MM-DD or null"}
Rules:
- Title: crisp and specific to the email's ask.
- Notes: 2-4 short lines with key details; include senders/links/dates only if useful.
- dueDate: ISO date (YYYY-MM-DD) when clear, otherwise null.
- No markdown, no bullets or numbering.`;

export async function generateAutoSummary(input: {
  subject: string;
  headline: string;
  summary: string;
  nextStep: string;
  participants: string[];
  transcript: string;
}): Promise<AutoSummaryResult> {
  const context = buildContext(input);
  if (!openai) {
    return fallbackAutoSummary(input);
  }
  try {
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      temperature: 0,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: AUTO_SUMMARY_PROMPT },
        { role: 'user', content: context }
      ]
    });
    const raw = completion.choices[0]?.message?.content || '';
    const parsed = parseAutoSummary(raw);
    if (parsed) return parsed;
    console.error('auto summary parse failed', raw);
  } catch (err) {
    console.error('auto summary generation failed', err);
  }
  return fallbackAutoSummary(input);
}

export async function generateTaskDraft(input: {
  subject: string;
  headline: string;
  summary: string;
  nextStep: string;
  participants?: string[];
  transcript: string;
}): Promise<TaskDraft> {
  const context = buildContext(input);
  if (!openai) return fallbackDraft(input);
  try {
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      temperature: 0.2,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: TASK_DRAFT_PROMPT },
        { role: 'user', content: context }
      ]
    });
    const raw = completion.choices[0]?.message?.content || '';
    const parsed = parseDraft(raw);
    if (parsed) return parsed;
  } catch (err) {
    console.error('task draft generation failed', err);
  }
  return fallbackDraft(input);
}

function buildContext(input: {
  subject: string;
  headline: string;
  summary: string;
  nextStep: string;
  participants?: string[];
  transcript: string;
}) {
  const participants = input.participants?.length ? input.participants.join(', ') : 'Unknown participants';
  const trimmedTranscript = input.transcript?.length > 6000
    ? input.transcript.slice(-6000)
    : input.transcript || '(no transcript provided)';
  return `Subject: ${input.subject || '(no subject)'}
Headline: ${input.headline || '(no headline)'}
Summary: ${input.summary || '(no summary)'}
Next step: ${input.nextStep || 'No action specified'}
Participants: ${participants}

Thread (oldest→newest):
${trimmedTranscript}`;
}

function parseAutoSummary(payload: string): AutoSummaryResult | null {
  const safe = extractJson(payload);
  if (!safe) return null;
  const mustKnow = normalizeText(safe.must_know || safe.mustKnow || safe.mustknow);
  const actionType = normalizeActionType(safe?.suggested_action?.actionType || safe?.suggested_action?.action_type);
  const prompt = normalizeText(safe?.suggested_action?.userFacingPrompt || safe?.suggested_action?.prompt);
  if (!mustKnow || !actionType || !prompt) return null;
  return {
    mustKnow,
    suggestedAction: { actionType, userFacingPrompt: prompt }
  };
}

function parseDraft(payload: string): TaskDraft | null {
  const safe = extractJson(payload);
  if (!safe) return null;
  const title = normalizeText(safe.title);
  const notes = normalizeText(safe.notes);
  const due = normalizeDueDate(safe.dueDate);
  if (!title) return null;
  return { title, notes: notes || '', dueDate: due };
}

function fallbackAutoSummary(input: {
  subject: string;
  headline: string;
  summary: string;
  nextStep: string;
  participants: string[];
}): AutoSummaryResult {
  const mustKnow = normalizeText(input.summary || input.headline || input.subject || 'New email in your inbox.');
  const prompt = actionPrompt('skip');
  return { mustKnow, suggestedAction: { actionType: 'skip', userFacingPrompt: prompt } };
}

function fallbackDraft(input: { subject: string; headline: string; summary: string; nextStep: string }): TaskDraft {
  const base = input.headline || input.subject || 'New task';
  const title = normalizeText(base);
  const notes = normalizeText(input.summary || input.nextStep || 'Captured from email thread.');
  return {
    title: title || 'New task',
    notes,
    dueDate: null
  };
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

function normalizeText(value: unknown) {
  const text = typeof value === 'string' ? value.trim() : '';
  if (!text) return '';
  return text.replace(/\s+/g, ' ');
}

function normalizeActionType(value: unknown): ActionType | null {
  const text = typeof value === 'string' ? value.trim().toLowerCase() : '';
  if (text === 'archive' || text === 'create_task' || text === 'more_info' || text === 'skip') {
    return text as ActionType;
  }
  return null;
}

function normalizeDueDate(value: unknown): string | null {
  if (value == null) return null;
  const text = typeof value === 'string' ? value.trim() : '';
  if (!text) return null;
  // Accept YYYY-MM-DD or ISO strings
  const iso = text.length === 10 ? `${text}T23:59:00Z` : text;
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed.toISOString().slice(0, 10);
}

function guessAction(nextStep: string, summary: string): ActionType {
  const next = (nextStep || '').toLowerCase();
  if (next.includes('archive')) return 'archive';
  if (next.includes('task') || next.includes('remind') || next.includes('follow up') || next.includes('follow-up')) {
    return 'create_task';
  }
  const text = `${summary || ''}`.toLowerCase();
  if (text.includes('fyi') || text.includes('newsletter') || text.includes('update')) return 'skip';
  return 'more_info';
}

function actionPrompt(action: ActionType) {
  if (action === 'archive') return 'Archive this thread to keep your inbox clear?';
  if (action === 'create_task') return 'Draft a quick task so this stays on your radar?';
  if (action === 'more_info') return 'Want me to pull more context or clarifications?';
  return 'Skip for now and move to the next email?';
}
