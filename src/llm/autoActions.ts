import OpenAI from 'openai';
import { type ActionType, type ExternalActionPayload } from '../actions/persistence.js';

export type SuggestedActionPayload = {
  actionType: ActionType;
  userFacingPrompt: string;
  externalAction?: ExternalActionPayload | null;
};

export type AutoSummaryResult = {
  mustKnow: string;
  suggestedAction: SuggestedActionPayload;
  suggestedActions: SuggestedActionPayload[];
};

export type TaskDraft = {
  title: string;
  notes: string;
  dueDate: string | null;
};

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

const AUTO_SUMMARY_PROMPT = `You summarize email threads for a busy professional and propose a short ordered sequence of actions.
Output STRICT JSON:
{"must_know":"<plain text essentials only>","suggested_actions":[{"action_type":"ARCHIVE|CREATE_TASK|MORE_INFO|REPLY|EXTERNAL_ACTION|UNSUBSCRIBE|NONE","userFacingPrompt":"<concise action prompt>","external_action":{"steps":"<1-2 sentences>","links":[{"label":"<human readable>","url":"https://..."}]}}]}
Rules:
- "must_know": concise, essential content only (no actions, no emojis, no markdown).
- "suggested_actions": 1-3 items, in the exact order the user should do them. Prefer ARCHIVE when no action is needed. Use NONE only for truly ignorable content.
- "suggested_actions[].userFacingPrompt": required when action_type is ARCHIVE/CREATE_TASK/MORE_INFO/REPLY/UNSUBSCRIBE. Keep it short, direct, and explicitly mention the action being suggested (e.g., for ARCHIVE say to archive). No markdown, no quotes, no emojis.
- "suggested_actions[].external_action": only include when action_type is EXTERNAL_ACTION.
- EXTERNAL_ACTION is for meaningful user actions outside the app (security alerts, verification, compliance, account lock risk, past-due notices, required RSVP/forms). Marketing/promotional CTAs like “buy now”, “upgrade”, “shop” must NOT trigger EXTERNAL_ACTION.
- CREATE_TASK is only for reminders when the email cannot be resolved immediately and the user should handle it later (future follow-ups, deadlines, planned check-ins). If there is a direct link to take action now (review budget, verify account, RSVP, submit form), prefer EXTERNAL_ACTION.
- REPLY is for straightforward responses the user can handle by email without external steps.
- UNSUBSCRIBE is only for promotional/bulk email that clearly provides a list-unsubscribe option.
- "suggested_actions[].external_action.steps": two sentences max. First sentence says why it matters. Second sentence says what to do.
- "suggested_actions[].external_action.links": 1-3 items. Prefer explicit account/security/action URLs from the email. Avoid generic homepage links.
- If action_type is NONE, still include it in suggested_actions with a brief userFacingPrompt.`;

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
  const suggestedActions = normalizeSuggestedActions(safe);
  if (!mustKnow || !suggestedActions.length) return null;
  return {
    mustKnow,
    suggestedAction: suggestedActions[0],
    suggestedActions
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
  const action: SuggestedActionPayload = { actionType: 'skip', userFacingPrompt: prompt };
  return { mustKnow, suggestedAction: action, suggestedActions: [action] };
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

function normalizeSuggestedActions(raw: any): SuggestedActionPayload[] {
  const fromArray = Array.isArray(raw?.suggested_actions || raw?.suggestedActions)
    ? (raw.suggested_actions || raw.suggestedActions)
    : null;
  const topLevelExternal = normalizeExternalAction(raw?.external_action || raw?.externalAction);

  const actions: SuggestedActionPayload[] = [];
  if (fromArray) {
    for (const entry of fromArray) {
      const normalized = normalizeSuggestedActionEntry(entry);
      if (normalized) actions.push(normalized);
    }
  }

  if (!actions.length) {
    const actionType = normalizeActionType(
      raw?.action_type ||
        raw?.actionType ||
        raw?.suggested_action?.actionType ||
        raw?.suggested_action?.action_type
    );
    const prompt = normalizeText(raw?.suggested_action?.userFacingPrompt || raw?.suggested_action?.prompt);
    const externalAction = actionType === 'external_action' ? topLevelExternal : null;
    const fallback = buildSuggestedAction(actionType, prompt, externalAction);
    if (fallback) actions.push(fallback);
  }

  const deduped = dedupeSuggestedActions(actions).slice(0, 3);
  const externalOnly = deduped.find(action => action.actionType === 'external_action');
  if (externalOnly) return [externalOnly];
  return deduped;
}

function normalizeSuggestedActionEntry(entry: any): SuggestedActionPayload | null {
  if (!entry || typeof entry !== 'object') return null;
  const actionType = normalizeActionType(entry.action_type || entry.actionType);
  const prompt = normalizeText(entry.userFacingPrompt || entry.prompt);
  const externalAction = actionType === 'external_action'
    ? normalizeExternalAction(entry.external_action || entry.externalAction)
    : null;
  return buildSuggestedAction(actionType, prompt, externalAction);
}

function buildSuggestedAction(
  actionType: ActionType | null,
  prompt: string,
  externalAction: ExternalActionPayload | null
): SuggestedActionPayload | null {
  if (!actionType) return null;
  let finalPrompt = prompt;
  if (!finalPrompt) {
    if (actionType === 'external_action') {
      finalPrompt = normalizeText(externalAction?.steps);
    } else {
      finalPrompt = actionPrompt(actionType);
    }
  }
  if (!finalPrompt) return null;
  if (actionType === 'archive' && !/\barchive\b/i.test(finalPrompt)) {
    finalPrompt = actionPrompt('archive');
  }
  if (actionType === 'external_action' && !externalAction?.steps) return null;
  return { actionType, userFacingPrompt: finalPrompt, externalAction };
}

function dedupeSuggestedActions(actions: SuggestedActionPayload[]) {
  const seen = new Set<ActionType>();
  return actions.filter(action => {
    if (seen.has(action.actionType)) return false;
    seen.add(action.actionType);
    return true;
  });
}

function normalizeText(value: unknown) {
  const text = typeof value === 'string' ? value.trim() : '';
  if (!text) return '';
  return text.replace(/\s+/g, ' ');
}

function normalizeActionType(value: unknown): ActionType | null {
  const text = typeof value === 'string' ? value.trim().toLowerCase() : '';
  if (text === 'none') return 'archive';
  if (text === 'archive' || text === 'create_task' || text === 'more_info' || text === 'skip' || text === 'external_action' || text === 'reply' || text === 'unsubscribe') {
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
  if (action === 'external_action') return 'This needs your attention outside the app. Want the key links?';
  if (action === 'reply') return 'Draft a reply to the sender so you can respond quickly?';
  if (action === 'unsubscribe') return 'Unsubscribe from this sender to stop these promos?';
  return 'Skip for now and move to the next email?';
}

function normalizeExternalAction(raw: unknown): ExternalActionPayload | null {
  if (!raw || typeof raw !== 'object') return null;
  const candidate = raw as { steps?: unknown; links?: unknown };
  const steps = normalizeText(candidate.steps);
  const links = normalizeExternalLinks(candidate.links);
  if (!steps) return null;
  return { steps, links };
}

function normalizeExternalLinks(raw: unknown): ExternalActionPayload['links'] {
  if (!Array.isArray(raw)) return [];
  const links = raw.map(item => normalizeExternalLink(item)).filter(Boolean) as ExternalActionPayload['links'];
  return links.slice(0, 3);
}

function normalizeExternalLink(raw: unknown) {
  if (!raw || typeof raw !== 'object') return null;
  const link = raw as { label?: unknown; url?: unknown };
  const label = normalizeText(link.label);
  const url = normalizeUrl(link.url);
  if (!url) return null;
  return { label: label || url, url };
}

function normalizeUrl(raw: unknown): string {
  const text = typeof raw === 'string' ? raw.trim() : '';
  if (!text) return '';
  try {
    const parsed = new URL(text);
    if (parsed.protocol === 'http:' || parsed.protocol === 'https:') return parsed.toString();
  } catch {
    return '';
  }
  return '';
}
