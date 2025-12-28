import { Prisma, type ActionFlow, type TranscriptMessage } from '@prisma/client';
import { prisma } from '../store/db.js';
import { AutoSummaryResult, generateAutoSummary, generateTaskDraft, TaskDraft } from '../llm/autoActions.js';

export type ActionType = 'archive' | 'create_task' | 'more_info' | 'skip' | 'external_action';
export type ActionState = 'suggested' | 'draft_ready' | 'editing' | 'executing' | 'completed' | 'failed';
export type TranscriptType =
  | 'must_know'
  | 'suggested_action'
  | 'draft_details'
  | 'inline_editor'
  | 'action_result';

export type ExternalActionLink = {
  label: string;
  url: string;
};

export type ExternalActionPayload = {
  steps: string;
  links: ExternalActionLink[];
};

export type TimelineMessage = {
  id: string;
  threadId: string;
  type: TranscriptType;
  content: string;
  payload: any;
  createdAt: string;
};

export type AutoSummaryContext = {
  userId: string;
  threadId: string;
  lastMessageId: string;
  subject: string;
  headline: string;
  summary: string;
  nextStep: string;
  participants: string[];
  transcript: string;
};

export async function ensureAutoSummaryCards(ctx: AutoSummaryContext) {
  const [existingFlow, existingMessages] = await Promise.all([
    prisma.actionFlow.findUnique({
      where: { userId_threadId: { userId: ctx.userId, threadId: ctx.threadId } }
    }),
    prisma.transcriptMessage.findMany({
      where: { userId: ctx.userId, threadId: ctx.threadId },
      orderBy: { createdAt: 'asc' }
    })
  ]);

  if (existingFlow && existingFlow.lastMessageId === ctx.lastMessageId) {
    if (existingFlow.actionType === 'skip' && existingFlow.state === 'completed') {
      return { flow: existingFlow, messages: serializeMessages(existingMessages) };
    }
    if (existingMessages.length) {
      return { flow: existingFlow, messages: serializeMessages(existingMessages) };
    }
  }

  const generated = await generateAutoSummary({
    subject: ctx.subject,
    headline: ctx.headline,
    summary: ctx.summary,
    nextStep: ctx.nextStep,
    participants: ctx.participants,
    transcript: ctx.transcript
  });

  const actionType: ActionType = generated.suggestedAction?.actionType || 'skip';
  const externalAction = generated.suggestedAction?.externalAction || null;
  const prompt = generated.suggestedAction?.userFacingPrompt
    || (actionType === 'external_action'
      ? 'This needs your attention outside the app. Here are the key links.'
      : 'Skip for now and move to the next email?');
  const mustKnow = generated.mustKnow || ctx.summary || ctx.headline || ctx.subject || 'New email.';

  const result = await prisma.$transaction(async tx => {
    await tx.transcriptMessage.deleteMany({ where: { userId: ctx.userId, threadId: ctx.threadId } });
    const flow = await tx.actionFlow.upsert({
      where: { userId_threadId: { userId: ctx.userId, threadId: ctx.threadId } },
      update: {
        actionType,
        state: 'suggested',
        draftPayload: null,
        lastMessageId: ctx.lastMessageId
      },
      create: {
        userId: ctx.userId,
        threadId: ctx.threadId,
        actionType,
        state: 'suggested',
        draftPayload: null,
        lastMessageId: ctx.lastMessageId
      }
    });

    const mustKnowMsg = await tx.transcriptMessage.create({
      data: {
        userId: ctx.userId,
        threadId: ctx.threadId,
        type: 'must_know',
        content: mustKnow,
        payload: null
      }
    });

    const suggestedMsg = await tx.transcriptMessage.create({
      data: {
        userId: ctx.userId,
        threadId: ctx.threadId,
        type: 'suggested_action',
        content: prompt,
        payload: packPayload({ actionType, externalAction })
      }
    });

    return { flow, messages: [mustKnowMsg, suggestedMsg] };
  });

  return {
    flow: result.flow,
    messages: serializeMessages(result.messages)
  };
}

export async function fetchTimeline(userId: string, threadId: string): Promise<TimelineMessage[]> {
  const messages = await prisma.transcriptMessage.findMany({
    where: { userId, threadId },
    orderBy: { createdAt: 'asc' }
  });
  return serializeMessages(messages);
}

export async function generateDraftDetails(ctx: {
  userId: string;
  threadId: string;
  subject: string;
  headline: string;
  summary: string;
  nextStep: string;
  transcript: string;
  lastMessageId?: string;
}) {
  const draft = await generateTaskDraft({
    subject: ctx.subject,
    headline: ctx.headline,
    summary: ctx.summary,
    nextStep: ctx.nextStep,
    transcript: ctx.transcript
  });
  const payload = sanitizeDraft(draft);
  const result = await prisma.$transaction(async tx => {
    const flow = await tx.actionFlow.upsert({
      where: { userId_threadId: { userId: ctx.userId, threadId: ctx.threadId } },
      update: {
        actionType: 'create_task',
        state: 'draft_ready',
        draftPayload: packPayload(payload),
        ...(ctx.lastMessageId ? { lastMessageId: ctx.lastMessageId } : {})
      },
      create: {
        userId: ctx.userId,
        threadId: ctx.threadId,
        actionType: 'create_task',
        state: 'draft_ready',
        draftPayload: packPayload(payload),
        lastMessageId: ctx.lastMessageId ?? null
      }
    });
    const message = await tx.transcriptMessage.create({
      data: {
        userId: ctx.userId,
        threadId: ctx.threadId,
        type: 'draft_details',
        content: draftSummary(payload),
        payload: packPayload(payload)
      }
    });
    return { flow, message };
  });
  return { flow: result.flow, message: serializeMessage(result.message) };
}

export async function openInlineEditor(userId: string, threadId: string) {
  const flow = await prisma.actionFlow.findUnique({
    where: { userId_threadId: { userId, threadId } }
  });
  if (!flow || !flow.draftPayload) throw new Error('Draft not ready to edit.');
  const payload = sanitizeDraft(unpackPayload(flow.draftPayload) as TaskDraft);
  const [updatedFlow, message] = await prisma.$transaction([
    prisma.actionFlow.update({
      where: { userId_threadId: { userId, threadId } },
      data: { state: 'editing', draftPayload: packPayload(payload) }
    }),
    prisma.transcriptMessage.create({
      data: {
        userId,
        threadId,
        type: 'inline_editor',
        content: 'Edit the task before creating.',
        payload: packPayload(payload)
      }
    })
  ]);
  return { flow: updatedFlow, message: serializeMessage(message) };
}

export async function saveEditedDraft(userId: string, threadId: string, draft: TaskDraft) {
  const payload = sanitizeDraft(draft);
  const result = await prisma.$transaction(async tx => {
    const flow = await tx.actionFlow.update({
      where: { userId_threadId: { userId, threadId } },
      data: {
        state: 'draft_ready',
        draftPayload: packPayload(payload),
        actionType: 'create_task'
      }
    });
    const message = await tx.transcriptMessage.create({
      data: {
        userId,
        threadId,
        type: 'draft_details',
        content: draftSummary(payload),
        payload: packPayload(payload)
      }
    });
    return { flow, message };
  });
  return { flow: result.flow, message: serializeMessage(result.message) };
}

export async function appendActionResult(userId: string, threadId: string, content: string, payload: any = null) {
  const message = await prisma.transcriptMessage.create({
    data: {
      userId,
      threadId,
      type: 'action_result',
      content: content || 'Done.',
      payload: packPayload(payload)
    }
  });
  return serializeMessage(message);
}

function sanitizeDraft(draft: TaskDraft): TaskDraft {
  const title = typeof draft?.title === 'string' ? draft.title.trim() : '';
  const notes = typeof draft?.notes === 'string' ? draft.notes.trim() : '';
  const dueDate = normalizeDate(draft?.dueDate);
  return {
    title: title || 'New task',
    notes: notes || '',
    dueDate
  };
}

function normalizeDate(raw: unknown): string | null {
  if (!raw) return null;
  const text = typeof raw === 'string' ? raw.trim() : '';
  if (!text) return null;
  const iso = text.length === 10 ? `${text}T23:59:00Z` : text;
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed.toISOString().slice(0, 10);
}

function draftSummary(draft: TaskDraft) {
  const bits = [`${draft.title || 'Task'}`];
  if (draft.dueDate) bits.push(`Due ${draft.dueDate}`);
  return bits.join(' — ');
}

function serializeMessages(messages: TranscriptMessage[]): TimelineMessage[] {
  return messages.map(serializeMessage);
}

export function serializeTimelineMessages(messages: TranscriptMessage[]): TimelineMessage[] {
  return serializeMessages(messages);
}

export function serializeMessage(message: TranscriptMessage): TimelineMessage {
  return {
    id: message.id,
    threadId: message.threadId,
    type: message.type as TranscriptType,
    content: message.content,
    payload: unpackPayload(message.payload),
    createdAt: message.createdAt.toISOString()
  };
}

function packPayload(value: any): string | null {
  if (value == null) return null;
  try {
    return JSON.stringify(value);
  } catch {
    return null;
  }
}

function unpackPayload(raw?: string | null) {
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}
