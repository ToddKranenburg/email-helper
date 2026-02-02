import { Router, Request, Response } from 'express';
import { ingestInbox } from '../gmail/fetch.js';
import { ensureThreadSummary } from '../gmail/summary.js';
import { loadPriorityQueue, type PriorityEntry } from '../actions/priorityQueue.js';
import { prisma } from '../store/db.js';
import fs from 'node:fs/promises';
import path from 'node:path';
import { chatAboutEmail, MAX_CHAT_TURNS, type ChatTurn } from '../llm/secretaryChat.js';
import { gmailClient } from '../gmail/client.js';
import { getAuthedClient, getMissingGmailScopes, MissingScopeError } from '../auth/google.js';
import { normalizeBody } from '../gmail/normalize.js';
import { GaxiosError } from 'gaxios';
import { performance } from 'node:perf_hooks';
import crypto from 'node:crypto';
import type { Summary, ThreadIndex, ActionFlow } from '@prisma/client';
import { classifyIntent, detectArchiveIntent } from '../llm/intentClassifier.js';
import { createGoogleTask, normalizeDueDate } from '../tasks/createTask.js';
import { generateGuidedReplyDraft, generateReplyDraft } from '../llm/replyDraft.js';
import {
  ensureAutoSummaryCards,
  fetchTimeline,
  generateDraftDetails,
  openInlineEditor,
  saveEditedDraft,
  appendActionResult,
  type ActionType,
  type ActionState,
  type TimelineMessage
} from '../actions/persistence.js';

export const router = Router();
const PAGE_SIZE = 20;
const PRIORITY_LIMIT = 6;
const ingestStatus = new Map<string, { status: 'idle' | 'running' | 'done' | 'error'; updatedAt: number; error?: string }>();
const REVIEW_PROMPT = 'Give me a concise, easy-to-digest rundown of this email. Hit the key points, any asks or decisions, deadlines, and suggested follow-ups in short bullets. Keep it scannable.';
const SCOPE_UPGRADE_PATH = '/auth/google?upgrade=1';
const REPLY_DRAFT_MIN_CONFIDENCE = 0.75;

router.use(async (req, _res, next) => {
  const sessionData = req.session as any;
  if (sessionData?.user?.id && sessionData?.user?.email) {
    try {
      await ensureUserRecord(sessionData);
    } catch (err) {
      console.warn('Failed to update user activity', err);
    }
  }
  next();
});

function clearGoogleSession(sessionData: any) {
  if (!sessionData) return;
  delete sessionData.googleTokens;
}

function ensureScopesForPage(sessionData: any, res: Response) {
  const missingScopes = getMissingGmailScopes(sessionData?.googleTokens);
  if (!missingScopes.length) return false;
  clearGoogleSession(sessionData);
  res.redirect(SCOPE_UPGRADE_PATH);
  return true;
}

async function ensureRefreshTokenForPage(sessionData: any, res: Response) {
  const userId = sessionData?.user?.id;
  if (!userId) return false;
  const token = await prisma.googleToken.findUnique({
    where: { userId },
    select: { refreshToken: true }
  });
  if (token?.refreshToken) return false;
  clearGoogleSession(sessionData);
  res.redirect('/auth/google?upgrade=1&reason=missing_refresh');
  return true;
}

function ensureScopesForApi(sessionData: any, res: Response) {
  const missingScopes = getMissingGmailScopes(sessionData?.googleTokens);
  if (!missingScopes.length) return false;
  clearGoogleSession(sessionData);
  res.status(403).json({
    error: 'Google permissions changed. Please reconnect your Google account to continue.',
    missingScopes,
    reconnectUrl: '/auth/google'
  });
  return true;
}

async function ensureRefreshTokenForApi(sessionData: any, res: Response) {
  const userId = sessionData?.user?.id;
  if (!userId) return false;
  const token = await prisma.googleToken.findUnique({
    where: { userId },
    select: { refreshToken: true }
  });
  if (token?.refreshToken) return false;
  clearGoogleSession(sessionData);
  res.status(403).json({
    error: 'Google needs a refresh token to keep syncing in the background. Please reconnect your Google account.',
    reconnectUrl: '/auth/google?upgrade=1&reason=missing_refresh'
  });
  return true;
}

function handleMissingScopeError(err: unknown, sessionData: any, res: Response) {
  if (!(err instanceof MissingScopeError)) return false;
  clearGoogleSession(sessionData);
  res.status(403).json({
    error: 'Google permissions changed. Please reconnect your Google account to continue.',
    missingScopes: err.missingScopes,
    reconnectUrl: '/auth/google'
  });
  return true;
}

function handleInsufficientScopeFromGaxios(err: unknown, sessionData: any, res: Response) {
  const gaxios = err instanceof GaxiosError ? err : undefined;
  if (!gaxios) return false;
  const header = gaxios.response?.headers?.['www-authenticate'];
  const authHeader = Array.isArray(header) ? header.join(' ') : header;
  if (typeof authHeader !== 'string' || !authHeader.includes('insufficient_scope')) return false;
  const scopeMatch = authHeader.match(/scope="([^"]+)"/);
  const scopeList = scopeMatch?.[1] ? scopeMatch[1].split(/\s+/).filter(Boolean) : [];
  clearGoogleSession(sessionData);
  res.status(403).json({
    error: 'Google permissions changed. Please reconnect your Google account to continue.',
    missingScopes: scopeList,
    reconnectUrl: '/auth/google'
  });
  return true;
}

router.get('/', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  const tokens = sessionData.googleTokens;
  if (tokens?.access_token && sessionData.user?.id) {
    // ✅ Already authorized: go straight to dashboard
    if (ensureScopesForPage(sessionData, res)) return;
    return res.redirect('/dashboard');
  }
  // ❌ Not authorized yet: show connect link
  res.send(`<a href="/auth/google">Connect Gmail</a>`);
});

type PageMeta = {
  totalItems: number;
  pageSize: number;
  currentPage: number;
  hasMore: boolean;
  nextPage: number | null;
};

type SummaryWithThreadIndex = Summary & { threadIndex: ThreadIndex | null };
type ThreadListItem = { thread: ThreadIndex; summary: Summary | null };
type ThreadContext = { summary: SummaryWithThreadIndex; transcript: string; participants: string[] };
type SecretaryThread = {
  threadId: string;
  messageId: string;
  headline: string;
  from: string;
  subject: string;
  summary: string;
  nextStep: string;
  link: string;
  category: string;
  receivedAt: string;
  convo: string;
  participants: string[];
  actionFlow: ActionFlow | null;
  timeline: TimelineMessage[];
};

router.get('/dashboard', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) return res.redirect('/auth/google');
  if (ensureScopesForPage(sessionData, res)) return;
  if (await ensureRefreshTokenForPage(sessionData, res)) return;
  const userId = sessionData.user.id;
  const traceId = createTraceId();
  const log = scopedLogger(`dashboard:${traceId}`);
  const routeStart = performance.now();
  log('start', { userId });
  await ensureUserRecord(sessionData);

  const requestedPage = typeof req.query.page === 'string' ? parseInt(req.query.page, 10) : 1;
  const listStart = performance.now();
  const pageData = await loadPage(userId, requestedPage);
  log('loaded summaries', { durationMs: elapsedMs(listStart), returned: pageData.items.length, page: pageData.currentPage });

  // Decide whether to auto-ingest AFTER rendering (first-time/empty state).
  const hasSummaries = pageData.totalItems > 0;
  if (hasSummaries) {
    sessionData.skipAutoIngest = true;
  }
  const autoIngest = !hasSummaries && !sessionData.skipAutoIngest;

  const templateStart = performance.now();
  const layout = await fs.readFile(path.join(process.cwd(), 'src/web/views/layout.html'), 'utf8');
  const body = await fs.readFile(path.join(process.cwd(), 'src/web/views/dashboard.html'), 'utf8');
  log('templates read', { durationMs: elapsedMs(templateStart) });
  const threads = await buildThreadsPayload(userId, pageData.items);
  const priorityQueue = await loadPriorityQueue(userId, { limit: PRIORITY_LIMIT });
  const [prioritizedCount, totalCount, latestCompletedBatch] = await Promise.all([
    prisma.threadIndex.count({ where: { userId, inPrimaryInbox: true, priorityScore: { not: null } } }),
    prisma.threadIndex.count({ where: { userId, inPrimaryInbox: true } }),
    prisma.prioritizationBatch.findFirst({
      where: { userId, status: 'completed', finishedAt: { not: null } },
      orderBy: { finishedAt: 'desc' }
    })
  ]);
  const priorityThreads = priorityQueue.items.length
    ? await buildThreadsPayload(userId, priorityQueue.items)
    : [];
  const mergedThreads = mergeThreads(threads, priorityThreads);

  // Inject a small flag the client script can read to auto-trigger ingest
  const renderStart = performance.now();
  const pageMeta: PageMeta = {
    totalItems: pageData.totalItems,
    pageSize: PAGE_SIZE,
    currentPage: pageData.currentPage,
    hasMore: pageData.hasMore,
    nextPage: pageData.nextPage
  };
  const withFlag = `${render(
    body,
    mergedThreads,
    pageMeta,
    priorityQueue.priority,
    { prioritizedCount, totalCount, batchFinishedAt: latestCompletedBatch?.finishedAt ?? null }
  )}
  <script>window.AUTO_INGEST = ${autoIngest ? 'true' : 'false'};</script>`;

  const html = layout.replace('<!--CONTENT-->', withFlag);
  log('html rendered', { durationMs: elapsedMs(renderStart) });
  res.send(html);
  log('completed', { durationMs: elapsedMs(routeStart) });
});

router.post('/ingest', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).send('auth first');
  }
  if (ensureScopesForApi(sessionData, res)) return;
  if (await ensureRefreshTokenForApi(sessionData, res)) return;
  const userId = sessionData.user.id;
  await ensureUserRecord(sessionData);
  sessionData.skipAutoIngest = true;

  const existing = ingestStatus.get(userId);
  if (existing?.status === 'running') {
    return res.json({ status: 'running' });
  }

  const sessionSnapshot = cloneSessionForIngest(sessionData);
  if (!sessionSnapshot) {
    return res.status(400).json({ status: 'error', message: 'Missing session data for ingest.' });
  }

  markIngestStatus(userId, 'running');
  triggerBackgroundIngest(sessionSnapshot, userId);
  res.json({ status: 'running' });
});

router.get('/ingest/status', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  const userId = sessionData?.user?.id;
  if (!sessionData?.googleTokens || !userId) {
    return res.status(401).json({ status: 'unauthorized' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  if (await ensureRefreshTokenForApi(sessionData, res)) return;
  const state = ingestStatus.get(userId) || { status: 'idle', updatedAt: Date.now() };
  res.json({ status: state.status, updatedAt: state.updatedAt, error: state.error });
});

router.post('/secretary/chat', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;

  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  const question = typeof req.body?.question === 'string' ? req.body.question.trim() : '';
  const history = normalizeHistory(req.body?.history);

  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });
  if (!question) return res.status(400).json({ error: 'Ask a specific question.' });

  const existingUserTurns = history.filter(turn => turn.role === 'user').length;
  if (existingUserTurns >= MAX_CHAT_TURNS) {
    return res.status(429).json({ error: 'Chat limit reached for this thread.' });
  }

  const context = await loadThreadContext(sessionData.user.id, threadId, req);
  if (!context) return res.status(404).json({ error: 'Email summary not found.' });
  if (!context.transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  try {
    const reply = await chatAboutEmail({
      subject: context.summary.threadIndex?.subject || '',
      headline: context.summary.headline,
      tldr: context.summary.tldr,
      nextStep: context.summary.nextStep,
      participants: context.participants,
      transcript: context.transcript,
      history,
      question
    });
    res.json({ reply });
  } catch (err) {
    console.error('secretary chat failed', err);
    res.status(500).json({ error: 'Unable to chat about this email right now. Please try again.' });
  }
});

router.post('/secretary/auto-summarize', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  const forceFresh = Boolean(req.body?.fresh);
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });
  const context = await loadThreadContext(sessionData.user.id, threadId, req);
  if (!context) return res.status(404).json({ error: 'Email summary not found.' });
  if (!context.transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  try {
    const ensured = await ensureAutoSummaryCards({
      userId: sessionData.user.id,
      threadId,
      lastMessageId: context.summary.lastMsgId,
      subject: context.summary.threadIndex?.subject || '',
      headline: context.summary.headline,
      summary: context.summary.tldr,
      nextStep: context.summary.nextStep,
      participants: context.participants,
      transcript: context.transcript
    }, { forceFresh });
    const timeline = await fetchTimeline(sessionData.user.id, threadId);
    return res.json({ flow: ensured.flow, timeline });
  } catch (err) {
    console.error('auto summarize failed', err);
    return res.status(500).json({ error: 'Unable to prepare this email right now. Please try again.' });
  }
});

router.post('/secretary/review', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });

  const context = await loadThreadContext(sessionData.user.id, threadId, req);
  if (!context) return res.status(404).json({ error: 'Email summary not found.' });
  if (!context.transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  try {
    const review = await chatAboutEmail({
      subject: context.summary.threadIndex?.subject || '',
      headline: context.summary.headline,
      tldr: context.summary.tldr,
      nextStep: context.summary.nextStep,
      participants: context.participants,
      transcript: context.transcript,
      history: [],
      question: REVIEW_PROMPT
    });
    return res.json({ review });
  } catch (err) {
    console.error('secretary review failed', err);
    return res.status(500).json({ error: 'Unable to review this email right now. Please try again.' });
  }
});

router.post('/secretary/reply-draft', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });

  const context = await loadThreadContext(sessionData.user.id, threadId, req);
  if (!context) return res.status(404).json({ error: 'Email summary not found.' });
  if (!context.transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  try {
    const draft = await generateReplyDraft({
      subject: context.summary.Thread?.subject || '',
      headline: context.summary.headline,
      summary: context.summary.tldr,
      nextStep: context.summary.nextStep,
      participants: context.participants,
      transcript: context.transcript,
      fromLine: formatSender(context.summary.Thread)
    });
    const eligible = Boolean(draft.safeToDraft && draft.body && draft.confidence >= REPLY_DRAFT_MIN_CONFIDENCE);
    const signoffUser = await resolveUserForSignoff(sessionData);
    const body = eligible ? appendReplySignoff(draft.body, signoffUser) : '';
    return res.json({
      body,
      confidence: draft.confidence,
      safe: draft.safeToDraft,
      suggested: eligible,
      reason: draft.reason
    });
  } catch (err) {
    console.error('reply draft failed', err);
    return res.status(500).json({ error: 'Unable to draft a reply right now.' });
  }
});

router.post('/secretary/reply-intent-draft', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  const text = typeof req.body?.text === 'string' ? req.body.text.trim() : '';
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });
  if (!text) return res.status(400).json({ error: 'Provide reply guidance.' });

  const context = await loadThreadContext(sessionData.user.id, threadId, req);
  if (!context) return res.status(404).json({ error: 'Email summary not found.' });
  if (!context.transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  try {
    const draft = await generateGuidedReplyDraft({
      subject: context.summary.Thread?.subject || '',
      headline: context.summary.headline,
      summary: context.summary.tldr,
      nextStep: context.summary.nextStep,
      participants: context.participants,
      transcript: context.transcript,
      fromLine: formatSender(context.summary.Thread),
      userInstruction: text
    });
    const eligible = Boolean(draft.safeToDraft && draft.body && draft.confidence >= REPLY_DRAFT_MIN_CONFIDENCE);
    const signoffUser = await resolveUserForSignoff(sessionData);
    const body = eligible ? appendReplySignoff(draft.body, signoffUser) : '';
    return res.json({
      body,
      confidence: draft.confidence,
      safe: draft.safeToDraft,
      suggested: eligible,
      reason: draft.reason
    });
  } catch (err) {
    console.error('reply intent draft failed', err);
    return res.status(500).json({ error: 'Unable to draft a reply right now.' });
  }
});

router.post('/secretary/archive-intent', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const text = typeof req.body?.text === 'string' ? req.body.text.trim() : '';
  if (!text) return res.status(400).json({ error: 'Provide text to evaluate.' });

  try {
    const decision = await detectArchiveIntent(text);
    return res.json(decision);
  } catch (err) {
    console.error('archive intent detection failed', err);
    return res.status(500).json({ error: 'Unable to evaluate archive intent right now.' });
  }
});

router.post('/secretary/intent', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const text = typeof req.body?.text === 'string' ? req.body.text.trim() : '';
  if (!text) return res.status(400).json({ error: 'Provide text to evaluate.' });

  try {
    const decision = await classifyIntent(text);
    return res.json(decision);
  } catch (err) {
    console.error('secretary intent detection failed', err);
    return res.status(500).json({ error: 'Unable to evaluate intent right now.' });
  }
});

router.post('/secretary/action/draft', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });
  const mode = req.body?.mode === 'edit' ? 'edit' : req.body?.mode === 'save' ? 'save' : 'generate';
  const context = await loadThreadContext(sessionData.user.id, threadId, req);
  if (!context) return res.status(404).json({ error: 'Email summary not found.' });
  if (!context.transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  try {
    let flow: ActionFlow | null = null;
    if (mode === 'edit') {
      const edited = await openInlineEditor(sessionData.user.id, threadId);
      flow = edited.flow;
    } else if (mode === 'save') {
      const draft = normalizeDraftInput(req.body?.draft);
      const saved = await saveEditedDraft(sessionData.user.id, threadId, draft);
      flow = saved.flow;
    } else {
      const generated = await generateDraftDetails({
        userId: sessionData.user.id,
        threadId,
        subject: context.summary.threadIndex?.subject || '',
        headline: context.summary.headline,
        summary: context.summary.tldr,
        nextStep: context.summary.nextStep,
        transcript: context.transcript,
        lastMessageId: context.summary.lastMsgId
      });
      flow = generated.flow;
    }
    const timeline = await fetchTimeline(sessionData.user.id, threadId);
    return res.json({ flow, timeline });
  } catch (err) {
    console.error('action draft failed', err);
    return res.status(500).json({ error: 'Unable to prepare the draft right now. Please try again.' });
  }
});

router.post('/secretary/action/execute', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });
  const actionType = normalizeActionTypeInput(req.body?.actionType || req.body?.action_type);
  if (!actionType) return res.status(400).json({ error: 'Unsupported action type.' });

  const context = await loadThreadContext(sessionData.user.id, threadId, req);
  if (!context) return res.status(404).json({ error: 'Email summary not found.' });
  if (!context.transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  try {
    if (actionType === 'archive') {
      const auth = getAuthedClient(sessionData);
      const gmail = gmailClient(auth);
      await gmail.users.threads.modify({
        userId: 'me',
        id: threadId,
        requestBody: { removeLabelIds: ['INBOX'] }
      });
      await prisma.summary.deleteMany({ where: { userId: sessionData.user.id, threadId } });
      await prisma.threadIndex.update({
        where: { threadId_userId: { threadId, userId: sessionData.user.id } },
        data: { inPrimaryInbox: false }
      });
      const flow = await prisma.actionFlow.upsert({
        where: { userId_threadId: { userId: sessionData.user.id, threadId } },
        update: {
          actionType,
          state: 'completed',
          draftPayload: null,
          lastMessageId: context.summary.lastMsgId
        },
        create: {
          userId: sessionData.user.id,
          threadId,
          actionType,
          state: 'completed',
          draftPayload: null,
          lastMessageId: context.summary.lastMsgId
        }
      });
      await appendActionResult(sessionData.user.id, threadId, 'Archived in Gmail.');
      const timeline = await fetchTimeline(sessionData.user.id, threadId);
      return res.json({ status: 'archived', flow, timeline });
    }

    if (actionType === 'skip') {
      const flow = await prisma.actionFlow.upsert({
        where: { userId_threadId: { userId: sessionData.user.id, threadId } },
        update: {
          actionType,
          state: 'completed',
          draftPayload: null,
          lastMessageId: context.summary.lastMsgId
        },
        create: {
          userId: sessionData.user.id,
          threadId,
          actionType,
          state: 'completed',
          draftPayload: null,
          lastMessageId: context.summary.lastMsgId
        }
      });
      await appendActionResult(sessionData.user.id, threadId, 'Skipped. No new suggestions until this thread changes.');
      const timeline = await fetchTimeline(sessionData.user.id, threadId);
      return res.json({ status: 'skipped', flow, timeline });
    }

    if (actionType === 'more_info') {
      const reply = await chatAboutEmail({
        subject: context.summary.threadIndex?.subject || '',
        headline: context.summary.headline,
        tldr: context.summary.tldr,
        nextStep: context.summary.nextStep,
        participants: context.participants,
        transcript: context.transcript,
        history: [],
        question: 'Share extra context and clarifications about this email. List any open questions or missing details.'
      });
      const flow = await prisma.actionFlow.upsert({
        where: { userId_threadId: { userId: sessionData.user.id, threadId } },
        update: {
          actionType,
          state: 'suggested',
          lastMessageId: context.summary.lastMsgId
        },
        create: {
          userId: sessionData.user.id,
          threadId,
          actionType,
          state: 'suggested',
          lastMessageId: context.summary.lastMsgId,
          draftPayload: null
        }
      });
      await appendActionResult(sessionData.user.id, threadId, reply);
      const timeline = await fetchTimeline(sessionData.user.id, threadId);
      return res.json({ status: 'ok', flow, timeline });
    }

    if (actionType === 'reply') {
      const body = normalizeReplyBody(req.body?.draft?.body ?? req.body?.body);
      if (!body) {
        return res.status(400).json({ error: 'Write a reply before sending.' });
      }

      const auth = getAuthedClient(sessionData);
      const gmail = gmailClient(auth);
      const replyMeta = await fetchReplyMetadata(gmail, context.summary.lastMsgId);
      const replyTarget = selectReplyTarget(replyMeta, context.summary.Thread, context.participants, sessionData.user.email);
      if (!replyTarget.to) {
        return res.status(400).json({ error: 'Unable to identify a sender to reply to.' });
      }
      const normalizedUser = normalizeEmail(sessionData.user.email);
      if (replyTarget.email && normalizedUser && replyTarget.email.toLowerCase() === normalizedUser) {
        return res.status(400).json({ error: 'Unable to find a sender to reply to in this thread.' });
      }
      const subject = buildReplySubject(replyMeta.subject || context.summary.Thread?.subject || '');
      const references = mergeReferences(replyMeta.references, replyMeta.messageId);
      const raw = buildReplyMessage({
        to: replyTarget.to,
        subject,
        body,
        inReplyTo: replyMeta.messageId,
        references
      });

      const executingFlow = await prisma.actionFlow.upsert({
        where: { userId_threadId: { userId: sessionData.user.id, threadId } },
        update: {
          actionType,
          state: 'executing',
          draftPayload: null,
          lastMessageId: context.summary.lastMsgId
        },
        create: {
          userId: sessionData.user.id,
          threadId,
          actionType,
          state: 'executing',
          draftPayload: null,
          lastMessageId: context.summary.lastMsgId
        }
      });

      try {
        await gmail.users.messages.send({
          userId: 'me',
          requestBody: {
            raw,
            threadId
          }
        });
        const completedFlow = await prisma.actionFlow.update({
          where: { id: executingFlow.id },
          data: { state: 'completed' }
        });
        const recipientLabel = replyTarget.label || replyTarget.to;
        await appendActionResult(
          sessionData.user.id,
          threadId,
          `✅ Reply sent to ${recipientLabel || 'the sender'}.`
        );
        const timeline = await fetchTimeline(sessionData.user.id, threadId);
        return res.json({ status: 'sent', flow: completedFlow, timeline });
      } catch (err) {
        if (handleMissingScopeError(err, sessionData, res)) return;
        if (handleInsufficientScopeFromGaxios(err, sessionData, res)) return;
        await prisma.actionFlow.update({
          where: { id: executingFlow.id },
          data: { state: 'failed' }
        });
        console.error('Failed to send reply', err);
        const message = err instanceof Error ? err.message : 'Unable to send that reply.';
        await appendActionResult(sessionData.user.id, threadId, `Reply failed: ${message}`);
        const timeline = await fetchTimeline(sessionData.user.id, threadId);
        return res.status(500).json({ error: 'Unable to send that reply right now.', timeline });
      }
    }

    if (actionType === 'create_task') {
      const draftInput = normalizeDraftInput(req.body?.draft);
      const existingFlow = await prisma.actionFlow.findUnique({
        where: { userId_threadId: { userId: sessionData.user.id, threadId } }
      });
      const draft = selectDraftPayload(draftInput, existingFlow?.draftPayload, context);
      if (!draft.title.trim()) {
        return res.status(400).json({ error: 'Add a task title before creating.' });
      }

      const executingFlow = await prisma.actionFlow.upsert({
        where: { userId_threadId: { userId: sessionData.user.id, threadId } },
        update: {
          actionType,
          state: 'executing',
          draftPayload: encodeDraft(draft),
          lastMessageId: context.summary.lastMsgId
        },
        create: {
          userId: sessionData.user.id,
          threadId,
          actionType,
          state: 'executing',
          draftPayload: encodeDraft(draft),
          lastMessageId: context.summary.lastMsgId
        }
      });

      try {
        const auth = getAuthedClient(sessionData);
        const task = await createGoogleTask(auth, {
          title: draft.title,
          notes: draft.notes,
          due: normalizeDueDate(draft.dueDate || undefined)
        });
        const listId = extractTaskListId(task);
        const taskUrl = buildTasksLink(task.id, listId);
        const completedFlow = await prisma.actionFlow.update({
          where: { id: executingFlow.id },
          data: { state: 'completed', draftPayload: encodeDraft(draft) }
        });
        await appendActionResult(sessionData.user.id, threadId, buildTaskResultMessage({
          title: task.title || draft.title,
          due: task.due ?? draft.dueDate,
          url: taskUrl
        }), {
          taskId: task.id,
          taskUrl
        });
        const timeline = await fetchTimeline(sessionData.user.id, threadId);
        return res.json({
          status: 'created',
          taskId: task.id,
          taskUrl,
          flow: completedFlow,
          timeline
        });
      } catch (err) {
        if (handleMissingScopeError(err, sessionData, res)) return;
        if (handleInsufficientScopeFromGaxios(err, sessionData, res)) return;
        await prisma.actionFlow.update({
          where: { id: executingFlow.id },
          data: { state: 'failed' }
        });
        console.error('Failed to create Google Task', err);
        const message = err instanceof Error ? err.message : 'Unable to create that task right now.';
        await appendActionResult(sessionData.user.id, threadId, `Task creation failed: ${message}`);
        const timeline = await fetchTimeline(sessionData.user.id, threadId);
        return res.status(500).json({ error: 'Unable to create that task right now.', timeline });
      }
    }

    return res.status(400).json({ error: 'Unsupported action type.' });
  } catch (err) {
    if (handleMissingScopeError(err, sessionData, res)) return;
    const gaxios = err instanceof GaxiosError ? err : undefined;
    const reason = gaxios?.response?.status === 403
      ? 'Google blocked this request. Please reconnect Google and try again.'
      : 'Unable to process that action right now. Please try again.';
    console.error('action execute failed', err);
    return res.status(500).json({ error: reason });
  }
});

router.get('/api/threads', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const userId = sessionData.user.id;
  const traceId = createTraceId();
  const log = scopedLogger(`threads:${traceId}`);
  const requestedPage = typeof req.query.page === 'string' ? parseInt(req.query.page, 10) : 1;
  const targetPage = Number.isFinite(requestedPage) ? Math.max(requestedPage, 1) : 1;

  try {
    const existing = await prisma.threadIndex.count({ where: { userId, inPrimaryInbox: true } });
    if (!existing) {
      log('ingesting initial index', { requestedPage: targetPage });
      await ingestInbox(sessionData);
    }

    const pageData = await loadPage(userId, targetPage);
    log('page ready', { page: pageData.currentPage, returned: pageData.items.length });
    const threads = await buildThreadsPayload(userId, pageData.items);
    return res.json({
      threads,
      meta: {
        totalItems: pageData.totalItems,
        pageSize: PAGE_SIZE,
        currentPage: pageData.currentPage,
        hasMore: pageData.hasMore,
        nextPage: pageData.nextPage
      }
    });
  } catch (err) {
    if (handleMissingScopeError(err, sessionData, res)) return;
    log('failed to load page', { error: err instanceof Error ? err.message : String(err) });
    return res.status(500).json({ error: 'Unable to load more emails right now. Please try again.' });
  }
});

router.get('/api/priority', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  if (await ensureRefreshTokenForApi(sessionData, res)) return;
  const userId = sessionData.user.id;
  try {
    const priorityQueue = await loadPriorityQueue(userId, { limit: PRIORITY_LIMIT });
    const priorityThreads = priorityQueue.items.length
      ? await buildThreadsPayload(userId, priorityQueue.items)
      : [];
    const [prioritizedCount, totalCount, latestCompletedBatch] = await Promise.all([
      prisma.threadIndex.count({ where: { userId, inPrimaryInbox: true, priorityScore: { not: null } } }),
      prisma.threadIndex.count({ where: { userId, inPrimaryInbox: true } }),
      prisma.prioritizationBatch.findFirst({
        where: { userId, status: 'completed', finishedAt: { not: null } },
        orderBy: { finishedAt: 'desc' }
      })
    ]);
    return res.json({
      priority: priorityQueue.priority,
      threads: priorityThreads,
      progress: { prioritizedCount, totalCount },
      batchFinishedAt: latestCompletedBatch?.finishedAt
        ? latestCompletedBatch.finishedAt.toISOString()
        : null
    });
  } catch (err) {
    if (handleMissingScopeError(err, sessionData, res)) return;
    return res.status(500).json({ error: 'Unable to load priority queue right now. Please try again.' });
  }
});

router.post('/api/archive', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  if (ensureScopesForApi(sessionData, res)) return;
  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });

  const thread = await prisma.threadIndex.findUnique({
    where: { threadId_userId: { threadId, userId: sessionData.user.id } }
  });
  if (!thread) return res.status(404).json({ error: 'Email not found in your queue.' });

  const traceId = createTraceId();
  const log = scopedLogger(`archive:${traceId}`);

  try {
    const auth = getAuthedClient(sessionData);
    const gmail = gmailClient(auth);
    await gmail.users.threads.modify({
      userId: 'me',
      id: threadId,
      requestBody: {
        removeLabelIds: ['INBOX']
      }
    });
    await prisma.summary.deleteMany({ where: { userId: sessionData.user.id, threadId } });
    await prisma.threadIndex.update({
      where: { threadId_userId: { threadId, userId: sessionData.user.id } },
      data: { inPrimaryInbox: false }
    });
    log('archived thread', { threadId, userId: sessionData.user.id });
    return res.json({ status: 'archived' });
  } catch (err) {
    if (handleMissingScopeError(err, sessionData, res)) return;
    const gaxios = err instanceof GaxiosError ? err : undefined;
    const reason = gaxios?.response?.status === 403
      ? 'Google blocked the archive request. Please reconnect Google and try again.'
      : 'Unable to archive this email right now. Please try again.';
    log('archive failed', { threadId, error: err instanceof Error ? err.message : String(err) });
    return res.status(500).json({ error: reason });
  }
});

router.post('/api/tasks', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }

  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  const messageId = typeof req.body?.messageId === 'string' ? req.body.messageId.trim() : '';
  const title = typeof req.body?.title === 'string' ? req.body.title.trim() : '';
  const notes = typeof req.body?.notes === 'string' ? req.body.notes.trim() : '';
  const due = typeof req.body?.due === 'string' ? req.body.due.trim() : '';

  if (!title) return res.status(400).json({ error: 'Add a task title before saving.' });
  if (!threadId && !messageId) {
    return res.status(400).json({ error: 'Missing email identifier for this task.' });
  }

  const summary = await prisma.summary.findFirst({
    where: {
      userId: sessionData.user.id,
      ...(messageId ? { lastMsgId: messageId } : {}),
      ...(threadId ? { threadId } : {})
    }
  });
  if (!summary) {
    return res.status(404).json({ error: 'Email not found in your queue.' });
  }

  try {
    const auth = getAuthedClient(sessionData);
    const task = await createGoogleTask(auth, { title, notes, due });
    const listId = extractTaskListId(task);
    return res.json({
      status: 'created',
      taskId: task.id,
      taskUrl: buildTasksLink(task.id, listId),
      due: task.due ?? null,
      title: task.title ?? title
    });
  } catch (err) {
    const gaxios = err instanceof GaxiosError ? err : undefined;
    const reason = gaxios?.response?.status === 403
      ? 'Google Tasks access is missing. Please reconnect Google and try again.'
      : 'Unable to create that task right now. Please try again.';
    console.error('Failed to create Google Task', err);
    return res.status(500).json({ error: reason });
  }
});

function buildTasksLink(taskId?: string | null, listId?: string | null) {
  if (!taskId) return '';
  const list = listId || '@default';
  const base = 'https://tasks.google.com/embed/';
  const query = new URLSearchParams({
    list,
    task: taskId,
    origin: 'https://mail.google.com'
  });
  return `${base}?${query.toString()}`;
}

function extractTaskListId(task: { selfLink?: string | null }) {
  const selfLink = typeof task?.selfLink === 'string' ? task.selfLink : '';
  if (!selfLink) return '';
  const match = selfLink.match(/\/lists\/([^/]+)\/tasks\/[^/]+/);
  return match?.[1] || '';
}

async function loadPage(userId: string, requestedPage: number, opts: { assumeMore?: boolean } = {}) {
  return loadPageWithOpts(userId, requestedPage, opts);
}

async function loadPageWithOpts(userId: string, requestedPage: number, opts: { assumeMore?: boolean }) {
  const currentPage = Number.isFinite(requestedPage) ? Math.max(requestedPage, 1) : 1;
  const skip = (currentPage - 1) * PAGE_SIZE;
  const [threads, totalItems] = await Promise.all([
    prisma.threadIndex.findMany({
      where: { userId, inPrimaryInbox: true },
      orderBy: [{ lastMessageDate: 'desc' }],
      skip,
      take: PAGE_SIZE + 1 // fetch one extra so we can keep rendering full pages
    }),
    prisma.threadIndex.count({ where: { userId, inPrimaryInbox: true } })
  ]);
  const hasExtraRecord = threads.length > PAGE_SIZE;
  const pageThreads = hasExtraRecord ? threads.slice(0, PAGE_SIZE) : threads;
  const threadIds = pageThreads.map(thread => thread.threadId);
  const summaries = threadIds.length
    ? await prisma.summary.findMany({ where: { userId, threadId: { in: threadIds } } })
    : [];
  const summaryMap = new Map(summaries.map(summary => [summary.threadId, summary]));
  const items = pageThreads.map(thread => ({ thread, summary: summaryMap.get(thread.threadId) || null }));
  const hasMore = totalItems > skip + items.length || Boolean(opts.assumeMore);
  const baseTotalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
  const totalPages = hasMore ? Math.max(baseTotalPages, currentPage + 1) : baseTotalPages;
  const nextPage = hasMore ? currentPage + 1 : null;
  return { items, totalItems, totalPages, currentPage, hasMore, nextPage };
}

async function buildThreadsPayload(_userId: string, items: ThreadListItem[]): Promise<SecretaryThread[]> {
  const threads: SecretaryThread[] = [];
  for (const item of items) {
    const summary = item.summary;
    const participants = parseParticipants(item.thread.participants);
    const emailTs = item.thread.lastMessageDate ? new Date(item.thread.lastMessageDate) : new Date();

    threads.push({
      threadId: item.thread.threadId,
      messageId: summary?.lastMsgId || item.thread.lastMessageId || '',
      headline: summary?.headline || '',
      from: formatSender(item.thread),
      subject: item.thread.subject || '(no subject)',
      summary: summary?.tldr || item.thread.snippet || '',
      nextStep: summary?.nextStep || '',
      link: item.thread.threadId ? `https://mail.google.com/mail/u/0/#all/${encodeURIComponent(item.thread.threadId)}` : '',
      category: summary?.category || '',
      receivedAt: emailTs.toISOString(),
      convo: summary?.convoText || '',
      participants,
      actionFlow: null,
      timeline: []
    });
  }

  return threads;
}


function mergeThreads(primary: SecretaryThread[], extra: SecretaryThread[]) {
  if (!extra.length) return primary;
  const seen = new Set(primary.map(thread => thread.threadId));
  const merged = primary.slice();
  extra.forEach(thread => {
    if (seen.has(thread.threadId)) return;
    seen.add(thread.threadId);
    merged.push(thread);
  });
  return merged;
}

function emojiForCategory(cat: string): string {
  const c = (cat || '').toLowerCase();
  if (c.startsWith('marketing')) return '🏷️';
  if (c.startsWith('personal event')) return '📅';
  if (c.startsWith('billing')) return '💳';
  if (c.startsWith('introduction')) return '🤝';
  if (c.startsWith('catch up')) return '👋';
  if (c.startsWith('editorial')) return '📰';
  if (c.startsWith('personal request')) return '🙏';
  if (c.startsWith('fyi')) return 'ℹ️';
  return '📎';
}

function render(
  tpl: string,
  items: SecretaryThread[],
  meta: PageMeta,
  priority: PriorityEntry[],
  priorityProgress: { prioritizedCount: number; totalCount: number; batchFinishedAt: Date | null }
) {
  const secretaryScript = renderSecretaryAssistant(items, meta, priority, priorityProgress);
  return `${tpl}\n${secretaryScript}`;
}

function renderSecretaryAssistant(
  items: SecretaryThread[],
  meta: PageMeta,
  priority: PriorityEntry[],
  priorityProgress: { prioritizedCount: number; totalCount: number; batchFinishedAt: Date | null }
) {
  const payload = safeJson({
    threads: items,
    priority,
    priorityProgress: {
      prioritizedCount: priorityProgress.prioritizedCount,
      totalCount: priorityProgress.totalCount
    },
    priorityBatchFinishedAt: priorityProgress.batchFinishedAt
      ? priorityProgress.batchFinishedAt.toISOString()
      : null,
    maxTurns: MAX_CHAT_TURNS,
    totalItems: meta.totalItems,
    pageSize: meta.pageSize,
    hasMore: meta.hasMore,
    currentPage: meta.currentPage,
    nextPage: meta.nextPage
  }); // consumed by src/web/public/secretary.js
  return `
<script id="secretary-bootstrap">window.SECRETARY_BOOTSTRAP = ${payload};</script>
<script src="https://cdn.jsdelivr.net/npm/marked@12.0.2/lib/marked.umd.min.js" defer></script>
<script src="https://cdn.jsdelivr.net/npm/linkify-it@5.0.0/dist/linkify-it.min.js" defer></script>
<script src="https://cdn.jsdelivr.net/npm/dompurify@3.1.6/dist/purify.min.js" defer></script>
<script src="/secretary.js" defer></script>
`;
}

function escapeHtml(s: string) {
  return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]!));
}

function safeJson(value: unknown) {
  return JSON.stringify(value).replace(/</g, '\\u003c');
}

function formatSender(thread?: { fromName?: string | null; fromEmail?: string | null } | null) {
  const name = thread?.fromName ? String(thread.fromName).trim() : '';
  const email = thread?.fromEmail ? String(thread.fromEmail).trim() : '';
  if (!name && !email) return '';
  if (name && email) return `${name} <${email}>`;
  if (name) return name;
  return `<${email}>`;
}

function formatFriendlyDate(raw?: string | null) {
  if (!raw) return '';
  const iso = raw.length === 10 ? `${raw}T00:00:00Z` : raw;
  const parsed = new Date(iso);
  if (Number.isNaN(parsed.getTime())) return '';
  const opts: Intl.DateTimeFormatOptions = { month: 'short', day: 'numeric' };
  if (parsed.getFullYear() !== new Date().getFullYear()) {
    opts.year = 'numeric';
  }
  return parsed.toLocaleDateString(undefined, opts);
}

function normalizeHistory(raw: any): ChatTurn[] {
  if (!Array.isArray(raw)) return [];
  const entries: ChatTurn[] = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue;
    const role = item.role === 'assistant' ? 'assistant' : item.role === 'user' ? 'user' : null;
    if (!role || typeof item.content !== 'string') continue;
    const content = item.content.trim();
    if (!content) continue;
    entries.push({ role, content });
  }
  return entries.slice(-MAX_CHAT_TURNS * 2);
}

async function fetchTranscript(threadId: string, req: Request) {
  try {
    const auth = getAuthedClient((req as any).session);
    const gmail = gmailClient(auth);
    const thread = await gmail.users.threads.get({ userId: 'me', id: threadId });
    const messages = (thread.data.messages || []).slice(-3);
    return messages
      .map(msg => normalizeBody(msg.payload))
      .filter(Boolean)
      .reverse()
      .join('\\n\\n---\\n\\n');
  } catch (err) {
    console.error('Failed to fetch Gmail transcript', err);
    return '';
  }
}

async function loadThreadContext(userId: string, threadId: string, req: Request): Promise<ThreadContext | null> {
  let summary = await prisma.summary.findFirst({
    where: { userId, threadId },
    include: { threadIndex: true }
  }) as SummaryWithThreadIndex | null;
  if (!summary) {
    try {
      const auth = getAuthedClient((req as any).session);
      summary = await ensureThreadSummary(auth, userId, threadId);
    } catch (err) {
      console.error('Unable to generate summary for thread', err);
    }
  }
  if (!summary) return null;

  let transcript = summary.convoText || '';
  if (!transcript) {
    transcript = await fetchTranscript(threadId, req);
    if (transcript) {
      await prisma.summary.update({ where: { id: summary.id }, data: { convoText: transcript } });
      summary.convoText = transcript;
    }
  }
  const participants = parseParticipants(summary.threadIndex?.participants);
  return { summary, transcript, participants };
}

function parseParticipants(raw?: string | null): string[] {
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.map(value => String(value || '').trim()).filter(Boolean);
    }
  } catch {
    // ignore
  }
  return [];
}

function normalizeActionTypeInput(raw: unknown): ActionType | null {
  const value = typeof raw === 'string' ? raw.trim().toLowerCase() : '';
  if (value === 'archive' || value === 'create_task' || value === 'more_info' || value === 'skip' || value === 'reply') {
    return value as ActionType;
  }
  return null;
}

function normalizeDraftInput(raw: any): { title: string; notes: string; dueDate: string | null } {
  const title = typeof raw?.title === 'string' ? raw.title.trim() : '';
  const notes = typeof raw?.notes === 'string' ? raw.notes.trim() : '';
  const due = typeof raw?.dueDate === 'string'
    ? raw.dueDate.trim()
    : typeof raw?.due === 'string'
      ? raw.due.trim()
      : '';
  return {
    title,
    notes,
    dueDate: due || null
  };
}

function normalizeReplyBody(raw: unknown): string {
  if (typeof raw !== 'string') return '';
  return raw.replace(/\r\n/g, '\n').trim();
}

function deriveFirstName(name?: string | null, email?: string | null): string {
  const cleanName = typeof name === 'string' ? name.replace(/["<>]/g, '').trim() : '';
  if (cleanName) {
    const first = cleanName.split(/\s+/)[0];
    if (first) return first;
  }
  const cleanEmail = typeof email === 'string' ? email.trim() : '';
  if (!cleanEmail || !cleanEmail.includes('@')) return '';
  const local = cleanEmail.split('@')[0] || '';
  const chunk = local.split(/[._-]+/).filter(Boolean)[0] || '';
  if (!chunk) return '';
  return chunk.charAt(0).toUpperCase() + chunk.slice(1);
}

function appendReplySignoff(body: string, user?: { name?: string | null; email?: string | null } | null): string {
  const trimmed = String(body || '').trim();
  if (!trimmed) return '';
  const firstName = deriveFirstName(user?.name ?? null, user?.email ?? null);
  if (!firstName) return trimmed;
  return `${trimmed}\n\n${firstName}`;
}

async function resolveUserForSignoff(sessionData: any): Promise<{ name?: string | null; email?: string | null } | null> {
  const sessionUser = sessionData?.user;
  const hasName = typeof sessionUser?.name === 'string' && sessionUser.name.trim();
  const hasEmail = typeof sessionUser?.email === 'string' && sessionUser.email.trim();
  if (hasName || hasEmail) return { name: sessionUser?.name ?? null, email: sessionUser?.email ?? null };
  const userId = typeof sessionUser?.id === 'string' ? sessionUser.id : '';
  if (!userId) return null;
  const dbUser = await prisma.user.findUnique({ where: { id: userId } });
  if (!dbUser) return null;
  return { name: dbUser.name ?? null, email: dbUser.email ?? null };
}

async function fetchReplyMetadata(gmail: ReturnType<typeof gmailClient>, messageId: string) {
  if (!messageId) {
    return { subject: '', replyTo: '', from: '', messageId: '', references: '' };
  }
  const message = await gmail.users.messages.get({
    userId: 'me',
    id: messageId,
    format: 'metadata',
    metadataHeaders: ['Subject', 'From', 'Reply-To', 'Message-ID', 'References']
  });
  const headers = message.data.payload?.headers || [];
  return {
    subject: findHeader(headers, 'Subject'),
    replyTo: findHeader(headers, 'Reply-To'),
    from: findHeader(headers, 'From'),
    messageId: findHeader(headers, 'Message-ID'),
    references: findHeader(headers, 'References')
  };
}

function findHeader(headers: Array<{ name?: string | null; value?: string | null }>, name: string) {
  const target = name.toLowerCase();
  const header = headers.find(item => (item?.name || '').toLowerCase() === target);
  return header?.value || '';
}

function selectReplyTarget(
  meta: { replyTo?: string; from?: string },
  thread?: { fromEmail?: string | null; fromName?: string | null } | null,
  participants: string[] = [],
  userEmail?: string | null
) {
  const parsed = parseAddress(meta.replyTo || meta.from || '');
  let email = parsed.email;
  let name = parsed.name;
  const fallbackEmail = thread?.fromEmail ? String(thread.fromEmail).trim() : '';
  const fallbackName = thread?.fromName ? String(thread.fromName).trim() : '';
  if (!email && fallbackEmail) {
    email = fallbackEmail;
    name = name || fallbackName;
  }
  const normalizedUser = normalizeEmail(userEmail);
  if (email && normalizedUser && email.toLowerCase() === normalizedUser) {
    const alternate = selectParticipantRecipient(participants, normalizedUser);
    if (alternate?.email) {
      email = alternate.email;
      name = alternate.name;
    }
  }
  const to = email ? formatAddress(name, email) : '';
  const label = name || email;
  return { to, label, email };
}

function parseAddress(raw: string) {
  const value = String(raw || '').trim();
  if (!value) return { email: '', name: '' };
  const angleMatch = value.match(/<([^>]+)>/);
  if (angleMatch) {
    const email = angleMatch[1].trim();
    const name = value.replace(angleMatch[0], '').trim().replace(/^"|"$/g, '');
    return { email, name };
  }
  const emailMatch = value.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  if (emailMatch) {
    const email = emailMatch[0].trim();
    const name = value.replace(emailMatch[0], '').trim().replace(/^"|"$/g, '');
    return { email, name };
  }
  return { email: '', name: value.replace(/^"|"$/g, '') };
}

function selectParticipantRecipient(participants: string[], userEmail: string) {
  if (!Array.isArray(participants)) return null;
  for (const participant of participants) {
    const parsed = parseAddress(participant);
    if (parsed.email && parsed.email.toLowerCase() !== userEmail) {
      return parsed;
    }
  }
  return null;
}

function normalizeEmail(raw?: string | null) {
  return raw ? String(raw).trim().toLowerCase() : '';
}

function formatAddress(name: string, email: string) {
  const safeEmail = sanitizeHeaderValue(email);
  const safeName = sanitizeHeaderValue(name);
  return safeName ? `${safeName} <${safeEmail}>` : safeEmail;
}

function buildReplySubject(subject: string) {
  const clean = sanitizeHeaderValue(subject);
  if (!clean) return 'Re:';
  if (/^re:/i.test(clean)) return clean;
  return `Re: ${clean}`;
}

function mergeReferences(references: string, messageId: string) {
  const ref = sanitizeHeaderValue(references);
  const msg = sanitizeHeaderValue(messageId);
  if (!msg) return ref;
  if (!ref) return msg;
  if (ref.includes(msg)) return ref;
  return `${ref} ${msg}`;
}

function buildReplyMessage(input: {
  to: string;
  subject: string;
  body: string;
  inReplyTo?: string;
  references?: string;
}) {
  const headers = [
    `To: ${sanitizeHeaderValue(input.to)}`,
    `Subject: ${sanitizeHeaderValue(input.subject)}`,
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset="UTF-8"',
    'Content-Transfer-Encoding: 7bit'
  ];
  const inReplyTo = sanitizeHeaderValue(input.inReplyTo || '');
  if (inReplyTo) headers.splice(2, 0, `In-Reply-To: ${inReplyTo}`);
  const references = sanitizeHeaderValue(input.references || '');
  if (references) headers.splice(3, 0, `References: ${references}`);
  const body = normalizeReplyBody(input.body).replace(/\n/g, '\r\n');
  const raw = `${headers.join('\r\n')}\r\n\r\n${body}`;
  return encodeBase64Url(raw);
}

function encodeBase64Url(value: string) {
  return Buffer.from(value)
    .toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=+$/g, '');
}

function sanitizeHeaderValue(value: string) {
  return String(value || '').replace(/[\r\n]+/g, ' ').trim();
}

function selectDraftPayload(
  inputDraft: { title: string; notes: string; dueDate: string | null },
  existingPayload: any,
  context: ThreadContext
) {
  const parsedExisting = parseDraftPayload(existingPayload);
  const existing = typeof parsedExisting === 'object' && parsedExisting !== null ? parsedExisting : {};
  const title = inputDraft.title || String(existing.title || context.summary.headline || context.summary.tldr || context.summary.threadIndex?.subject || 'New task');
  const notes = inputDraft.notes ?? String(existing.notes || context.summary.tldr || '');
  const dueDate = inputDraft.dueDate ?? (typeof existing.dueDate === 'string' ? existing.dueDate : null);
  return {
    title: title.trim(),
    notes: notes.trim(),
    dueDate: dueDate || null
  };
}

function buildTaskResultMessage(input: { title?: string | null; due?: string | null; url?: string | null }) {
  const bits = ['✅ Task created'];
  const title = typeof input.title === 'string' ? input.title.trim() : '';
  if (title) bits.push(title);
  const dueLabel = formatFriendlyDate(input.due);
  if (dueLabel) bits.push(`Due ${dueLabel}`);
  if (input.url) bits.push(`[Open in Google Tasks](${input.url})`);
  return bits.join(' — ');
}

function encodeDraft(draft: { title: string; notes: string; dueDate: string | null }) {
  try {
    return JSON.stringify(draft);
  } catch {
    return null;
  }
}

function parseDraftPayload(raw: any) {
  if (!raw) return {};
  if (typeof raw === 'string') {
    try {
      return JSON.parse(raw);
    } catch {
      return {};
    }
  }
  return typeof raw === 'object' ? raw : {};
}

async function ensureUserRecord(sessionData: any) {
  const user = sessionData?.user;
  if (!user?.id || !user?.email) return;
  await prisma.user.upsert({
    where: { id: user.id },
    update: {
      email: user.email,
      name: user.name ?? undefined,
      picture: user.picture ?? undefined,
      lastActiveAt: new Date()
    },
    create: {
      id: user.id,
      email: user.email,
      name: user.name ?? null,
      picture: user.picture ?? null,
      lastActiveAt: new Date()
    }
  });
}

function scopedLogger(scope: string) {
  return (message: string, extra?: Record<string, unknown>) => {
    if (extra && Object.keys(extra).length) {
      console.log(`[${scope}] ${message}`, extra);
    } else {
      console.log(`[${scope}] ${message}`);
    }
  };
}

function elapsedMs(start: number) {
  return Math.round((performance.now() - start) * 10) / 10;
}

function createTraceId() {
  if (typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function triggerBackgroundIngest(sessionData: any, userId: string) {
  ingestInbox(sessionData)
    .then(() => {
      markIngestStatus(userId, 'done');
    })
    .catch((err: unknown) => {
      if (err instanceof MissingScopeError) {
        markIngestStatus(userId, 'error', 'Google permissions changed. Please reconnect your Google account.');
        return;
      }
      const gaxios = err instanceof GaxiosError ? err : undefined;
      const message = gaxios?.response?.status === 403
        ? 'Gmail refused to share inbox data. Please reconnect your Google account.'
        : 'Unable to sync your Gmail inbox right now. Please try again.';
      markIngestStatus(userId, 'error', message);
    });
}

function markIngestStatus(userId: string, status: 'idle' | 'running' | 'done' | 'error', error?: string) {
  ingestStatus.set(userId, { status, updatedAt: Date.now(), error });
  if (status === 'done' || status === 'error') {
    setTimeout(() => ingestStatus.delete(userId), 5 * 60 * 1000); // clean up after a bit
  }
}

function cloneSessionForIngest(sessionData: any) {
  if (!sessionData?.googleTokens || !sessionData?.user?.id) return null;
  return {
    googleTokens: { ...(sessionData.googleTokens || {}) },
    user: { ...(sessionData.user || {}) }
  };
}
