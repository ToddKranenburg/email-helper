import { Router, Request, Response } from 'express';
import { ingestInbox } from '../gmail/fetch.js';
import { prisma } from '../store/db.js';
import { generateChatPrimers, type ChatPrimerInput } from '../llm/chatPrimer.js';
import fs from 'node:fs/promises';
import path from 'node:path';
import { chatAboutEmail, MAX_CHAT_TURNS, type ChatTurn } from '../llm/secretaryChat.js';
import { gmailClient } from '../gmail/client.js';
import { getAuthedClient } from '../auth/google.js';
import { normalizeBody } from '../gmail/normalize.js';
import { GaxiosError } from 'gaxios';
import { performance } from 'node:perf_hooks';
import crypto from 'node:crypto';
import type { Summary, Thread } from '@prisma/client';

export const router = Router();
const PAGE_SIZE = 20;
const PRIMER_BACKGROUND_BATCH = 6;
const primerCache = new Map<string, string>();
const primerPending = new Map<string, Promise<string>>();
const ingestStatus = new Map<string, { status: 'idle' | 'running' | 'done' | 'error'; updatedAt: number; error?: string }>();

router.get('/', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  const tokens = sessionData.googleTokens;
  if (tokens?.access_token && sessionData.user?.id) {
    // ✅ Already authorized: go straight to dashboard
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

type SummaryWithThread = Summary & { Thread: Thread | null };
type DecoratedSummary = SummaryWithThread & { chatPrimer?: string };

router.get('/dashboard', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) return res.redirect('/auth/google');
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
  const primerInputs = pageData.items.map(buildPrimerInputFromSummary);
  const decorated = decorateSummariesWithPrimers(userId, pageData.items);

  // Inject a small flag the client script can read to auto-trigger ingest
  const renderStart = performance.now();
  const pageMeta: PageMeta = {
    totalItems: pageData.totalItems,
    pageSize: PAGE_SIZE,
    currentPage: pageData.currentPage,
    hasMore: pageData.hasMore,
    nextPage: pageData.nextPage
  };
  const withFlag = `${render(body, decorated, pageMeta)}
  <script>window.AUTO_INGEST = ${autoIngest ? 'true' : 'false'};</script>`;

  const html = layout.replace('<!--CONTENT-->', withFlag);
  log('html rendered', { durationMs: elapsedMs(renderStart) });
  res.send(html);
  log('completed', { durationMs: elapsedMs(routeStart) });
  queuePrimerPrefetch(userId, primerInputs, traceId);
});

router.post('/ingest', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).send('auth first');
  }
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
  const state = ingestStatus.get(userId) || { status: 'idle', updatedAt: Date.now() };
  res.json({ status: state.status, updatedAt: state.updatedAt, error: state.error });
});

router.post('/secretary/chat', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }

  const threadId = typeof req.body?.threadId === 'string' ? req.body.threadId.trim() : '';
  const question = typeof req.body?.question === 'string' ? req.body.question.trim() : '';
  const history = normalizeHistory(req.body?.history);

  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });
  if (!question) return res.status(400).json({ error: 'Ask a specific question.' });

  const existingUserTurns = history.filter(turn => turn.role === 'user').length;
  if (existingUserTurns >= MAX_CHAT_TURNS) {
    return res.status(429).json({ error: 'Chat limit reached for this thread.' });
  }

  const summary = await prisma.summary.findFirst({
    where: { userId: sessionData.user.id, threadId },
    include: { Thread: true }
  });

  if (!summary) return res.status(404).json({ error: 'Email summary not found.' });

  let transcript = summary.convoText || '';
  if (!transcript) {
    transcript = await fetchTranscript(threadId, req);
    if (transcript) {
      await prisma.summary.update({ where: { id: summary.id }, data: { convoText: transcript } });
    }
  }
  if (!transcript) return res.status(400).json({ error: 'Unable to load that email thread. Try re-ingesting your inbox.' });

  const participants = parseParticipants(summary.Thread?.participants);

  try {
    const reply = await chatAboutEmail({
      subject: summary.Thread?.subject || '',
      headline: summary.headline,
      tldr: summary.tldr,
      nextStep: summary.nextStep,
      participants,
      transcript,
      history,
      question
    });
    res.json({ reply });
  } catch (err) {
    console.error('secretary chat failed', err);
    res.status(500).json({ error: 'Unable to chat about this email right now. Please try again.' });
  }
});

router.get('/secretary/primer/:threadId', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  const userId = sessionData.user.id;
  const rawId = typeof req.params.threadId === 'string' ? req.params.threadId : '';
  const threadId = rawId.trim();
  if (!threadId) return res.status(400).json({ error: 'Missing thread id.' });

  const cached = getCachedPrimer(userId, threadId);
  if (cached) {
    return res.json({ primer: cached, status: 'ready' });
  }

  const pending = primerPending.get(primerCacheKey(userId, threadId));
  if (pending) {
    const primer = await pending.catch(() => '');
    if (primer) return res.json({ primer, status: 'ready' });
    return res.status(202).json({ primer: '', status: 'pending' });
  }

  const summary = await prisma.summary.findFirst({
    where: { userId, threadId },
    include: { Thread: true }
  });
  if (!summary) return res.status(404).json({ error: 'Email summary not found.' });

  const traceId = createTraceId();
  const input = buildPrimerInputFromSummary(summary as SummaryWithThread);
  const primer = await fetchPrimerForThread(userId, input, traceId);
  if (primer) {
    return res.json({ primer, status: 'ready' });
  }
  return res.status(202).json({ primer: '', status: 'pending' });
});

router.get('/api/threads', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) {
    return res.status(401).json({ error: 'Authenticate with Google first.' });
  }
  const userId = sessionData.user.id;
  const traceId = createTraceId();
  const log = scopedLogger(`threads:${traceId}`);
  const requestedPage = typeof req.query.page === 'string' ? parseInt(req.query.page, 10) : 1;

  try {
    const pageData = await loadPage(userId, requestedPage);
    log('page ready', { page: pageData.currentPage, returned: pageData.items.length });
    const decorated = decorateSummariesWithPrimers(userId, pageData.items);
    queuePrimerPrefetch(userId, pageData.items.map(buildPrimerInputFromSummary), traceId);
    return res.json({
      threads: summariesToThreads(decorated),
      meta: {
        totalItems: pageData.totalItems,
        pageSize: PAGE_SIZE,
        currentPage: pageData.currentPage,
        hasMore: pageData.hasMore,
        nextPage: pageData.nextPage
      }
    });
  } catch (err) {
    log('failed to load page', { error: err instanceof Error ? err.message : String(err) });
    return res.status(500).json({ error: 'Unable to load more emails right now. Please try again.' });
  }
});

async function loadPage(userId: string, requestedPage: number) {
  const currentPage = Number.isFinite(requestedPage) ? Math.max(requestedPage, 1) : 1;
  const skip = (currentPage - 1) * PAGE_SIZE;
  const results = await prisma.summary.findMany({
    where: { userId },
    include: { Thread: true },
    orderBy: [
      { Thread: { lastMessageTs: 'desc' } },
      { createdAt: 'desc' }
    ],
    skip,
    take: PAGE_SIZE + 1 // fetch one extra to infer "has more" without full count
  }) as SummaryWithThread[];
  const hasMore = results.length > PAGE_SIZE;
  const items = hasMore ? results.slice(0, PAGE_SIZE) : results;
  const totalItems = skip + items.length + (hasMore ? 1 : 0); // lower bound: we know there is at least one more when hasMore
  const totalPages = hasMore ? currentPage + 1 : currentPage;
  const nextPage = hasMore ? currentPage + 1 : null;
  return { items, totalItems, totalPages, currentPage, hasMore, nextPage };
}

function decorateSummariesWithPrimers(userId: string, items: SummaryWithThread[]): DecoratedSummary[] {
  return items.map(item => ({
    ...item,
    chatPrimer: getCachedPrimer(userId, item.threadId) || ''
  }));
}

function summariesToThreads(items: DecoratedSummary[]): SecretaryEmail[] {
  return items.map(x => {
    const emailTs = x.Thread?.lastMessageTs ? new Date(x.Thread.lastMessageTs) : new Date(x.createdAt);
    return {
      threadId: x.threadId,
      headline: x.headline || '',
      from: formatSender(x.Thread),
      subject: x.Thread?.subject || '(no subject)',
      summary: x.tldr || '',
      nextStep: x.nextStep || '',
      link: x.threadId ? `https://mail.google.com/mail/u/0/#all/${encodeURIComponent(x.threadId)}` : '',
      primer: x.chatPrimer || '',
      category: x.category || '',
      receivedAt: emailTs.toISOString(),
      convo: x.convoText || ''
    };
  });
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

type SecretaryEmail = {
  threadId: string;
  headline: string;
  from: string;
  subject: string;
  summary: string;
  nextStep: string;
  link: string;
  primer: string;
  category: string;
  receivedAt: string;
  convo: string;
};

function render(tpl: string, items: DecoratedSummary[], meta: PageMeta) {
  const secretaryScript = renderSecretaryAssistant(items, meta);
  return `${tpl}\n${secretaryScript}`;
}

function renderSecretaryAssistant(items: DecoratedSummary[], meta: PageMeta) {
  const threads = summariesToThreads(items);
  const payload = safeJson({
    threads,
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

async function ensureUserRecord(sessionData: any) {
  const user = sessionData?.user;
  if (!user?.id || !user?.email) return;
  await prisma.user.upsert({
    where: { id: user.id },
    update: {
      email: user.email,
      name: user.name ?? undefined,
      picture: user.picture ?? undefined
    },
    create: {
      id: user.id,
      email: user.email,
      name: user.name ?? null,
      picture: user.picture ?? null
    }
  });
}

function buildPrimerInputFromSummary(item: SummaryWithThread): ChatPrimerInput {
  return {
    threadId: item.threadId,
    subject: item.Thread?.subject || '',
    summary: item.tldr || '',
    nextStep: item.nextStep || '',
    headline: item.headline || '',
    fromLine: formatSender(item.Thread)
  };
}

function queuePrimerPrefetch(userId: string, inputs: ChatPrimerInput[], traceId: string) {
  const work = inputs.filter(item => shouldGeneratePrimer(userId, item.threadId));
  if (!work.length) return;
  const log = scopedLogger(`primerPrefetch:${traceId}`);
  const batches = chunkArray(work, PRIMER_BACKGROUND_BATCH);
  log('queueing background primer generation', { count: work.length, batches: batches.length });
  let chain = Promise.resolve();
  batches.forEach((batch, index) => {
    const batchTrace = `${traceId}-bg${index + 1}`;
    const batchPromise = chain.then(() => runPrimerBatch(userId, batch, batchTrace));
    for (const input of batch) {
      const key = primerCacheKey(userId, input.threadId);
      const perThread = batchPromise
        .then(result => result[input.threadId] || '')
        .finally(() => primerPending.delete(key));
      primerPending.set(key, perThread);
    }
    chain = batchPromise.then(() => undefined).catch(() => undefined);
  });
}

async function fetchPrimerForThread(userId: string, input: ChatPrimerInput, traceId: string) {
  const existing = getCachedPrimer(userId, input.threadId);
  if (existing) return existing;
  const key = primerCacheKey(userId, input.threadId);
  if (primerPending.has(key)) {
    return primerPending.get(key)!;
  }
  const batchPromise = runPrimerBatch(userId, [input], traceId);
  const perThread = batchPromise
    .then(result => result[input.threadId] || '')
    .finally(() => primerPending.delete(key));
  primerPending.set(key, perThread);
  return perThread;
}

function runPrimerBatch(userId: string, inputs: ChatPrimerInput[], traceId: string): Promise<Record<string, string>> {
  const log = scopedLogger(`primerBatch:${traceId}`);
  return generateChatPrimers(inputs, { traceId }).then(result => {
    cachePrimerResults(userId, result);
    return result;
  }).catch(err => {
    log('failed to generate primers', { error: (err as Error).message || err });
    return {} as Record<string, string>;
  });
}

function cachePrimerResults(userId: string, primers: Record<string, string>) {
  for (const [threadId, primer] of Object.entries(primers)) {
    if (!primer) continue;
    primerCache.set(primerCacheKey(userId, threadId), primer);
  }
}

function getCachedPrimer(userId: string, threadId: string) {
  return primerCache.get(primerCacheKey(userId, threadId)) || '';
}

function shouldGeneratePrimer(userId: string, threadId: string) {
  const key = primerCacheKey(userId, threadId);
  return threadId && !primerCache.has(key) && !primerPending.has(key);
}

function primerCacheKey(userId: string, threadId: string) {
  return `${userId}:${threadId}`;
}

function chunkArray<T>(items: T[], size: number): T[][] {
  if (size <= 0) return [items.slice()];
  const chunks: T[][] = [];
  for (let i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
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
