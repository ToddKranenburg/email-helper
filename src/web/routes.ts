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
const PRIMER_SYNC_BATCH = 6;
const primerCache = new Map<string, string>();
const primerPending = new Map<string, Promise<string>>();

type SummaryWithThread = Summary & { Thread: Thread | null };

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

router.get('/dashboard', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) return res.redirect('/auth/google');
  const userId = sessionData.user.id;
  const traceId = createTraceId();
  const log = scopedLogger(`dashboard:${traceId}`);
  const routeStart = performance.now();
  log('start', { userId });
  await ensureUserRecord(sessionData);

  // Decide whether to auto-ingest AFTER rendering (first-time/empty state).
  const countStart = performance.now();
  const totalItems = await prisma.summary.count({ where: { userId } });
  log('summary count complete', { durationMs: elapsedMs(countStart), totalItems });
  const hasSummaries = totalItems > 0;
  if (hasSummaries) {
    sessionData.skipAutoIngest = true;
  }
  const autoIngest = !hasSummaries && !sessionData.skipAutoIngest;

  const requestedPage = typeof req.query.page === 'string' ? parseInt(req.query.page, 10) : 1;
  const totalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
  const currentPage = Number.isFinite(requestedPage) ? Math.min(Math.max(requestedPage, 1), totalPages) : 1;
  const skip = (currentPage - 1) * PAGE_SIZE;

  const listStart = performance.now();
  const visible = await prisma.summary.findMany({
    where: { userId },
    include: { Thread: true },
    orderBy: [
      { Thread: { lastMessageTs: 'desc' } },
      { createdAt: 'desc' }
    ],
    skip,
    take: PAGE_SIZE
  }) as SummaryWithThread[];
  log('loaded summaries', { durationMs: elapsedMs(listStart), returned: visible.length, page: currentPage });

  const templateStart = performance.now();
  const layout = await fs.readFile(path.join(process.cwd(), 'src/web/views/layout.html'), 'utf8');
  const body = await fs.readFile(path.join(process.cwd(), 'src/web/views/dashboard.html'), 'utf8');
  log('templates read', { durationMs: elapsedMs(templateStart) });
  const primerInputs = visible.map(buildPrimerInputFromSummary);
  const initialInputs = primerInputs.slice(0, PRIMER_SYNC_BATCH);
  const remainingInputs = primerInputs.slice(PRIMER_SYNC_BATCH);
  const primerStart = performance.now();
  const primers = await generateChatPrimers(initialInputs, { traceId });
  cachePrimerResults(userId, primers);
  log('chat primers ready', { durationMs: elapsedMs(primerStart), count: initialInputs.length });
  const decorated = visible.map(item => ({
    ...item,
    chatPrimer: primers[item.threadId] || getCachedPrimer(userId, item.threadId) || ''
  }));

  // Inject a small flag the client script can read to auto-trigger ingest
  const renderStart = performance.now();
  const withFlag = `${render(body, decorated, totalItems)}
  <script>window.AUTO_INGEST = ${autoIngest ? 'true' : 'false'};</script>`;

  const html = layout.replace('<!--CONTENT-->', withFlag);
  log('html rendered', { durationMs: elapsedMs(renderStart) });
  res.send(html);
  log('completed', { durationMs: elapsedMs(routeStart) });
  queuePrimerPrefetch(userId, remainingInputs, traceId);
});

router.post('/ingest', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) return res.status(401).send('auth first');
  const userId = sessionData.user.id;
  await ensureUserRecord(sessionData);
  sessionData.skipAutoIngest = true;

  // Clear current summaries so the dashboard shows only the latest pull
  await prisma.summary.deleteMany({ where: { userId } });
  try { await prisma.processing.deleteMany({ where: { userId } }); } catch { /* optional table */ }

  try {
    await ingestInbox(req);
    res.redirect('/dashboard');
  } catch (err) {
    console.error('Inbox ingest failed', err);
    if (err instanceof GaxiosError && err.response?.status === 403) {
      return res
        .status(403)
        .send('Gmail refused to share inbox data. Please disconnect and reconnect your Google account.');
    }
    res.status(500).send('Unable to sync your Gmail inbox right now. Please try again.');
  }
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

function render(tpl: string, items: any[], totalItems: number) {
  const secretaryScript = renderSecretaryAssistant(items, totalItems);
  return `${tpl}\n${secretaryScript}`;
}

function renderSecretaryAssistant(items: any[], totalItems: number) {
  const threads: SecretaryEmail[] = items.map(x => {
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
  const payload = safeJson({ threads, maxTurns: MAX_CHAT_TURNS, totalItems }); // consumed by src/web/public/secretary.js
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
  log('queueing background primer generation', { count: work.length });
  const batchPromise = runPrimerBatch(userId, work, traceId);
  for (const input of work) {
    const key = primerCacheKey(userId, input.threadId);
    const perThread = batchPromise
      .then(result => result[input.threadId] || '')
      .finally(() => primerPending.delete(key));
    primerPending.set(key, perThread);
  }
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
