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

export const router = Router();
const PAGE_SIZE = 20;

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
  await ensureUserRecord(sessionData);

  // Decide whether to auto-ingest AFTER rendering (first-time/empty state).
  const totalItems = await prisma.summary.count({ where: { userId } });
  const hasSummaries = totalItems > 0;
  if (hasSummaries) {
    sessionData.skipAutoIngest = true;
  }
  const autoIngest = !hasSummaries && !sessionData.skipAutoIngest;

  const requestedPage = typeof req.query.page === 'string' ? parseInt(req.query.page, 10) : 1;
  const totalPages = Math.max(1, Math.ceil(totalItems / PAGE_SIZE));
  const currentPage = Number.isFinite(requestedPage) ? Math.min(Math.max(requestedPage, 1), totalPages) : 1;
  const skip = (currentPage - 1) * PAGE_SIZE;

  const visible = await prisma.summary.findMany({
    where: { userId },
    include: { Thread: true },
    orderBy: [
      { Thread: { lastMessageTs: 'desc' } },
      { createdAt: 'desc' }
    ],
    skip,
    take: PAGE_SIZE
  });

  const layout = await fs.readFile(path.join(process.cwd(), 'src/web/views/layout.html'), 'utf8');
  const body = await fs.readFile(path.join(process.cwd(), 'src/web/views/dashboard.html'), 'utf8');
  const primerInputs: ChatPrimerInput[] = visible.map(item => ({
    threadId: item.threadId,
    subject: item.Thread?.subject || '',
    summary: item.tldr || '',
    nextStep: item.nextStep || '',
    headline: item.headline || '',
    fromLine: formatSender(item.Thread)
  }));
  const primers = await generateChatPrimers(primerInputs);
  const decorated = visible.map(item => ({
    ...item,
    chatPrimer: primers[item.threadId] || ''
  }));

  // Inject a small flag the client script can read to auto-trigger ingest
  const withFlag = `${render(body, decorated, totalItems)}
  <script>window.AUTO_INGEST = ${autoIngest ? 'true' : 'false'};</script>`;

  const html = layout.replace('<!--CONTENT-->', withFlag);
  res.send(html);
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
<script src="/secretary.js" defer></script>
`;
}

function escapeHtml(s: string) {
  return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]!));
}

function safeJson(value: unknown) {
  return JSON.stringify(value).replace(/</g, '\\u003c');
}

function formatSender(thread?: { fromName?: string | null; fromEmail?: string | null }) {
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
