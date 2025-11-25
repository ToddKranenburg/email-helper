import { Router, Request, Response } from 'express';
import { ingestInbox } from '../gmail/fetch.js';
import { prisma } from '../store/db.js';
import { buildInboxBrief } from '../llm/morningBrief.js';
import type { InboxBrief } from '../llm/morningBrief.js';
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

  const pagination = {
    page: currentPage,
    totalPages,
    totalItems,
    hasPrevious: currentPage > 1,
    hasNext: currentPage < totalPages,
    start: totalItems ? skip + 1 : 0,
    end: totalItems ? skip + visible.length : 0
  };

  const layout = await fs.readFile(path.join(process.cwd(), 'src/web/views/layout.html'), 'utf8');
  const body = await fs.readFile(path.join(process.cwd(), 'src/web/views/dashboard.html'), 'utf8');
  const brief = await buildInboxBrief(visible);
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
  const withFlag = `${render(body, decorated, brief, pagination)}
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

type PaginationState = {
  page: number;
  totalPages: number;
  totalItems: number;
  hasPrevious: boolean;
  hasNext: boolean;
  start: number;
  end: number;
};

function render(tpl: string, items: any[], brief: InboxBrief, pagination: PaginationState) {
  const rows = items.map(x => {
    const emailTs = x.Thread?.lastMessageTs ? new Date(x.Thread.lastMessageTs) : new Date(x.createdAt);
    const when = emailTs.toLocaleString();
    const senderText = formatSender(x.Thread);
    const sender = senderText ? escapeHtml(senderText) : '';
    const emoji = emojiForCategory(x.category);

    return `
    <div class="card">
      <div class="superhead"><span class="emoji">${emoji}</span><span class="superhead-text">${escapeHtml(x.category)}</span></div>
      <div class="headline">${escapeHtml(x.headline || '')}</div>
      <div class="meta">${when}</div>
      ${sender ? `<div class="meta"><span class="label">From:</span> ${sender}</div>` : ''}
      <div class="meta"><span class="label">Subject:</span> ${escapeHtml(x.Thread.subject || '(no subject)')}</div>

      <p>${escapeHtml(x.tldr)}</p>
      <p class="next"><span class="label">Next:</span> ${escapeHtml(x.nextStep || 'No action')}</p>
      <a href="https://mail.google.com/mail/u/0/#all/${x.threadId}" target="_blank">Open in Gmail</a>
    </div>
  `;
  }).join('\n');
  const briefHtml = renderBrief(brief);
  const paginationHtml = renderPagination(pagination);
  const secretaryScript = renderSecretaryAssistant(items);
  return `${tpl
    .replace('<!--BRIEF-->', briefHtml)
    .replace('<!--ROWS-->', rows)
    .replace('<!--PAGINATION-->', paginationHtml)
  }\n${secretaryScript}`;
}

function renderPagination(pagination: PaginationState) {
  if (!pagination.totalItems) {
    return '<p class="pagination-status">No email summaries yet. Ingest your inbox to get started.</p>';
  }
  const status = `Showing ${pagination.start}&ndash;${pagination.end} of ${pagination.totalItems} email${pagination.totalItems === 1 ? '' : 's'}`;
  const pageInfo = pagination.totalPages > 1 ? `Page ${pagination.page} of ${pagination.totalPages}` : '';
  const prevLabel = pagination.page === 2 ? 'Newest' : `Newer ${PAGE_SIZE}`;
  const nextLabel = `Older ${PAGE_SIZE}`;
  const prevControl = pagination.hasPrevious
    ? `<a class="pagination-btn" href="?page=${pagination.page - 1}" rel="prev">${escapeHtml(prevLabel)}</a>`
    : '<span class="pagination-btn disabled">Newer</span>';
  const nextControl = pagination.hasNext
    ? `<a class="pagination-btn" href="?page=${pagination.page + 1}" rel="next">${escapeHtml(nextLabel)}</a>`
    : '<span class="pagination-btn disabled">Older</span>';
  const note = pagination.totalPages > 1
    ? `<p class="pagination-note">Use Newer/Older to browse messages in batches of ${PAGE_SIZE}.</p>`
    : '';
  return `
    <div class="pagination-inner">
      <div>
        <div class="pagination-status">${status}</div>
        ${pageInfo ? `<div class="pagination-page">${pageInfo}</div>` : ''}
      </div>
      <div class="pagination-controls">
        ${prevControl}
        ${nextControl}
      </div>
    </div>
    ${note}
  `;
}

function renderBrief(brief: InboxBrief) {
  const bullets = (brief.highlights || []).map((point, index) => {
    const threadId = brief.highlightTargets?.[index];
    if (threadId) {
      const url = `https://mail.google.com/mail/u/0/#all/${encodeURIComponent(threadId)}`;
      return `<li><a class="brief-link" href="${escapeHtml(url)}" target="_blank" rel="noreferrer">${escapeHtml(point)}<span aria-hidden="true" class="brief-link-icon">↗</span></a></li>`;
    }
    return `<li>${escapeHtml(point)}</li>`;
  }).join('');
  const list = bullets ? `<ul class="brief-highlights">${bullets}</ul>` : '';
  return `
    <div class="card brief-card">
      <div class="brief-label">Morning Brief</div>
      <div class="brief-title">${escapeHtml(brief.title)}</div>
      <p>${escapeHtml(brief.overview)}</p>
      ${list}
    </div>
  `;
}

function renderSecretaryAssistant(items: any[]) {
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
  const payload = safeJson({ threads, maxTurns: MAX_CHAT_TURNS }); // consumed by src/web/public/secretary.js
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
