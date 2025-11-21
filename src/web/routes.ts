import { Router, Request, Response } from 'express';
import { ingestInbox } from '../gmail/fetch.js';
import { prisma } from '../store/db.js';
import { buildInboxBrief } from '../llm/morningBrief.js';
import type { InboxBrief } from '../llm/morningBrief.js';
import fs from 'node:fs/promises';
import path from 'node:path';
import { chatAboutEmail, MAX_CHAT_TURNS, type ChatTurn } from '../llm/secretaryChat.js';
import { gmailClient } from '../gmail/client.js';
import { getAuthedClient } from '../auth/google.js';
import { normalizeBody } from '../gmail/normalize.js';

export const router = Router();

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

  // Decide whether to auto-ingest AFTER rendering (first-time/empty state).
  const existingCount = await prisma.summary.count({ where: { userId } });
  const autoIngest = existingCount === 0;

  // Pull whatever is there (maybe empty), then sort newest message first
  const summaries = await prisma.summary.findMany({
    where: { userId },
    include: { Thread: true }
  });

  const sorted = summaries.sort((a, b) => {
    const at = a.Thread?.lastMessageTs ? new Date(a.Thread.lastMessageTs).getTime() : new Date(a.createdAt).getTime();
    const bt = b.Thread?.lastMessageTs ? new Date(b.Thread.lastMessageTs).getTime() : new Date(b.createdAt).getTime();
    return bt - at; // descending
  });

  const layout = await fs.readFile(path.join(process.cwd(), 'src/web/views/layout.html'), 'utf8');
  const body = await fs.readFile(path.join(process.cwd(), 'src/web/views/dashboard.html'), 'utf8');
  const brief = await buildInboxBrief(sorted);

  // Inject a small flag the client script can read to auto-trigger ingest
  const withFlag = `${render(body, sorted, brief)}
  <script>window.AUTO_INGEST = ${autoIngest ? 'true' : 'false'};</script>`;

  const html = layout.replace('<!--CONTENT-->', withFlag);
  res.send(html);
});

router.post('/ingest', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) return res.status(401).send('auth first');
  const userId = sessionData.user.id;

  // Clear current summaries so the dashboard shows only the latest pull
  await prisma.summary.deleteMany({ where: { userId } });
  try { await prisma.processing.deleteMany({ where: { userId } }); } catch { /* optional table */ }

  await ingestInbox(req);
  res.redirect('/dashboard');
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
  subject: string;
  summary: string;
  nextStep: string;
  link: string;
};

function render(tpl: string, items: any[], brief: InboxBrief) {
  const rows = items.map(x => {
    const emailTs = x.Thread?.lastMessageTs ? new Date(x.Thread.lastMessageTs) : new Date(x.createdAt);
    const when = emailTs.toLocaleString();
    const sender = x.Thread?.fromName || x.Thread?.fromEmail
      ? `${escapeHtml(x.Thread?.fromName || '')}${x.Thread?.fromName && x.Thread?.fromEmail ? ' ' : ''}${x.Thread?.fromEmail ? '&lt;' + escapeHtml(x.Thread.fromEmail) + '&gt;' : ''}`
      : '';
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
  const secretaryScript = renderSecretaryAssistant(items);
  return `${tpl
    .replace('<!--BRIEF-->', briefHtml)
    .replace('<!--ROWS-->', rows)
  }\n${secretaryScript}`;
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
  const data: SecretaryEmail[] = items.map(x => ({
    threadId: x.threadId,
    headline: x.headline || '',
    subject: x.Thread?.subject || '(no subject)',
    summary: x.tldr || '',
    nextStep: x.nextStep || '',
    link: x.threadId ? `https://mail.google.com/mail/u/0/#all/${encodeURIComponent(x.threadId)}` : ''
  }));
  const payload = safeJson(data);
  return `
<script>
(function () {
  const threads = ${payload};
  const MAX_TURNS = ${MAX_CHAT_TURNS};
  const messageEl = document.getElementById('secretary-message');
  const detailEl = document.getElementById('secretary-email');
  const summaryEl = document.getElementById('secretary-summary');
  const subjectEl = document.getElementById('secretary-subject');
  const nextEl = document.getElementById('secretary-next');
  const headlineEl = document.getElementById('secretary-headline');
  const linkEl = document.getElementById('secretary-link');
  const buttonEl = document.getElementById('secretary-button');
  const chatContainer = document.getElementById('secretary-chat');
  const chatLog = document.getElementById('secretary-chat-log');
  const chatForm = document.getElementById('secretary-chat-form');
  const chatInput = document.getElementById('secretary-chat-input');
  const chatHint = document.getElementById('secretary-chat-hint');
  const chatError = document.getElementById('secretary-chat-error');
  const chatPlaceholderHtml = '<div class="chat-placeholder">Ask for more details or clarifications about this thread.</div>';
  if (!buttonEl || !messageEl) return;

  if (!threads.length) {
    messageEl.textContent = 'Morning! Inbox is clear—nothing for us to review.';
    buttonEl.disabled = true;
    buttonEl.textContent = 'No emails';
    return;
  }

  let index = -1;
  let activeThreadId = '';
  const chatHistories = new Map();
  buttonEl.dataset.state = 'idle';

  if (chatForm) {
    chatForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      if (!activeThreadId || !chatInput) return;
      const question = chatInput.value.trim();
      if (!question) return;
      const history = ensureHistory(activeThreadId);
      const asked = history.filter(turn => turn.role === 'user').length;
      if (asked >= MAX_TURNS) {
        setChatError('Chat limit reached for this thread.');
        return;
      }
      setChatError('');
      const pending = { role: 'user', content: question };
      history.push(pending);
      renderChat(activeThreadId);
      chatInput.value = '';
      chatInput.disabled = true;
      const submitBtn = chatForm.querySelector('button[type="submit"]');
      if (submitBtn) submitBtn.disabled = true;
      try {
        const historyPayload = history.slice(0, -1);
        const resp = await fetch('/secretary/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            threadId: activeThreadId,
            question,
            history: historyPayload
          })
        });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) {
          history.pop();
          renderChat(activeThreadId);
          setChatError(data?.error || 'Something went wrong. Please try again.');
          chatInput.value = question;
          return;
        }
        history.push({ role: 'assistant', content: data.reply || 'No response received.' });
        renderChat(activeThreadId);
      } catch (err) {
        history.pop();
        renderChat(activeThreadId);
        setChatError('Failed to reach the assistant. Check your connection.');
        chatInput.value = question;
      } finally {
        chatInput.disabled = false;
        const submitBtn = chatForm.querySelector('button[type="submit"]');
        if (submitBtn) submitBtn.disabled = false;
        updateChatHint(activeThreadId);
      }
    });
  }

  function ensureHistory(threadId) {
    if (!chatHistories.has(threadId)) chatHistories.set(threadId, []);
    return chatHistories.get(threadId);
  }

  function renderChat(threadId) {
    if (!chatLog) return;
    const history = ensureHistory(threadId);
    if (!history.length) {
      chatLog.innerHTML = chatPlaceholderHtml;
      updateChatHint(threadId);
      return;
    }
    chatLog.innerHTML = history.map(turn => {
      return '<div class="chat-row chat-' + turn.role + '"><div class="chat-bubble">' + htmlEscape(turn.content) + '</div></div>';
    }).join('');
    chatLog.scrollTop = chatLog.scrollHeight;
    updateChatHint(threadId);
  }

  function updateChatHint(threadId) {
    if (!chatHint) return;
    const history = ensureHistory(threadId);
    const asked = history.filter(turn => turn.role === 'user').length;
    const remaining = Math.max(0, MAX_TURNS - asked);
    chatHint.textContent = remaining
      ? remaining + ' question' + (remaining === 1 ? '' : 's') + ' remaining in this chat.'
      : 'Chat limit reached. Wrap up this thread to move on.';
  }

  function setChatError(message) {
    if (!chatError) return;
    if (message) {
      chatError.textContent = message;
      chatError.classList.remove('hidden');
    } else {
      chatError.textContent = '';
      chatError.classList.add('hidden');
    }
  }

  function resetThreadView() {
    activeThreadId = '';
    if (detailEl) detailEl.classList.add('hidden');
    if (chatContainer) chatContainer.classList.add('hidden');
    if (chatLog) chatLog.innerHTML = chatPlaceholderHtml;
    if (chatInput) {
      chatInput.value = '';
      chatInput.disabled = false;
    }
    if (chatForm) {
      const submitBtn = chatForm.querySelector('button[type="submit"]');
      if (submitBtn) submitBtn.disabled = false;
    }
    if (chatHint) chatHint.textContent = 'Ask up to ' + MAX_TURNS + ' questions per thread.';
    setChatError('');
  }

  function htmlEscape(value) {
    const div = document.createElement('div');
    div.textContent = value || '';
    return div.innerHTML;
  }

  buttonEl.addEventListener('click', () => {
    const state = buttonEl.dataset.state;
    if (state === 'complete') return;
    if (state === 'ready-to-close') {
      buttonEl.dataset.state = 'complete';
      buttonEl.disabled = true;
      messageEl.textContent = 'All caught up—ping me if you want another pass.';
      chatHistories.clear();
      resetThreadView();
      return;
    }
    if (index === -1) {
      messageEl.textContent = \`Great, here's email 1 of \${threads.length}.\`;
    } else {
      const nextPosition = index + 2;
      messageEl.textContent = \`Next up, email \${nextPosition} of \${threads.length}.\`;
    }
    advance();
  });

  function advance() {
    index += 1;
    const current = threads[index];
    if (!current) return;
    if (detailEl) detailEl.classList.remove('hidden');
    if (chatContainer) chatContainer.classList.remove('hidden');
    activeThreadId = current.threadId;
    ensureHistory(activeThreadId);
    renderChat(activeThreadId);
    setChatError('');

    if (headlineEl) {
      if (current.headline) {
        headlineEl.textContent = current.headline;
        headlineEl.classList.remove('hidden');
      } else {
        headlineEl.textContent = '';
        headlineEl.classList.add('hidden');
      }
    }
    if (subjectEl) subjectEl.textContent = current.subject;
    if (summaryEl) summaryEl.textContent = current.summary || 'No summary captured.';
    if (nextEl) nextEl.textContent = current.nextStep || 'No next step needed.';
    if (linkEl) {
      if (current.link) {
        linkEl.href = current.link;
        linkEl.classList.remove('hidden');
      } else {
        linkEl.removeAttribute('href');
        linkEl.classList.add('hidden');
      }
    }

    if (index === threads.length - 1) {
      buttonEl.textContent = 'Done';
      buttonEl.dataset.state = 'ready-to-close';
      messageEl.textContent = threads.length === 1
        ? 'Only one email waiting. Tap Done when you are finished.'
        : \`Last one—email \${threads.length} of \${threads.length}.\`;
    } else {
      buttonEl.textContent = 'Next email';
      buttonEl.dataset.state = 'chatting';
    }
  }
})();
</script>
`;
}

function escapeHtml(s: string) {
  return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]!));
}

function safeJson(value: unknown) {
  return JSON.stringify(value).replace(/</g, '\\u003c');
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
