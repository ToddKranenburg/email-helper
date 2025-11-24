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
  const autoIngest = totalItems === 0;

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

  // Inject a small flag the client script can read to auto-trigger ingest
  const withFlag = `${render(body, visible, brief, pagination)}
  <script>window.AUTO_INGEST = ${autoIngest ? 'true' : 'false'};</script>`;

  const html = layout.replace('<!--CONTENT-->', withFlag);
  res.send(html);
});

router.post('/ingest', async (req: Request, res: Response) => {
  const sessionData = req.session as any;
  if (!sessionData.googleTokens || !sessionData.user?.id) return res.status(401).send('auth first');
  const userId = sessionData.user.id;
  await ensureUserRecord(sessionData);

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
  from: string;
  subject: string;
  summary: string;
  nextStep: string;
  link: string;
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
  const data: SecretaryEmail[] = items.map(x => ({
    threadId: x.threadId,
    headline: x.headline || '',
    from: formatSender(x.Thread),
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
  const fromEl = document.getElementById('secretary-from');
  const summaryEl = document.getElementById('secretary-summary');
  const subjectEl = document.getElementById('secretary-subject');
  const nextEl = document.getElementById('secretary-next');
  const headlineEl = document.getElementById('secretary-headline');
  const linkEl = document.getElementById('secretary-link');
  const buttonEl = document.getElementById('secretary-button');
  const backButtonEl = document.getElementById('secretary-back-button');
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
    if (backButtonEl) backButtonEl.disabled = true;
    return;
  }

  let index = -1;
  let activeThreadId = '';
  const chatHistories = new Map();
  let chatAdvanceTimer = 0;
  buttonEl.dataset.state = 'idle';
  if (backButtonEl) backButtonEl.disabled = true;

  if (chatForm) {
    chatForm.addEventListener('submit', async (event) => {
      event.preventDefault();
      if (!activeThreadId || !chatInput) return;
      const question = chatInput.value.trim();
      if (!question) return;
      const history = ensureHistory(activeThreadId);
      if (handleNextIntent(question, history)) {
        return;
      }
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

  if (chatForm && chatInput) {
    chatInput.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter' || event.shiftKey) return;
      event.preventDefault();
      if (typeof chatForm.requestSubmit === 'function') {
        chatForm.requestSubmit();
      } else {
        chatForm.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
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

  function handleNextIntent(question, history) {
    if (!activeThreadId || !shouldTriggerNextIntent(question)) return false;
    if (!history || !Array.isArray(history)) return false;
    setChatError('');
    history.push({ role: 'user', content: question });
    renderChat(activeThreadId);
    if (chatInput) {
      chatInput.value = '';
      chatInput.disabled = true;
    }
    const submitBtn = chatForm ? chatForm.querySelector('button[type="submit"]') : null;
    if (submitBtn) submitBtn.disabled = true;
    const isLastThread = index >= threads.length - 1;
    const response = isLastThread
      ? 'That was the final email in your queue. Tap Done when you are ready.'
      : 'Proceeding to the next email...';
    history.push({ role: 'assistant', content: response });
    renderChat(activeThreadId);
    if (chatAdvanceTimer) window.clearTimeout(chatAdvanceTimer);
    const delay = isLastThread ? 900 : 700;
    chatAdvanceTimer = window.setTimeout(() => {
      chatAdvanceTimer = 0;
      if (chatInput) chatInput.disabled = false;
      if (submitBtn) submitBtn.disabled = false;
      if (!isLastThread) {
        showThreadAt(index + 1, 'next');
      }
    }, delay);
    return true;
  }

  function shouldTriggerNextIntent(rawText) {
    if (!rawText) return false;
    const normalized = rawText.toLowerCase();
    const simple = normalized.replace(/[^a-z0-9\\s]/g, ' ').replace(/\\s+/g, ' ').trim();
    if (!simple) return false;
    const matchable = simple.replace(/(?: thanks?| thank you)+$/, '').trim() || simple;

    const directPatterns = [
      /^next( (email|one|thread|message|item|mail))?( please)?$/,
      /^(?:skip|pass)(?: (?:this|it)(?: one)?)?( please)?$/,
      /^(?:let s|lets|let us|shall we|we can|can we|could we|please) move on(?: (now|then))?$/,
      /^(?:onto|on to) the next( (one|email|thread|message|item))?$/,
      /^time for the next( (one|email|thread|message|item))?$/,
      /^(?:ready|i m ready|im ready|we re ready|were ready|ok|okay) (?:for|to move on to) (?:the )?next( (email|one|thread|message|item))?( please)?$/,
      /^(?:all set|done|im done|i m done|we re done|were done) (?:here|with (?:this|it)(?: one| email| thread| message)?)$/,
      /^that s all (?:for|with) (?:this|it)(?: one| email| thread| message)?$/
    ];
    const segments = matchable.split(/[,;]+/).map(part => part.trim()).filter(Boolean);
    const targets = segments.length ? segments : [matchable];
    if (targets.some(part => directPatterns.some(pattern => pattern.test(part)))) {
      return true;
    }

    if (/\\bnext( (email|one|thread|message|item|mail))? please$/.test(matchable)) {
      return true;
    }

    const targetedCombos = [
      'move on to the next email',
      'move on to the next one',
      'move on to the next thread',
      'move on to the next message',
      'move onto the next email',
      'move onto the next one',
      'move onto the next thread',
      'go to the next email',
      'go to the next one',
      'go to the next thread',
      'go on to the next email',
      'go on to the next one',
      'go on to the next thread',
      'proceed to the next email',
      'proceed to the next one',
      'proceed to the next thread',
      'show me the next email',
      'show me the next one',
      'ready for the next email',
      'ready for the next one',
      'ready to move on to the next',
      'done with this one',
      'done with this email',
      'done with this thread',
      'done with this message',
      'all set with this one',
      'all set with this thread',
      'skip this email',
      'skip this thread',
      'skip this one',
      'pass this email',
      'pass this thread',
      'let s move on to the next',
      'lets move on to the next',
      'let us move on to the next',
      'let s move on',
      'lets move on',
      'let us move on',
      'onto the next email',
      'onto the next one',
      'on to the next email',
      'on to the next one',
      'time for the next email',
      'time for the next one'
    ];
    return targetedCombos.some(text => matchable.includes(text));
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
    if (backButtonEl) backButtonEl.disabled = true;
    index = -1;
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
      if (backButtonEl) backButtonEl.disabled = true;
      messageEl.textContent = 'All caught up—ping me if you want another pass.';
      chatHistories.clear();
      resetThreadView();
      return;
    }
    const targetIndex = index + 1;
    showThreadAt(targetIndex, 'next');
  });

  if (backButtonEl) {
    backButtonEl.addEventListener('click', () => {
      if (backButtonEl.disabled) return;
      const previousIndex = index - 1;
      showThreadAt(previousIndex, 'back');
    });
  }

  function showThreadAt(newIndex, direction) {
    if (typeof newIndex !== 'number') return;
    if (newIndex < 0 || newIndex >= threads.length) return;
    index = newIndex;
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
    if (fromEl) {
      fromEl.textContent = current.from || 'Sender unknown';
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

    updateMessageAfterNavigation(direction);
    updateNavigationButtons();
  }

  function updateNavigationButtons() {
    if (index === threads.length - 1) {
      buttonEl.textContent = 'Done';
      buttonEl.dataset.state = 'ready-to-close';
    } else {
      buttonEl.textContent = 'Next email';
      buttonEl.dataset.state = 'chatting';
    }
    if (backButtonEl) backButtonEl.disabled = index <= 0;
  }

  function updateMessageAfterNavigation(direction) {
    if (!messageEl) return;
    if (threads.length === 1) {
      messageEl.textContent = 'Only one email waiting. Tap Done when you are finished.';
      return;
    }
    const position = index + 1;
    if (direction === 'back') {
      messageEl.textContent = 'Back to email ' + position + ' of ' + threads.length + '.';
      return;
    }
    if (position === 1) {
      messageEl.textContent = "Great, here's email 1 of " + threads.length + '.';
      return;
    }
    if (position === threads.length) {
      messageEl.textContent = 'Last one—email ' + threads.length + ' of ' + threads.length + '.';
      return;
    }
    messageEl.textContent = 'Reviewing email ' + position + ' of ' + threads.length + '.';
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
