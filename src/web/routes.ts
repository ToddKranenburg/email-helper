import { Router, Request, Response } from 'express';
import { ingestInbox } from '../gmail/fetch.js';
import { prisma } from '../store/db.js';
import { buildInboxBrief } from '../llm/morningBrief.js';
import type { InboxBrief } from '../llm/morningBrief.js';
import fs from 'node:fs/promises';
import path from 'node:path';

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
  const messageEl = document.getElementById('secretary-message');
  const detailEl = document.getElementById('secretary-email');
  const summaryEl = document.getElementById('secretary-summary');
  const subjectEl = document.getElementById('secretary-subject');
  const nextEl = document.getElementById('secretary-next');
  const headlineEl = document.getElementById('secretary-headline');
  const linkEl = document.getElementById('secretary-link');
  const buttonEl = document.getElementById('secretary-button');
  if (!buttonEl || !messageEl) return;

  if (!threads.length) {
    messageEl.textContent = 'Morning! Inbox is clear—nothing for us to review.';
    buttonEl.disabled = true;
    buttonEl.textContent = 'No emails';
    return;
  }

  let index = -1;
  buttonEl.dataset.state = 'idle';

  buttonEl.addEventListener('click', () => {
    const state = buttonEl.dataset.state;
    if (state === 'complete') return;
    if (state === 'ready-to-close') {
      buttonEl.dataset.state = 'complete';
      buttonEl.disabled = true;
      messageEl.textContent = 'All caught up—ping me if you want another pass.';
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
