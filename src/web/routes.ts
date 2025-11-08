import { Router, Request, Response } from 'express';
import { ingestInbox } from '../gmail/fetch.js';
import { prisma } from '../store/db.js';
import { buildInboxBrief } from '../llm/morningBrief.js';
import type { InboxBrief } from '../llm/morningBrief.js';
import fs from 'node:fs/promises';
import path from 'node:path';

export const router = Router();

router.get('/', async (req: Request, res: Response) => {
  const tokens = (req.session as any).googleTokens;
  if (tokens && tokens.access_token) {
    // ✅ Already authorized: go straight to dashboard
    return res.redirect('/dashboard');
  }
  // ❌ Not authorized yet: show connect link
  res.send(`<a href="/auth/google">Connect Gmail</a>`);
});

router.get('/dashboard', async (req: Request, res: Response) => {
  if (!(req.session as any).googleTokens) return res.redirect('/auth/google');

  // Decide whether to auto-ingest AFTER rendering (first-time/empty state).
  const existingCount = await prisma.summary.count();
  const autoIngest = existingCount === 0;

  // Pull whatever is there (maybe empty), then sort newest message first
  const summaries = await prisma.summary.findMany({
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
  if (!(req.session as any).googleTokens) return res.status(401).send('auth first');

  // Clear current summaries so the dashboard shows only the latest pull
  await prisma.summary.deleteMany({});
  try { await prisma.processing.deleteMany({}); } catch { /* optional table */ }

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
  return tpl
    .replace('<!--BRIEF-->', briefHtml)
    .replace('<!--ROWS-->', rows);
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

function escapeHtml(s: string) {
  return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]!));
}
