import { Router } from 'express';
import { getAuthedClient } from '../auth/google.js';
import { ingestInbox } from '../gmail/fetch.js';
import { prisma } from '../store/db.js';
import fs from 'node:fs/promises';
import path from 'node:path';

export const router = Router();

router.get('/', (_req, res) => res.send(`<a href="/auth/google">Connect Gmail</a>`));

router.get('/dashboard', async (req, res) => {
  if (!(req.session as any).googleTokens) return res.redirect('/auth/google');
  const summaries = await prisma.summary.findMany({
    orderBy: { createdAt: 'desc' },
    include: { Thread: true },
    take: 200
  });
  const layout = await fs.readFile(path.join(process.cwd(), 'src/web/views/layout.html'), 'utf8');
  const body = await fs.readFile(path.join(process.cwd(), 'src/web/views/dashboard.html'), 'utf8');
  const html = layout.replace('<!--CONTENT-->', render(body, summaries));
  res.send(html);
});

router.post('/ingest', async (req, res) => {
  if (!(req.session as any).googleTokens) return res.status(401).send('auth first');
  await ingestInbox(req);
  res.redirect('/dashboard');
});

function render(tpl: string, items: any[]) {
  const rows = items.map(x => `
    <div class="card">
      <div class="subject">${escapeHtml(x.Thread.subject || '(no subject)')}</div>
      <div class="meta">${new Date(x.createdAt).toLocaleString()} • ${x.category} • ${x.confidence}</div>
      <p>${escapeHtml(x.tldr)}</p>
      <p class="next">Next: ${escapeHtml(x.nextStep || 'None')}</p>
      <a href="https://mail.google.com/mail/u/0/#all/${x.threadId}" target="_blank">Open in Gmail</a>
    </div>
  `).join('\n');
  return tpl.replace('<!--ROWS-->', rows);
}

function escapeHtml(s: string) {
  return s.replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]!));
}
