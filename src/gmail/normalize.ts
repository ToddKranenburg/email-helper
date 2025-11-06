import { htmlToText } from 'html-to-text';
import { stripQuoted } from '../util/quoteStrip.js';

export function normalizeBody(payload: any): string {
  // walk MIME tree to find text/plain, else text/html
  function* walk(p: any): any {
    if (!p) return;
    if (p.mimeType?.startsWith('text/')) yield p;
    if (p.parts) for (const part of p.parts) yield* walk(part);
  }

  let best: any = null;
  for (const part of walk(payload)) {
    if (part.mimeType === 'text/plain') { best = part; break; }
    if (!best && part.mimeType === 'text/html') best = part;
  }
  if (!best) return '';
  const data = Buffer.from(best.body.data || '', 'base64').toString('utf8');
  const text = best.mimeType === 'text/html' ? htmlToText(data, { wordwrap: false }) : data;
  return stripQuoted(text);
}
