import type { gmail_v1 } from 'googleapis';
import { normalizeBody } from './normalize.js';

function headerValue(headers: gmail_v1.Schema$MessagePartHeader[] | undefined, name: string) {
  if (!headers?.length) return null;
  const match = headers.find(h => h?.name === name);
  return match?.value ?? null;
}

function formatMessage(msg: gmail_v1.Schema$Message) {
  const headers = msg.payload?.headers || [];
  const from = headerValue(headers, 'From') || 'Unknown sender';
  const to = headerValue(headers, 'To');
  const date = headerValue(headers, 'Date');
  const body = normalizeBody(msg.payload);
  if (!body) return '';
  const headerLines = [`From: ${from}`];
  if (to) headerLines.push(`To: ${to}`);
  if (date) headerLines.push(`Date: ${date}`);
  return `${headerLines.join('\n')}\n\n${body.trim()}`;
}

export function buildTranscript(messages: gmail_v1.Schema$Message[], opts: { maxChars?: number } = {}) {
  const parts = messages.map(msg => formatMessage(msg)).filter(Boolean);
  const combined = parts.join('\n\n---\n\n');
  if (opts.maxChars && combined.length > opts.maxChars) {
    return combined.slice(0, opts.maxChars);
  }
  return combined;
}
