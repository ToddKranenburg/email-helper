import pLimit from 'p-limit';
import { gmailClient, INBOX_QUERY } from './client.js';
import { normalizeBody } from './normalize.js';
import { getAuthedClient } from '../auth/google.js';
import { prisma } from '../store/db.js';
import type { Request } from 'express';

function parseFrom(headerValue: string | undefined): { name?: string; email?: string } {
  if (!headerValue) return {};
  const emailMatch = headerValue.match(/<([^>]+)>/);
  const email = emailMatch ? emailMatch[1].trim() : (/@/.test(headerValue) ? headerValue.trim() : undefined);
  let name: string | undefined = undefined;
  if (emailMatch) {
    name = headerValue.replace(emailMatch[0], '').trim().replace(/^"|"$/g, '');
  } else if (email) {
    name = undefined;
  } else {
    name = headerValue.trim();
  }
  return { name, email };
}

export async function ingestInbox(req: Request) {
  const auth = getAuthedClient(req.session);
  const gmail = gmailClient(auth);

  let pageToken: string | undefined = undefined;
  const limit = pLimit(6);

  do {
    const list = await gmail.users.threads.list({
      userId: 'me',
      q: INBOX_QUERY,
      pageToken
    });

    const threads = list.data.threads || [];
    await Promise.all(threads.map(t => limit(async () => {
      const full = await gmail.users.threads.get({ userId: 'me', id: t.id! });
      const msgs = (full.data.messages || []).slice(-3);
      const latest = msgs[msgs.length - 1];

      const hdrs = latest?.payload?.headers || [];
      const subject = hdrs.find(h => h.name === 'Subject')?.value || '';
      const fromRaw = hdrs.find(h => h.name === 'From')?.value;
      const { name: fromName, email: fromEmail } = parseFrom(fromRaw);
      const latestMsgId = latest?.id!;

      const participants = Array.from(
        new Set(
          hdrs
            .filter(h => ['From', 'To', 'Cc'].includes(h.name))
            .flatMap(h => h.value?.split(',') || [])
            .map(s => s.trim())
        )
      );

      await prisma.thread.upsert({
        where: { id: full.data.id! },
        update: {
          subject,
          participants: JSON.stringify(participants),
          lastMessageTs: new Date(Number(latest?.internalDate)),
          historyId: full.data.historyId,
          fromName,
          fromEmail
        },
        create: {
          id: full.data.id!,
          subject,
          participants: JSON.stringify(participants),
          lastMessageTs: new Date(Number(latest?.internalDate)),
          historyId: full.data.historyId,
          fromName,
          fromEmail
        }
      });

      // Skip if we've already summarized this latest message
      const existing = await prisma.summary.findUnique({ where: { lastMsgId: latestMsgId } });
      if (existing) return;

      const convoText = msgs
        .map(m => normalizeBody(m.payload))
        .filter(Boolean)
        .reverse()
        .join('\n\n---\n\n');

      await summarizeAndStore(full.data.id!, latestMsgId, subject, participants, convoText);
    })));

    pageToken = list.data.nextPageToken || undefined;
  } while (pageToken);
}

import { summarize } from '../llm/summarize.js';
async function summarizeAndStore(
  threadId: string,
  lastMsgId: string,
  subject: string,
  people: string[],
  convoText: string
) {
  const s = await summarize({ subject, people, convoText });
  await prisma.summary.create({
    data: {
      threadId,
      lastMsgId,
      tldr: s.tldr,
      category: s.category,
      nextStep: s.next_step,
      confidence: s.confidence
    }
  });
}
