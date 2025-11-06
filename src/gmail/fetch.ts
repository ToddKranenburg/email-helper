import pLimit from 'p-limit';
import { gmailClient, INBOX_QUERY } from './client.js';
import { normalizeBody } from './normalize.js';
import { getAuthedClient } from '../auth/google.js';
import { prisma } from '../store/db.js';
import type { Request } from 'express';

export async function ingestInbox(req: Request) {
  const auth = getAuthedClient(req.session);
  const gmail = gmailClient(auth);

  let pageToken: string | undefined = undefined;
  const limit = pLimit(6); // API concurrency

  do {
    const list = await gmail.users.threads.list({
      userId: 'me',
      q: INBOX_QUERY,
      pageToken
    });

    const threads = list.data.threads || [];
    await Promise.all(threads.map(t => limit(async () => {
      const full = await gmail.users.threads.get({ userId: 'me', id: t.id! });
      const msgs = (full.data.messages || []).slice(-3); // last N messages
      const latest = msgs[msgs.length - 1];

      const subject = latest?.payload?.headers?.find(h => h.name === 'Subject')?.value || '';
      const participants = Array.from(
        new Set(
          (latest?.payload?.headers || [])
            .filter(h => ['From', 'To', 'Cc'].includes(h.name))
            .flatMap(h => h.value?.split(',') || [])
            .map(s => s.trim())
        )
      );

      const body = normalizeBody(latest?.payload);
      // upsert thread (without summary yet)
      await prisma.thread.upsert({
        where: { id: full.data.id! },
        update: {
          subject,
          participants: JSON.stringify(participants),
          lastMessageTs: new Date(Number(latest?.internalDate)),
          historyId: full.data.historyId
        },
        create: {
          id: full.data.id!,
          subject,
          participants: JSON.stringify(participants),
          lastMessageTs: new Date(Number(latest?.internalDate)),
          historyId: full.data.historyId
        }
      });

      // Summarize last N messages (joined, newest→oldest)
      const convoText = msgs
        .map(m => normalizeBody(m.payload))
        .filter(Boolean)
        .reverse() // oldest→newest for summarizer
        .join('\n\n---\n\n');

      await summarizeAndStore(full.data.id!, subject, participants, convoText);
    })));

    pageToken = list.data.nextPageToken || undefined;
  } while (pageToken);
}

import { summarize } from '../llm/summarize.js';
async function summarizeAndStore(threadId: string, subject: string, people: string[], convoText: string) {
  const s = await summarize({ subject, people, convoText });
  await prisma.summary.create({
    data: {
      threadId,
      tldr: s.tldr,
      category: s.category,
      nextStep: s.next_step,
      confidence: s.confidence
    }
  });
}
