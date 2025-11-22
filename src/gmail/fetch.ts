import pLimit from 'p-limit';
import { gmailClient, INBOX_QUERY } from './client.js';
import { normalizeBody } from './normalize.js';
import { getAuthedClient } from '../auth/google.js';
import { prisma } from '../store/db.js';
import type { Request } from 'express';
import type { gmail_v1 } from 'googleapis';
import type { GaxiosResponse } from 'gaxios';

const THREAD_BATCH_SIZE = 20;

function parseFrom(headerValue: string | null | undefined): { name?: string; email?: string } {
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
  const session = req.session as any;
  const user = session?.user;
  if (!user?.id) throw new Error('User session missing during ingest');

  const auth = getAuthedClient(req.session);
  const gmail: gmail_v1.Gmail = gmailClient(auth);

  let pageToken: string | undefined = undefined;
  const limit = pLimit(6);

  do {
    // IMPORTANT: rely on Gmail search "category:primary" to mirror the UI
    const list: GaxiosResponse<gmail_v1.Schema$ListThreadsResponse> = await gmail.users.threads.list({
      userId: 'me',
      q: INBOX_QUERY,
      pageToken,
      maxResults: THREAD_BATCH_SIZE
      // NOTE: no labelIds filter here; Gmail's search operator is the source of truth
    });

    const threads: gmail_v1.Schema$Thread[] = (list.data.threads || []).filter(
      (thread: gmail_v1.Schema$Thread | null | undefined): thread is gmail_v1.Schema$Thread =>
        Boolean(thread && thread.id)
    );

    await Promise.all(
      threads.map((thread: gmail_v1.Schema$Thread) =>
        limit(async () => {
          const full = await gmail.users.threads.get({ userId: 'me', id: thread.id! });
          const msgs = (full.data.messages || [])
            .filter((msg: gmail_v1.Schema$Message | null | undefined): msg is gmail_v1.Schema$Message => Boolean(msg))
            .slice(-3);
          const latest = msgs[msgs.length - 1];
          if (!latest) return;

          const hdrs = latest.payload?.headers || [];
          const findHeader = (name: string) =>
            hdrs.find((h): h is gmail_v1.Schema$MessagePartHeader & { name: string } => Boolean(h?.name) && h.name === name);
          const subject = findHeader('Subject')?.value || '';
          const fromRaw = findHeader('From')?.value;
          const { name: fromName, email: fromEmail } = parseFrom(fromRaw);
          const latestMsgId = latest.id;
          if (!latestMsgId) return;

          const participants = Array.from(
            new Set(
              hdrs
                .filter((h): h is gmail_v1.Schema$MessagePartHeader & { name: string } => {
                  if (!h?.name) return false;
                  return ['From', 'To', 'Cc'].includes(h.name);
                })
                .flatMap(h => (h.value?.split(',') || []))
                .map(s => s.trim())
                .filter(Boolean)
            )
          );

          const threadId = full.data.id!;
          if (!threadId) return;

          await prisma.thread.upsert({
            where: { id_userId: { id: threadId, userId: user.id } },
            update: {
              subject,
              participants: JSON.stringify(participants),
              lastMessageTs: new Date(latest.internalDate ? Number(latest.internalDate) : Date.now()),
              historyId: full.data.historyId,
              fromName,
              fromEmail,
              userId: user.id
            },
            create: {
              id: threadId,
              userId: user.id,
              subject,
              participants: JSON.stringify(participants),
              lastMessageTs: new Date(latest.internalDate ? Number(latest.internalDate) : Date.now()),
              historyId: full.data.historyId,
              fromName,
              fromEmail
            }
          });

          // Skip if we've already summarized this latest message
          const existing = await prisma.summary.findUnique({
            where: {
              userId_lastMsgId: {
                userId: user.id,
                lastMsgId: latestMsgId
              }
            }
          });
          if (existing) return;

          const convoText = msgs
            .map(m => normalizeBody(m.payload))
            .filter(Boolean)
            .reverse()
            .join('\n\n---\n\n');

          await summarizeAndStore(threadId, latestMsgId, subject, participants, convoText, user.id);
        })
      )
    );

    pageToken = list.data.nextPageToken || undefined;
  } while (pageToken);
}

import { summarize } from '../llm/summarize.js';
async function summarizeAndStore(
  threadId: string,
  lastMsgId: string,
  subject: string,
  people: string[],
  convoText: string,
  userId: string
) {
  const s = await summarize({ subject, people, convoText });
  await prisma.summary.create({
    data: {
      threadId,
      userId,
      lastMsgId,
      headline: (s as any).headline || '',
      tldr: s.tldr,
      category: s.category,
      nextStep: s.next_step,
      convoText,
      confidence: s.confidence
    }
  });
}
