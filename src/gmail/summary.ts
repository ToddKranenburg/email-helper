import type { gmail_v1 } from 'googleapis';
import type { OAuth2Client } from 'google-auth-library';
import { gmailClient } from './client.js';
import { normalizeBody } from './normalize.js';
import { prisma } from '../store/db.js';
import type { Summary, ThreadIndex } from '@prisma/client';
import { summarize } from '../llm/summarize.js';
import { ensureAutoSummaryCards } from '../actions/persistence.js';

const SUMMARY_MESSAGE_LIMIT = 3;

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

function headerValue(headers: gmail_v1.Schema$MessagePartHeader[] | undefined, name: string) {
  if (!headers?.length) return null;
  const match = headers.find(h => h?.name === name);
  return match?.value ?? null;
}

function collectParticipants(headersList: gmail_v1.Schema$MessagePartHeader[][]): string[] {
  const participants = new Set<string>();
  headersList.forEach(headers => {
    for (const name of ['From', 'To', 'Cc']) {
      const raw = headerValue(headers, name);
      if (!raw) continue;
      raw.split(',')
        .map(item => item.trim())
        .filter(Boolean)
        .forEach(value => participants.add(value));
    }
  });
  return Array.from(participants);
}

function buildContentVersion(lastMessageDate: Date | null, lastMessageId: string | null) {
  if (!lastMessageDate) return null;
  const stamp = lastMessageDate.toISOString();
  return lastMessageId ? `${stamp}:${lastMessageId}` : stamp;
}

export async function ensureThreadSummary(
  auth: OAuth2Client,
  userId: string,
  threadId: string,
  opts: { forceFresh?: boolean } = {}
): Promise<(Summary & { threadIndex: ThreadIndex | null }) | null> {
  const existing = await prisma.summary.findFirst({
    where: { userId, threadId },
    include: { threadIndex: true }
  });

  const threadIndex = existing?.threadIndex || await prisma.threadIndex.findUnique({
    where: { threadId_userId: { threadId, userId } }
  });

  const latestMessageId = threadIndex?.lastMessageId || null;
  if (existing && !opts.forceFresh && latestMessageId && existing.lastMsgId === latestMessageId) {
    return existing;
  }

  const gmail = gmailClient(auth);
  const full = await gmail.users.threads.get({ userId: 'me', id: threadId });
  const messages = (full.data.messages || []).filter(
    (msg: gmail_v1.Schema$Message | null | undefined): msg is gmail_v1.Schema$Message => Boolean(msg)
  );
  if (!messages.length) return existing;

  const sorted = messages.slice().sort((a, b) => {
    const aDate = a.internalDate ? Number(a.internalDate) : 0;
    const bDate = b.internalDate ? Number(b.internalDate) : 0;
    return aDate - bDate;
  });
  const recent = sorted.slice(-SUMMARY_MESSAGE_LIMIT);
  const latest = recent[recent.length - 1];
  if (!latest?.id) return existing;

  const headersList = recent.map(msg => msg.payload?.headers || []);
  const latestHeaders = latest.payload?.headers || [];
  const subject = headerValue(latestHeaders, 'Subject') || threadIndex?.subject || '';
  const fromRaw = headerValue(latestHeaders, 'From');
  const { name: fromName, email: fromEmail } = parseFrom(fromRaw);
  const participants = collectParticipants(headersList);
  const unreadCount = recent.filter(msg => msg.labelIds?.includes('UNREAD')).length;
  const labelIds = latest.labelIds || [];
  const inPrimaryInbox = labelIds.includes('INBOX') && labelIds.includes('CATEGORY_PRIMARY') && !labelIds.includes('CHAT');

  const convoText = recent
    .map(msg => normalizeBody(msg.payload))
    .filter(Boolean)
    .reverse()
    .join('\n\n---\n\n');

  if (!threadIndex) {
    const lastMessageDate = new Date(latest.internalDate ? Number(latest.internalDate) : Date.now());
    await prisma.threadIndex.create({
      data: {
        threadId,
        userId,
        subject: subject || undefined,
        participants: participants.length ? JSON.stringify(participants) : undefined,
        fromName: fromName ?? undefined,
        fromEmail: fromEmail ?? undefined,
        lastMessageId: latest.id,
        lastMessageDate,
        unreadCount,
        snippet: full.data.snippet ?? undefined,
        gmailLabelIds: labelIds,
        inPrimaryInbox,
        contentVersion: buildContentVersion(lastMessageDate, latest.id)
      }
    });
  }

  const existingLatest = await prisma.summary.findUnique({
    where: {
      userId_lastMsgId: {
        userId,
        lastMsgId: latest.id
      }
    }
  });
  if (existingLatest && !opts.forceFresh) {
    return existingLatest;
  }

  const s = await summarize({ subject, people: participants, convoText });
  const created = await prisma.summary.create({
    data: {
      threadId,
      userId,
      lastMsgId: latest.id,
      headline: (s as any).headline || '',
      tldr: s.tldr,
      category: s.category,
      nextStep: s.next_step,
      convoText,
      confidence: s.confidence
    }
  });

  try {
    await ensureAutoSummaryCards({
      userId,
      threadId,
      lastMessageId: latest.id,
      subject,
      headline: (s as any).headline || '',
      summary: s.tldr,
      nextStep: s.next_step,
      participants,
      transcript: convoText
    });
  } catch (err) {
    console.error('failed to build auto summary cards', err);
  }

  await prisma.summary.deleteMany({
    where: {
      userId,
      threadId,
      id: { not: created.id }
    }
  });

  const lastMessageDate = new Date(latest.internalDate ? Number(latest.internalDate) : Date.now());
  await prisma.threadIndex.update({
    where: { threadId_userId: { threadId, userId } },
    data: {
      subject: subject || undefined,
      participants: participants.length ? JSON.stringify(participants) : undefined,
      fromName: fromName ?? undefined,
      fromEmail: fromEmail ?? undefined,
      lastMessageId: latest.id,
      lastMessageDate,
      unreadCount,
      gmailLabelIds: labelIds,
      contentVersion: buildContentVersion(lastMessageDate, latest.id)
    }
  }).catch(() => null);

  return prisma.summary.findFirst({
    where: { id: created.id },
    include: { threadIndex: true }
  });
}
