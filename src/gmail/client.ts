import { google, type gmail_v1 } from 'googleapis';
import type { OAuth2Client } from 'google-auth-library';

export type GmailClient = gmail_v1.Gmail;

export function gmailClient(auth: OAuth2Client): GmailClient {
  return google.gmail({ version: 'v1', auth });
}

/**
 * Use Gmail's search operator to mirror the UI Primary tab exactly.
 * - "category:primary" tracks what you see in the Primary tab (after Gmail heuristics)
 * - Exclude chats; do not time-limit so the list matches the UI
 * - Keep "in:inbox" so archived mail doesn't show
 *
 * You can override via GMAIL_QUERY in .env if desired.
 */
export const INBOX_QUERY =
  process.env.GMAIL_QUERY?.trim() ||
  'in:inbox category:primary -is:chat';
