import { google } from 'googleapis';
import type { OAuth2Client } from 'google-auth-library';

export function gmailClient(auth: OAuth2Client) {
  return google.gmail({ version: 'v1', auth });
}

/**
 * Use Gmail's search operator to mirror the UI Primary tab exactly.
 * - "category:primary" tracks what you see in the Primary tab (after Gmail heuristics)
 * - Keep the 30d window and exclude chats
 * - Keep "in:inbox" so archived mail doesn't show
 *
 * You can override via GMAIL_QUERY in .env if desired.
 */
export const INBOX_QUERY =
  process.env.GMAIL_QUERY?.trim() ||
  'in:inbox category:primary newer_than:30d -is:chat';
