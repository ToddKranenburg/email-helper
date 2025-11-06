import { google } from 'googleapis';
import type { OAuth2Client } from 'google-auth-library';

export function gmailClient(oauth: OAuth2Client) {
  return google.gmail({ version: 'v1', auth: oauth });
}

export const INBOX_QUERY =
  "in:inbox newer_than:30d \
  -label:^smartlabel_promo \
  -label:^smartlabel_social \
  -label:^smartlabel_forums \
  -label:^smartlabel_updates \
  -is:chat";