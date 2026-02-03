import { getAuthedClient } from '../auth/google.js';
import type { OAuth2Client } from 'google-auth-library';
import { syncPrimaryInbox } from './sync.js';

export async function ingestInbox(session: any, opts: { skipPriorityEnqueue?: boolean } = {}) {
  const user = session?.user;
  if (!user?.id) throw new Error('User session missing during ingest');
  const auth = getAuthedClient(session);
  return ingestInboxWithClient(auth, user.id, opts);
}

export async function ingestInboxWithClient(
  auth: OAuth2Client,
  userId: string,
  opts: { skipPriorityEnqueue?: boolean } = {}
) {
  return syncPrimaryInbox(auth, userId, opts);
}
