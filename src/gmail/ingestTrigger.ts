import { ingestInbox } from './fetch.js';
import { markIngestStatus } from './ingestStatus.js';
import { MissingScopeError } from '../auth/google.js';
import { GaxiosError } from 'gaxios';

export function triggerBackgroundIngest(sessionData: any, userId: string) {
  markIngestStatus(userId, 'running');
  ingestInbox(sessionData)
    .then(() => {
      markIngestStatus(userId, 'done');
    })
    .catch((err: unknown) => {
      if (err instanceof MissingScopeError) {
        markIngestStatus(userId, 'error', 'Google permissions changed. Please reconnect your Google account.');
        return;
      }
      const gaxios = err instanceof GaxiosError ? err : undefined;
      const message = gaxios?.response?.status === 403
        ? 'Gmail refused to share inbox data. Please reconnect your Google account.'
        : 'Unable to sync your Gmail inbox right now. Please try again.';
      markIngestStatus(userId, 'error', message);
    });
}
