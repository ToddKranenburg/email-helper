export type IngestState = {
  status: 'idle' | 'running' | 'done' | 'error';
  updatedAt: number;
  error?: string;
};

const ingestStatus = new Map<string, IngestState>();

export function markIngestStatus(userId: string, status: IngestState['status'], error?: string) {
  ingestStatus.set(userId, { status, updatedAt: Date.now(), error });
  if (status === 'done' || status === 'error') {
    setTimeout(() => ingestStatus.delete(userId), 5 * 60 * 1000);
  }
}

export function getIngestStatus(userId: string) {
  return ingestStatus.get(userId) || { status: 'idle', updatedAt: Date.now() };
}
