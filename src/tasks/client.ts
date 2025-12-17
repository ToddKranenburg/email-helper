import { google, type tasks_v1 } from 'googleapis';
import type { OAuth2Client } from 'google-auth-library';

export type TasksClient = tasks_v1.Tasks;

export function tasksClient(auth: OAuth2Client): TasksClient {
  return google.tasks({ version: 'v1', auth });
}
