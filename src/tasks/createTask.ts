import type { OAuth2Client } from 'google-auth-library';
import { tasksClient } from './client.js';

export type CreateTaskInput = {
  title: string;
  notes?: string;
  due?: string;
};

export async function createGoogleTask(auth: OAuth2Client, input: CreateTaskInput) {
  const tasks = tasksClient(auth);
  const due = normalizeDueDate(input.due);
  const task = await tasks.tasks.insert({
    tasklist: '@default',
    requestBody: {
      title: input.title,
      notes: input.notes,
      due: due ?? undefined
    }
  });
  return task.data;
}

export function normalizeDueDate(raw?: string | null) {
  if (!raw) return undefined;
  const trimmed = String(raw).trim();
  if (!trimmed) return undefined;
  const hasTime = trimmed.includes('T');
  const target = hasTime ? trimmed : `${trimmed}T23:59:00Z`;
  const parsed = new Date(target);
  if (Number.isNaN(parsed.getTime())) return undefined;
  return parsed.toISOString();
}
