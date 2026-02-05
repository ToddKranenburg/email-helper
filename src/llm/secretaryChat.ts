import OpenAI from 'openai';
import type { ChatCompletionMessageParam } from 'openai/resources/chat/completions';
import { formatUserIdentity, type UserIdentity } from './userContext.js';

export type ChatTurn = {
  role: 'user' | 'assistant';
  content: string;
};

export const MAX_CHAT_TURNS = 20;

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

const SYSTEM_PROMPT = `You are a trusted email secretary who answers follow-up questions about a single Gmail thread.
You receive a sanitized transcript of the thread plus the assistant's summary.
Use ONLY the provided transcript as ground truth. Quote relevant lines when clarifying details.
If information is missing from the transcript, say you don't have that detail instead of guessing.
Address the user directly as "you".
The user's identity is provided. Treat messages from the user's email as sent by the user, and treat mentions of their name/email as referring to the user.
Keep responses concise (<=180 words) and helpful. When listing steps, use short bullet points.`;

function buildContext(input: {
  subject: string;
  headline: string;
  summary: string;
  nextStep: string;
  participants: string[];
  transcript: string;
  user?: UserIdentity | null;
}) {
  const participants = input.participants.join(', ') || 'Unknown participants';
  const trimmedTranscript = input.transcript.length > 9000
    ? input.transcript.slice(-9000)
    : input.transcript;
  const userLines = formatUserIdentity(input.user)
    .split('\n')
    .map(line => `- ${line}`)
    .join('\n');

  return `Thread overview:
- Subject: ${input.subject || '(no subject)'}
- Headline: ${input.headline || '(none)'}
- Summary: ${input.summary || '(no summary)'}
- Recommended next step: ${input.nextStep || 'No action'}
- Participants: ${participants}
${userLines}

Transcript (oldest → newest):
${trimmedTranscript || '(Transcript unavailable)'}`;
}

export async function chatAboutEmail(input: {
  subject: string;
  headline: string;
  tldr: string;
  nextStep: string;
  participants: string[];
  transcript: string;
  history: ChatTurn[];
  question: string;
  user?: UserIdentity | null;
}): Promise<string> {
  if (!openai) throw new Error('OpenAI API key is not configured');

  const context = buildContext({
    subject: input.subject,
    headline: input.headline,
    summary: input.tldr,
    nextStep: input.nextStep,
    participants: input.participants,
    transcript: input.transcript,
    user: input.user
  });

  const messages: ChatCompletionMessageParam[] = [
    { role: 'system', content: SYSTEM_PROMPT },
    { role: 'user', content: context },
    { role: 'assistant', content: 'Context noted. Ready for follow-up questions.' }
  ];

  for (const turn of input.history) {
    if (turn.role === 'assistant' || turn.role === 'user') {
      messages.push({ role: turn.role, content: turn.content });
    }
  }
  messages.push({ role: 'user', content: input.question });

  const completion = await openai.chat.completions.create({
    model: 'gpt-4o-mini',
    temperature: 0.2,
    messages
  });

  const reply = completion.choices[0]?.message?.content?.trim();
  if (!reply) throw new Error('No reply from chat model');
  return reply;
}
