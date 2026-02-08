import OpenAI, { toFile } from 'openai';

const openai = process.env.OPENAI_API_KEY
  ? new OpenAI({ apiKey: process.env.OPENAI_API_KEY })
  : null;

const DEFAULT_TRANSCRIBE_MODEL = process.env.OPENAI_TRANSCRIBE_MODEL || 'gpt-4o-transcribe';

type TranscribeInput = {
  buffer: Buffer;
  filename: string;
  mimeType?: string;
  language?: string;
  prompt?: string;
};

export async function transcribeAudio(input: TranscribeInput) {
  if (!openai) throw new Error('OpenAI not configured');
  const { buffer, filename, mimeType, language, prompt } = input;
  const fallbackExt = mimeType?.includes('webm')
    ? 'webm'
    : mimeType?.includes('mp4')
      ? 'm4a'
      : mimeType?.includes('mpeg')
        ? 'mp3'
        : 'wav';
  const safeName = filename.includes('.') ? filename : `${filename}.${fallbackExt}`;
  const file = await toFile(buffer, safeName);
  const response = await openai.audio.transcriptions.create({
    file,
    model: DEFAULT_TRANSCRIBE_MODEL,
    language: language && language.length <= 8 ? language : undefined,
    prompt: prompt && prompt.length <= 400 ? prompt : undefined,
    response_format: 'json',
    temperature: 0
  });
  return response.text || '';
}
