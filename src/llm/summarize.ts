import OpenAI from 'openai';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

const SYSTEM = `You summarize email threads for a busy professional.
Return strict JSON with keys: tldr, category, next_step, confidence.
Categories (one): Scheduling | FYI | Approval | Billing | Docs/Links | Intro | Issue.
Keep TL;DR to 1–2 lines. Next step is actionable text or "None".`;

const CATEGORY_SET = new Set([
  'Scheduling','FYI','Approval','Billing','Docs/Links','Intro','Issue'
]);

function extractJSONObject(s: string) {
  // 1) if fenced, grab inside ```...```
  const fence = s.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  const candidate = fence ? fence[1] : s;

  // 2) grab the first {...} block
  const brace = candidate.match(/\{[\s\S]*\}/);
  const jsonStr = brace ? brace[0] : candidate;

  return JSON.parse(jsonStr);
}

function normalize(obj: any) {
  const o = {
    tldr: String(obj.tldr || '').trim(),
    category: String(obj.category || '').trim(),
    next_step: String(obj.next_step || 'None').trim(),
    confidence: String(obj.confidence || 'low').trim()
  };

  // Normalize casing/values
  const catTitle = o.category.replace(/\s+/g, ' ')
    .replace(/(^\w|\s\w)/g, c => c.toUpperCase());
  o.category = CATEGORY_SET.has(catTitle) ? catTitle : 'FYI';

  const conf = o.confidence.toLowerCase();
  o.confidence = conf === 'high' ? 'High' : conf === 'medium' ? 'Medium' : 'Low';

  if (!o.tldr) o.tldr = '(No summary)';
  return o;
}

export async function summarize(input: {
  subject: string;
  people: string[];
  convoText: string;
}) {
  const user = `Subject: ${input.subject}
Participants: ${input.people.join(', ')}

Thread (oldest→newest):
${input.convoText.slice(0, 6000)}

Respond ONLY as a JSON object with keys: tldr, category, next_step, confidence.`;

  const resp = await openai.chat.completions.create({
    model: 'gpt-4o-mini',
    messages: [{ role: 'system', content: SYSTEM }, { role: 'user', content: user }],
    temperature: 0.0,
    // Force JSON mode when available; ignored by older models
    response_format: { type: 'json_object' as const }
  });

  const raw = resp.choices[0]?.message?.content ?? '{}';

  try {
    return normalize(extractJSONObject(raw));
  } catch {
    // Last-resort: make a best-effort TL;DR, don't break the UI
    return {
      tldr: raw.slice(0, 240),
      category: 'FYI',
      next_step: 'None',
      confidence: 'Low'
    };
  }
}
