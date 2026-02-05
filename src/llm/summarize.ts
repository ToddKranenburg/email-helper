import OpenAI from 'openai';
import { formatUserIdentity, type UserIdentity } from './userContext.js';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

/**
 * Taxonomy (exactly as provided):
 * - Marketing/Promotion: bulk emails, promos, newsletters, invites to marketing/promo events etc. Sender often corporate/branded (marketing@..., info@...).
 * - Personal Event: non promotional — personal invites or events I've actively decided to attend; scheduling around a personal event.
 * - Billing: need to pay, receipt, payment failed, upcoming payment, refunds, invoices.
 * - Introduction: someone personally introducing me to another person.
 * - Catch Up: “checking in / long time / what’s new” from a person.
 * - Editorial/Writing: news content, Substacks, articles, essays being shared.
 * - Personal Request: requests for help or favors (non-scheduling).
 * - FYI: personal info updates that are not events.
 */
const TAXONOMY = [
  'Marketing/Promotion',
  'Personal Event',
  'Billing',
  'Introduction',
  'Catch Up',
  'Editorial/Writing',
  'Personal Request',
  'FYI'
] as const;

type Category = typeof TAXONOMY[number];
const ALLOWED = new Set<string>(TAXONOMY);

/** System prompt tailored to your taxonomy + headline requirement. */
const SYSTEM = `You summarize email threads for a busy professional.

OUTPUT: Return STRICT JSON with keys exactly: headline, tldr, category, next_step, confidence.

CONSTRAINTS:
- "headline": 3–7 words, max-clarity, no emojis, no trailing punctuation. Prefer imperative or crisp noun phrase (e.g., "Confirm Friday Meeting", "Invoice for November").
- "category" MUST be one of:
  -> Marketing/Promotion: bulk emails, promos, newsletters, invites to marketing/promo events etc. the sender is a good clue -- is it corporate/branded? marketing@... info@...?
  -> Personal Event: non promotional -- should be personal, or an event i've actively decided to attend
  -> Billing: need to pay, receipt, payment failed, upcoming payment, etc...
  -> Introduction: someone personally introducing me to another person
  -> Catch Up: “checking in / long time / what’s new” from a person
  -> Editorial/Writing: news content, Substacks, articles, essays being shared
  -> Personal Request: requests for help or favors
  -> FYI: personal info updates that are not events
- "tldr": 1–2 lines, concise, no emojis.
- "next_step": a short imperative ("RSVP by Friday", "Reply with availability", "No action").
  Use "No action" when nothing is required.
- "confidence": High | Medium | Low (your confidence in category and next_step).
- The user's identity is provided; treat messages from the user's email as sent by the user, and treat mentions of their name/email as referring to the user.`;

/** Extract JSON even if wrapped in fences or with noise. */
function extractJSONObject(s: string) {
  const fenced = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced ? fenced[1] : s;
  const brace = candidate.match(/\{[\s\S]*\}/);
  const jsonStr = brace ? brace[0] : candidate;
  return JSON.parse(jsonStr);
}

function titleCaseWords(s: string) {
  return String(s || '').trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .replace(/\b([a-z])/g, (_, c) => c.toUpperCase());
}

function cleanHeadline(h: string) {
  let x = String(h || '').trim();
  // Remove emojis and trailing punctuation
  x = x.replace(/\p{Extended_Pictographic}/gu, '').trim();
  x = x.replace(/[.!?;:]+$/g, '').trim();
  // Collapse spaces
  x = x.replace(/\s+/g, ' ').trim();
  // Limit to ~7 words, keep clarity
  const words = x.split(' ');
  if (words.length > 7) x = words.slice(0, 7).join(' ');
  // Title-ish case, but preserve common small words
  const small = new Set(['a','an','the','and','or','of','to','in','on','for','with','at','by','from']);
  x = x.split(' ').map((w, i) => {
    const lw = w.toLowerCase();
    if (i > 0 && small.has(lw)) return lw;
    return lw.charAt(0).toUpperCase() + lw.slice(1);
  }).join(' ');
  return x || '(No Headline)';
}

/** Normalize model output and guardrails. */
function normalize(obj: any) {
  const o = {
    headline: cleanHeadline(obj.headline || ''),
    tldr: String(obj.tldr || '').trim(),
    category: titleCaseWords(String(obj.category || '')),
    next_step: String(obj.next_step || 'No action').trim(),
    confidence: titleCaseWords(String(obj.confidence || 'Low').trim())
  };

  if (!o.tldr) o.tldr = '(No summary)';
  if (!ALLOWED.has(o.category)) o.category = 'FYI';

  // Normalize confidence to High/Medium/Low (UI appends "Confidence")
  const lc = o.confidence.toLowerCase();
  o.confidence = lc.startsWith('high') ? 'High' : lc.startsWith('med') ? 'Medium' : 'Low';

  return o;
}

/** Heuristic corrections for obvious mislabels. */
function correctCategory(input: {
  subject: string; convoText: string; modelCategory: string;
}): Category {
  const subject = (input.subject || '').toLowerCase();
  const text = (input.convoText || '').toLowerCase();
  const blob = `${subject}\n${text}`;

  // Marketing/Promotion
  if (/\bunsubscribe\b|\bview this email in your browser\b|\bmanage preferences\b|\bprivacy policy\b|\bupdate preferences\b/.test(blob)) {
    return 'Marketing/Promotion';
  }

  // Billing
  if (/\binvoice\b|\breceipt\b|\bpayment\b|\bpaid\b|\brefunded?\b|\bbilling\b|\bcharge(?:d|s)?\b|\bfailed payment\b|\bpast due\b/.test(blob)) {
    return 'Billing';
  }

  // Personal Event (invites, RSVPs, tickets, calendars, meetups)
  if (
    /\bevent\b|\brsvp\b|\bticket\b|\btickets\b|\bvenue\b|\bconcert\b|\bbirthday\b|\bparty\b|\bwebinar\b|\bmeetup\b|\breservation\b|\bbook(?:ing)?\b/.test(blob) ||
    /\bcalendar\b|\bics\b|\binvite\b|\bhold the date\b|\bsave the date\b|\bavailability\b|\bpick a time\b|\bschedule\b|\breschedule\b|\bzoom\b|\bmeet\b|\bcall\b/.test(blob)
  ) {
    return 'Personal Event';
  }

  // Introduction
  if (/\bintro(?:duction)?\b|\bintroduce\b|\bconnecting you\b|\bconnect you\b|\bmeet (?:my|our)\b/.test(blob)) {
    return 'Introduction';
  }

  // Editorial/Writing
  if (
    /\barticle\b|\bessay\b|\bop[- ]?ed\b|\bread more\b|\bread the full\b|\bcolumn\b|\bblog\b/.test(blob) ||
    /\bsubstack\.com\b|\bmedium\.com\b|\bnytimes\.com\b|\bnewyorker\.com\b|\bbloomberg\.com\b|\btheatlantic\.com\b|\bft\.com\b|\bwired\.com\b/.test(blob)
  ) {
    return 'Editorial/Writing';
  }

  // Personal Request (non-scheduling favors/asks)
  const isScheduling =
    /\bavailability\b|\bpick a time\b|\bschedule\b|\breschedule\b|\bcalendar\b|\bzoom\b|\bmeet\b|\bcall\b/.test(blob);
  const requestPhrases = [
    'can you','could you','would you','would you mind','please can you','pls can you',
    'do you mind','i need your help','favor','help me','review this','take a look',
    'share feedback','give feedback','look over','send me','forward me','introduce me to','cover me','pick up','grab','borrow'
  ];
  if (!isScheduling && requestPhrases.some(p => blob.includes(p))) {
    return 'Personal Request';
  }

  // Catch Up
  if (/\bcatch(?:ing)? up\b|\bcheck(?:ing)? in\b|\blong time\b|\bhow (?:have|are) you\b|\bwhat['’]s new\b|\bbeen a while\b/.test(blob)) {
    return 'Catch Up';
  }

  // FYI or model suggestion if valid
  const normalized = titleCaseWords(input.modelCategory);
  if (ALLOWED.has(normalized)) return normalized as Category;
  return 'FYI';
}

export async function summarize(input: {
  subject: string;
  people: string[];
  convoText: string;
  user?: UserIdentity | null;
}) {
  const userIdentity = formatUserIdentity(input.user);
  const user = `Subject: ${input.subject}
Participants: ${input.people.join(', ')}
${userIdentity}

Thread (oldest→newest):
${input.convoText.slice(0, 6000)}

Respond ONLY as a JSON object with keys: headline, tldr, category, next_step, confidence.`;

  const resp = await openai.chat.completions.create({
    model: 'gpt-4o-mini',
    messages: [{ role: 'system', content: SYSTEM }, { role: 'user', content: user }],
    temperature: 0.0,
    response_format: { type: 'json_object' as const }
  });

  const raw = resp.choices[0]?.message?.content ?? '{}';

  let base = normalize(extractJSONObject(raw));

  // Heuristic category correction
  base.category = correctCategory({
    subject: input.subject || '',
    convoText: input.convoText || '',
    modelCategory: base.category
  });

  return base;
}
