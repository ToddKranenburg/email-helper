import OpenAI from 'openai';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

/**
 * Custom taxonomy (exactly as requested):
 * - Event — concerts, birthdays, invitations, scheduling around an event, reservations, tickets
 * - Billing — invoices, receipts, failed payments, refunds
 * - Introduction — connecting people / intros
 * - Marketing/Promotion — company blasts, promos, “view in browser”, unsubscribe footers
 * - FYI — personal info updates that are not events
 * - Catch Up — “checking in / long time / what’s new” from a person
 * - Editorial/Writing — news content, Substacks, articles, essays being shared
 */
const TAXONOMY = [
  'Personal Event',
  'Billing',
  'Introduction',
  'Marketing/Promotion',
  'FYI',
  'Catch Up',
  'Editorial/Writing',
  'Personal Request'
] as const;

type Category = typeof TAXONOMY[number];
const ALLOWED = new Set<string>(TAXONOMY);

/** System prompt tailored to the taxonomy and strict JSON output. */
const SYSTEM = `You summarize email threads for a busy professional.

OUTPUT: Return STRICT JSON with keys exactly: tldr, category, next_step, confidence.

CONSTRAINTS:
- "category" MUST be one of:
  -> Marketing/Promotion: bulk emails, promos, newsletters, invites to marketing/promo events etc. the sender is a good clue -- is it corporate/branded? marketing@... info@...?
  -> Personal Event: non promotional -- should be personal, or an event i've actively decided to attend
  -> Billing: need to pay, receipt, payment failed, upcoming payment, etc...
  -> Introduction: someone personally introducing me to another person
  -> Catch Up: “checking in / long time / what’s new” from a person
  -> Editorial/Writing: news content, Substacks, articles, essays being shared
  -> Personal Request: requests for help or favors
  -> FYI: personal info updates that are not events
- TL;DR: 1–2 lines, concise, no emojis.
- "next_step": a short imperative ("RSVP by Friday", "Reply with availability", "No action").
  Use "No action" when nothing is required.
- "confidence": High | Medium | Low (your confidence in category and next_step).`;

/** Extract JSON even if wrapped in fences or with noise. */
function extractJSONObject(s: string) {
  const fenced = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced ? fenced[1] : s;
  const brace = candidate.match(/\{[\s\S]*\}/);
  const jsonStr = brace ? brace[0] : candidate;
  return JSON.parse(jsonStr);
}

function titleCase(s: string) {
  return String(s || '').toLowerCase().replace(/\b([a-z])/g, (_, c) => c.toUpperCase());
}

/** Normalize model output and guardrails. */
function normalize(obj: any) {
  const o = {
    tldr: String(obj.tldr || '').trim(),
    category: titleCase(String(obj.category || '')),
    next_step: String(obj.next_step || 'No action').trim(),
    confidence: titleCase(String(obj.confidence || 'Low').trim())
  };

  if (!ALLOWED.has(o.category)) o.category = 'FYI';
  if (!o.tldr) o.tldr = '(No summary)';

  // Normalize confidence to one of High/Medium/Low for storage (UI adds "Confidence")
  const lc = o.confidence.toLowerCase();
  o.confidence = lc.startsWith('high') ? 'High' : lc.startsWith('med') ? 'Medium' : 'Low';

  return o;
}

/** Heuristic corrections based on subject + body + participants. */
function correctCategory(input: {
  subject: string;
  convoText: string;
  people: string[];
  modelCategory: string;
}): Category {
  const subject = (input.subject || '').toLowerCase();
  const text = (input.convoText || '').toLowerCase();
  const blob = `${subject}\n${text}`;

  // --- Marketing/Promotion ---
  // Strong indicators of bulk/marketing mail
  if (
    /\bunsubscribe\b|\bview this email in your browser\b|\bmanage preferences\b|\bprivacy policy\b|\bupdate preferences\b/.test(
      blob
    )
  ) {
    return 'Marketing/Promotion';
  }

  // --- Billing ---
  if (
    /\binvoice\b|\breceipt\b|\bpayment\b|\bpaid\b|\brefunded?\b|\bbilling\b|\bcharge(?:d|s)?\b|\bfailed payment\b|\bpast due\b/.test(
      blob
    )
  ) {
    return 'Billing';
  }

  // --- Personal Event ---
  // Invitations, RSVPs, tickets, reservations, calendars, parties, webinars, meetups
  if (
    /\bevent\b|\brsvp\b|\bticket\b|\btickets\b|\bvenue\b|\bconcert\b|\bbirthday\b|\bparty\b|\bwebinar\b|\bmeetup\b|\breservation\b|\bbook(?:ing)?\b/.test(
      blob
    ) ||
    /\bcalendar\b|\bics\b|\binvite\b|\bhold the date\b|\bsave the date\b|\bavailability\b|\bpick a time\b|\bschedule\b|\breschedule\b/.test(
      blob
    )
  ) {
    return 'Personal Event';
  }

    // --- Personal Request ---
  // Direct favor/ask from a person, excluding scheduling-specific phrases (those are Event).
  const requestPhrases = [
    'can you', 'could you', 'would you', 'would you mind', 'please can you', 'pls can you',
    'do you mind', 'i need your help', 'favor', 'help me', 'can u', 'could u',
    'review this', 'take a look', 'share feedback', 'give feedback', 'look over',
    'send me', 'forward me', 'introduce me to', 'cover me', 'pick up', 'grab', 'borrow'
  ];
  const isSchedulingCue =
    /\bavailability\b|\bpick a time\b|\bschedule\b|\breschedule\b|\bcalendar\b|\bzoom\b|\bmeet\b|\bcall\b/.test(blob);
  if (!isSchedulingCue && requestPhrases.some(p => blob.includes(p))) {
    return 'Personal Request';
  }

  // --- Introduction ---
  // Connecting people signals
  if (
    /\bintro(?:duction)?\b|\bintroduce\b|\bconnecting you\b|\bconnect you\b|\bmeet (?:my|our|x)\b|\bcc['’']?ing\b/.test(
      blob
    )
  ) {
    return 'Introduction';
  }

   // --- Catch Up ---
  if (
    /\bcatch(?:ing)? up\b|\bcheck(?:ing)? in\b|\blong time\b|\bhow (?:have|are) you\b|\bwhat['’]s new\b|\bbeen a while\b/.test(
      blob
    )
  ) {
    return 'Catch Up';
  }

  // --- Editorial/Writing ---
  // Articles, essays, Substack, news outlets, "read more"
  if (
    /\barticle\b|\bessay\b|\bop[- ]?ed\b|\bread more\b|\bread the full\b|\bcolumn\b|\bblog\b/.test(blob) ||
    /\bsubstack\.com\b|\bmedium\.com\b|\bnytimes\.com\b|\bnewyorker\.com\b|\bbloomberg\.com\b|\btheatlantic\.com\b|\bft\.com\b|\bwired\.com\b/.test(
      blob
    )
  ) {
    return 'Editorial/Writing';
  }

  // --- FYI ---
  // If it's clearly informative from a person but not an event, default to FYI.
  // (We keep it as the final fallback.)
  const normalized = titleCase(input.modelCategory);
  if (ALLOWED.has(normalized)) return normalized as Category;
  return 'FYI';
}

export async function summarize(input: {
  subject: string;
  people: string[];     // e.g., ["Jane <jane@x.com>", "You <you@x.com>"]
  convoText: string;    // oldest→newest concatenated (already normalized/quote-stripped upstream)
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
    response_format: { type: 'json_object' as const }
  });

  const raw = resp.choices[0]?.message?.content ?? '{}';

  let base = normalize(extractJSONObject(raw));

  // Heuristic correction pass (prevents common mislabels, e.g., survey invitations → Event/Marketing, not Scheduling)
  base.category = correctCategory({
    subject: input.subject || '',
    convoText: input.convoText || '',
    people: input.people || [],
    modelCategory: base.category
  });

  // Ensure confidence formatting consistency (UI will append "Confidence")
  base.confidence = base.confidence.startsWith('H')
    ? 'High'
    : base.confidence.startsWith('M')
    ? 'Medium'
    : 'Low';

  return base;
}
