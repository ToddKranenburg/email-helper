import type { gmail_v1 } from 'googleapis';

export type UnsubscribeMetadata = {
  listUnsubscribe: string | null;
  listUnsubscribePost: string | null;
  listId: string | null;
  precedence: string | null;
  unsubscribeUrl: string | null;
  unsubscribeMailto: string | null;
  oneClick: boolean;
  supported: boolean;
  bulk: boolean;
};

export type MailtoPayload = {
  to: string;
  subject: string;
  body: string;
};

export function extractUnsubscribeMetadata(
  headers: gmail_v1.Schema$MessagePartHeader[] | undefined | null
): UnsubscribeMetadata | null {
  if (!headers?.length) return null;
  const listUnsubscribe = headerValue(headers, 'List-Unsubscribe');
  const listUnsubscribePost = headerValue(headers, 'List-Unsubscribe-Post');
  const listId = headerValue(headers, 'List-Id');
  const precedence = headerValue(headers, 'Precedence');

  if (!listUnsubscribe && !listId && !precedence) return null;

  const parsed = parseListUnsubscribe(listUnsubscribe || '');
  const oneClick = listUnsubscribePost
    ? /one-?click/i.test(listUnsubscribePost) || /list-unsubscribe=one-click/i.test(listUnsubscribePost)
    : false;
  const supported = Boolean(parsed.mailto || (parsed.url && oneClick));
  const bulk = Boolean(listId)
    || /bulk|list|junk/i.test(precedence || '')
    || Boolean(listUnsubscribe);

  return {
    listUnsubscribe: listUnsubscribe || null,
    listUnsubscribePost: listUnsubscribePost || null,
    listId: listId || null,
    precedence: precedence || null,
    unsubscribeUrl: parsed.url || null,
    unsubscribeMailto: parsed.mailto || null,
    oneClick,
    supported,
    bulk
  };
}

export function parseMailto(raw: string | null | undefined): MailtoPayload | null {
  const value = typeof raw === 'string' ? raw.trim() : '';
  if (!value) return null;
  const normalized = value.startsWith('mailto:') ? value : `mailto:${value}`;
  try {
    const url = new URL(normalized);
    const to = decodeURIComponent(url.pathname || '').trim();
    if (!to) return null;
    const subject = (url.searchParams.get('subject') || '').trim();
    const body = (url.searchParams.get('body') || '').trim();
    return {
      to,
      subject,
      body
    };
  } catch {
    return null;
  }
}

function parseListUnsubscribe(raw: string): { mailto: string | null; url: string | null } {
  const cleaned = String(raw || '').trim();
  if (!cleaned) return { mailto: null, url: null };
  const tokens: string[] = [];
  const matches = cleaned.match(/<[^>]+>/g);
  if (matches?.length) {
    matches.forEach(match => {
      const val = match.replace(/[<>]/g, '').trim();
      if (val) tokens.push(val);
    });
  } else {
    cleaned.split(',').forEach(chunk => {
      const val = chunk.trim();
      if (val) tokens.push(val);
    });
  }
  let mailto: string | null = null;
  let url: string | null = null;
  for (const token of tokens) {
    const value = token.trim();
    if (!value) continue;
    if (!mailto && value.toLowerCase().startsWith('mailto:')) {
      mailto = value;
    } else if (!url && /^https?:/i.test(value)) {
      url = value;
    }
  }
  return { mailto, url };
}

function headerValue(headers: gmail_v1.Schema$MessagePartHeader[], name: string) {
  const target = name.toLowerCase();
  const match = headers.find(h => typeof h?.name === 'string' && h.name.toLowerCase() === target);
  return match?.value ?? null;
}
