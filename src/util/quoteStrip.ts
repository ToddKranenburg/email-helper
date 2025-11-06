const QUOTE_MARKERS = [
  /^On .* wrote:$/i,
  /^From:\s/i,
  /^Sent:\s/i,
  /^Subject:\s/i,
  /^-----Original Message-----$/i
];

export function stripQuoted(text: string): string {
  const lines = text.split('\n');
  const cut = lines.findIndex(l => QUOTE_MARKERS.some(rx => rx.test(l.trim())));
  const head = (cut >= 0 ? lines.slice(0, cut) : lines)
    .filter(l => !l.match(/^>+/)) // remove quoted prefix
    .join('\n');
  return head.replace(/\n{3,}/g, '\n\n').trim();
}
