export type UserIdentity = {
  name?: string | null;
  email?: string | null;
};

function normalize(value: string | null | undefined) {
  return typeof value === 'string' ? value.trim() : '';
}

function deriveFirstName(name: string) {
  const trimmed = normalize(name);
  if (!trimmed) return '';
  const first = trimmed.split(/\s+/)[0] || '';
  return first;
}

function deriveEmailHandle(email: string) {
  const trimmed = normalize(email);
  if (!trimmed || !trimmed.includes('@')) return '';
  const local = trimmed.split('@')[0] || '';
  return local.trim();
}

export function formatUserIdentity(user?: UserIdentity | null) {
  const name = normalize(user?.name ?? '');
  const email = normalize(user?.email ?? '');
  const line = name && email
    ? `${name} <${email}>`
    : name
      ? name
      : email
        ? `<${email}>`
        : '(unknown)';

  const aliases = new Set<string>();
  const firstName = deriveFirstName(name);
  const handle = deriveEmailHandle(email);
  if (firstName) aliases.add(firstName);
  if (handle) aliases.add(handle);
  const aliasLine = aliases.size ? `User aliases: ${Array.from(aliases).join(', ')}` : '';

  return aliasLine ? `User: ${line}\n${aliasLine}` : `User: ${line}`;
}
