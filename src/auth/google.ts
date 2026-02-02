import { Router, Request, Response } from 'express';
import { google } from 'googleapis';
import { prisma } from '../store/db.js';

export const authRouter = Router();

function createOAuthClient() {
  return new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID!,
    process.env.GOOGLE_CLIENT_SECRET!,
    process.env.GOOGLE_REDIRECT_URI!
  );
}

const sharedClient = createOAuthClient();

export const GMAIL_SCOPES = [
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/gmail.send',
  'https://www.googleapis.com/auth/tasks'
] as const;

const PROFILE_SCOPES = [
  'openid',
  'email',
  'profile'
] as const;

function parseScopes(raw: string | undefined | null): string[] {
  if (!raw) return [];
  return raw
    .split(/[,\s]+/)
    .map(s => s.trim())
    .filter(Boolean);
}

const ENV_SCOPES = parseScopes(process.env.GOOGLE_SCOPES);
const REQUESTED_SCOPES = Array.from(new Set([...GMAIL_SCOPES, ...PROFILE_SCOPES, ...ENV_SCOPES]));
const SCOPES = [...REQUESTED_SCOPES];

export class MissingScopeError extends Error {
  missingScopes: string[];

  constructor(missingScopes: string[]) {
    super('Missing required Google scopes');
    this.name = 'MissingScopeError';
    this.missingScopes = missingScopes;
  }
}

function parseGrantedScopes(raw: unknown): string[] {
  if (Array.isArray(raw)) return raw.filter(Boolean).map(String);
  if (typeof raw === 'string') {
    return raw.split(/\s+/).map(s => s.trim()).filter(Boolean);
  }
  return [];
}

export function getMissingGmailScopes(tokens: { scope?: string | string[] } | null | undefined): string[] {
  const grantedScopes = parseGrantedScopes(tokens?.scope);
  if (!grantedScopes.length) return [...REQUESTED_SCOPES];
  const scopeSet = new Set(grantedScopes);
  const hasAny = (values: string[]) => values.some(value => scopeSet.has(value));
  return REQUESTED_SCOPES.filter(scope => {
    if (scope === 'email') {
      return !hasAny(['email', 'https://www.googleapis.com/auth/userinfo.email']);
    }
    if (scope === 'profile') {
      return !hasAny(['profile', 'https://www.googleapis.com/auth/userinfo.profile']);
    }
    return !scopeSet.has(scope);
  });
}

authRouter.get('/google', (req: Request, res: Response) => {
  const url = sharedClient.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent',
    include_granted_scopes: true
  });
  res.redirect(url);
});

authRouter.get('/google/callback', async (req: Request, res: Response) => {
  const { code } = req.query as { code: string };
  const oauthClient = createOAuthClient();
  const { tokens } = await oauthClient.getToken(code);
  oauthClient.setCredentials(tokens);

  let missingScopes = getMissingGmailScopes(tokens);
  if (missingScopes.length && tokens.access_token) {
    try {
      const tokenInfo = await oauthClient.getTokenInfo(tokens.access_token);
      const infoScopes = Array.isArray(tokenInfo?.scopes)
        ? tokenInfo.scopes.map((s: string) => s.trim()).filter(Boolean)
        : [];
      if (infoScopes.length) {
        tokens.scope = infoScopes.join(' ');
        missingScopes = getMissingGmailScopes(tokens);
      }
    } catch (err) {
      console.warn('Failed to verify token scopes', err);
    }
  }
  if (missingScopes.length) {
    console.error('User did not grant the required Gmail scopes', {
      missingScopes,
      grantedScopes: tokens.scope
    });
    return res
      .status(403)
      .send('Google did not grant Gmail access. Please remove the app from your Google Account permissions and try again.');
  }

  const gmail = google.gmail({ version: 'v1', auth: oauthClient });
  const { data: gmailProfile } = await gmail.users.getProfile({ userId: 'me' });
  const email = gmailProfile.emailAddress;
  if (!email) {
    return res.status(400).send('Unable to retrieve Gmail profile information.');
  }

  let profileName: string | null = null;
  let profilePicture: string | null = null;
  try {
    const oauth2 = google.oauth2({ version: 'v2', auth: oauthClient });
    const { data: userInfo } = await oauth2.userinfo.get();
    profileName = typeof userInfo?.name === 'string' ? userInfo.name.trim() : null;
    profilePicture = typeof userInfo?.picture === 'string' ? userInfo.picture.trim() : null;
  } catch (err) {
    console.warn('Unable to fetch Google profile info', err);
  }

  const existingUser = await prisma.user.findUnique({ where: { email } });
  const userId = existingUser?.id ?? email;

  const userPayload = {
    id: userId,
    email,
    name: profileName || (existingUser?.name ?? null),
    picture: profilePicture || (existingUser?.picture ?? null)
  };

  await prisma.user.upsert({
    where: { id: userId },
    update: {
      email: userPayload.email,
      name: userPayload.name ?? undefined,
      picture: userPayload.picture ?? undefined
    },
    create: userPayload
  });

  await upsertGoogleToken(userId, tokens);

  (req.session as any).googleTokens = tokens;
  (req.session as any).user = userPayload;
  res.redirect('/dashboard');
});

export function getAuthedClient(sessionObj: any) {
  if (!sessionObj?.googleTokens) {
    throw new Error('Missing Google tokens on session');
  }
  const missingScopes = getMissingGmailScopes(sessionObj.googleTokens);
  if (missingScopes.length) {
    throw new MissingScopeError(missingScopes);
  }
  const client = createOAuthClient();
  client.setCredentials(sessionObj.googleTokens);
  return client;
}

type StoredToken = {
  refreshToken: string;
  accessToken?: string | null;
  scope?: string | null;
  tokenType?: string | null;
  expiryDate?: Date | null;
};

export function getAuthedClientFromStoredToken(userId: string, token: StoredToken) {
  const client = createOAuthClient();
  client.setCredentials({
    refresh_token: token.refreshToken,
    access_token: token.accessToken ?? undefined,
    scope: token.scope ?? undefined,
    token_type: token.tokenType ?? undefined,
    expiry_date: token.expiryDate ? token.expiryDate.getTime() : undefined
  });
  attachTokenListener(client, userId);
  return client;
}

async function upsertGoogleToken(
  userId: string,
  tokens: {
    refresh_token?: string | null;
    access_token?: string | null;
    scope?: string | null;
    token_type?: string | null;
    expiry_date?: number | null;
  }
) {
  const existing = await prisma.googleToken.findUnique({ where: { userId } });
  const refreshToken = tokens.refresh_token ?? existing?.refreshToken;
  if (!refreshToken) {
    console.warn('No refresh token available for user', { userId });
    return;
  }
  const data = {
    userId,
    refreshToken,
    accessToken: tokens.access_token ?? existing?.accessToken ?? null,
    scope: tokens.scope ?? existing?.scope ?? null,
    tokenType: tokens.token_type ?? existing?.tokenType ?? null,
    expiryDate: tokens.expiry_date
      ? new Date(tokens.expiry_date)
      : existing?.expiryDate ?? null
  };
  if (existing) {
    await prisma.googleToken.update({ where: { userId }, data });
  } else {
    await prisma.googleToken.create({ data });
  }
}

function attachTokenListener(client: ReturnType<typeof createOAuthClient>, userId: string) {
  client.on('tokens', (tokens) => {
    if (!tokens) return;
    const hasUpdates = Boolean(
      tokens.access_token || tokens.refresh_token || tokens.expiry_date || tokens.scope || tokens.token_type
    );
    if (!hasUpdates) return;
    void prisma.googleToken.findUnique({ where: { userId } }).then(existing => {
      if (!existing) return;
      const refreshToken = tokens.refresh_token ?? existing.refreshToken;
      return prisma.googleToken.update({
        where: { userId },
        data: {
          refreshToken,
          accessToken: tokens.access_token ?? existing.accessToken,
          scope: tokens.scope ?? existing.scope,
          tokenType: tokens.token_type ?? existing.tokenType,
          expiryDate: tokens.expiry_date ? new Date(tokens.expiry_date) : existing.expiryDate
        }
      });
    }).catch(err => {
      console.warn('Failed to persist refreshed Google tokens', { userId, error: err instanceof Error ? err.message : err });
    });
  });
}
