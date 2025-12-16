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
  'https://www.googleapis.com/auth/gmail.send'
] as const;

function parseScopes(raw: string | undefined | null): string[] {
  if (!raw) return [];
  return raw
    .split(/[,\s]+/)
    .map(s => s.trim())
    .filter(Boolean);
}

const ENV_SCOPES = parseScopes(process.env.GOOGLE_SCOPES);
const REQUESTED_SCOPES = Array.from(new Set([...GMAIL_SCOPES, ...ENV_SCOPES]));
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
  return REQUESTED_SCOPES.filter(scope => !scopeSet.has(scope));
}

authRouter.get('/google', (req: Request, res: Response) => {
  const url = sharedClient.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'
  });
  res.redirect(url);
});

authRouter.get('/google/callback', async (req: Request, res: Response) => {
  const { code } = req.query as { code: string };
  const oauthClient = createOAuthClient();
  const { tokens } = await oauthClient.getToken(code);
  oauthClient.setCredentials(tokens);

  const missingScopes = getMissingGmailScopes(tokens);
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

  const existingUser = await prisma.user.findUnique({ where: { email } });
  const userId = existingUser?.id ?? email;

  const userPayload = {
    id: userId,
    email,
    name: existingUser?.name ?? null,
    picture: existingUser?.picture ?? null
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
