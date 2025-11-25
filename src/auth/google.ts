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

const GMAIL_SCOPES = [
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/gmail.send'
] as const;

const SCOPES = [...GMAIL_SCOPES];

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

  const grantedScopes = Array.isArray(tokens.scope)
    ? tokens.scope
    : (typeof tokens.scope === 'string' ? tokens.scope.split(/\s+/).filter(Boolean) : []);
  const scopeSet = new Set(grantedScopes);
  const missingScopes = grantedScopes.length ? GMAIL_SCOPES.filter(scope => !scopeSet.has(scope)) : [];
  if (missingScopes.length) {
    console.error('User did not grant the required Gmail scopes', {
      missingScopes,
      grantedScopes
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
  const client = createOAuthClient();
  client.setCredentials(sessionObj.googleTokens);
  return client;
}
