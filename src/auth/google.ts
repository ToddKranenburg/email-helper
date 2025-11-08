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

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'openid', 'email', 'profile'
];

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
  const oauth2 = google.oauth2('v2');
  const { data: profile } = await oauth2.userinfo.get({ auth: oauthClient });

  const userId = profile.id || (profile as { sub?: string }).sub;
  if (!userId || !profile.email) {
    return res.status(400).send('Unable to retrieve Google profile information.');
  }

  const userPayload = {
    id: userId,
    email: profile.email,
    name: profile.name || null,
    picture: profile.picture || null
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
