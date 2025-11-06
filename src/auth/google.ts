import { Router } from 'express';
import { google } from 'googleapis';

export const authRouter = Router();

const oauth2Client = new google.auth.OAuth2(
  process.env.GOOGLE_CLIENT_ID!,
  process.env.GOOGLE_CLIENT_SECRET!,
  process.env.GOOGLE_REDIRECT_URI!
);

const SCOPES = [
  'https://www.googleapis.com/auth/gmail.readonly',
  'openid', 'email', 'profile'
];

authRouter.get('/google', (req, res) => {
  const url = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent'
  });
  res.redirect(url);
});

authRouter.get('/google/callback', async (req, res) => {
  const { code } = req.query as { code: string };
  const { tokens } = await oauth2Client.getToken(code);
  (req.session as any).googleTokens = tokens;
  res.redirect('/dashboard');
});

export function getAuthedClient(sessionObj: any) {
  const client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID!,
    process.env.GOOGLE_CLIENT_SECRET!,
    process.env.GOOGLE_REDIRECT_URI!
  );
  client.setCredentials(sessionObj.googleTokens);
  return client;
}
