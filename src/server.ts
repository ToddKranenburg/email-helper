import 'dotenv/config';
import express from 'express';
import session from 'cookie-session';
import { router as web } from './web/routes.js';
import { authRouter } from './auth/google.js';

const app = express();

app.use(session({
  name: 'sess',
  secret: process.env.SESSION_SECRET!,
  maxAge: 7 * 24 * 60 * 60 * 1000
}));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.use('/auth', authRouter);
app.use('/', web);

const port = Number(process.env.PORT) || 3000;
app.listen(port, () => console.log(`http://localhost:${port}`));
