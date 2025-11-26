import 'dotenv/config';
import express from 'express';
import session from 'cookie-session';
import { router as web } from './web/routes.js';
import { authRouter } from './auth/google.js';
import path from 'path';

const app = express();

app.use((req, res, next) => {
  const start = process.hrtime.bigint();
  res.on('finish', () => {
    const delta = Number(process.hrtime.bigint() - start) / 1_000_000;
    const formatted = delta.toFixed(1);
    console.log(
      `[http] ${req.method} ${req.originalUrl} -> ${res.statusCode} ${formatted}ms`
    );
  });
  next();
});

// ✅ Serve static assets like favicon, CSS, images
app.use(express.static(path.join(process.cwd(), 'src/web/public')));

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
