require('dotenv').config();
const fs = require('fs');
const crypto = require('crypto');
const express = require('express');
const session = require('express-session');
const path = require('path');

const authRoutes = require('./routes/auth');
const reportRoutes = require('./routes/report');
const clientsRoutes = require('./routes/clients');

// GA4_PROPERTY_ID / GOOGLE_SHEET_ID are no longer required here — they're only
// used as one-time fallback defaults to seed the first client (see services/clients.js).
// Per-report GA4/Sheet IDs now come from the active client.
const requiredEnv = ['GOOGLE_CLIENT_ID', 'GOOGLE_CLIENT_SECRET', 'GOOGLE_REDIRECT_URI'];
const missing = requiredEnv.filter((key) => !process.env[key]);
if (missing.length > 0) {
  console.error(`Missing required .env values: ${missing.join(', ')}`);
  console.error('Copy .env.example to .env and fill in your credentials before starting.');
  process.exit(1);
}

// SESSION_SECRET signs the session cookie that ties a browser session to a
// signed-in Google account's email. Generate one locally on first boot if the
// user hasn't set it yet, so per-user sessions can't be forged.
if (!process.env.SESSION_SECRET) {
  const generated = crypto.randomBytes(32).toString('hex');
  process.env.SESSION_SECRET = generated;
  try {
    fs.appendFileSync(path.join(__dirname, '.env'), `\nSESSION_SECRET=${generated}\n`);
    console.log('Generated a new SESSION_SECRET and saved it to .env');
  } catch (e) {
    console.warn('Could not persist a generated SESSION_SECRET to .env; sessions will not survive a restart.');
  }
}

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(
  session({
    secret: process.env.SESSION_SECRET,
    name: 'reportgen.sid',
    resave: false,
    saveUninitialized: false,
    cookie: { httpOnly: true, sameSite: 'lax', maxAge: 30 * 24 * 60 * 60 * 1000 },
  })
);
app.use(express.static(path.join(__dirname, 'public')));

app.use('/auth', authRoutes);
app.use('/api', reportRoutes);
app.use('/api', clientsRoutes);

app.listen(PORT, () => {
  console.log(`ReportGen running at http://localhost:${PORT}`);
});
