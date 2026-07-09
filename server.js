require('dotenv').config();
const express = require('express');
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

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/auth', authRoutes);
app.use('/api', reportRoutes);
app.use('/api', clientsRoutes);

app.listen(PORT, () => {
  console.log(`ReportGen running at http://localhost:${PORT}`);
});
