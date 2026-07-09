const fs = require('fs');
const path = require('path');
const { OAuth2Client } = require('google-auth-library');

const TOKENS_PATH = path.join(__dirname, '..', 'tokens.json');
const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets.readonly';
const SCOPES = ['https://www.googleapis.com/auth/analytics.readonly', SHEETS_SCOPE];

function createOAuthClient() {
  return new OAuth2Client(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
}

function loadTokens() {
  if (!fs.existsSync(TOKENS_PATH)) return null;
  try {
    return JSON.parse(fs.readFileSync(TOKENS_PATH, 'utf8'));
  } catch {
    return null;
  }
}

function saveTokens(tokens) {
  fs.writeFileSync(TOKENS_PATH, JSON.stringify(tokens, null, 2), { mode: 0o600 });
}

function isAuthenticated() {
  const tokens = loadTokens();
  return !!(tokens && tokens.refresh_token);
}

// The stored token only reflects the Sheets scope if it was granted during the
// consent screen that produced it — tokens saved before this scope was added
// won't have it, so callers must detect this and prompt re-auth rather than
// assume it's present.
function hasSheetsScope() {
  const tokens = loadTokens();
  if (!tokens || !tokens.refresh_token) return false;
  const granted = (tokens.scope || '').split(' ').filter(Boolean);
  return granted.includes(SHEETS_SCOPE);
}

function getAuthUrl() {
  const client = createOAuthClient();
  return client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: SCOPES,
  });
}

async function handleCallback(code) {
  const client = createOAuthClient();
  const { tokens } = await client.getToken(code);
  const existing = loadTokens();
  // Google only returns refresh_token on the first consent; preserve it on re-auth.
  if (!tokens.refresh_token && existing && existing.refresh_token) {
    tokens.refresh_token = existing.refresh_token;
  }
  saveTokens(tokens);
  return tokens;
}

async function getAuthorizedClient() {
  const tokens = loadTokens();
  if (!tokens || !tokens.refresh_token) {
    const err = new Error('Not authenticated. Please sign in with Google first.');
    err.code = 'NOT_AUTHENTICATED';
    throw err;
  }
  const client = createOAuthClient();
  client.setCredentials(tokens);
  client.on('tokens', (newTokens) => {
    const merged = { ...tokens, ...newTokens };
    if (!merged.refresh_token && tokens.refresh_token) {
      merged.refresh_token = tokens.refresh_token;
    }
    saveTokens(merged);
  });

  try {
    await client.getAccessToken();
  } catch (e) {
    const err = new Error('Google authentication expired or was revoked. Please sign in again.');
    err.code = 'AUTH_EXPIRED';
    throw err;
  }

  return client;
}

module.exports = {
  isAuthenticated,
  hasSheetsScope,
  getAuthUrl,
  handleCallback,
  getAuthorizedClient,
};
