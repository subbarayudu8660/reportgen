const fs = require('fs');
const path = require('path');
const { OAuth2Client } = require('google-auth-library');

const TOKENS_PATH = path.join(__dirname, '..', 'tokens.json');
const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets.readonly';
const USERINFO_SCOPE = 'https://www.googleapis.com/auth/userinfo.email';
const SCOPES = ['https://www.googleapis.com/auth/analytics.readonly', SHEETS_SCOPE, USERINFO_SCOPE, 'openid'];
const USERINFO_URL = 'https://www.googleapis.com/oauth2/v2/userinfo';

function createOAuthClient() {
  return new OAuth2Client(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI
  );
}

// tokens.json is a single file shared across all signed-in Google accounts,
// keyed by email: { "user@gmail.com": { access_token, refresh_token, expiry_date, scope } }.
function loadAllTokens() {
  if (!fs.existsSync(TOKENS_PATH)) return {};
  try {
    const parsed = JSON.parse(fs.readFileSync(TOKENS_PATH, 'utf8'));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

function saveAllTokens(allTokens) {
  fs.writeFileSync(TOKENS_PATH, JSON.stringify(allTokens, null, 2), { mode: 0o600 });
}

function loadTokens(email) {
  if (!email) return null;
  return loadAllTokens()[email] || null;
}

function saveTokens(email, tokens) {
  const all = loadAllTokens();
  all[email] = tokens;
  saveAllTokens(all);
}

function isAuthenticated(email) {
  const tokens = loadTokens(email);
  return !!(tokens && tokens.refresh_token);
}

// The stored token only reflects the Sheets scope if it was granted during the
// consent screen that produced it — tokens saved before this scope was added
// won't have it, so callers must detect this and prompt re-auth rather than
// assume it's present.
function hasSheetsScope(email) {
  const tokens = loadTokens(email);
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

async function fetchUserEmail(client) {
  const accessToken = await client.getAccessToken();
  const res = await fetch(USERINFO_URL, {
    headers: { Authorization: `Bearer ${accessToken.token}` },
  });
  if (!res.ok) {
    throw new Error(`Failed to fetch Google account info (${res.status})`);
  }
  const info = await res.json();
  if (!info.email) {
    throw new Error('Google did not return an email address for this account.');
  }
  return info.email;
}

async function handleCallback(code) {
  const client = createOAuthClient();
  const { tokens } = await client.getToken(code);
  client.setCredentials(tokens);
  const email = await fetchUserEmail(client);

  const existing = loadTokens(email);
  // Google only returns refresh_token on the first consent; preserve it on re-auth.
  if (!tokens.refresh_token && existing && existing.refresh_token) {
    tokens.refresh_token = existing.refresh_token;
  }
  saveTokens(email, tokens);
  return { tokens, email };
}

async function getAuthorizedClient(email) {
  const tokens = loadTokens(email);
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
    saveTokens(email, merged);
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
