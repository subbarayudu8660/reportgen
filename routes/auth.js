const crypto = require('crypto');
const express = require('express');
const { getAuthUrl, handleCallback, isAuthenticated, hasSheetsScope } = require('../services/googleAuth');

const router = express.Router();

router.get('/status', (req, res) => {
  const email = req.session.email || null;
  const authenticated = isAuthenticated(email);
  res.json({ authenticated, needsReauth: authenticated && !hasSheetsScope(email), email: authenticated ? email : null });
});

router.get('/google', (req, res) => {
  const state = crypto.randomBytes(24).toString('hex');
  req.session.oauthState = state;
  res.redirect(getAuthUrl(state));
});

router.get('/callback', async (req, res) => {
  const { code, error, state } = req.query;
  const expectedState = req.session.oauthState;
  delete req.session.oauthState;

  if (error) {
    return res.redirect(`/?authError=${encodeURIComponent('Google sign-in was cancelled or denied.')}`);
  }

  if (!state || !expectedState || state !== expectedState) {
    return res.status(400).json({ error: 'Invalid or missing OAuth state parameter. Please try signing in again.' });
  }

  if (!code) {
    return res.redirect(`/?authError=${encodeURIComponent('Missing authorization code from Google.')}`);
  }

  try {
    const { email } = await handleCallback(code);
    req.session.email = email;
    res.redirect('/?authSuccess=1');
  } catch (e) {
    res.redirect(`/?authError=${encodeURIComponent('Failed to complete Google sign-in. Please try again.')}`);
  }
});

router.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.clearCookie('reportgen.sid');
    res.redirect('/');
  });
});

module.exports = router;
