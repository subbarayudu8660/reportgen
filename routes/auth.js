const express = require('express');
const { getAuthUrl, handleCallback, isAuthenticated, hasSheetsScope } = require('../services/googleAuth');

const router = express.Router();

router.get('/status', (req, res) => {
  const authenticated = isAuthenticated();
  res.json({ authenticated, needsReauth: authenticated && !hasSheetsScope() });
});

router.get('/google', (req, res) => {
  res.redirect(getAuthUrl());
});

router.get('/callback', async (req, res) => {
  const { code, error } = req.query;

  if (error) {
    return res.redirect(`/?authError=${encodeURIComponent('Google sign-in was cancelled or denied.')}`);
  }

  if (!code) {
    return res.redirect(`/?authError=${encodeURIComponent('Missing authorization code from Google.')}`);
  }

  try {
    await handleCallback(code);
    res.redirect('/?authSuccess=1');
  } catch (e) {
    res.redirect(`/?authError=${encodeURIComponent('Failed to complete Google sign-in. Please try again.')}`);
  }
});

module.exports = router;
