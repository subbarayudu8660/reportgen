const express = require('express');
const { getClients, addClient, deleteClient } = require('../services/clients');

const router = express.Router();

function requireAuth(req, res, next) {
  if (!req.session.email) {
    return res.status(401).json({ error: 'Not authenticated. Please sign in first.' });
  }
  next();
}

router.use(requireAuth);

router.get('/clients', (req, res) => {
  res.json(getClients());
});

router.post('/clients', (req, res) => {
  const { name, ga4PropertyId, sheetId, metaAdAccountId } = req.body || {};
  if (!name || !name.trim() || !ga4PropertyId || !ga4PropertyId.trim()) {
    return res.status(400).json({ error: 'Client name and GA4 Property ID are required.' });
  }
  const client = addClient({
    name: name.trim(),
    ga4PropertyId: ga4PropertyId.trim(),
    sheetId: (sheetId || '').trim(),
    metaAdAccountId: (metaAdAccountId || '').trim(),
  });
  res.status(201).json(client);
});

router.delete('/clients/:id', (req, res) => {
  const deleted = deleteClient(req.params.id);
  if (!deleted) {
    return res.status(404).json({ error: 'Client not found.' });
  }
  res.status(204).send();
});

module.exports = router;
