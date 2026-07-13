const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const CLIENTS_PATH = path.join(__dirname, '..', 'clients.json');

function loadClients() {
  if (!fs.existsSync(CLIENTS_PATH)) return [];
  try {
    const parsed = JSON.parse(fs.readFileSync(CLIENTS_PATH, 'utf8'));
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveClients(clients) {
  fs.writeFileSync(CLIENTS_PATH, JSON.stringify(clients, null, 2));
}

// Meta ad account IDs must be prefixed "act_" for the Graph API; normalize so
// users can enter either the bare numeric ID or the act_-prefixed form.
function normalizeMetaAdAccountId(id) {
  const trimmed = (id || '').trim();
  if (!trimmed) return '';
  return trimmed.startsWith('act_') ? trimmed : `act_${trimmed}`;
}

// Seeds clients.json from the .env GA4/Sheet defaults exactly once, the first
// time the app is used with no clients configured yet. Once any client exists
// (even after all are later deleted down to zero... no: only when the file is
// genuinely empty/missing) this never runs again.
function ensureSeeded() {
  const clients = loadClients();
  if (clients.length > 0 || fs.existsSync(CLIENTS_PATH)) return clients;

  const { GA4_PROPERTY_ID, GOOGLE_SHEET_ID, META_AD_ACCOUNT_ID } = process.env;
  if (!GA4_PROPERTY_ID) return clients;

  const seeded = [
    {
      id: crypto.randomUUID(),
      name: 'Default Client',
      ga4PropertyId: GA4_PROPERTY_ID,
      sheetId: GOOGLE_SHEET_ID || '',
      metaAdAccountId: normalizeMetaAdAccountId(META_AD_ACCOUNT_ID),
      createdAt: new Date().toISOString(),
    },
  ];
  saveClients(seeded);
  return seeded;
}

function getClients() {
  return ensureSeeded();
}

function getClientById(id) {
  return getClients().find((c) => c.id === id) || null;
}

function addClient({ name, ga4PropertyId, sheetId, metaAdAccountId }) {
  const clients = loadClients();
  const client = {
    id: crypto.randomUUID(),
    name,
    ga4PropertyId,
    sheetId: sheetId || '',
    metaAdAccountId: normalizeMetaAdAccountId(metaAdAccountId),
    createdAt: new Date().toISOString(),
  };
  clients.push(client);
  saveClients(clients);
  return client;
}

function deleteClient(id) {
  const clients = loadClients();
  const idx = clients.findIndex((c) => c.id === id);
  if (idx === -1) return false;
  clients.splice(idx, 1);
  saveClients(clients);
  return true;
}

module.exports = { getClients, getClientById, addClient, deleteClient };
