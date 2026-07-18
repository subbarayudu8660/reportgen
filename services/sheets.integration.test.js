// Exercises getSeoOverview's off-page fallback end-to-end (tab resolution ->
// summary attempt -> detail-tab row counting) against a stubbed Sheets API,
// since the real network/auth layer can't be unit-tested directly. A fresh
// require of ./sheets is used per test so each test's fetch/auth stubs don't
// leak into the plain unit tests in sheets.test.js.
const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('path');

function freshSheetsModule({ tabs, valuesByRange }) {
  const googleAuthPath = require.resolve('./googleAuth');
  const sheetsPath = require.resolve('./sheets');
  delete require.cache[sheetsPath];
  require.cache[googleAuthPath] = {
    id: googleAuthPath,
    filename: googleAuthPath,
    loaded: true,
    exports: {
      getAuthorizedClient: async () => ({ getAccessToken: async () => ({ token: 'fake' }) }),
      hasSheetsScope: () => true,
    },
  };

  const originalFetch = global.fetch;
  global.fetch = async (url) => {
    if (url.includes('fields=sheets.properties.title')) {
      return { ok: true, json: async () => ({ sheets: tabs.map((title) => ({ properties: { title } })) }) };
    }
    for (const [needle, values] of Object.entries(valuesByRange)) {
      if (url.includes(encodeURIComponent(needle)) || url.includes(needle)) {
        return { ok: true, json: async () => ({ values }) };
      }
    }
    return { ok: true, json: async () => ({ values: [] }) };
  };

  const mod = require('./sheets');
  return { mod, restore: () => { global.fetch = originalFetch; delete require.cache[googleAuthPath]; } };
}

test('getSeoOverview falls back to the detail-log tab when KPI_2 has no data for the current month', async () => {
  const { mod, restore } = freshSheetsModule({
    tabs: ['Keyword Ranking Report', 'KPI_2', 'Off-page SEO Submission Tracker'],
    valuesByRange: {
      // Keywords tab: minimal, just needs to resolve without erroring.
      'Keyword Ranking Report': [['Keyword', '6/30/2026'], ['kw1', '5']],
      // KPI_2 summary tab: has a June date column, but a zero total (no rows).
      KPI_2: [['Metric', '6/30/2026'], ['Off Page Submissions', '0']],
      // Detail log: one row per submission, dated in column A.
      'Off-page SEO Submission Tracker': [
        ['Date', 'Site'],
        ['6/2/2026', 'site-a.com'],
        ['6/14/2026', 'site-b.com'],
        ['6/29/2026', 'site-c.com'],
        ['5/3/2026', 'site-d.com'], // comparison-period row
      ],
    },
  });

  try {
    const seo = await mod.getSeoOverview('2026-06', '2026-05', 'fake-sheet-id', 'user@example.com');
    assert.equal(seo.offPage.current, 3);
    assert.equal(seo.offPage.previous, 1);
    assert.equal(seo.offPage.hasData, true);
    assert.equal(seo.offPage.error, null);
  } finally {
    restore();
  }
});

test('getSeoOverview uses the KPI_2 summary tab directly when it already has real data (no fallback needed)', async () => {
  const { mod, restore } = freshSheetsModule({
    tabs: ['Keyword Ranking Report', 'KPI_2', 'Off-page SEO Submission Tracker'],
    valuesByRange: {
      'Keyword Ranking Report': [['Keyword', '6/30/2026'], ['kw1', '5']],
      KPI_2: [['Metric', '6/30/2026', '5/30/2026'], ['Off Page Submissions', '42', '30']],
      // Detail tab has different numbers — if this were used instead, the
      // assertions below would fail, proving the summary tab won.
      'Off-page SEO Submission Tracker': [['Date'], ['6/2/2026'], ['6/14/2026']],
    },
  });

  try {
    const seo = await mod.getSeoOverview('2026-06', '2026-05', 'fake-sheet-id', 'user@example.com');
    assert.equal(seo.offPage.current, 42);
    assert.equal(seo.offPage.previous, 30);
  } finally {
    restore();
  }
});

test('getSeoOverview returns null with a warning when neither off-page source has usable data', async () => {
  const { mod, restore } = freshSheetsModule({
    tabs: ['Keyword Ranking Report', 'KPI_2'],
    valuesByRange: {
      'Keyword Ranking Report': [['Keyword', '6/30/2026'], ['kw1', '5']],
      KPI_2: [['Metric', '5/30/2026'], ['Off Page Submissions', '10']], // no June column at all
    },
  });

  try {
    const seo = await mod.getSeoOverview('2026-06', '2026-05', 'fake-sheet-id', 'user@example.com');
    assert.equal(seo.offPage.current, null);
    assert.equal(seo.offPage.hasData, false);
    assert.ok(seo.offPage.error.includes('KPI_2'));
  } finally {
    restore();
  }
});
