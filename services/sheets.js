const { getAuthorizedClient, hasSheetsScope } = require('./googleAuth');

const SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

// Tab names vary slightly across clients' sheets (typos, pluralization, punctuation).
// Each list is tried in order; the first one that matches an actual tab wins.
// The off-page/KPI tab is the same underlying data across clients — "KPI's",
// "KPI_2" etc. are all names seen for it, not separate tabs.
const KEYWORD_TAB_CANDIDATES = ['Keyword Ranking Report', 'Keywords Ranking Report', 'Keyword Rankings', 'Keywords'];
const KPI_TAB_CANDIDATES = ['KPI_2', 'KPI 2', 'KPI2', "KPI's", 'KPIs', 'KPI'];

function ensureSheetsScope(email) {
  if (!hasSheetsScope(email)) {
    const err = new Error(
      'Google Sheets access has not been authorized on the current session. Please sign in with Google again to grant Sheets access.'
    );
    err.code = 'SHEETS_SCOPE_MISSING';
    throw err;
  }
}

// Sheet name literals in an A1 range need internal single quotes doubled,
// e.g. a tab named KPI's becomes 'KPI''s'!A1:ZZ500.
function quoteSheetName(name) {
  return `'${name.replace(/'/g, "''")}'`;
}

async function getSheetValues(range, sheetId, email) {
  if (!sheetId) {
    throw new Error('No Google Sheet ID configured for the active client.');
  }
  const client = await getAuthorizedClient(email);
  const accessToken = await client.getAccessToken();

  const url = `${SHEETS_API_BASE}/${sheetId}/values/${encodeURIComponent(range)}?majorDimension=ROWS`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken.token}` },
  });

  if (!res.ok) {
    const text = await res.text();
    const err = new Error(`Sheets API request failed (${res.status}): ${text}`);
    err.code = 'SHEETS_API_ERROR';
    err.status = res.status;
    throw err;
  }

  const data = await res.json();
  return data.values || [];
}

// Fetches just the tab titles of the spreadsheet, used to resolve which
// naming variant a client's sheet actually uses.
async function getSheetTabTitles(sheetId, email) {
  if (!sheetId) {
    throw new Error('No Google Sheet ID configured for the active client.');
  }
  const client = await getAuthorizedClient(email);
  const accessToken = await client.getAccessToken();

  const url = `${SHEETS_API_BASE}/${sheetId}?fields=sheets.properties.title`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken.token}` },
  });

  if (!res.ok) {
    const text = await res.text();
    const err = new Error(`Sheets API request failed (${res.status}): ${text}`);
    err.code = 'SHEETS_API_ERROR';
    err.status = res.status;
    throw err;
  }

  const data = await res.json();
  return (data.sheets || []).map((s) => s.properties.title);
}

// Matches candidates against actual tab titles case-insensitively, in
// priority order, so e.g. "keywords" still matches a tab literally named "Keywords".
function resolveTabName(candidates, availableTitles) {
  const byLowerCase = new Map(availableTitles.map((title) => [title.toLowerCase(), title]));
  for (const candidate of candidates) {
    const match = byLowerCase.get(candidate.toLowerCase());
    if (match) return match;
  }
  return null;
}

function tabNotFoundMessage(availableTitles) {
  return `Could not find the expected sheet tabs. Found tabs: [${availableTitles.join(', ')}]. Please check your sheet structure.`;
}

// Sheet dates are plain "M/D/YYYY" strings (not ISO), so they need manual parsing.
function parseSheetDate(str) {
  if (!str) return null;
  const parts = String(str).trim().split('/');
  if (parts.length !== 3) return null;
  const [m, d, y] = parts.map(Number);
  if (!m || !d || !y) return null;
  return { year: y, month: m, day: d };
}

function dateColumns(headerRow) {
  const cols = [];
  for (let col = 1; col < headerRow.length; col++) {
    const parsed = parseSheetDate(headerRow[col]);
    if (parsed) cols.push({ col, date: parsed });
  }
  return cols;
}

function colsForMonth(dateCols, monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  return dateCols.filter((dc) => dc.date.year === year && dc.date.month === month);
}

// Multiple snapshot dates can fall in the same month; the most recent one
// before month-end is the most representative ranking for that month.
function latestColForMonth(dateCols, monthStr) {
  const matches = colsForMonth(dateCols, monthStr);
  if (matches.length === 0) return null;
  matches.sort((a, b) => a.date.day - b.date.day);
  return matches[matches.length - 1].col;
}

function emptyBuckets() {
  return { top10: 0, top11_30: 0, top31_50: 0, top51_100: 0, pending: 0 };
}

function classifyRank(raw) {
  if (raw === undefined || raw === null) return 'pending';
  const str = String(raw).trim();
  if (str === '' || str.includes('100+')) return 'pending';
  const n = Number(str);
  if (!Number.isFinite(n)) return 'pending';
  if (n <= 10) return 'top10';
  if (n <= 30) return 'top11_30';
  if (n <= 50) return 'top31_50';
  if (n <= 100) return 'top51_100';
  return 'pending';
}

async function getKeywordRankings(currentMonth, comparisonMonth, sheetId, email, tabName) {
  ensureSheetsScope(email);
  const range = `${quoteSheetName(tabName)}!A10:ZZ2000`;
  const values = await getSheetValues(range, sheetId, email);

  if (values.length === 0) {
    return { current: emptyBuckets(), previous: emptyBuckets(), hasData: false, error: null };
  }

  const [headerRow, ...keywordRows] = values;
  const rows = keywordRows.filter((r) => r[0]);
  const dateCols = dateColumns(headerRow);

  const currCol = latestColForMonth(dateCols, currentMonth);
  const prevCol = latestColForMonth(dateCols, comparisonMonth);

  function bucketsForCol(colIdx) {
    const buckets = emptyBuckets();
    if (colIdx === null) return buckets;
    rows.forEach((row) => {
      buckets[classifyRank(row[colIdx])]++;
    });
    return buckets;
  }

  return {
    current: bucketsForCol(currCol),
    previous: bucketsForCol(prevCol),
    hasData: currCol !== null && rows.length > 0,
    error: null,
  };
}

async function getOffPageSubmissions(currentMonth, comparisonMonth, sheetId, email, tabName) {
  ensureSheetsScope(email);
  const range = `${quoteSheetName(tabName)}!A1:ZZ500`;
  const values = await getSheetValues(range, sheetId, email);

  if (values.length === 0) {
    return { current: 0, previous: 0, hasData: false, error: null };
  }

  const [headerRow, ...activityRows] = values;
  const rows = activityRows.filter((r) => r[0]);
  const dateCols = dateColumns(headerRow);

  function sumForMonth(m) {
    const cols = colsForMonth(dateCols, m).map((dc) => dc.col);
    let total = 0;
    rows.forEach((row) => {
      cols.forEach((col) => {
        const n = Number(row[col]);
        if (Number.isFinite(n)) total += n;
      });
    });
    return total;
  }

  const hasCurrentCols = colsForMonth(dateCols, currentMonth).length > 0;
  const hasComparisonCols = colsForMonth(dateCols, comparisonMonth).length > 0;

  return {
    current: sumForMonth(currentMonth),
    previous: sumForMonth(comparisonMonth),
    hasData: rows.length > 0 && (hasCurrentCols || hasComparisonCols),
    error: null,
  };
}

function pctChange(current, previous) {
  if (!previous) return current > 0 ? 100 : 0;
  return ((current - previous) / previous) * 100;
}

async function getSeoOverview(currentMonth, comparisonMonth, sheetId, email) {
  ensureSheetsScope(email);
  const tabTitles = await getSheetTabTitles(sheetId, email);
  const keywordTabName = resolveTabName(KEYWORD_TAB_CANDIDATES, tabTitles);
  const kpiTabName = resolveTabName(KPI_TAB_CANDIDATES, tabTitles);
  const notFoundError = tabNotFoundMessage(tabTitles);

  const [keywordRankings, offPage] = await Promise.all([
    keywordTabName
      ? getKeywordRankings(currentMonth, comparisonMonth, sheetId, email, keywordTabName)
      : Promise.resolve({ current: emptyBuckets(), previous: emptyBuckets(), hasData: false, error: notFoundError }),
    kpiTabName
      ? getOffPageSubmissions(currentMonth, comparisonMonth, sheetId, email, kpiTabName)
      : Promise.resolve({ current: 0, previous: 0, hasData: false, error: notFoundError }),
  ]);

  const bucketKeys = ['top10', 'top11_30', 'top31_50', 'top51_100', 'pending'];
  const changes = {};
  bucketKeys.forEach((key) => {
    changes[key] = pctChange(keywordRankings.current[key], keywordRankings.previous[key]);
  });

  return {
    keywordRankings: { ...keywordRankings, changes },
    offPage: {
      current: offPage.current,
      previous: offPage.previous,
      change: pctChange(offPage.current, offPage.previous),
      hasData: offPage.hasData,
      error: offPage.error,
    },
  };
}

module.exports = { getSeoOverview };
