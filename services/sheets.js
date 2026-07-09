const { getAuthorizedClient, hasSheetsScope } = require('./googleAuth');

const SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

const KEYWORD_SHEET_RANGE = "'Keyword Ranking Report'!A10:ZZ2000";
const KPI_SHEET_RANGE = "'KPI_2'!A1:ZZ500";

function ensureSheetsScope() {
  if (!hasSheetsScope()) {
    const err = new Error(
      'Google Sheets access has not been authorized on the current session. Please sign in with Google again to grant Sheets access.'
    );
    err.code = 'SHEETS_SCOPE_MISSING';
    throw err;
  }
}

async function getSheetValues(range, sheetId) {
  if (!sheetId) {
    throw new Error('No Google Sheet ID configured for the active client.');
  }
  const client = await getAuthorizedClient();
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

async function getKeywordRankings(currentMonth, comparisonMonth, sheetId) {
  ensureSheetsScope();
  const values = await getSheetValues(KEYWORD_SHEET_RANGE, sheetId);

  if (values.length === 0) {
    return { current: emptyBuckets(), previous: emptyBuckets(), hasData: false };
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
  };
}

async function getOffPageSubmissions(currentMonth, comparisonMonth, sheetId) {
  ensureSheetsScope();
  const values = await getSheetValues(KPI_SHEET_RANGE, sheetId);

  if (values.length === 0) {
    return { current: 0, previous: 0, hasData: false };
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
  };
}

function pctChange(current, previous) {
  if (!previous) return current > 0 ? 100 : 0;
  return ((current - previous) / previous) * 100;
}

async function getSeoOverview(currentMonth, comparisonMonth, sheetId) {
  const [keywordRankings, offPage] = await Promise.all([
    getKeywordRankings(currentMonth, comparisonMonth, sheetId),
    getOffPageSubmissions(currentMonth, comparisonMonth, sheetId),
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
    },
  };
}

module.exports = { getSeoOverview };
