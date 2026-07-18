const { getAuthorizedClient, hasSheetsScope } = require('./googleAuth');
const { pctChange } = require('./utils');

const SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

// ---------------------------------------------------------------------------
// Label / text normalization
// ---------------------------------------------------------------------------

// Lowercases, trims, and normalizes curly apostrophes so labels like
// "KPI's" / "KPI’s" / "kpi's" / "KPIs" all compare equal after stripping
// punctuation differences the user might type into a sheet.
function normalizeLabel(value) {
  return String(value === undefined || value === null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/[‘’]/g, "'");
}

function labelContainsAny(value, variants) {
  const norm = normalizeLabel(value);
  if (!norm) return false;
  return variants.some((v) => norm.includes(v));
}

// ---------------------------------------------------------------------------
// Flexible date parsing — handles slash dates (M/D/Y and D/M/Y), ISO dates,
// spelled-out month names/abbreviations, and Google Sheets serial numbers.
// Returns { year, month, day } or null if the cell isn't a recognizable date.
// Never throws — a cell that isn't a date is simply not a date column.
// ---------------------------------------------------------------------------

const MONTH_NAMES = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december',
];

function monthNameToNumber(raw) {
  const norm = normalizeLabel(raw).replace(/\.$/, '');
  if (!norm) return null;
  const idx = MONTH_NAMES.findIndex((name) => name === norm || name.startsWith(norm));
  return idx === -1 ? null : idx + 1;
}

function makeDate(year, month, day) {
  if (!year || !month || !day) return null;
  if (month < 1 || month > 12) return null;
  if (day < 1 || day > 31) return null;
  return { year, month, day };
}

// Google Sheets/Excel serial dates count days since 1899-12-30. Only treat a
// bare number as a serial date if it falls in a plausible calendar range
// (roughly 1955-2064) so ordinary numeric cells (rankings, counts) aren't
// misread as dates.
function parseSerialDate(raw) {
  if (!/^\d{4,6}(\.\d+)?$/.test(raw)) return null;
  const serial = Number(raw);
  if (!(serial > 20000 && serial < 80000)) return null;
  const epochMs = Date.UTC(1899, 11, 30);
  const d = new Date(epochMs + serial * 86400000);
  return makeDate(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}

function parseFlexibleDate(raw) {
  if (raw === undefined || raw === null) return null;
  const str = String(raw).trim();
  if (str === '') return null;

  // ISO: YYYY-MM-DD
  let m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return makeDate(Number(m[1]), Number(m[2]), Number(m[3]));

  // Slash-separated: M/D/YYYY, assumed month-first (matches how these sheets
  // are authored); if the first part isn't a valid month but the second is,
  // treat it as day-first (D/M/YYYY) instead.
  m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    let month = a;
    let day = b;
    if (month > 12 && day <= 12) {
      month = b;
      day = a;
    }
    return makeDate(year, month, day);
  }

  // "June 1, 2026" / "Jun 1 2026"
  m = str.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m) {
    const month = monthNameToNumber(m[1]);
    if (month) return makeDate(Number(m[3]), month, Number(m[2]));
  }

  // "1 June 2026" / "1 Jun, 2026"
  m = str.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(\d{4})$/);
  if (m) {
    const month = monthNameToNumber(m[2]);
    if (month) return makeDate(Number(m[3]), month, Number(m[1]));
  }

  return parseSerialDate(str);
}

// For headers that name a month without a day (e.g. the KPI tab's "JAN",
// "01", "January 2026") rather than a full date.
function parseMonthLabel(raw) {
  if (raw === undefined || raw === null) return null;
  const str = String(raw).trim();
  if (str === '') return null;

  const withYear = str.match(/^([A-Za-z]{3,9})\.?\s*'?(\d{2,4})?$/);
  if (withYear) {
    const month = monthNameToNumber(withYear[1]);
    if (month) {
      let year = withYear[2] ? Number(withYear[2]) : null;
      if (year !== null && year < 100) year += 2000;
      return { month, year };
    }
  }

  const numeric = str.match(/^(\d{1,2})$/);
  if (numeric) {
    const month = Number(numeric[1]);
    if (month >= 1 && month <= 12) return { month, year: null };
  }

  // Fall back to full-date parsing in case the header is an actual date
  // (e.g. "6/1/2026") rather than a bare month label.
  const asDate = parseFlexibleDate(str);
  if (asDate) return { month: asDate.month, year: asDate.year };

  return null;
}

// ---------------------------------------------------------------------------
// Sheets API access
// ---------------------------------------------------------------------------

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

// Fetches just the tab titles of the spreadsheet, used to resolve which tab
// (by name) covers which section, since different clients name them differently.
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

// ---------------------------------------------------------------------------
// Tab resolution — tabs are found by what their name *contains*, not an
// exact/candidate-list match, so any reasonably-named client sheet works
// without code changes.
// ---------------------------------------------------------------------------

function isKeywordTabName(name) {
  return normalizeLabel(name).includes('keyword');
}

function isOffPageTabName(name) {
  return labelContainsAny(name, ['kpi_2', 'kpi 2', 'off-page', 'offpage', 'submission', 'activity tracker']);
}

// The monthly-traffic/DA-PA "KPI" tab is anything containing "kpi" that
// ISN'T the off-page "KPI_2"/"KPI 2" tab.
function isKpiTrafficTabName(name) {
  const norm = normalizeLabel(name);
  return norm.includes('kpi') && !isOffPageTabName(name);
}

// When more than one tab matches a category (e.g. "Keywords" and "Keywords
// Old"), prefer whichever doesn't look like a stale/backup copy.
function pickBestTab(matchingNames) {
  if (matchingNames.length === 0) return null;
  const fresh = matchingNames.filter((name) => !/\b(old|archive|archived|backup|deprecated)\b/i.test(name));
  return (fresh.length ? fresh : matchingNames)[0];
}

function resolveTabs(availableTitles) {
  const keywordTab = pickBestTab(availableTitles.filter(isKeywordTabName));
  const offPageTab = pickBestTab(availableTitles.filter(isOffPageTabName));
  const kpiTrafficTab = pickBestTab(availableTitles.filter(isKpiTrafficTabName));
  return { keywordTab, offPageTab, kpiTrafficTab };
}

function tabNotFoundMessage(sectionLabel, availableTitles) {
  return `${sectionLabel} tab not found. Available tabs: [${availableTitles.join(', ')}].`;
}

// ---------------------------------------------------------------------------
// Shared date-column helpers
// ---------------------------------------------------------------------------

// Scans a header row (starting at `startCol`) for cells that parse as full
// dates, returning { col, date } for each one found.
function findDateColumns(headerRow, startCol) {
  const cols = [];
  for (let col = startCol; col < headerRow.length; col++) {
    const parsed = parseFlexibleDate(headerRow[col]);
    if (parsed) cols.push({ col, date: parsed });
  }
  return cols;
}

function colsForMonth(dateCols, monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  return dateCols.filter((dc) => dc.date.year === year && dc.date.month === month);
}

// Multiple snapshot dates can fall in the same month; the most recent one
// is the most representative ranking for that month.
function latestColForMonth(dateCols, monthStr) {
  const matches = colsForMonth(dateCols, monthStr);
  if (matches.length === 0) return null;
  matches.sort((a, b) => a.date.day - b.date.day);
  return matches[matches.length - 1].col;
}

// Finds the row index (within `rows`) whose first column matches one of the
// given label variants. Returns -1 if not found.
function findLabelRow(rows, variants) {
  for (let i = 0; i < rows.length; i++) {
    if (labelContainsAny(rows[i][0], variants)) return i;
  }
  return -1;
}

function collectColumnALabels(rows) {
  return rows
    .map((r) => (r[0] === undefined || r[0] === null ? '' : String(r[0]).trim()))
    .filter((label) => label !== '');
}

// ---------------------------------------------------------------------------
// Keyword ranking buckets
// ---------------------------------------------------------------------------

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

// Finds the header row anywhere in the tab: the first row whose column A
// contains "keyword" (not assumed to be a fixed row number like 10 or 11).
function findKeywordHeaderRowIndex(values) {
  for (let i = 0; i < values.length; i++) {
    if (labelContainsAny(values[i][0], ['keyword'])) return i;
  }
  return -1;
}

async function getKeywordRankings(currentMonth, comparisonMonth, sheetId, email, tabName) {
  ensureSheetsScope(email);
  const range = `${quoteSheetName(tabName)}!A1:ZZ2000`;
  const values = await getSheetValues(range, sheetId, email);

  if (values.length === 0) {
    return {
      current: emptyBuckets(),
      previous: emptyBuckets(),
      hasData: false,
      error: `The "${tabName}" tab is empty.`,
    };
  }

  const headerRowIdx = findKeywordHeaderRowIndex(values);
  if (headerRowIdx === -1) {
    const foundLabels = collectColumnALabels(values).slice(0, 20);
    return {
      current: emptyBuckets(),
      previous: emptyBuckets(),
      hasData: false,
      error: `Could not find a "Keyword" header row in the "${tabName}" tab. Detected labels in column A: [${foundLabels.join(', ')}].`,
    };
  }

  const headerRow = values[headerRowIdx];
  // Data rows run from immediately after the header until the first row
  // with an empty column A.
  const dataRows = [];
  for (let i = headerRowIdx + 1; i < values.length; i++) {
    if (!values[i][0] || String(values[i][0]).trim() === '') break;
    dataRows.push(values[i]);
  }

  const dateCols = findDateColumns(headerRow, 1);
  const currCol = latestColForMonth(dateCols, currentMonth);
  const prevCol = latestColForMonth(dateCols, comparisonMonth);

  function bucketsForCol(colIdx) {
    const buckets = emptyBuckets();
    if (colIdx === null) return buckets;
    dataRows.forEach((row) => {
      buckets[classifyRank(row[colIdx])]++;
    });
    return buckets;
  }

  let error = null;
  if (currCol === null) {
    const foundDates = dateCols.map((dc) => `${dc.date.month}/${dc.date.day}/${dc.date.year}`);
    error = `No date columns found for ${currentMonth} in the "${tabName}" tab. Found date columns: [${foundDates.join(', ')}].`;
  }

  return {
    current: bucketsForCol(currCol),
    previous: bucketsForCol(prevCol),
    hasData: currCol !== null && dataRows.length > 0,
    error,
  };
}

// ---------------------------------------------------------------------------
// Off-page / submissions activity tab
// ---------------------------------------------------------------------------

async function getOffPageSubmissions(currentMonth, comparisonMonth, sheetId, email, tabName) {
  ensureSheetsScope(email);
  const range = `${quoteSheetName(tabName)}!A1:ZZ500`;
  const values = await getSheetValues(range, sheetId, email);

  if (values.length === 0) {
    return { current: 0, previous: 0, hasData: false, error: `The "${tabName}" tab is empty.` };
  }

  const [headerRow, ...activityRows] = values;
  const rows = activityRows.filter((r) => r[0] && String(r[0]).trim() !== '');
  const dateCols = findDateColumns(headerRow, 1);

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

  let error = null;
  if (!hasCurrentCols) {
    const foundDates = dateCols.map((dc) => `${dc.date.month}/${dc.date.day}/${dc.date.year}`);
    error = `No date columns found for ${currentMonth} in the "${tabName}" tab. Found date columns: [${foundDates.join(', ')}].`;
  }

  return {
    current: sumForMonth(currentMonth),
    previous: sumForMonth(comparisonMonth),
    hasData: rows.length > 0 && (hasCurrentCols || hasComparisonCols),
    error,
  };
}

// ---------------------------------------------------------------------------
// KPI tab: traffic-by-channel breakdown + DA/PA (Domain/Page Authority)
// This is additive data not currently rendered in the deck, but resilient
// scanning is implemented the same way so it's available without further
// hardcoding once a slide/consumer wants it.
// ---------------------------------------------------------------------------

const TRAFFIC_TYPE_VARIANTS = {
  organic: ['organic traffic', 'organic search', 'organic'],
  direct: ['direct traffic', 'direct'],
  referral: ['referral traffic', 'referral'],
  social: ['social traffic', 'organic social', 'social'],
  paidSearch: ['paid search', 'paid traffic'],
  unassigned: ['unassigned'],
};

// Row 1 (month labels) can have merged cells, which the Sheets values API
// returns as blank for every column but the first in the merged range.
// Forward-filling reconstructs the intended label for every column.
function forwardFillRow(row) {
  const filled = [];
  let last = null;
  for (let i = 0; i < row.length; i++) {
    const cell = row[i] === undefined || row[i] === null ? '' : String(row[i]).trim();
    if (cell !== '') last = cell;
    filled.push(last);
  }
  return filled;
}

function findMonthColumns(monthHeaderRow, monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  const filled = forwardFillRow(monthHeaderRow);
  const cols = [];
  for (let col = 1; col < filled.length; col++) {
    const parsed = parseMonthLabel(filled[col]);
    if (!parsed) continue;
    if (parsed.month !== month) continue;
    if (parsed.year !== null && parsed.year !== year) continue;
    cols.push(col);
  }
  return cols;
}

function getKpiTrafficAndAuthority(values, currentMonth, comparisonMonth) {
  const warnings = [];
  if (values.length < 2) {
    return { trafficByChannel: null, domainAuthority: null, warnings: ['The KPI tab has no rows to read.'] };
  }

  const monthHeaderRow = values[0];
  const dataRows = values.slice(2); // row 1 = months, row 2 = "Week 1..4" sub-header
  const labels = collectColumnALabels(dataRows);

  const currentCols = findMonthColumns(monthHeaderRow, currentMonth);
  const comparisonCols = findMonthColumns(monthHeaderRow, comparisonMonth);

  if (currentCols.length === 0) {
    warnings.push(`No columns found for ${currentMonth} in the KPI tab's month header row. Detected labels: [${labels.slice(0, 20).join(', ')}].`);
  }

  function sumRowAcrossCols(row, cols) {
    if (cols.length === 0) return null;
    let total = 0;
    let any = false;
    cols.forEach((col) => {
      const n = Number(row[col]);
      if (Number.isFinite(n)) {
        total += n;
        any = true;
      }
    });
    return any ? total : null;
  }

  function firstValueAcrossCols(row, cols) {
    for (const col of cols) {
      const raw = row[col];
      if (raw !== undefined && raw !== null && String(raw).trim() !== '') {
        const n = Number(raw);
        return Number.isFinite(n) ? n : String(raw).trim();
      }
    }
    return null;
  }

  const trafficByChannel = {};
  Object.entries(TRAFFIC_TYPE_VARIANTS).forEach(([key, variants]) => {
    const rowIdx = findLabelRow(dataRows, variants);
    if (rowIdx === -1) {
      trafficByChannel[key] = { current: null, previous: null };
      warnings.push(`Could not find "${variants[0]}" row in the KPI tab. Detected labels in column A: [${labels.slice(0, 30).join(', ')}].`);
      return;
    }
    trafficByChannel[key] = {
      current: sumRowAcrossCols(dataRows[rowIdx], currentCols),
      previous: sumRowAcrossCols(dataRows[rowIdx], comparisonCols),
    };
  });

  const daRowIdx = findLabelRow(dataRows, ['domain authority', 'da']);
  const paRowIdx = findLabelRow(dataRows, ['page authority', 'pa']);

  const domainAuthority = {
    da: daRowIdx === -1 ? null : {
      current: firstValueAcrossCols(dataRows[daRowIdx], currentCols),
      previous: firstValueAcrossCols(dataRows[daRowIdx], comparisonCols),
    },
    pa: paRowIdx === -1 ? null : {
      current: firstValueAcrossCols(dataRows[paRowIdx], currentCols),
      previous: firstValueAcrossCols(dataRows[paRowIdx], comparisonCols),
    },
  };
  if (daRowIdx === -1) warnings.push(`Could not find a "Domain Authority"/"DA" row in the KPI tab.`);
  if (paRowIdx === -1) warnings.push(`Could not find a "Page Authority"/"PA" row in the KPI tab.`);

  return { trafficByChannel, domainAuthority, warnings };
}

// ---------------------------------------------------------------------------
// Public entry point
// ---------------------------------------------------------------------------

async function getSeoOverview(currentMonth, comparisonMonth, sheetId, email) {
  ensureSheetsScope(email);
  const tabTitles = await getSheetTabTitles(sheetId, email);
  const { keywordTab, offPageTab, kpiTrafficTab } = resolveTabs(tabTitles);

  const warnings = [];

  const keywordRankingsPromise = keywordTab
    ? getKeywordRankings(currentMonth, comparisonMonth, sheetId, email, keywordTab)
    : Promise.resolve({
        current: emptyBuckets(),
        previous: emptyBuckets(),
        hasData: false,
        error: tabNotFoundMessage('Keyword rankings', tabTitles),
      });

  const offPagePromise = offPageTab
    ? getOffPageSubmissions(currentMonth, comparisonMonth, sheetId, email, offPageTab)
    : Promise.resolve({ current: 0, previous: 0, hasData: false, error: tabNotFoundMessage('Off-page submissions', tabTitles) });

  const kpiTrafficPromise = kpiTrafficTab
    ? getSheetValues(`${quoteSheetName(kpiTrafficTab)}!A1:ZZ500`, sheetId, email).then((values) =>
        getKpiTrafficAndAuthority(values, currentMonth, comparisonMonth)
      )
    : Promise.resolve({ trafficByChannel: null, domainAuthority: null, warnings: [tabNotFoundMessage('KPI (traffic/authority)', tabTitles)] });

  const [keywordRankings, offPage, kpiExtra] = await Promise.all([
    keywordRankingsPromise,
    offPagePromise,
    kpiTrafficPromise,
  ]);

  if (keywordRankings.error) warnings.push(keywordRankings.error);
  if (offPage.error) warnings.push(offPage.error);
  if (kpiExtra.warnings) warnings.push(...kpiExtra.warnings);

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
    // Additive data, not currently rendered by pptx.js — available for a
    // future slide without further sheet-parsing changes.
    trafficByChannel: kpiExtra.trafficByChannel,
    domainAuthority: kpiExtra.domainAuthority,
    warnings,
  };
}

module.exports = {
  getSeoOverview,
  // Exported for unit testing.
  _internal: {
    normalizeLabel,
    labelContainsAny,
    parseFlexibleDate,
    parseMonthLabel,
    monthNameToNumber,
    classifyRank,
    isKeywordTabName,
    isOffPageTabName,
    isKpiTrafficTabName,
    pickBestTab,
    resolveTabs,
    findDateColumns,
    colsForMonth,
    latestColForMonth,
    findLabelRow,
    forwardFillRow,
    findMonthColumns,
    findKeywordHeaderRowIndex,
    getKpiTrafficAndAuthority,
  },
};
