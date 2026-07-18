const { getAuthorizedClient, hasSheetsScope } = require('./googleAuth');
const { pctChange } = require('./utils');

const SHEETS_API_BASE = 'https://sheets.googleapis.com/v4/spreadsheets';

// ---------------------------------------------------------------------------
// Label / text normalization
// ---------------------------------------------------------------------------

// Lowercases, trims, and normalizes curly apostrophes/backticks so labels
// like "KPI's" / "KPI’s" / "kpi`s" / "KPIs" all compare equal after
// stripping punctuation differences the user might type into a sheet.
function normalizeLabel(value) {
  return String(value === undefined || value === null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/[‘’`]/g, "'");
}

// Short acronym variants (e.g. "DA", "PA") require an exact match rather
// than substring-contains — otherwise "DA" would false-positive against
// "Data Studio Link" and "PA" against "Paid Search".
function matchesVariant(label, variant) {
  const normLabel = normalizeLabel(label);
  const normVariant = normalizeLabel(variant);
  if (!normLabel || !normVariant) return false;
  if (normVariant.length <= 3) return normLabel === normVariant;
  return normLabel.includes(normVariant);
}

function labelContainsAny(value, variants) {
  return variants.some((v) => matchesVariant(value, v));
}

// A cell containing only punctuation/whitespace (a stray apostrophe, a
// lone dash, backticks left over from a copy-paste) reads as "blank" for
// the purposes of finding the end of a data range or a real label.
function isBlankCell(value) {
  const norm = normalizeLabel(value);
  if (norm === '') return true;
  return norm.replace(/[^a-z0-9]/g, '') === '';
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

// "2026-06" -> "June 2026", for human-readable warnings.
function monthYearLabel(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  const name = MONTH_NAMES[month - 1];
  return `${name.charAt(0).toUpperCase()}${name.slice(1)} ${year}`;
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
// Tab resolution.
//
// Fast path: an exact (case/whitespace/apostrophe-insensitive) match against
// a short list of tab names we've actually seen in client sheets. This is
// just a shortcut for the common case — it is NOT the source of truth.
//
// Fallback: if no variant matches exactly, dynamically scan every tab title
// for the relevant substring, so a sheet named something we've never seen
// still resolves correctly without a code change.
// ---------------------------------------------------------------------------

const TAB_VARIANTS = {
  keywords: ['Keyword Ranking Report', 'Keywords Ranking Report', 'Keyword Rankings'],
  kpi: ["KPI's", 'KPIs', 'KPI', 'Monthly KPIs'],
  offpage: ['KPI_2', 'KPI 2', 'Off-page SEO Submission Tracker'],
};

function exactMatchTab(variants, availableTitles) {
  const byNormalized = new Map(availableTitles.map((title) => [normalizeLabel(title), title]));
  for (const variant of variants) {
    const match = byNormalized.get(normalizeLabel(variant));
    if (match) return match;
  }
  return null;
}

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

function resolveTab(exactVariants, dynamicPredicate, availableTitles) {
  return (
    exactMatchTab(exactVariants, availableTitles) ||
    pickBestTab(availableTitles.filter(dynamicPredicate))
  );
}

function resolveTabs(availableTitles) {
  // Off-page is resolved before the generic KPI tab so that a sheet with
  // both "KPI_2" and "KPI" tabs never lets the KPI-traffic fallback (which
  // only excludes tabs it itself classifies as off-page) accidentally
  // swallow the off-page tab first.
  const offPageTab = resolveTab(TAB_VARIANTS.offpage, isOffPageTabName, availableTitles);
  const keywordTab = resolveTab(TAB_VARIANTS.keywords, isKeywordTabName, availableTitles);
  const kpiTrafficTab = resolveTab(TAB_VARIANTS.kpi, isKpiTrafficTabName, availableTitles);
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

// Finds the first row index (within `rows`) whose `columnIndex` cell matches
// one of the given label variants — the universal label matcher used
// everywhere a "which row is X in" question needs answering. Never assumes
// a fixed row number; if a label appears more than once, the first match
// wins. Returns -1 if not found.
function findRowByLabel(rows, columnIndex, variants) {
  return rows.findIndex((row) => labelContainsAny(row[columnIndex], variants));
}

function collectColumnALabels(rows) {
  return rows
    .map((r) => (r[0] === undefined || r[0] === null ? '' : String(r[0]).trim()))
    .filter((label) => label !== '' && !isBlankCell(label));
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

// Finds the header row anywhere in the tab — any row whose column A
// contains "keyword" (not assumed to be a fixed row number like 10 or 11).
// A decorative title row ("Keyword Ranking Report") can also contain the
// word "keyword" above the real header, so when several rows match, the
// real header is picked as whichever candidate has the most parseable date
// columns beside it; ties go to the lowest (later) row, since a title row
// always sits above its header, never below it.
function findKeywordHeaderRowIndex(values) {
  const candidates = [];
  for (let i = 0; i < values.length; i++) {
    if (labelContainsAny(values[i][0], ['keyword'])) candidates.push(i);
  }
  if (candidates.length === 0) return -1;
  if (candidates.length === 1) return candidates[0];

  let best = candidates[0];
  let bestDateCount = -1;
  candidates.forEach((idx) => {
    const dateCount = findDateColumns(values[idx], 1).length;
    if (dateCount >= bestDateCount) {
      bestDateCount = dateCount;
      best = idx;
    }
  });
  return best;
}

async function getKeywordRankings(currentMonth, comparisonMonth, sheetId, email, tabName) {
  ensureSheetsScope(email);
  const range = `${quoteSheetName(tabName)}!A1:ZZ2000`;
  const values = await getSheetValues(range, sheetId, email);

  if (values.length === 0) {
    return {
      current: null,
      previous: null,
      hasData: false,
      error: `The "${tabName}" tab is empty.`,
      comparisonError: null,
    };
  }

  const headerRowIdx = findKeywordHeaderRowIndex(values);
  if (headerRowIdx === -1) {
    const foundLabels = collectColumnALabels(values).slice(0, 20);
    return {
      current: null,
      previous: null,
      hasData: false,
      error: `Could not find a "Keyword" header row in the "${tabName}" tab. Detected labels in column A: [${foundLabels.join(', ')}].`,
      comparisonError: null,
    };
  }

  const headerRow = values[headerRowIdx];
  // Data rows run from immediately after the header until the first row
  // with an empty column A.
  const dataRows = [];
  for (let i = headerRowIdx + 1; i < values.length; i++) {
    if (isBlankCell(values[i][0])) break;
    dataRows.push(values[i]);
  }

  const dateCols = findDateColumns(headerRow, 1);
  const currCol = latestColForMonth(dateCols, currentMonth);
  const prevCol = latestColForMonth(dateCols, comparisonMonth);

  const foundDatesList = dateCols.map((dc) => `${dc.date.month}/${dc.date.day}/${dc.date.year}`);
  function logColumnUsed(monthStr, colIdx) {
    if (colIdx === null) return;
    const dc = dateCols.find((d) => d.col === colIdx);
    console.log(
      `[sheets] Keyword rankings ("${tabName}"): using ${dc.date.month}/${dc.date.day}/${dc.date.year} (column ${dc.col}) as the representative snapshot for ${monthStr}.`
    );
  }
  logColumnUsed(currentMonth, currCol);
  logColumnUsed(comparisonMonth, prevCol);

  // Returns null — not zeroed buckets — when this specific month has no
  // matching date column, so "no data available" is never rendered the
  // same way as a real zero-count result.
  function bucketsForCol(colIdx) {
    if (colIdx === null || dataRows.length === 0) return null;
    const buckets = emptyBuckets();
    dataRows.forEach((row) => {
      buckets[classifyRank(row[colIdx])]++;
    });
    return buckets;
  }

  function missingDateError(monthStr) {
    return dateCols.length > 0
      ? `No keyword ranking data found for ${monthYearLabel(monthStr)} — nearest available dates in sheet: [${foundDatesList.join(', ')}].`
      : `No date columns found in the "${tabName}" tab.`;
  }

  let error = null;
  if (currCol === null) {
    error = missingDateError(currentMonth);
  } else if (dataRows.length === 0) {
    error = `The "${tabName}" tab has a header row but no keyword data rows beneath it.`;
  }

  // Comparison-period-only gap: current period has data but the comparison
  // period doesn't (e.g. the sheet's tracking doesn't go back that far).
  // Reported independently of `error` so a real current-period result isn't
  // hidden behind a message that's really about the comparison period.
  const comparisonError = currCol !== null && dataRows.length > 0 && prevCol === null
    ? missingDateError(comparisonMonth)
    : null;

  return {
    current: bucketsForCol(currCol),
    previous: bucketsForCol(prevCol),
    hasData: currCol !== null && dataRows.length > 0,
    error,
    comparisonError,
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
  const rows = activityRows.filter((r) => !isBlankCell(r[0]));
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
// KPI tab: DA/PA (Domain/Page Authority) only.
//
// Organic/Direct/Referral/Social/Paid-Search traffic is intentionally NOT
// pulled from this tab — GA4 (services/ga4.js) is already the source of
// truth for channel traffic on the Traffic Overview slide, and duplicating
// it from the sheet would just create a second, possibly-conflicting number
// for the same metric.
// ---------------------------------------------------------------------------

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

function getKpiAuthority(values, currentMonth, comparisonMonth) {
  const warnings = [];
  if (values.length < 2) {
    return { domainAuthority: null, warnings: ['The KPI tab has no rows to read.'] };
  }

  const monthHeaderRow = values[0];
  const dataRows = values.slice(2); // row 1 = months, row 2 = "Week 1..4" sub-header

  const currentCols = findMonthColumns(monthHeaderRow, currentMonth);
  const comparisonCols = findMonthColumns(monthHeaderRow, comparisonMonth);

  if (currentCols.length === 0) {
    const labels = collectColumnALabels(dataRows);
    warnings.push(`No columns found for ${currentMonth} in the KPI tab's month header row. Detected labels: [${labels.slice(0, 20).join(', ')}].`);
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

  const daRowIdx = findRowByLabel(dataRows, 0, ['domain authority', 'da']);
  const paRowIdx = findRowByLabel(dataRows, 0, ['page authority', 'pa']);

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

  return { domainAuthority, warnings };
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
        current: null,
        previous: null,
        hasData: false,
        error: tabNotFoundMessage('Keyword rankings', tabTitles),
        comparisonError: null,
      });

  const offPagePromise = offPageTab
    ? getOffPageSubmissions(currentMonth, comparisonMonth, sheetId, email, offPageTab)
    : Promise.resolve({ current: 0, previous: 0, hasData: false, error: tabNotFoundMessage('Off-page submissions', tabTitles) });

  const kpiAuthorityPromise = kpiTrafficTab
    ? getSheetValues(`${quoteSheetName(kpiTrafficTab)}!A1:ZZ500`, sheetId, email).then((values) =>
        getKpiAuthority(values, currentMonth, comparisonMonth)
      )
    : Promise.resolve({ domainAuthority: null, warnings: [tabNotFoundMessage('KPI (DA/PA)', tabTitles)] });

  const [keywordRankings, offPage, kpiExtra] = await Promise.all([
    keywordRankingsPromise,
    offPagePromise,
    kpiAuthorityPromise,
  ]);

  if (keywordRankings.error) warnings.push(keywordRankings.error);
  if (keywordRankings.comparisonError) warnings.push(keywordRankings.comparisonError);
  if (offPage.error) warnings.push(offPage.error);
  if (kpiExtra.warnings) warnings.push(...kpiExtra.warnings);

  // Bucket-by-bucket % change only makes sense when BOTH periods actually
  // have data; `undefined` (not `pctChange`'s own `null` "New" sentinel)
  // signals "not comparable at all" so `fmtPct()` renders "N/A" rather than
  // "New" for a period that's genuinely missing, not just starting from zero.
  const bucketKeys = ['top10', 'top11_30', 'top31_50', 'top51_100', 'pending'];
  const changes = {};
  bucketKeys.forEach((key) => {
    changes[key] =
      keywordRankings.current === null || keywordRankings.previous === null
        ? undefined
        : pctChange(keywordRankings.current[key], keywordRankings.previous[key]);
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
    domainAuthority: kpiExtra.domainAuthority,
    warnings,
  };
}

module.exports = {
  getSeoOverview,
  // Exported for unit testing only — not meant to be imported by application code.
  _internal: {
    normalizeLabel,
    labelContainsAny,
    matchesVariant,
    isBlankCell,
    parseFlexibleDate,
    parseMonthLabel,
    monthNameToNumber,
    monthYearLabel,
    classifyRank,
    TAB_VARIANTS,
    exactMatchTab,
    isKeywordTabName,
    isOffPageTabName,
    isKpiTrafficTabName,
    pickBestTab,
    resolveTab,
    resolveTabs,
    findDateColumns,
    colsForMonth,
    latestColForMonth,
    findRowByLabel,
    collectColumnALabels,
    forwardFillRow,
    findMonthColumns,
    findKeywordHeaderRowIndex,
    getKpiAuthority,
  },
};
