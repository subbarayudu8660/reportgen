const { createClient } = require('@supabase/supabase-js');

const SUPABASE_TIMEOUT_MS = 15000;

let cachedClient = null;

// Lazily creates a single Supabase client using the service-role key from
// .env. Throws SUPABASE_NOT_CONFIGURED if either env var is missing, rather
// than letting the SDK fail with an opaque error deeper in the call chain.
function getSupabaseClient() {
  if (cachedClient) return cachedClient;

  const { SUPABASE_URL, SUPABASE_SERVICE_KEY } = process.env;
  if (!SUPABASE_URL || !SUPABASE_SERVICE_KEY) {
    const err = new Error('SUPABASE_URL and/or SUPABASE_SERVICE_KEY are not configured in the environment.');
    err.code = 'SUPABASE_NOT_CONFIGURED';
    throw err;
  }

  cachedClient = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);
  return cachedClient;
}

function withTimeout(promise, message) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => {
      const err = new Error(message);
      err.code = 'SUPABASE_TIMEOUT';
      reject(err);
    }, SUPABASE_TIMEOUT_MS);
  });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

// Returns the first/last calendar day of "YYYY-MM" as Date objects covering
// the whole month in UTC, so scan_date (timestamptz) comparisons don't miss
// scans near the month boundary due to timezone drift.
function monthBounds(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  const start = new Date(Date.UTC(year, month - 1, 1, 0, 0, 0));
  const end = new Date(Date.UTC(year, month, 1, 0, 0, 0));
  return { start, end };
}

// Picks the best scan for a target month out of all scans for a website:
// the one with the latest scan_date within that month, preferring
// country = 'in' when multiple scans share the exact latest date.
function pickScanForMonth(scans, monthStr) {
  const { start, end } = monthBounds(monthStr);
  const inMonth = scans.filter((s) => {
    const t = new Date(s.scan_date).getTime();
    return t >= start.getTime() && t < end.getTime();
  });
  if (inMonth.length === 0) return null;

  const latestTime = Math.max(...inMonth.map((s) => new Date(s.scan_date).getTime()));
  const latestScans = inMonth.filter((s) => new Date(s.scan_date).getTime() === latestTime);

  if (latestScans.length === 1) return latestScans[0];
  return latestScans.find((s) => s.country === 'in') || latestScans[0];
}

function classifyRank(rank) {
  if (rank === undefined || rank === null) return 'pending';
  const str = String(rank).trim();
  if (!str || str.toLowerCase() === 'not found') return 'pending';
  const n = Number(str);
  if (!Number.isFinite(n)) return 'pending';
  if (n >= 1 && n <= 10) return 'top10';
  if (n >= 11 && n <= 30) return 'top11to30';
  if (n >= 31 && n <= 50) return 'top31to50';
  if (n >= 51 && n <= 100) return 'top51to100';
  return 'pending';
}

function emptyBuckets() {
  return { top10: 0, top11to30: 0, top31to50: 0, top51to100: 0, pending: 0 };
}

async function fetchBucketsForScan(supabase, scan) {
  const { data: rows, error } = await withTimeout(
    supabase.from('ranking_results').select('keyword, rank').eq('scan_id', scan.id),
    'Supabase query timed out while fetching ranking_results.'
  );
  if (error) {
    const err = new Error(`Supabase query failed for ranking_results: ${error.message}`);
    err.code = 'SUPABASE_API_ERROR';
    throw err;
  }

  const buckets = emptyBuckets();
  (rows || []).forEach((row) => {
    buckets[classifyRank(row.rank)] += 1;
  });
  return buckets;
}

function monthYearLabel(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  return new Date(year, month - 1, 1).toLocaleString('en-US', { month: 'long', year: 'numeric' });
}

// Fetches keyword ranking bucket counts for a website from Supabase, for both
// the current and comparison reporting periods. Selects, per period, the
// scan with the latest scan_date within that calendar month (preferring
// country = 'in' on same-date ties). Throws SUPABASE_NOT_CONFIGURED,
// SUPABASE_NO_SCANS (no scans at all for this website), SUPABASE_TIMEOUT, or
// SUPABASE_API_ERROR — callers should catch and skip the slide, same
// resilience pattern as Meta Ads / Google Ads.
async function getKeywordRankings(websiteUrl, currentMonth, comparisonMonth) {
  const supabase = getSupabaseClient();

  const { data: scans, error } = await withTimeout(
    supabase
      .from('scans')
      .select('id, scan_date, country')
      .eq('website', websiteUrl)
      .order('scan_date', { ascending: false }),
    'Supabase query timed out while fetching scans.'
  );
  if (error) {
    const err = new Error(`Supabase query failed for scans: ${error.message}`);
    err.code = 'SUPABASE_API_ERROR';
    throw err;
  }
  if (!scans || scans.length === 0) {
    const err = new Error(`No scans found in Supabase for website "${websiteUrl}".`);
    err.code = 'SUPABASE_NO_SCANS';
    throw err;
  }

  const warnings = [];
  const recentDates = scans.slice(0, 5).map((s) => s.scan_date);

  async function resolvePeriod(monthStr) {
    const scan = pickScanForMonth(scans, monthStr);
    if (!scan) {
      warnings.push(
        `No keyword scan found for ${monthYearLabel(monthStr)} for ${websiteUrl}. Available scan dates: [${recentDates.join(', ')}]`
      );
      return { buckets: null, scan: null };
    }
    console.log(`Supabase keyword scan selected for ${monthStr}: id=${scan.id} scan_date=${scan.scan_date} country=${scan.country}`);
    const buckets = await fetchBucketsForScan(supabase, scan);
    return { buckets, scan: { id: scan.id, scanDate: scan.scan_date, country: scan.country } };
  }

  const [current, comparison] = await Promise.all([
    resolvePeriod(currentMonth),
    resolvePeriod(comparisonMonth),
  ]);

  return {
    current: current.buckets,
    comparison: comparison.buckets,
    currentScan: current.scan,
    comparisonScan: comparison.scan,
    warnings,
  };
}

module.exports = { getKeywordRankings };
