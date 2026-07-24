const { getAuthorizedClient, hasGoogleAdsScope } = require('./googleAuth');
const { monthDateRange } = require('./utils');

const GOOGLE_ADS_API_VERSION = 'v24';
const GOOGLE_ADS_API_BASE = `https://googleads.googleapis.com/${GOOGLE_ADS_API_VERSION}`;

function fmtCurrency(n) {
  return `$${n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

// ListAccessibleCustomers reflects which accounts the signed-in user has
// DIRECT access to (i.e. no manager-account relationship needed). It rarely
// changes, so it's cached per signed-in email rather than re-fetched on every
// report — this is an in-memory, per-process cache (not a database), which is
// fine since it's just an optimization: a stale/missing entry just falls back
// to re-fetching.
const ACCESSIBLE_CUSTOMERS_TTL_MS = 60 * 60 * 1000; // 1 hour
const accessibleCustomersCache = new Map(); // email -> { ids: Set<string>, fetchedAt: number }

async function fetchAccessibleCustomerIds(accessToken, developerToken) {
  const res = await fetch(`${GOOGLE_ADS_API_BASE}/customers:listAccessibleCustomers`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'developer-token': developerToken,
    },
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) {
    const err = new Error((body.error && body.error.message) || `listAccessibleCustomers failed (${res.status})`);
    err.code = 'GOOGLE_ADS_API_ERROR';
    err.status = res.status;
    throw err;
  }
  const resourceNames = Array.isArray(body.resourceNames) ? body.resourceNames : [];
  return new Set(resourceNames.map((name) => name.replace(/^customers\//, '')));
}

async function getAccessibleCustomerIds(email, accessToken, developerToken) {
  const cached = accessibleCustomersCache.get(email);
  if (cached && Date.now() - cached.fetchedAt < ACCESSIBLE_CUSTOMERS_TTL_MS) {
    return cached.ids;
  }
  const ids = await fetchAccessibleCustomerIds(accessToken, developerToken);
  accessibleCustomersCache.set(email, { ids, fetchedAt: Date.now() });
  return ids;
}

// Same pattern as services/sheets.js#ensureSheetsScope — tokens saved before
// the adwords scope was added won't have it, so this must be checked
// explicitly rather than assumed present just because the user is signed in.
function ensureGoogleAdsScope(email) {
  if (!hasGoogleAdsScope(email)) {
    const err = new Error('Your Google session needs additional permissions for Google Ads. Please sign out and sign in again.');
    err.code = 'GOOGLE_ADS_SCOPE_MISSING';
    throw err;
  }
}

// Fetches Google Ads campaign performance for one reporting month via the
// googleAds:search REST endpoint. Throws with `.code` of GOOGLE_ADS_NOT_CONFIGURED,
// GOOGLE_ADS_NO_ACCOUNT, GOOGLE_ADS_TOKEN_REQUIRED, GOOGLE_ADS_PERMISSION_DENIED,
// or GOOGLE_ADS_API_ERROR — callers should catch and skip the slide rather than
// fail the report, same pattern as services/meta.js#getMetaAdsData.
async function getGoogleAdsData(googleAdsCustomerId, monthStr, email) {
  const { GOOGLE_ADS_DEVELOPER_TOKEN, GOOGLE_ADS_LOGIN_CUSTOMER_ID } = process.env;
  if (!GOOGLE_ADS_DEVELOPER_TOKEN || !GOOGLE_ADS_LOGIN_CUSTOMER_ID) {
    const err = new Error('GOOGLE_ADS_DEVELOPER_TOKEN or GOOGLE_ADS_LOGIN_CUSTOMER_ID is not configured in the environment.');
    err.code = 'GOOGLE_ADS_NOT_CONFIGURED';
    throw err;
  }
  if (!googleAdsCustomerId) {
    const err = new Error('No Google Ads Customer ID configured for this client.');
    err.code = 'GOOGLE_ADS_NO_ACCOUNT';
    throw err;
  }
  ensureGoogleAdsScope(email);

  const client = await getAuthorizedClient(email);
  const accessToken = await client.getAccessToken();

  const { startDate, endDate } = monthDateRange(monthStr);
  const query = `
    SELECT campaign.name, campaign.status, metrics.impressions, metrics.clicks, metrics.cost_micros, metrics.conversions
    FROM campaign
    WHERE segments.date BETWEEN '${startDate}' AND '${endDate}'
  `.trim();

  // Google's docs: login-customer-id must be omitted (or equal to the target
  // customer ID) when the signed-in user has DIRECT access to that account,
  // and is only required when accessing it THROUGH a manager (MCC) account.
  // Sending it unconditionally breaks direct-access accounts that aren't
  // linked to the configured MCC, so pick the header based on
  // ListAccessibleCustomers rather than always sending it.
  let hasDirectAccess = false;
  try {
    const accessibleIds = await getAccessibleCustomerIds(email, accessToken.token, GOOGLE_ADS_DEVELOPER_TOKEN);
    hasDirectAccess = accessibleIds.has(googleAdsCustomerId);
  } catch (listErr) {
    console.error('Google Ads listAccessibleCustomers error:', listErr.code, listErr.message);
    // Can't determine direct access — fall through and try the MCC path first.
  }

  async function runSearch(loginCustomerId) {
    const headers = {
      Authorization: `Bearer ${accessToken.token}`,
      'developer-token': GOOGLE_ADS_DEVELOPER_TOKEN,
      'Content-Type': 'application/json',
    };
    if (loginCustomerId) headers['login-customer-id'] = loginCustomerId;

    let res;
    try {
      res = await fetch(`${GOOGLE_ADS_API_BASE}/customers/${googleAdsCustomerId}/googleAds:search`, {
        method: 'POST',
        headers,
        body: JSON.stringify({ query }),
      });
    } catch (networkErr) {
      const err = new Error(`Could not reach the Google Ads API: ${networkErr.message}`);
      err.code = 'GOOGLE_ADS_API_ERROR';
      throw err;
    }

    const body = await res.json().catch(() => ({}));
    if (res.ok) return body;

    const apiError = body.error || {};
    const details = Array.isArray(apiError.details) ? apiError.details : [];
    const errorCode = details.reduce((acc, d) => acc || (d.errors && d.errors[0] && d.errors[0].errorCode), null);
    const errorCodeKeys = errorCode ? Object.keys(errorCode) : [];

    console.error('Google Ads API error response:', JSON.stringify({
      httpStatus: res.status,
      errorCode: apiError.code,
      message: apiError.message,
      status: apiError.status,
      details: apiError.details,
      loginCustomerIdSent: loginCustomerId || null,
    }, null, 2));

    if (errorCodeKeys.includes('developerTokenError') || errorCodeKeys.includes('authenticationError')) {
      const err = new Error(
        'Google Ads developer token is pending Basic Access approval. Google Ads data will appear automatically once approved.'
      );
      err.code = 'GOOGLE_ADS_TOKEN_REQUIRED';
      throw err;
    }
    if (errorCodeKeys.includes('authorizationError') || res.status === 403) {
      const err = new Error(
        'Your Google account does not have access to this Google Ads account. Ask the account owner to add your email as an Admin.'
      );
      err.code = 'GOOGLE_ADS_PERMISSION_DENIED';
      throw err;
    }

    const err = new Error((apiError.message) || `Google Ads API request failed (${res.status})`);
    err.code = 'GOOGLE_ADS_API_ERROR';
    err.status = res.status;
    throw err;
  }

  const primaryLoginCustomerId = hasDirectAccess ? null : GOOGLE_ADS_LOGIN_CUSTOMER_ID;
  const fallbackLoginCustomerId = hasDirectAccess ? GOOGLE_ADS_LOGIN_CUSTOMER_ID : null;

  let body;
  try {
    body = await runSearch(primaryLoginCustomerId);
  } catch (primaryErr) {
    if (primaryErr.code !== 'GOOGLE_ADS_PERMISSION_DENIED') throw primaryErr;
    // Wrong access path assumed — retry the other way before giving up.
    body = await runSearch(fallbackLoginCustomerId);
  }

  const rows = body.results || [];

  let totalImpressions = 0;
  let totalClicks = 0;
  let totalCostMicros = 0;
  let totalConversions = 0;
  let campaignsWithImpressions = 0;

  rows.forEach((row) => {
    const impressions = Number(row.metrics && row.metrics.impressions) || 0;
    totalImpressions += impressions;
    totalClicks += Number(row.metrics && row.metrics.clicks) || 0;
    totalCostMicros += Number(row.metrics && row.metrics.costMicros) || 0;
    totalConversions += Number(row.metrics && row.metrics.conversions) || 0;
    if (impressions > 0) campaignsWithImpressions += 1;
  });

  const totalSpendAmount = totalCostMicros / 1_000_000;

  return {
    totalCampaigns: campaignsWithImpressions,
    totalSpend: fmtCurrency(totalSpendAmount),
    impressions: totalImpressions,
    clicks: totalClicks,
    conversions: totalConversions,
    cpc: totalClicks > 0 ? fmtCurrency(totalSpendAmount / totalClicks) : fmtCurrency(0),
  };
}

module.exports = { getGoogleAdsData };
