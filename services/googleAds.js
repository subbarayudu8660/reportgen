const { getAuthorizedClient } = require('./googleAuth');
const { monthDateRange } = require('./utils');

const GOOGLE_ADS_API_VERSION = 'v24';
const GOOGLE_ADS_API_BASE = `https://googleads.googleapis.com/${GOOGLE_ADS_API_VERSION}`;

function fmtCurrency(n) {
  return `$${n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
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

  const client = await getAuthorizedClient(email);
  const accessToken = await client.getAccessToken();

  const { startDate, endDate } = monthDateRange(monthStr);
  const query = `
    SELECT campaign.name, campaign.status, metrics.impressions, metrics.clicks, metrics.cost_micros, metrics.conversions
    FROM campaign
    WHERE segments.date BETWEEN '${startDate}' AND '${endDate}'
  `.trim();

  let res;
  try {
    res = await fetch(`${GOOGLE_ADS_API_BASE}/customers/${googleAdsCustomerId}/googleAds:search`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken.token}`,
        'developer-token': GOOGLE_ADS_DEVELOPER_TOKEN,
        'login-customer-id': GOOGLE_ADS_LOGIN_CUSTOMER_ID,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ query }),
    });
  } catch (networkErr) {
    const err = new Error(`Could not reach the Google Ads API: ${networkErr.message}`);
    err.code = 'GOOGLE_ADS_API_ERROR';
    throw err;
  }

  const body = await res.json().catch(() => ({}));

  if (!res.ok) {
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
