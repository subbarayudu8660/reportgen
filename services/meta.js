const META_API_VERSION = 'v19.0';
const META_BASE = `https://graph.facebook.com/${META_API_VERSION}`;
const META_TIMEOUT_MS = 30000;

// Computes the first/last calendar day of a "YYYY-MM" month as YYYY-MM-DD strings
// for the Meta Insights API's time_range parameter.
function monthDateRange(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  const since = new Date(Date.UTC(year, month - 1, 1));
  const until = new Date(Date.UTC(year, month, 0));
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { since: fmt(since), until: fmt(until) };
}

async function metaFetch(url, accessToken) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), META_TIMEOUT_MS);

  let res;
  try {
    res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` }, signal: controller.signal });
  } catch (networkErr) {
    if (networkErr.name === 'AbortError') {
      const err = new Error('Meta Ads API timed out after 30 seconds. Please try again.');
      err.code = 'META_TIMEOUT';
      throw err;
    }
    const err = new Error(`Could not reach the Meta Graph API: ${networkErr.message}`);
    err.code = 'META_API_ERROR';
    throw err;
  } finally {
    clearTimeout(timeout);
  }

  const body = await res.json().catch(() => ({}));

  if (!res.ok || body.error) {
    if (body.error && body.error.code === 190) {
      const err = new Error(
        'Meta Ads access token has expired. Please generate a new token and update META_ACCESS_TOKEN in your environment variables.'
      );
      err.code = 'META_TOKEN_EXPIRED';
      throw err;
    }
    const err = new Error((body.error && body.error.message) || `Meta API request failed (${res.status})`);
    err.code = 'META_API_ERROR';
    err.status = res.status;
    throw err;
  }

  return body;
}

function sumAction(actions, actionType) {
  if (!Array.isArray(actions)) return 0;
  const match = actions.find((a) => a.action_type === actionType);
  return match ? Number(match.value) || 0 : 0;
}

function fmtSpend(spend) {
  const n = Number(spend);
  if (!Number.isFinite(n)) return null;
  return `$${n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

// Fetches Meta Ads insights + active campaign count for one reporting month.
// Throws with `.code` of META_NOT_CONFIGURED, META_NO_ACCOUNT, META_TOKEN_EXPIRED,
// or META_API_ERROR — callers should catch and skip the slide rather than fail the report.
async function getMetaAdsData(adAccountId, monthStr) {
  const accessToken = process.env.META_ACCESS_TOKEN;
  if (!accessToken) {
    const err = new Error('META_ACCESS_TOKEN is not configured in the environment.');
    err.code = 'META_NOT_CONFIGURED';
    throw err;
  }
  if (!adAccountId) {
    const err = new Error('No Meta Ad Account ID configured for this client.');
    err.code = 'META_NO_ACCOUNT';
    throw err;
  }

  const { since, until } = monthDateRange(monthStr);
  const timeRange = encodeURIComponent(JSON.stringify({ since, until }));

  const insightsUrl = `${META_BASE}/${adAccountId}/insights?time_range=${timeRange}&fields=campaign_name,impressions,clicks,spend,reach,actions&level=account`;
  const campaignsUrl = `${META_BASE}/${adAccountId}/campaigns?fields=id,name,status`;

  const [insightsBody, campaignsBody] = await Promise.all([
    metaFetch(insightsUrl, accessToken),
    metaFetch(campaignsUrl, accessToken),
  ]);

  const insightsRow = (insightsBody.data && insightsBody.data[0]) || {};
  const activeCampaigns = (campaignsBody.data || []).filter((c) => c.status === 'ACTIVE');

  return {
    totalCampaigns: activeCampaigns.length,
    totalSpend: fmtSpend(insightsRow.spend),
    impressions: insightsRow.impressions !== undefined ? Number(insightsRow.impressions) : null,
    clicks: insightsRow.clicks !== undefined ? Number(insightsRow.clicks) : null,
    reach: insightsRow.reach !== undefined ? Number(insightsRow.reach) : null,
    formLeads: sumAction(insightsRow.actions, 'lead'),
    messageConversations: sumAction(insightsRow.actions, 'onsite_conversion.messaging_conversation_started_7d'),
  };
}

module.exports = { getMetaAdsData };
