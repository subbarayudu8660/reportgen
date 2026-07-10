const { getAuthorizedClient } = require('./googleAuth');

const GA4_API_BASE = 'https://analyticsdata.googleapis.com/v1beta';

async function runReport(body, propertyId) {
  if (!propertyId) {
    throw new Error('No GA4 property ID configured for the active client.');
  }
  const client = await getAuthorizedClient();
  const accessToken = await client.getAccessToken();

  const res = await fetch(`${GA4_API_BASE}/properties/${propertyId}:runReport`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken.token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const text = await res.text();
    let apiErrorStatus = null;
    let apiErrorMessage = text;
    try {
      const parsed = JSON.parse(text);
      apiErrorStatus = parsed.error && parsed.error.status;
      apiErrorMessage = (parsed.error && parsed.error.message) || text;
    } catch (parseErr) {
      // Response body wasn't JSON — fall back to the raw text above.
    }
    const err = new Error(`GA4 API request failed (${res.status}): ${apiErrorMessage}`);
    err.code = 'GA4_API_ERROR';
    err.status = res.status;
    err.apiErrorStatus = apiErrorStatus;
    throw err;
  }

  return res.json();
}

// Maps a raw GA4_API_ERROR into a specific, user-actionable message + HTTP status.
// `err` must have `.status` (HTTP status from the GA4 API) and optionally `.apiErrorStatus`
// (the GA4 API's own error status string, e.g. "PERMISSION_DENIED").
function mapGa4Error(err) {
  const status = err.status;
  const apiStatus = err.apiErrorStatus;

  if (status === 403 || apiStatus === 'PERMISSION_DENIED') {
    return {
      httpStatus: 403,
      message:
        'Your Google account does not have access to this GA4 property. Ask the property owner to add your email as a Viewer in GA4 → Admin → Property Access Management.',
    };
  }
  if (status === 404 || apiStatus === 'NOT_FOUND') {
    return {
      httpStatus: 404,
      message:
        'The GA4 Property ID you entered does not exist. Please double-check the Property ID in your client settings — it should be a 9-digit number found in GA4 → Admin → Property Settings.',
    };
  }
  if (status === 400 || apiStatus === 'INVALID_ARGUMENT') {
    return {
      httpStatus: 400,
      message: 'The GA4 Property ID format is invalid. It should be a plain 9-digit number with no extra characters.',
    };
  }
  if (status === 429 || apiStatus === 'RESOURCE_EXHAUSTED') {
    return {
      httpStatus: 429,
      message: 'Google Analytics API quota exceeded. Please wait a few minutes and try again.',
    };
  }
  if (status === 401 || apiStatus === 'UNAUTHENTICATED') {
    return {
      httpStatus: 401,
      message: 'Your Google session has expired. Please sign out and sign in again.',
    };
  }
  return {
    httpStatus: 502,
    message: `Google Analytics returned an unexpected error (${status || 'unknown'}). Please try again or contact support.`,
  };
}

function monthDateRange(monthStr) {
  // monthStr like "2026-06"
  const [year, month] = monthStr.split('-').map(Number);
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const endDateObj = new Date(year, month, 0); // last day of month
  const endDate = `${endDateObj.getFullYear()}-${String(endDateObj.getMonth() + 1).padStart(2, '0')}-${String(endDateObj.getDate()).padStart(2, '0')}`;
  return { startDate, endDate };
}

function rowsToMap(report) {
  const dims = (report.dimensionHeaders || []).map((h) => h.name);
  const mets = (report.metricHeaders || []).map((h) => h.name);
  return (report.rows || []).map((row) => {
    const entry = {};
    row.dimensionValues.forEach((v, i) => (entry[dims[i]] = v.value));
    row.metricValues.forEach((v, i) => (entry[mets[i]] = Number(v.value)));
    return entry;
  });
}

function sumMetric(report, metricName) {
  const mets = (report.metricHeaders || []).map((h) => h.name);
  const idx = mets.indexOf(metricName);
  if (idx === -1 || !report.rows || report.rows.length === 0) return 0;
  return report.rows.reduce((sum, row) => sum + Number(row.metricValues[idx].value), 0);
}

async function getMonthSessionTotal(monthStr, propertyId) {
  const { startDate, endDate } = monthDateRange(monthStr);
  const report = await runReport(
    {
      dateRanges: [{ startDate, endDate }],
      metrics: [{ name: 'sessions' }],
    },
    propertyId
  );
  return sumMetric(report, 'sessions');
}

async function getOrganicSearchMetrics(monthStr, propertyId) {
  const { startDate, endDate } = monthDateRange(monthStr);
  const report = await runReport(
    {
      dateRanges: [{ startDate, endDate }],
      dimensions: [{ name: 'sessionDefaultChannelGroup' }],
      metrics: [
        { name: 'sessions' },
        { name: 'engagedSessions' },
        { name: 'averageSessionDuration' },
        { name: 'eventCount' },
      ],
      dimensionFilter: {
        filter: {
          fieldName: 'sessionDefaultChannelGroup',
          stringFilter: { value: 'Organic Search', matchType: 'EXACT' },
        },
      },
    },
    propertyId
  );

  const rows = rowsToMap(report);
  const row = rows[0] || {};
  return {
    sessions: row.sessions || 0,
    engagedSessions: row.engagedSessions || 0,
    avgEngagementTime: row.averageSessionDuration || 0,
    totalEvents: row.eventCount || 0,
    hasData: rows.length > 0,
  };
}

async function getEcommerceFunnel(monthStr, propertyId) {
  const { startDate, endDate } = monthDateRange(monthStr);
  const eventReport = await runReport(
    {
      dateRanges: [{ startDate, endDate }],
      dimensions: [{ name: 'eventName' }],
      metrics: [{ name: 'eventCount' }],
      dimensionFilter: {
        filter: {
          fieldName: 'eventName',
          inListFilter: {
            values: ['add_to_cart', 'begin_checkout', 'add_payment_info', 'purchase'],
          },
        },
      },
    },
    propertyId
  );

  const revenueReport = await runReport(
    {
      dateRanges: [{ startDate, endDate }],
      metrics: [{ name: 'purchaseRevenue' }],
    },
    propertyId
  );

  const rows = rowsToMap(eventReport);
  const counts = { add_to_cart: 0, begin_checkout: 0, add_payment_info: 0, purchase: 0 };
  rows.forEach((r) => {
    if (counts.hasOwnProperty(r.eventName)) counts[r.eventName] = r.eventCount;
  });

  const purchaseRevenue = sumMetric(revenueReport, 'purchaseRevenue');

  const totalActivity = counts.add_to_cart + counts.begin_checkout + counts.add_payment_info + counts.purchase + purchaseRevenue;

  return {
    addToCart: counts.add_to_cart,
    beginCheckout: counts.begin_checkout,
    addPaymentInfo: counts.add_payment_info,
    purchase: counts.purchase,
    purchaseRevenue,
    hasData: totalActivity > 0,
  };
}

async function getSessionsByPagePath(monthStr, pagePaths, propertyId) {
  const { startDate, endDate } = monthDateRange(monthStr);
  const report = await runReport(
    {
      dateRanges: [{ startDate, endDate }],
      dimensions: [{ name: 'pagePath' }],
      metrics: [{ name: 'sessions' }, { name: 'engagedSessions' }],
      dimensionFilter: {
        filter: {
          fieldName: 'pagePath',
          inListFilter: { values: pagePaths },
        },
      },
    },
    propertyId
  );

  const rows = rowsToMap(report);
  const map = {};
  rows.forEach((r) => {
    map[r.pagePath] = { sessions: r.sessions || 0, engagedSessions: r.engagedSessions || 0 };
  });
  return map;
}

async function getTopLandingPages(currentMonth, comparisonMonth, propertyId) {
  const { startDate, endDate } = monthDateRange(currentMonth);
  const report = await runReport(
    {
      dateRanges: [{ startDate, endDate }],
      dimensions: [{ name: 'pagePath' }],
      metrics: [{ name: 'sessions' }, { name: 'engagedSessions' }],
      orderBys: [{ metric: { metricName: 'sessions' }, desc: true }],
      limit: 10,
    },
    propertyId
  );

  const rows = rowsToMap(report);
  if (rows.length === 0) {
    return { pages: [], hasData: false };
  }

  const topPaths = rows.map((r) => r.pagePath);
  const comparisonMap = await getSessionsByPagePath(comparisonMonth, topPaths, propertyId);

  const pages = rows
    .map((r) => {
      const sessions = r.sessions || 0;
      const engagedSessions = r.engagedSessions || 0;
      const comparison = comparisonMap[r.pagePath];
      const prevSessions = comparison ? comparison.sessions : null;
      const change = prevSessions === null ? null : pctChange(sessions, prevSessions);
      return { pagePath: r.pagePath, sessions, engagedSessions, prevSessions, change };
    })
    .sort((a, b) => b.sessions - a.sessions);

  return { pages, hasData: true };
}

async function getTrafficOverview(currentMonth, comparisonMonth, propertyId) {
  const channels = ['Organic Search', 'Referral', 'Organic Social', 'Direct'];

  async function fetchChannelSessions(m) {
    const { startDate, endDate } = monthDateRange(m);
    const report = await runReport(
      {
        dateRanges: [{ startDate, endDate }],
        dimensions: [{ name: 'sessionDefaultChannelGroup' }],
        metrics: [{ name: 'sessions' }],
        dimensionFilter: {
          filter: {
            fieldName: 'sessionDefaultChannelGroup',
            inListFilter: { values: channels },
          },
        },
      },
      propertyId
    );
    const rows = rowsToMap(report);
    const map = {};
    channels.forEach((c) => (map[c] = 0));
    rows.forEach((r) => {
      map[r.sessionDefaultChannelGroup] = r.sessions || 0;
    });
    return map;
  }

  const [curr, prev] = await Promise.all([fetchChannelSessions(currentMonth), fetchChannelSessions(comparisonMonth)]);

  const rows = channels.map((c) => ({
    channel: c,
    current: curr[c],
    previous: prev[c],
    change: pctChange(curr[c], prev[c]),
  }));

  return {
    rows,
    hasData: rows.some((r) => r.current > 0 || r.previous > 0),
  };
}

function pctChange(current, previous) {
  if (!previous) return current > 0 ? 100 : 0;
  return ((current - previous) / previous) * 100;
}

async function getFullReportData(currentMonth, comparisonMonth, propertyId) {
  const [
    organicCurr,
    organicPrev,
    funnelCurr,
    funnelPrev,
    landingPages,
    trafficOverview,
    currSessionTotal,
    prevSessionTotal,
  ] = await Promise.all([
    getOrganicSearchMetrics(currentMonth, propertyId),
    getOrganicSearchMetrics(comparisonMonth, propertyId),
    getEcommerceFunnel(currentMonth, propertyId),
    getEcommerceFunnel(comparisonMonth, propertyId),
    getTopLandingPages(currentMonth, comparisonMonth, propertyId),
    getTrafficOverview(currentMonth, comparisonMonth, propertyId),
    getMonthSessionTotal(currentMonth, propertyId),
    getMonthSessionTotal(comparisonMonth, propertyId),
  ]);

  // Whether GA4 returned *any* sessions at all for a month, independent of channel/event
  // filters — distinguishes "before tracking started" from a section-specific empty result.
  const hasAnyGa4Data = currSessionTotal > 0;
  const comparisonMonthHasGa4Data = prevSessionTotal > 0;

  return {
    currentMonth,
    comparisonMonth,
    hasAnyGa4Data,
    comparisonMonthHasGa4Data,
    organicSearch: {
      current: organicCurr,
      previous: organicPrev,
      changes: comparisonMonthHasGa4Data
        ? {
            sessions: pctChange(organicCurr.sessions, organicPrev.sessions),
            engagedSessions: pctChange(organicCurr.engagedSessions, organicPrev.engagedSessions),
            avgEngagementTime: pctChange(organicCurr.avgEngagementTime, organicPrev.avgEngagementTime),
            totalEvents: pctChange(organicCurr.totalEvents, organicPrev.totalEvents),
          }
        : { sessions: null, engagedSessions: null, avgEngagementTime: null, totalEvents: null },
    },
    ecommerce: {
      current: funnelCurr,
      previous: funnelPrev,
      changes: comparisonMonthHasGa4Data
        ? {
            addToCart: pctChange(funnelCurr.addToCart, funnelPrev.addToCart),
            beginCheckout: pctChange(funnelCurr.beginCheckout, funnelPrev.beginCheckout),
            addPaymentInfo: pctChange(funnelCurr.addPaymentInfo, funnelPrev.addPaymentInfo),
            purchase: pctChange(funnelCurr.purchase, funnelPrev.purchase),
            purchaseRevenue: pctChange(funnelCurr.purchaseRevenue, funnelPrev.purchaseRevenue),
          }
        : { addToCart: null, beginCheckout: null, addPaymentInfo: null, purchase: null, purchaseRevenue: null },
    },
    landingPages,
    trafficOverview: comparisonMonthHasGa4Data
      ? trafficOverview
      : { ...trafficOverview, rows: trafficOverview.rows.map((r) => ({ ...r, change: null })) },
  };
}

module.exports = {
  getFullReportData,
  monthDateRange,
  mapGa4Error,
};
