const express = require('express');
const { getFullReportData, mapGa4Error } = require('../services/ga4');
const { getSeoOverview } = require('../services/sheets');
const { getMetaAdsData } = require('../services/meta');
const { generateReportPptx, monthLabel } = require('../services/pptx');
const { getClientById } = require('../services/clients');

const router = express.Router();

const MONTH_RE = /^\d{4}-\d{2}$/;

function emptySeoData() {
  const buckets = { top10: 0, top11_30: 0, top31_50: 0, top51_100: 0, pending: 0 };
  const changes = { top10: null, top11_30: null, top31_50: null, top51_100: null, pending: null };
  return {
    keywordRankings: { current: buckets, previous: { ...buckets }, changes, hasData: false },
    offPage: { current: 0, previous: 0, change: null, hasData: false },
  };
}

router.post('/generate-report', async (req, res) => {
  const { clientId, currentMonth, comparisonMonth, paidMedia } = req.body;

  if (!clientId) {
    return res.status(400).json({ error: 'Please select a client.' });
  }
  if (!currentMonth || !MONTH_RE.test(currentMonth) || !comparisonMonth || !MONTH_RE.test(comparisonMonth)) {
    return res.status(400).json({ error: 'Please provide a valid current period and comparison period in YYYY-MM format.' });
  }

  const client = getClientById(clientId);
  if (!client) {
    return res.status(400).json({ error: 'Selected client was not found. It may have been removed.' });
  }

  const email = req.session.email;

  try {
    const [gaData, seoData] = await Promise.all([
      getFullReportData(currentMonth, comparisonMonth, client.ga4PropertyId, email),
      client.sheetId ? getSeoOverview(currentMonth, comparisonMonth, client.sheetId, email) : Promise.resolve(emptySeoData()),
    ]);

    // Meta Ads is fetched separately and never fails the whole report — a bad
    // token or missing ad account just skips the slide and surfaces a warning.
    let metaAdsData = null;
    let metaAdsWarning = null;
    if (client.metaAdAccountId) {
      try {
        metaAdsData = await getMetaAdsData(client.metaAdAccountId, currentMonth);
      } catch (metaErr) {
        console.error('Meta Ads API error:', metaErr.code, metaErr.message);
        metaAdsWarning = `Meta Ads data could not be fetched: ${metaErr.message}`;
      }
    }

    // When a Meta Ad Account ID is configured, API data replaces the manual Meta
    // Ads form fields on the Paid Media slide; manual entry is only a fallback
    // for clients with no ad account configured.
    const effectivePaidMedia = {
      googleAds: paidMedia && paidMedia.googleAds,
      metaAds: client.metaAdAccountId ? metaAdsData : paidMedia && paidMedia.metaAds,
    };

    const data = { ...gaData, seo: seoData, clientName: client.name, paidMedia: effectivePaidMedia };
    const buffer = await generateReportPptx(data);

    const currentLabel = monthLabel(currentMonth);
    const comparisonLabel = monthLabel(comparisonMonth);

    const warnings = [];
    if (!gaData.hasAnyGa4Data) {
      warnings.push(
        `No data found in GA4 for ${currentLabel}. This period may be before tracking was set up, or no traffic was recorded during this time.`
      );
    }
    if (!seoData.keywordRankings.hasData) {
      warnings.push(
        `No keyword ranking data found for ${currentLabel} in the SEO tracker. The sheet may not have a date column within this month.`
      );
    }
    if (!seoData.offPage.hasData) {
      warnings.push(`No off-page submission data found for ${currentLabel}.`);
    }
    if (gaData.hasAnyGa4Data && !gaData.comparisonMonthHasGa4Data) {
      warnings.push(`No comparison period data available — ${comparisonLabel} has no recorded GA4 sessions.`);
    }
    if (metaAdsWarning) {
      warnings.push(metaAdsWarning);
    }

    const filename = `WebrocketAI-Report-${currentLabel.replace(' ', '-')}-vs-${comparisonLabel.replace(' ', '-')}.pptx`;

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    );
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('X-Report-Warnings', encodeURIComponent(JSON.stringify(warnings)));
    res.setHeader('Access-Control-Expose-Headers', 'X-Report-Warnings, Content-Disposition');
    res.send(buffer);
  } catch (e) {
    if (e.code === 'NOT_AUTHENTICATED' || e.code === 'AUTH_EXPIRED' || e.code === 'SHEETS_SCOPE_MISSING') {
      return res.status(401).json({ error: e.message, code: e.code });
    }
    if (e.code === 'GA4_API_ERROR') {
      console.error('GA4 API error:', e.status, e.apiErrorStatus, e.message);
      const { httpStatus, message } = mapGa4Error(e);
      return res.status(httpStatus).json({ error: message });
    }
    if (e.code === 'SHEETS_API_ERROR') {
      console.error('Sheets API error:', e.status, e.message);
      return res.status(502).json({
        error: 'Google Sheets returned an error while fetching SEO data for this period. Please try again or check the sheet configuration.',
      });
    }
    console.error('Report generation error:', e.message);
    return res.status(500).json({ error: 'Failed to generate report. Please try again.' });
  }
});

module.exports = router;
