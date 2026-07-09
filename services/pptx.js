const PptxGenJS = require('pptxgenjs');

const BG = 'F8F9FA';
const WHITE = 'FFFFFF';
const ACCENT = '4361EE';
const POSITIVE = '2DC653';
const NEGATIVE = 'E63946';
const NEUTRAL = '6C757D';
const TEXT_DARK = '212529';
const ROW_ALT = 'F1F3F5';

const BRAND = 'WebrocketAI';

function monthLabel(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  return new Date(year, month - 1, 1).toLocaleString('en-US', { month: 'long', year: 'numeric' });
}

function fmtPct(n) {
  const sign = n >= 0 ? '+' : '';
  return `${sign}${n.toFixed(1)}%`;
}

function fmtNum(n) {
  return Math.round(n).toLocaleString('en-US');
}

function fmtSeconds(n) {
  const mins = Math.floor(n / 60);
  const secs = Math.round(n % 60);
  return `${mins}m ${secs}s`;
}

function fmtCurrency(n) {
  return `$${n.toLocaleString('en-US', { maximumFractionDigits: 0 })}`;
}

function trend(n) {
  return n >= 0 ? 'increasing' : 'decreasing';
}

function changeColor(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) return NEUTRAL;
  if (n > 0) return POSITIVE;
  if (n < 0) return NEGATIVE;
  return NEUTRAL;
}

// Marks a table cell as a colored % change value (green/red/grey), handling N/A.
function pctCell(n) {
  if (n === null || n === undefined || !Number.isFinite(n)) {
    return { text: 'N/A', highlight: NEUTRAL };
  }
  return { text: fmtPct(n), highlight: changeColor(n) };
}

function addSlideChrome(slide, title) {
  slide.background = { color: BG };
  slide.addText(title, {
    x: 0.5,
    y: 0.35,
    w: 9,
    h: 0.65,
    fontSize: 32,
    bold: true,
    color: ACCENT,
    fontFace: 'Arial',
  });
  slide.addShape('line', {
    x: 0.5,
    y: 1.05,
    w: 2.2,
    h: 0,
    line: { color: ACCENT, width: 2.5 },
  });
}

function addSummary(slide, text) {
  slide.addText(text, {
    x: 0.5,
    y: 1.3,
    w: 9,
    h: 0.9,
    fontSize: 13,
    color: TEXT_DARK,
    fontFace: 'Arial',
    valign: 'top',
  });
}

function addFooter(slide, monthStr) {
  slide.addText(`${BRAND} — ${monthLabel(monthStr)} Report`, {
    x: 6.0,
    y: 5.35,
    w: 3.5,
    h: 0.25,
    fontSize: 9,
    color: NEUTRAL,
    fontFace: 'Arial',
    align: 'right',
  });
}

function addTable(slide, rows, colWidths, opts = {}) {
  const { x = 0.5, y = 2.25, w = 9 } = opts;
  const tableRows = rows.map((row, rowIdx) => {
    const isHeader = rowIdx === 0;
    const isAltRow = !isHeader && (rowIdx - 1) % 2 === 1;

    return row.map((cell) => {
      const isSpecial = typeof cell === 'object' && cell !== null;
      const text = isSpecial ? cell.text : String(cell);
      const color = isHeader ? WHITE : isSpecial ? cell.highlight : TEXT_DARK;

      return {
        text,
        options: {
          color,
          fill: { color: isHeader ? ACCENT : isAltRow ? ROW_ALT : WHITE },
          bold: isHeader || isSpecial,
          fontSize: isHeader ? 12 : 11.5,
          fontFace: 'Arial',
          align: 'left',
          valign: 'middle',
        },
      };
    });
  });

  slide.addTable(tableRows, {
    x,
    y,
    w,
    colW: colWidths,
    border: { type: 'solid', color: ROW_ALT, pt: 1 },
    autoPage: false,
  });
}

function buildTrafficOverviewSlide(pptx, data) {
  const slide = pptx.addSlide();
  addSlideChrome(slide, 'Traffic Overview');

  const { rows: channelRows, hasData } = data.trafficOverview;
  const currLabel = monthLabel(data.currentMonth);
  const prevLabel = monthLabel(data.comparisonMonth);

  if (!hasData) {
    addSummary(
      slide,
      `No session data by channel was returned for ${currLabel} or ${prevLabel}. Confirm channel grouping is configured for this property.`
    );
    addFooter(slide, data.currentMonth);
    return slide;
  }

  const prevHasData = data.comparisonMonthHasGa4Data;

  let summary;
  if (!prevHasData) {
    summary = `Session data by channel for ${currLabel} is shown below. No comparison period data available — ${prevLabel} had no recorded GA4 sessions.`;
  } else {
    const sorted = [...channelRows].sort((a, b) => b.change - a.change);
    const topGrowth = sorted[0];
    const topDecline = sorted[sorted.length - 1];

    summary = `Total traffic across the four core channels ${trend(
      channelRows.reduce((s, r) => s + r.change, 0)
    )} in ${currLabel} compared to ${prevLabel}. ${topGrowth.channel} traffic grew the most, ${trend(
      topGrowth.change
    )} by ${fmtPct(topGrowth.change)}`;

    if (topDecline.channel !== topGrowth.channel) {
      summary += `, while ${topDecline.channel} traffic changed by ${fmtPct(topDecline.change)}`;
    }
    summary += '.';
  }

  addSummary(slide, summary);

  const rows = [['Channel', currLabel, prevLabel, '% Change']];
  channelRows.forEach((r) => {
    rows.push([
      r.channel,
      fmtNum(r.current),
      prevHasData ? fmtNum(r.previous) : 'No comparison data',
      pctCell(r.change),
    ]);
  });

  addTable(slide, rows, [3, 2, 2, 2]);
  addFooter(slide, data.currentMonth);

  return slide;
}

const KEYWORD_BUCKETS = [
  ['Top 10', 'top10'],
  ['Top 11-30', 'top11_30'],
  ['Top 31-50', 'top31_50'],
  ['Top 51-100', 'top51_100'],
  ['Pending', 'pending'],
];

function buildSeoOverviewSlide(pptx, data) {
  const slide = pptx.addSlide();
  addSlideChrome(slide, 'SEO Performance Overview');

  const seo = data.seo;
  const currLabel = monthLabel(data.currentMonth);
  const prevLabel = monthLabel(data.comparisonMonth);

  const hasAnyData = seo && (seo.keywordRankings.hasData || seo.offPage.hasData);

  if (!hasAnyData) {
    addSummary(
      slide,
      `No SEO data (keyword rankings or off-page submissions) was found in the connected Google Sheet for ${currLabel} or ${prevLabel}.`
    );
    addFooter(slide, data.currentMonth);
    return slide;
  }

  const { keywordRankings, offPage } = seo;

  const summaryParts = [];
  if (keywordRankings.hasData) {
    summaryParts.push(
      `Keywords ranking in the Top 10 ${trend(keywordRankings.changes.top10)} from ${
        keywordRankings.previous.top10
      } to ${keywordRankings.current.top10} versus ${prevLabel}.`
    );
  } else {
    summaryParts.push('Keyword ranking data was not available for this period.');
  }
  if (offPage.hasData) {
    summaryParts.push(
      `Off-page submissions ${trend(offPage.change)} to ${fmtNum(offPage.current)} in ${currLabel}, compared to ${fmtNum(
        offPage.previous
      )} in ${prevLabel} (${fmtPct(offPage.change)}).`
    );
  } else {
    summaryParts.push('Off-page submission data was not available for this period.');
  }
  addSummary(slide, summaryParts.join(' '));

  const leftRows = [['Bucket', prevLabel, currLabel, '% Change']];
  if (keywordRankings.hasData) {
    KEYWORD_BUCKETS.forEach(([label, key]) => {
      leftRows.push([
        label,
        fmtNum(keywordRankings.previous[key]),
        fmtNum(keywordRankings.current[key]),
        pctCell(keywordRankings.changes[key]),
      ]);
    });
  } else {
    leftRows.push(['Keyword Rankings', 'Not configured', 'Not configured', 'Not configured']);
  }
  addTable(slide, leftRows, [1.6, 1.1, 1.1, 1.2], { x: 0.5, y: 2.25, w: 5 });

  const rightRows = [['Metric', prevLabel, currLabel, '% Change']];
  if (offPage.hasData) {
    rightRows.push([
      'Off-Page Submissions',
      fmtNum(offPage.previous),
      fmtNum(offPage.current),
      pctCell(offPage.change),
    ]);
  } else {
    rightRows.push(['Off-Page Submissions', 'Not configured', 'Not configured', 'Not configured']);
  }
  addTable(slide, rightRows, [1.8, 1.1, 1.1, 1.2], { x: 5.7, y: 2.25, w: 4 });

  addFooter(slide, data.currentMonth);

  return slide;
}

function buildOrganicSearchSlide(pptx, data) {
  const slide = pptx.addSlide();
  addSlideChrome(slide, 'Organic Search Performance');

  const { current, changes } = data.organicSearch;
  const label = monthLabel(data.currentMonth);

  if (!current.hasData) {
    addSummary(
      slide,
      `No organic search data was returned for ${label}. This may mean the property has no Organic Search traffic in this period, or channel grouping is not configured.`
    );
    addFooter(slide, data.currentMonth);
    return slide;
  }

  const prevHasData = data.comparisonMonthHasGa4Data;
  const previous = data.organicSearch.previous;

  const summary = prevHasData
    ? `Organic search delivered ${trend(changes.sessions)} results in ${label}, with sessions ${trend(
        changes.sessions
      )} by ${fmtPct(changes.sessions)} and engaged sessions ${trend(changes.engagedSessions)} by ${fmtPct(
        changes.engagedSessions
      )} versus the comparison period. Average engagement time moved ${fmtPct(
        changes.avgEngagementTime
      )}, while total events tied to organic sessions changed by ${fmtPct(changes.totalEvents)}.`
    : `Organic search delivered ${fmtNum(current.sessions)} sessions and ${fmtNum(
        current.engagedSessions
      )} engaged sessions in ${label}. No comparison period data available.`;

  addSummary(slide, summary);

  addTable(
    slide,
    [
      ['Metric', 'Current Month', 'Comparison Period', '% Change'],
      [
        'Sessions',
        fmtNum(current.sessions),
        prevHasData ? fmtNum(previous.sessions) : 'No comparison data',
        pctCell(changes.sessions),
      ],
      [
        'Engaged Sessions',
        fmtNum(current.engagedSessions),
        prevHasData ? fmtNum(previous.engagedSessions) : 'No comparison data',
        pctCell(changes.engagedSessions),
      ],
      [
        'Avg Engagement Time',
        fmtSeconds(current.avgEngagementTime),
        prevHasData ? fmtSeconds(previous.avgEngagementTime) : 'No comparison data',
        pctCell(changes.avgEngagementTime),
      ],
      [
        'Total Events',
        fmtNum(current.totalEvents),
        prevHasData ? fmtNum(previous.totalEvents) : 'No comparison data',
        pctCell(changes.totalEvents),
      ],
    ],
    [3, 2, 2, 2]
  );
  addFooter(slide, data.currentMonth);

  return slide;
}

function buildEcommerceSlide(pptx, data) {
  const slide = pptx.addSlide();
  addSlideChrome(slide, 'Ecommerce Funnel Performance');

  const { current, previous, changes } = data.ecommerce;
  const label = monthLabel(data.currentMonth);

  if (!current.hasData) {
    addSummary(
      slide,
      'Ecommerce tracking is not configured for this property. The following events need to be set up to populate this slide: add_to_cart, begin_checkout, add_payment_info, purchase.'
    );

    addTable(
      slide,
      [
        ['Metric', 'Current Month', 'Comparison Period', '% Change'],
        ['Add to Cart', 'Not configured', 'Not configured', 'Not configured'],
        ['Checkout', 'Not configured', 'Not configured', 'Not configured'],
        ['Payment', 'Not configured', 'Not configured', 'Not configured'],
        ['Purchase', 'Not configured', 'Not configured', 'Not configured'],
        ['Revenue', 'Not configured', 'Not configured', 'Not configured'],
      ],
      [3, 2, 2, 2]
    );
    addFooter(slide, data.currentMonth);

    return slide;
  }

  const prevHasData = data.comparisonMonthHasGa4Data;

  const summary = prevHasData
    ? `The ecommerce funnel ${trend(changes.purchase)} in ${label}, with purchases ${trend(
        changes.purchase
      )} by ${fmtPct(changes.purchase)} and purchase revenue ${trend(changes.purchaseRevenue)} by ${fmtPct(
        changes.purchaseRevenue
      )} versus the comparison period. Upper-funnel activity showed add-to-cart events changing by ${fmtPct(
        changes.addToCart
      )} and checkout initiations changing by ${fmtPct(changes.beginCheckout)}.`
    : `The ecommerce funnel recorded ${fmtNum(current.purchase)} purchases and ${fmtCurrency(
        current.purchaseRevenue
      )} in revenue during ${label}. No comparison period data available.`;

  addSummary(slide, summary);

  addTable(
    slide,
    [
      ['Metric', 'Current Month', 'Comparison Period', '% Change'],
      [
        'Add to Cart',
        fmtNum(current.addToCart),
        prevHasData ? fmtNum(previous.addToCart) : 'No comparison data',
        pctCell(changes.addToCart),
      ],
      [
        'Begin Checkout',
        fmtNum(current.beginCheckout),
        prevHasData ? fmtNum(previous.beginCheckout) : 'No comparison data',
        pctCell(changes.beginCheckout),
      ],
      [
        'Add Payment Info',
        fmtNum(current.addPaymentInfo),
        prevHasData ? fmtNum(previous.addPaymentInfo) : 'No comparison data',
        pctCell(changes.addPaymentInfo),
      ],
      [
        'Purchase',
        fmtNum(current.purchase),
        prevHasData ? fmtNum(previous.purchase) : 'No comparison data',
        pctCell(changes.purchase),
      ],
      [
        'Purchase Revenue',
        fmtCurrency(current.purchaseRevenue),
        prevHasData ? fmtCurrency(previous.purchaseRevenue) : 'No comparison data',
        pctCell(changes.purchaseRevenue),
      ],
    ],
    [3, 2, 2, 2]
  );
  addFooter(slide, data.currentMonth);

  return slide;
}

function displayPagePath(pagePath) {
  return pagePath === '/' ? 'Homepage' : pagePath;
}

function buildLandingPagesSlide(pptx, data) {
  const slide = pptx.addSlide();
  addSlideChrome(slide, 'Top Landing Pages');

  const { pages, hasData } = data.landingPages;
  const label = monthLabel(data.currentMonth);

  if (!hasData) {
    addSummary(slide, `No landing page session data was returned for ${label}.`);
    addFooter(slide, data.currentMonth);
    return slide;
  }

  const topPage = pages[0];
  const growthPages = pages
    .filter((p) => p.change !== null && Number.isFinite(p.change))
    .sort((a, b) => b.change - a.change);
  const topGrowthPage = growthPages[0];

  let summary = `The top landing page in ${label} was ${displayPagePath(topPage.pagePath)}, driving ${fmtNum(
    topPage.sessions
  )} sessions with ${fmtNum(topPage.engagedSessions)} engaged sessions.`;

  if (topGrowthPage && topGrowthPage.change > 0) {
    summary += ` ${displayPagePath(topGrowthPage.pagePath)} showed the strongest growth versus the comparison period, with sessions up ${fmtPct(
      topGrowthPage.change
    )}.`;
  }

  summary += ' The top 10 pages below are ranked by current month sessions.';

  if (!data.comparisonMonthHasGa4Data) {
    summary += ` No comparison period data available — ${monthLabel(data.comparisonMonth)} had no recorded GA4 sessions.`;
  }

  addSummary(slide, summary);

  const rows = [['Page Path', 'Sessions', 'Prev. Sessions', 'Engaged Sessions', '% Change']];
  pages.forEach((p) => {
    rows.push([
      displayPagePath(p.pagePath),
      fmtNum(p.sessions),
      p.prevSessions === null ? 'N/A' : fmtNum(p.prevSessions),
      fmtNum(p.engagedSessions),
      pctCell(p.change),
    ]);
  });

  addTable(slide, rows, [3.5, 1.5, 1.5, 1.5, 1.5]);
  addFooter(slide, data.currentMonth);

  return slide;
}

async function generateReportPptx(data) {
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: 10, height: 5.63 });
  pptx.layout = 'WIDE';

  buildTrafficOverviewSlide(pptx, data);
  buildSeoOverviewSlide(pptx, data);
  buildOrganicSearchSlide(pptx, data);
  buildEcommerceSlide(pptx, data);
  buildLandingPagesSlide(pptx, data);

  return pptx.write({ outputType: 'nodebuffer' });
}

module.exports = { generateReportPptx, monthLabel };
