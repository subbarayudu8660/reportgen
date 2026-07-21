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

// `null` means "previous period was 0, current is nonzero" (pctChange's "New"
// case) — a real percent change is undefined in that scenario, so render "New"
// rather than a misleading number. `undefined`/NaN means no comparison data at all.
function fmtPct(n) {
  if (n === null) return 'New';
  if (n === undefined || !Number.isFinite(n)) return 'N/A';
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
  if (n === null) return POSITIVE;
  if (n === undefined || !Number.isFinite(n)) return NEUTRAL;
  if (n > 0) return POSITIVE;
  if (n < 0) return NEGATIVE;
  return NEUTRAL;
}

// Marks a table cell as a colored % change value (green/red/grey), handling
// "New" (null — previous period was 0) and "N/A" (undefined/NaN — no comparison data).
function pctCell(n) {
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

function addFooter(slide, monthStr, clientName) {
  slide.addText(`${clientName || BRAND} | Digital Marketing Report | ${monthLabel(monthStr)}`, {
    x: 4.5,
    y: 5.35,
    w: 4.5,
    h: 0.25,
    fontSize: 9,
    color: NEUTRAL,
    fontFace: 'Arial',
    align: 'right',
  });
  slide.slideNumber = {
    x: 9.3,
    y: 5.35,
    w: 0.4,
    h: 0.25,
    fontSize: 9,
    color: NEUTRAL,
    fontFace: 'Arial',
  };
}

function addTable(slide, rows, colWidths, opts = {}) {
  const { x = 0.5, y = 2.25, w = 9, rowH, headerFontSize = 12, bodyFontSize = 11.5 } = opts;
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
          fontSize: isHeader ? headerFontSize : bodyFontSize,
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
    ...(rowH ? { rowH } : {}),
    border: { type: 'solid', color: ROW_ALT, pt: 1 },
    autoPage: false,
  });
}

// Picks a row height (and matching font sizes) for a table with `numRows`
// (header + data) rows so it fits within `availableH` inches without
// overlapping the footer. Prefers the ideal row height; falls back to a
// smaller fixed row height, then finally shrinks both row height and font
// size together so a full 10-row landing-pages table never overflows.
function fitTableRows(numRows, availableH) {
  const IDEAL_ROW_H = 0.34;
  const REDUCED_ROW_H = 0.28;

  if (numRows * IDEAL_ROW_H <= availableH) {
    return { rowH: IDEAL_ROW_H, headerFontSize: 12, bodyFontSize: 11.5 };
  }
  if (numRows * REDUCED_ROW_H <= availableH) {
    return { rowH: REDUCED_ROW_H, headerFontSize: 12, bodyFontSize: 11.5 };
  }
  const rowH = Math.max(0.22, availableH / numRows);
  return { rowH, headerFontSize: 10.5, bodyFontSize: 9.5 };
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
    addFooter(slide, data.currentMonth, data.clientName);
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
  addFooter(slide, data.currentMonth, data.clientName);

  return slide;
}

const KEYWORD_BUCKETS = [
  ['Top 10', 'top10'],
  ['Top 11-30', 'top11_30'],
  ['Top 31-50', 'top31_50'],
  ['Top 51-100', 'top51_100'],
  ['Pending (100+)', 'pending'],
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
    addFooter(slide, data.currentMonth, data.clientName);
    return slide;
  }

  const { keywordRankings, offPage } = seo;

  const summaryParts = [];
  if (keywordRankings.hasData) {
    if (keywordRankings.previous !== null) {
      summaryParts.push(
        `Keywords ranking in the Top 10 ${trend(keywordRankings.changes.top10)} from ${
          keywordRankings.previous.top10
        } to ${keywordRankings.current.top10} versus ${prevLabel}.`
      );
    } else {
      summaryParts.push(
        `Keywords ranking in the Top 10 stood at ${keywordRankings.current.top10} for ${currLabel}. No keyword ranking data was available for ${prevLabel}.`
      );
    }
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
    // `current` is guaranteed non-null when hasData is true; `previous` can
    // still be null on its own (e.g. the comparison period predates the
    // sheet's tracking), in which case it renders as "No data", not "0".
    KEYWORD_BUCKETS.forEach(([label, key]) => {
      leftRows.push([
        label,
        keywordRankings.previous === null ? 'No data' : fmtNum(keywordRankings.previous[key]),
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

  addFooter(slide, data.currentMonth, data.clientName);

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
    addFooter(slide, data.currentMonth, data.clientName);
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
  addFooter(slide, data.currentMonth, data.clientName);

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
    addFooter(slide, data.currentMonth, data.clientName);

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
  addFooter(slide, data.currentMonth, data.clientName);

  return slide;
}

function displayPagePath(pagePath) {
  return pagePath === '/' ? 'Homepage' : pagePath;
}

// A stacked hero-stat box: a big bold number (optionally followed by a small
// inline suffix run, e.g. "pages/user") over a smaller descriptive label,
// on a light rounded card — used for the Top Landing Pages slide's stat column.
function addStatCard(slide, { x, y, w, h, valueRuns, label }) {
  slide.addShape('roundRect', {
    x,
    y,
    w,
    h,
    fill: { color: ROW_ALT },
    line: { type: 'none' },
    rectRadius: 0.08,
  });

  const valueH = h * 0.42;
  slide.addText(valueRuns, {
    x,
    y: y + 0.06,
    w,
    h: valueH,
    align: 'center',
    valign: 'bottom',
    fontFace: 'Arial',
  });

  slide.addText(label, {
    x: x + 0.12,
    y: y + 0.06 + valueH,
    w: w - 0.24,
    h: h - valueH - 0.12,
    align: 'center',
    valign: 'top',
    fontSize: 8.5,
    color: NEUTRAL,
    fontFace: 'Arial',
  });
}

function buildLandingPagesSlide(pptx, data) {
  const slide = pptx.addSlide();
  addSlideChrome(slide, 'Top Landing Pages');

  const { pages, hasData, screenPageViews, activeUsers, pagesPerUser } = data.landingPages;
  const label = monthLabel(data.currentMonth);

  if (!hasData) {
    addSummary(slide, `No landing page session data was returned for ${label}.`);
    addFooter(slide, data.currentMonth, data.clientName);
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

  // Cap defensively at 10 rows — `getTopLandingPages` already limits the GA4
  // query to 10 pages, but the slide shouldn't rely solely on the upstream cap.
  const topPages = pages.slice(0, 10);

  const rows = [['Page Path', 'Sessions', 'Prev. Sessions', 'Engaged Sessions', '% Change']];
  topPages.forEach((p) => {
    rows.push([
      displayPagePath(p.pagePath),
      fmtNum(p.sessions),
      p.prevSessions === null ? 'N/A' : fmtNum(p.prevSessions),
      fmtNum(p.engagedSessions),
      pctCell(p.change),
    ]);
  });

  // Table takes ~65% of the content width; the remaining ~30% (with a small
  // gap) holds three stacked hero stat cards.
  const tableX = 0.5;
  const tableY = 2.25;
  const tableW = 6.2;
  const FOOTER_Y = 5.35;
  const MIN_FOOTER_GAP = 0.3;

  const numTableRows = rows.length; // header + data rows
  const availableH = FOOTER_Y - MIN_FOOTER_GAP - tableY;
  const { rowH, headerFontSize, bodyFontSize } = fitTableRows(numTableRows, availableH);
  const tableH = numTableRows * rowH;

  addTable(slide, rows, [2.3, 1.0, 1.05, 1.05, 0.8], {
    x: tableX,
    y: tableY,
    w: tableW,
    rowH,
    headerFontSize,
    bodyFontSize,
  });

  const cardX = 7.0;
  const cardW = 2.5;
  let cardH = 0.85;
  let cardGap = 0.15;
  let cardsTotalH = 3 * cardH + 2 * cardGap;

  // Shrink the cards (keeping their proportions) if the full stack wouldn't
  // fit in the space above the footer — mirrors the table's own fallback.
  if (cardsTotalH > availableH) {
    const scale = availableH / cardsTotalH;
    cardH *= scale;
    cardGap *= scale;
    cardsTotalH = availableH;
  }

  // Vertically center the stat card column on the table rather than pinning
  // both to the same top edge, so a shorter/taller table stays balanced.
  const tableCenterY = tableY + tableH / 2;
  const cardsStartY = Math.min(
    Math.max(tableY, tableCenterY - cardsTotalH / 2),
    FOOTER_Y - MIN_FOOTER_GAP - cardsTotalH
  );

  addStatCard(slide, {
    x: cardX,
    y: cardsStartY,
    w: cardW,
    h: cardH,
    valueRuns: [{ text: fmtNum(screenPageViews), options: { fontSize: 26, bold: true, color: ACCENT } }],
    label: 'Total page views recorded during this reporting period',
  });

  addStatCard(slide, {
    x: cardX,
    y: cardsStartY + cardH + cardGap,
    w: cardW,
    h: cardH,
    valueRuns: [{ text: fmtNum(activeUsers), options: { fontSize: 26, bold: true, color: ACCENT } }],
    label: 'Active users who viewed the site',
  });

  const pagesPerUserText = pagesPerUser === null ? 'N/A' : pagesPerUser.toFixed(2);
  addStatCard(slide, {
    x: cardX,
    y: cardsStartY + 2 * (cardH + cardGap),
    w: cardW,
    h: cardH,
    valueRuns: [
      { text: pagesPerUserText, options: { fontSize: 26, bold: true, color: ACCENT } },
      { text: ' pages/user', options: { fontSize: 10, color: TEXT_DARK } },
    ],
    label:
      pagesPerUser === null
        ? 'Pages-per-user could not be calculated for this period.'
        : `On average, each visitor viewed ${pagesPerUserText} pages during their visit, indicating they explored multiple pages across the website.`,
  });

  addFooter(slide, data.currentMonth, data.clientName);

  return slide;
}

// Parses a free-text currency string like "CAD $154.77" into a prefix + numeric amount.
function parseCurrencyString(str) {
  if (!str) return null;
  const match = String(str).trim().match(/^(.*?)([\d,]+(?:\.\d+)?)\s*$/);
  if (!match) return null;
  const amount = parseFloat(match[2].replace(/,/g, ''));
  if (!Number.isFinite(amount)) return null;
  return { prefix: match[1], amount };
}

// Combines spend + GST into a display string. Adds numerically when both share a
// recognizable currency format; otherwise falls back to a plain concatenation.
function withGstText(spend, gst) {
  const parsedSpend = parseCurrencyString(spend);
  const parsedGst = parseCurrencyString(gst);
  if (parsedSpend && parsedGst) {
    const total = parsedSpend.amount + parsedGst.amount;
    return `${parsedSpend.prefix}${total.toLocaleString('en-US', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
  }
  return `${spend} + ${gst} (GST)`;
}

function hasAnyValue(obj) {
  return Boolean(obj) && Object.values(obj).some((v) => v !== null && v !== undefined && v !== '');
}

function paidMediaCell(value) {
  return value === null || value === undefined || value === '' ? '—' : String(value);
}

function addPaidMediaSectionHeader(slide, label, y) {
  slide.addShape('rect', {
    x: 0.5,
    y,
    w: 9,
    h: 0.35,
    fill: { color: ACCENT },
    line: { type: 'none' },
  });
  slide.addText(label, {
    x: 0.6,
    y,
    w: 8.8,
    h: 0.35,
    fontSize: 13,
    bold: true,
    color: WHITE,
    fontFace: 'Arial',
    valign: 'middle',
  });
}

// API-sourced Google Ads data carries a `cpc` field (computed server-side from
// totalSpend/clicks); manual form entries never set it. Column order/set differs
// slightly between the two since CPC only makes sense once real spend+click data exists.
function buildGoogleAdsTable(slide, googleAds, y) {
  const numOrDash = (n) => (n === null || n === undefined ? null : fmtNum(n));
  const isApiSourced = googleAds.cpc !== null && googleAds.cpc !== undefined;

  if (isApiSourced) {
    const header = ['Total Campaigns', 'Spend', 'Impressions', 'Clicks', 'Conversions', 'CPC'];
    const dataRow = [
      'Total',
      paidMediaCell(numOrDash(googleAds.totalCampaigns)),
      paidMediaCell(googleAds.totalSpend),
      paidMediaCell(numOrDash(googleAds.impressions)),
      paidMediaCell(numOrDash(googleAds.clicks)),
      paidMediaCell(numOrDash(googleAds.conversions)),
      paidMediaCell(googleAds.cpc),
    ];
    addTable(slide, [['', ...header], dataRow], [0.9, 1.4, 1.3, 1.3, 1.2, 1.4, 1.5], { x: 0.5, y, w: 9 });
    return;
  }

  const header = ['Total Campaigns', 'Conversions', 'Spend', 'Impressions', 'Clicks'];
  const dataRow = [
    'Total',
    paidMediaCell(numOrDash(googleAds.totalCampaigns)),
    paidMediaCell(numOrDash(googleAds.conversions)),
    paidMediaCell(googleAds.totalSpend),
    paidMediaCell(numOrDash(googleAds.impressions)),
    paidMediaCell(numOrDash(googleAds.clicks)),
  ];

  const rows = [['', ...header], dataRow];
  if (googleAds.gstAmount) {
    rows.push(['With GST', '', '', withGstText(googleAds.totalSpend, googleAds.gstAmount), '', '']);
  }

  addTable(slide, rows, [1.4, 1.6, 1.4, 1.8, 1.4, 1.4], { x: 0.5, y, w: 9 });
}

function googleAdsSummaryText(googleAds, monthStr) {
  return `Google Ads delivered ${fmtNum(googleAds.impressions)} impressions and ${fmtNum(
    googleAds.clicks
  )} clicks in ${monthLabel(monthStr)}, with a total spend of ${googleAds.totalSpend} across ${
    googleAds.totalCampaigns
  } campaigns and ${fmtNum(googleAds.conversions)} conversions.`;
}

function buildMetaAdsTable(slide, metaAds, y) {
  const header = ['Total Campaigns', 'Form Leads', 'Message Leads', 'Spend', 'Impressions', 'Clicks', 'Reach'];
  const numOrDash = (n) => (n === null || n === undefined ? null : fmtNum(n));
  const dataRow = [
    'Total',
    paidMediaCell(numOrDash(metaAds.totalCampaigns)),
    paidMediaCell(numOrDash(metaAds.formLeads)),
    paidMediaCell(numOrDash(metaAds.messageConversations)),
    paidMediaCell(metaAds.totalSpend),
    paidMediaCell(numOrDash(metaAds.impressions)),
    paidMediaCell(numOrDash(metaAds.clicks)),
    paidMediaCell(numOrDash(metaAds.reach)),
  ];

  const rows = [['', ...header], dataRow];
  if (metaAds.gstAmount) {
    rows.push(['With GST', '', '', '', withGstText(metaAds.totalSpend, metaAds.gstAmount), '', '', '']);
  }

  addTable(slide, rows, [1.05, 1.15, 1.15, 1.05, 1.4, 1.15, 1.0, 1.05], { x: 0.5, y, w: 9 });
}

function metaAdsSummaryText(metaAds, monthStr) {
  return `Meta Ads delivered ${fmtNum(metaAds.impressions)} impressions and ${fmtNum(
    metaAds.clicks
  )} clicks in ${monthLabel(monthStr)}, with a total spend of ${metaAds.totalSpend} across ${
    metaAds.totalCampaigns
  } campaigns, generating ${fmtNum(metaAds.formLeads)} form leads and ${fmtNum(
    metaAds.messageConversations
  )} message leads.`;
}

function buildPaidMediaSlide(pptx, data) {
  const googleAds = (data.paidMedia && data.paidMedia.googleAds) || {};
  const metaAds = (data.paidMedia && data.paidMedia.metaAds) || {};
  const hasGoogle = hasAnyValue(googleAds);
  const hasMeta = hasAnyValue(metaAds);

  if (!hasGoogle && !hasMeta) return null;

  const slide = pptx.addSlide();
  slide.background = { color: BG };

  slide.addText('Paid Media Performance', {
    x: 0.5,
    y: 0.3,
    w: 9,
    h: 0.55,
    fontSize: 28,
    bold: true,
    color: ACCENT,
    fontFace: 'Arial',
  });
  slide.addText(`Google Ads & Meta Ads — ${monthLabel(data.currentMonth)}`, {
    x: 0.5,
    y: 0.85,
    w: 9,
    h: 0.3,
    fontSize: 13,
    color: NEUTRAL,
    fontFace: 'Arial',
  });
  slide.addShape('line', {
    x: 0.5,
    y: 1.2,
    w: 2.2,
    h: 0,
    line: { color: ACCENT, width: 2.5 },
  });

  let y = 1.45;

  addPaidMediaSectionHeader(slide, 'Google Ads', y);
  y += 0.45;
  if (hasGoogle) {
    buildGoogleAdsTable(slide, googleAds, y);
    y += googleAds.gstAmount ? 1.15 : 0.85;
    if (googleAds.cpc !== null && googleAds.cpc !== undefined) {
      slide.addText(googleAdsSummaryText(googleAds, data.currentMonth), {
        x: 0.5,
        y,
        w: 9,
        h: 0.4,
        fontSize: 10.5,
        color: TEXT_DARK,
        fontFace: 'Arial',
        valign: 'top',
      });
      y += 0.45;
    }
  } else {
    slide.addText('Not configured for this period', {
      x: 0.5,
      y,
      w: 9,
      h: 0.4,
      fontSize: 12,
      italic: true,
      color: NEUTRAL,
      fontFace: 'Arial',
    });
    y += 0.55;
  }

  y += 0.2;
  addPaidMediaSectionHeader(slide, 'Meta Ads', y);
  y += 0.45;
  if (hasMeta) {
    buildMetaAdsTable(slide, metaAds, y);
    y += metaAds.gstAmount ? 1.15 : 0.85;
    slide.addText(metaAdsSummaryText(metaAds, data.currentMonth), {
      x: 0.5,
      y,
      w: 9,
      h: 0.4,
      fontSize: 10.5,
      color: TEXT_DARK,
      fontFace: 'Arial',
      valign: 'top',
    });
  } else {
    slide.addText('Not configured for this period', {
      x: 0.5,
      y,
      w: 9,
      h: 0.4,
      fontSize: 12,
      italic: true,
      color: NEUTRAL,
      fontFace: 'Arial',
    });
  }

  slide.addText(`${data.clientName || BRAND} | Digital Marketing Report | ${monthLabel(data.currentMonth)}`, {
    x: 0.5,
    y: 5.35,
    w: 6.0,
    h: 0.25,
    fontSize: 9,
    color: NEUTRAL,
    fontFace: 'Arial',
  });
  slide.slideNumber = {
    x: 9.3,
    y: 5.35,
    w: 0.4,
    h: 0.25,
    fontSize: 9,
    color: NEUTRAL,
    fontFace: 'Arial',
  };

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
  buildPaidMediaSlide(pptx, data);

  return pptx.write({ outputType: 'nodebuffer' });
}

module.exports = { generateReportPptx, monthLabel };
