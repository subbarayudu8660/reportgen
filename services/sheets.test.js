const test = require('node:test');
const assert = require('node:assert/strict');
const { _internal } = require('./sheets');

const {
  normalizeLabel,
  labelContainsAny,
  parseFlexibleDate,
  parseMonthLabel,
  classifyRank,
  isKeywordTabName,
  isOffPageTabName,
  isKpiTrafficTabName,
  pickBestTab,
  findDateColumns,
  colsForMonth,
  latestColForMonth,
  findLabelRow,
  forwardFillRow,
  findMonthColumns,
  findKeywordHeaderRowIndex,
  getKpiTrafficAndAuthority,
} = _internal;

test('normalizeLabel trims, lowercases, and normalizes apostrophes', () => {
  assert.equal(normalizeLabel("  KPI'S  "), "kpi's");
  assert.equal(normalizeLabel('KPI’s'), "kpi's");
  assert.equal(normalizeLabel(null), '');
});

test('labelContainsAny matches case-insensitively regardless of surrounding text', () => {
  assert.equal(labelContainsAny('Organic Traffic (Sessions)', ['organic traffic']), true);
  assert.equal(labelContainsAny('ORGANIC', ['organic traffic', 'organic']), true);
  assert.equal(labelContainsAny('Direct', ['organic']), false);
});

test('parseFlexibleDate handles M/D/YYYY', () => {
  assert.deepEqual(parseFlexibleDate('6/1/2026'), { year: 2026, month: 6, day: 1 });
});

test('parseFlexibleDate handles D/M/YYYY when month-first is invalid', () => {
  assert.deepEqual(parseFlexibleDate('25/1/2026'), { year: 2026, month: 1, day: 25 });
});

test('parseFlexibleDate handles ISO dates', () => {
  assert.deepEqual(parseFlexibleDate('2026-06-01'), { year: 2026, month: 6, day: 1 });
});

test('parseFlexibleDate handles spelled-out month names', () => {
  assert.deepEqual(parseFlexibleDate('June 1, 2026'), { year: 2026, month: 6, day: 1 });
  assert.deepEqual(parseFlexibleDate('Jun 1 2026'), { year: 2026, month: 6, day: 1 });
  assert.deepEqual(parseFlexibleDate('1 June 2026'), { year: 2026, month: 6, day: 1 });
});

test('parseFlexibleDate handles Google Sheets serial numbers', () => {
  // 45813 => 2025-06-05 in the Sheets/Excel epoch (1899-12-30).
  assert.deepEqual(parseFlexibleDate('45813'), { year: 2025, month: 6, day: 5 });
});

test('parseFlexibleDate returns null for non-dates', () => {
  assert.equal(parseFlexibleDate('Organic Traffic'), null);
  assert.equal(parseFlexibleDate(''), null);
  assert.equal(parseFlexibleDate(null), null);
  assert.equal(parseFlexibleDate('42'), null); // too small to plausibly be a serial date
});

test('parseMonthLabel handles names, abbreviations, and numbers', () => {
  assert.deepEqual(parseMonthLabel('JAN'), { month: 1, year: null });
  assert.deepEqual(parseMonthLabel('January'), { month: 1, year: null });
  assert.deepEqual(parseMonthLabel('01'), { month: 1, year: null });
  assert.deepEqual(parseMonthLabel('1'), { month: 1, year: null });
  assert.equal(parseMonthLabel('Week 1'), null);
});

test('classifyRank buckets correctly and treats blank/100+ as pending', () => {
  assert.equal(classifyRank('5'), 'top10');
  assert.equal(classifyRank('10'), 'top10');
  assert.equal(classifyRank('11'), 'top11_30');
  assert.equal(classifyRank('30'), 'top11_30');
  assert.equal(classifyRank('31'), 'top31_50');
  assert.equal(classifyRank('45'), 'top31_50');
  assert.equal(classifyRank('75'), 'top51_100');
  assert.equal(classifyRank('100+'), 'pending');
  assert.equal(classifyRank(''), 'pending');
  assert.equal(classifyRank(null), 'pending');
  assert.equal(classifyRank('150'), 'pending');
});

test('tab name matching: keyword tab', () => {
  assert.equal(isKeywordTabName('Keyword Ranking Report'), true);
  assert.equal(isKeywordTabName('keywords'), true);
  assert.equal(isKeywordTabName('KEYWORDS OLD'), true);
  assert.equal(isKeywordTabName('KPI'), false);
});

test('tab name matching: off-page tab covers KPI_2 variants and submission wording', () => {
  assert.equal(isOffPageTabName('KPI_2'), true);
  assert.equal(isOffPageTabName('KPI 2'), true);
  assert.equal(isOffPageTabName('Off-Page Submissions'), true);
  assert.equal(isOffPageTabName('OffPage'), true);
  assert.equal(isOffPageTabName('Activity Tracker'), true);
  assert.equal(isOffPageTabName("KPI's"), false);
});

test('tab name matching: KPI traffic tab excludes the off-page KPI_2/KPI 2 tab', () => {
  assert.equal(isKpiTrafficTabName("KPI's"), true);
  assert.equal(isKpiTrafficTabName('KPI'), true);
  assert.equal(isKpiTrafficTabName('KPI_2'), false);
  assert.equal(isKpiTrafficTabName('KPI 2'), false);
});

test('pickBestTab prefers a non-"old"/"archive" tab when multiple match', () => {
  assert.equal(pickBestTab(['Keywords Old', 'Keywords']), 'Keywords');
  assert.equal(pickBestTab(['Keywords']), 'Keywords');
  assert.equal(pickBestTab([]), null);
});

test('findDateColumns / colsForMonth / latestColForMonth pick the most recent snapshot in a month', () => {
  const header = ['Keyword', '6/1/2026', '6/15/2026', '7/1/2026'];
  const dateCols = findDateColumns(header, 1);
  assert.equal(dateCols.length, 3);
  const juneCols = colsForMonth(dateCols, '2026-06');
  assert.equal(juneCols.length, 2);
  assert.equal(latestColForMonth(dateCols, '2026-06'), 2);
  assert.equal(latestColForMonth(dateCols, '2026-08'), null);
});

test('findLabelRow finds the first matching row and ignores others', () => {
  const rows = [['Direct'], ['Organic Traffic'], ['Organic Search']];
  assert.equal(findLabelRow(rows, ['organic traffic', 'organic search', 'organic']), 1);
  assert.equal(findLabelRow(rows, ['unassigned']), -1);
});

test('forwardFillRow fills blanks left over from merged header cells', () => {
  assert.deepEqual(forwardFillRow(['Metric', 'JAN', '', '', 'FEB', '', '', '']), [
    'Metric', 'JAN', 'JAN', 'JAN', 'FEB', 'FEB', 'FEB', 'FEB',
  ]);
});

test('findMonthColumns matches forward-filled merged month headers', () => {
  const header = ['Metric', 'JAN', '', '', 'FEB', '', '', ''];
  assert.deepEqual(findMonthColumns(header, '2026-01'), [1, 2, 3]);
  assert.deepEqual(findMonthColumns(header, '2026-02'), [4, 5, 6, 7]);
  assert.deepEqual(findMonthColumns(header, '2026-03'), []);
});

test('findKeywordHeaderRowIndex finds the header wherever it is, not a fixed row', () => {
  const values = [[], [''], ['Notes'], [], ['Keyword', '6/1/2026'], ['seo term', '5']];
  assert.equal(findKeywordHeaderRowIndex(values), 4);
});

test('findKeywordHeaderRowIndex returns -1 when no keyword header exists', () => {
  assert.equal(findKeywordHeaderRowIndex([['Foo'], ['Bar']]), -1);
});

test('getKpiTrafficAndAuthority scans labels/columns dynamically and warns on missing rows', () => {
  const values = [
    ['Metric', 'JAN', '', 'FEB', ''],
    ['', 'Week 1', 'Week 2', 'Week 1', 'Week 2'],
    ['Organic Traffic', '100', '150', '80', '90'],
    ['Direct', '10', '20', '5', '5'],
    ['DA', '45', '45', '46', '46'],
  ];
  const result = getKpiTrafficAndAuthority(values, '2026-01', '2026-02');
  assert.equal(result.trafficByChannel.organic.current, 250);
  assert.equal(result.trafficByChannel.organic.previous, 170);
  assert.equal(result.trafficByChannel.direct.current, 30);
  assert.equal(result.trafficByChannel.referral.current, null);
  assert.equal(result.domainAuthority.da.current, 45);
  assert.equal(result.domainAuthority.pa, null);
  assert.ok(result.warnings.some((w) => w.includes('Referral')) || result.warnings.some((w) => w.includes('referral')));
  assert.ok(result.warnings.some((w) => w.toLowerCase().includes('page authority')));
});

test('getKpiTrafficAndAuthority warns when no columns match the target month', () => {
  const values = [
    ['Metric', 'JAN'],
    ['', 'Week 1'],
    ['Organic Traffic', '100'],
  ];
  const result = getKpiTrafficAndAuthority(values, '2026-06', '2026-05');
  assert.ok(result.warnings.some((w) => w.includes('2026-06')));
});
