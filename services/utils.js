// Percent change from `previous` to `current`. When `previous` is 0 and `current`
// is nonzero, the "% increase" is mathematically undefined (not a real 100%) —
// returns null so callers/UI can render "New" instead of a misleading number.
function pctChange(current, previous) {
  if (!previous) return current > 0 ? null : 0;
  return ((current - previous) / previous) * 100;
}

// Computes the first/last calendar day of a "YYYY-MM" month as YYYY-MM-DD strings.
function monthDateRange(monthStr) {
  const [year, month] = monthStr.split('-').map(Number);
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const endDateObj = new Date(year, month, 0); // last day of month
  const endDate = `${endDateObj.getFullYear()}-${String(endDateObj.getMonth() + 1).padStart(2, '0')}-${String(endDateObj.getDate()).padStart(2, '0')}`;
  return { startDate, endDate };
}

module.exports = { pctChange, monthDateRange };
