// format.js: number formatting. Money is held in GBP thousands throughout the workbook.
export const isNum = v => typeof v === 'number' && Number.isFinite(v);
const grp = (n, dp) => n.toLocaleString('en-GB', { minimumFractionDigits: dp, maximumFractionDigits: dp });

export function fmtK(v, dp = 0) {                 // GBP thousands, accounting style
  if (!isNum(v)) return v === '' || v === null || v === undefined ? '' : String(v);
  const t = grp(Math.abs(v), dp);
  return v < 0 ? `(${t})` : t;
}
export function fmtMoney(v) {                      // headline style: GBP m above 1,000k
  if (!isNum(v)) return String(v ?? '');
  const a = Math.abs(v), sign = v < 0 ? '-' : '';
  return a >= 1000 ? `${sign}£${grp(a / 1000, 1)}m` : `${sign}£${grp(a, 0)}k`;
}
export const fmtPct = (v, dp = 1) => isNum(v) ? `${grp(v * 100, dp)}%` : String(v ?? '');
export const fmtX = (v, dp = 1) => isNum(v) ? `${grp(v, dp)}x` : String(v ?? '');
export const fmtMo = (v) => isNum(v) ? `${grp(v, 1)} months` : String(v ?? '');
export const fmtNum = (v, dp = 1) => isNum(v) ? grp(v, dp) : String(v ?? '');

/** Format a workbook value by the unit text stored beside it on Outputs_Web. */
export function fmtByUnit(v, unit) {
  if (!isNum(v)) return v === '' || v === null || v === undefined ? '' : String(v);
  switch (unit) {
    case '%': case '%/mo': return fmtPct(v);
    case 'x': return fmtX(v);
    case 'months': return fmtNum(v, 1);
    case 'GBP k': return fmtK(v, 0);
    case '#': case 'FTE': return fmtNum(v, 1);
    default: return fmtNum(v, 1);
  }
}
export function inputDisplay(v, unit) {            // what the number box shows for an input
  if (/%|pts/.test(unit || '')) return Math.round(v * 1e6) / 1e4;   // 0.065 -> 6.5
  return Math.round(v * 1e6) / 1e6;
}
export function inputParse(text, unit) {
  const n = Number(text);
  return /%|pts/.test(unit || '') ? n / 100 : n;
}
export const isPercentUnit = unit => /%|pts/.test(unit || '');
export function excelDate(serial) {
  const d = new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
  return d.toLocaleDateString('en-GB', { month: 'short', year: '2-digit', timeZone: 'UTC' });
}
