import { calculateLoan, parseAmount, monthsBetween } from './calc.js';

const norm = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
const MONEY_FMT = '#,##0.00';

// Header aliases (normalised). First match wins.
const isPrincipal = (h) => /^new?princ(i|a)p(a|l|e)l?(amount)?$/.test(h) || /^newprin/.test(h);
const FIND = {
  principal: (h) => isPrincipal(h),
  fallbackPrincipal: (h) => h === 'principal' || h === 'principle',
  interest: (h) => h === 'interest' || h === 'totalinterest',
  gross: (h) => h === 'grossloan' || h === 'totalrepayment' || h === 'grossloantotalrepayment',
  monthly: (h) => /^monthlyrepay/.test(h),
  rate: (h) => /^monthlyflatrate|^flatrate|^interestrate$/.test(h),
  monthlyInterest: (h) => h === 'monthlyinterest',
  tenure: (h) => h === 'loantenure' || h === 'tenure',
  start: (h) => h === 'startdate' || h === 'start',
  end: (h) => h === 'enddate' || h === 'end',
  status: (h) => h === 'status',
  name: (h) => /^(borrower|customer|client|member|staff)?(name|fullname)$/.test(h) || h === 'borrower' || h === 'customer',
};

function findHeaderRow(XLSX, ws, range) {
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 15); r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && FIND.principal(norm(cell.v))) return r;
    }
  }
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 15); r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && FIND.fallbackPrincipal(norm(cell.v))) return r;
    }
  }
  return -1;
}

function cellDateParts(XLSX, cell) {
  if (!cell || cell.v === undefined || cell.v === '') return null;
  if (typeof cell.v === 'number') {
    const p = XLSX.SSF.parse_date_code(cell.v); // serial → parts, no timezone involved
    return p ? { y: p.y, m: p.m, d: p.d } : null;
  }
  const t = new Date(String(cell.v) + ' UTC');
  return isNaN(t) ? null : { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

/**
 * Calculates every loan row of the first sheet that has a "New Principal" column and
 * writes results back into that SAME worksheet (original columns/data/formats preserved).
 * settings: { rate: number(%), tenure: number|null, autoTenure: boolean }
 */
export function processWorkbook(XLSX, wb, settings) {
  const sheetName = wb.SheetNames.find((n) => wb.Sheets[n]['!ref'] && findHeaderRow(XLSX, wb.Sheets[n], XLSX.utils.decode_range(wb.Sheets[n]['!ref'])) >= 0);
  if (!sheetName) {
    const first = wb.Sheets[wb.SheetNames[0]];
    if (!first || !first['!ref']) throw new Error('The Excel file is empty.');
    throw new Error('Could not find a "New Principal" column. Please make sure your sheet has a header named "New Principal".');
  }
  const ws = wb.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref']);
  const hr = findHeaderRow(XLSX, ws, range);

  const cols = {}; // key -> column index
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: hr, c })];
    const h = norm(cell?.v);
    headers[c] = cell?.v;
    if (!h) continue;
    for (const key of Object.keys(FIND)) {
      if (key === 'principal' || key === 'fallbackPrincipal') continue;
      if (cols[key] === undefined && FIND[key](h)) cols[key] = c;
    }
    if (cols.principal === undefined && FIND.principal(h)) cols.principal = c;
    else if (cols.fallback === undefined && FIND.fallbackPrincipal(h)) cols.fallback = c;
  }
  if (cols.principal === undefined) cols.principal = cols.fallback;

  // Add missing output columns at the end (existing columns are updated in place).
  let nextCol = range.e.c + 1;
  const addCol = (key, title) => {
    if (cols[key] !== undefined) return;
    cols[key] = nextCol++;
    ws[XLSX.utils.encode_cell({ r: hr, c: cols[key] })] = { t: 's', v: title };
  };
  addCol('rate', 'Monthly Flat Rate');
  addCol('monthlyInterest', 'Monthly Interest');
  addCol('tenure', 'Loan Tenure');
  addCol('interest', 'Interest');
  addCol('gross', 'Gross Loan');
  addCol('monthly', 'Monthly repayment');

  const setNum = (r, key, v, fmt) => {
    const addr = XLSX.utils.encode_cell({ r, c: cols[key] });
    const old = ws[addr];
    const z = old && old.z && old.z !== 'General' ? old.z : fmt;
    ws[addr] = { ...(old || {}), t: 'n', v, z };
    delete ws[addr].w; delete ws[addr].f;
  };

  const rows = [], invalid = [];
  let detected = 0;
  for (let r = hr + 1; r <= range.e.r; r++) {
    let hasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined && String(cell.v).trim() !== '') { hasData = true; break; }
    }
    if (!hasData) continue; // empty row
    detected++;
    const cellAt = (key) => (cols[key] === undefined ? undefined : ws[XLSX.utils.encode_cell({ r, c: cols[key] })]);
    const nameCell = cellAt('name');
    const label = nameCell && nameCell.v ? String(nameCell.v) : `Row ${r + 1}`;
    const statusCell = cellAt('status');
    const status = statusCell && statusCell.w ? statusCell.w : statusCell?.v ?? '';
    const fail = (reason) => invalid.push({ row: r + 1, label, reason });

    const pc = cellAt('principal');
    if (!pc || pc.v === undefined || String(pc.v).trim() === '') { fail('New Principal is empty.'); continue; }
    const principal = parseAmount(pc.v);
    if (Number.isNaN(principal)) { fail(`New Principal "${pc.v}" is not a valid number.`); continue; }
    if (principal < 0) { fail('New Principal is negative.'); continue; }
    if (principal === 0) { fail('New Principal is zero.'); continue; }

    let tenure = settings.tenure;
    if (settings.autoTenure) {
      const s = cellDateParts(XLSX, cellAt('start')), e = cellDateParts(XLSX, cellAt('end'));
      tenure = s && e ? monthsBetween(s, e) : NaN;
      if (!(tenure >= 1)) { fail('Auto tenure needs a valid Start Date and End Date (end after start).'); continue; }
    }

    const res = calculateLoan(principal, settings.rate, tenure);
    setNum(r, 'principal', res.principal, pc.z && pc.z !== 'General' ? pc.z : MONEY_FMT);
    setNum(r, 'rate', res.rate / 100, '0.00%');
    setNum(r, 'monthlyInterest', res.monthlyInterest, MONEY_FMT);
    setNum(r, 'tenure', res.tenure, '0');
    setNum(r, 'interest', res.totalInterest, MONEY_FMT);
    setNum(r, 'gross', res.totalRepayment, MONEY_FMT);
    setNum(r, 'monthly', res.monthlyRepayment, MONEY_FMT);
    rows.push({ row: r + 1, label, status, ...res });
  }

  ws['!ref'] = XLSX.utils.encode_range({ s: range.s, e: { r: range.e.r, c: Math.max(range.e.c, nextCol - 1) } });
  if (ws['!cols']) for (let c = range.e.c + 1; c < nextCol; c++) ws['!cols'][c] = { wch: 18 };
  return { sheetName, detected, rows, invalid, usedFallbackPrincipal: cols.principal === cols.fallback && cols.fallback !== undefined };
}

export function summarize(rows) {
  const sum = (k) => Math.round(rows.reduce((a, r) => a + r[k], 0) * 100) / 100;
  return {
    count: rows.length,
    principal: sum('principal'), monthlyInterest: sum('monthlyInterest'),
    totalInterest: sum('totalInterest'), totalRepayment: sum('totalRepayment'),
    avgMonthly: rows.length ? Math.round((sum('monthlyRepayment') / rows.length) * 100) / 100 : 0,
  };
}
