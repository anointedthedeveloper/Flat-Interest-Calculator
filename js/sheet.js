import { calculateLoan, parseAmount, monthsBetween, grossFromBank } from './calc.js';

const norm = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
const MONEY_FMT = '#,##0.00';

// Header aliases (normalised: lowercase, letters and digits only).
const FIND = {
  principal: (h) => /^newprin/.test(h),
  fallbackPrincipal: (h) => /^princip/.test(h),
  bank: (h) => /^bankpay/.test(h),
  grossBank: (h) => /^grossbankpay/.test(h),
  balance: (h) => /^(balance|bal)(bf|broughtforward|restruc\w*)?$/.test(h),
  interest: (h) => h === 'interest' || h === 'totalinterest',
  gross: (h) => h === 'grossloan' || h === 'totalrepayment' || /^totaldebt/.test(h) || h === 'grossloantotalrepayment',
  monthly: (h) => /^monthlyrepay/.test(h) || /^monthlyemi/.test(h) || h === 'emi',
  rate: (h) => /^monthlyflatrate|^flatrate|^interestrate$/.test(h),
  monthlyInterest: (h) => h === 'monthlyinterest',
  tenure: (h) => h === 'loantenure' || h === 'tenure' || h === 'tenor',
  start: (h) => h === 'startdate' || h === 'start',
  end: (h) => h === 'enddate' || h === 'end',
  status: (h) => h === 'status',
  name: (h) => /^(borrower|customer|client|clients|member|staff)?(name|fullname)$/.test(h) || h === 'borrower' || h === 'customer',
};
const isAnchor = (h) => FIND.principal(h) || FIND.fallbackPrincipal(h) || FIND.bank(h);

function findHeaderRow(XLSX, ws, range) {
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 15); r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === 'string' && isAnchor(norm(cell.v))) return r;
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
    throw new Error('Could not find a "New Principal" or "Bank payment" column. Please make sure your sheet has a header row containing one of them.');
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
    if (FIND.principal(h)) { if (cols.principal === undefined) cols.principal = c; continue; }
    if (FIND.fallbackPrincipal(h)) { if (cols.fallback === undefined) cols.fallback = c; continue; }
    for (const key of Object.keys(FIND)) {
      if (key === 'principal' || key === 'fallbackPrincipal') continue;
      if (cols[key] === undefined && FIND[key](h)) { cols[key] = c; break; }
    }
  }
  if (cols.principal === undefined) cols.principal = cols.fallback;
  if (cols.principal === undefined && cols.bank === undefined) throw new Error('Could not find a "New Principal" or "Bank payment" column.');
  const hasTenorColumn = cols.tenure !== undefined;
  const deduction = settings.deduction ?? 4;
  // Principal column may be missing when it is derived from Bank payment: create it.

  // Add missing output columns at the end (existing columns are updated in place).
  let nextCol = range.e.c + 1;
  const dry = !!settings.dryRun; // dry run: analyse only, never touch the sheet
  const addCol = (key, title) => {
    if (cols[key] !== undefined) return;
    cols[key] = nextCol++;
    if (!dry) ws[XLSX.utils.encode_cell({ r: hr, c: cols[key] })] = { t: 's', v: title };
  };
  addCol('principal', 'Principal');
  addCol('rate', 'Monthly Flat Rate');
  addCol('monthlyInterest', 'Monthly Interest');
  addCol('tenure', 'Loan Tenure');
  addCol('interest', 'Interest');
  addCol('gross', 'Gross Loan');
  addCol('monthly', 'Monthly repayment');

  const setNum = (r, key, v, fmt) => {
    if (dry) return;
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

    // Principal: use the sheet's value; if it is blank/placeholder (e.g. "XXXX") derive it from
    // Balance B/F + Bank payment / (1 - deduction).
    const pc = cellAt('principal');
    const pEmpty = !pc || pc.v === undefined || String(pc.v).trim() === '';
    let principal = pEmpty ? NaN : parseAmount(pc.v);
    let derived = false, gross = null;
    if (Number.isNaN(principal) && cols.bank !== undefined) {
      const bc = cellAt('bank');
      const bank = bc ? parseAmount(bc.v) : NaN;
      if (Number.isFinite(bank)) {
        if (bank < 0) { fail('Bank payment is negative.'); continue; }
        const bf = cellAt('balance');
        const bal = !bf || bf.v === undefined || String(bf.v).trim() === '' ? 0 : parseAmount(bf.v);
        if (Number.isNaN(bal)) { fail(`Balance B/F "${bf.v}" is not a valid number.`); continue; }
        gross = grossFromBank(bank, deduction);
        principal = Math.round((bal + gross + Number.EPSILON) * 100) / 100;
        derived = true;
      }
    }
    if (!derived) {
      if (pEmpty) { fail(cols.bank !== undefined ? 'No principal and no Bank payment to work it out from.' : 'New Principal is empty.'); continue; }
      if (Number.isNaN(principal)) { fail(`New Principal "${pc.v}" is not a valid number.`); continue; }
    }
    if (principal < 0) { fail('Principal is negative.'); continue; }
    if (principal === 0) { fail('Principal is zero.'); continue; }

    let tenure = settings.tenure;
    if (settings.useFileTenor && hasTenorColumn) {
      const tc = cellAt('tenure');
      if (tc && tc.v !== undefined && String(tc.v).trim() !== '') {
        tenure = parseAmount(tc.v);
        if (!Number.isInteger(tenure) || tenure < 1) { fail(`Tenor "${tc.v}" is not a valid number of months.`); continue; }
      }
    }
    if (settings.autoTenure) {
      const s = cellDateParts(XLSX, cellAt('start')), e = cellDateParts(XLSX, cellAt('end'));
      tenure = s && e ? monthsBetween(s, e) : NaN;
      if (!(tenure >= 1)) { fail('Auto tenure needs a valid Start Date and End Date (end after start).'); continue; }
    }

    const res = calculateLoan(principal, settings.rate, tenure);
    if (derived && cols.grossBank !== undefined) setNum(r, 'grossBank', gross, MONEY_FMT);
    setNum(r, 'principal', res.principal, MONEY_FMT);
    setNum(r, 'rate', res.rate / 100, '0.00%');
    setNum(r, 'monthlyInterest', res.monthlyInterest, MONEY_FMT);
    setNum(r, 'tenure', res.tenure, '0');
    setNum(r, 'interest', res.totalInterest, MONEY_FMT);
    setNum(r, 'gross', res.totalRepayment, MONEY_FMT);
    setNum(r, 'monthly', res.monthlyRepayment, MONEY_FMT);
    rows.push({ row: r + 1, label, status, derived, gross, ...res });
  }

  if (!dry) ws['!ref'] = XLSX.utils.encode_range({ s: range.s, e: { r: range.e.r, c: Math.max(range.e.c, nextCol - 1) } });
  if (!dry && ws['!cols']) for (let c = range.e.c + 1; c < nextCol; c++) ws['!cols'][c] = { wch: 18 };
  return { sheetName, detected, rows, invalid, hasTenorColumn, derivedCount: rows.filter((r) => r.derived).length };
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

/** First rows of the detected sheet as display text, for the pre-calculation preview. */
export function previewSheet(XLSX, wb, limit = 8) {
  const name = wb.SheetNames.find((n) => wb.Sheets[n]['!ref']) || wb.SheetNames[0];
  const ws = wb.Sheets[name];
  if (!ws || !ws['!ref']) return { headers: [], rows: [], total: 0 };
  const range = XLSX.utils.decode_range(ws['!ref']);
  const hr = Math.max(findHeaderRow(XLSX, ws, range), range.s.r);
  const text = (r, c) => { const x = ws[XLSX.utils.encode_cell({ r, c })]; return x ? String(x.w ?? x.v ?? '') : ''; };
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) headers.push(text(hr, c));
  const rows = [];
  for (let r = hr + 1; r <= range.e.r; r++) {
    const row = headers.map((_, i) => text(r, range.s.c + i));
    if (row.some((v) => v.trim() !== '')) rows.push(row);
  }
  return { headers, rows: rows.slice(0, limit), total: rows.length };
}
