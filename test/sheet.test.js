import test from 'node:test';
import assert from 'node:assert/strict';
import vm from 'node:vm';
import fs from 'node:fs';
import { processWorkbook } from '../js/sheet.js';
const ctx = { console, TextEncoder, TextDecoder };
ctx.self = ctx; ctx.window = ctx;
vm.runInNewContext(fs.readFileSync(new URL('../vendor/xlsx.full.min.js', import.meta.url), 'utf8'), ctx);
const XLSX = ctx.XLSX;

function build(aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Loans');
  return XLSX.read(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }), { type: 'array', cellNF: true, cellStyles: true });
}
const H = ['Payment Date', 'Bal. restruc', 'Bank payment', 'Gross bank paym', 'new principal', 'Interest', 'Gross Loan', 'Monthly repayment', 'Start Date', 'End date', 'Status'];

test('calculates, updates existing columns, flags bad rows', () => {
  const wb = build([H,
    ['01-Sep-26', '', 144000, 150000, 150000, 90000, 240000, 20000, '1-Oct-2026', '30-Sep-2027', 'RENEWAL'],
    [],
    ['04-Sep-26', 165375, 144000, 150000, 315375, 0, 0, 0, '1-Oct-2026', '30-Sep-2027', 'TOP UP'],
    ['x', '', '', '', 'abc'], ['x', '', '', '', -5], ['x', '', '', '', '']]);
  const out = processWorkbook(XLSX, wb, { rate: 5, tenure: 12, autoTenure: false });
  assert.equal(out.rows.length, 2);
  assert.equal(out.invalid.length, 3);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets.Loans);
  assert.equal(rows[1]['Interest'], 189225);
  assert.equal(rows[1]['Gross Loan'], 504600);
  assert.equal(rows[1]['Monthly repayment'], 42050);
  assert.equal(rows[1]['Monthly Interest'], 15768.75);
  assert.equal(rows[1]['Monthly Flat Rate'], 0.05);
  assert.equal(rows[1]['Loan Tenure'], 12);
  assert.equal(Object.keys(rows[0]).filter((k) => /interest/i.test(k)).length, 2); // no duplicate Interest col
});
test('missing principal column / empty file', () => {
  assert.throws(() => processWorkbook(XLSX, build([['a', 'b'], [1, 2]]), { rate: 5, tenure: 12 }), /New Principal/);
});
test('auto tenure', () => {
  const wb = build([H, ['', '', '', '', 100000, '', '', '', '1-Oct-2026', '30-Sep-2027', '']]);
  const out = processWorkbook(XLSX, wb, { rate: 5, tenure: 12, autoTenure: true });
  assert.equal(out.rows[0].tenure, 12);
});

test('bank payment layout derives principal and fills placeholders', () => {
  const wb = build([
    ['s/n', 'Clients Name', 'Tenor', 'Balance B/F', 'Bank payment', 'Gross bank payment (column G/0.96)', 'Principal (column F + Column H)', 'Total Interest ', 'Total debt (Columns I + J)', 'monthly EMI'],
    [1, 'AUDU', 12, 0, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
    [2, 'OMOLORO', 12, 165375, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
    [3, 'BAD', 'x', 0, 100000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
  ]);
  const out = processWorkbook(XLSX, wb, { rate: 5, tenure: 12, deduction: 4, useFileTenor: true });
  assert.equal(out.rows.length, 2);
  assert.equal(out.invalid.length, 1);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets.Loans);
  assert.equal(rows[0]['Principal (column F + Column H)'], 150000);
  assert.equal(rows[0]['Total Interest '], 90000);
  assert.equal(rows[0]['monthly EMI'], 20000);
  assert.equal(rows[1]['Gross bank payment (column G/0.96)'], 150000);
  assert.equal(rows[1]['Principal (column F + Column H)'], 315375);
  assert.equal(rows[1]['Total debt (Columns I + J)'], 504600);
  assert.equal(rows[1]['monthly EMI'], 42050);
  assert.equal(Object.keys(rows[0]).filter((k) => /^tenor|tenure/i.test(k)).length, 1);
});
