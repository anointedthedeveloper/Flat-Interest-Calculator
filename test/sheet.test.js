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
const S = { rate: 5, deduction: 4 };
const H = ['Tenor', 'Bank payment', 'new principal', 'Interest', 'Gross Loan', 'Monthly repayment', 'Status'];

test('New Principal layout: updates existing columns, flags bad rows, no extra columns', () => {
  const wb = build([H,
    [12, 144000, 150000, 90000, 240000, 20000, 'RENEWAL'],
    [],
    [12, 144000, 315375, 0, 0, 0, 'TOP UP'],
    [12, '', 'abc'], [12, '', -5], [null, null, null, null, null, null, 'note']]);
  const out = processWorkbook(XLSX, wb, S);
  assert.equal(out.rows.length, 2);
  assert.equal(out.invalid.length, 2); // the text-only row is a note, not a loan
  const rows = XLSX.utils.sheet_to_json(wb.Sheets.Loans);
  assert.equal(rows[1]['Interest'], 189225);
  assert.equal(rows[1]['Gross Loan'], 504600);
  assert.equal(rows[1]['Monthly repayment'], 42050);
  assert.deepEqual(Object.keys(rows[0]), H); // no extra columns
});
test('missing principal/bank column', () => {
  assert.throws(() => processWorkbook(XLSX, build([['a', 'b'], [1, 2]]), S), /New Principal/);
});
test('missing Tenor column asks the user to add it', () => {
  assert.throws(() => processWorkbook(XLSX, build([['Bank payment'], [1000]]), S), /Tenor/);
});
test('blank or invalid tenor is reported, never assumed', () => {
  const wb = build([['Tenor', 'Bank payment', 'Interest'], [null, 96000], ['abc', 96000], [0, 96000], [12, 96000]]);
  const out = processWorkbook(XLSX, wb, S);
  assert.equal(out.rows.length, 1);
  assert.equal(out.invalid.length, 3);
  assert.match(out.invalid[0].reason, /Tenor is empty/);
});

test('bank payment layout derives principal and fills placeholders', () => {
  const wb = build([
    ['s/n', 'Clients Name', 'Tenor', 'Balance B/F', 'Bank payment', 'Gross bank payment (column G/0.96)', 'Principal (column F + Column H)', 'Total Interest ', 'Total debt (Columns I + J)', 'monthly EMI'],
    [1, 'AUDU', 12, 0, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
    [2, 'OMOLORO', 12, 165375, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
    [3, 'BAD', 'x', 0, 100000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
  ]);
  const out = processWorkbook(XLSX, wb, S);
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
});

test('output columns equal input columns', () => {
  const wb = build([
    ['s/n', 'Clients ID', 'Clients Name', 'Tenor', 'Balance B/F', 'Bank payment', 'Gross bank payment (column I/0.96)', 'Principal (column H + Column J)', 'Total Interest ', 'Total debt (Columns K + L)', 'monthly EMI'],
    [1, 156, 'AUDU', 12, 0, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX']]);
  const before = XLSX.utils.sheet_to_json(wb.Sheets.Loans, { header: 1 })[0];
  processWorkbook(XLSX, wb, S);
  assert.deepEqual(XLSX.utils.sheet_to_json(wb.Sheets.Loans, { header: 1 })[0], before);
  assert.equal(XLSX.utils.sheet_to_json(wb.Sheets.Loans)[0]['Clients ID'], 156);
});

test('building layout: blank cells, different tenors, note rows skipped', () => {
  const wb = build([
    ['S/N', 'Clients Name', 'Tenor', 'Balance B/Fwd', 'Bank payment', 'Gross Payment (Column I/.96)', 'Principal (H+J)', 'Interest', 'Gross Loan (K+L)', 'EMI', 'Status'],
    [1, 'OKOH', 12, 36012.38, 96000, null, null, null, null, null, 'TOP UP'],
    [2, 'ABURU', 6, null, 240000, null, null, null, null, null, 'NEW'],
    [null, 'SCHOOL CUSTOMER'],
    [3, 'EWARAMI', 18, null, 288000, 300000, 300000, 180000, 480000, 40000, 'NEW'],
  ]);
  const out = processWorkbook(XLSX, wb, S);
  assert.equal(out.skipped, 1);
  assert.equal(out.rows.length, 3);
  const r = XLSX.utils.sheet_to_json(wb.Sheets.Loans);
  assert.equal(r[0]['Gross Payment (Column I/.96)'], 100000);
  assert.equal(r[0]['Principal (H+J)'], 136012.38);
  assert.equal(r[0]['Interest'], 81607.43);
  assert.equal(r[0]['EMI'], 18134.98);
  assert.equal(r[1]['Interest'], 75000); // 6 months
  assert.equal(r[1]['EMI'], 54166.67);
  assert.equal(r[3]['Gross Loan (K+L)'], 570000); // 18 months
  assert.equal(r[3]['EMI'], 31666.67);
});
