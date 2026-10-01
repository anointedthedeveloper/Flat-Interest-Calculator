import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { patchSheetXml, patchXlsx } from '../js/patch.js';
import { processWorkbook } from '../js/sheet.js';
const ctx = { console, TextEncoder, TextDecoder, setTimeout, clearTimeout, setImmediate, Promise, Uint8Array, ArrayBuffer }; ctx.self = ctx; ctx.window = ctx;
vm.runInNewContext(fs.readFileSync(new URL('../vendor/xlsx.full.min.js', import.meta.url), 'utf8'), ctx);
vm.runInNewContext(fs.readFileSync(new URL('../vendor/jszip.min.js', import.meta.url), 'utf8'), ctx);
const XLSX = ctx.XLSX, JSZip = ctx.JSZip;

test('patchSheetXml changes only target cells and keeps styles', () => {
  const xml = '<worksheet><dimension ref="A1:C2"/><sheetData><row r="1" spans="1:3"><c r="A1" t="s"><v>0</v></c><c r="B1" s="4" t="s"><v>1</v></c></row><row r="2"><c r="A2" s="3"><v>5</v></c><c r="B2" s="7" t="s"><v>9</v></c></row></sheetData></worksheet>';
  const out = patchSheetXml(xml, [{ addr: 'B2', value: 150000 }, { addr: 'D2', value: 1.5 }, { addr: 'D1', text: 'New & Co' }]);
  assert.match(out, /<c r="B2" s="7"><v>150000<\/v><\/c><c r="D2"><v>1.5<\/v><\/c>/);
  assert.match(out, /<c r="A2" s="3"><v>5<\/v><\/c>/);
  assert.match(out, /<c r="B1" s="4" t="s"><v>1<\/v><\/c><c r="D1" t="inlineStr"><is><t>New &amp; Co<\/t><\/is><\/c>/);
  assert.match(out, /<dimension ref="A1:D2"\/>/);
});

const REAL = '/root/.claude/uploads/676f5c45-2ab5-5cb7-bc50-a53211555387/1541f6d4-emi_calculator_WORKINGS.xlsx';
test('real workbook: only the XXXX cells change, every other file in the package is identical', { skip: !fs.existsSync(REAL) }, async () => {
  const buf = fs.readFileSync(REAL);
  const wb = XLSX.read(buf, { type: 'buffer', cellNF: true, cellStyles: true });
  const out = processWorkbook(XLSX, wb, { rate: 5, tenure: 12, deduction: 4, useFileTenor: true });
  assert.equal(out.edits.length, 20); // 4 rows x 5 calculated columns
  const patched = await patchXlsx(JSZip, buf, out.sheetName, out.edits);
  const a = await JSZip.loadAsync(buf), b = await JSZip.loadAsync(patched);
  for (const name of Object.keys(a.files)) {
    if (a.files[name].dir || name === 'xl/worksheets/sheet1.xml') continue;
    assert.equal(await b.file(name).async('string'), await a.file(name).async('string'), name);
  }
  const rows = XLSX.utils.sheet_to_json(XLSX.read(patched, { type: 'buffer' }).Sheets.Sheet1);
  assert.equal(rows[1]['Principal (column H + Column J)'], 315375);
  assert.equal(rows[1]['monthly EMI'], 42050);
  assert.equal(rows[1]['Clients ID'], 451);
});

test('keep as is: nothing removed, empty and invalid rows untouched', async () => {
  const aoa = [['s/n', 'Name', 'Tenor', 'Balance B/F', 'Bank payment', 'Gross bank payment (column E/0.96)', 'Principal (column D + Column F)', 'Total Interest', 'Total debt', 'monthly EMI'],
    [1, 'A', 12, 0, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'], [], [2, 'BAD', 12, 0, 'abc', 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'], ['note', 'footer text']];
  const wb0 = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb0, XLSX.utils.aoa_to_sheet(aoa), 'Loans'); XLSX.utils.book_append_sheet(wb0, XLSX.utils.aoa_to_sheet([['other']]), 'Other');
  const buf = XLSX.write(wb0, { type: 'buffer', bookType: 'xlsx' });
  const wb = XLSX.read(buf, { type: 'buffer' });
  const out = processWorkbook(XLSX, wb, { rate: 5, tenure: 12, deduction: 4, useFileTenor: true });
  const patched = await patchXlsx(JSZip, buf, out.sheetName, out.edits);
  const before = XLSX.read(buf, { type: 'buffer' }), after = XLSX.read(patched, { type: 'buffer' });
  assert.deepEqual(after.SheetNames, before.SheetNames);
  assert.equal(after.Sheets.Loans['!ref'], before.Sheets.Loans['!ref']);
  const edited = new Set(out.edits.map((e) => e.addr));
  for (const addr of Object.keys(before.Sheets.Loans)) {
    if (addr[0] === '!' || edited.has(addr)) continue;
    assert.equal(after.Sheets.Loans[addr].v, before.Sheets.Loans[addr].v, addr);
  }
  assert.equal(after.Sheets.Loans.G2.v, 150000);
  assert.equal(after.Sheets.Loans.G4.v, 'XXXX'); // invalid row left as is
  assert.equal(after.Sheets.Other.A1.v, 'other');
});
