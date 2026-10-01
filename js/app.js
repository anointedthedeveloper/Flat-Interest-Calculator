import { validateSettings, formatNaira } from './calc.js';
import { processWorkbook, summarize, previewSheet } from './sheet.js';
import { patchXlsx } from './patch.js';

const $ = (id) => document.getElementById(id);
const show = (el, on = true) => { el.hidden = !on; };
const fail = (el, msg) => { el.textContent = msg; show(el, !!msg); };
const esc = (t) => String(t ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const nextPaint = () => new Promise((res) => requestAnimationFrame(() => setTimeout(res, 30)));
const wait = (ms) => new Promise((res) => setTimeout(res, ms));

let buffer = null, lastWb = null, lastOut = null, hasFile = false, isXlsx = true;
let results = [], filtered = [], page = 0, showGross = false, showStatus = false;

const readWorkbook = (buf) => XLSX.read(buf, { type: 'array', cellNF: true, cellStyles: true, cellDates: false });

/* ---------- drop zone states: idle | over | reading | done | error ---------- */
const drop = $('drop');
let settled = 'idle'; // state to return to when a drag leaves
function setDrop(state, { name = '', meta = '', msg = '' } = {}) {
  drop.dataset.state = state;
  drop.querySelectorAll('.fname').forEach((el) => { if (name) el.textContent = name; });
  if (meta) drop.querySelector('.fmeta').textContent = meta;
  if (msg) drop.querySelector('.errmsg').textContent = msg;
  if (state !== 'over') settled = state;
  if (state === 'done') { // restart the check-mark animation
    const t = drop.querySelector('.tick'); t.replaceWith(t.cloneNode(true));
  }
}
const openPicker = () => $('file').click();
drop.querySelectorAll('.choose').forEach((b) => b.addEventListener('click', (e) => { e.stopPropagation(); openPicker(); }));
drop.addEventListener('click', () => { if (drop.dataset.state === 'idle' || drop.dataset.state === 'error') openPicker(); });
drop.addEventListener('keydown', (e) => { if ((e.key === 'Enter' || e.key === ' ') && e.target === drop) { e.preventDefault(); openPicker(); } });
$('file').addEventListener('change', (e) => { const f = e.target.files[0]; e.target.value = ''; if (f) loadFile(f); });

let depth = 0; // dragenter/dragleave fire for child elements too
const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
drop.addEventListener('dragenter', (e) => { if (!hasFiles(e)) return; e.preventDefault(); depth++; if (drop.dataset.state !== 'reading') setDrop('over'); });
drop.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
drop.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (depth === 0 && drop.dataset.state === 'over') drop.dataset.state = settled; });
drop.addEventListener('drop', (e) => {
  e.preventDefault(); depth = 0;
  const f = e.dataTransfer.files[0];
  if (f) loadFile(f); else drop.dataset.state = settled;
});
// A file dropped outside the box should not make the browser navigate away.
['dragover', 'drop'].forEach((t) => window.addEventListener(t, (e) => e.preventDefault()));

const fmtSize = (n) => (n < 1048576 ? Math.max(1, Math.round(n / 1024)) + ' KB' : (n / 1048576).toFixed(1) + ' MB');

async function loadFile(f) {
  fail($('error'), ''); fail($('settingsError'), '');
  show($('results'), false); show($('preview'), false);
  hasFile = false; $('calc').disabled = true;
  if (!/\.(xlsx|xls)$/i.test(f.name)) { setDrop('error', { name: f.name, msg: 'Unsupported file type. Please upload an .xlsx or .xls file.' }); return; }
  isXlsx = /\.xlsx$/i.test(f.name);
  setDrop('reading', { name: f.name });
  await nextPaint();
  try {
    const t0 = performance.now();
    buffer = await f.arrayBuffer();
    const wb = readWorkbook(buffer);
    const pv = previewSheet(XLSX, wb);
    const probe = processWorkbook(XLSX, wb, { rate: 5, tenure: 12, deduction: 4, dryRun: true });
    const spent = performance.now() - t0;
    if (spent < 700) await wait(700 - spent); // let the drop animation be seen on small files
    hasFile = true;
    const n = probe.detected;
    setDrop('done', { name: f.name, meta: `${fmtSize(f.size)}, ${n.toLocaleString()} loan record${n === 1 ? '' : 's'} detected in "${probe.sheetName}"` });
    setupSettings(probe);
    $('previewTable').tHead.innerHTML = '<tr>' + pv.headers.map((h) => `<th>${esc(h)}</th>`).join('') + '</tr>';
    $('previewTable').tBodies[0].innerHTML = pv.rows.map((r) => '<tr>' + r.map((v) => `<td>${esc(v)}</td>`).join('') + '</tr>').join('');
    $('previewNote').textContent = `first ${pv.rows.length} of ${pv.total.toLocaleString()} rows`;
    show($('preview'));
    $('calc').disabled = false;
  } catch (e) {
    buffer = null;
    setDrop('error', { name: f.name, msg: e.message || 'Could not read this Excel file.' });
  }
}

/* ---------- settings ---------- */
function setupSettings(probe) {
  const opts = [['fixed', 'The number above']];
  if (probe.hasTenorColumn) opts.unshift(['column', 'Tenor column in the file']);
  if (probe.hasDates) opts.push(['dates', 'Start Date and End date']);
  $('tsrc').innerHTML = opts.map(([v, t]) => `<option value="${v}">${t}</option>`).join('');
  show($('srcRow'), opts.length > 1);
  show($('dedRow'), probe.usesBank);
  syncTenure();
}
const syncTenure = () => { $('tenure').disabled = $('tsrc').value === 'dates'; };
$('tsrc').onchange = syncTenure;

$('calc').onclick = async () => {
  fail($('settingsError'), ''); fail($('error'), '');
  const rate = parseFloat($('rate').value), tenure = Number($('tenure').value), auto = $('tsrc').value === 'dates';
  const deduction = $('dedRow').hidden ? 4 : parseFloat($('deduction').value);
  const useFileTenor = $('tsrc').value === 'column';
  const err = validateSettings(rate, tenure, auto, deduction);
  if (err) { fail($('settingsError'), err); return; }
  const btn = $('calc'); btn.disabled = true;
  const label = btn.innerHTML; btn.textContent = 'Calculating...';
  await nextPaint();
  try {
    lastWb = readWorkbook(buffer); // fresh copy every run
    const out = processWorkbook(XLSX, lastWb, { rate, tenure, autoTenure: auto, deduction, useFileTenor });
    lastOut = out;
    render(out);
    show($('xlsNote'), !isXlsx);
  } catch (e) { fail($('error'), e.message); show($('results'), false); }
  btn.innerHTML = label; btn.disabled = false;
};

/* ---------- results ---------- */
function render(out) {
  const s = summarize(out.rows);
  const stat = (l, v, i) => `<div class="stat" style="animation-delay:${i * 50}ms"><span>${l}</span><b>${v}</b></div>`;
  $('cards').innerHTML = [['Total loans', s.count.toLocaleString()], ['Total principal', formatNaira(s.principal)],
    ['Total monthly interest', formatNaira(s.monthlyInterest)], ['Total interest', formatNaira(s.totalInterest)],
    ['Total repayment', formatNaira(s.totalRepayment)], ['Average monthly repayment', formatNaira(s.avgMonthly)]]
    .map(([l, v], i) => stat(l, v, i)).join('');

  const inv = $('invalid'), LIMIT = 200;
  if (out.invalid.length) {
    inv.innerHTML = `<strong>${out.invalid.length.toLocaleString()} row${out.invalid.length === 1 ? '' : 's'} could not be calculated (left unchanged in the file):</strong><ul>` +
      out.invalid.slice(0, LIMIT).map((r) => `<li>Row ${r.row}${r.label.startsWith('Row ') ? '' : ' (' + esc(r.label) + ')'}: ${esc(r.reason)}</li>`).join('') +
      (out.invalid.length > LIMIT ? `<li>and ${(out.invalid.length - LIMIT).toLocaleString()} more</li>` : '') + '</ul>';
    show(inv);
  } else show(inv, false);

  results = out.rows; showGross = out.derivedCount > 0; showStatus = out.rows.some((r) => r.status);
  $('thead').innerHTML = ['Borrower / Row', showGross && 'Gross Bank Payment', 'Principal', 'Monthly Interest', 'Total Interest', 'Total Repayment', 'Monthly Repayment', 'Tenure', showStatus && 'Status']
    .filter(Boolean).map((h, i) => `<th class="${i > 0 && h !== 'Status' ? 'r' : ''}">${h}</th>`).join('');
  $('q').value = ''; page = 0; applyFilter();
  show($('results'));
  $('results').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function applyFilter() {
  const q = $('q').value.trim().toLowerCase();
  filtered = q ? results.filter((r) => `${r.label} row ${r.row} ${r.status}`.toLowerCase().includes(q)) : results;
  page = Math.min(page, Math.max(0, Math.ceil(filtered.length / pageSize()) - 1));
  drawPage();
}
const pageSize = () => Number($('pageSize').value);
function drawPage() {
  const size = pageSize(), start = page * size, slice = filtered.slice(start, start + size);
  $('table').tBodies[0].innerHTML = slice.map((r) => {
    const money = (v) => `<td class="r">${v == null ? '' : formatNaira(v)}</td>`;
    return `<tr><td>${esc(r.label)}</td>${showGross ? money(r.gross) : ''}${[r.principal, r.monthlyInterest, r.totalInterest, r.totalRepayment, r.monthlyRepayment].map(money).join('')}` +
      `<td class="r">${r.tenure} mo</td>${showStatus ? `<td>${esc(r.status)}</td>` : ''}</tr>`;
  }).join('') || `<tr><td colspan="10" class="muted">No matching rows.</td></tr>`;
  const pages = Math.max(1, Math.ceil(filtered.length / size));
  $('range').textContent = filtered.length ? `Showing ${(start + 1).toLocaleString()}-${(start + slice.length).toLocaleString()} of ${filtered.length.toLocaleString()}` : '0 rows';
  $('pageNo').textContent = `Page ${page + 1} of ${pages}`;
  $('prev').disabled = page === 0; $('next').disabled = page >= pages - 1;
}
$('q').oninput = () => { page = 0; applyFilter(); };
$('pageSize').onchange = () => { page = 0; drawPage(); };
$('prev').onclick = () => { page--; drawPage(); };
$('next').onclick = () => { page++; drawPage(); };

// Every sheet of the original must still exist, with at least the same used range.
function integrityProblem(bytes) {
  const o = XLSX.read(buffer, { type: 'array', sheetStubs: true }), n = XLSX.read(bytes, { type: 'array', sheetStubs: true });
  if (o.SheetNames.join('|') !== n.SheetNames.join('|')) return 'Sheet list changed';
  for (const name of o.SheetNames) {
    const a = o.Sheets[name]['!ref'], b = n.Sheets[name]['!ref'];
    if (!a) continue;
    if (!b) return `Sheet "${name}" lost its data`;
    const ra = XLSX.utils.decode_range(a), rb = XLSX.utils.decode_range(b);
    if (rb.e.r < ra.e.r || rb.e.c < ra.e.c) return `Sheet "${name}" lost rows or columns`;
  }
  return null;
}
function save(blobParts, name, type) {
  const url = URL.createObjectURL(new Blob(blobParts, { type }));
  const a = Object.assign(document.createElement('a'), { href: url, download: name });
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
$('download').onclick = async () => {
  const d = new Date(), p = (n) => String(n).padStart(2, '0');
  const name = `loan_calculated_${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.xlsx`;
  const xlsxType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  fail($('error'), ''); show($('dlNote'), false);
  if (isXlsx) {
    try { // .xlsx: change only the calculated cells inside the original file, so all formatting survives
      const bytes = await patchXlsx(JSZip, buffer, lastOut.sheetName, lastOut.edits);
      const problem = integrityProblem(bytes);
      if (problem) throw new Error(problem);
      save([bytes], name, xlsxType);
      const sheets = XLSX.read(buffer, { type: 'array', bookSheets: true }).SheetNames.length;
      $('dlNote').textContent = `Checked: all ${sheets} sheet${sheets === 1 ? '' : 's'}, rows and columns were kept. ${lastOut.edits.length.toLocaleString()} cell${lastOut.edits.length === 1 ? '' : 's'} changed.`;
      show($('dlNote'));
      return;
    } catch (e) { console.warn('In-place save failed, using standard export', e); }
  }
  XLSX.writeFile(lastWb, name, { bookType: 'xlsx', cellStyles: true });
};

/* ---------- sample files ---------- */
const downloadSample = (name, aoa, widths) => {
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  ws['!cols'] = widths.map((w) => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Loans');
  XLSX.writeFile(wb, name);
};
const D = (y, m, d) => new Date(Date.UTC(y, m - 1, d));
$('sample').onclick = () => downloadSample('sample_new_principal.xlsx', [
  ['Payment Date', 'Bal. restruc', 'Bank payment', 'Gross bank paym', 'New Principal', 'Interest', 'Gross Loan', 'Monthly repayment', 'Start Date', 'End date', 'Status'],
  [D(2026, 9, 1), null, 144000, 150000, 150000, null, null, null, D(2026, 10, 1), D(2027, 9, 30), 'RENEWAL'],
  [D(2026, 9, 4), 165375, 144000, 150000, 315375, null, null, null, D(2026, 10, 1), D(2027, 9, 30), 'TOP UP'],
], Array(11).fill(16));
$('sample2').onclick = () => downloadSample('sample_bank_payment.xlsx', [
  ['s/n', 'Clients Name', 'IPPIS NO', 'Ministry', 'Tenor', 'Payment Date', 'Balance B/F', 'Bank payment', 'Gross bank payment (Bank payment / 0.96)', 'Principal (Balance B/F + Gross)', 'Total Interest', 'Total debt', 'monthly EMI'],
  [1, 'AUDU DANJUMA', 86679, 'OSGF', 12, D(2026, 9, 1), 0, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
  [2, 'OMOLORO OLUWASEYI', 437602, 'OSGF', 12, D(2026, 9, 4), 165375, 144000, 'XXXX', 'XXXX', 'XXXX', 'XXXX', 'XXXX'],
], [6, 28, 12, 14, 8, 14, 14, 14, 24, 24, 16, 16, 16]);
