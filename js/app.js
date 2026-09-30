import { validateSettings, formatNaira } from './calc.js';
import { processWorkbook, summarize, previewSheet } from './sheet.js';

const $ = (id) => document.getElementById(id);
let file = null, buffer = null, lastWb = null;

const show = (el, on = true) => { el.hidden = !on; };
const fail = (el, msg) => { el.textContent = msg; show(el, !!msg); };

function readWorkbook(buf) {
  return XLSX.read(buf, { type: 'array', cellNF: true, cellStyles: true, cellDates: false });
}

const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const icon = (id) => `<svg><use href="#i-${id}"/></svg>`;
function setStatus(kind, html) { // kind: busy | ok | err
  const el = $('status');
  el.className = 'status ' + kind;
  el.innerHTML = icon(kind === 'busy' ? 'spin' : kind === 'ok' ? 'check' : 'alert') + `<div>${html}</div>`;
  show(el);
}

async function loadFile(f) {
  fail($('error'), '');
  show($('results'), false); show($('preview'), false);
  file = null; $('calc').disabled = true;
  if (!/\.(xlsx|xls)$/i.test(f.name)) { setStatus('err', `<b>${esc(f.name)}</b><small>Unsupported file type. Please upload an .xlsx or .xls file.</small>`); return; }
  setStatus('busy', `Reading <b>${esc(f.name)}</b>...`);
  try {
    buffer = await f.arrayBuffer();
    const wb = readWorkbook(buffer);
    const pv = previewSheet(XLSX, wb);
    const probe = processWorkbook(XLSX, readWorkbook(buffer), { rate: 5, tenure: 12 });
    file = f;
    const kb = f.size < 1024 * 1024 ? Math.max(1, Math.round(f.size / 1024)) + ' KB' : (f.size / 1048576).toFixed(1) + ' MB';
    setStatus('ok', `<b>${esc(f.name)}</b> uploaded (${kb})<small>${probe.detected} loan record${probe.detected === 1 ? '' : 's'} detected in sheet "${esc(probe.sheetName)}"${probe.usedFallbackPrincipal ? ' — using the "Principal" column' : ''}. Ready to calculate.</small>`);
    $('previewTable').tHead.innerHTML = '<tr>' + pv.headers.map((h) => `<th>${esc(h)}</th>`).join('') + '</tr>';
    $('previewTable').tBodies[0].innerHTML = pv.rows.map((r) => '<tr>' + r.map((v) => `<td>${esc(v)}</td>`).join('') + '</tr>').join('');
    $('previewNote').textContent = `(showing ${pv.rows.length} of ${pv.total} rows)`;
    show($('preview'));
    $('calc').disabled = false;
  } catch (e) {
    setStatus('err', `<b>${esc(f.name)}</b><small>${esc(e.message || 'Could not read this Excel file.')}</small>`);
  }
}

$('choose').onclick = () => $('file').click();
$('file').onchange = (e) => e.target.files[0] && loadFile(e.target.files[0]);
const drop = $('drop');
['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); setStatus('busy', 'Release to upload your file'); $('status').className = 'status'; }));
['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));
drop.addEventListener('dragleave', () => { if (!file) show($('status'), false); });
drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
$('auto').onchange = () => { $('tenure').disabled = $('auto').checked; };

$('calc').onclick = () => {
  fail($('settingsError'), ''); fail($('error'), '');
  const rate = parseFloat($('rate').value), tenure = Number($('tenure').value), auto = $('auto').checked;
  const err = validateSettings(rate, tenure, auto);
  if (err) { fail($('settingsError'), err); return; }
  try {
    lastWb = readWorkbook(buffer); // fresh copy each run
    const out = processWorkbook(XLSX, lastWb, { rate, tenure, autoTenure: auto });
    render(out);
  } catch (e) { fail($('error'), e.message); show($('results'), false); }
};

function render(out) {
  const s = summarize(out.rows);
  const stat = (l, v) => `<div class="stat"><span>${l}</span><b>${v}</b></div>`;
  $('cards').innerHTML = stat('Total Loans', s.count) + stat('Total Principal', formatNaira(s.principal)) +
    stat('Total Monthly Interest', formatNaira(s.monthlyInterest)) + stat('Total Interest', formatNaira(s.totalInterest)) +
    stat('Total Repayment', formatNaira(s.totalRepayment)) + stat('Avg Monthly Repayment', formatNaira(s.avgMonthly));

  const inv = $('invalid');
  if (out.invalid.length) {
    inv.innerHTML = `<strong>${out.invalid.length} row${out.invalid.length === 1 ? '' : 's'} could not be calculated (left unchanged in the file):</strong><ul>` +
      out.invalid.map((r) => `<li>Row ${r.row}${r.label.startsWith('Row ') ? '' : ' (' + esc(r.label) + ')'}: ${esc(r.reason)}</li>`).join('') + '</ul>';
    show(inv);
  } else show(inv, false);

  $('table').tBodies[0].innerHTML = out.rows.map((r) => `<tr><td>${esc(r.label)}</td>` +
    [r.principal, r.monthlyInterest, r.totalInterest, r.totalRepayment, r.monthlyRepayment].map((v) => `<td class="r">${formatNaira(v)}</td>`).join('') +
    `<td class="r">${r.tenure} mo</td><td class="r">${r.rate}%</td><td>${esc(r.status)}</td></tr>`).join('');
  show($('results'));
  $('results').scrollIntoView({ behavior: 'smooth' });
}

$('download').onclick = () => {
  const d = new Date(), p = (n) => String(n).padStart(2, '0');
  XLSX.writeFile(lastWb, `loan_calculated_${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.xlsx`, { bookType: 'xlsx', cellStyles: true });
};

$('sample').onclick = () => {
  const H = ['Payment Date', 'Bal. restruc', 'Bank payment', 'Gross bank paym', 'New Principal', 'Interest', 'Gross Loan', 'Monthly repayment', 'Start Date', 'End date', 'Status'];
  const D = (y, m, d) => new Date(Date.UTC(y, m - 1, d));
  const ws = XLSX.utils.aoa_to_sheet([H,
    [D(2026, 9, 1), null, 144000, 150000, 150000, null, null, null, D(2026, 10, 1), D(2027, 9, 30), 'RENEWAL'],
    [D(2026, 9, 4), 165375, 144000, 150000, 315375, null, null, null, D(2026, 10, 1), D(2027, 9, 30), 'TOP UP'],
  ], { cellDates: true });
  ws['!cols'] = H.map(() => ({ wch: 16 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Loans');
  XLSX.writeFile(wb, 'sample_loans.xlsx');
};
