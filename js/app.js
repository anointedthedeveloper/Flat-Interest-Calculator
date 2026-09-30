import { validateSettings, formatNaira } from './calc.js';
import { processWorkbook, summarize } from './sheet.js';

const $ = (id) => document.getElementById(id);
let file = null, buffer = null, lastWb = null;

const show = (el, on = true) => { el.hidden = !on; };
const fail = (el, msg) => { el.textContent = msg; show(el, !!msg); };

function readWorkbook(buf) {
  return XLSX.read(buf, { type: 'array', cellNF: true, cellStyles: true, cellDates: false });
}

async function loadFile(f) {
  fail($('error'), '');
  show($('results'), false);
  if (!/\.(xlsx|xls)$/i.test(f.name)) { fail($('error'), 'Please choose an .xlsx or .xls file.'); return; }
  try {
    buffer = await f.arrayBuffer();
    const wb = readWorkbook(buffer);
    // count rows by dry-run on a throwaway copy
    const probe = processWorkbook(XLSX, readWorkbook(buffer), { rate: 5, tenure: 12 });
    file = f;
    $('fileInfo').textContent = `📄 ${f.name} — ${probe.detected} loan record${probe.detected === 1 ? '' : 's'} detected (sheet "${probe.sheetName}")`;
    show($('fileInfo'));
    $('calc').disabled = false;
  } catch (e) {
    file = null; $('calc').disabled = true;
    $('fileInfo').textContent = `📄 ${f.name}`; show($('fileInfo'));
    fail($('error'), e.message || 'Could not read this Excel file.');
  }
}

$('choose').onclick = () => $('file').click();
$('file').onchange = (e) => e.target.files[0] && loadFile(e.target.files[0]);
const drop = $('drop');
['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));
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

  const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
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
