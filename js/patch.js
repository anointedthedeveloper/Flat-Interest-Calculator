// Writes calculated values straight into the original .xlsx package, changing ONLY the listed cells.
// Everything else (styles, fonts, colours, widths, other sheets, formulas, number formats) stays byte-for-byte.

const attr = (tag, name) => { const m = tag.match(new RegExp(`\\s${name}="([^"]*)"`)); return m ? m[1] : null; };
const colNum = (letters) => letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const split = (addr) => { const m = addr.match(/^([A-Z]+)(\d+)$/); return { col: colNum(m[1]), row: Number(m[2]), letters: m[1] }; };
const esc = (t) => String(t).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));

function sheetPath(workbookXml, relsXml, sheetName) {
  const unesc = (t) => t.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
  const tag = (workbookXml.match(/<sheet\b[^>]*>/g) || []).find((t) => unesc(attr(t, 'name') || '') === sheetName);
  if (!tag) throw new Error('sheet not found');
  const rid = attr(tag, 'r:id');
  const rel = (relsXml.match(/<Relationship\b[^>]*>/g) || []).find((t) => attr(t, 'Id') === rid);
  if (!rel) throw new Error('relationship not found');
  let target = attr(rel, 'Target');
  target = target.startsWith('/') ? target.slice(1) : 'xl/' + target;
  return target;
}

function cellXml(addr, styleAttr, edit) {
  const s = styleAttr ? ` s="${styleAttr}"` : '';
  if (edit.text !== undefined) return `<c r="${addr}"${s} t="inlineStr"><is><t>${esc(edit.text)}</t></is></c>`;
  return `<c r="${addr}"${s}><v>${edit.value}</v></c>`;
}

export function patchSheetXml(xml, edits) {
  let maxCol = 0, maxRow = 0;
  const byRow = new Map();
  for (const e of edits) {
    const a = split(e.addr);
    maxCol = Math.max(maxCol, a.col); maxRow = Math.max(maxRow, a.row);
    if (!byRow.has(a.row)) byRow.set(a.row, []);
    byRow.get(a.row).push({ ...e, ...a });
  }
  const rowRe = /<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g;
  xml = xml.replace(rowRe, (rowXml) => {
    const open = rowXml.match(/^<row\b[^>]*?>/)[0];
    const rowNo = Number(attr(open, 'r'));
    const list = byRow.get(rowNo);
    if (!list) return rowXml;
    byRow.delete(rowNo);
    const selfClosed = open.endsWith('/>');
    const head = selfClosed ? open.slice(0, -2) + '>' : open;
    let inner = selfClosed ? '' : rowXml.slice(open.length, rowXml.length - '</row>'.length);
    // split existing cells
    const cells = [];
    inner.replace(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g, (c) => { cells.push(c); return c; });
    const map = new Map(cells.map((c) => [attr(c.match(/^<c\b[^>]*?>/)[0], 'r'), c]));
    for (const e of list) {
      const existing = map.get(e.addr);
      const style = existing ? attr(existing.match(/^<c\b[^>]*?>/)[0], 's') : null;
      map.set(e.addr, cellXml(e.addr, style, e));
    }
    const sorted = [...map.entries()].sort((a, b) => split(a[0]).col - split(b[0]).col).map((x) => x[1]);
    return head.replace(/\sspans="[^"]*"/, '') + sorted.join('') + '</row>';
  });
  if (byRow.size) { // rows that do not exist yet: append before </sheetData>
    const rows = [...byRow.entries()].sort((a, b) => a[0] - b[0]).map(([r, list]) =>
      `<row r="${r}">` + list.sort((a, b) => a.col - b.col).map((e) => cellXml(e.addr, null, e)).join('') + '</row>');
    xml = xml.replace('</sheetData>', rows.join('') + '</sheetData>');
  }
  // widen the declared used range if we added columns/rows
  xml = xml.replace(/<dimension\b[^>]*?ref="([^"]*)"[^>]*?\/>/, (full, ref) => {
    const m = ref.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!m) return full;
    const endCol = Math.max(m[3] ? colNum(m[3]) : colNum(m[1]), maxCol);
    const endRow = Math.max(m[4] ? Number(m[4]) : Number(m[2]), maxRow);
    const letters = (n) => { let s = ''; while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); } return s; };
    return full.replace(ref, `${m[1]}${m[2]}:${letters(endCol)}${endRow}`);
  });
  return xml;
}

/** JSZip: the JSZip class. Returns a Uint8Array of the patched .xlsx. */
export async function patchXlsx(JSZip, buffer, sheetName, edits) {
  const zip = await JSZip.loadAsync(buffer);
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const path = sheetPath(wbXml, relsXml, sheetName);
  const xml = await zip.file(path).async('string');
  zip.file(path, patchSheetXml(xml, edits));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
