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

const colLetters = (n) => { let s = ''; while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); } return s; };

/** Widen (never narrow) columns that are too thin to show the numbers we wrote. widths: { zeroBasedColIndex: chars } */
export function widenColumns(xml, widths) {
  const targets = Object.entries(widths).map(([c, w]) => [Number(c) + 1, w]).sort((a, b) => a[0] - b[0]);
  if (!targets.length) return xml;
  const m = xml.match(/<cols>([\s\S]*?)<\/cols>/);
  let defs = m ? (m[1].match(/<col\b[^>]*?\/>/g) || []).map((t) => ({ min: Number(attr(t, 'min')), max: Number(attr(t, 'max')), tag: t })) : [];
  const withW = (d, min, max, width) => {
    let t = d ? d.tag : '<col/>';
    t = t.replace(/\smin="[^"]*"/, '').replace(/\smax="[^"]*"/, '').replace(/\swidth="[^"]*"/, '').replace(/\scustomWidth="[^"]*"/, '').replace(/\/>$/, '');
    return { min, max, tag: `${t} min="${min}" max="${max}"${width != null ? ` width="${width}" customWidth="1"` : ''}/>` };
  };
  for (const [col, need] of targets) {
    const i = defs.findIndex((d) => d.min <= col && col <= d.max);
    if (i < 0) { defs.push(withW(null, col, col, need)); continue; }
    const d = defs[i], cur = Number(attr(d.tag, 'width') || 0);
    if (cur >= need) continue;
    const pieces = [];
    if (d.min < col) pieces.push({ ...d, tag: d.tag.replace(/\smax="[^"]*"/, ` max="${col - 1}"`), max: col - 1 });
    pieces.push(withW(d, col, col, need));
    if (d.max > col) pieces.push({ ...d, tag: d.tag.replace(/\smin="[^"]*"/, ` min="${col + 1}"`), min: col + 1 });
    defs.splice(i, 1, ...pieces);
  }
  defs.sort((a, b) => a.min - b.min);
  const block = `<cols>${defs.map((d) => d.tag).join('')}</cols>`;
  if (m) return xml.replace(m[0], block);
  return xml.replace('<sheetData', block + '<sheetData');
}

export function patchSheetXml(xml, edits, widths = {}) {
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
  return widenColumns(xml, widths);
}

/** JSZip: the JSZip class. Returns a Uint8Array of the patched .xlsx. */
export async function patchXlsx(JSZip, buffer, sheetName, edits, widths = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const path = sheetPath(wbXml, relsXml, sheetName);
  const xml = await zip.file(path).async('string');
  zip.file(path, patchSheetXml(xml, edits, widths));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
