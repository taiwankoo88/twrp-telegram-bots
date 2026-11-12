// Order workbooks without the host PC: the static order page (site/) builds the same file as excel.mjs in the
// browser. It starts from a base workbook that excel.mjs itself generated (site/base/<project>-<material>.xlsx,
// made by gen-base.mjs: one title, one item) and rewrites only the sheet's rows from row 7 down, reusing the
// base's own cell styles for each kind of row. Layout rules (row order, heights, formulas) mirror excel.mjs
// `generate` / `generateGasket`; change both together and run test-xlsx.mjs, which compares the two outputs.
// Plain JavaScript only (runs in the browser and in Node).
import {describeItem} from './items.mjs';

// ---- zip (stored or deflated entries, as written by the base workbooks) ----
const through = async (bytes, stream) => new Uint8Array(await new Response(new Blob([bytes]).stream().pipeThrough(stream)).arrayBuffer());
export async function unzip(buffer) {
  const b = new Uint8Array(buffer), v = new DataView(b.buffer, b.byteOffset, b.byteLength), text = new TextDecoder();
  let end = b.length - 22;
  while (end >= 0 && v.getUint32(end, true) !== 0x06054b50) end--;
  if (end < 0) throw Error('Not a workbook.');
  const files = [];
  for (let i = 0, p = v.getUint32(end + 16, true); i < v.getUint16(end + 10, true); i++) {
    const method = v.getUint16(p + 10, true), size = v.getUint32(p + 20, true), nameLength = v.getUint16(p + 28, true), at = v.getUint32(p + 42, true);
    const start = at + 30 + v.getUint16(at + 26, true) + v.getUint16(at + 28, true), data = b.subarray(start, start + size);
    files.push({name: text.decode(b.subarray(p + 46, p + 46 + nameLength)), data: method ? await through(data, new DecompressionStream('deflate-raw')) : data.slice()});
    p += 46 + nameLength + v.getUint16(p + 30, true) + v.getUint16(p + 32, true);
  }
  return files;
}
const CRC = Array.from({length: 256}, (_, n) => { for (let k = 0; k < 8; k++) n = n & 1 ? 0xEDB88320 ^ (n >>> 1) : n >>> 1; return n >>> 0; });
const crc32 = bytes => { let c = ~0; for (const x of bytes) c = CRC[(c ^ x) & 255] ^ (c >>> 8); return ~c >>> 0; };
export async function zip(files) {
  const parts = [], central = [];
  let offset = 0;
  for (const {name, data} of files) {
    const packed = await through(data, new CompressionStream('deflate-raw')), label = new TextEncoder().encode(name);
    // version 20, UTF-8 names, deflate, fixed date 1 Jan 2026
    const fields = [20, 0x0800, 8, 0, 0x5C21, crc32(data), packed.length, data.length, label.length, 0];
    const header = (signature, extra) => {
      const h = new DataView(new ArrayBuffer(extra ? 46 : 30));
      h.setUint32(0, signature, true);
      let p = 4;
      if (extra) { h.setUint16(p, 20, true); p += 2; }
      for (const [i, value] of fields.entries()) { if (i >= 5 && i <= 7) { h.setUint32(p, value, true); p += 4; } else { h.setUint16(p, value, true); p += 2; } }
      if (extra) h.setUint32(42, offset, true);
      return new Uint8Array(h.buffer);
    };
    parts.push(header(0x04034b50, false), label, packed);
    central.push(header(0x02014b50, true), label);
    offset += 30 + label.length + packed.length;
  }
  const size = central.reduce((n, x) => n + x.length, 0), tail = new DataView(new ArrayBuffer(22));
  tail.setUint32(0, 0x06054b50, true); tail.setUint16(8, files.length, true); tail.setUint16(10, files.length, true);
  tail.setUint32(12, size, true); tail.setUint32(16, offset, true);
  return new Uint8Array(await new Blob([...parts, ...central, tail.buffer]).arrayBuffer());
}

// ---- sheet ----
const esc = s => String(s).replace(/[^\x09\x0A\x0D\x20-�]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const str = v => ({t: 'str', v}), n = v => ({t: 'n', v});

// Rows of the base sheet: attributes, and each cell's column, style and (for the shared reinforcement text) content.
function readRows(xml) {
  const rows = {};
  for (const [whole, r, attrs, inner = ''] of xml.matchAll(/<x:row r="(\d+)"([^>]*?)(?:\s*\/>|>(.*?)<\/x:row>)/gs)) {
    rows[r] = {xml: whole, attrs, cells: [...inner.matchAll(/<x:c r="([A-Z]+)\d+"([^>]*?)(?:\s*\/>|>(.*?)<\/x:c>)/gs)].map(([cell, col, a, body]) => ({col, s: a.match(/ s="(\d+)"/)?.[1], t: a.match(/ t="(\w+)"/)?.[1], body, xml: cell}))};
  }
  return rows;
}

export const orderFileName = order => order.material === 'EPDM' ? `${order.project} EPDM Gasket Order.xlsx` : `${order.project} ${order.material} Fastener Order.xlsx`;

// base: bytes of site/base/<project>-<material>.xlsx. Returns the bytes of the order workbook.
export async function buildWorkbook(order, base, today = new Date()) {
  const files = await unzip(base), sheet = files.find(f => f.name === 'xl/worksheets/sheet1.xml');
  const xml = new TextDecoder().decode(sheet.data), gasket = order.material === 'EPDM';
  const cols = gasket ? 'ABCDEFGHIJ' : 'ABCDEFGHI', rows = readRows(xml);
  // Base layout: 7 title, 8 item, 9 separator, 10 document note, 11 empty, 12 total, 13 onwards untouched template rows.
  const TITLE = 7, ITEM = 8, NOTE = 10, TOTAL = 12, last = Math.max(...Object.keys(rows).map(Number));
  // What the template leaves in a row outside the order columns (a row height, styled cells to the right).
  const leftover = r => r >= 13 ? rows[r] : rows[13];
  const renumber = (row, r) => row.xml.replace(/(<x:row r=")\d+/, `$1${r}`).replace(/(<x:c r="[A-Z]+)\d+/g, `$1${r}`);
  const empty = (r, height) => {
    const row = leftover(r);
    if (height !== undefined) return `<x:row r="${r}" ht="${height}" customHeight="1">${(row?.cells || []).map(c => c.xml.replace(/(r="[A-Z]+)\d+/, `$1${r}`)).join('')}</x:row>`;
    return row ? renumber(row, r) : '';
  };
  const filled = (r, like, values, height) => {
    const styles = Object.fromEntries(rows[like].cells.map(c => [c.col, c.s]));
    const cells = [...cols].map(col => {
      const {t, v, f, raw} = values[col] || {}, open = `<x:c r="${col}${r}"${styles[col] ? ` s="${styles[col]}"` : ''}${t ? ` t="${t}"` : ''}`;
      if (raw !== undefined) return `${open}>${raw}</x:c>`;
      return f === undefined && v === undefined ? `${open} />` : `${open}>${f === undefined ? '' : `<x:f>${esc(f)}</x:f>`}${v === undefined ? '' : `<x:v>${esc(v)}</x:v>`}</x:c>`;
    });
    const rest = (leftover(r)?.cells || []).filter(c => !cols.includes(c.col)).map(c => c.xml.replace(/(r="[A-Z]+)\d+/, `$1${r}`));
    return `<x:row r="${r}"${height === undefined ? leftover(r)?.attrs || '' : ` ht="${height}" customHeight="1"`}>${cells.join('')}${rest.join('')}</x:row>`;
  };

  const out = [];
  let r = 7;
  if (gasket) {
    const reinforcement = rows[ITEM].cells.find(c => c.col === 'D');
    for (const sec of order.sections) {
      out.push(filled(r, TITLE, {B: str(`Title: ${sec.title}`)}, Math.max(17.4, Math.ceil((7 + sec.title.length) / 42) * 15))); r++;
      for (const [j, item] of sec.items.entries()) {
        const {designation, standard, remark} = describeItem(order.material, item);
        out.push(filled(r, ITEM, {
          A: n(j + 1), B: str(designation), C: str(standard), D: {t: reinforcement.t, raw: reinforcement.body}, E: n(item.pn), F: str('EPDM'),
          G: item.parts.length > 1 ? {t: 'n', f: `SUM(${item.parts.join(',')})`, v: item.qty} : n(item.qty),
          H: remark ? str(remark) : {}, J: {t: 'str', f: `IF(I${r}="","",G${r}*I${r})`},
        }, Math.max(32, Math.ceil((remark || '').length / 20) * 15))); r++;
      }
      out.push(empty(r, 21)); r++;
    }
    out.push(filled(r, NOTE, {A: str('Document Required: Certificate for all items to be attached together with DO.')}, 20));
    const note = r; r++;
    out.push(empty(r)); r++;
    out.push(filled(r, TOTAL, {H: str('Total'), J: {t: 'str', f: `IF(COUNT(I8:I${r - 3})=0,"",SUM(J8:J${r - 3}))`}}, 21));
    var merge = `A${note}:H${note}`, baseMerge = 'A10:H10';
  } else {
    for (const [ix, sec] of order.sections.entries()) {
      out.push(filled(r, TITLE, {A: n(ix + 1), B: str(sec.title)}, Math.max(26, Math.ceil(sec.title.length / 52) * 18))); r++;
      for (const [j, item] of sec.items.entries()) {
        const {designation, standard, grade, remark} = describeItem(order.material, item);
        out.push(filled(r, ITEM, {
          A: n(j + 1), B: str(designation), C: str(standard), D: str(order.material), E: str(grade),
          F: item.parts.length > 1 ? {t: 'n', f: `SUM(${item.parts.join(',')})`, v: item.qty} : n(item.qty),
          G: remark ? str(remark) : {}, I: {t: 'str', f: `IF(H${r}="","",F${r}*H${r})`},
        }, Math.max(26, Math.ceil((remark || '').length / 24) * 17))); r++;
      }
      out.push(empty(r)); r++;
    }
    out.push(filled(r, NOTE, {A: str('Document Required: Mill Test Certificate (MTC) for all fastener items to be attached together with DO.')}, 30));
    const note = r; r++;
    out.push(empty(r)); r++;
    out.push(filled(r, TOTAL, {G: str('Total'), I: {t: 'str', f: `IF(COUNT(H8:H${r - 3})=0,"",SUM(I8:I${r - 3}))`}}));
    var merge = `A${note}:G${note}`, baseMerge = 'A10:G10';
  }
  for (r++; r <= last; r++) out.push(empty(r));

  // Date in Singapore time and requester, in the header rows kept from the base.
  const date = new Intl.DateTimeFormat('en-CA', {timeZone: 'Asia/Singapore', year: 'numeric', month: '2-digit', day: '2-digit'}).format(today);
  const cell = (text, ref, value) => text.replace(new RegExp(`(<x:c r="${ref}"[^>]*>)<x:v>.*?</x:v>`), (_, open) => `${open}<x:v>${esc(value)}</x:v>`);
  const from = xml.indexOf('<x:row r="7"'), to = xml.indexOf('</x:sheetData>');
  if (from < 0 || to < 0 || !xml.includes(`<x:mergeCell ref="${baseMerge}" />`)) throw Error('Unexpected base workbook.');
  const head = cell(cell(xml.slice(0, from), 'B3', `${date}T00:00:00.000Z`), 'B4', order.requester);
  sheet.data = new TextEncoder().encode(head + out.join('') + xml.slice(to).replace(`ref="${baseMerge}"`, `ref="${merge}"`));
  return zip(files);
}
