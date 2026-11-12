// Checks that xlsx.mjs (the static order page's generator) writes the same workbook as excel.mjs (the bot's):
// every row height, cell style, type, value and formula of the sheet, the merged cells, and every other part
// of the file. Run after any change to excel.mjs, xlsx.mjs or the templates (then run gen-base.mjs first).
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {parseOrder} from './parser.mjs';
import {generate} from './excel.mjs';
import {buildWorkbook, orderFileName, unzip} from './xlsx.mjs';

const root = path.dirname(fileURLToPath(import.meta.url));
const long = n => Array.from({length: n}, (_, i) => `M${[10, 12, 16, 20, 24, 30][i % 6]}x${40 + i * 5}mm bolt = ${(i + 1) * 10}`).join('\n');
const orders = [
  `C3D2 HDG\nRequester Name: Web Tester\nTitle: WDL21 Platform handrail\nSets: 50\nM12x135mm bolt = 1\nM12 nut = 3\nM12x3mm washer = 4\nTitle: WDL12 Chemical Hanging Piperack add-on, replacement to ST311-22048, a long title that wraps over more than one line in the sheet\nM16x60mm bolt = 400\nM16 nut = 500\nM16 nut = 55\nM30 washer = 12`,
  `C3D1 SS316\nRequester Name: A & B <Test>\nTitle: WDL13 Membrane Tank Pipe Header Valve Connection\nM16x110mm bolt = 2880\nM16 nut = 2880\nM16x130mm bolt = 5760\nM16 nut = 5760\nM16 nut = 10\nM16 nut = 20\nM16 nut = 30\nM16x3mm flat washer = 17280\nM16x2m stud rod = 7\nTitle: =Second\nM24 washer = 5`,
  `C3D1 SS316\nRequester Name: Long\n${Array.from({length: 6}, (_, s) => `Title: WDL1${s} Area\n${long(14)}`).join('\n')}`,
  `C3D2 HDG\nRequester Name: One\nTitle: Single\nM20 nut = 1`,
  `C3D2 SS316\nRequester Name: =Deep\nTitle: First half\n${long(48)}\nTitle: Second half\n${long(48)}`,
  `C3D1 Gasket\nRequester Name: Web Tester\nTitle: WDL12 Pump Gallery 2 and 3 PSD\nDN400 PN10 = 10\nDN200 PN10 = 30\nDN200 PN10 = 20\nDN100 PN16 5mm = 6\nTitle: WDL21 Dampers with a rather long location name that should wrap in the title cell\nSets: 3\nDN150 PN10 = 4`,
  `C3D2 Gasket\nRequester Name: G\nTitle: ${'Many '.repeat(4)}\n${Array.from({length: 40}, (_, i) => `DN${100 + i * 50} PN10 = ${i + 1}`).join('\n')}`,
];

const text = new TextDecoder();
function sheetOf(xml) {
  const rows = {};
  for (const [, r, attrs, inner = ''] of xml.matchAll(/<x:row r="(\d+)"([^>]*?)(?:\s*\/>|>(.*?)<\/x:row>)/gs)) {
    const cells = {};
    for (const [, ref, a, body = ''] of inner.matchAll(/<x:c r="([A-Z]+\d+)"([^>]*?)(?:\s*\/>|>(.*?)<\/x:c>)/gs)) {
      // An unstyled empty cell is the same as no cell.
      if (a.trim() || body) cells[ref] = `${a.trim()}|${body}`;
    }
    const ht = attrs.match(/ht="([\d.]+)"/), custom = /customHeight="1"/.test(attrs);
    // Heights are stored as 32-bit floats by excel.mjs (17.4 becomes 17.399999618530273).
    rows[r] = {height: ht ? `${Math.round(Number(ht[1]) * 100) / 100}${custom ? ' custom' : ''}` : '', cells};
  }
  for (const r of Object.keys(rows)) if (!rows[r].height && !Object.keys(rows[r].cells).length) delete rows[r];
  return {rows, rest: xml.replace(/<x:sheetData>.*<\/x:sheetData>/s, '')};
}

let failed = 0;
for (const [i, source] of orders.entries()) {
  const order = parseOrder(source), name = `${order.project} ${order.material} #${i + 1}`;
  const want = await unzip(await fs.readFile(await generate(order, `xlsx-test-${i + 1}`)));
  const built = await buildWorkbook(order, await fs.readFile(path.join(root, 'site', 'base', `${order.project}-${order.material}.xlsx`)));
  const dir = path.join(root, 'outputs', `xlsx-test-${i + 1}-site`);
  await fs.mkdir(dir, {recursive: true});
  await fs.writeFile(path.join(dir, orderFileName(order)), built);
  const got = await unzip(built), problems = [];
  // Beyond row 101 the template has no cells of its own, so the bot's note row ("Document Required") picks up
  // the column defaults there (vertical centring, a number format on the empty F cell) and adds styles for it.
  // The page keeps the standard note style at any depth; only that row's style numbers may differ.
  const rowCount = 7 + order.sections.reduce((t, sec) => t + sec.items.length + 2, 0), deep = order.material !== 'EPDM' && rowCount > 101;
  const plain = cell => deep ? cell?.replace(/s="\d+" ?/, '') : cell;
  if (want.map(f => f.name).join() !== got.map(f => f.name).join()) problems.push('different parts');
  for (const part of want) {
    const a = text.decode(part.data), b = text.decode(got.find(f => f.name === part.name)?.data ?? new Uint8Array());
    if (part.name !== 'xl/worksheets/sheet1.xml') {
      // Relationship ids are random in every export.
      const ids = s => s.replace(/R[0-9a-f]{16}/g, 'RID');
      if (ids(a) !== ids(b) && !(deep && part.name === 'xl/styles.xml')) problems.push(`${part.name} differs`);
      continue;
    }
    const x = sheetOf(a), y = sheetOf(b);
    if (x.rest !== y.rest) problems.push('sheet settings / merged cells differ');
    for (const r of new Set([...Object.keys(x.rows), ...Object.keys(y.rows)])) {
      const p = x.rows[r] || {height: '', cells: {}}, q = y.rows[r] || {height: '', cells: {}};
      if (p.height !== q.height) problems.push(`row ${r} height: bot "${p.height}" page "${q.height}"`);
      for (const ref of new Set([...Object.keys(p.cells), ...Object.keys(q.cells)])) if (Number(r) === rowCount ? plain(p.cells[ref]) !== plain(q.cells[ref]) : p.cells[ref] !== q.cells[ref]) problems.push(`${ref}: bot "${p.cells[ref]}" page "${q.cells[ref]}"`);
    }
  }
  if (problems.length) { failed++; console.log(`FAILED ${name}: ${problems.length} differences\n  ${problems.slice(0, 12).join('\n  ')}`); }
  else console.log(`passed ${name}: ${deep ? 'same as the workbook from the bot except the note row style beyond row 101' : 'identical to the workbook from the bot'} (${order.sections.length} titles, ${order.sections.reduce((t, s) => t + s.items.length, 0)} items)`);
}
console.log(failed ? `${failed} of ${orders.length} FAILED` : `all ${orders.length} passed`);
