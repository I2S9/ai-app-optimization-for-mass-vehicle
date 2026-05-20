import fs from 'fs';
import path from 'path';
const TMP = path.join(process.env.TEMP, 'xlsm-syn-colors');
const stylesXml = fs.readFileSync(path.join(TMP, 'xl/styles.xml'), 'utf8');
const sheetXml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet4.xml'), 'utf8');
const INDEXED = { 44: '#CCFFCC', 5: '#FFFF00', 64: '#333399' };
function resolve(node) {
  if (!node) return null;
  const rgb = node.match(/rgb="([^"]+)"/)?.[1];
  if (rgb) return `#${rgb.length === 8 ? rgb.slice(2) : rgb}`.toLowerCase();
  const i = node.match(/indexed="(\d+)"/)?.[1];
  return i != null ? (INDEXED[i] ?? `idx:${i}`) : null;
}
const fills = [null];
for (const m of stylesXml.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) {
  const b = m[1];
  const fg = resolve(b.match(/<fgColor[^>]*\/>/)?.[0]) || resolve(b.match(/<fgColor[^>]*>[\s\S]*?<\/fgColor>/)?.[0]);
  fills.push(fg);
}
const xfs = [];
for (const m of stylesXml.matchAll(/<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g)) {
  const fillId = parseInt(m[1].match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
  const apply = fillId > 0 || m[1].includes('applyFill="1"');
  xfs.push(apply ? fills[fillId] : null);
}
for (const col of ['A', 'F']) {
  console.log(`\n=== Column ${col} ===`);
  for (let row = 15; row <= 45; row++) {
    const m = sheetXml.match(new RegExp(`<c r="${col}${row}"([^>]*)>`));
    if (!m) continue;
    const s = parseInt(m[1].match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    const bg = s >= 0 ? xfs[s] : null;
    if (bg) console.log(row, bg);
  }
}
