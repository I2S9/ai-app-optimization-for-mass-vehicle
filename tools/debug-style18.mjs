import fs from 'fs';
import path from 'path';
const TMP = path.join(process.env.TEMP, 'xlsm-syn-colors');
const INDEXED_COLORS = {
  44: '#CCFFCC', 47: '#FF99CC', 64: '#333399',
};
function resolveColor(node) {
  if (!node) return null;
  const rgb = node.match(/rgb="([^"]+)"/)?.[1];
  if (rgb) {
    const hex = rgb.length === 8 ? rgb.slice(2) : rgb;
    return `#${hex}`.toLowerCase();
  }
  const indexed = node.match(/indexed="(\d+)"/)?.[1];
  if (indexed != null) return INDEXED_COLORS[indexed] ?? `idx:${indexed}`;
  return null;
}
const stylesXml = fs.readFileSync(path.join(TMP, 'xl/styles.xml'), 'utf8');
const fills = [{ fg: null }];
const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
let fm;
let fillIdx = 0;
while ((fm = fillRe.exec(stylesXml)) !== null) {
  const block = fm[1];
  const fg =
    resolveColor(block.match(/<fgColor[^>]*\/>/)?.[0] || block.match(/<fgColor[^>]*>/)?.[0]) ||
    resolveColor(block.match(/<fgColor[^>]*[^/]>[\s\S]*?<\/fgColor>/)?.[0]);
  const bg =
    resolveColor(block.match(/<bgColor[^>]*\/>/)?.[0] || block.match(/<bgColor[^>]*>/)?.[0]);
  fills.push({ fg: fg || bg });
  if (fillIdx === 15 || fillIdx === 16) console.log('at push', fillIdx, '-> fills', fills.length - 1, { fg, bg });
  fillIdx++;
}
const xfs = [];
const xfRe = /<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g;
let xfm;
while ((xfm = xfRe.exec(stylesXml)) !== null) {
  const attrs = xfm[1];
  const fillId = parseInt(attrs.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
  const applyFill = fillId > 0 || attrs.includes('applyFill="1"');
  const f = fills[fillId] || {};
  xfs.push(applyFill && f.fg ? f.fg : null);
}
const block = fills[16];
let i = 0;
for (const m of stylesXml.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) {
  if (i++ === 16) {
    console.log('raw', m[1]);
    const node = m[1].match(/<fgColor[^>]*\/>/)?.[0];
    console.log('node', node, 'resolved', resolveColor(node));
  }
}
console.log('fill16', fills[16]);
console.log('xf18', xfs[18]);
