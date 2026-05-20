import fs from 'fs';
import path from 'path';

const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-syn-colors');
const INDEXED = {
  5: '#FFFF00', 13: '#FFFF00', 44: '#CCFFCC', 46: '#99CCFF',
  41: '#CCFFFF', 64: '#333399',
};

function resolveFill(block) {
  if (!block) return null;
  const indexed = block.match(/indexed="(\d+)"/)?.[1];
  if (indexed != null) return INDEXED[indexed] ?? `indexed:${indexed}`;
  const rgb = block.match(/rgb="([^"]+)"/)?.[1];
  if (rgb) return `#${rgb.length === 8 ? rgb.slice(2) : rgb}`;
  return block.slice(0, 80);
}

const stylesXml = fs.readFileSync(path.join(TMP, 'xl', 'styles.xml'), 'utf8');
const sheetXml = fs.readFileSync(path.join(TMP, 'xl', 'worksheets', 'sheet4.xml'), 'utf8');
const fills = [];
const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
let fm;
while ((fm = fillRe.exec(stylesXml)) !== null) fills.push(fm[1]);
const xfs = [];
const xfRe = /<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g;
let xfm;
while ((xfm = xfRe.exec(stylesXml)) !== null) {
  const attrs = xfm[1];
  const fillId = parseInt(attrs.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
  const apply = fillId > 0 || attrs.includes('applyFill="1"');
  xfs.push(apply ? resolveFill(fills[fillId]) : null);
}

for (const row of [25, 26, 41, 42, 100, 421]) {
  console.log(`\n=== Row ${row} ===`);
  const re = new RegExp(`<c r="([A-Z]+)${row}"([^>]*)>`, 'g');
  let m;
  while ((m = re.exec(sheetXml)) !== null) {
    const col = m[1];
    if (col > 'N') break;
    const s = parseInt(m[2].match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    if (s < 0) continue;
    const bg = xfs[s];
    if (bg) console.log(col, 's=', s, 'bg=', bg);
  }
}
