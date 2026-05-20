import fs from 'fs';
import path from 'path';

const TMP = path.join(process.env.TEMP, 'xlsm-syn-colors');
const stylesXml = fs.readFileSync(path.join(TMP, 'xl/styles.xml'), 'utf8');
const sheetXml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet4.xml'), 'utf8');

const fills = [];
const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
let fm;
while ((fm = fillRe.exec(stylesXml)) !== null) fills.push(fm[1]);

for (const fillId of [16, 48, 71, 167]) {
  console.log('fill', fillId, fills[fillId]?.slice(0, 100));
}

const xfs = [];
const xfRe = /<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g;
let xfm, i = 0;
while ((xfm = xfRe.exec(stylesXml)) !== null) {
  const attrs = xfm[1];
  const fillId = parseInt(attrs.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
  if ([16, 48, 71, 167].includes(fillId)) {
    console.log('xf', i, 'fillId', fillId, 'attrs', attrs.slice(0, 120));
  }
  xfs.push(fillId);
  i++;
}

const byStyle = new Map();
const re = /<c r="F(\d+)"([^>]*)>/g;
let m;
while ((m = re.exec(sheetXml)) !== null) {
  const row = +m[1];
  const s = parseInt(m[2].match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
  if (s < 0) continue;
  if (!byStyle.has(s)) byStyle.set(s, []);
  byStyle.get(s).push(row);
}

for (const [s, rows] of [...byStyle.entries()].sort((a, b) => a[0] - b[0])) {
  if (rows.length <= 30 || rows.includes(26) || rows.includes(25)) {
    console.log('F col style', s, 'fillId', xfs[s], 'rows sample', rows.slice(0, 8), 'count', rows.length);
  }
}
