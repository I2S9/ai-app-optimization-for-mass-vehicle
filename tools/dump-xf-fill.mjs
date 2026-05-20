import fs from 'fs';
import path from 'path';
const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-syn-colors');
const xml = fs.readFileSync(path.join(TMP, 'xl', 'styles.xml'), 'utf8');
const fills = [];
const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
let fm;
while ((fm = fillRe.exec(xml)) !== null) fills.push(fm[1]);
const xfs = [];
const xfRe = /<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g;
let xfm;
while ((xfm = xfRe.exec(xml)) !== null) xfs.push(xfm[1]);
for (const idx of [18, 126, 740, 741, 742, 743, 744, 745, 755, 4]) {
  const xf = xfs[idx];
  const fillId = parseInt(xf.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
  console.log('xf', idx, 'fillId', fillId);
  console.log(fills[fillId]);
  console.log('---');
}
