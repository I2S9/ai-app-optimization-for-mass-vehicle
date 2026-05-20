import fs from 'fs';
import path from 'path';
const sheet = fs.readFileSync(
  path.join(process.env.TEMP, 'xlsm-bd-export', 'xl/worksheets/sheet3.xml'),
  'utf8'
);
const shared = [];
const sx = fs.readFileSync(
  path.join(process.env.TEMP, 'xlsm-bd-export', 'xl/sharedStrings.xml'),
  'utf8'
);
let sm;
while ((sm = /<si>([\s\S]*?)<\/si>/g.exec(sx)) !== null) {
  const p = [];
  let t;
  while ((t = /<t[^>]*>([^<]*)<\/t>/g.exec(sm[1])) !== null) p.push(t[1]);
  shared.push(p.join(''));
}
for (let col of 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('').concat(['AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV'])) {
  const m = sheet.match(new RegExp(`<c r="${col}5"([^>]*)>([\\s\\S]*?)<\\/c>`));
  if (!m) continue;
  const inner = m[2];
  const f = inner.match(/<f[^>]*>([^<]*)<\/f>/)?.[1];
  const v = inner.match(/<v>([^<]*)<\/v>/)?.[1];
  const t = m[1].match(/t="s"/) ? shared[parseInt(v, 10)] : v;
  if (t || f) console.log(col + '5:', t || '', f || '');
}
const merges = [...sheet.matchAll(/mergeCell ref="([^"]+)"/g)].map((m) => m[1]);
console.log('merges on row 5', merges.filter((r) => r.includes('5')));
