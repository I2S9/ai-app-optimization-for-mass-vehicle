import fs from 'fs';
import path from 'path';
const TMP = path.join(process.env.TEMP, 'xlsm-syn-colors');
const ss = fs.readFileSync(path.join(TMP, 'xl/sharedStrings.xml'), 'utf8');
const sh = [];
for (const m of ss.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
  const p = [];
  let t;
  const tr = /<t[^>]*>([^<]*)<\/t>/g;
  while ((t = tr.exec(m[1])) !== null) p.push(t[1]);
  sh.push(p.join(''));
}
const xml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet4.xml'), 'utf8');
function label(row) {
  const m = xml.match(new RegExp(`<c r="F${row}"([^>]*)>([\\s\\S]*?)<\\/c>`));
  if (!m) return null;
  const t = m[1].match(/\bt="([^"]+)"/)?.[1];
  const v = m[2].match(/<v>([^<]*)<\/v>/)?.[1];
  return t === 's' ? sh[parseInt(v, 10)] : v;
}
for (const row of [51, 60, 62, 75, 177, 204, 297, 318, 406, 416, 25, 26, 41, 42, 421, 422, 100]) {
  console.log(row, JSON.stringify(label(row)));
}
