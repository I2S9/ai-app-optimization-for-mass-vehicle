import fs from 'fs';
import path from 'path';

const TMP = path.join(process.env.TEMP, 'xlsm-bd-export');
const xml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet4.xml'), 'utf8');
const sx = fs.readFileSync(path.join(TMP, 'xl/sharedStrings.xml'), 'utf8');
const ss = [];
const re = /<si>([\s\S]*?)<\/si>/g;
let m;
while ((m = re.exec(sx)) !== null) {
  const p = [];
  let t;
  const tr = /<t[^>]*>([^<]*)<\/t>/g;
  while ((t = tr.exec(m[1])) !== null) p.push(t[1]);
  ss.push(p.join(''));
}

for (const ref of [
  'F15', 'F16', 'F17', 'F18', 'F19', 'F20', 'F21', 'F22', 'D15', 'D16',
]) {
  const cre = new RegExp(`<c r="${ref}"([^>]*)>([\\s\\S]*?)<\\/c>`);
  const cm = cre.exec(xml);
  if (!cm) {
    console.log(ref, 'missing');
    continue;
  }
  const t = cm[1].match(/\bt="([^"]+)"/)?.[1];
  const vM = cm[2].match(/<v>([^<]*)<\/v>/);
  const val = vM
    ? t === 's'
      ? ss[parseInt(vM[1], 10)]
      : vM[1]
    : null;
  console.log(ref, val);
}
