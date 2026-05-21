import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-bd-export');
const shared = [];
for (const m of fs
  .readFileSync(path.join(TMP, 'xl/sharedStrings.xml'), 'utf8')
  .matchAll(/<si>([\s\S]*?)<\/si>/g)) {
  const parts = [];
  for (const t of m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)) parts.push(t[1]);
  shared.push(parts.join(''));
}

function resolveCellValue(raw, t) {
  if (raw == null || raw === '') return raw;
  if (t === 's') return shared[+raw] ?? String(raw);
  const s = String(raw).trim();
  if (/^\d+$/.test(s)) {
    const hit = shared[+s];
    if (hit != null && hit !== '' && !/^-?\d+(\.\d+)?$/.test(String(hit).trim())) return hit;
  }
  return s;
}

const sheetXml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet3.xml'), 'utf8');

function parseCol(colLetter) {
  const re = new RegExp(
    `<c r="(${colLetter})(\\d+)"([^/]*?)(\\/>|>([\\s\\S]*?)<\\/c>)`,
    'g'
  );
  const cells = [];
  let m;
  while ((m = re.exec(sheetXml)) !== null) {
    const row = +m[2];
    const inner = m[3] === '/>' ? '' : m[4] || '';
    const t = m[3].match(/\bt="([^"]+)"/)?.[1];
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
    let v = vM ? resolveCellValue(vM[1], t) : '';
    if (fM && !v) v = `(f)${fM[1].slice(0, 30)}`;
    if (!v && !fM) continue;
    cells.push({ row, v: String(v).slice(0, 50), f: fM?.[1]?.slice(0, 40) });
  }
  return cells;
}

for (const col of ['B', 'AU', 'AT', 'AS', 'AP']) {
  const c = parseCol(col);
  const withV = c.filter((x) => x.v && !x.v.startsWith('(f)'));
  console.log(col, 'total', c.length, 'with v', withV.length, 'sample', withV.slice(0, 6));
}

// Compare transform
const { transformBdSheet } = await import('../web/js/sheetTransform.js');
const raw = JSON.parse(
  fs.readFileSync(path.join(__dirname, '../web/public/data/bd-sheet.json'), 'utf8')
);
const t = transformBdSheet(raw);
for (const col of ['B', 'AW', 'AV', 'AU', 'AS', 'AR']) {
  const cells = t.cells.filter((c) => c.c === col && c.v);
  console.log('json', col, cells.length, cells.slice(0, 4).map((c) => ({ r: c.r, v: c.v })));
}
