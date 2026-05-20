import fs from 'fs';
import path from 'path';
const xml = fs.readFileSync(
  path.join(process.env.TEMP, 'xlsm-analyze/xl/worksheets/sheet4.xml'),
  'utf8'
);
const re = /<c r="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/c>/g;
let m, c = 0;
while ((m = re.exec(xml)) && c < 2) {
  const inner = m[2];
  const f = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
  if (f && f[1].includes('SUMPRODUCT')) {
    console.log('CELL', m[1]);
    console.log(
      f[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&')
    );
    console.log('---\n');
    c++;
  }
}
// SYNTHESIS layout: row 14 = first data row with vehicle config, row 15+ = component groups
// Find column groups - row 4 project names
import { readFileSync } from 'fs';
const shared = [];
const sx = readFileSync(
  path.join(process.env.TEMP, 'xlsm-analyze/xl/sharedStrings.xml'),
  'utf8'
);
const sre = /<si>([\s\S]*?)<\/si>/g;
let sm;
while ((sm = sre.exec(sx)) !== null) {
  const p = [];
  let t;
  const tr = /<t[^>]*>([^<]*)<\/t>/g;
  while ((t = tr.exec(sm[1])) !== null) p.push(t[1]);
  shared.push(p.join(''));
}
function parseCells(file) {
  const x = readFileSync(file, 'utf8');
  const cells = new Map();
  const cre = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
  let cm;
  while ((cm = cre.exec(x)) !== null) {
    const ref = cm[1];
    const attrs = cm[2];
    const inner = cm[3];
    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    let value = null,
      formula = null;
    const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
    if (fM)
      formula = fM[1]
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&');
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = t === 's' ? shared[parseInt(vM[1], 10)] : vM[1];
    cells.set(ref, { formula, value });
  }
  return cells;
}
const syn = parseCells(
  path.join(process.env.TEMP, 'xlsm-analyze/xl/worksheets/sheet4.xml')
);
console.log('=== SYNTHESIS vehicle columns row 4 (G onwards sample) ===');
for (const col of ['G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']) {
  const cell = syn.get(`${col}4`);
  if (cell) console.log(col, cell.value || cell.formula?.slice(0, 30));
}
console.log('\n=== First SUMPRODUCT cell with value - scan G column rows 15-100 ===');
for (let r = 15; r <= 100; r++) {
  const cell = syn.get(`G${r}`);
  if (cell?.formula?.includes('SUMPRODUCT')) {
    console.log(`G${r} label F:`, syn.get(`F${r}`)?.value);
    console.log(cell.formula.slice(0, 1200));
    break;
  }
}
console.log('\n=== SYNTHESIS editable header block rows 3-13 col G-T ===');
for (let r = 3; r <= 13; r++) {
  const vals = [];
  for (const col of ['F', 'G', 'I', 'P', 'S', 'T']) {
    const cell = syn.get(`${col}${r}`);
    if (cell?.value && !cell.formula) vals.push(`${col}=${cell.value}`);
    else if (cell?.formula) vals.push(`${col}=formula`);
  }
  if (vals.length) console.log(`R${r}:`, vals.join(', '));
}
