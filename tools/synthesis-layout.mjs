import fs from 'fs';
import path from 'path';

const shared = [];
const sx = fs.readFileSync(path.join(process.env.TEMP, 'xlsm-analyze/xl/sharedStrings.xml'), 'utf8');
let sm;
const sre = /<si>([\s\S]*?)<\/si>/g;
while ((sm = sre.exec(sx)) !== null) {
  const p = [];
  let t;
  const tr = /<t[^>]*>([^<]*)<\/t>/g;
  while ((t = tr.exec(sm[1])) !== null) p.push(t[1]);
  shared.push(p.join(''));
}

const xml = fs.readFileSync(path.join(process.env.TEMP, 'xlsm-analyze/xl/worksheets/sheet4.xml'), 'utf8');
const cells = new Map();
const cre = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
let cm;
while ((cm = cre.exec(xml)) !== null) {
  const ref = cm[1];
  const attrs = cm[2];
  const inner = cm[3];
  const t = attrs.match(/\bt="([^"]+)"/)?.[1];
  let value = null,
    formula = null;
  const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
  if (fM) formula = fM[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
  const vM = inner.match(/<v>([^<]*)<\/v>/);
  if (vM) value = t === 's' ? shared[parseInt(vM[1], 10)] : vM[1];
  cells.set(ref, { formula, value });
}

function colNum(c) {
  let n = 0;
  for (const x of c) n = n * 26 + x.charCodeAt(0) - 64;
  return n;
}
function numCol(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Vehicle columns: cols with SUMPRODUCT in row 27
const vehicleCols = [];
for (let i = 7; i <= colNum('NI'); i++) {
  const col = numCol(i);
  const cell = cells.get(`${col}27`);
  if (cell?.formula?.includes('SUMPRODUCT')) vehicleCols.push(col);
}
console.log('Vehicle columns with SUMPRODUCT (count):', vehicleCols.length);
console.log('First 20:', vehicleCols.slice(0, 20).join(', '));
console.log('Last 10:', vehicleCols.slice(-10).join(', '));

// Data rows: F column labels from row 15 to 530
const dataRows = [];
for (let r = 15; r <= 530; r++) {
  const f = cells.get(`F${r}`);
  if (f?.value || f?.formula) dataRows.push({ row: r, label: f.value || '(formula)' });
}
console.log('\nSynthesis aggregation rows (F col):', dataRows.length);
console.log('Sample labels:', dataRows.slice(0, 15).map((x) => x.label));
console.log('Last labels:', dataRows.slice(-5).map((x) => x.label));

// Map synthesis filter rows to BD columns (from SUMPRODUCT)
const filterMap = {
  R3: 'A Date/Org',
  R4: 'B Project',
  R5: 'C Silhouette',
  R6: 'E Hybridation',
  R7: 'D Plaque',
  R8: 'F Seats',
  R9: 'G Tech spec',
  R11: 'K Pole',
  R12: 'L Energy',
  R13: 'I Pack',
  R14: 'O Finish',
};
console.log('\nFilter row mapping (SYNTHESIS row → BD column):');
for (const [sr, label] of Object.entries(filterMap)) console.log(sr, label);

for (const ref of ['G27', 'H27', 'I27', 'P27', 'R27']) {
  const c = cells.get(ref);
  console.log(ref, c?.formula ? 'FORMULA ' + c.formula.slice(0, 80) : 'VALUE ' + c?.value);
}
