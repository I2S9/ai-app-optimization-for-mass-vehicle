import { readFileSync } from 'fs';
const j = JSON.parse(
  readFileSync('web/public/data/bd-sheet.json', 'utf8')
);
function buildMap(cells, hr) {
  const m = new Map();
  for (const c of cells) m.set(`${c.r}:${c.c}`, c);
  for (const [rs, cols] of Object.entries(hr || {})) {
    const r = +rs;
    for (const [col, c] of Object.entries(cols)) {
      m.set(`${r}:${col}`, { r, c: col, v: c.v });
    }
  }
  return m;
}
const m = buildMap(j.cells, j.headerRows);
const g = (r, c) => {
  const x = m.get(`${r}:${c}`);
  return x?.v != null ? String(x.v) : '';
};
const gas = (r) => g(r, 'AS');
function isCaps(v) {
  if (!v || v.length < 3 || v.length > 45) return false;
  if (v.startsWith('_') || v.startsWith('-')) return false;
  if (['TT', 'STLA/S', 'STLA-S', 'PTF', 'STATUS', 'TARGET'].includes(v))
    return false;
  return v === v.toUpperCase() && /^[A-Z0-9][A-Z0-9 /-]+$/.test(v);
}
const subs = [];
const secs = [];
for (let r = 2; r <= 350; r++) {
  const as = gas(r);
  if (as.startsWith('_') && as !== gas(r - 1)) subs.push(`${r}:${as}`);
  else if (r === 5) secs.push(`${r}:-ADAPTATION`);
  else if (r === 139) secs.push(`${r}:-ADTH`);
  else if (isCaps(g(r, 'A')) && g(r, 'A') !== g(r - 1, 'A'))
    secs.push(`${r}:A:${g(r, 'A')}`);
  else if (
    isCaps(g(r, 'AP')) &&
    g(r, 'AP') !== g(r - 1, 'AP') &&
    !as.startsWith('_')
  )
    secs.push(`${r}:AP:${g(r, 'AP')}`);
}
console.log('sections', secs.length);
secs.forEach((x) => console.log(' S', x));
console.log('subsections', subs.length);
subs.slice(0, 20).forEach((x) => console.log(' s', x));
