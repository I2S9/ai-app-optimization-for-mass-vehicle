import fs from 'fs';
const j = JSON.parse(
  fs.readFileSync('web/public/data/synthesis-sheet.json', 'utf8')
);
for (let r = 3; r <= 14; r++) {
  const cols = {};
  for (const c of j.cells.filter((x) => x.r === r && 'A' <= x.c && x.c <= 'F')) {
    cols[c.c] = c.v;
  }
  const hr = j.headerRows?.[r] || {};
  for (const [col, cell] of Object.entries(hr)) {
    if (col <= 'F') cols[col] = cols[col] ?? cell.v;
  }
  console.log('row', r, cols);
}
// vehicle pillar cols
for (const col of ['G', 'H', 'L', 'M', 'P', 'Q']) {
  const cells = j.cells.filter((x) => x.c === col && x.r >= 3 && x.r <= 14);
  const merge = j.merges?.find(
    (m) => m.startCol === col && m.startRow <= 14 && m.endRow >= 3
  );
  console.log(
    'col',
    col,
    'merge',
    merge?.ref,
    'r3',
    cells.find((x) => x.r === 3)?.v,
    'bg sample',
    cells.find((x) => x.bg)?.bg
  );
}
