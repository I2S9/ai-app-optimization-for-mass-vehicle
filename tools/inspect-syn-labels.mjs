import fs from 'fs';
const j = JSON.parse(
  fs.readFileSync('web/public/data/synthesis-sheet.json', 'utf8')
);
const FILTER_LABELS = {
  3: 'Date',
  4: 'Projet',
  5: 'Silhouette',
  6: 'Hybridation',
  7: 'Plaque de conception',
  8: 'Sièges',
  9: 'Spécificité technique',
  10: '',
  11: 'Pôle',
  12: 'Energie',
  13: 'Pack technique',
  14: 'Finition',
};
for (let r = 3; r <= 14; r++) {
  const f = j.cells.find((c) => c.r === r && c.c === 'F')?.v;
  const d = j.cells.find((c) => c.r === r && c.c === 'D')?.v;
  const e = j.cells.find((c) => c.r === r && c.c === 'E')?.v;
  const a = j.cells.find((c) => c.r === r && c.c === 'A')?.v;
  console.log(r, 'expected:', FILTER_LABELS[r], '| F:', f, '| D:', d, '| E:', e, '| A:', a);
}
// merges for G and P vertical headers
for (const m of j.merges || []) {
  if (
    (m.startCol === 'G' || m.startCol === 'P' || m.startCol === 'H') &&
    m.startRow <= 14
  ) {
    const master = j.cells.find(
      (c) => c.r === m.startRow && c.c === m.startCol
    );
    console.log('merge', m.ref, 'master v:', master?.v, 'bg:', master?.bg);
  }
}
