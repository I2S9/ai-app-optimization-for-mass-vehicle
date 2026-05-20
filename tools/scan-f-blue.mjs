import fs from 'fs';
import path from 'path';

const TMP = path.join(process.env.TEMP, 'xlsm-syn-colors');
const stylesXml = fs.readFileSync(path.join(TMP, 'xl/styles.xml'), 'utf8');
const sheetXml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet4.xml'), 'utf8');

const fills = [];
for (const m of stylesXml.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) fills.push(m[1]);

const xfs = [];
for (const m of stylesXml.matchAll(/<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g)) {
  const fillId = parseInt(m[1].match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
  xfs.push({ fillId, fill: fills[fillId] || '' });
}

const blueRows = [];
const yellowRows = [];
const greenRows = [];
for (const m of sheetXml.matchAll(/<c r="F(\d+)"([^>]*)>/g)) {
  const row = +m[1];
  const s = parseInt(m[2].match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
  if (s < 0) continue;
  const { fillId, fill } = xfs[s] || {};
  if (/FF00B0F0/i.test(fill)) blueRows.push({ row, s, fillId });
  if (/FFFFFF00/i.test(fill)) yellowRows.push({ row, s, fillId });
  if (/indexed="44"/i.test(fill)) greenRows.push({ row, s, fillId });
}

console.log('F blue rows', blueRows.length, blueRows.slice(0, 20));
console.log('F yellow rows', yellowRows.length, yellowRows.slice(0, 20));
console.log('F indexed44 rows', greenRows.length, greenRows);

for (const row of [26, 42, 27, 100]) {
  const cols = [];
  for (const m of sheetXml.matchAll(new RegExp(`<c r="([A-Z]+)${row}"([^>]*)>`, 'g'))) {
    const col = m[1];
    if (col > 'P') continue;
    const s = parseInt(m[2].match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    const { fillId, fill } = xfs[s] || {};
    if (fill && !/none/i.test(fill)) cols.push({ col, s, fillId, fill: fill.slice(0, 60) });
  }
  console.log('row', row, cols);
}
