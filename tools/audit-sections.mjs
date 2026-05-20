import { readFileSync } from 'fs';
import { transformBdSheet } from '../web/js/sheetTransform.js';
import { buildCellMap, computeSectionHeaderRows, computeOutlineRows, getRowLabel, displayCellValue } from '../web/js/bdStore.js';
import { isSectionLabel } from '../web/js/bdStore.js';

const sheet = transformBdSheet(
  JSON.parse(readFileSync('web/public/data/bd-sheet.json', 'utf8'))
);
const map = buildCellMap(sheet.cells, sheet.headerRows);
const sh = sheet.sectionHeaderRows;
const canon = sheet.canonicalSectionByLabel || {};

function raw(r, c) {
  return map.get(`${r}:${c}`)?.v ?? '';
}

const secRows = [...sh].sort((a, b) => a - b);
const fromAp = [];
const fromA = [];
const fromAs = [];

let lastAp = '';
for (let r = 6; r <= sheet.lastRow; r++) {
  const ap = raw(r, 'AP');
  if (isSectionLabel(ap) && ap !== lastAp) {
    fromAp.push({ r, ap });
    lastAp = ap;
  }
  const a = raw(r, 'A');
  if (isSectionLabel(a) && a !== lastAp) {
    fromA.push({ r, a });
    lastAp = a;
  }
}

console.log('sectionHeaderRows count', secRows.length);
console.log('from AP only logic', fromAp.length);
console.log('from A extra', fromA.length, fromA.slice(0, 20));

const apBy = {};
for (const r of secRows) {
  const label =
    displayCellValue(map, r, 'AP', sh, canon) ||
    getRowLabel(map, r, sh, canon) ||
    (r === 5 ? '-ADAPTATION' : r === 139 ? '-ADTH' : '');
  if (!label) continue;
  (apBy[label] ||= []).push(r);
}
const dup = Object.entries(apBy).filter(([, rs]) => rs.length > 1);
console.log('duplicate labels on section rows', dup.length);
dup.forEach(([l, rs]) => console.log(' DUP', l, rs.join(', ')));

console.log('\noutline', sheet.outlineRows.length, 'rows');
for (const r of sheet.outlineRows) {
  const l =
    getRowLabel(map, r, sh, canon) ||
    displayCellValue(map, r, 'AP', sh, canon) ||
    displayCellValue(map, r, 'A', sh, canon) ||
    displayCellValue(map, r, 'AS', sh, canon);
  console.log(' ', r, l);
}
