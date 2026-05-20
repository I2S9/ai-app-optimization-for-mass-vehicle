import { readFileSync } from 'fs';
import { buildCellMap, getAsLabel, displayValue, getCell } from '../web/js/bdStore.js';
import { transformBdSheet } from '../web/js/sheetTransform.js';

const sheet = transformBdSheet(
  JSON.parse(readFileSync('web/public/data/bd-sheet.json', 'utf8'))
);
const map = buildCellMap(sheet.cells, sheet.headerRows);

function isAllCapsSection(v) {
  if (!v || v.length < 3 || v.length > 45) return false;
  if (v.startsWith('_') || v.startsWith('-')) return false;
  if (['TT', 'STLA/S', 'STLA-S', 'PTF', 'STATUS', 'TARGET', 'FWD', 'AWD'].includes(v))
    return false;
  return v === v.toUpperCase() && /^[A-Z0-9][A-Z0-9 \/\-–]*$/.test(v);
}

function isSection(map, row) {
  if (row === 5) return true;
  const a = displayValue(getCell(map, row, 'A'));
  if (isAllCapsSection(a)) {
    const pa = displayValue(getCell(map, row - 1, 'A'));
    if (a !== pa) return true;
  }
  const ap = displayValue(getCell(map, row, 'AP'));
  if (isAllCapsSection(ap)) {
    const pap = displayValue(getCell(map, row - 1, 'AP'));
    if (ap !== pap && !getAsLabel(map, row).startsWith('_')) return true;
  }
  for (const col of ['W', 'AP', 'A']) {
    const v = displayValue(getCell(map, row, col));
    if (v === '-ADAPTATION' && row === 5) return true;
    if (v === '-ADTH') {
      const as = getAsLabel(map, row);
      if (!as.startsWith('_') || row === 139) return true;
      if (row === 139 || (col === 'W' && v === '-ADTH' && !as.startsWith('_DIVERS')))
        return true;
    }
  }
  if (row === 139) return true;
  return false;
}

function isSubSection(map, row) {
  const as = getAsLabel(map, row);
  if (!as.startsWith('_')) return false;
  return as !== getAsLabel(map, row - 1);
}

const sections = [];
const subs = [];
for (let r = 2; r <= 400; r++) {
  if (isSubSection(map, r)) subs.push({ r, as: getAsLabel(map, r) });
  else if (isSection(map, r))
    sections.push({
      r,
      a: displayValue(getCell(map, r, 'A')),
      as: getAsLabel(map, r),
    });
}
console.log('sections', sections.length, sections.slice(0, 25));
console.log('subs', subs.length, subs.slice(0, 25));
