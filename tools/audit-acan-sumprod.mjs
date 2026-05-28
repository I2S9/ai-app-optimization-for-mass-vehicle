import { readFileSync } from 'fs';
import { buildCellMap, getCell, displayCellValue, bdSubsystemL2Col, bdSubsystemL1Col } from '../web/js/bdStore.js';
import { transformSynthesisSheet, transformBdSheet } from '../web/js/sheetTransform.js';
import { displayToExcelCol } from '../web/js/synthesisPerf.js';
import {
  buildBdColumnIndex,
  buildSumproductL2Index,
  buildVehColFilterIndex,
  computeSumproduct,
} from '../web/js/synthesisCalc.js';
import { BD_SUBSYSTEM_L2_COL } from '../web/js/bdColumnConfig.js';
import { translateSubsystemLabel, canonicalL2MatchKey } from '../web/js/bdTranslate.js';

const syn = transformSynthesisSheet(
  JSON.parse(readFileSync('./web/public/data/synthesis-sheet.json', 'utf8'))
);
const bd = transformBdSheet(JSON.parse(readFileSync('./web/public/data/bd-sheet.json', 'utf8')));
const synMap = buildCellMap(syn.cells, syn.headerRows);
const bdMap = buildCellMap(bd.cells, bd.headerRows);
const bdCols = buildBdColumnIndex(bd);
const l1 = bdSubsystemL1Col(bd);
const l2 = bdSubsystemL2Col(bd);
const getSyn = (r, c) => {
  const x = getCell(synMap, r, c);
  return x?.v != null && x.v !== '' ? String(x.v) : '';
};

function getBdDisplay(r, c) {
  return displayCellValue(
    bdMap,
    r,
    c,
    bd.sectionHeaderRows,
    bd.canonicalSectionByLabel,
    l1,
    l2
  );
}
function getBdRaw(r, c) {
  return bdCols[c]?.[r] ?? '';
}

console.log('BD L2 col visible', l2, 'config', BD_SUBSYSTEM_L2_COL);

for (const d of ['M', 'AC']) {
  const ex = displayToExcelCol(d);
  const cell = getCell(synMap, 30, ex);
  const f = cell?.f || '';
  console.log('\nrow30', d, 'excel', ex, 'v', cell?.v);
  if (f) {
    console.log('  has AS', f.includes('AS'), 'has AU', f.includes('AU'));
    console.log('  filter refs', [...f.matchAll(/([A-Z]+)\$3/g)].slice(0, 2).map((m) => m[0]));
  }
}

const acCol = displayToExcelCol('AC');
const vehCol = acCol;
const row = 30;
const label = translateSubsystemLabel(getSyn(row, 'F').trim());
console.log('\nF30 label', label);

const filt = buildVehColFilterIndex(bdCols, vehCol, getSyn, getBdDisplay);
console.log('filter count', filt.length, 'sample filters', [3, 4, 5, 14].map((r) => getSyn(r, vehCol)));

const idx = buildSumproductL2Index(bdCols, getBdDisplay);
const calc = computeSumproduct(bdCols, row, vehCol, getSyn, getBdDisplay, idx, filt);
console.log('engine calc', calc, 'static', getCell(synMap, row, vehCol)?.v);

let manual = 0;
for (const r of filt) {
  const l2v = translateSubsystemLabel(String(getBdDisplay(r, l2)).trim());
  if (l2v !== label) continue;
  manual += parseFloat(String(getBdDisplay(r, 'V')).replace(',', '.')) || 0;
}
console.log('manual sum (display BD, l2 col', l2 + ')', Math.round(manual * 10000) / 10000);

const acCols = 'AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN'.split(',').map(displayToExcelCol);
let miss = 0;
for (let r = 26; r <= 120; r++) {
  for (const col of acCols) {
    const cell = getCell(synMap, r, col);
    const st = cell?.v;
    if (st == null || st === '') continue;
    const n = parseFloat(String(st).replace(',', '.'));
    if (!Number.isFinite(n)) continue;
    const filtC = buildVehColFilterIndex(bdCols, col, getSyn, getBdDisplay);
    const c = computeSumproduct(
      bdCols,
      r,
      col,
      getSyn,
      getBdDisplay,
      buildSumproductL2Index(bdCols, getBdDisplay),
      filtC
    );
    if (Math.abs(c - n) > 0.05) {
      miss++;
      if (miss <= 12) {
        console.log('MISS', 'row', r, 'col', col, 'static', n, 'calc', c, 'F', getSyn(r, 'F').slice(0, 25));
      }
    }
  }
}
console.log('\nTotal misses 26-120 ACAN (numeric static)', miss);
