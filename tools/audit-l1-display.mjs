import fs from 'fs';
import { transformBdSheet } from '../web/js/sheetTransform.js';
import {
  buildCellMap,
  displayCellValue,
  bdSubsystemL1Col,
  isStructureRow,
  isSectionRow,
  shouldDateColBlue,
  displayValue,
  getCell,
  colATitle,
} from '../web/js/bdStore.js';
import { translateValue } from '../web/js/bdTranslate.js';

const TMP = `${process.env.TEMP}/xlsm-bd-export`;
const sheetXml = fs.readFileSync(`${TMP}/xl/worksheets/sheet3.xml`, 'utf8');
const shared = [];
for (const m of fs.readFileSync(`${TMP}/xl/sharedStrings.xml`, 'utf8').matchAll(
  /<si>([\s\S]*?)<\/si>/g
)) {
  const parts = [];
  for (const t of m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)) parts.push(t[1]);
  shared.push(parts.join(''));
}

function excelAp(row) {
  const re = new RegExp(
    `<c r="AP${row}"([^/]*?)(\\/>|>([\\s\\S]*?)<\\/c>)`
  );
  const m = sheetXml.match(re);
  if (!m) return '';
  const inner = m[2] === '/>' ? '' : m[3] || '';
  const t = m[1].match(/\bt="([^"]+)"/)?.[1];
  const vM = inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return '';
  const raw = vM[1];
  return t === 's' ? shared[+raw] ?? raw : raw;
}

const raw = JSON.parse(fs.readFileSync('web/public/data/bd-sheet.json', 'utf8'));
const t = transformBdSheet(raw);
const map = buildCellMap(t.cells, t.headerRows);
const l1 = bdSubsystemL1Col(t);
const sh = t.sectionHeaderRows;
const canon = t.canonicalSectionByLabel;

let hidden = 0;
let mism = 0;
for (let r = 6; r <= 800; r++) {
  const ex = excelAp(r);
  if (!ex) continue;
  const web = displayCellValue(map, r, l1, sh, canon, l1, 'AU');
  const cellV = map.get(`${r}:${l1}`)?.v ?? '';
  const en = translateValue(ex);
  if (ex && !web) hidden++;
  if (web && en && web !== en && web !== cellV) mism++;
}
console.log('L1 col', l1, 'hidden excel values', hidden);
let match = 0;
let diff = 0;
for (let r = 6; r <= 1200; r++) {
  const ex = excelAp(r);
  if (!ex) continue;
  const web = displayCellValue(map, r, l1, sh, canon, l1, 'AU');
  const expected = translateValue(ex);
  if (web === expected || web === ex) match++;
  else {
    diff++;
    if (diff <= 5) console.log('diff', r, { ex, expected, web });
  }
}
console.log('match', match, 'diff', diff);
for (const r of [5, 6, 7, 8, 9, 21, 246, 272]) {
  const ex = excelAp(r);
  const web = displayCellValue(map, r, l1, sh, canon, l1, 'AU');
  const raw = map.get(`${r}:${l1}`)?.v;
  const cell = map.get(`${r}:${l1}`);
  console.log('row', r, {
    ex,
    raw,
    web,
    dv: displayValue(cell),
    blue: shouldDateColBlue(map, r, sh),
    a: colATitle(map, r),
    l1Col: l1,
    colIsL1: l1 === l1,
  });
}
