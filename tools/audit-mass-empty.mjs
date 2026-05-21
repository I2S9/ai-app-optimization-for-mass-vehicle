import fs from 'fs';
import { transformBdSheet } from '../web/js/sheetTransform.js';
import {
  buildCellMap,
  displayCellValue,
  bdSubsystemL1Col,
  bdSubsystemL2Col,
  bdMassCol,
} from '../web/js/bdStore.js';

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

function excelV(row) {
  const re = new RegExp(
    `<c r="V${row}"([^/]*?)(\\/>|>([\\s\\S]*?)<\\/c>)`
  );
  const m = sheetXml.match(re);
  if (!m) return '';
  const inner = m[2] === '/>' ? '' : m[3] || '';
  const t = m[1].match(/\bt="([^"]+)"/)?.[1];
  const vM = inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return '';
  const raw = vM[1];
  if (t === 's') return shared[+raw] ?? raw;
  return raw;
}

const raw = JSON.parse(fs.readFileSync('web/public/data/bd-sheet.json', 'utf8'));
const t = transformBdSheet(raw);
const map = buildCellMap(t.cells, t.headerRows);
const mass = bdMassCol(t);
const l1 = bdSubsystemL1Col(t);
const l2 = bdSubsystemL2Col(t);
const sh = t.sectionHeaderRows;
const canon = t.canonicalSectionByLabel;

let excelHasWebEmpty = 0;
let jsonHasWebEmpty = 0;
const samples = [];
for (let r = 6; r <= t.lastRow; r++) {
  const ex = excelV(r);
  const cell = map.get(`${r}:${mass}`);
  const jsonV = cell?.v != null && cell.v !== '' ? String(cell.v) : '';
  const web = displayCellValue(map, r, mass, sh, canon, l1, l2);
  if (ex && !web) {
    excelHasWebEmpty++;
    if (samples.length < 12)
      samples.push({ r, excel: ex, jsonV, f: cell?.f?.slice(0, 30) });
  }
  if (jsonV && !web) {
    jsonHasWebEmpty++;
    if (samples.length < 20)
      samples.push({ r, jsonV, excel: ex, web });
  }
}
console.log('mass col', mass);
console.log('excel has value, web empty', excelHasWebEmpty);
console.log('json has value, web empty', jsonHasWebEmpty);
console.log('samples', samples);
