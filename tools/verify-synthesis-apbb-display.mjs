/**
 * Verify display columns AP…BB, display rows 3–23 (same row numbers as the grid).
 * Display row 19 = blank spacer between Excel 18 and 19; display 20–23 = Excel 19–22.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { transformSynthesisSheet } from '../web/js/sheetTransform.js';
import { buildCellMap, getCell } from '../web/js/bdStore.js';
import { synDisplayValue, buildSynPillarColumns } from '../web/js/synStore.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const JSON_PATH = path.join(__dirname, '..', 'web', 'public', 'data', 'synthesis-sheet.json');

const COLS = ['AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB'];

const IMAGE_BY_DISPLAY_ROW = {
  3: COLS.map(() => 'STLA/S'),
  4: COLS.map(() => 'SP2'),
  5: COLS.map(() => 'P3S'),
  6: [
    'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV',
    'MHEVP2', 'MHEVP2', 'HEV', 'HEV',
  ],
  7: COLS.map(() => 'EMEA'),
  8: ['', '', 'SBW', '', 'SBW', '', 'SBW', 'SBW', 'SBW', '', '', '', ''],
  9: [
    'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'AWD',
    'FWD', 'FWD', 'FWD', 'FWD',
  ],
  10: [
    'HR', 'HR', 'HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'XR',
    'TT', 'TT', 'TT', 'TT',
  ],
  11: [
    'HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'HIGH_Range',
    'X_Range', 'X_Range', 'X_Range', 'X_Range',
    '', '', '', '',
  ],
  12: COLS.map(() => ''),
  13: COLS.map(() => 'TARGET'),
  14: [
    'L2', 'L3', 'L3', 'GT', 'GT', 'L3', 'L3', 'GT', 'GTI',
    'L2', 'L3', 'L3', 'GT',
  ],
  15: COLS.map(() => ''),
  16: [
    '1806', '1822', '1834', '1711', '1724', '1835', '1848', '1738', '1876',
    '1497', '1516', '1489', '1359',
  ],
  17: ['1801', '', '', '', '', '1823', '', '', '2000', '1415', '', '1425', ''],
  18: [
    '1809', '1825', '1837', '1839', '1851', '1830', '1843', '1857', '1989',
    '1444', '1463', '1485', '1504',
  ],
  19: COLS.map(() => ''),
  20: [
    '-2.7', '-3.3', '-3.3', '-127.2', '-127.2', '4.9', '4.9', '-119', '-112.5',
    '53.9', '52.3', '4.4', '-144.5',
  ],
  21: ['4.9', '', '', '', '', '12.3', '', '', '-123.8', '82.4', '', '63.9', ''],
  22: ['', '', '', '', '', '', '', '', '', 'Step 3', 'Step 3', 'Step 3', 'Step 3'],
  23: COLS.map(() => ''),
};

const HIDDEN_OFFSET = 5;
function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}
function numToCol(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function displayToExcelCol(displayCol) {
  return numToCol(colToNum(displayCol) + HIDDEN_OFFSET);
}
function displayRowToExcel(displayRow) {
  if (displayRow <= 18) return displayRow;
  if (displayRow === 19) return null;
  if (displayRow >= 20 && displayRow <= 23) return displayRow - 1;
  return null;
}

function normalize(v) {
  const s = String(v ?? '').trim();
  if (!s) return '';
  const n = parseFloat(s.replace(',', '.'));
  if (Number.isFinite(n) && /^-?\d/.test(s)) {
    const rounded = Math.round(n * 10) / 10;
    return Number.isInteger(rounded) ? String(rounded) : String(rounded);
  }
  return s.replace(/\s+/g, ' ').trim();
}

const raw = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
const sheet = transformSynthesisSheet(raw);
const cellMap = buildCellMap(sheet.cells, sheet.headerRows);
const pillarColumns = buildSynPillarColumns(sheet, cellMap);

let errors = [];
for (let dr = 3; dr <= 23; dr++) {
  const er = displayRowToExcel(dr);
  const expected = IMAGE_BY_DISPLAY_ROW[dr];
  for (let i = 0; i < COLS.length; i++) {
    const col = COLS[i];
    const exp = normalize(expected[i]);
    if (er === null) {
      if (exp !== '') errors.push({ dr, col, exp, actual: '(gap)' });
      continue;
    }
    const excelCol = displayToExcelCol(col);
    const cell = getCell(cellMap, er, excelCol);
    const shown = normalize(synDisplayValue(cell, cellMap, er, excelCol, sheet, pillarColumns));
    if (shown !== exp) {
      errors.push({ dr, col, excel: `${excelCol}${er}`, exp, actual: shown });
    }
  }
}

if (errors.length) {
  console.error(`${errors.length} mismatch(es):`);
  for (const e of errors.slice(0, 40)) {
    console.error(
      `  L${e.dr} ${e.col}${e.excel ? ' (' + e.excel + ')' : ''}: attendu "${e.exp}", affiché "${e.actual}"`
    );
  }
  if (errors.length > 40) console.error(`  … et ${errors.length - 40} autres`);
  process.exit(1);
}

console.log('OK — 273 cellules AP…BB (lignes 3–23): STLA/S, SP2, P3S, … conformes.');
