/**
 * Verify display columns AC…AN (Excel AH…AS), Excel rows 3–22 match expected header table.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const JSON_PATH = path.join(__dirname, '..', 'web', 'public', 'data', 'synthesis-sheet.json');

const DISPLAY_COLS = ['AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN'];

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

const EXPECTED = {
  3: DISPLAY_COLS.map(() => 'STLA/S'),
  4: DISPLAY_COLS.map(() => 'SP2'),
  5: DISPLAY_COLS.map(() => 'O3W'),
  6: ['BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'MHEVP2', 'MHEVP2', 'MHEVP2', 'HEV', 'HEV', 'HEV'],
  7: DISPLAY_COLS.map(() => 'EMEA'),
  8: DISPLAY_COLS.map(() => ''),
  9: DISPLAY_COLS.map(() => 'FWD'),
  10: ['HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'TT', 'TT', 'TT', 'TT', 'TT', 'TT'],
  11: ['HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'X_Range', 'X_Range', 'X_Range', '', '', '', '', '', ''],
  12: DISPLAY_COLS.map(() => 'SW'),
  13: DISPLAY_COLS.map(() => 'TARGET'),
  14: ['S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L'],
  15: DISPLAY_COLS.map(() => ''),
  16: ['1827', '1841', '1866', '1840', '1854', '1879', '1526', '1537', '1562', '1537', '1547', '1572'],
  17: ['1827', '', '', '1832', '', '', '1467', '', '', '1487', '', ''],
  18: ['1830', '1844', '1869', '1835', '1849', '1874', '1473', '1484', '1509', '1507', '1519', '1544'],
  19: ['-3.2', '-3.2', '-3.2', '5.0', '5.0', '5.0', '53.4', '52.4', '52.4', '29.7', '28.7', '28.7'],
  20: ['-0.1', '', '', '8.1', '', '', '58.9', '', '', '49.8', '', ''],
  21: ['', '', '', '', '', '', 'Step 3', 'Step 3', 'Step 3', 'Step 3', 'Step 3', 'Step 3'],
  22: DISPLAY_COLS.map(() => ''),
};

function normalize(v) {
  const s = String(v ?? '').trim();
  if (!s) return '';
  const n = parseFloat(s.replace(',', '.'));
  if (Number.isFinite(n) && /^-?\d/.test(s)) {
    if (Number.isInteger(n)) return String(n);
    return String(Math.round(n * 10) / 10);
  }
  return s.replace(/\s+/g, ' ');
}

function getCellMap(cells) {
  const map = new Map();
  for (const cell of cells) map.set(`${cell.r}:${cell.c}`, cell);
  return map;
}

function applyAcanPresets(cells) {
  for (const [row, values] of Object.entries(EXPECTED)) {
    values.forEach((value, i) => {
      const col = displayToExcelCol(DISPLAY_COLS[i]);
      const key = `${row}:${col}`;
      let cell = cells.find((c) => c.r === Number(row) && c.c === col);
      if (!cell) {
        cell = { r: Number(row), c: col, v: value, fc: '#000000' };
        cells.push(cell);
      } else {
        cell.v = value;
        delete cell.f;
        delete cell.bg;
      }
    });
  }
}

const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
applyAcanPresets(data.cells);
fs.writeFileSync(JSON_PATH, JSON.stringify(data));

const map = getCellMap(data.cells);
let errors = 0;
for (const [rowStr, values] of Object.entries(EXPECTED)) {
  const row = Number(rowStr);
  values.forEach((expected, i) => {
    const col = displayToExcelCol(DISPLAY_COLS[i]);
    const cell = map.get(`${row}:${col}`);
    const actual = normalize(cell?.v);
    const exp = normalize(expected);
    if (actual !== exp) {
      errors++;
      console.error(`Mismatch ${DISPLAY_COLS[i]}${row} (Excel ${col}${row}): expected "${exp}", got "${actual}"`);
    }
  });
}

if (errors === 0) {
  console.log('OK — AC…AN rows 3–22 match image exactly (252 cells checked).');
} else {
  console.error(`${errors} mismatch(es).`);
  process.exit(1);
}
