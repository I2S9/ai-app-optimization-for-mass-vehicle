/**
 * Verify display columns AC…AN, display rows 3–23 against reference image.
 * Display row 19 = blank spacer between Excel 18 and 19; display 20–23 = Excel 19–22.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const JSON_PATH = path.join(__dirname, '..', 'web', 'public', 'data', 'synthesis-sheet.json');

const COLS = ['AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN'];

const IMAGE_BY_DISPLAY_ROW = {
  3: COLS.map(() => 'STLA/S'),
  4: COLS.map(() => 'SP2'),
  5: COLS.map(() => 'O3W'),
  6: ['BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'MHEVP2', 'MHEVP2', 'MHEVP2', 'HEV', 'HEV', 'HEV'],
  7: COLS.map(() => 'EMEA'),
  8: COLS.map(() => ''),
  9: COLS.map(() => 'FWD'),
  10: ['HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'TT', 'TT', 'TT', 'TT', 'TT', 'TT'],
  11: ['HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'X_Range', 'X_Range', 'X_Range', '', '', '', '', '', ''],
  12: COLS.map(() => 'SW'),
  13: COLS.map(() => 'TARGET'),
  14: ['S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L'],
  15: COLS.map(() => ''),
  16: ['1827', '1841', '1866', '1840', '1854', '1879', '1526', '1537', '1562', '1537', '1547', '1572'],
  17: ['1827', '', '', '1832', '', '', '1467', '', '', '1487', '', ''],
  18: ['1830', '1844', '1869', '1835', '1849', '1874', '1473', '1484', '1509', '1507', '1519', '1544'],
  19: COLS.map(() => ''),
  20: ['-3.2', '-3.2', '-3.2', '5.0', '5.0', '5.0', '53.4', '52.4', '52.4', '29.7', '28.7', '28.7'],
  21: ['-0.1', '', '', '8.1', '', '', '58.9', '', '', '49.8', '', ''],
  22: ['', '', '', '', '', '', 'Step 3', 'Step 3', 'Step 3', 'Step 3', 'Step 3', 'Step 3'],
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

const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
const map = new Map(data.cells.map((c) => [`${c.r}:${c.c}`, c.v ?? '']));

let errors = [];
for (let dr = 3; dr <= 23; dr++) {
  const er = displayRowToExcel(dr);
  const expected = IMAGE_BY_DISPLAY_ROW[dr];
  for (let i = 0; i < COLS.length; i++) {
    const col = COLS[i];
    const exp = normalize(expected[i]);
    if (er === null) {
      if (exp !== '') errors.push({ dr, col, exp, actual: '(gap row)' });
      continue;
    }
    const excelCol = displayToExcelCol(col);
    const actual = normalize(map.get(`${er}:${excelCol}`));
    if (actual !== exp) {
      errors.push({ dr, col, excel: `${excelCol}${er}`, exp, actual });
    }
  }
}

if (errors.length) {
  console.error(`${errors.length} mismatch(es):`);
  for (const e of errors) {
    console.error(
      `  L${e.dr} ${e.col}${e.excel ? ' (Excel ' + e.excel + ')' : ''}: attendu "${e.exp}", trouvé "${e.actual}"`
    );
  }
  process.exit(1);
}

console.log('OK — 264 cellules AC…AN (lignes affichage 3–23) conformes à l’image.');
