/**
 * Live SUMPRODUCT for SYNTHESIS ↔ BD (vehicle columns, filter rows 3–14).
 * Same pattern as Excel: mass in BD!V, match BD!AS to F{row}, filters per vehicle column.
 */
const BD_END_ROW = 3480;
const FILTER_ROWS = [
  ['A', 3],
  ['B', 4],
  ['C', 5],
  ['D', 7],
  ['E', 6],
  ['F', 8],
  ['G', 9],
  ['I', 13],
  ['K', 11],
  ['L', 12],
  ['O', 14],
];
function parseNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
}
/** Column-oriented sparse arrays indexed by Excel row. */
export function buildBdColumnIndex(bdRaw) {
  const cols = {};
  const add = (col, row, v) => {
    if (!cols[col]) cols[col] = [];
    cols[col][row] = v;
  };
  for (const cell of bdRaw.cells || []) {
    add(cell.c, cell.r, cell.v);
  }
  for (const [rowStr, rowCols] of Object.entries(bdRaw.headerRows || {})) {
    const row = parseInt(rowStr, 10);
    for (const [col, cell] of Object.entries(rowCols)) {
      add(col, row, cell.v);
    }
  }
  return cols;
}
function columnMatches(bdVal, crit) {
  const v = String(bdVal ?? '').trim();
  const c = String(crit ?? '').trim();
  if (c === 'TT') return true;
  if (!c) return v === 'TT' || v === '';
  return v === c || v === 'TT';
}
function rowMatchesFilters(bdCols, vehCol, synRow, getSynCell, getBdValue) {
  for (const [bdCol, filterRow] of FILTER_ROWS) {
    const crit = getSynCell(filterRow, vehCol);
    const bdVal = getBdValue
      ? getBdValue(synRow, bdCol)
      : bdCols[bdCol]?.[synRow];
    if (!columnMatches(bdVal, crit)) return false;
  }
  const q = String(
    (getBdValue ? getBdValue(synRow, 'Q') : bdCols.Q?.[synRow]) ?? ''
  ).trim();
  if (q !== 'S' && q !== '') return false;
  return true;
}
export function computeSumproduct(
  bdCols,
  synRow,
  vehCol,
  getSynCell,
  getBdValue
) {
  const label = String(getSynCell(synRow, 'F') ?? '').trim();
  const weights = bdCols.V;
  if (!weights) return 0;
  let sum = 0;
  const end = Math.min(BD_END_ROW, 4000);
  for (let r = 2; r <= end; r++) {
    const as = String(bdCols.AS?.[r] ?? '').trim();
    if (label && as !== label) continue;
    if (!rowMatchesFilters(bdCols, vehCol, r, getSynCell, getBdValue)) continue;
    const mass = getBdValue
      ? getBdValue(r, 'V')
      : weights[r];
    sum += parseNum(mass);
  }
  return sum;
}
export function isSumproductCell(cell) {
  return Boolean(cell?.f && /SUMPRODUCT/i.test(cell.f));
}
