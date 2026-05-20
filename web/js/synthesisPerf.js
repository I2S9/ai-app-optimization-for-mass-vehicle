/** Excel columns A–E hidden in UI; F→A, G→B, … */
export const SYN_HIDDEN_COLS = new Set(['A', 'B', 'C', 'D', 'E']);
export const SYN_STICKY_COL = 'F';
const HIDDEN_OFFSET = 5;

export function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Excel column letter shown in the grid (F → A, G → B, …). */
export function excelToDisplayCol(excelCol) {
  const n = colToNum(excelCol) - HIDDEN_OFFSET;
  return n > 0 ? numToCol(n) : excelCol;
}

export function filterSynDisplayColumns(columns = []) {
  return columns.filter((c) => !SYN_HIDDEN_COLS.has(c));
}

export function isSynFilterEdit(row, col) {
  return row >= 3 && row <= 14;
}
