/** Excel columns A–E hidden in UI; F→A, G→B, … */
export const SYN_HIDDEN_COLS = new Set(['A', 'B', 'C', 'D', 'E']);
export const SYN_STICKY_COL = 'F';
/** Display column A (Excel F) — sub-system labels. */
export const SYN_LABEL_COL_MIN_W = 240;
/** SP1 / SP2 TARGET pillar — wide column for large vertical label. */
export const SYN_PILLAR_COL_WIDTH = 72;

export function synStickyColWidth(sheet) {
  const fromSheet = sheet?.colWidths?.find((w) => w.col === SYN_STICKY_COL)?.width;
  return Math.max(fromSheet || SYN_LABEL_COL_MIN_W, SYN_LABEL_COL_MIN_W);
}

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
