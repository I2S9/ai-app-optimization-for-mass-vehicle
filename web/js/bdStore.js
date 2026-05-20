/** In-memory BD sheet model (loaded once from static JSON). */

export function buildCellMap(cells) {
  const map = new Map();
  for (const cell of cells) {
    map.set(`${cell.r}:${cell.c}`, cell);
  }
  return map;
}

export function buildWidthMap(colWidths, columns) {
  const map = new Map();
  for (const w of colWidths || []) {
    map.set(w.col, w.width);
  }
  const defaultW = 72;
  for (const col of columns) {
    if (!map.has(col)) map.set(col, defaultW);
  }
  return map;
}

export function getCell(map, row, col) {
  return map.get(`${row}:${col}`);
}

export function displayValue(cell) {
  if (!cell) return '';
  if (cell.v != null && cell.v !== '') return String(cell.v);
  if (cell.f) return cell.f.startsWith('=') ? cell.f : cell.f;
  return '';
}

export function isReadonlyCell(cell, row, dataStartRow) {
  if (row < dataStartRow) return true;
  if (!cell) return false;
  return Boolean(cell.f);
}

export function isAddBlueRow(map, row) {
  const a = getCell(map, row, 'A');
  return a?.v === '_ADDBLUE';
}
