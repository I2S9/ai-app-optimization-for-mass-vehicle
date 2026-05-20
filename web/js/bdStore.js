/** In-memory BD sheet model (loaded once from static JSON). */

export function buildCellMap(cells, headerRows = {}) {
  const map = new Map();
  for (const cell of cells) {
    map.set(`${cell.r}:${cell.c}`, cell);
  }
  for (const [rowStr, cols] of Object.entries(headerRows)) {
    const row = parseInt(rowStr, 10);
    if (Number.isNaN(row)) continue;
    for (const [col, c] of Object.entries(cols)) {
      const key = `${row}:${col}`;
      const existing = map.get(key);
      map.set(key, {
        r: row,
        c: col,
        v: c.v ?? existing?.v,
        f: c.f ?? existing?.f,
        bg: c.s?.backgroundColor ?? existing?.bg,
      });
    }
  }
  return map;
}

export function buildWidthMap(colWidths, columns, headers = {}) {
  const map = new Map();
  for (const w of colWidths || []) {
    map.set(w.col, w.width);
  }
  for (const col of columns) {
    const label = String(headers[col] || col);
    const fromHeader = Math.ceil(label.length * 7.5 + 24);
    const fromExcel = map.get(col) ?? 72;
    const wide = Math.max(fromExcel, fromHeader, 56);
    map.set(col, Math.min(wide, 280));
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

/** Label for template / section rows (column A when empty). */
export function getRowLabel(map, row) {
  const a = displayValue(getCell(map, row, 'A'));
  if (a) return a;
  if (row === 5) return '-ADAPTATION';
  for (const col of ['AS', 'AP', 'W']) {
    const v = displayValue(getCell(map, row, col));
    if (v?.startsWith('_')) return v;
  }
  return '';
}

export function isReadonlyCell(cell, row, displayStartRow) {
  if (row < displayStartRow) return true;
  if (!cell) return false;
  return Boolean(cell.f);
}

/** Row 5 — orange "-ADAPTATION" band (Excel); not row 6+ copies in AP/W. */
export function isAdaptationRow(map, row) {
  return row === 5;
}

/** Row 4 — dark blue separator. */
export function isSeparatorRow(row) {
  return row === 4;
}

/** Row 7+ — teal section headers (_ADDBLUE, _CARBURANT, …). */
export function isSectionHeaderRow(map, row) {
  if (isAdaptationRow(map, row) || isSeparatorRow(row)) return false;
  const label = getRowLabel(map, row);
  if (!label.startsWith('_')) return false;
  return true;
}

export function isAddBlueRow(map, row) {
  const label = getRowLabel(map, row);
  return label === '_ADDBLUE' || label === 'ADDBLUE';
}

export function isRecopierRow(map, row) {
  return getRowLabel(map, row) === 'A recopier';
}

/** Data rows with STLA/S in column A — light green first column. */
export function isDataGreenColA(map, row) {
  const a = displayValue(getCell(map, row, 'A'));
  return a === 'STLA/S' || a === 'STLA-S';
}

export function rowStyleClass(map, row) {
  if (isSeparatorRow(row)) return 'row-separator';
  if (isAdaptationRow(map, row)) return 'row-adaptation';
  if (isAddBlueRow(map, row) || isSectionHeaderRow(map, row)) return 'row-section';
  if (isRecopierRow(map, row)) return 'row-recopier';
  return '';
}

export function cellInlineStyle(cell, map, row, col) {
  const style = {};
  if (cell?.bg) style.backgroundColor = cell.bg;
  if (cell?.fc) style.color = cell.fc;

  if (col === 'A' && isDataGreenColA(map, row)) {
    style.backgroundColor = '#c6efce';
  }
  if (col === 'A' && (isSectionHeaderRow(map, row) || isAddBlueRow(map, row))) {
    style.color = '#c00000';
    style.fontWeight = 'bold';
  }
  if (col === 'A' && isAdaptationRow(map, row)) {
    style.fontWeight = 'bold';
    style.color = '#000000';
  }
  return style;
}

/** Column A shows section labels when Excel leaves A empty. */
export function displayCellValue(map, row, col) {
  const cell = getCell(map, row, col);
  const v = displayValue(cell);
  if (col !== 'A' || v) return v;
  const label = getRowLabel(map, row);
  if (label) return label;
  if (row === 5) return '-ADAPTATION';
  return '';
}
