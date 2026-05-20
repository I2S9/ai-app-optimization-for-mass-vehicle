/** Prepare BD sheet JSON for the web grid (columns, English labels & values). */
import { computeOutlineRows, computeSectionHeaderRows } from './bdStore.js';
import {
  HEADER_FR_EN,
  translateValue,
} from './bdTranslate.js';
const HIDDEN_COLUMNS = new Set([
  'AI',
  'AJ',
  'AK',
  'AL',
  'AM',
  'AN',
  'AT',
  'AV',
  'Z', // Source
  'AA', // Pack
  'AB', // Reference
]);
const HEADER_EN = {
  AP: 'Sub-system L1',
  AQ: 'Sub-system L1 ref',
  AS: 'Sub-system L2',
  AT: 'Sub-system L2 ref',
  AU: 'Sub-System Design Dpt',
};
function translateCellValue(v) {
  if (v == null || v === '') return v;
  return translateValue(String(v));
}
/** First row where column A is FIN (end of usable data). */
export function findFinRow(sheet) {
  let finRow = null;
  const match = (r, col, v) => {
    if (col !== 'A' || v == null) return;
    if (String(v).trim().toUpperCase() !== 'FIN') return;
    if (finRow == null || r < finRow) finRow = r;
  };
  for (const cell of sheet.cells || []) match(cell.r, cell.c, cell.v);
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const r = parseInt(rowStr, 10);
    if (cols.A) match(r, 'A', cols.A.v);
  }
  return finRow;
}
/** Drop empty tail rows after the FIN marker. */
export function trimSheetAfterFin(sheet) {
  const finRow = findFinRow(sheet);
  if (finRow == null) return sheet;
  const cells = (sheet.cells || []).filter((c) => c.r <= finRow);
  const headerRows = {};
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const r = parseInt(rowStr, 10);
    if (r <= finRow) headerRows[rowStr] = cols;
  }
  return { ...sheet, lastRow: finRow, finRow, cells, headerRows };
}
export function transformBdSheet(sheet) {
  sheet = trimSheetAfterFin(sheet);
  const columns = sheet.columns.filter((c) => !HIDDEN_COLUMNS.has(c));
  const headers = { ...sheet.headers };
  for (const col of [...HIDDEN_COLUMNS]) {
    delete headers[col];
  }
  headers.AU = 'Sub-System Design Dpt';
  for (const col of columns) {
    const raw = headers[col];
    if (HEADER_EN[col]) headers[col] = HEADER_EN[col];
    else if (raw && HEADER_FR_EN[raw]) headers[col] = HEADER_FR_EN[raw];
    else if (raw) headers[col] = translateValue(raw);
  }
  const cells = sheet.cells
    .filter((c) => !HIDDEN_COLUMNS.has(c.c))
    .map((c) => {
      if (c.v == null || c.v === '') return c;
      const v = translateCellValue(String(c.v));
      return v === c.v ? c : { ...c, v };
    });
  const headerRows = {};
  for (const [row, cols] of Object.entries(sheet.headerRows || {})) {
    headerRows[row] = {};
    for (const [col, cell] of Object.entries(cols)) {
      if (HIDDEN_COLUMNS.has(col)) continue;
      let v = cell.v;
      if (v != null && v !== '') v = translateCellValue(String(v));
      headerRows[row][col] = { ...cell, v };
    }
  }
  const colWidths = (sheet.colWidths || []).filter((w) => !HIDDEN_COLUMNS.has(w.col));
  const prepared = {
    ...sheet,
    columns,
    headers,
    cells,
    headerRows,
    colWidths,
  };
  const { rows, canonicalByLabel } = computeSectionHeaderRows(prepared);
  prepared.sectionHeaderRows = rows;
  prepared.canonicalSectionByLabel = Object.fromEntries(canonicalByLabel);
  prepared.outlineRows = computeOutlineRows(prepared).filter(
    (r) => r <= prepared.lastRow
  );
  return prepared;
}

/** Synthesis grid: config rows + vehicle comparison columns. */
export function transformSynthesisSheet(sheet) {
  const columns = [...(sheet.columns || [])];
  const headers = { ...(sheet.headers || {}) };
  for (const col of columns) {
    const raw = headers[col];
    if (raw && HEADER_FR_EN[raw]) headers[col] = HEADER_FR_EN[raw];
    else if (raw) headers[col] = translateValue(String(raw));
  }
  const cells = (sheet.cells || []).map((c) => {
    if (c.v == null || c.v === '') return c;
    const v = translateCellValue(String(c.v));
    return v === c.v ? c : { ...c, v };
  });
  const headerRows = {};
  for (const [row, cols] of Object.entries(sheet.headerRows || {})) {
    headerRows[row] = {};
    for (const [col, cell] of Object.entries(cols)) {
      let v = cell.v;
      if (v != null && v !== '') v = translateCellValue(String(v));
      headerRows[row][col] = { ...cell, v };
    }
  }
  return {
    ...sheet,
    columns,
    headers,
    cells,
    headerRows,
    dataStartRow: sheet.dataStartRow || 3,
    sectionHeaderRows: new Set(),
    outlineRows: [],
  };
}
