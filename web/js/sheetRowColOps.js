/**
 * Insert blank row/column and copy/paste entire row or column on raw sheets.
 */
import { colToIndex, indexToCol } from './formulaUtil.js';

export function shiftFormulaRows(formula, firstShiftRow, delta) {
  if (!formula || delta === 0) return formula;
  return String(formula).replace(
    /(\$?)([A-Z]{1,3})(\$?)(\d+)/gi,
    (m, d1, col, d2, row) => {
      const r = parseInt(row, 10);
      if (r < firstShiftRow) return m;
      return `${d1}${col}${d2}${r + delta}`;
    }
  );
}

export function shiftFormulaColumns(formula, firstShiftIdx, delta) {
  if (!formula || delta === 0) return formula;
  return String(formula).replace(
    /(\$?)([A-Z]{1,3})(\$?)(\d+)/gi,
    (m, d1, col, d2, row) => {
      const idx = colToIndex(col.toUpperCase());
      if (idx < firstShiftIdx) return m;
      return `${d1}${indexToCol(idx + delta)}${d2}${row}`;
    }
  );
}

function remapSheetCol(col, afterIdx, delta) {
  const idx = colToIndex(col);
  if (idx <= afterIdx) return col;
  return indexToCol(idx + delta);
}

function invalidateRawIndex(raw) {
  if (raw) delete raw._cellIndex;
}

function cloneCellPayload(cell) {
  const out = {};
  if (cell.v != null && String(cell.v) !== '') out.v = String(cell.v);
  if (cell.f) out.f = String(cell.f);
  if (cell.userEdited) out.userEdited = true;
  return out;
}

/** Insert a blank row below `afterRow` (Excel row index). */
export function insertBlankRowAt(sheet, afterRow) {
  const insertAt = Number(afterRow) + 1;
  if (!Number.isFinite(insertAt) || insertAt < 1) return sheet;

  const cells = (sheet.cells || []).map((cell) => {
    const entry = { ...cell };
    if (entry.r >= insertAt) {
      entry.r += 1;
      if (entry.f) entry.f = shiftFormulaRows(entry.f, insertAt, 1);
    }
    return entry;
  });

  const headerRows = {};
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const r = parseInt(rowStr, 10);
    const nr = r >= insertAt ? r + 1 : r;
    const row = {};
    for (const [col, cell] of Object.entries(cols)) {
      const entry = { ...cell };
      if (entry.f) entry.f = shiftFormulaRows(entry.f, insertAt, 1);
      row[col] = entry;
    }
    headerRows[String(nr)] = row;
  }

  const deletedRows = (sheet.deletedRows || [])
    .map(Number)
    .filter(Number.isFinite)
    .map((r) => (r >= insertAt ? r + 1 : r));

  const matrixColors = {};
  for (const [oldR, color] of Object.entries(sheet.matrixColors || {})) {
    const r = parseInt(oldR, 10);
    matrixColors[r >= insertAt ? r + 1 : r] = color;
  }

  const merges = (sheet.merges || []).map((m) => {
    const entry = { ...m };
    if (entry.startRow >= insertAt) entry.startRow += 1;
    if (entry.endRow >= insertAt) entry.endRow += 1;
    return entry;
  });

  const lastRow = Math.max(sheet.lastRow || 0, sheet.finRow || 0, insertAt - 1) + 1;

  const out = {
    ...sheet,
    cells,
    headerRows,
    deletedRows,
    matrixColors,
    merges,
    lastRow,
    finRow: lastRow,
  };
  invalidateRawIndex(out);
  return out;
}

/** Insert a blank column to the right of `afterCol`. */
export function insertBlankColumnAt(sheet, afterCol) {
  const afterIdx = colToIndex(String(afterCol));
  const delta = 1;
  const newCol = indexToCol(afterIdx + 1);

  const oldColumns = sheet.columns || [];
  const columns = [];
  let inserted = false;
  for (const col of oldColumns) {
    const idx = colToIndex(col);
    if (idx <= afterIdx) {
      columns.push(col);
      if (idx === afterIdx && !inserted) {
        columns.push(newCol);
        inserted = true;
      }
    } else {
      columns.push(remapSheetCol(col, afterIdx, delta));
    }
  }
  if (!inserted) columns.push(newCol);

  const remapKey = (obj) => {
    const out = {};
    for (const [col, val] of Object.entries(obj || {})) {
      out[remapSheetCol(col, afterIdx, delta)] = val;
    }
    return out;
  };

  const headers = remapKey(sheet.headers);
  const headersRaw = remapKey(sheet.headersRaw || sheet.headers);
  if (headers[newCol] == null) headers[newCol] = '';
  if (headersRaw[newCol] == null) headersRaw[newCol] = '';

  const colWidths = (sheet.colWidths || []).map((w) => ({
    ...w,
    col: remapSheetCol(w.col, afterIdx, delta),
  }));

  const headerRows = {};
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const row = {};
    for (const [col, cell] of Object.entries(cols)) {
      const nc = remapSheetCol(col, afterIdx, delta);
      const entry = { ...cell };
      if (entry.f) entry.f = shiftFormulaColumns(entry.f, afterIdx + 1, delta);
      row[nc] = entry;
    }
    headerRows[rowStr] = row;
  }

  const cells = (sheet.cells || []).map((cell) => {
    const nc = remapSheetCol(cell.c, afterIdx, delta);
    const entry = { ...cell, c: nc };
    if (entry.f) entry.f = shiftFormulaColumns(entry.f, afterIdx + 1, delta);
    return entry;
  });

  const out = {
    ...sheet,
    columns,
    headers,
    headersRaw,
    colWidths,
    headerRows,
    cells,
  };
  invalidateRawIndex(out);
  return out;
}

export function collectRowCells(sheet, row) {
  const r = Number(row);
  const out = [];
  const seen = new Set();
  for (const cell of sheet.cells || []) {
    if (Number(cell.r) !== r) continue;
    if (!cell.v && !cell.f) continue;
    out.push({ c: cell.c, ...cloneCellPayload(cell) });
    seen.add(cell.c);
  }
  const hdr = sheet.headerRows?.[String(r)] || sheet.headerRows?.[r];
  if (hdr) {
    for (const [col, cell] of Object.entries(hdr)) {
      if (seen.has(col) || (!cell.v && !cell.f)) continue;
      out.push({ c: col, ...cloneCellPayload(cell) });
    }
  }
  return out;
}

export function collectColCells(sheet, col) {
  const c = String(col);
  const out = [];
  const seen = new Set();
  for (const cell of sheet.cells || []) {
    if (cell.c !== c || (!cell.v && !cell.f)) continue;
    out.push({ r: cell.r, ...cloneCellPayload(cell) });
    seen.add(cell.r);
  }
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const r = parseInt(rowStr, 10);
    if (seen.has(r)) continue;
    const cell = cols[c];
    if (!cell || (!cell.v && !cell.f)) continue;
    out.push({ r, ...cloneCellPayload(cell) });
  }
  return out;
}

export function setRawCellPayload(raw, row, col, payload) {
  if (!raw) return;
  invalidateRawIndex(raw);
  const r = Number(row);
  const c = String(col);
  const hasContent = Boolean(payload && (payload.v != null || payload.f));

  let cell = (raw.cells || []).find((x) => x.r === r && x.c === c);
  if (!hasContent) {
    if (cell) raw.cells = raw.cells.filter((x) => x !== cell);
    const hdr =
      (raw.headerRows && raw.headerRows[String(r)]) ||
      (raw.headerRows && raw.headerRows[r]);
    if (hdr && hdr[c]) delete hdr[c];
    return;
  }

  if (!cell) {
    cell = { r, c };
    if (!raw.cells) raw.cells = [];
    raw.cells.push(cell);
  }
  if (payload.v != null) cell.v = String(payload.v);
  else delete cell.v;
  if (payload.f) cell.f = String(payload.f);
  else delete cell.f;
  cell.userEdited = true;
}

export function clearRowCells(raw, row) {
  const r = Number(row);
  raw.cells = (raw.cells || []).filter((c) => Number(c.r) !== r);
  if (raw.headerRows) {
    delete raw.headerRows[String(r)];
    delete raw.headerRows[r];
  }
  invalidateRawIndex(raw);
}

export function clearColCells(raw, col) {
  const c = String(col);
  raw.cells = (raw.cells || []).filter((cell) => cell.c !== c);
  for (const cols of Object.values(raw.headerRows || {})) {
    if (cols[c]) delete cols[c];
  }
  invalidateRawIndex(raw);
}

export function pasteRowCells(raw, targetRow, payloads) {
  clearRowCells(raw, targetRow);
  for (const p of payloads || []) {
    if (!p.c) continue;
    setRawCellPayload(raw, targetRow, p.c, p);
  }
}

export function pasteColCells(raw, targetCol, payloads) {
  clearColCells(raw, targetCol);
  for (const p of payloads || []) {
    if (p.r == null) continue;
    setRawCellPayload(raw, p.r, targetCol, p);
  }
}
