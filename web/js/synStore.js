/** SYNTHESIS sheet display helpers (filter band, labels, merges). */
import { displayValue, getCell, PROJECT_COLS } from './bdStore.js';

export const SYN_FILTER_ROWS = new Set([3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]);
export const SYN_LABEL_COL = 'F';
export const SYN_VEHICLE_COL_START = 'G';

/** Merges disabled in web grid (virtual scroll). */
export function buildMergeMaps() {
  return { master: new Map(), covered: new Set() };
}

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

export function isMergeCovered(mergeMaps, row, col) {
  return mergeMaps?.covered?.has(`${row}:${col}`) ?? false;
}

export function getMergeSpan(mergeMaps, row, col) {
  const m = mergeMaps?.master?.get(`${row}:${col}`);
  if (!m) return null;
  // Large vertical merges break virtualized tables — keep layout value in master cell only.
  if (m.rowspan > 3 || m.colspan > 8) return null;
  return { rowspan: m.rowspan, colspan: m.colspan };
}

export function synLabel(map, row) {
  const v = displayValue(getCell(map, row, SYN_LABEL_COL));
  if (v) return String(v).trim();
  const d = displayValue(getCell(map, row, 'D'));
  const t = d ? String(d).trim() : '';
  if (t.startsWith('-') || t.startsWith('_')) return t;
  return '';
}

export function isSynFilterRow(row, sheet) {
  const rows = sheet?.filterRows || [...SYN_FILTER_ROWS];
  return rows.includes(row);
}

export function isSynSeparatorRow(row) {
  return row === 4;
}

export function isSynSectionLabelRow(map, row) {
  const label = synLabel(map, row);
  if (!label) return false;
  return label.startsWith('-') || label.startsWith('_');
}

export function synRowStyleClass(map, row, sheet) {
  if (isSynSeparatorRow(row)) return 'syn-row-separator';
  if (isSynFilterRow(row, sheet)) return 'syn-row-filter';
  if (isSynSectionLabelRow(map, row)) {
    if (synLabel(map, row).startsWith('-')) return 'syn-row-section';
    return 'syn-row-subsection';
  }
  return 'syn-row-data';
}

/** Row CSS handles colours — skip per-cell bg for scroll performance. */
export function synCellInlineStyle(cell, map, row, col, sheet) {
  const style = {};
  if (col === SYN_LABEL_COL) style.textAlign = 'left';
  if (isSynFilterRow(row, sheet) && colToNum(col) >= colToNum(SYN_VEHICLE_COL_START)) {
    style.fontWeight = '600';
  }
  return style;
}

export function synProjectCellClass(displayText, col) {
  if (!PROJECT_COLS.has(col) || !displayText) return '';
  const v = String(displayText).trim().toUpperCase();
  if (!v) return '';
  if (v === 'TT') return 'cell-proj-tt';
  if (v === 'TARGET') return 'cell-proj-target';
  if (v === 'SPC' || /^SP\d+$/.test(v)) return 'cell-proj-spc';
  if (['BEV', 'HEV', 'MHEVP2', 'PHEV', 'ICE', 'MHEV', 'PHEV2'].includes(v)) {
    return 'cell-proj-energy';
  }
  if (v === 'HR' || v === 'XR' || v === 'SR') return 'cell-proj-range';
  return 'cell-proj-value';
}

export function synIsReadonly(cell, row, sheet) {
  if (row < 2) return true;
  if (!cell) return false;
  if (cell.f) return true;
  return false;
}

export function synDisplayValue(cell, map, row, col) {
  if (!cell) return '';
  if (col === SYN_LABEL_COL) {
    const label = synLabel(map, row);
    if (isSynSectionLabelRow(map, row) && label) return label;
  }
  return displayValue(cell);
}

/** Yellow section / blue subsection / separator rows (outline / eye view). */
export function isSynOutlineRow(map, row, sheet) {
  const cls = synRowStyleClass(map, row, sheet);
  return (
    cls === 'syn-row-section' ||
    cls === 'syn-row-subsection' ||
    cls === 'syn-row-separator'
  );
}

export function computeSynBodyRows(sheet, cellMap, outlineOnly = false) {
  const map = cellMap || new Map();
  const rows = [];
  let displayRow = 1;
  for (let r = 2; r <= sheet.lastRow; r++) {
    if (outlineOnly && !isSynOutlineRow(map, r, sheet)) continue;
    rows.push({ excelRow: r, displayRow: displayRow++ });
  }
  return rows;
}

export function synRowHeightPx(sheet, excelRow, defaultPx = 21) {
  const ht = sheet?.rowHeights?.[String(excelRow)] ?? sheet?.rowHeights?.[excelRow];
  return ht || defaultPx;
}
