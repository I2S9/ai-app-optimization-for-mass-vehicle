/**
 * Output for CDC column W (Curb Mass) ↔ Synthesis row 16.
 * Each CDC data row mirrors one Synthesis vehicle column (rows transposed from columns).
 */
import { displayToExcelCol } from './synthesisPerf.js';
import {
  synMaaExcelCols,
  synAcanExcelCols,
  synApbbExcelCols,
} from './synthesisCalc.js?v=grid-perf2';

export const CDC_CURB_COL = 'W';

/** Data rows carrying a vehicle (skip header/spacer bands). */
export const CDC_CURB_ROW_SKIP = new Set([1, 2, 3, 4, 5, 6, 120, 121, 122, 123]);

function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function numToCol(n) {
  let s = '';
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

/** Excel columns for a Synthesis summary table (display start…end). */
export function synTableExcelCols(displayStart, displayEnd) {
  const out = [];
  for (let d = colToNum(displayStart); d <= colToNum(displayEnd); d++) {
    out.push(displayToExcelCol(numToCol(d)));
  }
  return out;
}

/**
 * CDC vehicle blocks → matching Synthesis table columns (row index within block
 * maps to column index). Row counts are aligned with table widths where possible;
 * extra CDC rows without a matching column stay blank.
 */
export const CDC_SYN_CURB_BLOCKS = [
  { from: 7, to: 21, label: 'M…AA (O3H)', getSynCols: synMaaExcelCols },
  { from: 22, to: 33, label: 'AC…AN (O3W)', getSynCols: synAcanExcelCols },
  { from: 34, to: 50, label: 'AP…BB (P3S)', getSynCols: synApbbExcelCols },
  { from: 51, to: 62, label: 'BD…BO (P3W)', getSynCols: () => synTableExcelCols('BD', 'BO') },
  { from: 63, to: 79, label: 'CI…CY (P3U)', getSynCols: () => synTableExcelCols('CI', 'CY') },
  { from: 80, to: 92, label: 'BS…CE (P3H)', getSynCols: () => synTableExcelCols('BS', 'CE') },
  { from: 93, to: 104, label: 'EF…EQ (J1X)', getSynCols: () => synTableExcelCols('EF', 'EQ') },
  { from: 105, to: 119, label: 'DR…ED (J2U)', getSynCols: () => synTableExcelCols('DR', 'ED') },
  { from: 124, to: 140, label: 'CI…CY (P3H SP2)', getSynCols: () => synTableExcelCols('CI', 'CY') },
  { from: 141, to: 156, label: 'FJ…FZ (P3W SP2)', getSynCols: () => synTableExcelCols('FJ', 'FZ') },
  { from: 157, to: 169, label: 'AP…BB (O3H SP2)', getSynCols: synApbbExcelCols },
  { from: 170, to: 181, label: 'AC…AN (O3W SP2)', getSynCols: synAcanExcelCols },
  { from: 182, to: 194, label: 'DA…DP (A3H)', getSynCols: () => synTableExcelCols('DA', 'DP') },
];

export function cdcCurbBlockForRow(row) {
  const r = Number(row);
  if (!Number.isFinite(r)) return null;
  return CDC_SYN_CURB_BLOCKS.find((b) => r >= b.from && r <= b.to) || null;
}

/** Synthesis Excel column for CDC row (positional index within the block's table). */
export function cdcRowToSynExcelCol(row) {
  if (CDC_CURB_ROW_SKIP.has(Number(row))) return null;
  const block = cdcCurbBlockForRow(row);
  if (!block) return null;
  const cols = block.getSynCols();
  const idx = Number(row) - block.from;
  if (idx < 0 || idx >= cols.length) return null;
  return cols[idx];
}

export function isCdcCurbLinkedCell(row, col) {
  return col === CDC_CURB_COL && cdcRowToSynExcelCol(row) != null;
}

/** Human-readable link target for tooltips. */
export function cdcCurbLinkTitle(row) {
  const block = cdcCurbBlockForRow(row);
  const synCol = cdcRowToSynExcelCol(row);
  if (!block || !synCol) return '';
  return `Lié à Synthesis — Curb mass ligne 16, table ${block.label}, colonne ${synCol}`;
}
