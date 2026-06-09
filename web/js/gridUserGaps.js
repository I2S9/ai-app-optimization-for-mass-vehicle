/**
 * User-inserted blank rows/columns — editable overlay (no cell/formula remapping).
 */

export const USER_COL_GAP_PREFIX = '__UGAP_';
export const USER_ROW_GAP_NAV_PREFIX = '__UGR_';

export function isUserGapCol(col) {
  return typeof col === 'string' && col.startsWith(USER_COL_GAP_PREFIX);
}

export function userGapColId(col) {
  if (!isUserGapCol(col)) return null;
  return col.slice(USER_COL_GAP_PREFIX.length);
}

export function userGapRowNavKey(gapId) {
  return `${USER_ROW_GAP_NAV_PREFIX}${gapId}`;
}

export function isUserGapNavRow(row) {
  return typeof row === 'string' && row.startsWith(USER_ROW_GAP_NAV_PREFIX);
}

export function userGapRowIdFromNav(row) {
  if (!isUserGapNavRow(row)) return null;
  return row.slice(USER_ROW_GAP_NAV_PREFIX.length);
}

export function isUserGapCellAddress(row, col) {
  return isUserGapNavRow(row) || isUserGapCol(col);
}

export function userGapColAnchor(sheet, gapCol) {
  const id = userGapColId(gapCol);
  if (!id) return null;
  const gap = userColGapsForSheet(sheet).find((g) => g.id === id);
  return gap ? gap.afterCol : null;
}

function normalizeUserGapCells(raw) {
  if (!raw || !Array.isArray(raw.userGapCells)) return [];
  return raw.userGapCells
    .map((c) => {
      const out = { v: c.v != null ? String(c.v) : '' };
      if (c.gapRowId) out.gapRowId = String(c.gapRowId);
      if (c.gapColId) out.gapColId = String(c.gapColId);
      if (c.col) out.col = String(c.col);
      if (c.row != null && Number.isFinite(Number(c.row))) out.row = Number(c.row);
      return out;
    })
    .filter((c) => c.v !== '' && (c.gapRowId || c.gapColId));
}

export function userGapCellsForSheet(sheet) {
  return normalizeUserGapCells(sheet);
}

function gapCellMatches(cell, addr) {
  if (addr.gapRowId && cell.gapRowId !== addr.gapRowId) return false;
  if (addr.gapColId && cell.gapColId !== addr.gapColId) return false;
  if (addr.col && cell.col !== addr.col) return false;
  if (addr.row != null && cell.row !== addr.row) return false;
  return true;
}

/** Resolve storage address for a user-gap row/col intersection. */
export function resolveUserGapCellAddr(row, col) {
  const gapRowId = isUserGapNavRow(row) ? userGapRowIdFromNav(row) : null;
  const gapColId = isUserGapCol(col) ? userGapColId(col) : null;
  if (!gapRowId && !gapColId) return null;
  const addr = {};
  if (gapRowId) addr.gapRowId = gapRowId;
  if (gapColId) addr.gapColId = gapColId;
  if (gapRowId && !gapColId) addr.col = String(col);
  if (gapColId && !gapRowId) addr.row = Number(row);
  return addr;
}

export function getUserGapCellValue(sheet, row, col) {
  const addr = resolveUserGapCellAddr(row, col);
  if (!addr) return undefined;
  const hit = userGapCellsForSheet(sheet).find((c) => gapCellMatches(c, addr));
  return hit ? hit.v : '';
}

export function setUserGapCell(sheet, row, col, value) {
  const addr = resolveUserGapCellAddr(row, col);
  if (!addr || !sheet) return false;
  if (!sheet.userGapCells) sheet.userGapCells = [];
  const v = value != null ? String(value) : '';
  const idx = sheet.userGapCells.findIndex((c) => gapCellMatches(c, addr));
  if (v === '') {
    if (idx >= 0) sheet.userGapCells.splice(idx, 1);
  } else if (idx >= 0) {
    sheet.userGapCells[idx].v = v;
  } else {
    sheet.userGapCells.push({ ...addr, v });
  }
  return true;
}

/** Column width = average of the anchor column and its right neighbor. */
export function userGapColAverageWidth(sheet, gapCol, getWidth, fallback = 72) {
  const anchor = userGapColAnchor(sheet, gapCol);
  if (!anchor) return fallback;
  const leftW = getWidth(anchor) ?? fallback;
  const cols = sheet.columns || [];
  const anchorIdx = cols.indexOf(anchor);
  const rightCol =
    anchorIdx >= 0 && anchorIdx < cols.length - 1 ? cols[anchorIdx + 1] : null;
  const rightW = rightCol ? (getWidth(rightCol) ?? leftW) : leftW;
  return Math.round((leftW + rightW) / 2);
}

export function gridRowKeyFromEntry(entry) {
  if (!entry) return null;
  if (entry.userGap && entry.gapKey) return userGapRowNavKey(entry.gapKey);
  return entry.excelRow != null ? entry.excelRow : null;
}

export function cloneUserGapsSnapshot(raw) {
  return {
    rowGaps: userRowGapsForSheet(raw).map((g) => ({ ...g })),
    colGaps: userColGapsForSheet(raw).map((g) => ({ ...g })),
    gapCells: userGapCellsForSheet(raw).map((c) => ({ ...c })),
  };
}

/** Strip user gaps (e.g. legacy corrupted session snapshots). */
export function clearUserGaps(raw) {
  if (!raw) return;
  raw.userRowGaps = [];
  raw.userColGaps = [];
  raw.userGapCells = [];
}

function normalizeUserRowGaps(raw) {
  if (!raw || !Array.isArray(raw.userRowGaps)) return [];
  return raw.userRowGaps
    .map((g) => ({
      id: String(g.id || ''),
      afterExcelRow: Number(g.afterExcelRow),
    }))
    .filter((g) => g.id && Number.isFinite(g.afterExcelRow));
}

function normalizeUserColGaps(raw) {
  if (!raw || !Array.isArray(raw.userColGaps)) return [];
  return raw.userColGaps
    .map((g) => ({
      id: String(g.id || ''),
      afterCol: String(g.afterCol || ''),
    }))
    .filter((g) => g.id && g.afterCol);
}

export function userRowGapsForSheet(sheet) {
  return normalizeUserRowGaps(sheet);
}

export function userColGapsForSheet(sheet) {
  return normalizeUserColGaps(sheet);
}

export function gapsAfterExcelRow(sheet, excelRow) {
  const r = Number(excelRow);
  return userRowGapsForSheet(sheet).filter((g) => g.afterExcelRow === r);
}

export function addUserRowGap(raw, afterExcelRow) {
  if (!raw) return;
  const row = Number(afterExcelRow);
  if (!Number.isFinite(row)) return;
  if (!raw.userRowGaps) raw.userRowGaps = [];
  raw.userRowGaps.push({
    id: `ur${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
    afterExcelRow: row,
  });
}

export function addUserColGap(raw, afterCol) {
  if (!raw) return;
  const col = String(afterCol || '');
  if (!col || isUserGapCol(col)) return;
  if (!raw.userColGaps) raw.userColGaps = [];
  raw.userColGaps.push({
    id: `uc${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
    afterCol: col,
  });
}

/** Inject blank editable rows after anchored Excel rows. */
export function injectUserRowGaps(rows, sheet, nextDisplayRow) {
  let displayRow = nextDisplayRow;
  const out = [];
  for (const entry of rows) {
    out.push(entry);
    displayRow = entry.displayRow != null ? entry.displayRow + 1 : displayRow;
    const anchor = entry.excelRow;
    if (anchor == null) continue;
    for (const gap of gapsAfterExcelRow(sheet, anchor)) {
      out.push({
        userGap: true,
        gapKey: gap.id,
        afterExcelRow: anchor,
        excelRow: null,
        displayRow: displayRow++,
      });
    }
  }
  return out;
}

export function expandColumnsWithUserGaps(columns, sheet) {
  const gaps = userColGapsForSheet(sheet);
  if (!gaps.length) return columns || [];
  const byAfter = new Map();
  for (const g of gaps) {
    if (!byAfter.has(g.afterCol)) byAfter.set(g.afterCol, []);
    byAfter.get(g.afterCol).push(g);
  }
  const out = [];
  for (const col of columns || []) {
    out.push(col);
    const inserted = byAfter.get(col);
    if (inserted) {
      for (const g of inserted) {
        out.push(`${USER_COL_GAP_PREFIX}${g.id}`);
      }
    }
  }
  return out;
}

export function isSynUserGapEntry(entry) {
  return Boolean(entry && entry.userGap);
}
