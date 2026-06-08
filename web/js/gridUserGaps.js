/**
 * User-inserted blank rows/columns — display-only gaps (no cell/formula remapping).
 */

export const USER_COL_GAP_PREFIX = '__UGAP_';

export function isUserGapCol(col) {
  return typeof col === 'string' && col.startsWith(USER_COL_GAP_PREFIX);
}

export function userGapColId(col) {
  if (!isUserGapCol(col)) return null;
  return col.slice(USER_COL_GAP_PREFIX.length);
}

export function userGapColAnchor(sheet, gapCol) {
  const id = userGapColId(gapCol);
  if (!id) return null;
  const gap = userColGapsForSheet(sheet).find((g) => g.id === id);
  return gap ? gap.afterCol : null;
}

export function cloneUserGapsSnapshot(raw) {
  return {
    rowGaps: userRowGapsForSheet(raw).map((g) => ({ ...g })),
    colGaps: userColGapsForSheet(raw).map((g) => ({ ...g })),
  };
}

/** Synthesis must not use display-only gaps — strip if present in saved session. */
export function clearUserGaps(raw) {
  if (!raw) return;
  raw.userRowGaps = [];
  raw.userColGaps = [];
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

/** Inject white display rows after anchored Excel rows. */
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
