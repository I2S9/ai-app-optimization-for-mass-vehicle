/** BD sheet model: row types, colors, outline filter, display. */
import {
  translateValue,
  SECTION_ALLOWLIST,
} from './bdTranslate.js';
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
    let wide = Math.max(fromExcel, fromHeader, 56);
    if (col === 'AE') wide = Math.max(wide, 220);
    map.set(col, Math.min(wide, col === 'AE' ? 320 : 280));
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
/** Sub-system L2 label from AS (never surface formula text as a label). */
export function getAsLabel(map, row) {
  const cell = getCell(map, row, 'AS');
  if (!cell?.v || cell.v === '') return '';
  return String(cell.v);
}
/** Blue sub-section title (_ADDBLUE, _FUEL…) from column A or AS. */
export function getSubSectionLabel(map, row) {
  const a = displayValue(getCell(map, row, 'A'));
  if (a.startsWith('_') && !isUnassignedSectionLabel(a)) return a.trim();
  const as = getAsLabel(map, row);
  if (as.startsWith('_') && !isUnassignedSectionLabel(as)) return as.trim();
  return '';
}
const SKIP_LABELS = new Set([
  'TT',
  'STLA/S',
  'STLA-S',
  'PTF',
  'STATUS',
  'TARGET',
  'FWD',
  'AWD',
  'BEV',
  'HEV',
  'SP1',
  'SP2',
  'END',
  'FIN',
]);
/** Columns B–P: project / silhouette / vehicle filters (Excel row 5 + data rows). */
export const PROJECT_COLS = new Set([
  'B',
  'C',
  'D',
  'E',
  'F',
  'G',
  'H',
  'I',
  'J',
  'K',
  'L',
  'M',
  'N',
  'O',
  'P',
]);
const ENERGY_VALUES = new Set([
  'BEV',
  'HEV',
  'MHEVP2',
  'PHEV',
  'ICE',
  'MHEV',
  'PHEV2',
]);
function isUnassignedSectionLabel(v) {
  const t = String(v || '').trim();
  return (
    /^-Non affecté$/i.test(t) ||
    /^_Non affecté$/i.test(t) ||
    t === '-Unassigned' ||
    t === '_Unassigned'
  );
}
/** Blue band label: leading "_" + full uppercase (L2 and component title rows). */
export function formatBlueBandLabel(label) {
  const t = String(label || '').trim();
  if (!t || isUnassignedSectionLabel(t)) return t;
  const withPrefix = t.startsWith('_') ? t : `_${t}`;
  return withPrefix.toUpperCase();
}
/** @deprecated use formatBlueBandLabel */
export function ensureSubSectionPrefix(label) {
  return formatBlueBandLabel(label);
}
/** Yellow section title — must match the official section list. */
export function isSectionLabel(v) {
  const t = String(v || '').trim();
  if (!t) return false;
  if (isUnassignedSectionLabel(t)) return true;
  return SECTION_ALLOWLIST.has(t);
}
/** @deprecated */
export function isSubsystemL1Label(v) {
  return isSectionLabel(v);
}
/** Yellow CA chapter bands: -ADAPTATION (row 5) and -ADTH (row 139). */
export function isCaBandRow(map, row) {
  return row === 5 || row === 139;
}
export function isSectionBandRow(map, row) {
  return isCaBandRow(map, row);
}
function asSectionRowSet(sectionHeaderRows) {
  if (sectionHeaderRows instanceof Set) return sectionHeaderRows;
  return new Set(sectionHeaderRows || []);
}
/**
 * One yellow row per L1 section (column AP only + CA bands 5/139 + first _Unassigned).
 */
export function computeSectionHeaderRows(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const rows = new Set();
  const canonicalByLabel = new Map();
  rows.add(5);
  rows.add(139);
  canonicalByLabel.set('-ADAPTATION', 5);
  canonicalByLabel.set('-ADTH', 139);
  let lastAp = '';
  for (let r = sheet.dataStartRow || 6; r <= sheet.lastRow; r++) {
    const ap = displayValue(getCell(map, r, 'AP'));
    if (!ap) continue;
    const label = String(ap).trim();
    if (isSectionLabel(label) && label !== lastAp) {
      rows.add(r);
      canonicalByLabel.set(label, r);
      lastAp = label;
    }
  }
  let lastAs = '';
  for (let r = sheet.dataStartRow || 6; r <= sheet.lastRow; r++) {
    const as = getAsLabel(map, r);
    if (isUnassignedSectionLabel(as) && as !== lastAs) {
      rows.add(r);
      canonicalByLabel.set(as, r);
      lastAs = as;
    }
  }
  return { rows, canonicalByLabel };
}
export function isSectionRow(map, row, sectionHeaderRows) {
  if (sectionHeaderRows) return asSectionRowSet(sectionHeaderRows).has(row);
  if (isCaBandRow(map, row)) return true;
  return false;
}
/** @deprecated alias */
export function isSubSystemL1Row(map, row) {
  return isSectionRow(map, row) && !isCaBandRow(map, row);
}
/** Blue row: first row of each L2 block (_ADDBLUE, _FUEL…) per Excel. */
export function isSubSectionRow(map, row, sectionHeaderRows) {
  if (isSectionRow(map, row, sectionHeaderRows)) return false;
  const label = getSubSectionLabel(map, row);
  if (!label) return false;
  return label !== getSubSectionLabel(map, row - 1);
}
export function isSeparatorRow(row) {
  return row === 4;
}
/** Yellow or blue structure row (no project/silhouette data). */
export function isStructureRow(map, row, sectionHeaderRows) {
  return (
    isSeparatorRow(row) ||
    isSectionRow(map, row, sectionHeaderRows) ||
    isSubSectionRow(map, row, sectionHeaderRows)
  );
}
/** Row 5: project filter legend (TT, BEV, …) under -ADAPTATION. */
export function isProjectLegendRow(row) {
  return row === 5;
}
/** Vehicle header before STLA/S lines (SPC, TARGET, mass in V…). */
export function isProjectConfigRow(map, row, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  if (isDataGreenColA(map, row, sectionHeaderRows)) return false;
  if (isProjectLegendRow(row)) return true;
  const i = displayValue(getCell(map, row, 'I'));
  if (i === 'TARGET' || i === 'STATUS') return true;
  const b = displayValue(getCell(map, row, 'B'));
  if (b && b !== 'TT' && b !== '0' && !String(b).startsWith('_')) return true;
  const c = displayValue(getCell(map, row, 'C'));
  if (c && c !== 'TT' && c !== '0') return true;
  return false;
}
function isSignificantDataRow(map, row, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  const v = displayValue(getCell(map, row, 'V'));
  if (v && v !== '0' && !String(v).startsWith('#')) return true;
  const s = displayValue(getCell(map, row, 'S'));
  if (s && s !== '0' && !String(s).startsWith('#')) return true;
  const ae = displayValue(getCell(map, row, 'AE'));
  if (ae && ae !== '0') return true;
  return isProjectConfigRow(map, row, sectionHeaderRows);
}
export function shouldDisplayBodyRow(map, row, sheet) {
  const dataStart = sheet.dataStartRow || 6;
  const sh = sheet.sectionHeaderRows;
  if (row < dataStart && !isStructureRow(map, row, sh)) return false;
  if (isStructureRow(map, row, sh)) return true;
  if (colATitle(map, row)) return true;
  if (isSignificantDataRow(map, row, sh)) return true;
  return false;
}
/** Body rows for the grid — skips untitled white lines; consecutive display numbers. */
export function computeBodyDisplayRows(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const rows = [];
  let displayRow = 1;
  for (let r = 2; r <= sheet.lastRow; r++) {
    if (!shouldDisplayBodyRow(map, r, sheet)) continue;
    rows.push({ excelRow: r, displayRow: displayRow++ });
  }
  return rows;
}
export function getRowLabel(map, row, sectionHeaderRows, canonicalSectionByLabel) {
  if (!isSectionRow(map, row, sectionHeaderRows)) {
    const sub = getSubSectionLabel(map, row);
    if (sub) return sub;
    const a = displayValue(getCell(map, row, 'A'));
    if (a) return a;
    return '';
  }
  if (row === 5) return '-ADAPTATION';
  if (row === 139) return '-ADTH';
  const ap = displayValue(getCell(map, row, 'AP'));
  if (ap && isSectionLabel(ap)) return ap;
  const as = getAsLabel(map, row);
  if (isUnassignedSectionLabel(as)) return as;
  if (isCaBandRow(map, row)) {
    const w = displayValue(getCell(map, row, 'W'));
    if (w?.startsWith('-')) return w;
  }
  return '';
}
export function isReadonlyCell(cell, row, dataStartRow) {
  if (row < 2) return true;
  if (row < dataStartRow) return true;
  if (!cell) return false;
  return Boolean(cell.f);
}
/** Raw title in frozen column A (Date). */
export function colATitle(map, row) {
  const a = displayValue(getCell(map, row, 'A'));
  return a ? String(a).trim() : '';
}
/** Yellow/blue band or other title in frozen column A. */
export function hasFrozenTitle(map, row, sheet) {
  const sectionHeaderRows = sheet.sectionHeaderRows;
  if (isStructureRow(map, row, sectionHeaderRows)) return true;
  return Boolean(colATitle(map, row));
}
export function isFinDeLotRow(map, row) {
  const t = colATitle(map, row);
  return /^Fin de Lot\b/i.test(t) || /^End of lot\b/i.test(t);
}
export function isRecopierRow(map, row, sectionHeaderRows) {
  const t = colATitle(map, row);
  return t === 'A recopier' || t === 'To copy';
}
/** STLA/S rows: green frozen Date column (Excel). */
export function isDataGreenColA(map, row, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  const a = colATitle(map, row);
  return a === 'STLA/S' || a === 'STLA-S';
}
/** Blue title row (full row): label in A only, no component data. */
export function shouldDateColBlue(map, row, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  if (isFinDeLotRow(map, row)) return false;
  if (isRecopierRow(map, row, sectionHeaderRows)) return false;
  if (isDataGreenColA(map, row, sectionHeaderRows)) return false;
  return Boolean(colATitle(map, row));
}
export function rowStyleClass(map, row, sectionHeaderRows) {
  if (isSeparatorRow(row)) return 'row-separator';
  if (isSectionRow(map, row, sectionHeaderRows)) return 'row-section';
  if (isSubSectionRow(map, row, sectionHeaderRows)) return 'row-subsection';
  if (isFinDeLotRow(map, row)) return 'row-fin-lot';
  if (isRecopierRow(map, row, sectionHeaderRows)) return 'row-recopier';
  if (isDataGreenColA(map, row, sectionHeaderRows)) return 'row-date-green';
  if (isProjectConfigRow(map, row, sectionHeaderRows)) return 'row-project-config';
  if (shouldDateColBlue(map, row, sectionHeaderRows)) return 'row-date-blue';
  return '';
}
/** Semantic fill for project columns (SPC, TT, BEV, TARGET…). */
export function projectCellClass(displayText, col) {
  if (!PROJECT_COLS.has(col) || !displayText) return '';
  const v = String(displayText).trim().toUpperCase();
  if (!v) return '';
  if (v === 'TT') return 'cell-proj-tt';
  if (v === 'TARGET') return 'cell-proj-target';
  if (v === 'STATUS') return 'cell-proj-status';
  if (v === 'SPC' || /^SP\d+$/.test(v) || v === 'SP1' || v === 'SP2') return 'cell-proj-spc';
  if (ENERGY_VALUES.has(v)) return 'cell-proj-energy';
  if (v === 'FWD' || v === 'AWD' || v === 'RWD') return 'cell-proj-drivetrain';
  if (v === 'PTF') return 'cell-proj-ptf';
  return 'cell-proj-value';
}
/** Title-only markers (blue / fin de lot / to copy) — no project data. */
export function isTitleMarkerRow(map, row, sectionHeaderRows) {
  return (
    shouldDateColBlue(map, row, sectionHeaderRows) ||
    isFinDeLotRow(map, row) ||
    isRecopierRow(map, row, sectionHeaderRows)
  );
}
/** Outline view (eye): yellow L1 + all blue bands (L2 and component titles). */
export function isOutlineRow(map, row, sectionHeaderRows) {
  return (
    isStructureRow(map, row, sectionHeaderRows) ||
    shouldDateColBlue(map, row, sectionHeaderRows)
  );
}
/** Outline = yellow sections + every blue-highlighted row. */
export function computeOutlineRows(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const sectionRows = asSectionRowSet(
    sheet.sectionHeaderRows ||
      computeSectionHeaderRows(sheet).rows
  );
  const rows = [];
  for (let r = 2; r <= sheet.lastRow; r++) {
    if (isOutlineRow(map, r, sectionRows)) rows.push(r);
  }
  return rows;
}
export function cellInlineStyle(cell, map, row, col, sectionHeaderRows) {
  const style = {};
  if (cell?.bg) style.backgroundColor = cell.bg;
  if (cell?.fc) style.color = cell.fc;
  const cls = rowStyleClass(map, row, sectionHeaderRows);
  if (cls === 'row-date-green' || cls === 'row-project-config') {
    if (col === 'A' || PROJECT_COLS.has(col)) {
      if (col === 'A') style.backgroundColor = '#c6efce';
    }
  }
  if (col === 'A' && isDataGreenColA(map, row, sectionHeaderRows)) {
    style.backgroundColor = '#c6efce';
  }
  if (col === 'A' && (cls === 'row-subsection' || cls === 'row-date-blue')) {
    style.color = cls === 'row-subsection' ? '#c00000' : '#000';
    style.fontSize = '11px';
    style.fontWeight = '700';
  }
  if (col === 'A' && cls === 'row-section') {
    style.color = '#000';
    style.fontSize = '11px';
    style.fontWeight = '700';
  }
  return style;
}
/**
 * Yellow / blue bookmark rows: show section or sub-section title only, never data.
 * Returns null when the row is not a structure bookmark.
 */
function structureBookmarkDisplay(
  map,
  row,
  col,
  sectionHeaderRows,
  canonicalByLabel
) {
  if (!isStructureRow(map, row, sectionHeaderRows)) return null;
  if (isSeparatorRow(row)) {
    if (col === 'A') {
      const v = displayValue(getCell(map, row, 'A'));
      return v || '4';
    }
    return '';
  }
  if (isSectionRow(map, row, sectionHeaderRows)) {
    const title = getRowLabel(map, row, sectionHeaderRows, canonicalByLabel);
    if (!title) return '';
    if (col === 'A') return title;
    if (isCaBandRow(map, row)) {
      if (col === 'W') return title;
      return '';
    }
    if (col === 'AP' && canonicalByLabel.get(title) === row) return title;
    if (col === 'AS' && isUnassignedSectionLabel(title)) return title;
    return '';
  }
  if (isSubSectionRow(map, row, sectionHeaderRows)) {
    const title = formatBlueBandLabel(getSubSectionLabel(map, row));
    if (!title) return '';
    if (col === 'A' || col === 'AS') return title;
    return '';
  }
  return '';
}
/** Hide TT / 0 only on inherited filler rows — keep on STLA/S & project rows. */
function maskStructureValue(map, row, col, v, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return v;
  if (isProjectLegendRow(row)) return v;
  if (isDataGreenColA(map, row, sectionHeaderRows)) return v;
  if (isProjectConfigRow(map, row, sectionHeaderRows)) return v;
  if (v === 'TT' || v === '0') return '';
  return v;
}
function projectLegendDisplay(map, row, col) {
  if (!isProjectLegendRow(row) || !PROJECT_COLS.has(col)) return null;
  return displayValue(getCell(map, row, col)) || '';
}
function isInheritedBandOrSection(v) {
  if (!v) return false;
  const t = String(v).trim();
  return isSectionLabel(v) || (t.startsWith('-') && t.length > 1);
}
function canonicalByLabelMap(canonicalSectionByLabel) {
  if (!canonicalSectionByLabel) return new Map();
  if (canonicalSectionByLabel instanceof Map) return canonicalSectionByLabel;
  return new Map(Object.entries(canonicalSectionByLabel));
}
/** Section title visible on exactly one cell: W@5/139 or AP@canonical row. */
function isCanonicalSectionDisplay(map, row, col, v, canonicalByLabel) {
  if (!v || !isInheritedBandOrSection(v)) return false;
  const label = String(v).trim();
  if (isCaBandRow(map, row) && col === 'W') {
    return (
      (row === 5 && label === '-ADAPTATION') ||
      (row === 139 && label === '-ADTH')
    );
  }
  return canonicalByLabel.get(label) === row && col === 'AP';
}
/** Hide L1 / CA titles everywhere except their single canonical cell. */
function maskDuplicateSectionLabel(map, row, col, v, canonicalByLabel) {
  if (!v || !isInheritedBandOrSection(v)) return v;
  if (isCanonicalSectionDisplay(map, row, col, v, canonicalByLabel)) return v;
  return '';
}
/** _Unassigned only on the first row of the block (column AS). */
function maskRepeatedUnassigned(map, row, col, v) {
  if (col !== 'AS' || !isUnassignedSectionLabel(v)) return v;
  const prev = getAsLabel(map, row - 1);
  if (isUnassignedSectionLabel(prev) && String(prev).trim() === String(v).trim()) {
    return '';
  }
  return v;
}
/** L2 titles only on blue header rows — hide copies on data rows & duplicate AS when A shows it. */
function maskInheritedSubSectionLabel(map, row, col, v, sectionHeaderRows) {
  if (!v) return v;
  const t = String(v).trim();
  if (!t.startsWith('_') || isUnassignedSectionLabel(t)) return v;
  if (isSubSectionRow(map, row, sectionHeaderRows)) {
    if (col === 'AS') {
      const a = displayValue(getCell(map, row, 'A'));
      if (a && a.trim() === t) return '';
    }
    return formatBlueBandLabel(v);
  }
  return '';
}
function blueTitleDisplay(map, row, col, sectionHeaderRows) {
  if (!shouldDateColBlue(map, row, sectionHeaderRows) || col !== 'A') return null;
  const title = colATitle(map, row);
  if (!title) return '';
  return formatBlueBandLabel(title);
}
export function displayCellValue(
  map,
  row,
  col,
  sectionHeaderRows,
  canonicalSectionByLabel
) {
  const canonicalByLabel = canonicalByLabelMap(canonicalSectionByLabel);
  const legend = projectLegendDisplay(map, row, col);
  if (legend !== null) return legend;
  const bookmark = structureBookmarkDisplay(
    map,
    row,
    col,
    sectionHeaderRows,
    canonicalByLabel
  );
  if (bookmark !== null) return bookmark;
  const blueTitle = blueTitleDisplay(map, row, col, sectionHeaderRows);
  if (blueTitle !== null) return blueTitle;
  const cell = getCell(map, row, col);
  let v = displayValue(cell);
  if (col === 'AS' && cell && (cell.v == null || cell.v === '') && cell.f) v = '';
  v = maskStructureValue(map, row, col, v, sectionHeaderRows);
  v = maskDuplicateSectionLabel(map, row, col, v, canonicalByLabel);
  v = maskInheritedSubSectionLabel(map, row, col, v, sectionHeaderRows);
  v = maskRepeatedUnassigned(map, row, col, v);
  if (isTitleMarkerRow(map, row, sectionHeaderRows) && col !== 'A') {
    return '';
  }
  if (col !== 'A' || v) return v;
  return '';
}
