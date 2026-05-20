/** BD sheet model: row types, colors, outline filter, display. */

import {
  translateValue,
  isLabelColumn,
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

function isUnassignedSectionLabel(v) {
  const t = String(v || '').trim();
  return (
    /^-Non affecté$/i.test(t) ||
    /^_Non affecté$/i.test(t) ||
    t === '-Unassigned' ||
    t === '_Unassigned'
  );
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

export function isOutlineRow(map, row, sectionHeaderRows) {
  return isStructureRow(map, row, sectionHeaderRows);
}

/** Outline = separator + yellow L1 sections + blue L2 sub-sections. */
export function computeOutlineRows(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const sectionRows = asSectionRowSet(
    sheet.sectionHeaderRows ||
      computeSectionHeaderRows(sheet).rows
  );
  const rows = [];
  for (let r = 2; r <= sheet.lastRow; r++) {
    if (
      isSeparatorRow(r) ||
      sectionRows.has(r) ||
      isSubSectionRow(map, r, sectionRows)
    )
      rows.push(r);
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

export function isRecopierRow(map, row, sectionHeaderRows) {
  return getRowLabel(map, row, sectionHeaderRows) === 'To copy';
}

export function isDataGreenColA(map, row, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  const a = displayValue(getCell(map, row, 'A'));
  return a === 'STLA/S' || a === 'STLA-S';
}

export function rowStyleClass(map, row, sectionHeaderRows) {
  if (isSeparatorRow(row)) return 'row-separator';
  if (isSectionRow(map, row, sectionHeaderRows)) return 'row-section';
  if (isSubSectionRow(map, row, sectionHeaderRows)) return 'row-subsection';
  if (isRecopierRow(map, row, sectionHeaderRows)) return 'row-recopier';
  return '';
}

export function cellInlineStyle(cell, map, row, col, sectionHeaderRows) {
  const style = {};
  if (cell?.bg) style.backgroundColor = cell.bg;
  if (cell?.fc) style.color = cell.fc;

  const cls = rowStyleClass(map, row, sectionHeaderRows);
  if (col === 'A' && isDataGreenColA(map, row, sectionHeaderRows)) {
    style.backgroundColor = '#c6efce';
  }
  if (col === 'A' && cls === 'row-subsection') {
    style.color = '#c00000';
    style.fontWeight = 'bold';
  }
  if (col === 'A' && cls === 'row-section') {
    style.color = '#000';
    style.fontWeight = 'bold';
  }
  return style;
}

/** Hide TT / project fillers on structure rows except nomination columns. */
function maskStructureValue(map, row, col, v, sectionHeaderRows) {
  if (!isStructureRow(map, row, sectionHeaderRows)) return v;
  if (isLabelColumn(col)) return v;
  if (v === 'TT' || v === '0') return '';
  if (row === 5 && col !== 'A' && col !== 'W') return '';
  return '';
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
    return v;
  }
  return '';
}

export function displayCellValue(
  map,
  row,
  col,
  sectionHeaderRows,
  canonicalSectionByLabel
) {
  const canonicalByLabel = canonicalByLabelMap(canonicalSectionByLabel);
  const cell = getCell(map, row, col);
  let v = displayValue(cell);
  if (col === 'AS' && cell && (cell.v == null || cell.v === '') && cell.f) v = '';
  v = maskStructureValue(map, row, col, v, sectionHeaderRows);
  v = maskDuplicateSectionLabel(map, row, col, v, canonicalByLabel);
  v = maskInheritedSubSectionLabel(map, row, col, v, sectionHeaderRows);
  v = maskRepeatedUnassigned(map, row, col, v);

  // Yellow L1: title in frozen column A.
  if (col === 'A' && isSectionRow(map, row, sectionHeaderRows)) {
    const title = getRowLabel(map, row, sectionHeaderRows, canonicalByLabel);
    if (title) return title;
  }

  // Blue L2 (_ADDBLUE…): title in frozen column A (Excel: A7=_ADDBLUE).
  if (col === 'A' && isSubSectionRow(map, row, sectionHeaderRows)) {
    const title = getSubSectionLabel(map, row);
    if (title) return title;
  }

  if (col !== 'A' || v) return v;
  return '';
}
