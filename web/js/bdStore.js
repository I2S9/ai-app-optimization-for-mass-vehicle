/** BD sheet model: row types, colors, outline filter, display. */
import {
  translateValue,
  translateSubsystemLabel,
  repairClimateTypo,
  formatSectionDisplayLabel,
  formatCaBandStoredLabel,
  SECTION_ALLOWLIST,
} from './bdTranslate.js';
import {
  BD_CODIFICATION_COL,
  BD_FREE_FIELD_COL,
  BD_MASS_AV_AR_COLS,
  BD_POSITION_COLS,
  BD_PROJECT_COL,
  BD_SILHOUETTE_COL,
  BD_SUBSYSTEM_L1_COL,
  BD_SUBSYSTEM_L1_COL_RAW,
  BD_SUBSYSTEM_L2_COL,
  BD_SUBSYSTEM_L2_COL_RAW,
  BD_DESIGN_DEPT_COL,
  BD_DESIGN_DEPT_COL_RAW,
  BD_TITLE_COL,
  BD_MASS_COL,
  BD_CODIFICATION_WIDTH,
  BD_MODULAR_TYPE_COL,
  BD_TRADE_COL,
} from './bdColumnConfig.js';
import { colToIndex, indexToCol } from './formulaUtil.js';

/** Column letter for yellow L1 section titles (AP in JSON, AR after transform). */
export function bdSubsystemL1Col(sheet) {
  const headers = (sheet && sheet.headers) || {};
  for (const col of (sheet && sheet.columns) || Object.keys(headers)) {
    if (headers[col] === 'Sub-system L1') return col;
  }
  if (sheet && sheet.columns && sheet.columns.includes(BD_SUBSYSTEM_L1_COL_RAW)) {
    return BD_SUBSYSTEM_L1_COL_RAW;
  }
  return BD_SUBSYSTEM_L1_COL;
}
/** Column letter for L2 sub-section labels (_ADDBLUE, _FUEL…). */
export function bdSubsystemL2Col(sheet) {
  const headers = (sheet && sheet.headers) || {};
  for (const col of (sheet && sheet.columns) || Object.keys(headers)) {
    if (headers[col] === 'Sub-system L2') return col;
  }
  if (sheet && sheet.columns && sheet.columns.includes(BD_SUBSYSTEM_L2_COL_RAW)) {
    return BD_SUBSYSTEM_L2_COL_RAW;
  }
  return BD_SUBSYSTEM_L2_COL;
}
export function bdHeaderCol(sheet, headerLabel, fallback) {
  const headers = (sheet && sheet.headers) || {};
  for (const col of (sheet && sheet.columns) || Object.keys(headers)) {
    if (headers[col] === headerLabel) return col;
  }
  return fallback;
}
export function bdCodificationCol(sheet) {
  return bdHeaderCol(sheet, 'Codification', BD_CODIFICATION_COL);
}
export function bdTitleCol(sheet) {
  return bdHeaderCol(sheet, 'Title', BD_TITLE_COL);
}
export function bdMassCol(sheet) {
  return bdHeaderCol(sheet, 'Mass', BD_MASS_COL);
}
export function bdModularTypeCol(sheet) {
  return bdHeaderCol(sheet, 'Modular type', BD_MODULAR_TYPE_COL);
}
export function bdSilhouetteCol(sheet) {
  return bdHeaderCol(sheet, 'Silhouette', BD_SILHOUETTE_COL);
}
export function bdDesignDeptCol(sheet) {
  return bdHeaderCol(sheet, 'Sub-System Design Dpt', BD_DESIGN_DEPT_COL);
}
const EXCEL_ERROR_VALUES = new Set([
  '#NAME?',
  '#REF!',
  '#VALUE!',
  '#N/A',
  '#NULL!',
  '#DIV/0!',
  '#NUM!',
]);
export function stripExcelErrorValue(v) {
  if (v == null || v === '') return '';
  const t = String(v).trim();
  if (EXCEL_ERROR_VALUES.has(t) || (t.startsWith('#') && t.endsWith('!'))) return '';
  return t;
}
/** Columns C+ with alternating pale green / blue (not Codification / Title). */
export function isBdDataStripeCol(col, sheet) {
  if (col === 'A' || col === BD_PROJECT_COL) return false;
  const codif = bdCodificationCol(sheet);
  const title = bdTitleCol(sheet);
  if (col === codif || col === title) return false;
  const start = colToIndex(bdSilhouetteCol(sheet));
  return colToIndex(col) >= start;
}
export function bdColMetaClass(col, sheet) {
  if (
    col === bdCodificationCol(sheet) ||
    col === bdTitleCol(sheet) ||
    col === bdModularTypeCol(sheet)
  ) {
    return 'cell-col-gray';
  }
  if (isBdDataStripeCol(col, sheet)) return 'col-data-stripe';
  return '';
}
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
        v: c.v != null ? c.v : existing && existing.v,
        f: c.f != null ? c.f : existing && existing.f,
        bg:
          c.bg != null
            ? c.bg
            : c.s && c.s.backgroundColor != null
              ? c.s.backgroundColor
              : existing && existing.bg,
        fc:
          c.fc != null
            ? c.fc
            : c.s && c.s.color != null
              ? c.s.color
              : existing && existing.fc,
        b: c.b != null ? c.b : existing && existing.b,
      });
    }
  }
  return map;
}
/** Free-field column — long text; width from longest cell value. */
export function measureFreeFieldWidth(cells = []) {
  let maxLen = String('Free field').length;
  for (const cell of cells) {
    if (cell.c !== BD_FREE_FIELD_COL || cell.v == null || cell.v === '') continue;
    maxLen = Math.max(maxLen, String(cell.v).length);
  }
  return Math.min(720, Math.max(420, Math.ceil(maxLen * 6.5 + 32)));
}

export function buildWidthMap(
  colWidths,
  columns,
  headers = {},
  cells = [],
  freeFieldWidthCached
) {
  const map = new Map();
  const freeFieldWidth =
    freeFieldWidthCached != null ? freeFieldWidthCached : measureFreeFieldWidth(cells);
  for (const w of colWidths || []) {
    map.set(w.col, w.width);
  }
  for (const col of columns) {
    const label = String(headers[col] || col);
    const fromHeader = Math.ceil(label.length * 7.5 + 24);
    const fromExcel = map.has(col) ? map.get(col) : 72;
    let wide = Math.max(fromExcel, fromHeader, 56);
    if (col === BD_FREE_FIELD_COL) {
      map.set(col, Math.max(wide, freeFieldWidth));
      continue;
    }
    if (col === BD_CODIFICATION_COL) {
      map.set(col, Math.max(BD_CODIFICATION_WIDTH, fromHeader));
      continue;
    }
    if (col === BD_TITLE_COL) {
      map.set(col, Math.min(wide, 56));
      continue;
    }
    map.set(col, Math.min(wide, 280));
  }
  return map;
}
export function getCell(map, row, col) {
  return map.get(`${row}:${col}`);
}
export function displayValue(cell) {
  if (!cell) return '';
  if (cell.v != null && cell.v !== '') {
    const v = stripExcelErrorValue(String(cell.v));
    if (v) return v;
  }
  if (cell.f) return cell.f.startsWith('=') ? cell.f : cell.f;
  return '';
}

/** Excel formula text must never become a bookmark / section title. */
export function isFormulaLike(text) {
  const t = String(text || '').trim();
  if (!t) return false;
  if (t.startsWith('=')) return true;
  if (/^IF\s*\(/i.test(t)) return true;
  if (t.includes('#REF!') || t.includes('#NAME?')) return true;
  return false;
}

/** Display text for structure titles only (values, never formula strings). */
export function cellLabelValue(cell) {
  if (!cell) return '';
  if (cell.v != null && cell.v !== '') {
    const v = stripExcelErrorValue(String(cell.v));
    if (v && !isFormulaLike(v)) return v;
  }
  return '';
}

/**
 * Fixed canonical title for the two template CA chapter bands (rows 5 / 139).
 * These are structural chapters that must always match the Synthesis page, so
 * stray inline edits (e.g. "ADAPTATION TEST") are ignored in favour of the
 * canonical name.
 */
export function caBandCanonicalTitle(row) {
  if (row === 5) return 'ADAPTATION';
  if (row === 139) return 'CABIN CLIMATE TREATMENT SYSTEM';
  return '';
}

/**
 * Effective title for a CA chapter band (rows 5 / 139). A matrix rename is
 * persisted in column W (`-NEW NAME`); when present it wins so the chapter can
 * be renamed like any other section. When W is empty (untouched template) the
 * fixed canonical title is used, so the default display never regresses.
 */
export function caBandStoredTitle(map, row) {
  const w = cellLabelValue(getCell(map, row, 'W'));
  const stored = w ? formatSectionDisplayLabel(w) : '';
  return stored || caBandCanonicalTitle(row);
}

/** Yellow L1 / CA chapter title for matrix extract and Database row labels. */
export function getBdSectionTitleFromRow(map, row, sheet) {
  if (isCaBandRow(map, row)) {
    return caBandStoredTitle(map, row);
  }
  const l1Col = bdSubsystemL1Col(sheet);
  const ap = cellLabelValue(getCell(map, row, l1Col));
  if (ap) return translateSubsystemLabel(String(ap).trim());
  const a = cellLabelValue(getCell(map, row, 'A'));
  if (a && String(a).trim().startsWith('-')) {
    return translateSubsystemLabel(String(a).trim());
  }
  return '';
}

export function sanitizeStructureLabel(label) {
  const t = String(label || '').trim();
  if (!t || isFormulaLike(t)) return '';
  return t;
}
/** Sub-system L2 label (never surface formula text as a label). */
export function getAsLabel(map, row, l2Col = BD_SUBSYSTEM_L2_COL) {
  const cell = getCell(map, row, l2Col);
  if (!cell || cell.v == null || cell.v === '') return '';
  return String(cell.v);
}
/**
 * Blue sub-section title from the L2 sub-system column (Excel AS), then A.
 * A sub-section is the name of an L2 sub-system block. Most are prefixed with "_"
 * (_ADDBLUE, _FUEL…) but many sections (FENDERS, ALTERNATOR, SEATS, …) store the
 * block name WITHOUT the "_" (FRONT FENDERS, ALTERNATOR SYSTEM…). Both must be
 * recognised so the eye view / Bookmark Matrix list the SAME blue sub-sections
 * the Synthesis page shows. Section titles, "_Unassigned" and formulas are excluded.
 */
export function getSubSectionLabel(map, row, l2Col = BD_SUBSYSTEM_L2_COL) {
  const l2 = getAsLabel(map, row, l2Col).trim();
  // A genuine L2 block name (with or without "_"). Note we do NOT exclude values
  // that happen to match a section name (e.g. "FRONT BUMPER" used as the L2 block
  // under the FRONT section) — section header ROWS are already excluded upstream by
  // isSubSectionRow. We only drop terminal markers (END / FIN), "_Unassigned" and
  // formula text so they never appear as blue sub-sections.
  if (
    l2 &&
    !isUnassignedSectionLabel(l2) &&
    !isFormulaLike(l2) &&
    !SKIP_LABELS.has(l2.toUpperCase())
  ) {
    return l2;
  }
  const a = displayValue(getCell(map, row, 'A')).trim();
  if (a.startsWith('_') && !isUnassignedSectionLabel(a)) return a;
  // Data-entry gap: some components fill column A but leave the L2 cell empty
  // (e.g. "Vitre de toit" under HEADLINER / SR). Synthesis lists them as blue
  // sub-sections, so the Database eye view must too. Only accept a *real* label
  // that sits inside a named L1 section — this skips markers (To copy / End of
  // lot / STLA/S / formulas) and the trailing "not yet classified" zone (blank L1).
  if (
    !l2 &&
    a &&
    !isUnassignedSectionLabel(a) &&
    !isFormulaLike(a) &&
    !SKIP_LABELS.has(a.toUpperCase()) &&
    !/^-?[\d\s.,%]+$/.test(a) &&
    !isRecopierRow(map, row) &&
    !isFinDeLotRow(map, row) &&
    !isFormulesRow(map, row)
  ) {
    const l1Idx = colToIndex(l2Col) - 3;
    const l1Col = l1Idx >= 0 ? indexToCol(l1Idx) : null;
    const l1 = l1Col ? displayValue(getCell(map, row, l1Col)).trim() : '';
    // A === L1 is the signature of a yellow section header row (e.g. "STARTER"):
    // never treat those as a blue sub-section, or the real first sub-section right
    // below would be swallowed as a duplicate.
    if (
      l1 &&
      l1 !== '0' &&
      !isUnassignedSectionLabel(l1) &&
      a.toUpperCase() !== l1.toUpperCase()
    ) {
      return a;
    }
  }
  return '';
}
/**
 * Canonical (translated) sub-section key for band comparisons. A matrix rename
 * only rewrites the band's FIRST L2 cell; inherited data rows keep the OLD raw
 * value. Comparing raw labels then makes the next data row look like a NEW
 * sub-section (duplicate blue band). Compare canonical labels so a renamed band
 * and its inherited rows resolve to the same key → single band, no duplicates.
 */
export function canonicalSubSectionKey(label) {
  if (!label) return '';
  const t = repairClimateTypo(String(label).trim());
  if (!t) return '';
  return String(translateValue(t) || t)
    .trim()
    .toUpperCase();
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
/** Rows hidden from the grid (Excel duplicate tail meta only). */
const HIDDEN_META_ROWS = new Set([138]);
/**
 * Excel rows removed from the Database grid — display numbers are renumbered.
 * Row 21 (_AIR CIRCUIT sub-section) is kept so the section/sub-section list stays
 * identical to the Synthesis page; only the surrounding STLA/S data rows are hidden.
 */
export const BD_HIDDEN_DISPLAY_ROWS = new Set([19, 20, 22]);

export function isBdHiddenDisplayRow(row) {
  return BD_HIDDEN_DISPLAY_ROWS.has(Number(row));
}
/** Column-header row in thead; first body row is 2. */
export const BD_HEADER_DISPLAY_ROW = 1;
export const BD_BODY_DISPLAY_ROW_START = 2;
/** Yellow CA chapter bands: -ADAPTATION (row 5) and -ADTH (row 139). */
export function isCaBandRow(map, row) {
  return row === 5 || row === 139;
}
/** Label in Date column on yellow CA band rows only (one row per chapter). */
export function caChapterDateLabel(row) {
  return caBandCanonicalTitle(row);
}
export function isSectionBandRow(map, row) {
  return isCaBandRow(map, row);
}
/** Cache Array→Set so per-row render lookups don't rebuild the Set every call. */
const _sectionRowSetCache = new WeakMap();
function asSectionRowSet(sectionHeaderRows) {
  if (sectionHeaderRows instanceof Set) return sectionHeaderRows;
  if (!sectionHeaderRows) return new Set();
  const cached = _sectionRowSetCache.get(sectionHeaderRows);
  if (cached) return cached;
  const set = new Set(sectionHeaderRows);
  _sectionRowSetCache.set(sectionHeaderRows, set);
  return set;
}
/**
 * One yellow row per L1 section (column AP only + CA bands 5/139 + first _Unassigned).
 */
export function computeSectionHeaderRows(sheet, cellMap = null) {
  const map =
    cellMap != null
      ? cellMap
      :
    (sheet.cellMap instanceof Map
      ? sheet.cellMap
      : buildCellMap(sheet.cells, sheet.headerRows));
  const l1Col = bdSubsystemL1Col(sheet);
  const l2Col = bdSubsystemL2Col(sheet);
  const rows = new Set();
  const canonicalByLabel = new Map();
  rows.add(5);
  rows.add(139);
  canonicalByLabel.set('-ADAPTATION', 5);
  canonicalByLabel.set('-ADTH', 139);
  let lastAp = '';
  for (let r = sheet.dataStartRow || 6; r <= sheet.lastRow; r++) {
    if (HIDDEN_META_ROWS.has(r)) continue;
    const ap = cellLabelValue(getCell(map, r, l1Col));
    if (!ap) continue;
    const label = String(ap).trim();
    if (isFormulaLike(label)) continue;
    if (isSectionLabel(label) && label !== lastAp) {
      rows.add(r);
      canonicalByLabel.set(label, r);
      lastAp = label;
    }
  }
  let lastAs = '';
  for (let r = sheet.dataStartRow || 6; r <= sheet.lastRow; r++) {
    const as = getAsLabel(map, r, l2Col);
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
export function isSubSectionRow(map, row, sectionHeaderRows, l2Col = BD_SUBSYSTEM_L2_COL) {
  if (isSectionRow(map, row, sectionHeaderRows)) return false;
  const label = getSubSectionLabel(map, row, l2Col);
  if (!label) return false;
  const prev = getSubSectionLabel(map, row - 1, l2Col);
  return canonicalSubSectionKey(label) !== canonicalSubSectionKey(prev);
}
/** First row of an L2 bookmark band: _SUB rows or blue title rows (Front wings…). */
export function isBdL2BookmarkStart(
  map,
  row,
  sectionHeaderRows,
  l2Col = BD_SUBSYSTEM_L2_COL
) {
  if (isSubSectionRow(map, row, sectionHeaderRows, l2Col)) return true;
  if (!shouldDateColBlue(map, row, sectionHeaderRows)) return false;
  const title = colATitle(map, row);
  if (!title) return false;
  if (!shouldDateColBlue(map, row - 1, sectionHeaderRows)) return true;
  return title !== colATitle(map, row - 1);
}
/** Label for matrix / structure extract (matches Database blue-band display). */
export function getBdL2BookmarkLabel(
  map,
  row,
  sheet,
  sectionHeaderRows,
  l2Col = bdSubsystemL2Col(sheet)
) {
  const sub = getSubSectionLabel(map, row, l2Col);
  if (sub) return formatBlueBandLabel(translateValue(sub));
  if (shouldDateColBlue(map, row, sectionHeaderRows)) {
    const title = colATitle(map, row);
    if (title) return formatBlueBandLabel(translateValue(title));
  }
  return '';
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
/** Vehicle header before STLA/S lines (SPC, TARGET, mass in V…). */
export function isProjectConfigRow(map, row, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  if (isDataGreenColA(map, row, sectionHeaderRows)) return false;
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
  const freeField = displayValue(getCell(map, row, BD_FREE_FIELD_COL));
  if (freeField && freeField !== '0') return true;
  return isProjectConfigRow(map, row, sectionHeaderRows);
}
function isRowInArchivedBands(sheet, row) {
  const bands = sheet && sheet.archivedRowBands;
  if (!bands || !bands.length || row == null) return false;
  for (const b of bands) {
    if (row >= b.start && row <= b.end) return true;
  }
  return false;
}

export function shouldDisplayBodyRow(map, row, sheet) {
  if (HIDDEN_META_ROWS.has(row)) return false;
  if (isBdHiddenDisplayRow(row)) return false;
  if (isRowInArchivedBands(sheet, row)) return false;
  const dataStart = (sheet && sheet.dataStartRow) || 6;
  const sh = sheet.sectionHeaderRows;
  if (
    row < dataStart &&
    !isStructureRow(map, row, sh) &&
    !isPreBandMarkerRow(map, row, sh)
  ) {
    return false;
  }
  if (isStructureRow(map, row, sh)) return true;
  if (colATitle(map, row)) return true;
  if (isSignificantDataRow(map, row, sh)) return true;
  return false;
}
/** Body rows for the grid — skips untitled white lines; consecutive display numbers. */
export function computeBodyDisplayRows(sheet) {
  const map =
    sheet.cellMap instanceof Map
      ? sheet.cellMap
      : buildCellMap(sheet.cells, sheet.headerRows);
  const rows = [];
  let displayRow = BD_BODY_DISPLAY_ROW_START;
  const userGaps = Array.isArray(sheet.userRowGaps) ? sheet.userRowGaps : [];
  const gapsAfter = (excelRow) =>
    userGaps.filter(
      (g) => g && Number(g.afterExcelRow) === Number(excelRow) && g.id
    );
  for (let r = 2; r <= sheet.lastRow; r++) {
    if (!shouldDisplayBodyRow(map, r, sheet)) continue;
    rows.push({ excelRow: r, displayRow: displayRow++ });
    for (const gap of gapsAfter(r)) {
      rows.push({
        userGap: true,
        gapKey: gap.id,
        afterExcelRow: r,
        excelRow: null,
        displayRow: displayRow++,
      });
    }
  }
  return rows;
}
export function getRowLabel(
  map,
  row,
  sectionHeaderRows,
  canonicalSectionByLabel,
  l1Col = BD_SUBSYSTEM_L1_COL_RAW
) {
  if (!isSectionRow(map, row, sectionHeaderRows)) {
    const sub = getSubSectionLabel(map, row);
    if (sub) return sub;
    const a = displayValue(getCell(map, row, 'A'));
    if (a) return a;
    return '';
  }
  if (row === 5 || row === 139) {
    // CA chapter bands: a matrix rename (column W) wins; otherwise canonical.
    return caBandStoredTitle(map, row);
  }
  const ap = cellLabelValue(getCell(map, row, l1Col));
  if (ap && isSectionLabel(ap)) return ap;
  const as = getAsLabel(map, row);
  if (isUnassignedSectionLabel(as)) return as;
  if (isCaBandRow(map, row)) {
    const w = displayValue(getCell(map, row, 'W'));
    if (w && w.startsWith('-')) return w;
  }
  return '';
}
export function isReadonlyCell(cell, row, dataStartRow) {
  if (row < 2) return true;
  if (row < dataStartRow) return true;
  if (!cell) return false;
  return Boolean(cell.f);
}

/** Body rows users may edit (mass, vehicle cols, etc.) — not L1/L2 bands or headers. */
export function isBdBodyEditableCell(sheet, row, col) {
  const dataStart = sheet && sheet.dataStartRow != null ? sheet.dataStartRow : 6;
  if (row < dataStart) return false;
  const map =
    sheet && sheet.cellMap instanceof Map
      ? sheet.cellMap
      : buildCellMap(sheet && sheet.cells, sheet && sheet.headerRows);
  const sh = sheet && sheet.sectionHeaderRows;
  if (isStructureRow(map, row, sh) || isTitleMarkerRow(map, row, sh)) {
    return false;
  }
  const l1 = bdSubsystemL1Col(sheet);
  const l2 = bdSubsystemL2Col(sheet);
  const dept = bdDesignDeptCol(sheet);
  if (col === l1 || col === l2 || col === dept) return false;
  if (BD_POSITION_COLS.has(col) || BD_MASS_AV_AR_COLS.has(col)) return false;
  return true;
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
  const u = String(t || '')
    .trim()
    .replace(/^_+/, '');
  return (
    /^Fin de Lot\b/i.test(u) ||
    /^End of lot\b/i.test(u) ||
    /end of lot.*do not delete/i.test(u) ||
    /fin de lot.*ne pas supprimer/i.test(u)
  );
}
export function isFormulesRow(map, row, sectionHeaderRows) {
  const t = colATitle(map, row);
  return t === 'Ligne avec formules' || t === 'Row with formulas';
}
export function isRecopierRow(map, row, sectionHeaderRows) {
  const t = colATitle(map, row);
  return t === 'A recopier' || t === 'To copy';
}
/** White band before ADAPTATION (rows 2–3). */
export function isWhiteMarkerRow(map, row, sectionHeaderRows) {
  return (
    isFormulesRow(map, row, sectionHeaderRows) ||
    isRecopierRow(map, row, sectionHeaderRows)
  );
}
/** Blue spacer before ADAPTATION (Excel row 4). */
export function isPreAdaptBlueRow(row) {
  return isSeparatorRow(row);
}
/** Pre-ADAPTATION marker rows (white or blue band). */
export function isPreBandMarkerRow(map, row, sectionHeaderRows) {
  return (
    isWhiteMarkerRow(map, row, sectionHeaderRows) ||
    isPreAdaptBlueRow(row)
  );
}
/** STLA/S rows: green frozen Date column (Excel). */
export function isDataGreenColA(map, row, sectionHeaderRows) {
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  const a = colATitle(map, row);
  return a === 'STLA/S' || a === 'STLA-S';
}
/** Blue title row (full row): label in A only, no component data. */
export function shouldDateColBlue(map, row, sectionHeaderRows) {
  if (HIDDEN_META_ROWS.has(row)) return false;
  if (isStructureRow(map, row, sectionHeaderRows)) return false;
  if (isPreBandMarkerRow(map, row, sectionHeaderRows)) return false;
  if (isFinDeLotRow(map, row)) return false;
  if (isDataGreenColA(map, row, sectionHeaderRows)) return false;
  return Boolean(colATitle(map, row));
}
export function rowStyleClass(map, row, sectionHeaderRows) {
  if (isPreAdaptBlueRow(row)) return 'row-pre-adapt-blue';
  if (isWhiteMarkerRow(map, row, sectionHeaderRows)) return 'row-recopier';
  if (isCaBandRow(map, row)) return 'row-section';
  if (isSectionRow(map, row, sectionHeaderRows)) return 'row-section';
  if (isSubSectionRow(map, row, sectionHeaderRows)) return 'row-subsection';
  if (isFinDeLotRow(map, row)) return 'row-fin-lot';
  if (isDataGreenColA(map, row, sectionHeaderRows)) return 'row-date-green';
  if (isProjectConfigRow(map, row, sectionHeaderRows)) return 'row-project-config';
  if (shouldDateColBlue(map, row, sectionHeaderRows)) return 'row-date-blue';
  return '';
}
/** Alternating pastel bands from Excel row 6 on Silhouette+ (incl. STLA/S data rows). */
export function rowDataStripeClass(
  map,
  row,
  sectionHeaderRows,
  dataStartRow = 6
) {
  if (isStructureRow(map, row, sectionHeaderRows)) return '';
  if (shouldDateColBlue(map, row, sectionHeaderRows)) return '';
  if (isPreBandMarkerRow(map, row, sectionHeaderRows)) return '';
  const anchor = dataStartRow || 6;
  if (row < anchor) return '';
  return (row - anchor) % 2 === 0
    ? 'row-data-stripe-green'
    : 'row-data-stripe-blue';
}
/** Semantic fill for column Project (B) only. */
export function projectCellClass(displayText, col) {
  if (col !== BD_PROJECT_COL || !displayText) return '';
  const v = String(displayText).trim().toUpperCase();
  if (!v) return '';
  if (v === 'TT') return 'cell-proj-night';
  if (v === 'TARGET') return 'cell-proj-target';
  if (v === 'STATUS') return 'cell-proj-status';
  if (v === 'SPC') return 'cell-proj-spc';
  if (v === 'INFO') return 'cell-proj-info';
  if (ENERGY_VALUES.has(v)) return 'cell-proj-energy';
  if (v === 'FWD' || v === 'AWD' || v === 'RWD') return 'cell-proj-drivetrain';
  if (v === 'PTF') return 'cell-proj-ptf';
  return 'cell-proj-night';
}
/** Title-only markers (blue / fin de lot / to copy) — no project data. */
export function isTitleMarkerRow(map, row, sectionHeaderRows) {
  return (
    shouldDateColBlue(map, row, sectionHeaderRows) ||
    isFinDeLotRow(map, row) ||
    isPreBandMarkerRow(map, row, sectionHeaderRows)
  );
}
/**
 * Outline view (eye): CA bands (5/139) + yellow sections + blue `_` sub-sections only.
 * Component-level blue bands (shouldDateColBlue, e.g. _FRONT FENDERS, _ALTERNATOR SYSTEM)
 * are Database-detail only and are intentionally excluded so the eye view, the Synthesis
 * page and the Bookmark Matrix all show the SAME section / sub-section list.
 */
export function isOutlineRow(map, row, sectionHeaderRows) {
  if (HIDDEN_META_ROWS.has(row)) return false;
  if (isBdHiddenDisplayRow(row)) return false;
  if (isFinDeLotRow(map, row)) return false;
  if (isPreBandMarkerRow(map, row, sectionHeaderRows)) return false;
  return (
    isCaBandRow(map, row) ||
    isStructureRow(map, row, sectionHeaderRows)
  );
}
/** Outline = ADAPTATION/ADTH on rows 5/139 + structure rows. */
export function computeOutlineRows(sheet, cellMap = null) {
  const map =
    cellMap != null
      ? cellMap
      :
    (sheet.cellMap instanceof Map
      ? sheet.cellMap
      : buildCellMap(sheet.cells, sheet.headerRows));
  const sectionRows = asSectionRowSet(
    sheet.sectionHeaderRows != null
      ? sheet.sectionHeaderRows
      :
      computeSectionHeaderRows(sheet, map).rows
  );
  const rows = [];
  for (let r = 2; r <= sheet.lastRow; r++) {
    if (isOutlineRow(map, r, sectionRows)) rows.push(r);
  }
  return rows;
}
export function cellInlineStyle(
  cell,
  map,
  row,
  col,
  sectionHeaderRows,
  matrixColors
) {
  const style = {};
  if (cell && cell.bg) style.backgroundColor = cell.bg;
  if (cell && cell.fc) style.color = cell.fc;
  const cls = rowStyleClass(map, row, sectionHeaderRows);
  const rowColor =
    matrixColors && (matrixColors[row] != null ? matrixColors[row] : matrixColors[String(row)]);
  if (isCaBandRow(map, row)) {
    style.backgroundColor = '#ffff00';
    style.color = '#000';
    style.fontWeight = '700';
  } else if (rowColor && (cls === 'row-section' || cls === 'row-subsection')) {
    style.backgroundColor = rowColor;
  }
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
  if (cls === 'row-recopier') {
    style.backgroundColor = '#fff';
    style.color = '#000';
  }
  if (cls === 'row-pre-adapt-blue') {
    style.backgroundColor = '#0070c0';
    style.color = '#fff';
  }
  return style;
}
/**
 * Yellow / blue bookmark rows: show section or sub-section title only, never data.
 * Returns null when the row is not a structure bookmark.
 */
function isDesignDeptCol(col) {
  return col === BD_DESIGN_DEPT_COL || col === BD_DESIGN_DEPT_COL_RAW;
}

function structureBookmarkDisplay(
  map,
  row,
  col,
  sectionHeaderRows,
  canonicalByLabel,
  l1Col = BD_SUBSYSTEM_L1_COL_RAW,
  l2Col = BD_SUBSYSTEM_L2_COL
) {
  if (!isStructureRow(map, row, sectionHeaderRows)) return null;
  if (isDesignDeptCol(col)) return null;
  if (isSeparatorRow(row)) {
    return '';
  }
  if (isSectionRow(map, row, sectionHeaderRows)) {
    const rawTitle = getRowLabel(
      map,
      row,
      sectionHeaderRows,
      canonicalByLabel,
      l1Col
    );
    if (!rawTitle) return '';
    const title = formatSectionDisplayLabel(rawTitle);
    /* CA bands 5 / 139: full-row yellow; label without "_" or leading "-". */
    if (isCaBandRow(map, row)) {
      if (col === 'A' || col === 'W') return title || caChapterDateLabel(row);
      return '';
    }
    if (col === 'A') return title;
    if (col === l1Col && canonicalByLabel.get(rawTitle) === row) return title;
    if (col === l2Col && isUnassignedSectionLabel(rawTitle)) return title;
    return '';
  }
  if (isSubSectionRow(map, row, sectionHeaderRows)) {
    const title = formatBlueBandLabel(
      translateValue(getSubSectionLabel(map, row, l2Col))
    );
    if (!title) return '';
    if (col === 'A') return title;
    return '';
  }
  return null;
}
/** Hide TT / 0 only on inherited filler rows — keep on STLA/S & project rows. */
function maskStructureValue(map, row, col, v, sectionHeaderRows) {
  if (isMassCol(col)) return v;
  if (isStructureRow(map, row, sectionHeaderRows)) return v;
  if (isDataGreenColA(map, row, sectionHeaderRows)) return v;
  if (isProjectConfigRow(map, row, sectionHeaderRows)) return v;
  if (v === 'TT' || v === '0') return '';
  return v;
}
function isInheritedBandOrSection(v) {
  if (!v) return false;
  const t = String(v).trim();
  if (/^-?\d+(\.\d+)?$/.test(t)) return false;
  return isSectionLabel(v) || (t.startsWith('-') && t.length > 1);
}
/** Mass / numeric columns must not use section-label masking. */
function isMassCol(col) {
  return col === BD_MASS_COL;
}
const _canonByLabelCache = new WeakMap();
function canonicalByLabelMap(canonicalSectionByLabel) {
  if (!canonicalSectionByLabel) return new Map();
  if (canonicalSectionByLabel instanceof Map) return canonicalSectionByLabel;
  let cached = _canonByLabelCache.get(canonicalSectionByLabel);
  if (!cached) {
    cached = new Map(Object.entries(canonicalSectionByLabel));
    _canonByLabelCache.set(canonicalSectionByLabel, cached);
  }
  return cached;
}
/** Section title visible on exactly one cell: W@5/139 or L1@canonical row. */
function isCanonicalSectionDisplay(
  map,
  row,
  col,
  v,
  canonicalByLabel,
  l1Col = BD_SUBSYSTEM_L1_COL_RAW
) {
  if (!v || !isInheritedBandOrSection(v)) return false;
  const label = String(v).trim();
  if (isCaBandRow(map, row) && col === 'W') {
    // W is the canonical (editable) CA-chapter title cell, whatever its label.
    return row === 5 || row === 139;
  }
  return canonicalByLabel.get(label) === row && col === l1Col;
}
/** Hide L1 / CA titles everywhere except their single canonical cell. */
function maskDuplicateSectionLabel(
  map,
  row,
  col,
  v,
  canonicalByLabel,
  l1Col = BD_SUBSYSTEM_L1_COL_RAW
) {
  if (!v || !isInheritedBandOrSection(v)) return v;
  if (isCanonicalSectionDisplay(map, row, col, v, canonicalByLabel, l1Col)) {
    return v;
  }
  return '';
}
/** _Unassigned only on the first row of the block (L2 column). */
function maskRepeatedUnassigned(map, row, col, v, l2Col) {
  if (col !== l2Col || !isUnassignedSectionLabel(v)) return v;
  const prev = getAsLabel(map, row - 1, l2Col);
  if (isUnassignedSectionLabel(prev) && String(prev).trim() === String(v).trim()) {
    return '';
  }
  return v;
}
/** L2 titles only on blue header rows — hide copies on data rows & duplicate AS when A shows it. */
function maskInheritedSubSectionLabel(
  map,
  row,
  col,
  v,
  sectionHeaderRows,
  l2Col
) {
  if (!v) return v;
  const t = String(v).trim();
  if (!t.startsWith('_') || isUnassignedSectionLabel(t)) return v;
  if (col === l2Col) {
    if (isSubSectionRow(map, row, sectionHeaderRows)) {
      return formatBlueBandLabel(translateValue(v));
    }
    // L2 bookmark title only on the blue header row — hide on STLA/S data rows.
    return '';
  }
  if (isSubSectionRow(map, row, sectionHeaderRows)) {
    if (col === 'AS') {
      const a = displayValue(getCell(map, row, 'A'));
      if (a && a.trim() === t) return '';
    }
    return formatBlueBandLabel(translateValue(v));
  }
  return '';
}
/** Mass column: at most 4 decimal places (Excel Masse). */
export function formatMassDisplayValue(raw) {
  if (raw == null || raw === '') return '';
  const s = stripExcelErrorValue(String(raw).trim());
  if (!s) return '';
  if (!/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(s)) return s;
  const n = Number(s);
  if (Number.isNaN(n)) return s;
  const rounded = Math.round(n * 10000) / 10000;
  if (Number.isInteger(rounded)) return String(rounded);
  return parseFloat(rounded.toFixed(4)).toString();
}
/** Excel cached mass (V) — never masked as section labels; keep 0 and negatives. */
function massDisplayValue(map, row, col) {
  const cell = getCell(map, row, col);
  if (!cell) return '';
  if (cell.v != null && cell.v !== '') {
    return formatMassDisplayValue(cell.v);
  }
  return '';
}
function blueTitleDisplay(map, row, col, sectionHeaderRows) {
  if (!shouldDateColBlue(map, row, sectionHeaderRows) || col !== 'A') return null;
  const title = colATitle(map, row);
  if (!title) return '';
  return formatBlueBandLabel(translateValue(title));
}
export function displayCellValue(
  map,
  row,
  col,
  sectionHeaderRows,
  canonicalSectionByLabel,
  l1Col = BD_SUBSYSTEM_L1_COL_RAW,
  l2Col = BD_SUBSYSTEM_L2_COL
) {
  const canonicalByLabel =
    canonicalSectionByLabel instanceof Map
      ? canonicalSectionByLabel
      : canonicalByLabelMap(canonicalSectionByLabel);

  if (isPreAdaptBlueRow(row)) return '';

  if (isWhiteMarkerRow(map, row, sectionHeaderRows)) {
    if (col === 'A') {
      const t = colATitle(map, row);
      return t ? translateValue(t) : '';
    }
    return '';
  }

  if (isStructureRow(map, row, sectionHeaderRows)) {
    if (isMassCol(col)) {
      return massDisplayValue(map, row, col);
    }
    if (col === l1Col) {
      const cell = getCell(map, row, col);
      let v = stripExcelErrorValue(displayValue(cell));
      v = maskStructureValue(map, row, col, v, sectionHeaderRows);
      return v ? translateSubsystemLabel(v) : '';
    }
    const bookmark = structureBookmarkDisplay(
      map,
      row,
      col,
      sectionHeaderRows,
      canonicalByLabel,
      l1Col,
      l2Col
    );
    if (bookmark != null && bookmark !== '') return bookmark;
    return '';
  }

  if (shouldDateColBlue(map, row, sectionHeaderRows)) {
    if (col === 'A') {
      const title = colATitle(map, row);
      return title ? formatBlueBandLabel(translateValue(title)) : '';
    }
    if (isMassCol(col)) {
      return massDisplayValue(map, row, col);
    }
    if (
      col === l1Col ||
      col === l2Col ||
      col === BD_DESIGN_DEPT_COL ||
      col === BD_DESIGN_DEPT_COL_RAW
    ) {
      const cell = getCell(map, row, col);
      let v = stripExcelErrorValue(displayValue(cell));
      if (col === l2Col && cell && (cell.v == null || cell.v === '') && cell.f) {
        v = '';
      }
      return v ? translateSubsystemLabel(v) : '';
    }
    return '';
  }

  if (BD_MASS_AV_AR_COLS.has(col)) return '';
  if (BD_POSITION_COLS.has(col)) return '';
  if (col === BD_TITLE_COL) return '';
  if (isMassCol(col)) {
    return massDisplayValue(map, row, col);
  }
  const cell = getCell(map, row, col);
  let v = displayValue(cell);
  v = stripExcelErrorValue(v);
  if (col === l2Col && cell && (cell.v == null || cell.v === '') && cell.f) v = '';
  if (col === l1Col) {
    v = maskStructureValue(map, row, col, v, sectionHeaderRows);
    return v ? translateSubsystemLabel(v) : '';
  }
  if (col === l2Col) {
    v = maskStructureValue(map, row, col, v, sectionHeaderRows);
    v = maskInheritedSubSectionLabel(
      map,
      row,
      col,
      v,
      sectionHeaderRows,
      l2Col
    );
    v = maskRepeatedUnassigned(map, row, col, v, l2Col);
    if (
      isTitleMarkerRow(map, row, sectionHeaderRows) &&
      col !== 'A' &&
      col !== BD_TRADE_COL
    ) {
      return '';
    }
    return v ? translateSubsystemLabel(v) : '';
  }
  if (col === BD_DESIGN_DEPT_COL || col === BD_DESIGN_DEPT_COL_RAW) {
    if (
      isTitleMarkerRow(map, row, sectionHeaderRows) &&
      col !== 'A' &&
      col !== BD_TRADE_COL
    ) {
      return '';
    }
    if (cell && cell.v != null && cell.v !== '') {
      return translateSubsystemLabel(stripExcelErrorValue(String(cell.v)));
    }
    return '';
  }
  v = maskStructureValue(map, row, col, v, sectionHeaderRows);
  if (col !== l1Col && col !== l2Col && !isMassCol(col)) {
    v = maskDuplicateSectionLabel(map, row, col, v, canonicalByLabel, l1Col);
    v = maskInheritedSubSectionLabel(
      map,
      row,
      col,
      v,
      sectionHeaderRows,
      l2Col
    );
  }
  v = maskRepeatedUnassigned(map, row, col, v, l2Col);
  if (
    isTitleMarkerRow(map, row, sectionHeaderRows) &&
    col !== 'A' &&
    col !== BD_TRADE_COL
  ) {
    return '';
  }
  if (col !== 'A' || v) return v ? translateValue(v) : '';
  return '';
}
