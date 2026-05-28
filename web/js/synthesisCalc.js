/**
 * Live mass aggregation SYNTHESIS ↔ BD (Excel SUMPRODUCT equivalent).
 * Rows 26+ : blue = BD mass filtered; green = sum of blues until next section.
 */
import { BD_SUBSYSTEM_L2_COL, BD_MASS_COL } from './bdColumnConfig.js';
import { isSectionLabel } from './bdStore.js';
import { translateSubsystemLabel, canonicalL2MatchKey } from './bdTranslate.js';
import {
  colToNum,
  numToCol,
  displayToExcelCol,
  excelToDisplayCol,
  isSynSpacerDisplayExcelCol,
  isSynSp2DisplayExcelCol,
  isSynSp2RestartDisplayExcelCol,
  SYN_AC_AN_TABLE_DISPLAY_START,
  SYN_AC_AN_TABLE_DISPLAY_END,
  SYN_AP_BB_TABLE_DISPLAY_START,
  SYN_AP_BB_TABLE_DISPLAY_END,
  isSynHeaderPanelVehicleCol,
  isSynBuiltinPillarExcelCol,
} from './synthesisPerf.js';

const BD_END_ROW = 3480;

/** Synthesis filter row → BD column (same as Excel SUMPRODUCT). */
export const SYN_FILTER_BD_ROWS = [
  ['A', 3],
  ['B', 4],
  ['C', 5],
  ['D', 7],
  ['E', 6],
  ['F', 8],
  ['G', 9],
  ['I', 13],
  ['K', 11],
  ['L', 12],
  ['O', 14],
];

/** Human-readable map for the Database page. */
export const SYN_FILTER_BD_LABELS = [
  { synRow: 3, bdCol: 'A', label: 'Date / Org' },
  { synRow: 4, bdCol: 'B', label: 'Project' },
  { synRow: 5, bdCol: 'C', label: 'Silhouette' },
  { synRow: 6, bdCol: 'E', label: 'Hybridization' },
  { synRow: 7, bdCol: 'D', label: 'Design plate' },
  { synRow: 8, bdCol: 'F', label: 'Seats' },
  { synRow: 9, bdCol: 'G', label: 'Tech spec' },
  { synRow: 11, bdCol: 'K', label: 'Pole' },
  { synRow: 12, bdCol: 'L', label: 'Energy' },
  { synRow: 13, bdCol: 'I', label: 'Pack' },
  { synRow: 14, bdCol: 'O', label: 'Finish' },
];

/** First Excel row for live blue/green mass engine (-ADAPTATION green total, row 26+ blues). */
export const SYN_CALC_FIRST_ROW = 25;

/** Display M…AA (Excel R…AF) — project header band (subset of calc cols). */
export const SYN_MAA_DISPLAY_START = 'M';
export const SYN_MAA_DISPLAY_END = 'AA';

/** Display AC…AN (Excel AH…AS) — second project table (same SUMPRODUCT rules as M…AA). */
export const SYN_ACAN_DISPLAY_START = SYN_AC_AN_TABLE_DISPLAY_START;
export const SYN_ACAN_DISPLAY_END = SYN_AC_AN_TABLE_DISPLAY_END;

/** Display AP…BB (Excel AU…BG) — third project table (same SUMPRODUCT rules as M…AA). */
export const SYN_APBB_DISPLAY_START = SYN_AP_BB_TABLE_DISPLAY_START;
export const SYN_APBB_DISPLAY_END = SYN_AP_BB_TABLE_DISPLAY_END;

/** Rows blank in M…AA — no live blue mass (grey spacers). */
const SYN_SUMPRODUCT_SKIP_ROWS = new Set([
  34, 37, 40, 49, 50, 78, 79, 80, 83, 84, 105, 106, 107, 109,
]);

function parseNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
}

/** Excel column letters for display M…AA (vehicle filter axis per column). */
export function synMaaExcelCols() {
  const out = [];
  for (let d = colToNum(SYN_MAA_DISPLAY_START); d <= colToNum(SYN_MAA_DISPLAY_END); d++) {
    out.push(displayToExcelCol(numToCol(d)));
  }
  return out;
}

/** Grid stores Excel letters (R…); UI labels use display letters (M…, AC…). */
export function synMassDisplayCol(col) {
  if (!col) return col;
  const s = String(col);
  const n = colToNum(s);
  const maaExcelStart = colToNum(displayToExcelCol(SYN_MAA_DISPLAY_START));
  if (n >= maaExcelStart) {
    return excelToDisplayCol(s);
  }
  return s;
}

export function isSynMaaMassCol(col) {
  const d = synMassDisplayCol(col);
  const n = colToNum(d);
  return (
    n >= colToNum(SYN_MAA_DISPLAY_START) && n <= colToNum(SYN_MAA_DISPLAY_END)
  );
}

/** Excel columns for display AC…AN (vehicle filter axis AH…AS, rows 3–14). */
export function synAcanExcelCols() {
  const out = [];
  for (
    let d = colToNum(SYN_ACAN_DISPLAY_START);
    d <= colToNum(SYN_ACAN_DISPLAY_END);
    d++
  ) {
    out.push(displayToExcelCol(numToCol(d)));
  }
  return out;
}

export function isSynAcanMassCol(col) {
  const d = synMassDisplayCol(col);
  const n = colToNum(d);
  return (
    n >= colToNum(SYN_ACAN_DISPLAY_START) && n <= colToNum(SYN_ACAN_DISPLAY_END)
  );
}

/** Display AP…BB (Excel AU…BG) — header panel rows 3–22. */
export function isSynApbbMassCol(col) {
  const d = synMassDisplayCol(col);
  const n = colToNum(d);
  return (
    n >= colToNum(SYN_APBB_DISPLAY_START) &&
    n <= colToNum(SYN_APBB_DISPLAY_END)
  );
}

/** Excel columns for display AP…BB (vehicle filter axis AU…BG, rows 3–14). */
export function synApbbExcelCols() {
  const out = [];
  for (
    let d = colToNum(SYN_APBB_DISPLAY_START);
    d <= colToNum(SYN_APBB_DISPLAY_END);
    d++
  ) {
    out.push(displayToExcelCol(numToCol(d)));
  }
  return out;
}

/** Excel column for cellMap lookups and BD filter axis (display AC → AH). */
export function synVehicleMassExcelCol(col) {
  const d = synMassDisplayCol(col);
  const n = colToNum(d);
  if (
    n >= colToNum(SYN_APBB_DISPLAY_START) &&
    n <= colToNum(SYN_APBB_DISPLAY_END)
  ) {
    return displayToExcelCol(d);
  }
  if (
    n >= colToNum(SYN_ACAN_DISPLAY_START) &&
    n <= colToNum(SYN_ACAN_DISPLAY_END)
  ) {
    return displayToExcelCol(d);
  }
  if (
    n >= colToNum(SYN_MAA_DISPLAY_START) &&
    n <= colToNum(SYN_MAA_DISPLAY_END)
  ) {
    return displayToExcelCol(d);
  }
  return String(col);
}

/** Live SOMMEPROD columns: display M…AA + AC…AN + AP…BB (Excel R…AF, AH…AS, AU…BG). */
export function isSynVehicleMassCol(col) {
  return isSynMaaMassCol(col) || isSynAcanMassCol(col) || isSynApbbMassCol(col);
}

/** Filter-row edit (rows 3–14) on a project vehicle column — invalidates that column's BD filter cache. */
export function isSynFilterEdit(row, col) {
  if (row < 3 || row > 14) return false;
  if (isSynHeaderPanelVehicleCol(col)) return false;
  return isSynVehicleMassCol(col);
}

/** Display AB = Excel AG — Δ(V − Y) on the same row (display V=AA, Y=AD). */
export const SYN_AB_DIFF_DISPLAY_COL = 'AB';
export const SYN_AB_DIFF_V_DISPLAY_COL = 'V';
export const SYN_AB_DIFF_Y_DISPLAY_COL = 'Y';

export function synAbDiffExcelCol() {
  return displayToExcelCol(SYN_AB_DIFF_DISPLAY_COL);
}

export function synAbDiffVExcelCol() {
  return displayToExcelCol(SYN_AB_DIFF_V_DISPLAY_COL);
}

export function synAbDiffYExcelCol() {
  return displayToExcelCol(SYN_AB_DIFF_Y_DISPLAY_COL);
}

export function isSynAbDiffExcelCol(col) {
  return col === synAbDiffExcelCol();
}

/** Live formula cell: display AB = display V − display Y (same row). */
export function isSynAbDiffCell(row, col, sheet) {
  if (!isSynAbDiffExcelCol(col)) return false;
  return isSynCalcRow(row, sheet);
}

export function computeSynAbDiff(getNumeric, row) {
  const v = getNumeric(row, synAbDiffVExcelCol());
  const y = getNumeric(row, synAbDiffYExcelCol());
  return Math.round((v - y) * 10000) / 10000;
}

/** Vehicle / data column eligible for live blue or green mass (G…last, no pillars/spacer). */
export function isSynMassCalcCol(col) {
  if (col === 'F') return false;
  if (isSynAbDiffExcelCol(col)) return false;
  if (isSynSpacerDisplayExcelCol(col)) return false;
  if (isSynSp2DisplayExcelCol(col)) return false;
  if (isSynSp2RestartDisplayExcelCol(col)) return false;
  return colToNum(col) >= colToNum('G');
}

/** All Excel vehicle columns with live filter index (M…AA + AC…AN + AP…BB). */
export function synCalcExcelCols(_sheet) {
  return [...new Set([...synMaaExcelCols(), ...synAcanExcelCols(), ...synApbbExcelCols()])];
}

function isSynCalcRow(row, sheet) {
  const r = Number(row);
  if (!Number.isFinite(r) || r < SYN_CALC_FIRST_ROW) return false;
  const last = sheet?.effectiveLastRow ?? sheet?.lastRow ?? 422;
  return r <= last;
}

/** ADAPTATION block — row 26 = SUM(rows 27:41) per column, from display C (Excel H) through last data col. */
export const SYN_ADAPTATION_SUM_ROW = 26;
export const SYN_ADAPTATION_SUM_FROM_ROW = 27;
export const SYN_ADAPTATION_SUM_TO_ROW = 41;
/** Excel H = display C — first column included in row-26 totals. */
export const SYN_ROW26_SUM_EXCEL_COL_START = 'H';

function colToNumLocal(col) {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - 64);
  }
  return n;
}

/** Row 26 total column — all body cols from display C onward (not label F, pillars, spacers). */
export function isSynRow26SumCol(col) {
  if (!col || col === 'F') return false;
  if (isSynSpacerDisplayExcelCol(col)) return false;
  if (isSynSp2DisplayExcelCol(col)) return false;
  if (isSynSp2RestartDisplayExcelCol(col)) return false;
  if (isSynBuiltinPillarExcelCol(col)) return false;
  return colToNumLocal(String(col)) >= colToNumLocal(SYN_ROW26_SUM_EXCEL_COL_START);
}

/** @deprecated alias */
export function isSynAdaptationSumCol(col) {
  return isSynRow26SumCol(col);
}

export function isSynAdaptationSumCell(row, col) {
  return (
    Number(row) === SYN_ADAPTATION_SUM_ROW && isSynAdaptationSumCol(col)
  );
}

export function affectsAdaptationSum(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || r < SYN_ADAPTATION_SUM_FROM_ROW) return false;
  if (r > SYN_ADAPTATION_SUM_TO_ROW) return false;
  return isSynAdaptationSumCol(col);
}

/** @param {(row: number, col: string) => number} getNumeric */
export function computeAdaptationRowSum(
  getNumeric,
  col,
  fromRow = SYN_ADAPTATION_SUM_FROM_ROW,
  toRow = SYN_ADAPTATION_SUM_TO_ROW
) {
  let sum = 0;
  for (let r = fromRow; r <= toRow; r++) {
    sum += getNumeric(r, col);
  }
  return Math.round(sum * 10000) / 10000;
}

/** Yellow/green L1 section band — total row for blues below until next section. */
export function isSynGreenSectionRow(row, sheet, synLabel = '', rowClass = '') {
  if (
    rowClass === 'syn-row-filter' ||
    rowClass === 'syn-row-metric' ||
    rowClass === 'syn-row-separator' ||
    rowClass === 'syn-row-subsection'
  ) {
    return false;
  }
  if (rowClass === 'syn-row-section') return true;
  const band = sheet?.rowBands?.[String(row)] ?? sheet?.rowBands?.[row];
  if (band === 'section') return true;
  if (band === 'filter' || band === 'separator' || band === 'subsection') {
    return false;
  }
  const label = String(synLabel ?? '').trim();
  if (label.startsWith('-')) return true;
  return isSectionLabel(label);
}

/** Blue L2 subsection band (same rows as Database sub-section blues). */
export function isSynBlueSubsectionRow(row, sheet, synLabel = '', rowClass = '') {
  const r = Number(row);
  if (!Number.isFinite(r) || SYN_SUMPRODUCT_SKIP_ROWS.has(r)) return false;
  if (rowClass === 'syn-row-subsection') return true;
  if (
    rowClass &&
    rowClass !== 'syn-row-data' &&
    rowClass !== 'syn-row-metric'
  ) {
    return false;
  }
  const band = sheet?.rowBands?.[String(r)] ?? sheet?.rowBands?.[r];
  if (band === 'subsection') return true;
  if (band === 'section' || band === 'filter' || band === 'separator') return false;
  const label = String(synLabel ?? '').trim();
  if (label.startsWith('_')) return true;
  return false;
}

/** Column-oriented sparse arrays indexed by Excel row. */
export function buildBdColumnIndex(bdRaw) {
  const cols = {};
  const add = (col, row, v) => {
    if (!cols[col]) cols[col] = [];
    cols[col][row] = v;
  };
  for (const cell of bdRaw.cells || []) {
    add(cell.c, cell.r, cell.v);
  }
  for (const [rowStr, rowCols] of Object.entries(bdRaw.headerRows || {})) {
    const row = parseInt(rowStr, 10);
    for (const [col, cell] of Object.entries(rowCols)) {
      add(col, row, cell.v);
    }
  }
  return cols;
}

function columnMatches(bdVal, crit) {
  const v = String(bdVal ?? '').trim();
  const c = String(crit ?? '').trim();
  if (c === 'TT') return true;
  if (!c) return v === 'TT' || v === '';
  return v === c || v === 'TT';
}

function rowMatchesFilters(bdCols, vehCol, bdRow, getSynCell, getBdValue) {
  for (const [bdCol, filterRow] of SYN_FILTER_BD_ROWS) {
    const crit = getSynCell(filterRow, vehCol);
    const bdVal = getBdValue
      ? getBdValue(bdRow, bdCol)
      : bdCols[bdCol]?.[bdRow];
    if (!columnMatches(bdVal, crit)) return false;
  }
  const q = String(
    (getBdValue ? getBdValue(bdRow, 'Q') : bdCols.Q?.[bdRow]) ?? ''
  ).trim();
  if (q !== 'S') return false;
  return true;
}

/** BD rows passing column filters for one Synthesis vehicle column (cacheable). */
export function buildVehColFilterIndex(bdCols, vehCol, getSynCell, getBdValue) {
  const rows = [];
  const end = Math.min(BD_END_ROW, 4000);
  for (let r = 2; r <= end; r++) {
    if (rowMatchesFilters(bdCols, vehCol, r, getSynCell, getBdValue)) {
      rows.push(r);
    }
  }
  return rows;
}

/** Index BD body rows with Q=S by L2 key (Excel AS → AU on Database page). */
export function buildSumproductL2Index(bdCols, getBdValue) {
  const index = new Map();
  const end = Math.min(BD_END_ROW, 4000);
  for (let r = 2; r <= end; r++) {
    const q = String(
      (getBdValue ? getBdValue(r, 'Q') : bdCols.Q?.[r]) ?? ''
    ).trim();
    if (q !== 'S') continue;
    const l2Raw =
      (getBdValue ? getBdValue(r, BD_SUBSYSTEM_L2_COL) : bdCols[BD_SUBSYSTEM_L2_COL]?.[r]) ??
      '';
    const key = canonicalL2MatchKey(l2Raw);
    if (!key) continue;
    if (!index.has(key)) index.set(key, []);
    index.get(key).push(r);
  }
  return index;
}

/**
 * Exact Excel SOMMEPROD for one blue cell (display AC…AN → Excel AH…AS, or M…AA → R…AF):
 *   Σ BD.V  ×  (BD L2 = Syn F)  ×  filtres AH$3…$14  ×  (BD.Q = "S")
 * BD L2 = Excel AS, visible on Database as col. AU after export transform.
 */
export function computeSumproduct(
  bdCols,
  synRow,
  vehCol,
  getSynCell,
  getBdValue,
  l2Index = null,
  filterRows = null
) {
  const labelKey = canonicalL2MatchKey(getSynCell(synRow, 'F'));
  const weights = bdCols.V;
  if (!weights) return 0;

  let scan;
  if (labelKey && l2Index instanceof Map) {
    const l2Rows = l2Index.get(labelKey) || [];
    if (filterRows?.length) {
      const filt = new Set(filterRows);
      scan = l2Rows.filter((r) => filt.has(r));
    } else {
      scan = l2Rows;
    }
  } else if (filterRows?.length) {
    scan = filterRows;
  } else {
    scan = Array.from({ length: Math.min(BD_END_ROW, 4000) - 1 }, (_, i) => i + 2);
  }

  let sum = 0;
  for (const r of scan) {
    if (labelKey) {
      const l2Raw =
        (getBdValue
          ? getBdValue(r, BD_SUBSYSTEM_L2_COL)
          : bdCols[BD_SUBSYSTEM_L2_COL]?.[r]) ?? '';
      if (canonicalL2MatchKey(l2Raw) !== labelKey) continue;
    }
    if (
      !filterRows &&
      !rowMatchesFilters(bdCols, vehCol, r, getSynCell, getBdValue)
    ) {
      continue;
    }
    const mass = getBdValue ? getBdValue(r, BD_MASS_COL) : weights[r];
    sum += parseNum(mass);
  }
  return sum;
}

/** @deprecated Alias — same as {@link computeSumproduct} for display AC…AN. */
export const computeAcanSumproduct = computeSumproduct;

/** L2 sub-system registry from Database (for the Database page). */
export function buildBdL2Registry(bdCols, getBdValue) {
  const byLabel = new Map();
  const end = Math.min(BD_END_ROW, 4000);
  for (let r = 2; r <= end; r++) {
    const q = String(
      (getBdValue ? getBdValue(r, 'Q') : bdCols.Q?.[r]) ?? ''
    ).trim();
    if (q !== 'S') continue;
    const l2 = String(
      (getBdValue ? getBdValue(r, BD_SUBSYSTEM_L2_COL) : bdCols[BD_SUBSYSTEM_L2_COL]?.[r]) ??
        ''
    ).trim();
    if (!l2) continue;
    const mass = parseNum(getBdValue ? getBdValue(r, BD_MASS_COL) : bdCols.V?.[r]);
    const hit = byLabel.get(l2) || { label: l2, rows: 0, mass: 0 };
    hit.rows += 1;
    hit.mass += mass;
    byLabel.set(l2, hit);
  }
  return [...byLabel.values()]
    .map((e) => ({
      ...e,
      mass: Math.round(e.mass * 10000) / 10000,
    }))
    .sort((a, b) => a.label.localeCompare(b.label));
}

/** Sorted Excel rows for yellow/green L1 section headers (row 26+). */
export function buildSynSectionRowList(sheet, getLabel, getRowClass) {
  const first = SYN_CALC_FIRST_ROW;
  const last = sheet?.effectiveLastRow ?? sheet?.lastRow ?? 422;
  const rows = [];
  for (let r = first; r <= last; r++) {
    const label = getLabel(r);
    const rowClass = getRowClass(r);
    if (isSynGreenSectionRow(r, sheet, label, rowClass)) rows.push(r);
  }
  return rows;
}

/**
 * Green section row = sum of blue rows after this section until the next section.
 */
export function computeSynSectionSum(
  sectionRow,
  col,
  sectionRows,
  isBlueRow,
  getBlueValueAt
) {
  const idx = sectionRows.indexOf(sectionRow);
  if (idx < 0) return 0;
  const end = sectionRows[idx + 1] ?? Number.MAX_SAFE_INTEGER;
  let sum = 0;
  for (let r = sectionRow + 1; r < end; r++) {
    if (!isBlueRow(r)) continue;
    sum += parseNum(getBlueValueAt(r, col));
  }
  return Math.round(sum * 10000) / 10000;
}

/** @deprecated Use {@link computeSynSectionSum}. */
export const computeSynSectionMaaSum = computeSynSectionSum;

/** Green section total — sum of blue rows in the block (same column). */
export function isSynSectionSumDataCell(
  row,
  col,
  sheet,
  cell,
  synLabel = '',
  rowClass = ''
) {
  if (cell?.userEdited) return false;
  if (!isSynVehicleMassCol(col)) return false;
  if (!isSynCalcRow(row, sheet)) return false;
  return isSynGreenSectionRow(row, sheet, synLabel, rowClass);
}

/** Blue-row cell — live SOMMEPROD from Database (filtres col. lignes 3–14, L2=F, Q=S). */
export function isSynSumproductDataCell(
  row,
  col,
  sheet,
  cell,
  synLabel = '',
  rowClass = ''
) {
  if (cell?.userEdited) return false;
  if (!isSynVehicleMassCol(col)) return false;
  if (!isSynCalcRow(row, sheet)) return false;
  if (Number(row) === SYN_ADAPTATION_SUM_ROW) return false;
  if (isSynAdaptationSumCell(row, col)) return false;
  if (!String(synLabel ?? '').trim()) return false;
  return isSynBlueSubsectionRow(row, sheet, synLabel, rowClass);
}

/** Body cell computed live (blue BD mass or green section subtotal). */
export function isSynCalculatedMassCell(
  cell,
  row,
  col,
  sheet,
  synLabel = '',
  rowClass = ''
) {
  if (cell?.f && /SUMPRODUCT/i.test(cell.f)) return true;
  if (row == null || col == null || !sheet) {
    return Boolean(cell?.f && /SUMPRODUCT/i.test(cell.f));
  }
  return (
    isSynSectionSumDataCell(row, col, sheet, cell, synLabel, rowClass) ||
    isSynSumproductDataCell(row, col, sheet, cell, synLabel, rowClass) ||
    isSynAbDiffCell(row, col, sheet) ||
    isSynAdaptationSumCell(row, col)
  );
}

export function isSynCalculatedMaaCell(...args) {
  return isSynCalculatedMassCell(...args);
}

export function isSumproductCell(cell, row, col, sheet, synLabel = '', rowClass = '') {
  return isSynCalculatedMassCell(cell, row, col, sheet, synLabel, rowClass);
}

function isSynNumericRaw(raw) {
  const s = String(raw ?? '').trim();
  if (!s) return false;
  return /^-?\d+([.,]\d+)?([eE][+-]?\d+)?$/.test(s);
}

/** Excel cols with live mass engine on row 26+ (M…AA, AC…AN, AP…BB, AB Δ). */
export function synLiveMassExcelCols() {
  const set = new Set([
    ...synMaaExcelCols(),
    ...synAcanExcelCols(),
    ...synApbbExcelCols(),
    synAbDiffExcelCol(),
  ]);
  return [...set];
}

/** Drop stale export values on live calc bands so the engine always wins. */
export function sanitizeSynLiveMassCells(cells = []) {
  const cols = new Set(synLiveMassExcelCols());
  for (let i = cells.length - 1; i >= 0; i--) {
    const cell = cells[i];
    if (Number(cell.r) < SYN_CALC_FIRST_ROW) continue;
    if (!cols.has(cell.c)) continue;
    if (cell.userEdited) continue;
    cells.splice(i, 1);
  }
  return cells;
}

/** Drop stale export/session values on row 26 so the live SUM(27:41) always wins. */
export function sanitizeSynAdaptationSumCells(cells = []) {
  for (let i = cells.length - 1; i >= 0; i--) {
    const cell = cells[i];
    if (!isSynAdaptationSumCell(cell.r, cell.c)) continue;
    if (cell.userEdited) continue;
    cells.splice(i, 1);
  }
  return cells;
}

const SYN_FORMULA_FILTER_LABELS = new Map(
  SYN_FILTER_BD_LABELS.map((f) => [f.synRow, f.label])
);

function synFormulaFilterCrit(raw) {
  const c = String(raw ?? '').trim();
  if (!c) return '∅';
  if (c === 'TT') return 'TT (tous)';
  return `"${c}"`;
}

/**
 * Human-readable formula for the Synthesis formula bar (Database-driven, not Excel export).
 */
export function describeSynCellFormula(
  row,
  col,
  sheet,
  synLabelText = '',
  rowClass = '',
  getSynCell = () => ''
) {
  const r = Number(row);
  const excelCol = synVehicleMassExcelCol(col);
  const dispCol = excelToDisplayCol(excelCol);
  const colRef = dispCol;

  if (isSynAdaptationSumCell(r, col)) {
    return `=SOMME(${colRef}${SYN_ADAPTATION_SUM_FROM_ROW}:${colRef}${SYN_ADAPTATION_SUM_TO_ROW}) — total colonne (lignes bleues + SOMMEPROD BD)`;
  }

  if (
    isSynSectionSumDataCell(r, col, sheet, null, synLabelText, rowClass)
  ) {
    return `=SOMME(lignes bleues col. ${colRef} jusqu'à la prochaine ligne verte)`;
  }

  if (
    isSynSumproductDataCell(r, col, sheet, null, synLabelText, rowClass)
  ) {
    const l2 = String(synLabelText ?? '').trim() || '?';
    const filterParts = SYN_FILTER_BD_ROWS.map(([bdCol, synRow]) => {
      const crit = getSynCell(synRow, excelCol);
      const label = SYN_FORMULA_FILTER_LABELS.get(synRow) ?? bdCol;
      return `BD.${bdCol}=${synFormulaFilterCrit(crit)} (${label}, filtre ${colRef}${synRow})`;
    });
    return (
      `=SOMMEPROD(` +
      `BD.V × BD.AU="${l2}" × BD.Q="S" × ` +
      filterParts.join(' × ') +
      `) — colonne ${colRef} (filtres propres à cette colonne, ex. Silhouette ${colRef}5)`
    );
  }

  if (isSynAbDiffCell(r, col, sheet)) {
    return `=${excelToDisplayCol(synAbDiffVExcelCol())}${r} − ${excelToDisplayCol(synAbDiffYExcelCol())}${r}`;
  }

  return '';
}
