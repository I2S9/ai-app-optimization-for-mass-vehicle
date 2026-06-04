/** SYNTHESIS sheet display helpers (filter band, labels, merges). */
import { displayValue, getCell, isSectionLabel, formatBlueBandLabel } from './bdStore.js';
import { translateValue, translateSubsystemLabel } from './bdTranslate.js';
import {
  displayToExcelCol,
  excelToDisplayCol,
  isSynFilterGreyExcelCol,
  isSynAdaptGreyExcelCol,
  isSynAdaptFluoExcelCol,
  isSynAdaptFluoBandRow,
  isSynAdaptFluoIjOnlyCell,
  isSynSpotBlueCell,
  isSynProjHeaderGreenExcelCol,
  isSynProjHeaderBandExcelCol,
  isSynHdrSummaryTableCol,
  isSynHdrSummaryEvery3Col,
  SYN_PROJ_HDR_GREEN_DISPLAY_START,
  SYN_PROJ_HDR_GREEN_DISPLAY_END,
  SYN_AC_AN_TABLE_DISPLAY_START,
  SYN_AC_AN_TABLE_DISPLAY_END,
  SYN_AP_BB_TABLE_DISPLAY_START,
  SYN_AP_BB_TABLE_DISPLAY_END,
  SYN_BS_CE_TABLE_DISPLAY_START,
  SYN_BS_CE_TABLE_DISPLAY_END,
  SYN_BD_BO_TABLE_DISPLAY_START,
  SYN_BD_BO_TABLE_DISPLAY_END,
  SYN_CI_CY_TABLE_DISPLAY_START,
  SYN_CI_CY_TABLE_DISPLAY_END,
  SYN_DA_DP_TABLE_DISPLAY_START,
  SYN_DA_DP_TABLE_DISPLAY_END,
  SYN_DR_ED_TABLE_DISPLAY_START,
  SYN_DR_ED_TABLE_DISPLAY_END,
  SYN_EF_EQ_TABLE_DISPLAY_START,
  SYN_EF_EQ_TABLE_DISPLAY_END,
  SYN_ES_FE_TABLE_DISPLAY_START,
  SYN_ES_FE_TABLE_DISPLAY_END,
  SYN_FJ_FZ_TABLE_DISPLAY_START,
  SYN_FJ_FZ_TABLE_DISPLAY_END,
  isSynSpacerDisplayExcelCol,
  isSynSp2DisplayExcelCol,
  isSynSp2RestartDisplayExcelCol,
  isSynBuiltinPillarExcelCol,
  synPillarAccentClass,
} from './synthesisPerf.js';
import {
  isSynAdaptationSumCell,
  isSynAdaptBandExcelCol,
  isSynSumproductDataCell,
  isSynSectionSumDataCell,
  isSynAbDiffCell,
  isSynCalculatedMassCell,
  computeAdaptationRowSum,
  describeSynCellFormula,
  isSynVehicleMassCol,
  SYN_CALC_FIRST_ROW,
  SYN_ADAPTATION_SUM_ROW,
  SYN_ADAPTATION_SUM_FROM_ROW,
  synAdaptationSumRow,
  isSynAdaptationSumCol,
} from './synthesisCalc.js?v=syn-apbb8';

export {
  isSynSpacerDisplayExcelCol,
  isSynSp2DisplayExcelCol,
  isSynSp2RestartDisplayExcelCol,
  isSynBuiltinPillarExcelCol,
  synPillarAccentClass,
} from './synthesisPerf.js';

export const SYN_FILTER_ROWS = new Set([3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]);
export const SYN_LABEL_COL = 'F';
export const SYN_VEHICLE_COL_START = 'G';
/** Header panel colour band — display C…J (Excel H…O), rows 3–22. */
export const SYN_HDR_PANEL_COL_START = 'H';
export const SYN_HDR_PANEL_COL_END = 'O';

/**
 * Grid header letters (display columns). Excel A–E are hidden (F→A, G→B, …).
 * Do not use raw Excel letters here — e.g. display AB is Excel AG, not Excel AB (which shows as W).
 */
export const SYN_FORCE_WHITE_DISPLAY_COLS = [
  'AB',
  'AO',
  'BC',
  'BP',
  'BR',
  'CF',
  'CH',
  'CZ',
  'DQ',
];

/** Excel column keys for cell lookup (derived from display letters above). */
export const SYN_FORCE_WHITE_EXCEL_COLS = new Set(
  SYN_FORCE_WHITE_DISPLAY_COLS.map(displayToExcelCol)
);

export function isSynForceWhiteExcelCol(excelCol) {
  return SYN_FORCE_WHITE_EXCEL_COLS.has(excelCol);
}

export function synForceWhiteColClass(col) {
  return isSynForceWhiteExcelCol(col) ? 'syn-force-white-col' : '';
}
/** Rows 3–19, display C…J — bold + slightly larger text. */
export const SYN_HDR_PANEL_BOLD_LAST_ROW = 19;
/** Rows 18–19, display columns C–J (Excel H–O). */
export const SYN_HDR_METRIC_ROW_BG = '#ebf1de';
export const SYN_HDR_METRIC_BG_ROWS = new Set([18, 19]);
export const SYN_HDR_PANEL_GAP_COUNT = 2;
/** Blank rows before Date (display 1–2); pillars B/K/CG keep pillar fill. */
export const SYN_HDR_PANEL_TOP_GAP_COUNT = 2;

/** Filter band labels (Excel F); export sometimes omits shared-string cells. */
export const SYN_FILTER_ROW_LABELS = {
  3: 'Date',
  4: 'Project',
  5: 'Silhouette',
  6: 'Hybridization',
  7: 'Design plate',
  8: 'Seats',
  9: 'Technical spec',
  10: '',
  11: 'Pole',
  12: 'Energy',
  13: 'Technical pack',
  14: 'Trim',
};

/** Mass / portfolio summary rows between filter band and ADAPTATION (Excel F16–F22). */
export const SYN_METRIC_ROWS = new Set([15, 16, 17, 18, 19, 20, 21, 22]);
/** Metric panel rows 15–22: display C–J white band (not 18–19). */
export const SYN_METRIC_CJ_WHITE_ROWS = new Set([15, 16, 17, 20, 21, 22]);
/** Metric rows whose vehicle-column numbers are shown with a kg suffix. */
export const SYN_METRIC_KG_ROWS = new Set([16, 18, 19, 20]);
/** SP1 / SP2 pillars — vertical title rendered in a grid overlay (Excel G, P & CL). */
export const SYN_PILLAR_OVERLAY_COLS = new Set(['G', 'P', 'CL']);

/** Always-on pillar metadata (display B / K / CG) — not dependent on export merges. */
export const SYN_BUILTIN_PILLAR_META = {
  G: { title: 'SP1 TARGET', startRow: 3, endRow: 421 },
  P: { title: 'SP2 TARGET', startRow: 3, endRow: 421 },
  CL: { title: 'SP2 RESTART', startRow: 3, endRow: 421 },
};

export function getSynPillarMeta(col, pillarColumns) {
  if (pillarColumns && pillarColumns.has(col)) return pillarColumns.get(col);
  return SYN_BUILTIN_PILLAR_META[col] != null ? SYN_BUILTIN_PILLAR_META[col] : null;
}
export const SYN_METRIC_ROW_LABELS = {
  15: '',
  16: 'Curb mass:',
  17: 'PM pre-target',
  18: 'Curb mass: last update',
  19: 'Control',
  20: 'Portfolio',
  21: 'Forecast',
  22: '% Forecast/target',
};

/** English labels for Synthesis grid (values unchanged in sheet JSON). */
function synTranslateText(raw, col) {
  if (raw == null || raw === '') return '';
  const s = String(raw).trim();
  if (!s) return '';
  if (col === SYN_LABEL_COL) return translateSubsystemLabel(s);
  return translateValue(s);
}

export const SYN_HEADER_PANEL_LAST_ROW = 22;
/** Excel rows hidden in the web grid (continuous line numbers in the row column). */
export const SYN_SKIPPED_ROWS = [23, 24];
/** Reserved — no spacer rows between header panel (row 22) and ADAPTATION (row 25). */
export const SYN_PANEL_GAP_ROWS = new Set();

/** Empty spacer rows inside the header panel (between filter/metric blocks). */
export const SYN_HEADER_SPACER_ROWS = new Set([10, 15]);

/** @deprecated Use sequential displayRow from computeSynBodyRows. */
export function synDisplayRowNumber(excelRow) {
  let n = excelRow;
  for (const skip of SYN_SKIPPED_ROWS) {
    if (excelRow > skip) n--;
  }
  return n - (SYN_GRID_FIRST_ROW - 1);
}

/** Grey band for SP1 TARGET pillar (display column B). */
export const SYN_PILLAR_BG = '#bfbfbf';
/** ADAPTATION band rows 25–41 — fluorescent yellow (display D–G, I–J). */
export const SYN_ADAPT_FLUO_BG = '#ffff00';
/** Silhouette row 5 — Avenger like / P1X accent fills. */
export const SYN_VAL_AVENGER_BG = '#d8e4bc';
export const SYN_VAL_P1X_BG = '#c0504d';
/** Display column K (Excel P) + rows 3–4 project header (display M…AA). */
export const SYN_SP2_TARGET_BG = '#92d050';
export const SYN_COL_K_BG = SYN_SP2_TARGET_BG;
/** Display column CG (Excel CL) — SP2 RESTART pillar. */
export const SYN_SP2_RESTART_BG = '#c4d79b';
/** Spot highlights — same blue as Database sub-section bands. */
export const SYN_SPOT_BLUE_BG = '#00b0f0';
export const SYN_PROJ_HDR_GREEN_ROWS = new Set([3, 4, 13]);
export const SYN_PROJ_HDR_YELLOW_BG = '#ffff99';
export const SYN_PROJ_HDR_YELLOW_ROWS = new Set([5]);
export const SYN_PROJ_HDR_GREY_BG = '#bfbfbf';
export const SYN_PROJ_HDR_GREY_ROWS = new Set([11]);
/** HEV label red — same as syn-hdr-val-hev (#ff0000). */
export const SYN_HEV_RED_BG = '#ff0000';
/** BEV label green — rows 3–23 (Excel). */
export const SYN_BEV_TEXT_COLOR = '#00b050';
export const SYN_BEV_TEXT_FIRST_ROW = 3;
export const SYN_BEV_TEXT_LAST_ROW = 23;
export const SYN_PROJ_HDR_RED_BG = '#ffc7ce';
export const SYN_ROW19_MO_GREEN_BG = '#c6efce';
export const SYN_ROW19_PAA_RED_BG = '#ffc7ce';
export const SYN_ROW16_FLUO_BG = '#ffff00';
export const SYN_ROW25_MAA_GREEN_BG = '#92d050';
/** Display-row based green lines (same as rows 3–4: #92d050). */
export const SYN_DISPLAY_GREEN_BG = SYN_SP2_TARGET_BG;
export const SYN_DISPLAY_GREEN_ROWS = new Set([
  26, 42, 47, 52, 54, 59, 61, 63, 71, 73, 76, 84, 93, 97, 99, 165, 178, 181,
  205, 261, 276, 279, 288, 291, 296, 298, 307, 315, 319, 344, 353, 361, 368,
  372, 394, 398, 403, 407, 411, 415, 417, 422,
]);

/** Display-row based blue blocks — summary tables from display M (same span as grey). */
export const SYN_DISPLAY_BLUE_BG = SYN_SPOT_BLUE_BG;
export const SYN_DISPLAY_BLUE_ROWS = (() => {
  const s = new Set([
    78, 80, 85, 88, 91, 95, 106, 107, 111,
    210, 280, 287, 289, 412, 413, 414, 416,
  ]);
  for (let i = 113; i <= 119; i++) s.add(i);
  return s;
})();

/** Display-row based grey blocks — all summary tables from display M (M…AA, AC…AN, AP…BB, …). */
export const SYN_DISPLAY_GREY_MAA_BG = '#a6a6a6';
export const SYN_DISPLAY_GREY_MAA_ROWS = (() => {
  const s = new Set([
    27, 35, 38, 41, 50, 51, 77, 79, 81, 82, 83, 86, 87, 108, 109, 110, 112, 179, 180, 189,
    196, 199, 201, 202, 204, 221, 222, 224, 230, 239, 270, 272, 274, 275, 278,
    320, 321, 341, 342, 362, 373, 375, 378, 379, 410,
  ]);
  const addRange = (a, b) => {
    const lo = Math.min(a, b);
    const hi = Math.max(a, b);
    for (let i = lo; i <= hi; i++) s.add(i);
  };
  addRange(120, 123);
  addRange(125, 133);
  addRange(137, 157);
  addRange(161, 164);
  addRange(191, 193);
  addRange(252, 254);
  addRange(323, 339);
  addRange(418, 421);
  return s;
})();
/** Last Excel row exported / shown in the Synthesis grid. */
export const SYN_MAX_EXCEL_ROW = 422;
/** First Synthesis body row shown in the grid (Excel row 3 = Date). */
export const SYN_GRID_FIRST_ROW = 3;
/** SP1 pillar from Date (row 3) through last Synthesis row. */
export const SYN_PILLAR_FIRST_ROW = 3;
/** Vertical SP1/SP2 label: one letter every N Excel rows (2 = blank row between letters). */
export const SYN_PILLAR_LETTER_ROW_STEP = 2;
/** From this row, display columns C+ show 0,00 through last column (ADAPTATION band and below). */
export const SYN_ZERO_FILL_FIRST_ROW = 25;

/** Row 25 — display C…J preset values (Excel H…O); null = empty cell. */
const SYN_ROW_25_DISPLAY_VALUES = new Map([
  ['C', 35.8],
  ['D', null],
  ['E', 107.5],
  ['F', 104.0],
  ['G', 107.5],
  ['H', 36.6],
  ['I', 108.7],
  ['J', 115.2],
]);

export function synRow25PresetRaw(col) {
  if (!isSynHeaderPanelVehicleCol(col)) return undefined;
  const d = excelToDisplayCol(col);
  if (!SYN_ROW_25_DISPLAY_VALUES.has(d)) return undefined;
  return SYN_ROW_25_DISPLAY_VALUES.get(d);
}

/** O(1) row:col lookup while applying preset seeds (avoids cells.find on huge sheets). */
function synPresetCellIndex(cells) {
  const index = new Map();
  for (const c of cells) index.set(`${c.r}:${c.c}`, c);
  return index;
}

/** Row 25 vehicle cols are live SUM(27:40) — no static presets (see adaptationSumLocal). */
export function applySynRow25PresetCells(cells = []) {
  return cells;
}

function parseSynBandNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
}

/** Numeric value for ADAPTATION band rows 27–41 (presets unless userEdited). */
export function getSynAdaptBandNumeric(getCell, row, col) {
  const cell = getCell(row, col);
  if (cell && cell.userEdited) {
    return parseSynBandNum(displayValue(cell));
  }
  const preset = synRowCjPresetRaw(row, col);
  if (preset !== undefined) {
    if (preset == null || preset === '') return 0;
    return Number(preset);
  }
  const maaPreset = synRowMaaPresetRaw(row, col);
  if (maaPreset !== undefined) {
    if (maaPreset == null || maaPreset === '') return 0;
    return Number(maaPreset);
  }
  const raw = cell ? displayValue(cell) : '';
  if (raw && isSynNumericRaw(raw)) {
    return parseSynBandNum(raw);
  }
  if (cell && cell.v != null && cell.v !== '') {
    return parseSynBandNum(cell.v);
  }
  return 0;
}

/** Rows 27–422 — display C…J preset tables (null = empty cell). */
export const SYN_ADAPT_CJ_PRESET_FIRST_ROW = 27;
export const SYN_ADAPT_CJ_PRESET_LAST_ROW = 422;
const SYN_CJ_PRESET_DISPLAY_COLS = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

function synCjPresetRowMap(values) {
  const m = new Map();
  SYN_CJ_PRESET_DISPLAY_COLS.forEach((d, i) => m.set(d, values[i]));
  return m;
}

const SYN_CJ_PRESET_ZERO_ROW = synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0]);

function synCjPresetEntriesThrough116() {
  return [
  [27, synCjPresetRowMap([0, 30, 27.1, 30, 27.1, 0, 33.8, 33.8])],
  [28, synCjPresetRowMap([0, 6.8, 6.8, 6.8, 6.8, 0, 6.8, 6.8])],
  [29, synCjPresetRowMap([22.2, 20.1, 21.1, 20.1, 21.1, 21.6, 20.0, 21.0])],
  [30, synCjPresetRowMap([0, 0.9, 0.9, 0.9, 0.9, 0, 1.3, 1.3])],
  [31, synCjPresetRowMap([0, 21.1, 20.6, 21.1, 20.6, 0, 21.1, 20.6])],
  [32, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [33, synCjPresetRowMap([4.6, 4.2, 4.2, 4.2, 4.2, 4.6, 5.0, 5.0])],
  [34, synCjPresetRowMap([0, 0, 0, 0, 0, 0, null, null])],
  [35, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [36, synCjPresetRowMap([2, 2, 2, 2, 2, 2, 1.8, 1.8])],
  [37, synCjPresetRowMap([0, 0, 0, 0, 0, 0, null, null])],
  [38, synCjPresetRowMap([0, 9, 15, 9, 15, 0, 9, 15])],
  [39, synCjPresetRowMap([7.1, 9.9, 9.9, 9.9, 9.9, 8.5, 9.9, 9.9])],
  [40, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, null])],
  [41, synCjPresetRowMap([19.0, 15.6, 17.1, 15.6, 17.1, 18.6, 17.1, 17.1])],
  [42, synCjPresetRowMap([1.5, 3.0, 3.0, 3.0, 3.0, 0.7, 0.7, 0.7])],
  [43, synCjPresetRowMap([9.7, 7.0, 7.0, 7.0, 7.0, 10.0, 10.0, 10.0])],
  [44, synCjPresetRowMap([7.9, 5.6, 7.1, 5.6, 7.1, 7.9, 6.4, 6.4])],
  [45, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [46, synCjPresetRowMap([4.0, 4.0, 4.0, 4.0, 4.0, 4.8, 4.8, 4.8])],
  [47, synCjPresetRowMap([4.0, 4.0, 4.0, 4.0, 4.0, 4.8, 4.8, 4.8])],
  [48, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [49, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [50, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [51, synCjPresetRowMap([0, 8.6, 0, 8.6, 0, 0, 8.6, 0])],
  [52, synCjPresetRowMap([0, 8.6, 0, 8.6, 0, 0, 8.6, 0])],
  [53, synCjPresetRowMap([59.1, 59.1, 59.1, 53.6, 53.6, 58.9, 58.9, 58.9])],
  [54, synCjPresetRowMap([25.8, 25.8, 25.8, 22.6, 22.6, 25.2, 25.2, 25.2])],
  [55, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [56, synCjPresetRowMap([33.3, 33.3, 33.3, 31.0, 31.0, 33.7, 33.7, 33.7])],
  [57, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [58, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [59, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [60, synCjPresetRowMap([18.0, 12.0, 12.0, 12.0, 12.0, 17.6, 12.0, 12.0])],
  [61, synCjPresetRowMap([18.0, 12.0, 12.0, 12.0, 12.0, 17.6, 12.0, 12.0])],
  [62, synCjPresetRowMap([355.6, 14.4, 43.6, 14.4, 43.6, 456.9, 14.4, 43.6])],
  [63, synCjPresetRowMap([355.6, 12.0, 40.1, 12.0, 40.1, 409.0, 14.4, 40.1])],
  [64, synCjPresetRowMap([0, 2.4, 3.5, 2.4, 3.5, 47.9, 12.0, 3.5])],
  [65, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 2.4, 0])],
  [66, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [67, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [68, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [69, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [70, synCjPresetRowMap([9.8, 9.8, 9.8, 9.8, 9.8, 11.8, 11.8, 11.8])],
  [71, synCjPresetRowMap([9.8, 9.8, 9.8, 7.0, 7.0, 15.9, 15.9, 15.9])],
  [72, synCjPresetRowMap([15.3, 15.3, 15.3, 10.8, 10.8, 17.5, 17.5, 17.5])],
  [73, synCjPresetRowMap([15.3, 15.3, 15.3, 10.8, 10.8, 14.3, 14.3, 14.3])],
  [74, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [75, synCjPresetRowMap([0, 106.0, 117.0, 106.0, 117.0, 0, 106.0, 117.0])],
  [76, synCjPresetRowMap([0, 0, 0, null, null, 0, null, null])],
  [77, synCjPresetRowMap([0, 94.5, 105.5, 94.5, 105.5, 0, 94.5, 105.5])],
  [78, synCjPresetRowMap([0, 0, 0, null, null, 0, null, null])],
  [79, synCjPresetRowMap([0, 11.5, 11.5, 11.5, 11.5, 0, 11.5, 11.5])],
  [80, synCjPresetRowMap([0, 0, 0, null, null, 0, null, null])],
  [81, synCjPresetRowMap([0, 0, 0, null, null, 0, null, null])],
  [82, synCjPresetRowMap([0, 0, 0, null, null, 0, null, null])],
  [83, synCjPresetRowMap([328.8, 284.8, 284.8, 280.8, 280.8, 355.0, 301.0, 301.0])],
  [84, synCjPresetRowMap([39.2, 0, 0, 0, 0, 0, 0, 0])],
  [85, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [86, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [87, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [88, synCjPresetRowMap([328.8, 284.8, 284.8, 280.8, 280.8, 355.0, 311.0, 311.0])],
  [89, synCjPresetRowMap([15.1, 15.1, 15.1, 11.0, 11.0, 15.0, 15.0, 15.0])],
  [90, synCjPresetRowMap([1.5, 1.5, 1.5, 1.5, 1.5, 1.3, 1.3, 1.3])],
  [91, synCjPresetRowMap([13.6, 13.6, 13.6, 9.5, 9.5, 13.7, 13.7, 13.7])],
  [92, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [93, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [94, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [95, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [96, synCjPresetRowMap([5.5, 4.8, 4.8, 4.8, 4.8, 5.3, 5.3, 5.3])],
  [97, synCjPresetRowMap([5.5, 4.8, 4.8, 4.8, 4.8, 5.3, 5.3, 5.3])],
  [98, synCjPresetRowMap([8.7, 8.7, 8.7, 6.2, 6.2, 11.0, 11.0, 11.0])],
  [99, synCjPresetRowMap([6.9, 6.9, 6.9, 4.4, 4.4, 0, 0, 0])],
  [100, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [101, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [102, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [103, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [104, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [105, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [106, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [107, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [108, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [109, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [110, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [111, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [112, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [113, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [114, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [115, synCjPresetRowMap([0, 0, 0, 0, 0, 0, 0, 0])],
  [116, SYN_CJ_PRESET_ZERO_ROW],
  ];
}

function synCjPresetEntries117To157() {
  const entries = [];
  for (let row = 117; row <= 148; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([149, synCjPresetRowMap([1.8, 1.8, 1.8, 1.8, 1.8, 2.7, 2.7, 2.7])]);
  entries.push([150, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([151, SYN_CJ_PRESET_ZERO_ROW]);
  for (let row = 152; row <= 157; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  return entries;
}

function synCjPresetEntries158To198() {
  const entries = [];
  for (let row = 158; row <= 163; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([164, synCjPresetRowMap([13.2, 12.3, 12.3, 12.3, 12.3, 13.7, 12.7, 12.6])]);
  entries.push([165, synCjPresetRowMap([13.2, 12.3, 12.3, 12.3, 12.3, 0, 0, 0])]);
  for (let row = 166; row <= 180; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([181, synCjPresetRowMap([6.3, 5.5, 5.5, 5.5, 5.5, 5.6, 5.6, 5.6])]);
  entries.push([182, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([183, synCjPresetRowMap([1, 1, 1, 1, 1, 1, 1, 1])]);
  entries.push([184, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([185, synCjPresetRowMap([1, 1, 1, 1, 1, 0.4, 0.4, 0.4])]);
  entries.push([186, synCjPresetRowMap([2.4, 2.4, 2.4, 2.4, 2.4, 2.4, 2.4, 2.4])]);
  entries.push([187, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([188, synCjPresetRowMap([0.9, 0.1, 0.1, 0.1, 0.1, 0.8, 0.8, 0.8])]);
  for (let row = 189; row <= 198; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  return entries;
}

function synCjPresetEntries199To239() {
  const entries = [];
  entries.push([199, synCjPresetRowMap([1, 1, 1, 1, 1, 1, 1, 1])]);
  for (let row = 200; row <= 203; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([204, synCjPresetRowMap([9.3, 9.3, 9.3, 8.4, 8.4, 10.1, 10.1, 10.1])]);
  entries.push([205, synCjPresetRowMap([0.7, 0.7, 0.7, 0.7, 0.7, 0, 0, 0])]);
  for (let row = 206; row <= 207; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([208, synCjPresetRowMap([1.8, 1.8, 1.8, 1.8, 1.8, 2.4, 2.4, 2.4])]);
  entries.push([209, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([210, synCjPresetRowMap([1.2, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2, 1.2])]);
  entries.push([211, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([212, synCjPresetRowMap([0, 0, 0, 1, 1, 0, 0, 0])]);
  entries.push([213, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([214, synCjPresetRowMap([0.3, 0.3, 0.3, 0.3, 0.3, 0.2, 0.2, 0.2])]);
  for (let row = 215; row <= 228; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([229, synCjPresetRowMap([5.3, 5.3, 5.3, 3.3, 3.3, 5.2, 5.2, 5.2])]);
  for (let row = 230; row <= 239; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  return entries;
}

function synCjPresetEntries240To280() {
  const entries = [];
  for (let row = 240; row <= 263; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([264, synCjPresetRowMap([0.6, 0.6, 0.6, 0.6, 0.6, 0.5, 0.5, 0.5])]);
  entries.push([265, synCjPresetRowMap([0.4, 0.4, 0.4, 0.4, 0.4, 0.5, 0.5, 0.5])]);
  for (let row = 266; row <= 267; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([268, synCjPresetRowMap([0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1])]);
  for (let row = 269; row <= 276; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([277, synCjPresetRowMap([1.2, 0, 0, 0, 0, 1.1, 1.1, 1.1])]);
  entries.push([278, synCjPresetRowMap([1.2, 0, 0, 0, 0, 1.1, 1.1, 1.1])]);
  entries.push([279, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([280, synCjPresetRowMap([10.4, 10.4, 10.4, 9.4, 9.4, 20.3, 13.3, 13.3])]);
  return entries;
}

function synCjPresetEntries281To321() {
  const entries = [];
  entries.push([281, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([282, synCjPresetRowMap([0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1])]);
  for (let row = 283; row <= 286; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([287, synCjPresetRowMap([2.4, 2.4, 2.4, 2.4, 2.4, 8.9, 8.9, 8.9])]);
  entries.push([288, synCjPresetRowMap([7.8, 7.8, 7.8, 6.8, 6.8, 11.2, 11.2, 11.2])]);
  entries.push([289, synCjPresetRowMap([10.6, 7.5, 7.5, 7.5, 7.5, 9.9, 9.6, 9.6])]);
  entries.push([290, synCjPresetRowMap([8.1, 8.1, 8.1, 8.1, 8.1, 7.8, 7.8, 7.8])]);
  entries.push([291, synCjPresetRowMap([1, 1, 1, 1, 1, 0.8, 0.8, 0.8])]);
  entries.push([292, synCjPresetRowMap([3.7, 3.7, 3.7, 3.7, 3.7, 3.8, 3.8, 3.8])]);
  entries.push([293, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([294, synCjPresetRowMap([3.4, 3.4, 3.4, 3.4, 3.4, 3.3, 3.3, 3.3])]);
  entries.push([295, synCjPresetRowMap([4.9, 4.9, 4.9, 4.9, 4.9, 4.7, 4.7, 4.7])]);
  entries.push([296, synCjPresetRowMap([4.9, 4.9, 4.9, 4.9, 4.9, 0, 0, 0])]);
  entries.push([297, synCjPresetRowMap([33.8, 31.2, 31.2, 31.2, 31.2, 38.7, 39.7, 36.5])]);
  entries.push([298, synCjPresetRowMap([29.5, 27.0, 27.0, 27.0, 27.0, 33.7, 34.7, 30.5])]);
  entries.push([299, synCjPresetRowMap([4.2, 4.2, 4.2, 4.2, 4.2, 4.9, 4.9, 6.0])]);
  for (let row = 300; row <= 305; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  return entries;
}

/** Rows 306–321 — display C…J (capture fin de bande ADAPTATION). */
export const SYN_ADAPT_CJ_PRESET_ROW_306 = 306;
export const SYN_ADAPT_CJ_PRESET_ROW_321 = 321;

function synCjPresetEntries306To321() {
  const entries = [];
  entries.push([306, synCjPresetRowMap([58.4, 54.9, 54.9, 54.9, 54.9, 70.1, 63.5, 65.6])]);
  entries.push([307, synCjPresetRowMap([5.7, 5.3, 5.3, 5.3, 5.3, 5.7, 5.7, 5.3])]);
  entries.push([308, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([309, synCjPresetRowMap([4.3, 4.3, 4.3, 4.3, 4.3, 4.9, 4.9, 5.0])]);
  entries.push([310, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([311, synCjPresetRowMap([23.5, 24.2, 24.2, 24.2, 24.2, 28.4, 24.2, 24.2])]);
  entries.push([312, synCjPresetRowMap([24.9, 21.1, 21.1, 21.1, 21.1, 31.1, 28.7, 31.1])]);
  entries.push([313, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([314, synCjPresetRowMap([13.7, 13.7, 13.7, 13.7, 13.7, 15.9, 15.9, 15.9])]);
  entries.push([315, synCjPresetRowMap([1.3, 1.3, 1.3, 1.3, 1.3, 1.8, 1.8, 1.8])]);
  entries.push([316, synCjPresetRowMap([9.6, 9.6, 9.6, 9.6, 9.6, 10.9, 10.9, 10.9])]);
  entries.push([317, synCjPresetRowMap([2.7, 2.7, 2.7, 2.7, 2.7, 3.2, 3.2, 3.2])]);
  entries.push([318, synCjPresetRowMap([81.2, 94.0, 94.5, 94.0, 94.5, 104.6, 94.0, 94.5])]);
  entries.push([319, synCjPresetRowMap([0, null, null, null, null, 0, null, null])]);
  entries.push([320, synCjPresetRowMap([0, null, null, null, null, 0, null, null])]);
  entries.push([321, synCjPresetRowMap([0, 94.0, 94.5, 94.0, 94.5, 0, 94.0, 94.5])]);
  return entries;
}

/** Rows 402–422 — display C…J (capture fin de feuille). */
function synCjPresetEntries402To422() {
  const entries = [];
  entries.push([402, synCjPresetRowMap([11.3, 10.3, 10.3, 10.3, 10.3, 13.3, 14.8, 14.8])]);
  entries.push([403, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([404, synCjPresetRowMap([2.9, 2.9, 2.9, 2.9, 2.9, 5.9, 5.9, 5.4])]);
  entries.push([405, synCjPresetRowMap([8.5, 7.5, 7.5, 7.5, 7.5, 7.4, 7.4, 12.0])]);
  entries.push([406, synCjPresetRowMap([10.9, 10.9, 10.9, 10.9, 10.9, 12.0, 12.0, 12.0])]);
  entries.push([407, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([408, synCjPresetRowMap([10.9, 10.9, 10.9, 10.9, 10.9, 12.0, 12.0, 12.0])]);
  entries.push([409, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([410, synCjPresetRowMap([0, 0, 0, 0, 0, 0.2, 0.2, 0.2])]);
  entries.push([411, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([412, synCjPresetRowMap([0, 0, 0, 0, 0, 0.2, 0.2, 0.2])]);
  entries.push([413, SYN_CJ_PRESET_ZERO_ROW]);
  entries.push([414, synCjPresetRowMap([2.6, 2.6, 2.6, 2.6, 2.6, 2.6, 2.6, 2.6])]);
  entries.push([415, synCjPresetRowMap([2.6, 2.6, 2.6, 2.6, 2.6, 0, 0, 0])]);
  for (let row = 416; row <= 421; row++) {
    entries.push([row, SYN_CJ_PRESET_ZERO_ROW]);
  }
  entries.push([422, synCjPresetRowMap([0, 0, 0, 8.0, 8.0, 4.6, -18.6, -20.6])]);
  return entries;
}

const SYN_ROWS_CJ_PRESETS = new Map([
  ...synCjPresetEntriesThrough116(),
  ...synCjPresetEntries117To157(),
  ...synCjPresetEntries158To198(),
  ...synCjPresetEntries199To239(),
  ...synCjPresetEntries240To280(),
  ...synCjPresetEntries281To321(),
  ...synCjPresetEntries306To321(),
  ...synCjPresetEntries402To422(),
]);

export function synRowCjPresetRaw(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r)) return undefined;
  const rowMap = SYN_ROWS_CJ_PRESETS.get(r);
  if (!rowMap || !isSynHeaderPanelVehicleCol(col)) return undefined;
  const d = excelToDisplayCol(col);
  if (!rowMap.has(d)) return undefined;
  return rowMap.get(d);
}

/** Rows 27–321 — force display C…J presets (overrides legacy export). */
export function applySynRowsCjPresetCells(cells = []) {
  const index = synPresetCellIndex(cells);
  for (const [row, rowMap] of SYN_ROWS_CJ_PRESETS) {
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const key = `${row}:${col}`;
      let cell = index.get(key);
      if (cell && cell.userEdited) continue;
      if (value == null) {
        if (cell) {
          cell.v = '';
          delete cell.f;
        } else {
          cell = { r: row, c: col, v: '' };
          cells.push(cell);
          index.set(key, cell);
        }
        continue;
      }
      if (!cell) {
        cell = { r: row, c: col, v: String(value) };
        cells.push(cell);
        index.set(key, cell);
      } else {
        cell.v = String(value);
        delete cell.f;
      }
    }
  }
  return cells;
}

/** @deprecated Use applySynRowsCjPresetCells */
export function applySynRows27To41PresetCells(cells = []) {
  return applySynRowsCjPresetCells(cells);
}

/** Rows 3–117 (grid display 3–23 header + 26–118 body) — display M…AA presets. */
export const SYN_ADAPT_MAA_PRESET_FIRST_ROW = 3;
export const SYN_ADAPT_MAA_PRESET_LAST_ROW = 117;
/** Excel rows blank in M…AA (grid grey spacers in display 27, 35, 38, 41, 50, 51, 79–81, 84–85, 106–108, 110). */
export const SYN_MAA_GREY_SPACER_EXCEL_ROWS = new Set([
  26, 34, 37, 40, 49, 50, 78, 79, 80, 83, 84, 105, 106, 107, 109,
]);
const SYN_MAA_PRESET_DISPLAY_COLS = [
  'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA',
];

function synMaaPresetRowMap(values) {
  const m = new Map();
  SYN_MAA_PRESET_DISPLAY_COLS.forEach((d, i) => m.set(d, values[i]));
  return m;
}

function isSynMaaPresetExcelCol(col) {
  return isSynProjHeaderGreenExcelCol(col);
}

function isSynAcanExcelCol(col) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_START));
  const end = colToNum(displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_END));
  return n >= start && n <= end;
}

function isSynAcanPresetExcelCol(col) {
  return isSynAcanExcelCol(col);
}

export function isSynMaaGreySpacerExcelRow(row) {
  return SYN_MAA_GREY_SPACER_EXCEL_ROWS.has(Number(row));
}

const SYN_MAA_PRESET_NULL_ROW = synMaaPresetRowMap(
  SYN_MAA_PRESET_DISPLAY_COLS.map(() => null)
);
const SYN_MAA_PRESET_ZERO_ROW = synMaaPresetRowMap(
  SYN_MAA_PRESET_DISPLAY_COLS.map(() => 0)
);

/** Five triplet groups (M-O, P-R, S-U, V-X, Y-AA) — each value ×3. */
function synMaaPresetFiveGroups(g1, g2, g3, g4, g5) {
  const t = (v) => [v, v, v];
  return [...t(g1), ...t(g2), ...t(g3), ...t(g4), ...t(g5)];
}

/** Rows 54, 55, 57 — third triplet (S-U) uses three distinct values. */
function synMaaPresetSplitThirdTriplet(a, b, s, t, u, vA, vB) {
  return [...synMaaPresetFiveGroups(a, b, s, vA, vB).slice(0, 6), s, t, u, ...synMaaPresetFiveGroups(a, b, s, vA, vB).slice(9)];
}

function synMaaPresetAll(v) {
  return synMaaPresetRowMap(SYN_MAA_PRESET_DISPLAY_COLS.map(() => v));
}

function synMaaPresetTextRow(text) {
  return synMaaPresetAll(text);
}

function synMaaPresetSparse(pairs) {
  const m = new Map(SYN_MAA_PRESET_NULL_ROW);
  for (const [display, value] of pairs) m.set(display, value);
  return m;
}

/** Header panel rows 3–22 (grid display 3–18, gap 19, display 20–23). */
function synMaaPresetEntries3To22() {
  const N = SYN_MAA_PRESET_NULL_ROW;
  const txt = synMaaPresetTextRow;
  return [
    [3, txt('STLA/S')],
    [4, txt('SP2')],
    [5, txt('O3H')],
    [
      6,
      synMaaPresetRowMap([
        'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV',
        'MHEVP2', 'MHEVP2', 'MHEVP2', 'HEV', 'HEV', 'HEV',
      ]),
    ],
    [7, txt('EMEA')],
    [8, N],
    [
      9,
      synMaaPresetRowMap([
        'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'AWD', 'AWD', 'AWD',
        'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD',
      ]),
    ],
    [
      10,
      synMaaPresetRowMap([
        'HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'XR', 'XR', 'XR',
        'TT', 'TT', 'TT', 'TT', 'TT', 'TT',
      ]),
    ],
    [
      11,
      synMaaPresetRowMap([
        'HIGH_Range', 'HIGH_Range', 'HIGH_Range',
        'X_Range', 'X_Range', 'X_Range', 'X_Range', 'X_Range', 'X_Range',
        null, null, null, null, null, null,
      ]),
    ],
    [12, N],
    [13, txt('TARGET')],
    [
      14,
      synMaaPresetRowMap([
        'S', 'M', 'L', 'S', 'M', 'L', 'M', 'L', 'GSE', 'S', 'M', 'L', 'S', 'M', 'L',
      ]),
    ],
    [15, N],
    [
      16,
      synMaaPresetRowMap([
        1787, 1802, 1828, 1800, 1816, 1841, 1985, 2010, 2013, 1486, 1498, 1524,
        1496, 1509, 1535,
      ]),
    ],
    [
      17,
      synMaaPresetSparse([
        ['M', 1785],
        ['P', 1790],
        ['S', 1978],
        ['V', 1415],
        ['Y', 1425],
      ]),
    ],
    [
      18,
      synMaaPresetRowMap([
        1790, 1806, 1831, 1795, 1811, 1836, 1978, 2003, 2011, 1432, 1446, 1471,
        1467, 1480, 1506,
      ]),
    ],
    [
      19,
      synMaaPresetRowMap([
        -3.2, -3.2, -3.2, 5.0, 5.0, 5.0, 6.8, 6.8, 2.0, 53.4, 52.4, 52.4, 29.7,
        28.7, 28.7,
      ]),
    ],
    [
      20,
      synMaaPresetSparse([
        ['M', 1.6],
        ['P', 9.7],
        ['S', 6.6],
        ['V', 70.6],
        ['Y', 71.4],
      ]),
    ],
    [
      21,
      synMaaPresetSparse([
        ['V', 'Step 3'],
        ['W', 'Step 3'],
        ['X', 'Step 3'],
        ['Y', 'Step 3'],
        ['Z', 'Step 3'],
        ['AA', 'Step 3'],
      ]),
    ],
    [22, N],
  ];
}

/** Header panel rows 3–22 — display AC…AN presets (mirrors the AC–AN summary table). */
const SYN_ACAN_PRESET_DISPLAY_COLS = [
  'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
];

function synAcanPresetRowMap(values) {
  const m = new Map();
  SYN_ACAN_PRESET_DISPLAY_COLS.forEach((d, i) => m.set(d, values[i]));
  return m;
}

const SYN_ACAN_PRESET_NULL_ROW = synAcanPresetRowMap(
  SYN_ACAN_PRESET_DISPLAY_COLS.map(() => null)
);

function synAcanPresetTextRow(text) {
  return synAcanPresetRowMap(SYN_ACAN_PRESET_DISPLAY_COLS.map(() => text));
}

function synAcanPresetSparse(pairs) {
  const m = new Map(SYN_ACAN_PRESET_NULL_ROW);
  for (const [display, value] of pairs) m.set(display, value);
  return m;
}

function synAcanPresetEntries3To22() {
  const N = SYN_ACAN_PRESET_NULL_ROW;
  const txt = synAcanPresetTextRow;
  return [
    [3, txt('STLA/S')],
    [4, txt('SP2')],
    [5, txt('O3W')],
    [
      6,
      synAcanPresetRowMap([
        'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV',
        'MHEVP2', 'MHEVP2', 'MHEVP2',
        'HEV', 'HEV', 'HEV',
      ]),
    ],
    [7, txt('EMEA')],
    [8, N],
    [9, txt('FWD')],
    [
      10,
      synAcanPresetRowMap([
        'HR', 'HR', 'HR', 'XR', 'XR', 'XR',
        'TT', 'TT', 'TT', 'TT', 'TT', 'TT',
      ]),
    ],
    [
      11,
      synAcanPresetRowMap([
        'HIGH_Range', 'HIGH_Range', 'HIGH_Range',
        'X_Range', 'X_Range', 'X_Range',
        null, null, null, null, null, null,
      ]),
    ],
    [12, txt('SW')],
    [13, txt('TARGET')],
    [
      14,
      synAcanPresetRowMap([
        'S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L',
      ]),
    ],
    [15, N],
    [
      16,
      synAcanPresetRowMap([
        1827, 1841, 1866, 1840, 1854, 1879, 1526, 1537, 1562, 1537, 1547, 1572,
      ]),
    ],
    [
      17,
      synAcanPresetSparse([
        ['AC', 1827],
        ['AF', 1832],
        ['AI', 1467],
        ['AL', 1487],
      ]),
    ],
    [
      18,
      synAcanPresetRowMap([
        1830, 1844, 1869, 1835, 1849, 1874, 1473, 1484, 1509, 1507, 1519, 1544,
      ]),
    ],
    [
      19,
      synAcanPresetRowMap([
        -3.2, -3.2, -3.2, 5.0, 5.0, 5.0, 53.4, 52.4, 52.4, 29.7, 28.7, 28.7,
      ]),
    ],
    [
      20,
      synAcanPresetSparse([
        ['AC', -0.1],
        ['AF', 8.1],
        ['AI', 58.9],
        ['AL', 49.8],
      ]),
    ],
    [
      21,
      synAcanPresetSparse([
        ['AI', 'Step 3'],
        ['AJ', 'Step 3'],
        ['AK', 'Step 3'],
        ['AL', 'Step 3'],
        ['AM', 'Step 3'],
        ['AN', 'Step 3'],
      ]),
    ],
    [22, N],
  ];
}

const SYN_ROWS_ACAN_PRESETS = new Map([
  ...synAcanPresetEntries3To22(),
]);

export function synRowAcanPresetRaw(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r)) return undefined;
  const rowMap = SYN_ROWS_ACAN_PRESETS.get(r);
  if (!rowMap || !isSynAcanPresetExcelCol(col)) return undefined;
  const d = excelToDisplayCol(col);
  if (!rowMap.has(d)) return undefined;
  return rowMap.get(d);
}

/** Rows 3–22 — seed AC…AN presets (overrides legacy export unless userEdited). */
export function applySynRowsAcanPresetCells(cells = []) {
  const index = synPresetCellIndex(cells);
  for (const [row, rowMap] of SYN_ROWS_ACAN_PRESETS) {
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const key = `${row}:${col}`;
      let cell = index.get(key);
      if (cell && cell.userEdited) continue;
      if (value == null) {
        if (cell) {
          cell.v = '';
          delete cell.f;
          delete cell.bg;
        } else {
          cell = { r: row, c: col, v: '' };
          cells.push(cell);
          index.set(key, cell);
        }
        continue;
      }
      if (!cell) {
        cell = { r: row, c: col, v: String(value) };
        cells.push(cell);
        index.set(key, cell);
      } else {
        cell.v = String(value);
        delete cell.f;
        delete cell.bg;
      }
    }
  }
  return cells;
}

/** Header panel rows 3–22 — seed headerRows with AC…AN presets (keeps userEdited). */
export function applySynRowsAcanPresetHeaderRows(headerRows = {}) {
  for (const [row, rowMap] of SYN_ROWS_ACAN_PRESETS) {
    if (!isSynHeaderMaaPresetRow(row)) continue;
    const rowKey = String(row);
    if (!headerRows[rowKey]) headerRows[rowKey] = {};
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const existing = headerRows[rowKey][col];
      if (existing && existing.userEdited) continue;
      headerRows[rowKey][col] = {
        ...(existing || {}),
        v: value == null ? '' : String(value),
        f: undefined,
      };
      if (headerRows[rowKey][col].f === undefined) {
        delete headerRows[rowKey][col].f;
      }
    }
  }
  return headerRows;
}

/** Header panel rows 3–22 — display AP…BB presets (mirrors the AP–BB summary table). */
const SYN_APBB_PRESET_DISPLAY_COLS = [
  'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB',
];

function synApbbPresetRowMap(values) {
  const m = new Map();
  SYN_APBB_PRESET_DISPLAY_COLS.forEach((d, i) => m.set(d, values[i]));
  return m;
}

const SYN_APBB_PRESET_NULL_ROW = synApbbPresetRowMap(
  SYN_APBB_PRESET_DISPLAY_COLS.map(() => null)
);

function synApbbPresetTextRow(text) {
  return synApbbPresetRowMap(SYN_APBB_PRESET_DISPLAY_COLS.map(() => text));
}

function synApbbPresetSparse(pairs) {
  const m = new Map(SYN_APBB_PRESET_NULL_ROW);
  for (const [display, value] of pairs) m.set(display, value);
  return m;
}

function isSynApbbPresetExcelCol(col) {
  return isSynApBbTableCol(col);
}

function synApbbPresetEntries3To22() {
  const N = SYN_APBB_PRESET_NULL_ROW;
  const txt = synApbbPresetTextRow;
  return [
    [3, txt('STLA/S')],
    [4, txt('SP2')],
    [5, txt('P3S')],
    [
      6,
      synApbbPresetRowMap([
        'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV',
        'MHEVP2', 'MHEVP2', 'HEV', 'HEV',
      ]),
    ],
    [7, txt('EMEA')],
    [
      8,
      synApbbPresetSparse([
        ['AR', 'SBW'],
        ['AT', 'SBW'],
        ['AV', 'SBW'],
        ['AW', 'SBW'],
        ['AX', 'SBW'],
      ]),
    ],
    [
      9,
      synApbbPresetRowMap([
        'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'AWD',
        'FWD', 'FWD', 'FWD', 'FWD',
      ]),
    ],
    [
      10,
      synApbbPresetRowMap([
        'HR', 'HR', 'HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'XR',
        'TT', 'TT', 'TT', 'TT',
      ]),
    ],
    [
      11,
      synApbbPresetRowMap([
        'HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'HIGH_Range',
        'X_Range', 'X_Range', 'X_Range', 'X_Range',
        null, null, null, null,
      ]),
    ],
    [12, N],
    [13, txt('TARGET')],
    [
      14,
      synApbbPresetRowMap([
        'L2', 'L3', 'L3', 'GT', 'GT', 'L3', 'L3', 'GT', 'GTI',
        'L2', 'L3', 'L3', 'GT',
      ]),
    ],
    [15, N],
    [
      16,
      synApbbPresetRowMap([
        1806, 1822, 1834, 1711, 1724, 1835, 1848, 1738, 1876, 1497, 1516, 1489,
        1359,
      ]),
    ],
    [
      17,
      synApbbPresetSparse([
        ['AP', 1801],
        ['AU', 1823],
        ['AX', 2000],
        ['AY', 1415],
        ['BA', 1425],
      ]),
    ],
    [
      18,
      synApbbPresetRowMap([
        1809, 1825, 1837, 1839, 1851, 1830, 1843, 1857, 1989, 1444, 1463, 1485,
        1504,
      ]),
    ],
    [
      19,
      synApbbPresetRowMap([
        -2.7, -3.3, -3.3, -127.2, -127.2, 4.9, 4.9, -119.0, -112.5, 53.9, 52.3,
        4.4, -144.5,
      ]),
    ],
    [
      20,
      synApbbPresetSparse([
        ['AP', 4.9],
        ['AU', 12.3],
        ['AX', -123.8],
        ['AY', 82.4],
        ['BA', 63.9],
      ]),
    ],
    [
      21,
      synApbbPresetSparse([
        ['AY', 'Step 3'],
        ['AZ', 'Step 3'],
        ['BA', 'Step 3'],
        ['BB', 'Step 3'],
      ]),
    ],
    [22, N],
  ];
}

const SYN_ROWS_APBB_PRESETS = new Map([...synApbbPresetEntries3To22()]);

export function synRowApbbPresetRaw(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r)) return undefined;
  const rowMap = SYN_ROWS_APBB_PRESETS.get(r);
  if (!rowMap || !isSynApbbPresetExcelCol(col)) return undefined;
  const d = excelToDisplayCol(col);
  if (!rowMap.has(d)) return undefined;
  return rowMap.get(d);
}

/** Rows 3–22 — seed AP…BB presets (overrides legacy export unless userEdited). */
export function applySynRowsApbbPresetCells(cells = []) {
  const index = synPresetCellIndex(cells);
  for (const [row, rowMap] of SYN_ROWS_APBB_PRESETS) {
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const key = `${row}:${col}`;
      let cell = index.get(key);
      if (cell && cell.userEdited) continue;
      if (value == null) {
        if (cell) {
          cell.v = '';
          delete cell.f;
          delete cell.bg;
        } else {
          cell = { r: row, c: col, v: '' };
          cells.push(cell);
          index.set(key, cell);
        }
        continue;
      }
      if (!cell) {
        cell = { r: row, c: col, v: String(value) };
        cells.push(cell);
        index.set(key, cell);
      } else {
        cell.v = String(value);
        delete cell.f;
        delete cell.bg;
      }
    }
  }
  return cells;
}

/** Header panel rows 3–22 — seed headerRows with AP…BB presets (keeps userEdited). */
export function applySynRowsApbbPresetHeaderRows(headerRows = {}) {
  for (const [row, rowMap] of SYN_ROWS_APBB_PRESETS) {
    if (!isSynHeaderMaaPresetRow(row)) continue;
    const rowKey = String(row);
    if (!headerRows[rowKey]) headerRows[rowKey] = {};
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const existing = headerRows[rowKey][col];
      if (existing && existing.userEdited) continue;
      headerRows[rowKey][col] = {
        ...(existing || {}),
        v: value == null ? '' : String(value),
        f: undefined,
      };
      if (headerRows[rowKey][col].f === undefined) {
        delete headerRows[rowKey][col].f;
      }
    }
  }
  return headerRows;
}

const SYN_MAA_ROW_135_PATTERN = [
  0, 0, 13.5, 0, 0, 13.5, 0, 13.5, 13.5, 0, 0, 13.5, 0, 0, 13.5,
];

function synMaaPresetEntries76To117() {
  const Z = SYN_MAA_PRESET_ZERO_ROW;
  const N = SYN_MAA_PRESET_NULL_ROW;
  const G = synMaaPresetFiveGroups;
  const all = synMaaPresetAll;
  const p135 = () => synMaaPresetRowMap(SYN_MAA_ROW_135_PATTERN);
  return [
    [76, synMaaPresetRowMap(G(0, 0, 0, 94.5, 105.5))],
    [77, synMaaPresetRowMap(G(0, 0, 0, 11.5, 11.5))],
    [78, N],
    [79, N],
    [80, N],
    [81, synMaaPresetRowMap(G(413.1, 413.1, 413.1, 371.0, 371.0))],
    [82, synMaaPresetRowMap(G(42.1, 42.1, 42.1, 0, 0))],
    [83, N],
    [84, N],
    [85, Z],
    [86, all(371.0)],
    [87, all(12.0)],
    [88, all(1.4)],
    [89, all(10.6)],
    [90, p135()],
    [91, Z],
    [92, p135()],
    [93, Z],
    [94, all(4.8)],
    [95, all(4.8)],
    [96, all(6.7)],
    [97, Z],
    [98, Z],
    [99, all(6.7)],
    [100, Z],
    [101, Z],
    [102, Z],
    [103, Z],
    [104, Z],
    [105, N],
    [106, N],
    [107, N],
    [108, Z],
    [109, N],
    [110, Z],
    [111, Z],
    [112, Z],
    [113, Z],
    [114, Z],
    [115, Z],
    [116, Z],
    [117, Z],
  ];
}

function synMaaPresetEntries42To75() {
  const Z = SYN_MAA_PRESET_ZERO_ROW;
  const N = SYN_MAA_PRESET_NULL_ROW;
  const G = synMaaPresetFiveGroups;
  const all = synMaaPresetAll;
  return [
    [42, synMaaPresetRowMap(G(2.8, 2.8, 2.8, 3.0, 3.0))],
    [
      43,
      synMaaPresetRowMap([
        7.0, 7.5, 7.5, 7.0, 7.5, 7.5, 7.5, 7.5, 7.5, 7.0, 7.0, 7.0, 7.0, 7.0, 7.0,
      ]),
    ],
    [44, synMaaPresetRowMap(G(7.0, 7.0, 7.0, 5.6, 7.1))],
    [45, Z],
    [46, all(4.4)],
    [47, all(4.4)],
    [48, Z],
    [49, N],
    [50, N],
    [51, synMaaPresetRowMap(G(0, 0, 0, 8.6, 0))],
    [52, synMaaPresetRowMap(G(0, 0, 0, 8.6, 0))],
    [53, synMaaPresetRowMap(synMaaPresetSplitThirdTriplet(59.3, 59.7, 59.7, 60.5, 60.5, 59.3, 59.7))],
    [54, synMaaPresetRowMap(synMaaPresetSplitThirdTriplet(27.2, 27.6, 27.6, 28.0, 28.0, 27.2, 27.6))],
    [55, Z],
    [56, synMaaPresetRowMap(synMaaPresetSplitThirdTriplet(32.1, 32.1, 32.1, 32.5, 32.5, 32.1, 32.5))],
    [57, Z],
    [58, Z],
    [59, Z],
    [60, synMaaPresetRowMap(G(23.0, 23.0, 23.0, 12.0, 12.0))],
    [61, synMaaPresetRowMap(G(23.0, 23.0, 23.0, 12.0, 12.0))],
    [62, synMaaPresetRowMap(G(414.0, 419.6, 420.6, 14.4, 50.7))],
    [63, synMaaPresetRowMap(G(414.0, 419.6, 420.6, 12.0, 47.4))],
    [64, synMaaPresetRowMap(G(0, 0, 0, 2.4, 3.3))],
    [65, Z],
    [66, Z],
    [67, Z],
    [68, Z],
    [69, Z],
    [70, all(10.1)],
    [71, all(10.1)],
    [72, all(15.8)],
    [73, all(15.8)],
    [74, Z],
    [75, synMaaPresetRowMap(G(0, 0, 0, 106.0, 117.0))],
  ];
}

const SYN_ROWS_MAA_PRESETS = new Map([
  ...synMaaPresetEntries3To22(),
  [
    25,
    synMaaPresetRowMap([
      39.6, 44.7, 44.7, 39.6, 44.7, 44.7, 52.2, 52.2, 52.2, 108.7, 108.7, 108.7,
      81.5, 81.5, 81.5,
    ]),
  ],
  [26, SYN_MAA_PRESET_NULL_ROW],
  [
    27,
    synMaaPresetRowMap([
      0, 0, 0, 0, 0, 0, 0, 0, 0, 33.8, 33.8, 33.8, 0, 0, 0,
    ]),
  ],
  [
    28,
    synMaaPresetRowMap([
      0, 0, 0, 0, 0, 0, 0, 0, 0, 6.8, 6.8, 6.8, 6.8, 6.8, 6.8,
    ]),
  ],
  [
    29,
    synMaaPresetRowMap([
      22.1, 26.6, 26.6, 22.1, 26.6, 26.6, 31.9, 31.9, 31.9, 20.0, 20.0, 20.0,
      21.0, 21.0, 21.0,
    ]),
  ],
  [
    30,
    synMaaPresetRowMap([
      0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 1.3, 1.3, 1.3, 1.3, 1.3, 1.3,
    ]),
  ],
  [
    31,
    synMaaPresetRowMap([
      0, 0, 0, 0, 0, 0, 0, 0, 0, 21.1, 21.1, 21.1, 20.6, 20.6, 20.6,
    ]),
  ],
  [32, SYN_MAA_PRESET_ZERO_ROW],
  [
    33,
    synMaaPresetRowMap([
      5.4, 6.0, 6.0, 5.4, 6.0, 6.0, 6.0, 6.0, 6.0, 5.0, 5.0, 5.0, 5.0, 5.0, 5.0,
    ]),
  ],
  [34, SYN_MAA_PRESET_NULL_ROW],
  [35, SYN_MAA_PRESET_ZERO_ROW],
  [36, synMaaPresetRowMap(SYN_MAA_PRESET_DISPLAY_COLS.map(() => 1.8))],
  [37, SYN_MAA_PRESET_NULL_ROW],
  [
    38,
    synMaaPresetRowMap([
      0, 0, 0, 0, 0, 0, 0, 0, 0, 9.0, 9.0, 9.0, 15.0, 15.0, 15.0,
    ]),
  ],
  [
    39,
    synMaaPresetRowMap([
      9.9, 9.9, 9.9, 9.9, 9.9, 9.9, 12.1, 12.1, 12.1, 9.9, 9.9, 9.9, 9.9, 9.9,
      9.9,
    ]),
  ],
  [40, SYN_MAA_PRESET_NULL_ROW],
  [
    41,
    synMaaPresetRowMap([
      16.8, 17.3, 17.3, 16.8, 17.3, 17.3, 17.3, 17.3, 17.3, 15.6, 15.6, 15.6,
      17.1, 17.1, 17.1,
    ]),
  ],
  ...synMaaPresetEntries42To75(),
  ...synMaaPresetEntries76To117(),
]);

export function synRowMaaPresetRaw(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r)) return undefined;
  const rowMap = SYN_ROWS_MAA_PRESETS.get(r);
  if (!rowMap || !isSynMaaPresetExcelCol(col)) return undefined;
  const d = excelToDisplayCol(col);
  if (!rowMap.has(d)) return undefined;
  return rowMap.get(d);
}

function isSynHeaderMaaPresetRow(row) {
  const r = Number(row);
  return Number.isFinite(r) && r >= 3 && r <= SYN_HEADER_PANEL_LAST_ROW;
}

/** Rows 3–117 — seed M…AA presets (overrides legacy export unless userEdited). */
export function applySynRowsMaaPresetCells(cells = []) {
  const index = synPresetCellIndex(cells);
  for (const [row, rowMap] of SYN_ROWS_MAA_PRESETS) {
    // Body rows 26+ — live engine from Database (never seed static Excel totals).
    if (row >= SYN_ADAPTATION_SUM_ROW) continue;
    if (row === SYN_CALC_FIRST_ROW) continue;
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const key = `${row}:${col}`;
      let cell = index.get(key);
      if (cell && cell.userEdited) continue;
      if (value == null) {
        if (cell) {
          cell.v = '';
          delete cell.f;
        } else {
          cell = { r: row, c: col, v: '' };
          cells.push(cell);
          index.set(key, cell);
        }
        continue;
      }
      if (!cell) {
        cell = { r: row, c: col, v: String(value) };
        cells.push(cell);
        index.set(key, cell);
      } else {
        cell.v = String(value);
        delete cell.f;
      }
    }
  }
  return cells;
}

/** Header panel rows 3–22 — seed headerRows with M…AA presets (keeps userEdited). */
export function applySynRowsMaaPresetHeaderRows(headerRows = {}) {
  for (const [row, rowMap] of SYN_ROWS_MAA_PRESETS) {
    if (!isSynHeaderMaaPresetRow(row)) continue;
    const rowKey = String(row);
    if (!headerRows[rowKey]) headerRows[rowKey] = {};
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const existing = headerRows[rowKey][col];
      if (existing && existing.userEdited) continue;
      headerRows[rowKey][col] = {
        ...(existing || {}),
        v: value == null ? '' : String(value),
        f: undefined,
      };
      if (headerRows[rowKey][col].f === undefined) {
        delete headerRows[rowKey][col].f;
      }
    }
  }
  return headerRows;
}

/** First -ADAPTATION section row. */
export function findSynAdaptationRow(map, sheet) {
  const last = (sheet && sheet.lastRow) || SYN_MAX_EXCEL_ROW;
  for (let r = 2; r <= last; r++) {
    const raw =
      synLabel(map, r) || displayValue(getCell(map, r, SYN_LABEL_COL)) || '';
    const t = String(raw).trim().toUpperCase();
    if (t.startsWith('-ADAPTATION')) return r;
  }
  return 25;
}

/** First _ÉCHAPPEMENT row — vertical SP1 / SP2 label starts here. */
export function findSynEchappementRow(map, sheet) {
  const last = (sheet && sheet.lastRow) || SYN_MAX_EXCEL_ROW;
  for (let r = 2; r <= last; r++) {
    const raw =
      synLabel(map, r) || displayValue(getCell(map, r, SYN_LABEL_COL)) || '';
    const t = String(raw).trim().toUpperCase();
    if (t.includes('ECHAPPEMENT') || t.includes('ÉCHAPPEMENT')) return r;
  }
  const adapt = findSynAdaptationRow(map, sheet);
  for (let r = adapt; r <= Math.min(adapt + 15, last); r++) {
    const raw =
      synLabel(map, r) || displayValue(getCell(map, r, SYN_LABEL_COL)) || '';
    const t = String(raw).trim().toUpperCase();
    if (t.includes('ECHAPPEMENT') || t.includes('ÉCHAPPEMENT')) return r;
  }
  return 31;
}

/** "SP1 Target" → "SP1 TARGET"; "SP2 Restart" → "SP2 RESTART". */
export function normalizeSynPillarTitle(raw) {
  const t = String(raw != null ? raw : '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
  const target = t.match(/^(SP\d+)\s*TARGET$/i);
  if (target) return `${target[1]} TARGET`;
  const restart = t.match(/^(SP\d+)\s*RESTART$/i);
  if (restart) return `${restart[1]} RESTART`;
  return t;
}

/** One character per row; keeps the space between SP1 and TARGET. */
export function synPillarLettersFromTitle(title) {
  return [...normalizeSynPillarTitle(title)];
}

/** One letter every SYN_PILLAR_LETTER_ROW_STEP rows from Échappement downward. */
export function synPillarLetterForRow(row, col, pillarColumns, map, sheet) {
  const p = getSynPillarMeta(col, pillarColumns);
  if (!p) return '';
  const start = findSynEchappementRow(map, sheet);
  const letters = synPillarLettersFromTitle(p.title);
  const offset = row - start;
  if (offset < 0 || offset % SYN_PILLAR_LETTER_ROW_STEP !== 0) return '';
  const idx = offset / SYN_PILLAR_LETTER_ROW_STEP;
  if (idx >= letters.length) return '';
  const ch = letters[idx];
  return ch === ' ' ? '\u00a0' : ch;
}

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
  return Boolean(
    mergeMaps &&
      mergeMaps.covered &&
      mergeMaps.covered.has(`${row}:${col}`)
  );
}

export function getMergeSpan(mergeMaps, row, col) {
  const m = mergeMaps && mergeMaps.master ? mergeMaps.master.get(`${row}:${col}`) : null;
  if (!m) return null;
  // Large vertical merges break virtualized tables — keep layout value in master cell only.
  if (m.rowspan > 3 || m.colspan > 8) return null;
  return { rowspan: m.rowspan, colspan: m.colspan };
}

export function synLabel(map, row) {
  const v = displayValue(getCell(map, row, SYN_LABEL_COL));
  if (v) {
    const t = String(v).trim();
    if (isSynFilterRow(row) && /^\d{3,4}$/.test(t)) return '';
    return t;
  }
  const d = displayValue(getCell(map, row, 'D'));
  const t = d ? String(d).trim() : '';
  if (t.startsWith('-') || t.startsWith('_')) return t;
  return '';
}

export function synFilterRowLabel(map, row) {
  const fallback = SYN_FILTER_ROW_LABELS[row] != null ? SYN_FILTER_ROW_LABELS[row] : '';
  const fromCell = synLabel(map, row);
  if (fromCell) {
    const t = fromCell.trim();
    if (fallback && /^\d{3,5}$/.test(t)) return fallback;
    if (!fallback && /^\d{3,5}$/.test(t)) return '';
    return synTranslateText(fromCell, SYN_LABEL_COL);
  }
  return fallback;
}

export function isSynMetricRow(row) {
  return SYN_METRIC_ROWS.has(row);
}

export function isSynHeaderPanelRow(row) {
  if (row == null) return false;
  return row >= 3 && row <= SYN_HEADER_PANEL_LAST_ROW;
}

export function isSynPanelGapRow(row) {
  return SYN_PANEL_GAP_ROWS.has(row);
}

export function synMetricRowLabel(map, row) {
  const fromCell = synLabel(map, row);
  if (fromCell) return synTranslateText(fromCell, SYN_LABEL_COL);
  if (Object.prototype.hasOwnProperty.call(SYN_METRIC_ROW_LABELS, row)) {
    return SYN_METRIC_ROW_LABELS[row];
  }
  return '';
}

export function synHeaderPanelLabel(map, row) {
  if (isSynFilterRow(row)) return synFilterRowLabel(map, row);
  if (isSynMetricRow(row)) return synMetricRowLabel(map, row);
  return '';
}

/** True when the cell value is a plain number (not text like P1X). */
export function isSynNumericRaw(raw) {
  const s = String(raw != null ? raw : '').trim();
  if (!s) return false;
  return /^-?\d+([.,]\d+)?([eE][+-]?\d+)?$/.test(s);
}

/** Synthesis numbers — max 4 digit chars per cell, French comma (12 → 12,00 ; 1556 → 1556). */
export function formatSynNumericDisplay(raw) {
  if (raw == null || raw === '') return '';
  const s = String(raw).trim();
  if (!s || !isSynNumericRaw(s)) return s;
  let n = parseFloat(s.replace(',', '.'));
  if (!Number.isFinite(n)) return s;
  if (n === 0) return '0,00';

  const neg = n < 0;
  n = Math.abs(Number(n.toPrecision(4)));

  let intStr = String(Math.trunc(n));
  if (intStr.length > 4) {
    const div = 10 ** (intStr.length - 4);
    intStr = String(Math.round(n / div));
  }

  const intLen = intStr.length;
  const decPlaces = Math.max(0, 4 - intLen);

  if (decPlaces === 0) {
    return (neg ? '-' : '') + intStr;
  }

  const scaled = Math.round(n * 10 ** decPlaces) / 10 ** decPlaces;
  const i = Math.trunc(scaled);
  let frac = Math.round((scaled - i) * 10 ** decPlaces);
  if (frac >= 10 ** decPlaces) {
    return formatSynNumericDisplay(String((neg ? -1 : 1) * (i + 1)));
  }
  const fracStr = String(frac).padStart(decPlaces, '0');

  return (neg ? '-' : '') + String(i) + ',' + fracStr;
}

/** AP…BB — curb mass rows 16–18; control/portfolio variance rows 19–20. */
function formatSynHdrApbbMetricDisplay(row, raw) {
  const s = String(raw != null ? raw : '').trim();
  if (!s) return s;
  if (!isSynNumericRaw(s)) return s;
  const n = parseFloat(s.replace(',', '.'));
  if (!Number.isFinite(n)) return s;
  if (row === 16 || row === 17 || row === 18) {
    return `${formatSynNumericDisplay(n)} kg`;
  }
  if (row === 19 || row === 20) {
    const rounded = Math.round(n * 10) / 10;
    const neg = rounded < 0;
    const abs = Math.abs(rounded);
    const body = abs.toFixed(1).replace('.', ',');
    return (neg ? '-' : '') + body;
  }
  return formatSynNumericDisplay(n);
}

/** Header panel M…AA — mass rows with kg; control/portfolio with 1 decimal, no kg. */
function formatSynHdrMaaMetricDisplay(row, raw, col) {
  const s = String(raw != null ? raw : '').trim();
  if (!s) return s;
  if (!isSynNumericRaw(s)) return s;
  const n = parseFloat(s.replace(',', '.'));
  if (!Number.isFinite(n)) return s;
  if (col && isSynApBbTableCol(col)) {
    return formatSynHdrApbbMetricDisplay(row, raw);
  }
  if (row === 16 || row === 17 || row === 18) {
    return `${formatSynNumericDisplay(n)} kg`;
  }
  if (row === 19 || row === 20) {
    const rounded = Math.round(n * 10) / 10;
    const neg = rounded < 0;
    const abs = Math.abs(rounded);
    const body = abs.toFixed(1).replace('.', ',');
    return (neg ? '-' : '') + body;
  }
  return formatSynNumericDisplay(n);
}

/** Excel serial date → display (PM pre-target row). */
export function formatSynMetricValue(row, col, raw) {
  const s = String(raw != null ? raw : '').trim();
  if (!s || colToNum(col) < colToNum(SYN_VEHICLE_COL_START)) return s;
  if (row === 17 || row === 18) {
    const n = parseFloat(s.replace(',', '.'));
    if (Number.isFinite(n) && n > 30000 && n < 60000) {
      const d = new Date((n - 25569) * 86400000);
      if (!Number.isNaN(d.getTime())) {
        return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
      }
    }
  }
  if (isSynNumericRaw(s)) {
    const n = parseFloat(s.replace(',', '.'));
    if (!Number.isFinite(n)) return s;
    const formatted = formatSynNumericDisplay(n);
    if (SYN_METRIC_KG_ROWS.has(row) || Math.abs(n) >= 100) {
      return `${formatted} kg`;
    }
    return formatted;
  }
  if (
    SYN_METRIC_KG_ROWS.has(row) &&
    /\d/.test(s) &&
    !/\bkg\b/i.test(s)
  ) {
    const n = parseFloat(s.replace(/[^\d,.-]/g, '').replace(',', '.'));
    if (Number.isFinite(n)) {
      return `${formatSynNumericDisplay(String(n))} kg`;
    }
  }
  return s;
}

/** Metric data cells — Control green/red, Portfolio red, etc. */
export function synMetricCellClass(row, col, display) {
  if (colToNum(col) < colToNum(SYN_VEHICLE_COL_START)) return '';
  const s = String(display != null ? display : '').trim();
  if (!s) return '';
  if (isSynApBbTableCol(col)) {
    if (row === 19) {
      const n = parseFloat(s.replace(',', '.'));
      if (Number.isFinite(n)) {
        if (n < 0) return 'syn-metric-control-ok';
        if (n > 0) return 'syn-metric-control-warn';
      }
    }
    if (row === 20) return 'syn-metric-portfolio';
    if (row === 18) return 'syn-metric-stale syn-metric-mass';
    if (row === 16) return 'syn-metric-mass';
    if (row === 17) return 'syn-metric-pretarget';
    return '';
  }
  if (row === 19) {
    const n = parseFloat(s.replace(',', '.'));
    if (Number.isFinite(n)) {
      if (n < 0) return 'syn-metric-control-ok';
      if (n > 0) return 'syn-metric-control-warn';
    }
  }
  if (row === 20) return 'syn-metric-portfolio';
  if (row === 21) return 'syn-metric-forecast';
  if (row === 22) return 'syn-row22-hev-red';
  if (row === 18) return 'syn-metric-stale syn-metric-mass';
  if (row === 16) return 'syn-metric-mass';
  if (row === 17) return 'syn-metric-pretarget';
  return '';
}

/** Vertical SP1 TARGET / SP2 TARGET pillars (merged Excel columns). */
export function buildSynPillarColumns(sheet, cellMap) {
  const pillars = new Map();
  for (const [col, meta] of Object.entries(SYN_BUILTIN_PILLAR_META)) {
    pillars.set(col, { ...meta });
  }
  for (const m of (sheet && sheet.merges) || []) {
    if (m.colspan !== 1 || m.endRow - m.startRow < 50) continue;
    const cell = getCell(cellMap, m.startRow, m.startCol);
    const raw = String(cell && cell.v != null ? cell.v : '').trim();
    if (!raw || !/target|restart/i.test(raw)) continue;
    pillars.set(m.startCol, {
      title: normalizeSynPillarTitle(raw),
      startRow: m.startRow,
      endRow: m.endRow,
    });
  }
  return pillars;
}

export function isSynPillarCol(col, pillarColumns) {
  if (isSynBuiltinPillarExcelCol(col)) return true;
  return Boolean(pillarColumns && pillarColumns.has(col));
}

/** Pillar column — not on virtual panel gap rows (display 21–22). */
export function isSynPillarColAtRow(col, row, pillarColumns) {
  if (row == null) return false;
  return isSynPillarCol(col, pillarColumns) && row >= SYN_PILLAR_FIRST_ROW;
}

export function isSynPillarAnchor(row, col, pillarColumns) {
  const p = pillarColumns ? pillarColumns.get(col) : null;
  return Boolean(p && row === p.startRow);
}

export function isSynSp2PillarCol(col, pillarColumns) {
  if (!isSynPillarCol(col, pillarColumns)) return false;
  const meta = pillarColumns ? pillarColumns.get(col) : null;
  const title = meta && meta.title != null ? meta.title : '';
  return /^SP2\b/i.test(title);
}

/** Rows 3–4 & 13 — all summary tables from display M. */
export function isSynProjHeaderGreenCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || !SYN_PROJ_HDR_GREEN_ROWS.has(r)) return false;
  if (r === 13 && isSynApBbTableCol(col)) return false;
  return isSynHdrSummaryTableCol(col);
}

/** Rows 3–22 — cell text is STLA/S (or STLA-S) in any summary table from M. */
export function isSynStlaSlashGreenCell(row, col, display) {
  if (row == null || !isSynHeaderPanelRow(row)) return false;
  if (!isSynHdrSummaryTableCol(col)) return false;
  const t = String(display != null ? display : '').trim().toUpperCase();
  return t === 'STLA/S' || t === 'STLA-S';
}

/** Display column M+ — cell text contains "TARGET". */
export function isSynTargetTextGreenCell(col, display) {
  if (display == null || display === '') return false;
  const text = String(display).trim().toUpperCase();
  if (!text.includes('TARGET')) return false;
  return isSynHdrSummaryTableCol(col);
}

export function synProjHeaderGreenStyle() {
  return {
    background: SYN_SP2_TARGET_BG,
    backgroundColor: SYN_SP2_TARGET_BG,
    color: '#000',
  };
}

/** Row 5 — all summary tables from display M. */
export function isSynProjHeaderYellowCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || !SYN_PROJ_HDR_YELLOW_ROWS.has(r)) return false;
  if (r === 5 && isSynApBbTableCol(col)) return false;
  return isSynHdrSummaryTableCol(col);
}

export function synProjHeaderYellowStyle() {
  return {
    background: SYN_PROJ_HDR_YELLOW_BG,
    backgroundColor: SYN_PROJ_HDR_YELLOW_BG,
    color: '#000',
  };
}

/** Row 11 — all summary tables from display M; AP…BB HIGH_Range/X_Range on row 11. */
export function isSynProjHeaderGreyCol(row, col) {
  const r = Number(row);
  if (Number(r) === 11 && isSynApBbTableCol(col)) return true;
  if (!Number.isFinite(r) || !SYN_PROJ_HDR_GREY_ROWS.has(r)) return false;
  return isSynHdrSummaryTableCol(col);
}

export function synProjHeaderGreyStyle() {
  return {
    background: SYN_PROJ_HDR_GREY_BG,
    backgroundColor: SYN_PROJ_HDR_GREY_BG,
    color: '#000',
  };
}

/** Row 20 — extended summary tables only (DR…ED, EF…EQ, ES…FE, FJ…FZ). */
export function isSynProjHeaderRedCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || r !== 20) return false;
  if (!isSynProjHeaderBandExcelCol(col)) return false;
  const n = colToNum(col);
  const mStart = colToNum(displayToExcelCol('M'));
  const aaEnd = colToNum(displayToExcelCol('AA'));
  return n < mStart || n > aaEnd;
}

export function synProjHeaderRedStyle() {
  return {
    background: SYN_PROJ_HDR_RED_BG,
    backgroundColor: SYN_PROJ_HDR_RED_BG,
    color: '#9c0006',
  };
}

/** Row 19: display M–O green, P–AA red; AC–AE green, AF–AN red. */
export function isSynRow19MoGreenCol(row, col) {
  if (Number(row) !== 19) return false;
  if (isSynApBbTableCol(col)) return false;
  if (isSynAcanExcelCol(col)) {
    const n = colToNum(col);
    const start = colToNum(displayToExcelCol('AC'));
    const end = colToNum(displayToExcelCol('AE'));
    return n >= start && n <= end;
  }
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('M'));
  const end = colToNum(displayToExcelCol('O'));
  return n >= start && n <= end;
}

export function isSynRow19PaaRedCol(row, col) {
  if (Number(row) !== 19) return false;
  if (isSynApBbTableCol(col)) return false;
  if (isSynAcanExcelCol(col)) {
    const n = colToNum(col);
    const start = colToNum(displayToExcelCol('AF'));
    const end = colToNum(displayToExcelCol('AN'));
    return n >= start && n <= end;
  }
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('P'));
  const end = colToNum(displayToExcelCol('AA'));
  return n >= start && n <= end;
}

export function synRow19MoGreenStyle() {
  return {
    background: SYN_ROW19_MO_GREEN_BG,
    backgroundColor: SYN_ROW19_MO_GREEN_BG,
    color: '#000',
  };
}

export function synRow19PaaRedStyle() {
  return {
    background: SYN_ROW19_PAA_RED_BG,
    backgroundColor: SYN_ROW19_PAA_RED_BG,
    color: '#9c0006',
  };
}

/** Row 20 (portfolio): display M, P, S, V, Y — red; AC, AF, AI, AL — red. */
const SYN_ROW_20_PORTFOLIO_DISPLAY_COLS = new Set(['M', 'P', 'S', 'V', 'Y']);
const SYN_ROW_20_PORTFOLIO_YELLOW_DISPLAY_COLS = new Set(['V', 'Y']);
const SYN_ROW_20_ACAN_PORTFOLIO_DISPLAY_COLS = new Set(['AC', 'AF', 'AI', 'AL']);

export function isSynRow20PortfolioRedCol(row, col) {
  if (Number(row) !== 20) return false;
  if (isSynApBbTableCol(col)) return false;
  const d = excelToDisplayCol(col);
  if (isSynAcanExcelCol(col)) {
    return SYN_ROW_20_ACAN_PORTFOLIO_DISPLAY_COLS.has(d);
  }
  return (
    SYN_ROW_20_PORTFOLIO_DISPLAY_COLS.has(d) && isSynProjHeaderGreenExcelCol(col)
  );
}

export function isSynRow20PortfolioYellowCol(row, col) {
  if (Number(row) !== 20) return false;
  const d = excelToDisplayCol(col);
  if (isSynAcanExcelCol(col)) {
    return SYN_ROW_20_ACAN_PORTFOLIO_DISPLAY_COLS.has(d);
  }
  return SYN_ROW_20_PORTFOLIO_YELLOW_DISPLAY_COLS.has(d);
}

export function synRow20PortfolioRedStyle() {
  return {
    background: SYN_ROW19_PAA_RED_BG,
    backgroundColor: SYN_ROW19_PAA_RED_BG,
    color: '#9c0006',
  };
}

export function synRow20PortfolioYellowStyle() {
  return { color: '#ffff00' };
}

/** Row 21/22: Step 3 on display V–AA, AI–AN, or AY–BB. */
export function isSynRow21Step3Col(row, col, display) {
  if (!String(display != null ? display : '').trim()) return false;
  const n = colToNum(col);
  if (Number(row) === 21 && isSynApBbTableCol(col)) {
    const start = colToNum(displayToExcelCol('AY'));
    const end = colToNum(displayToExcelCol('BB'));
    return n >= start && n <= end;
  }
  if (Number(row) !== 21) return false;
  if (isSynAcanExcelCol(col)) {
    const start = colToNum(displayToExcelCol('AI'));
    const end = colToNum(displayToExcelCol('AN'));
    return n >= start && n <= end;
  }
  const start = colToNum(displayToExcelCol('V'));
  const end = colToNum(displayToExcelCol('AA'));
  return n >= start && n <= end;
}

function isSynEvery3FromDisplayStartCol(col, displayStart, displayEnd) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(displayStart));
  const end = colToNum(displayToExcelCol(displayEnd));
  if (n < start || n > end) return false;
  return (n - start) % 3 === 0;
}

export function synRow18MaaGreyStyle() {
  return { color: '#808080' };
}

/** @deprecated Use isSynRow20PortfolioRedCol */
export function isSynRow20MopGreenCol(_row, _col) {
  return false;
}

/** @deprecated Use isSynRow20PortfolioRedCol */
export function isSynRow20PaaRedCol(_row, _col) {
  return false;
}

export function synRow20MopGreenStyle() {
  return synRow20PortfolioRedStyle();
}

export function synRow20PaaRedStyle() {
  return synRow20PortfolioRedStyle();
}

/** @deprecated Step 3 row — no HEV band on Q–AA */
export function isSynRow21QaaHevRedCol(_row, _col) {
  return false;
}

/** @deprecated Step 3 row — M–P stay default */
export function isSynRow21MaaWhiteCol(_row, _col) {
  return false;
}

export function synRow21QaaHevRedStyle() {
  return synRow21MaaWhiteStyle();
}

/** @deprecated */
export function isSynRow22MaaHevRedCol(_row, _col, _display) {
  return false;
}

export function synRow22MaaHevRedStyle() {
  return synRow21MaaWhiteStyle();
}

export function synRow21MaaWhiteStyle() {
  return {
    background: '#fff',
    backgroundColor: '#fff',
    color: '#000',
  };
}

/** Row 25: display M–AA green. */
export function isSynRow25MaGreenCol(row, col) {
  if (Number(row) !== 25) return false;
  if (isSynSpacerDisplayExcelCol(col)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('M'));
  const end = colToNum(displayToExcelCol('AA'));
  return n >= start && n <= end;
}

export function synRow25MaGreenStyle() {
  return {
    background: SYN_ROW25_MAA_GREEN_BG,
    backgroundColor: SYN_ROW25_MAA_GREEN_BG,
    color: '#000',
  };
}

/**
 * Display M…AA — every 3 columns from M (M, P, S, V, Y).
 */
function isSynMaaEvery3FromMCol(col) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('M'));
  const end = colToNum(displayToExcelCol('AA'));
  if (n < start || n > end) return false;
  return (n - start) % 3 === 0;
}

/**
 * Row 16 (CURB MASS): fluo yellow on every column of each summary table —
 * full triplets (M–O, P–R, S–U… within M…AA; AC–AE, AF–AH… within AC…AN; etc.)
 * plus the first column of each table block (M, AC, AP, BS, …).
 */
export function isSynRow16FluoEvery3Col(row, col) {
  if (Number(row) !== 16) return false;
  if (isSynApBbTableCol(col)) return false;
  return isSynHdrSummaryTableCol(col);
}

/** Row 5 — AP…BB P3S silhouette band (black fill, white text). */
export function isSynApbbP3sBlackCol(row, col) {
  if (Number(row) !== 5) return false;
  return isSynApBbTableCol(col);
}

export function synApbbP3sBlackStyle() {
  return { backgroundColor: '#000000', color: '#ffffff' };
}

const SYN_APBB_ROW16_FLUO_DISPLAY_COLS = new Set(['AP', 'AU', 'AX', 'AY', 'BA']);

/** Row 16 — AP…BB curb mass (fluo yellow on populated summary columns). */
export function isSynApbbRow16FluoCol(row, col) {
  if (Number(row) !== 16) return false;
  if (!isSynApBbTableCol(col)) return false;
  return SYN_APBB_ROW16_FLUO_DISPLAY_COLS.has(excelToDisplayCol(col));
}

/** @deprecated use isSynApbbRow16FluoCol */
export function isSynApbbRow18FluoCol(row, col) {
  return isSynApbbRow16FluoCol(row, col);
}

/** @deprecated use isSynRow16FluoEvery3Col */
export function isSynRow16FluoEvery3FromMCol(row, col) {
  return isSynRow16FluoEvery3Col(row, col);
}

/** @deprecated use isSynRow16FluoEvery3Col */
export function isSynRow16FluoEvery3FromAcanCol(row, col) {
  return isSynRow16FluoEvery3Col(row, col);
}

/** Row 17: first col of each triplet — blue PM pre-target text. */
export function isSynRow17BlueEvery3Col(row, col) {
  if (Number(row) !== 17) return false;
  if (isSynApBbTableCol(col)) return false;
  return isSynHdrSummaryEvery3Col(col);
}

/** Row 17 — AP…BB PM pre-target (blue text on populated cells). */
export function isSynApbbRow17BlueCol(row, col) {
  if (Number(row) !== 17) return false;
  return isSynApBbTableCol(col);
}

/** @deprecated use isSynApbbRow17BlueCol */
export function isSynApbbRow19BlueCol(row, col) {
  return isSynApbbRow17BlueCol(row, col);
}

/** @deprecated use isSynRow17BlueEvery3Col */
export function isSynRow17MaaBlueCol(row, col) {
  return isSynRow17BlueEvery3Col(row, col);
}

/** @deprecated use isSynRow17BlueEvery3Col */
export function isSynRow17AcanBlueCol(row, col) {
  return isSynRow17BlueEvery3Col(row, col);
}

/** Row 18: all summary-table columns — grey reference mass text. */
export function isSynRow18GreyCol(row, col) {
  if (Number(row) !== 18) return false;
  if (isSynApBbTableCol(col)) return false;
  return isSynHdrSummaryTableCol(col);
}

/** Row 18 — AP…BB reference mass (grey text). */
export function isSynApbbRow18GreyCol(row, col) {
  if (Number(row) !== 18) return false;
  return isSynApBbTableCol(col);
}

/** @deprecated use isSynApbbRow18GreyCol */
export function isSynApbbRow20GreyCol(row, col) {
  return isSynApbbRow18GreyCol(row, col);
}

/** @deprecated use isSynRow18GreyCol */
export function isSynRow18MaaGreyCol(row, col) {
  return isSynRow18GreyCol(row, col);
}

/** @deprecated use isSynRow18GreyCol */
export function isSynRow18AcanGreyCol(row, col) {
  return isSynRow18GreyCol(row, col);
}

export function synRow17MaaBlueStyle() {
  return { color: '#0070c0' };
}

/** Row 17: display M, P, S, V, Y — blue PM pre-target text (not fluo). */
export function isSynRow17FluoEvery3FromMCol(_row, _col) {
  return false;
}

export function synRow16FluoStyle() {
  return {
    background: SYN_ROW16_FLUO_BG,
    backgroundColor: SYN_ROW16_FLUO_BG,
    color: '#000',
  };
}

export function isSynDisplayRowGreyMaaCol(displayRow, col) {
  if (!SYN_DISPLAY_GREY_MAA_ROWS.has(Number(displayRow))) return false;
  return isSynHdrSummaryTableCol(col);
}

export function synDisplayRowGreyMaaStyle() {
  return {
    background: SYN_DISPLAY_GREY_MAA_BG,
    backgroundColor: SYN_DISPLAY_GREY_MAA_BG,
    color: '#000',
  };
}

export function isSynDisplayRowGreenMaaCol(displayRow, col) {
  if (!SYN_DISPLAY_GREEN_ROWS.has(Number(displayRow))) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('M'));
  const end = colToNum(displayToExcelCol('AA'));
  return n >= start && n <= end;
}

/** Display row 26+ — AC…AN green bands (same rows & colour as M…AA). */
export function isSynDisplayRowGreenAcanCol(displayRow, col) {
  const dr = Number(displayRow);
  if (!Number.isFinite(dr) || dr < 26) return false;
  if (!SYN_DISPLAY_GREEN_ROWS.has(dr)) return false;
  return isSynAcAnTableCol(col);
}

export function synDisplayRowGreenMaaStyle() {
  return {
    background: SYN_DISPLAY_GREEN_BG,
    backgroundColor: SYN_DISPLAY_GREEN_BG,
    color: '#000',
  };
}

export function synDisplayRowGreenAcanStyle() {
  return synDisplayRowGreenMaaStyle();
}

export function isSynDisplayRowBlueCol(displayRow, col) {
  if (!SYN_DISPLAY_BLUE_ROWS.has(Number(displayRow))) return false;
  return isSynHdrSummaryTableCol(col);
}

export function synDisplayRowBlueStyle() {
  return {
    background: SYN_DISPLAY_BLUE_BG,
    backgroundColor: SYN_DISPLAY_BLUE_BG,
    color: '#000',
  };
}

/** Display column M and beyond (Excel R+). */
export function isSynDisplayColFromM(col) {
  return colToNum(col) >= colToNum(displayToExcelCol('M'));
}

/** Yellow fluo rows/bands → green from display M (sections L1, row 16 CURB MASS, row 18 AP…BB). */
export function isSynYellowFluoGreenFromMCol(row, col, map, sheet) {
  if (!isSynDisplayColFromM(col)) return false;
  if (isSynSpacerDisplayExcelCol(col)) return false;
  if (isSynForceWhiteExcelCol(col)) return false;
  if (synRowStyleClass(map, row, sheet) === 'syn-row-section') return true;
  if (isSynRow16FluoEvery3Col(row, col)) return true;
  if (isSynApbbRow16FluoCol(row, col)) return true;
  return false;
}

export function synYellowFluoGreenFromMStyle() {
  return synDisplayRowGreenMaaStyle();
}

/** Rows 3–22 — bold vertical line right of display C…I (between columns C–J). */
export function isSynHdrCjDividerRightCol(row, col) {
  if (!isSynHeaderPanelRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(SYN_HDR_PANEL_COL_START);
  const endJ = colToNum(SYN_HDR_PANEL_COL_END);
  return n >= start && n < endJ;
}

/** Rows 3–22 — bold vertical line right of display M…Z (between columns M–AA). */
export function isSynHdrMaDividerRightCol(row, col) {
  if (!isSynHeaderPanelRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_PROJ_HDR_GREEN_DISPLAY_START));
  const endAa = colToNum(displayToExcelCol(SYN_PROJ_HDR_GREEN_DISPLAY_END));
  return n >= start && n < endAa;
}

/** Bold vertical divider — right edge of display L (header panel rows 3–22). */
export function isSynHdrLmDividerRightCol(row, col) {
  if (!isSynHeaderPanelRow(row)) return false;
  return isSynSpacerDisplayExcelCol(col);
}

/** Bold vertical divider — left edge of display M (header panel rows 3–22). */
export function isSynHdrLmDividerLeftCol(row, col) {
  if (!isSynHeaderPanelRow(row)) return false;
  return col === displayToExcelCol(SYN_PROJ_HDR_GREEN_DISPLAY_START);
}

/** Bold vertical divider — right edge of display AA (header panel rows 3–22). */
export function isSynHdrAaDividerRightCol(row, col) {
  if (!isSynHeaderPanelRow(row)) return false;
  return col === displayToExcelCol(SYN_PROJ_HDR_GREEN_DISPLAY_END);
}

/** Display-row gap between Excel 18–19 — keep table grid borders on a white band. */
export function isSynHeaderTableGapBetweenEntry(entry) {
  return Boolean(entry && entry.gapBetween);
}

/** Excel row for table-frame helpers — gap row uses a panel row so col checks apply. */
export function resolveSynHeaderTableEntryRow(entry) {
  if (!entry) return null;
  if (entry.gapBetween) return SYN_GRID_FIRST_ROW;
  return entry.excelRow != null ? entry.excelRow : null;
}

/** M / AA frame — Excel rows 3–22 + display gap row 19. */
export function isSynHdrLmDividerLeftEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrLmDividerLeftCol(row, col);
}

export function isSynHdrAaDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrAaDividerRightCol(row, col);
}

export function isSynHdrCjDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrCjDividerRightCol(row, col);
}

export function isSynHdrMaDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrMaDividerRightCol(row, col);
}

export function isSynHdrLmDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrLmDividerRightCol(row, col);
}

/** @deprecated use isSynHdrLmDividerRightCol */
export function isSynHdrLmDividerCol(row, col) {
  return isSynHdrLmDividerRightCol(row, col);
}

/** Excel rows 3–22 — AC…AN summary table (display row 23 = Excel 22). */
export function isSynAcAnTableRow(row) {
  return isSynHeaderPanelRow(row);
}

export function isSynAcAnTableCol(col) {
  return isSynAcanExcelCol(col);
}

export function isSynAcAnTableCell(row, col) {
  return isSynAcAnTableRow(row) && isSynAcAnTableCol(col);
}

/** Bold vertical line right of display AC…AM (between columns AC–AN). */
export function isSynHdrAcAnDividerRightCol(row, col) {
  if (!isSynAcAnTableRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_START));
  const endAn = colToNum(displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_END));
  return n >= start && n < endAn;
}

/** Bold vertical line — left edge of display AC. */
export function isSynHdrAcAnDividerLeftCol(row, col) {
  if (!isSynAcAnTableRow(row)) return false;
  return col === displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_START);
}

/** Bold vertical line — right edge of display AN. */
export function isSynHdrAcAnDividerRightEdgeCol(row, col) {
  if (!isSynAcAnTableRow(row)) return false;
  return col === displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_END);
}

export function isSynAcAnTableCellEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynAcAnTableCell(row, col);
}

export function isSynHdrAcAnDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrAcAnDividerRightCol(row, col);
}

export function isSynHdrAcAnDividerLeftEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrAcAnDividerLeftCol(row, col);
}

export function isSynHdrAcAnDividerRightEdgeEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrAcAnDividerRightEdgeCol(row, col);
}

/** Excel rows 3–22 — AP…BB summary table (display row 23 = Excel 22). */
export function isSynApBbTableRow(row) {
  return isSynHeaderPanelRow(row);
}

export function isSynApBbTableCol(col) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_AP_BB_TABLE_DISPLAY_START));
  const end = colToNum(displayToExcelCol(SYN_AP_BB_TABLE_DISPLAY_END));
  return n >= start && n <= end;
}

export function isSynApBbTableCell(row, col) {
  return isSynApBbTableRow(row) && isSynApBbTableCol(col);
}

/** Bold vertical line right of display AP…BA (between columns AP–BB). */
export function isSynHdrApBbDividerRightCol(row, col) {
  if (!isSynApBbTableRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_AP_BB_TABLE_DISPLAY_START));
  const endBb = colToNum(displayToExcelCol(SYN_AP_BB_TABLE_DISPLAY_END));
  return n >= start && n < endBb;
}

/** Bold vertical line — left edge of display AP. */
export function isSynHdrApBbDividerLeftCol(row, col) {
  if (!isSynApBbTableRow(row)) return false;
  return col === displayToExcelCol(SYN_AP_BB_TABLE_DISPLAY_START);
}

/** Bold vertical line — right edge of display BB. */
export function isSynHdrApBbDividerRightEdgeCol(row, col) {
  if (!isSynApBbTableRow(row)) return false;
  return col === displayToExcelCol(SYN_AP_BB_TABLE_DISPLAY_END);
}

export function isSynApBbTableCellEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynApBbTableCell(row, col);
}

export function isSynHdrApBbDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrApBbDividerRightCol(row, col);
}

export function isSynHdrApBbDividerLeftEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrApBbDividerLeftCol(row, col);
}

export function isSynHdrApBbDividerRightEdgeEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrApBbDividerRightEdgeCol(row, col);
}

/** Excel rows 3–22 — BS…CE summary table (display row 23 = Excel 22). */
export function isSynBsCeTableRow(row) {
  return isSynHeaderPanelRow(row);
}

export function isSynBsCeTableCol(col) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_BS_CE_TABLE_DISPLAY_START));
  const end = colToNum(displayToExcelCol(SYN_BS_CE_TABLE_DISPLAY_END));
  return n >= start && n <= end;
}

export function isSynBsCeTableCell(row, col) {
  return isSynBsCeTableRow(row) && isSynBsCeTableCol(col);
}

export function isSynHdrBsCeDividerRightCol(row, col) {
  if (!isSynBsCeTableRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_BS_CE_TABLE_DISPLAY_START));
  const endCe = colToNum(displayToExcelCol(SYN_BS_CE_TABLE_DISPLAY_END));
  return n >= start && n < endCe;
}

export function isSynHdrBsCeDividerLeftCol(row, col) {
  if (!isSynBsCeTableRow(row)) return false;
  return col === displayToExcelCol(SYN_BS_CE_TABLE_DISPLAY_START);
}

export function isSynHdrBsCeDividerRightEdgeCol(row, col) {
  if (!isSynBsCeTableRow(row)) return false;
  return col === displayToExcelCol(SYN_BS_CE_TABLE_DISPLAY_END);
}

export function isSynBsCeTableCellEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynBsCeTableCell(row, col);
}

export function isSynHdrBsCeDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrBsCeDividerRightCol(row, col);
}

export function isSynHdrBsCeDividerLeftEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrBsCeDividerLeftCol(row, col);
}

export function isSynHdrBsCeDividerRightEdgeEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrBsCeDividerRightEdgeCol(row, col);
}

/** Excel rows 3–22 — BD…BO summary table (display row 23 = Excel 22). */
export function isSynBdBoTableRow(row) {
  return isSynHeaderPanelRow(row);
}

export function isSynBdBoTableCol(col) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_BD_BO_TABLE_DISPLAY_START));
  const end = colToNum(displayToExcelCol(SYN_BD_BO_TABLE_DISPLAY_END));
  return n >= start && n <= end;
}

export function isSynBdBoTableCell(row, col) {
  return isSynBdBoTableRow(row) && isSynBdBoTableCol(col);
}

export function isSynHdrBdBoDividerRightCol(row, col) {
  if (!isSynBdBoTableRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_BD_BO_TABLE_DISPLAY_START));
  const endBo = colToNum(displayToExcelCol(SYN_BD_BO_TABLE_DISPLAY_END));
  return n >= start && n < endBo;
}

export function isSynHdrBdBoDividerLeftCol(row, col) {
  if (!isSynBdBoTableRow(row)) return false;
  return col === displayToExcelCol(SYN_BD_BO_TABLE_DISPLAY_START);
}

export function isSynHdrBdBoDividerRightEdgeCol(row, col) {
  if (!isSynBdBoTableRow(row)) return false;
  return col === displayToExcelCol(SYN_BD_BO_TABLE_DISPLAY_END);
}

export function isSynBdBoTableCellEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynBdBoTableCell(row, col);
}

export function isSynHdrBdBoDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrBdBoDividerRightCol(row, col);
}

export function isSynHdrBdBoDividerLeftEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrBdBoDividerLeftCol(row, col);
}

export function isSynHdrBdBoDividerRightEdgeEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrBdBoDividerRightEdgeCol(row, col);
}

/** Excel rows 3–22 — CI…CY summary table (display row 23 = Excel 22). */
export function isSynCiCyTableRow(row) {
  return isSynHeaderPanelRow(row);
}

export function isSynCiCyTableCol(col) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_CI_CY_TABLE_DISPLAY_START));
  const end = colToNum(displayToExcelCol(SYN_CI_CY_TABLE_DISPLAY_END));
  return n >= start && n <= end;
}

export function isSynCiCyTableCell(row, col) {
  return isSynCiCyTableRow(row) && isSynCiCyTableCol(col);
}

export function isSynHdrCiCyDividerRightCol(row, col) {
  if (!isSynCiCyTableRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_CI_CY_TABLE_DISPLAY_START));
  const endCy = colToNum(displayToExcelCol(SYN_CI_CY_TABLE_DISPLAY_END));
  return n >= start && n < endCy;
}

export function isSynHdrCiCyDividerLeftCol(row, col) {
  if (!isSynCiCyTableRow(row)) return false;
  return col === displayToExcelCol(SYN_CI_CY_TABLE_DISPLAY_START);
}

export function isSynHdrCiCyDividerRightEdgeCol(row, col) {
  if (!isSynCiCyTableRow(row)) return false;
  return col === displayToExcelCol(SYN_CI_CY_TABLE_DISPLAY_END);
}

export function isSynCiCyTableCellEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynCiCyTableCell(row, col);
}

export function isSynHdrCiCyDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrCiCyDividerRightCol(row, col);
}

export function isSynHdrCiCyDividerLeftEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrCiCyDividerLeftCol(row, col);
}

export function isSynHdrCiCyDividerRightEdgeEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrCiCyDividerRightEdgeCol(row, col);
}

/** Excel rows 3–22 — DA…DP summary table (display row 23 = Excel 22). */
export function isSynDaDpTableRow(row) {
  return isSynHeaderPanelRow(row);
}

export function isSynDaDpTableCol(col) {
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_DA_DP_TABLE_DISPLAY_START));
  const end = colToNum(displayToExcelCol(SYN_DA_DP_TABLE_DISPLAY_END));
  return n >= start && n <= end;
}

export function isSynDaDpTableCell(row, col) {
  return isSynDaDpTableRow(row) && isSynDaDpTableCol(col);
}

export function isSynHdrDaDpDividerRightCol(row, col) {
  if (!isSynDaDpTableRow(row)) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_DA_DP_TABLE_DISPLAY_START));
  const endDp = colToNum(displayToExcelCol(SYN_DA_DP_TABLE_DISPLAY_END));
  return n >= start && n < endDp;
}

export function isSynHdrDaDpDividerLeftCol(row, col) {
  if (!isSynDaDpTableRow(row)) return false;
  return col === displayToExcelCol(SYN_DA_DP_TABLE_DISPLAY_START);
}

export function isSynHdrDaDpDividerRightEdgeCol(row, col) {
  if (!isSynDaDpTableRow(row)) return false;
  return col === displayToExcelCol(SYN_DA_DP_TABLE_DISPLAY_END);
}

export function isSynDaDpTableCellEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynDaDpTableCell(row, col);
}

export function isSynHdrDaDpDividerRightEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrDaDpDividerRightCol(row, col);
}

export function isSynHdrDaDpDividerLeftEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrDaDpDividerLeftCol(row, col);
}

export function isSynHdrDaDpDividerRightEdgeEntry(entry, col) {
  const row = resolveSynHeaderTableEntryRow(entry);
  if (row == null) return false;
  return isSynHdrDaDpDividerRightEdgeCol(row, col);
}

function createSynSummaryTableApi(displayStart, displayEnd) {
  const excelStart = () => displayToExcelCol(displayStart);
  const excelEnd = () => displayToExcelCol(displayEnd);

  function isTableRow(row) {
    return isSynHeaderPanelRow(row);
  }
  function isTableCol(col) {
    const n = colToNum(col);
    const start = colToNum(excelStart());
    const end = colToNum(excelEnd());
    return n >= start && n <= end;
  }
  function isTableCell(row, col) {
    return isTableRow(row) && isTableCol(col);
  }
  function isDividerRightCol(row, col) {
    if (!isTableRow(row)) return false;
    const n = colToNum(col);
    const start = colToNum(excelStart());
    const end = colToNum(excelEnd());
    return n >= start && n < end;
  }
  function isDividerLeftCol(row, col) {
    if (!isTableRow(row)) return false;
    return col === excelStart();
  }
  function isDividerRightEdgeCol(row, col) {
    if (!isTableRow(row)) return false;
    return col === excelEnd();
  }
  function isTableCellEntry(entry, col) {
    const row = resolveSynHeaderTableEntryRow(entry);
    if (row == null) return false;
    return isTableCell(row, col);
  }
  function isDividerRightEntry(entry, col) {
    const row = resolveSynHeaderTableEntryRow(entry);
    if (row == null) return false;
    return isDividerRightCol(row, col);
  }
  function isDividerLeftEntry(entry, col) {
    const row = resolveSynHeaderTableEntryRow(entry);
    if (row == null) return false;
    return isDividerLeftCol(row, col);
  }
  function isDividerRightEdgeEntry(entry, col) {
    const row = resolveSynHeaderTableEntryRow(entry);
    if (row == null) return false;
    return isDividerRightEdgeCol(row, col);
  }
  return {
    isTableRow,
    isTableCol,
    isTableCell,
    isDividerRightCol,
    isDividerLeftCol,
    isDividerRightEdgeCol,
    isTableCellEntry,
    isDividerRightEntry,
    isDividerLeftEntry,
    isDividerRightEdgeEntry,
  };
}

const _synDrEdTable = createSynSummaryTableApi(
  SYN_DR_ED_TABLE_DISPLAY_START,
  SYN_DR_ED_TABLE_DISPLAY_END
);
export const isSynDrEdTableRow = _synDrEdTable.isTableRow;
export const isSynDrEdTableCol = _synDrEdTable.isTableCol;
export const isSynDrEdTableCell = _synDrEdTable.isTableCell;
export const isSynHdrDrEdDividerRightCol = _synDrEdTable.isDividerRightCol;
export const isSynHdrDrEdDividerLeftCol = _synDrEdTable.isDividerLeftCol;
export const isSynHdrDrEdDividerRightEdgeCol = _synDrEdTable.isDividerRightEdgeCol;
export const isSynDrEdTableCellEntry = _synDrEdTable.isTableCellEntry;
export const isSynHdrDrEdDividerRightEntry = _synDrEdTable.isDividerRightEntry;
export const isSynHdrDrEdDividerLeftEntry = _synDrEdTable.isDividerLeftEntry;
export const isSynHdrDrEdDividerRightEdgeEntry = _synDrEdTable.isDividerRightEdgeEntry;

const _synEfEqTable = createSynSummaryTableApi(
  SYN_EF_EQ_TABLE_DISPLAY_START,
  SYN_EF_EQ_TABLE_DISPLAY_END
);
export const isSynEfEqTableRow = _synEfEqTable.isTableRow;
export const isSynEfEqTableCol = _synEfEqTable.isTableCol;
export const isSynEfEqTableCell = _synEfEqTable.isTableCell;
export const isSynHdrEfEqDividerRightCol = _synEfEqTable.isDividerRightCol;
export const isSynHdrEfEqDividerLeftCol = _synEfEqTable.isDividerLeftCol;
export const isSynHdrEfEqDividerRightEdgeCol = _synEfEqTable.isDividerRightEdgeCol;
export const isSynEfEqTableCellEntry = _synEfEqTable.isTableCellEntry;
export const isSynHdrEfEqDividerRightEntry = _synEfEqTable.isDividerRightEntry;
export const isSynHdrEfEqDividerLeftEntry = _synEfEqTable.isDividerLeftEntry;
export const isSynHdrEfEqDividerRightEdgeEntry = _synEfEqTable.isDividerRightEdgeEntry;

const _synEsFeTable = createSynSummaryTableApi(
  SYN_ES_FE_TABLE_DISPLAY_START,
  SYN_ES_FE_TABLE_DISPLAY_END
);
export const isSynEsFeTableRow = _synEsFeTable.isTableRow;
export const isSynEsFeTableCol = _synEsFeTable.isTableCol;
export const isSynEsFeTableCell = _synEsFeTable.isTableCell;
export const isSynHdrEsFeDividerRightCol = _synEsFeTable.isDividerRightCol;
export const isSynHdrEsFeDividerLeftCol = _synEsFeTable.isDividerLeftCol;
export const isSynHdrEsFeDividerRightEdgeCol = _synEsFeTable.isDividerRightEdgeCol;
export const isSynEsFeTableCellEntry = _synEsFeTable.isTableCellEntry;
export const isSynHdrEsFeDividerRightEntry = _synEsFeTable.isDividerRightEntry;
export const isSynHdrEsFeDividerLeftEntry = _synEsFeTable.isDividerLeftEntry;
export const isSynHdrEsFeDividerRightEdgeEntry = _synEsFeTable.isDividerRightEdgeEntry;

const _synFjFzTable = createSynSummaryTableApi(
  SYN_FJ_FZ_TABLE_DISPLAY_START,
  SYN_FJ_FZ_TABLE_DISPLAY_END
);
export const isSynFjFzTableRow = _synFjFzTable.isTableRow;
export const isSynFjFzTableCol = _synFjFzTable.isTableCol;
export const isSynFjFzTableCell = _synFjFzTable.isTableCell;
export const isSynHdrFjFzDividerRightCol = _synFjFzTable.isDividerRightCol;
export const isSynHdrFjFzDividerLeftCol = _synFjFzTable.isDividerLeftCol;
export const isSynHdrFjFzDividerRightEdgeCol = _synFjFzTable.isDividerRightEdgeCol;
export const isSynFjFzTableCellEntry = _synFjFzTable.isTableCellEntry;
export const isSynHdrFjFzDividerRightEntry = _synFjFzTable.isDividerRightEntry;
export const isSynHdrFjFzDividerLeftEntry = _synFjFzTable.isDividerLeftEntry;
export const isSynHdrFjFzDividerRightEdgeEntry = _synFjFzTable.isDividerRightEdgeEntry;

/** Display column L (Excel Q) — white gutter, all body rows. */
export function synSpacerColClass(col) {
  return isSynSpacerDisplayExcelCol(col) ? 'syn-spacer-col-l' : '';
}

export function synSpacerColStyle(col) {
  if (synSpacerColClass(col)) {
    return { backgroundColor: '#fff', color: '#000' };
  }
  return null;
}

export function isSynFilterRow(row, sheet) {
  const rows = (sheet && sheet.filterRows) || [...SYN_FILTER_ROWS];
  return rows.includes(row);
}

export function isSynSeparatorRow(row) {
  return row === 4;
}

/** Yellow L1 band (AILES, ALTERNATEUR, -ADAPTATION, BOUCLIER AR…). */
export function isSynL1SectionLabel(label) {
  const t = String(label != null ? label : '').trim();
  if (!t) return false;
  if (t.startsWith('-')) return true;
  return isSectionLabel(t);
}

export function isSynL2SubsectionLabel(label) {
  const t = String(label != null ? label : '').trim();
  return t.startsWith('_');
}

/**
 * Per-map cache of the L1-section structure (which rows carry an L1 title).
 * `isSynBetweenL1SectionRows` used to rescan every row of the sheet for every
 * cell on every render — the dominant scroll hotspot in Synthesis (~50% of CPU).
 * The L1 layout only depends on the label column, so we compute it once per map
 * (single pass) and answer subsequent queries in O(1).
 */
const _synL1Cache = new WeakMap();

function getSynL1Info(map, sheet) {
  const last =
    sheet && sheet.effectiveLastRow != null
      ? sheet.effectiveLastRow
      : sheet && sheet.lastRow != null
        ? sheet.lastRow
        : SYN_MAX_EXCEL_ROW;
  if (!map) return { last, set: new Set(), min: null, max: null };
  let info = _synL1Cache.get(map);
  if (info && info.last === last) return info;
  const set = new Set();
  let min = null;
  let max = null;
  for (let r = SYN_GRID_FIRST_ROW; r <= last; r++) {
    if (isSynL1SectionLabel(synLabel(map, r))) {
      set.add(r);
      if (min === null) min = r;
      max = r;
    }
  }
  info = { last, set, min, max };
  _synL1Cache.set(map, info);
  return info;
}

/** Memoized row→class map (a row's band class is identical across all columns). */
const _synRowClassCache = new WeakMap();

/** Drop cached structure for a map (call when a label/band cell is edited). */
export function clearSynStructureCache(map) {
  if (!map) return;
  _synL1Cache.delete(map);
  _synRowClassCache.delete(map);
}

/** Rows under an L1 title until the next L1 (e.g. between AILES and ALTERNATEUR). */
export function isSynBetweenL1SectionRows(map, row, sheet) {
  const info = getSynL1Info(map, sheet);
  if (info.set.has(row)) return false;
  if (info.min === null || info.max === null) return false;
  return info.min < row && info.max > row;
}

export function isSynSectionLabelRow(map, row, sheet) {
  const band =
    (sheet && sheet.rowBands && sheet.rowBands[String(row)]) ||
    (sheet && sheet.rowBands && sheet.rowBands[row]);
  if (band === 'section' || band === 'subsection') return true;
  const label = synLabel(map, row);
  if (!label) return false;
  return label.startsWith('-') || label.startsWith('_');
}

/** Row band from Excel label column (F) export, else label-prefix fallback. */
export function synRowStyleClass(map, row, sheet) {
  let rowCache = map ? _synRowClassCache.get(map) : null;
  if (rowCache) {
    const hit = rowCache.get(row);
    if (hit !== undefined) return hit;
  }
  const cls = computeSynRowStyleClass(map, row, sheet);
  if (map) {
    if (!rowCache) {
      rowCache = new Map();
      _synRowClassCache.set(map, rowCache);
    }
    rowCache.set(row, cls);
  }
  return cls;
}

function computeSynRowStyleClass(map, row, sheet) {
  if (isSynPanelGapRow(row)) return 'syn-panel-gap';
  const band =
    (sheet && sheet.rowBands && sheet.rowBands[String(row)]) ||
    (sheet && sheet.rowBands && sheet.rowBands[row]);
  if (band === 'separator') return 'syn-row-separator';
  if (band === 'filter') return 'syn-row-filter';
  if (band === 'section') return 'syn-row-section';
  if (band === 'subsection') return 'syn-row-subsection';
  if (isSynSeparatorRow(row)) return 'syn-row-separator';
  if (isSynMetricRow(row)) return 'syn-row-metric';
  if (isSynFilterRow(row, sheet)) return 'syn-row-filter';
  const label = synLabel(map, row);
  if (isSynL1SectionLabel(label)) return 'syn-row-section';
  if (isSynL2SubsectionLabel(label)) return 'syn-row-subsection';
  if (isSynBetweenL1SectionRows(map, row, sheet)) return 'syn-row-subsection';
  if (label.startsWith('-')) return 'syn-row-section';
  if (label.startsWith('_')) return 'syn-row-subsection';
  return 'syn-row-data';
}

/** Row bands coloured via CSS (same palette as Database page). */
const SYN_STRUCTURE_ROW_CLASSES = new Set([
  'syn-row-section',
  'syn-row-subsection',
  'syn-row-separator',
  'syn-row-filter',
]);

export function synCellInlineStyle(cell, map, row, col, sheet, pillarColumns) {
  const style = {};
  if (col === SYN_LABEL_COL) {
    style.textAlign = 'left';
    style.color = '#000';
  }
  const forceWhiteClass = synForceWhiteColClass(col);
  if (forceWhiteClass) {
    style.background = '#fff';
    style.backgroundColor = '#fff';
    style.color = '#000';
    return style;
  }
  const spacerStyle = synSpacerColStyle(col);
  if (spacerStyle) {
    Object.assign(style, spacerStyle);
    return style;
  }
  if (
    isSynPillarColAtRow(col, row, pillarColumns) ||
    (row >= SYN_PILLAR_FIRST_ROW && isSynSp2RestartDisplayExcelCol(col))
  ) {
    const bg = isSynSp2DisplayExcelCol(col)
      ? SYN_SP2_TARGET_BG
      : isSynSp2RestartDisplayExcelCol(col)
        ? SYN_SP2_RESTART_BG
        : SYN_PILLAR_BG;
    style.background = bg;
    style.backgroundColor = bg;
    style.color = '#000';
    style.border = 'none';
    return style;
  }
  const disp = synDisplayValue(cell, map, row, col, sheet, pillarColumns);
  const bevStyle = synHdrBevTextStyle(row, disp);
  if (bevStyle) {
    Object.assign(style, bevStyle);
    Object.assign(style, synHeaderPanelBoldFontStyle(row, col) || {});
    return style;
  }
  const energyStyle = synHdrEnergyValueStyle(row, col, disp);
  if (energyStyle) {
    Object.assign(style, energyStyle);
    Object.assign(style, synHeaderPanelBoldFontStyle(row, col) || {});
    return style;
  }
  const spotBlueStyle = synSpotBlueColStyle(row, col);
  if (spotBlueStyle) {
    Object.assign(style, spotBlueStyle);
    Object.assign(style, synHeaderPanelBoldFontStyle(row, col) || {});
    return style;
  }
  const greyStyle = synFilterGreyColStyle(row, col);
  if (greyStyle) {
    Object.assign(style, greyStyle);
    Object.assign(style, synHeaderPanelBoldFontStyle(row, col) || {});
    return style;
  }
  if (
    SYN_HDR_METRIC_BG_ROWS.has(row) &&
    isSynHeaderPanelVehicleCol(col) &&
    !isSynSpacerDisplayExcelCol(col)
  ) {
    style.background = SYN_HDR_METRIC_ROW_BG;
    style.backgroundColor = SYN_HDR_METRIC_ROW_BG;
    style.color = '#000';
    Object.assign(style, synHeaderPanelBoldFontStyle(row, col) || {});
    return style;
  }
  const adaptStyle = synAdaptBandColStyle(row, col, pillarColumns);
  if (adaptStyle) {
    Object.assign(style, adaptStyle);
    return style;
  }
  const metricWhiteStyle = synMetricCjWhiteColStyle(row, col);
  if (metricWhiteStyle) {
    Object.assign(style, metricWhiteStyle);
    return style;
  }
  if (
    isSynStlaSlashGreenCell(row, col, disp) ||
    isSynProjHeaderGreenCol(row, col) ||
    isSynTargetTextGreenCell(col, disp)
  ) {
    style.background = SYN_SP2_TARGET_BG;
    style.backgroundColor = SYN_SP2_TARGET_BG;
    style.color = '#000';
    return style;
  }
  if (isSynProjHeaderRedCol(row, col)) {
    style.background = SYN_PROJ_HDR_RED_BG;
    style.backgroundColor = SYN_PROJ_HDR_RED_BG;
    style.color = '#9c0006';
    return style;
  }
  const raw = cell ? displayValue(cell) : '';
  const accentStyle = synCellAccentStyle(raw);
  if (accentStyle) {
    Object.assign(style, accentStyle);
    Object.assign(style, synHeaderPanelBoldFontStyle(row, col) || {});
    return style;
  }
  Object.assign(style, synHeaderPanelBoldFontStyle(row, col) || {});
  const rowCls = synRowStyleClass(map, row, sheet);
  if (SYN_STRUCTURE_ROW_CLASSES.has(rowCls)) {
    const rowColor =
      sheet && sheet.matrixColors
        ? sheet.matrixColors[row] != null
          ? sheet.matrixColors[row]
          : sheet.matrixColors[String(row)]
        : undefined;
    if (
      rowColor &&
      (rowCls === 'syn-row-section' || rowCls === 'syn-row-subsection')
    ) {
      style.backgroundColor = rowColor;
    }
    return style;
  }
  if (cell && cell.b) style.fontWeight = '700';
  Object.assign(style, synAdaptCjBoldFontStyle(row, col, disp, sheet) || {});
  return style;
}

function isSynHeaderPanelVehicleCol(col) {
  const n = colToNum(col);
  return (
    n >= colToNum(SYN_HDR_PANEL_COL_START) && n <= colToNum(SYN_HDR_PANEL_COL_END)
  );
}

/** Rows 3–19, display columns C–J; M…AA; AC…AN; AP…BB; BS…CE; BD…BO; CI…CY; DA…DP. */
export function isSynHeaderPanelBoldCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || r < SYN_GRID_FIRST_ROW) return false;
  if (col === SYN_LABEL_COL) return false;
  if (
    r <= SYN_HEADER_PANEL_LAST_ROW &&
    (isSynAcAnTableCol(col) ||
      isSynApBbTableCol(col) ||
      isSynBsCeTableCol(col) ||
      isSynBdBoTableCol(col) ||
      isSynCiCyTableCol(col) ||
      isSynDaDpTableCol(col) ||
      isSynDrEdTableCol(col) ||
      isSynEfEqTableCol(col) ||
      isSynEsFeTableCol(col) ||
      isSynFjFzTableCol(col))
  ) {
    return true;
  }
  if (r > SYN_HDR_PANEL_BOLD_LAST_ROW) return false;
  const n = colToNum(col);
  if (n >= colToNum(SYN_HDR_PANEL_COL_START) && n <= colToNum(SYN_HDR_PANEL_COL_END)) {
    return true;
  }
  return isSynProjHeaderGreenExcelCol(col);
}

export function synHeaderPanelBoldFontStyle(row, col) {
  if (!isSynHeaderPanelBoldCol(row, col)) return null;
  return { fontWeight: '700', fontSize: '12px' };
}

/** Rows 25…last — display C–J with a displayed value (not empty / not 0,00 placeholder). */
export function isSynAdaptCjBoldRow(row, sheet) {
  const r = Number(row);
  if (!Number.isFinite(r) || r < SYN_ZERO_FILL_FIRST_ROW) return false;
  const last =
    sheet && sheet.effectiveLastRow != null
      ? sheet.effectiveLastRow
      : sheet && sheet.lastRow != null
        ? sheet.lastRow
        : SYN_MAX_EXCEL_ROW;
  return r <= last;
}

export function isSynAdaptCjValueBold(row, col, displayText, sheet) {
  if (!isSynAdaptCjBoldRow(row, sheet)) return false;
  if (!isSynHeaderPanelVehicleCol(col)) return false;
  if (isSynSpacerDisplayExcelCol(col)) return false;
  const t = String(displayText != null ? displayText : '').trim();
  if (!t || t === '0,00') return false;
  return true;
}

export function synAdaptCjBoldFontStyle(row, col, displayText, sheet) {
  if (!isSynAdaptCjValueBold(row, col, displayText, sheet)) return null;
  return { fontWeight: '700' };
}

/** Avenger like (#d8e4bc) / P1X (#c0504d) — exact cell value (row 5 Silhouette). */
export function synCellAccentClass(displayText) {
  const u = String(displayText != null ? displayText : '')
    .trim()
    .toUpperCase();
  if (!u) return '';
  if (u === 'P1X') return 'syn-val-p1x';
  if (u === 'AVENGER LIKE') return 'syn-val-avenger-like';
  return '';
}

export function synCellAccentStyle(displayText) {
  const cls = synCellAccentClass(displayText);
  if (cls === 'syn-val-p1x') {
    return { backgroundColor: SYN_VAL_P1X_BG, color: '#fff' };
  }
  if (cls === 'syn-val-avenger-like') {
    return { backgroundColor: SYN_VAL_AVENGER_BG, color: '#000' };
  }
  return null;
}

/** Row 3–14 energy / drivetrain labels in M…AA. */
export function isSynHdrBevTextRow(row) {
  const r = Number(row);
  return (
    Number.isFinite(r) &&
    r >= SYN_BEV_TEXT_FIRST_ROW &&
    r <= SYN_BEV_TEXT_LAST_ROW
  );
}

export function synHdrBevTextClass(row, displayText) {
  if (!isSynHdrBevTextRow(row)) return '';
  const v = String(displayText != null ? displayText : '').trim().toUpperCase();
  if (!v.includes('BEV')) return '';
  return 'syn-hdr-val-bev';
}

export function synHdrBevTextStyle(row, displayText) {
  if (!synHdrBevTextClass(row, displayText)) return null;
  return { color: SYN_BEV_TEXT_COLOR };
}

function escapeSynDisplayHtml(text) {
  return String(text != null ? text : '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Read-only display: highlight "BEV" in green within cell text (rows 3–23). */
export function synHdrBevDisplayHtml(row, displayText) {
  if (!isSynHdrBevTextRow(row)) return null;
  const text = String(displayText != null ? displayText : '');
  if (!/BEV/i.test(text)) return null;
  return escapeSynDisplayHtml(text).replace(
    /BEV/gi,
    (match) => `<span class="syn-hdr-bev-word">${match}</span>`
  );
}

export function synHdrEnergyValueClass(row, col, displayText) {
  if (row == null || !isSynHeaderPanelRow(row)) return '';
  if (!isSynHdrSummaryTableCol(col)) return '';
  const v = String(displayText != null ? displayText : '').trim().toUpperCase();
  if (!v) return '';
  if (v.includes('MHEVP2')) return 'syn-hdr-val-mhevp2';
  if (v.includes('HEV')) return 'syn-hdr-val-hev';
  if (v === 'AWD') return 'syn-hdr-val-awd';
  return '';
}

export function synHdrEnergyValueStyle(row, col, displayText) {
  const cls = synHdrEnergyValueClass(row, col, displayText);
  if (cls === 'syn-hdr-val-mhevp2') {
    return { background: '#ffff00', backgroundColor: '#ffff00', color: '#000' };
  }
  if (cls === 'syn-hdr-val-hev') {
    return { background: '#ff0000', backgroundColor: '#ff0000', color: '#fff' };
  }
  if (cls === 'syn-hdr-val-awd') {
    return {
      background: '#fff',
      backgroundColor: '#fff',
      color: '#ffc000',
    };
  }
  return null;
}

/** Rows 3–14, display columns C & H (Excel H & M) — #a6a6a6. */
export const SYN_FILTER_GREY_BG = '#a6a6a6';

export function synFilterGreyColClass(row, col) {
  if (row >= 3 && row <= 14 && isSynFilterGreyExcelCol(col)) {
    return 'syn-filter-col-grey';
  }
  return '';
}

export function synFilterGreyColStyle(row, col) {
  if (synFilterGreyColClass(row, col)) {
    return { backgroundColor: SYN_FILTER_GREY_BG, color: '#000' };
  }
  return null;
}

/** Rows 15, 16, 17, 20, 21, 22 — display columns C–J (Excel H–O) white. */
export function synMetricCjWhiteColClass(row, col) {
  if (!SYN_METRIC_CJ_WHITE_ROWS.has(row)) return '';
  if (!isSynHeaderPanelVehicleCol(col)) return '';
  return 'syn-metric-cj-white';
}

export function synMetricCjWhiteColStyle(row, col) {
  if (synMetricCjWhiteColClass(row, col)) {
    return { backgroundColor: '#fff', color: '#000' };
  }
  return null;
}

export function synSpotBlueColClass(row, col) {
  if (isSynSpotBlueCell(row, col)) return 'syn-spot-blue';
  return '';
}

export function synSpotBlueColStyle(row, col) {
  if (synSpotBlueColClass(row, col)) {
    return { backgroundColor: SYN_SPOT_BLUE_BG, color: '#000' };
  }
  return null;
}

/** Row 25+ — C/H grey; D–G & I–J fluo (rows 25–41 + listed extra rows), else grey like C/H. */
export function synAdaptBandColClass(row, col, pillarColumns) {
  if (isSynSpotBlueCell(row, col)) return '';
  if (row < SYN_ZERO_FILL_FIRST_ROW) return '';
  if (col === SYN_LABEL_COL) return '';
  if (isSynForceWhiteExcelCol(col)) return '';
  if (isSynSpacerDisplayExcelCol(col)) return '';
  if (isSynPillarColAtRow(col, row, pillarColumns)) return '';
  if (isSynAdaptGreyExcelCol(col)) return 'syn-adapt-col-grey';
  if (isSynAdaptFluoIjOnlyCell(row, col)) return 'syn-adapt-col-fluo';
  if (isSynAdaptFluoExcelCol(col)) {
    return isSynAdaptFluoBandRow(row) ? 'syn-adapt-col-fluo' : 'syn-adapt-col-grey';
  }
  return '';
}

export function synAdaptBandColStyle(row, col, pillarColumns) {
  const cls = synAdaptBandColClass(row, col, pillarColumns);
  if (cls === 'syn-adapt-col-grey') {
    return { backgroundColor: SYN_PILLAR_BG, color: '#000' };
  }
  if (cls === 'syn-adapt-col-fluo') {
    return { backgroundColor: SYN_ADAPT_FLUO_BG, color: '#000' };
  }
  return null;
}

/** Rows 3–22, columns C–J: P1H / HEV / MHEVP2 / metric rows / default grey. */
export function synHeaderPanelVehicleClass(row, col, displayText) {
  if (synMetricCjWhiteColClass(row, col)) return '';
  if (SYN_HEADER_SPACER_ROWS.has(row)) return '';
  if (!isSynHeaderPanelRow(row) || !isSynHeaderPanelVehicleCol(col)) return '';
  const bevCls = synHdrBevTextClass(row, displayText);
  if (bevCls) return bevCls;
  const accent = synCellAccentClass(displayText);
  if (accent) return accent;
  const energyCls = synHdrEnergyValueClass(row, col, displayText);
  if (energyCls) return energyCls;
  const v = String(displayText != null ? displayText : '')
    .trim()
    .toUpperCase();
  if (v.includes('P1H')) return 'syn-hdr-val-p1h';
  if ((row === 18 || row === 19) && !isSynFilterGreyExcelCol(col)) {
    return 'syn-hdr-row-metric-bg';
  }
  return 'syn-hdr-val-default';
}

/** Vehicle columns (G+): same legend colours as Database B–P. */
export function synProjectCellClass(displayText, col) {
  if (colToNum(col) < colToNum(SYN_VEHICLE_COL_START) || !displayText) return '';
  const accent = synCellAccentClass(displayText);
  if (accent) return accent;
  const v = String(displayText).trim().toUpperCase();
  if (!v) return '';
  if (v === 'TT') return 'cell-proj-tt';
  if (v === 'INFO') return 'cell-proj-info';
  if (v.includes('TARGET') && colToNum(col) >= colToNum(displayToExcelCol('M'))) {
    return 'syn-proj-hdr-green';
  }
  if (v === 'TARGET') return 'cell-proj-target';
  if (v === 'STATUS') return 'cell-proj-status';
  if (v === 'SPC') return 'cell-proj-spc';
  if (['BEV', 'PHEV', 'ICE', 'MHEV', 'PHEV2'].includes(v)) {
    return 'cell-proj-energy';
  }
  if (v === 'FWD' || v === 'AWD' || v === 'RWD') return 'cell-proj-drivetrain';
  if (v === 'PTF') return 'cell-proj-ptf';
  if (v === 'HR' || v === 'XR' || v === 'SR') return 'cell-proj-range';
  return 'cell-proj-night';
}

/** Synthesis grid: all body cells are editable in the UI (see SynthesisGrid cellReadonly). */
export function synIsReadonly(_cell, _row, _sheet) {
  return false;
}

/** Row 25+, display column C through last data column (not label A, not pillars). */
export function isSynZeroFillDataCol(row, col, pillarColumns) {
  if (row < SYN_ZERO_FILL_FIRST_ROW) return false;
  if (col === SYN_LABEL_COL) return false;
  if (isSynForceWhiteExcelCol(col)) return false;
  if (isSynSpacerDisplayExcelCol(col)) return false;
  if (isSynPillarColAtRow(col, row, pillarColumns)) return false;
  return colToNum(col) >= colToNum(SYN_HDR_PANEL_COL_START);
}

/** Editable cells that must receive a numeric value (adapt band + metric C–J). */
export function isSynNumericEntryCell(row, col, pillarColumns) {
  if (isSynZeroFillDataCol(row, col, pillarColumns)) return true;
  if (!isSynHeaderPanelVehicleCol(col) || isSynSpacerDisplayExcelCol(col)) {
    return false;
  }
  if (SYN_METRIC_CJ_WHITE_ROWS.has(row)) return true;
  if (row === 18 || row === 19) return true;
  return false;
}

export function synDisplayValue(cell, map, row, col, sheet, pillarColumns) {
  if (
    isSynPillarColAtRow(col, row, pillarColumns) ||
    (row >= SYN_PILLAR_FIRST_ROW && isSynSp2RestartDisplayExcelCol(col))
  ) {
    const raw = cell ? displayValue(cell) : '';
    if (raw && String(raw).trim()) return String(raw).trim();
    return synPillarLetterForRow(row, col, pillarColumns, map, sheet);
  }
  if (col === SYN_LABEL_COL && isSynHeaderPanelRow(row)) {
    const raw = cell ? displayValue(cell) : '';
    if (raw && String(raw).trim()) {
      return synTranslateText(String(raw).trim(), SYN_LABEL_COL);
    }
    return synHeaderPanelLabel(map, row);
  }
  if (isSynHeaderPanelRow(row) && isSynMaaPresetExcelCol(col)) {
    if (cell && cell.userEdited) {
      const rawEdited = displayValue(cell);
      if (isSynNumericRaw(rawEdited)) {
        return synTranslateText(formatSynHdrMaaMetricDisplay(row, rawEdited, col), col);
      }
      return synTranslateText(rawEdited, col);
    }
    const hdrMaaPreset = synRowMaaPresetRaw(row, col);
    if (hdrMaaPreset !== undefined) {
      if (hdrMaaPreset == null || hdrMaaPreset === '') return '';
      if (typeof hdrMaaPreset === 'number' || isSynNumericRaw(String(hdrMaaPreset))) {
        return synTranslateText(
          formatSynHdrMaaMetricDisplay(row, String(hdrMaaPreset), col),
          col
        );
      }
      return synTranslateText(String(hdrMaaPreset), col);
    }
  }
  if (isSynHeaderPanelRow(row) && isSynAcanPresetExcelCol(col)) {
    if (cell && cell.userEdited) {
      const rawEdited = displayValue(cell);
      if (isSynNumericRaw(rawEdited)) {
        return synTranslateText(formatSynHdrMaaMetricDisplay(row, rawEdited, col), col);
      }
      return synTranslateText(rawEdited, col);
    }
    const hdrAcanPreset = synRowAcanPresetRaw(row, col);
    if (hdrAcanPreset !== undefined) {
      if (hdrAcanPreset == null || hdrAcanPreset === '') return '';
      if (typeof hdrAcanPreset === 'number' || isSynNumericRaw(String(hdrAcanPreset))) {
        return synTranslateText(
          formatSynHdrMaaMetricDisplay(row, String(hdrAcanPreset), col),
          col
        );
      }
      return synTranslateText(String(hdrAcanPreset), col);
    }
  }
  if (isSynHeaderPanelRow(row) && isSynApbbPresetExcelCol(col)) {
    if (cell && cell.userEdited) {
      const rawEdited = displayValue(cell);
      if (isSynNumericRaw(rawEdited)) {
        return synTranslateText(formatSynHdrMaaMetricDisplay(row, rawEdited, col), col);
      }
      return synTranslateText(rawEdited, col);
    }
    const hdrApbbPreset = synRowApbbPresetRaw(row, col);
    if (hdrApbbPreset !== undefined) {
      if (hdrApbbPreset == null || hdrApbbPreset === '') return '';
      if (typeof hdrApbbPreset === 'number' || isSynNumericRaw(String(hdrApbbPreset))) {
        return synTranslateText(
          formatSynHdrMaaMetricDisplay(row, String(hdrApbbPreset), col),
          col
        );
      }
      return synTranslateText(String(hdrApbbPreset), col);
    }
  }
  if (isSynZeroFillDataCol(row, col, pillarColumns)) {
    const rowClass = synRowStyleClass(map, row, sheet);
    const label = synLabel(map, row);
    if (
      isSynAdaptBandExcelCol(col) &&
      row >= SYN_ADAPTATION_SUM_FROM_ROW &&
      !isSynAdaptationSumCell(row, col, sheet)
    ) {
      if (cell && cell.userEdited) {
        const rawEdited = displayValue(cell);
        if (isSynNumericRaw(rawEdited)) {
          return synTranslateText(formatSynNumericDisplay(rawEdited), col);
        }
        return synTranslateText(rawEdited, col);
      }
      const cjPreset = synRowCjPresetRaw(row, col);
      if (cjPreset !== undefined) {
        if (cjPreset == null || cjPreset === '') return '';
        return synTranslateText(formatSynNumericDisplay(String(cjPreset)), col);
      }
      const raw = cell ? displayValue(cell) : '';
      if (raw && String(raw).trim() !== '') {
        if (isSynNumericRaw(raw)) {
          return synTranslateText(formatSynNumericDisplay(raw), col);
        }
        return synTranslateText(String(raw).trim(), col);
      }
      return '';
    }
    const liveMassCell = isSynCalculatedMassCell(
      cell,
      row,
      col,
      sheet,
      label,
      rowClass
    );
    // Body M…AA / AC…AN — only show session-merged values; never Excel export or M…AA presets.
    if (liveMassCell && isSynVehicleMassCol(col)) {
      if (cell && cell.userEdited) {
        const rawEdited = displayValue(cell);
        if (isSynNumericRaw(rawEdited)) {
          return synTranslateText(formatSynNumericDisplay(rawEdited), col);
        }
        return synTranslateText(rawEdited, col);
      }
      const rawLive = cell ? displayValue(cell) : '';
      if (rawLive !== '' && isSynNumericRaw(rawLive)) {
        return synTranslateText(formatSynNumericDisplay(rawLive), col);
      }
      return '';
    }
    if (
      isSynAbDiffCell(row, col, sheet)
    ) {
      const raw = cell ? displayValue(cell) : '';
      if (raw !== '' && isSynNumericRaw(raw)) {
        return synTranslateText(formatSynNumericDisplay(raw), col);
      }
      return synTranslateText(formatSynNumericDisplay('0'), col);
    }
    if (
      isSynSectionSumDataCell(
        row,
        col,
        sheet,
        cell,
        synLabel(map, row),
        synRowStyleClass(map, row, sheet)
      )
    ) {
      return synTranslateText(formatSynNumericDisplay('0'), col);
    }
    if (
      isSynSumproductDataCell(
        row,
        col,
        sheet,
        cell,
        synLabel(map, row),
        synRowStyleClass(map, row, sheet)
      )
    ) {
      const raw = cell ? displayValue(cell) : '';
      if (raw !== '' && isSynNumericRaw(raw)) {
        return synTranslateText(formatSynNumericDisplay(raw), col);
      }
      return synTranslateText(formatSynNumericDisplay('0'), col);
    }
    if (cell && cell.userEdited) {
      const rawEdited = displayValue(cell);
      if (isSynNumericRaw(rawEdited)) {
        return synTranslateText(formatSynNumericDisplay(rawEdited), col);
      }
      return synTranslateText(rawEdited, col);
    }
    const liveSectionSum = isSynSectionSumDataCell(
      row,
      col,
      sheet,
      cell,
      label,
      rowClass
    );
    if (
      isSynMaaPresetExcelCol(col) &&
      !liveSectionSum &&
      !(isSynVehicleMassCol(col) && row >= SYN_CALC_FIRST_ROW)
    ) {
      const maaPreset = synRowMaaPresetRaw(row, col);
      if (maaPreset !== undefined) {
        if (maaPreset == null || maaPreset === '') return '';
        return synTranslateText(formatSynNumericDisplay(String(maaPreset)), col);
      }
      if (isSynMaaGreySpacerExcelRow(row)) return '';
    }
    if (
      row === SYN_ZERO_FILL_FIRST_ROW &&
      !liveSectionSum &&
      !isSynAdaptationSumCell(row, col, sheet)
    ) {
      const raw25 = cell ? displayValue(cell) : '';
      if (raw25 && String(raw25).trim() !== '') {
        if (isSynNumericRaw(raw25)) {
          return synTranslateText(formatSynNumericDisplay(raw25), col);
        }
        return synTranslateText(raw25, col);
      }
      return '0,00';
    }
    if (
      sheet &&
      sheet.adaptationHeaderRow != null &&
      Number(row) === Number(sheet.adaptationHeaderRow) + 1 &&
      isSynAdaptationSumCol(col) &&
      !isSynAdaptationSumCell(row, col, sheet)
    ) {
      const raw = cell ? displayValue(cell) : '';
      if (raw && String(raw).trim() !== '') {
        if (isSynNumericRaw(raw)) {
          return synTranslateText(formatSynNumericDisplay(raw), col);
        }
        return synTranslateText(String(raw).trim(), col);
      }
      return '';
    }
    if (liveSectionSum) {
      return synTranslateText(formatSynNumericDisplay('0'), col);
    }
    const cjPreset = synRowCjPresetRaw(row, col);
    if (cjPreset !== undefined) {
      if (cjPreset == null || cjPreset === '') return '';
      return synTranslateText(formatSynNumericDisplay(String(cjPreset)), col);
    }
    const raw = cell ? displayValue(cell) : '';
    if (raw && String(raw).trim() !== '') {
      if (isSynNumericRaw(raw)) {
        return synTranslateText(formatSynNumericDisplay(raw), col);
      }
      return synTranslateText(raw, col);
    }
    return '0,00';
  }
  if (!cell) return '';
  if (col === SYN_LABEL_COL) {
    const label = synLabel(map, row);
    if (label) {
      const rowCls = synRowStyleClass(map, row, sheet);
      const translated = synTranslateText(label, SYN_LABEL_COL);
      // Yellow L1 sections: NO leading prefix. Blue sub-sections: "_" + UPPERCASE,
      // exactly like the Database page and the Bookmark Matrix (formatBlueBandLabel),
      // so every blue sub-section reads "_NAME" with no space after the underscore.
      if (rowCls === 'syn-row-section') {
        return translated.replace(/^[-\u2013_\s]+/, '');
      }
      if (rowCls === 'syn-row-subsection') {
        return formatBlueBandLabel(translated);
      }
    }
  }
  const raw = displayValue(cell);
  if (isSynMetricRow(row)) {
    const formatted = formatSynMetricValue(row, col, raw);
    return synTranslateText(formatted, col);
  }
  if (isSynNumericRaw(raw)) {
    return synTranslateText(formatSynNumericDisplay(raw), col);
  }
  return synTranslateText(raw, col);
}

/** Yellow section / blue subsection rows (outline / eye view).
 * Separators (e.g. the "PROJECT" band) are intentionally excluded so the eye view
 * shows the SAME section + sub-section list as the Database page and Bookmark Matrix. */
export function isSynOutlineRow(map, row, sheet) {
  const cls = synRowStyleClass(map, row, sheet);
  return cls === 'syn-row-section' || cls === 'syn-row-subsection';
}

/** Last Excel row with label, structure band, or vehicle-column data. */
export function computeSynEffectiveLastRow(sheet, cellMap) {
  const map = cellMap || new Map();
  let last = SYN_HEADER_PANEL_LAST_ROW;
  const rowChecked = new Map();
  const rowHasContent = (r) => {
    if (rowChecked.has(r)) return rowChecked.get(r);
    const ok = synRowHasContent(map, r, sheet);
    rowChecked.set(r, ok);
    return ok;
  };
  for (const cell of (sheet && sheet.cells) || []) {
    if (cell.r <= SYN_HEADER_PANEL_LAST_ROW) continue;
    if (rowHasContent(cell.r)) last = Math.max(last, cell.r);
  }
  for (const rowStr of Object.keys((sheet && sheet.rowBands) || {})) {
    const r = parseInt(rowStr, 10);
    if (r > SYN_HEADER_PANEL_LAST_ROW && rowHasContent(r)) {
      last = Math.max(last, r);
    }
  }
  return Math.min(last, SYN_MAX_EXCEL_ROW);
}

/** Skip blank lines below the grid; keep filter band rows 3–14. */
function isSynRowArchived(sheet, row) {
  const bands = sheet && sheet.archivedRowBands;
  if (!bands || !bands.length || row == null) return false;
  for (const b of bands) {
    if (row >= b.start && row <= b.end) return true;
  }
  return false;
}

export function synRowHasContent(map, row, sheet) {
  // Row 26 is a computed SUM row (27..41) and must always be present in the grid.
  if (Number(row) === synAdaptationSumRow(sheet)) return true;
  if (isSynRowArchived(sheet, row)) return false;
  if (isSynPanelGapRow(row)) return true;
  if (isSynHeaderPanelRow(row)) return true;
  if (synLabel(map, row)) return true;
  const cls = synRowStyleClass(map, row, sheet);
  if (cls !== 'syn-row-data') return true;
  for (const col of (sheet && sheet.columns) || []) {
    if (colToNum(col) < colToNum(SYN_VEHICLE_COL_START)) continue;
    const cell = getCell(map, row, col);
    if (!cell) continue;
    if (cell.f) return true;
    const v = cell.v;
    if (v != null && String(v).trim() !== '') return true;
  }
  return false;
}

export function isSynPanelGapEntry(entry) {
  return Boolean(
    entry &&
      (entry.gapBeforePanel || entry.gapAfterPanel || entry.gapBetween)
  );
}

/** True when the ADAPTATION band (Excel 25–26, display ~26–27) is present. */
export function synGridLooksHealthy(sheet, cellMap = null) {
  if (!sheet) return false;
  const map =
    cellMap instanceof Map
      ? cellMap
      : sheet.cellMap instanceof Map
        ? sheet.cellMap
        : null;
  if (!map) return false;
  const l25 = String(synLabel(map, 25) || '').toUpperCase();
  const l26 = String(synLabel(map, 26) || '');
  if (!l25.includes('ADAPT')) return false;
  if (!l26.startsWith('_')) return false;
  const body = computeSynBodyRows(sheet, map, false);
  const d26 = body.find((e) => e.displayRow === 26);
  return Boolean(d26 && d26.excelRow === 25);
}

export function computeSynBodyRows(sheet, cellMap, outlineOnly = false) {
  const map = cellMap || new Map();
  const lastRow =
    sheet && sheet.effectiveLastRow != null
      ? sheet.effectiveLastRow
      : computeSynEffectiveLastRow(sheet, map);
  const rows = [];
  let displayRow = 1;
  if (!outlineOnly) {
    for (let g = 1; g <= SYN_HDR_PANEL_TOP_GAP_COUNT; g++) {
      rows.push({
        gapBeforePanel: true,
        gapIndex: g,
        excelRow: null,
        displayRow: displayRow++,
      });
    }
  }
  for (let r = SYN_GRID_FIRST_ROW; r <= lastRow; r++) {
    if (SYN_SKIPPED_ROWS.includes(r)) continue;
    if (r > SYN_HEADER_PANEL_LAST_ROW && !synRowHasContent(map, r, sheet)) continue;
    if (outlineOnly && !isSynOutlineRow(map, r, sheet)) continue;
    rows.push({ excelRow: r, displayRow: displayRow++ });
    // Blank display row between Excel rows 18 and 19 (Curb mass update → Control).
    if (!outlineOnly && r === 18) {
      rows.push({
        gapBetween: true,
        gapKey: 'between-18-19',
        excelRow: null,
        displayRow: displayRow++,
      });
    }
    if (!outlineOnly && r === SYN_HEADER_PANEL_LAST_ROW) {
      for (let g = 1; g <= SYN_HDR_PANEL_GAP_COUNT; g++) {
        rows.push({
          gapAfterPanel: true,
          gapIndex: g,
          excelRow: null,
          displayRow: displayRow++,
        });
      }
    }
  }
  return rows;
}

export function synRowHeightPx(sheet, excelRow, defaultPx = 21) {
  const ht =
    (sheet && sheet.rowHeights && sheet.rowHeights[String(excelRow)]) ||
    (sheet && sheet.rowHeights && sheet.rowHeights[excelRow]);
  return ht || defaultPx;
}
