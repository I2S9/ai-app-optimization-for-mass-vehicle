/** SYNTHESIS sheet display helpers (filter band, labels, merges). */
import { displayValue, getCell, isSectionLabel } from './bdStore.js';
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
  SYN_PROJ_HDR_GREEN_DISPLAY_START,
  SYN_PROJ_HDR_GREEN_DISPLAY_END,
  SYN_AC_AN_TABLE_DISPLAY_START,
  SYN_AC_AN_TABLE_DISPLAY_END,
  isSynSpacerDisplayExcelCol,
  isSynSp2DisplayExcelCol,
  isSynSp2RestartDisplayExcelCol,
  isSynBuiltinPillarExcelCol,
  synPillarAccentClass,
} from './synthesisPerf.js';

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
  return pillarColumns?.get(col) ?? SYN_BUILTIN_PILLAR_META[col] ?? null;
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
export const SYN_ROW19_MO_GREEN_BG = '#c6efce';
export const SYN_ROW19_PAA_RED_BG = '#ffc7ce';
export const SYN_ROW16_FLUO_BG = '#ffff00';
export const SYN_ROW25_MAA_GREEN_BG = '#92d050';
/** Display-row based green lines (same as rows 3–4: #92d050). */
export const SYN_DISPLAY_GREEN_BG = SYN_SP2_TARGET_BG;
export const SYN_DISPLAY_GREEN_ROWS = new Set([
  26, 42, 47, 52, 54, 59, 61, 63, 71, 73, 76, 82, 88, 91, 93, 95, 97, 99, 165, 178, 181,
  205, 261, 276, 279, 288, 291, 296, 298, 307, 315, 319, 344, 353, 361, 368,
  372, 280, 289, 394, 398, 403, 407, 411, 415, 417, 422,
]);

/** Display-row based grey blocks for columns M…AA. */
export const SYN_DISPLAY_GREY_MAA_BG = '#a6a6a6';
export const SYN_DISPLAY_GREY_MAA_ROWS = (() => {
  const s = new Set([
    27, 35, 38, 41, 287, 50, 51, 79, 80, 81, 84, 85, 86, 87, 106, 107, 108, 110, 112, 179, 180, 189,
    196, 199, 202, 204, 210, 221, 222, 224, 230, 239, 270, 272, 274, 275, 278,
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
  // User wrote "418 à 412" → interpret as 412…418 inclusive.
  addRange(412, 418);
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

/** Seed row 25 C–J when sheet cells are empty (keeps user edits). */
export function applySynRow25PresetCells(cells = []) {
  const row = SYN_ZERO_FILL_FIRST_ROW;
  for (const [display, value] of SYN_ROW_25_DISPLAY_VALUES) {
    if (value == null) continue;
    const col = displayToExcelCol(display);
    let cell = cells.find((c) => c.r === row && c.c === col);
    if (!cell) {
      cells.push({ r: row, c: col, v: String(value) });
    } else if (
      !cell.userEdited &&
      (!cell.v || String(cell.v).trim() === '')
    ) {
      cell.v = String(value);
    }
  }
  return cells;
}

export const SYN_ROW_26_ZERO_ROW = 26;
const SYN_ROW_26_ZERO_DISPLAY_COLS = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

/** Row 26 — display C…J always 0 (overrides legacy header copies like « Projet »). */
export function applySynRow26ZeroCells(cells = []) {
  const row = SYN_ROW_26_ZERO_ROW;
  for (const display of SYN_ROW_26_ZERO_DISPLAY_COLS) {
    const col = displayToExcelCol(display);
    let cell = cells.find((c) => c.r === row && c.c === col);
    if (!cell) {
      cells.push({ r: row, c: col, v: '0' });
    } else if (!cell.userEdited) {
      cell.v = '0';
      delete cell.f;
    }
  }
  return cells;
}

export function isSynRow26ZeroCol(row, col) {
  if (Number(row) !== SYN_ROW_26_ZERO_ROW) return false;
  return isSynHeaderPanelVehicleCol(col);
}

function parseSynBandNum(v) {
  if (v == null || v === '') return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
}

/** Numeric value for ADAPTATION band rows 26–40 (presets unless userEdited). */
export function getSynAdaptBandNumeric(getCell, row, col) {
  const cell = getCell(row, col);
  if (cell?.userEdited) {
    return parseSynBandNum(displayValue(cell));
  }
  if (isSynRow26ZeroCol(row, col)) {
    return 0;
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
  if (cell?.v != null && cell.v !== '') {
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
  for (const [row, rowMap] of SYN_ROWS_CJ_PRESETS) {
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const cell = cells.find((c) => c.r === row && c.c === col);
      if (cell?.userEdited) continue;
      if (value == null) {
        if (cell) {
          cell.v = '';
          delete cell.f;
        } else {
          cells.push({ r: row, c: col, v: '' });
        }
        continue;
      }
      if (!cell) {
        cells.push({ r: row, c: col, v: String(value) });
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

/** Rows 25–117 (grid display 26–118) — display M…AA presets; null = empty grey spacer. */
export const SYN_ADAPT_MAA_PRESET_FIRST_ROW = 25;
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

/** Rows 25–117 — force display M…AA presets (overrides legacy export). */
export function applySynRowsMaaPresetCells(cells = []) {
  for (const [row, rowMap] of SYN_ROWS_MAA_PRESETS) {
    for (const [display, value] of rowMap) {
      const col = displayToExcelCol(display);
      const cell = cells.find((c) => c.r === row && c.c === col);
      if (cell?.userEdited) continue;
      if (value == null) {
        if (cell) {
          cell.v = '';
          delete cell.f;
        } else {
          cells.push({ r: row, c: col, v: '' });
        }
        continue;
      }
      if (!cell) {
        cells.push({ r: row, c: col, v: String(value) });
      } else {
        cell.v = String(value);
        delete cell.f;
      }
    }
  }
  return cells;
}

/** First -ADAPTATION section row. */
export function findSynAdaptationRow(map, sheet) {
  const last = sheet?.lastRow || SYN_MAX_EXCEL_ROW;
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
  const last = sheet?.lastRow || SYN_MAX_EXCEL_ROW;
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
  const t = String(raw ?? '')
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
  const fallback = SYN_FILTER_ROW_LABELS[row] ?? '';
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
  const s = String(raw ?? '').trim();
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

/** Excel serial date → display (PM pre-target row). */
export function formatSynMetricValue(row, col, raw) {
  const s = String(raw ?? '').trim();
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
  const s = String(display ?? '').trim();
  if (!s) return '';
  if (row === 19) {
    const n = parseFloat(s.replace(',', '.'));
    if (Number.isFinite(n)) {
      if (n < 0) return 'syn-metric-control-ok';
      if (n > 0) return 'syn-metric-control-warn';
    }
  }
  if (row === 20) return 'syn-metric-portfolio';
  if (row === 21) return 'syn-metric-forecast';
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
  for (const m of sheet?.merges || []) {
    if (m.colspan !== 1 || m.endRow - m.startRow < 50) continue;
    const cell = getCell(cellMap, m.startRow, m.startCol);
    const raw = String(cell?.v ?? '').trim();
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
  return pillarColumns?.has(col) ?? false;
}

/** Pillar column — not on virtual panel gap rows (display 21–22). */
export function isSynPillarColAtRow(col, row, pillarColumns) {
  if (row == null) return false;
  return isSynPillarCol(col, pillarColumns) && row >= SYN_PILLAR_FIRST_ROW;
}

export function isSynPillarAnchor(row, col, pillarColumns) {
  const p = pillarColumns?.get(col);
  return Boolean(p && row === p.startRow);
}

export function isSynSp2PillarCol(col, pillarColumns) {
  if (!isSynPillarCol(col, pillarColumns)) return false;
  const title = pillarColumns.get(col)?.title ?? '';
  return /^SP2\b/i.test(title);
}

/** Rows 3–4, display columns M through AA (Excel R…AF). */
export function isSynProjHeaderGreenCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || !SYN_PROJ_HDR_GREEN_ROWS.has(r)) return false;
  return isSynProjHeaderGreenExcelCol(col);
}

export function synProjHeaderGreenStyle() {
  return {
    background: SYN_SP2_TARGET_BG,
    backgroundColor: SYN_SP2_TARGET_BG,
    color: '#000',
  };
}

/** Row 5, display columns M through AA (Excel R…AF). */
export function isSynProjHeaderYellowCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || !SYN_PROJ_HDR_YELLOW_ROWS.has(r)) return false;
  return isSynProjHeaderGreenExcelCol(col);
}

export function synProjHeaderYellowStyle() {
  return {
    background: SYN_PROJ_HDR_YELLOW_BG,
    backgroundColor: SYN_PROJ_HDR_YELLOW_BG,
    color: '#000',
  };
}

/** Row 11, display columns M through AA (Excel R…AF). */
export function isSynProjHeaderGreyCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || !SYN_PROJ_HDR_GREY_ROWS.has(r)) return false;
  return isSynProjHeaderGreenExcelCol(col);
}

export function synProjHeaderGreyStyle() {
  return {
    background: SYN_PROJ_HDR_GREY_BG,
    backgroundColor: SYN_PROJ_HDR_GREY_BG,
    color: '#000',
  };
}

/** Row 19: display M–O green, P–AA red. */
export function isSynRow19MoGreenCol(row, col) {
  if (Number(row) !== 19) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('M'));
  const end = colToNum(displayToExcelCol('O'));
  return n >= start && n <= end;
}

export function isSynRow19PaaRedCol(row, col) {
  if (Number(row) !== 19) return false;
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
 * Row 16: every 3 columns fluo starting at display M.
 * Pattern: M, P, S, V, Y (within M…AA).
 */
export function isSynRow16FluoEvery3FromMCol(row, col) {
  if (Number(row) !== 16) return false;
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('M'));
  const end = colToNum(displayToExcelCol('AA'));
  if (n < start || n > end) return false;
  return (n - start) % 3 === 0;
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
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol('M'));
  const end = colToNum(displayToExcelCol('AA'));
  return n >= start && n <= end;
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

export function synDisplayRowGreenMaaStyle() {
  return {
    background: SYN_DISPLAY_GREEN_BG,
    backgroundColor: SYN_DISPLAY_GREEN_BG,
    color: '#000',
  };
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

/** M / AA frame — Excel rows 3–22 only (not top gap rows). */
export function isSynHdrLmDividerLeftEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrLmDividerLeftCol(entry.excelRow, col);
}

export function isSynHdrAaDividerRightEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrAaDividerRightCol(entry.excelRow, col);
}

export function isSynHdrCjDividerRightEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrCjDividerRightCol(entry.excelRow, col);
}

export function isSynHdrMaDividerRightEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrMaDividerRightCol(entry.excelRow, col);
}

export function isSynHdrLmDividerRightEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrLmDividerRightCol(entry.excelRow, col);
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
  const n = colToNum(col);
  const start = colToNum(displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_START));
  const end = colToNum(displayToExcelCol(SYN_AC_AN_TABLE_DISPLAY_END));
  return n >= start && n <= end;
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
  if (!entry || entry.excelRow == null) return false;
  return isSynAcAnTableCell(entry.excelRow, col);
}

export function isSynHdrAcAnDividerRightEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrAcAnDividerRightCol(entry.excelRow, col);
}

export function isSynHdrAcAnDividerLeftEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrAcAnDividerLeftCol(entry.excelRow, col);
}

export function isSynHdrAcAnDividerRightEdgeEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrAcAnDividerRightEdgeCol(entry.excelRow, col);
}

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
  const rows = sheet?.filterRows || [...SYN_FILTER_ROWS];
  return rows.includes(row);
}

export function isSynSeparatorRow(row) {
  return row === 4;
}

/** Yellow L1 band (AILES, ALTERNATEUR, -ADAPTATION, BOUCLIER AR…). */
export function isSynL1SectionLabel(label) {
  const t = String(label ?? '').trim();
  if (!t) return false;
  if (t.startsWith('-')) return true;
  return isSectionLabel(t);
}

export function isSynL2SubsectionLabel(label) {
  const t = String(label ?? '').trim();
  return t.startsWith('_');
}

/** Rows under an L1 title until the next L1 (e.g. between AILES and ALTERNATEUR). */
export function isSynBetweenL1SectionRows(map, row, sheet) {
  if (isSynL1SectionLabel(synLabel(map, row))) return false;
  const last = sheet?.effectiveLastRow ?? sheet?.lastRow ?? SYN_MAX_EXCEL_ROW;
  let lastL1 = 0;
  for (let r = SYN_GRID_FIRST_ROW; r < row; r++) {
    if (isSynL1SectionLabel(synLabel(map, r))) lastL1 = r;
  }
  if (!lastL1) return false;
  for (let r = row + 1; r <= last; r++) {
    if (isSynL1SectionLabel(synLabel(map, r))) return true;
  }
  return false;
}

export function isSynSectionLabelRow(map, row, sheet) {
  const band = sheet?.rowBands?.[String(row)] ?? sheet?.rowBands?.[row];
  if (band === 'section' || band === 'subsection') return true;
  const label = synLabel(map, row);
  if (!label) return false;
  return label.startsWith('-') || label.startsWith('_');
}

/** Row band from Excel label column (F) export, else label-prefix fallback. */
export function synRowStyleClass(map, row, sheet) {
  if (isSynPanelGapRow(row)) return 'syn-panel-gap';
  const band = sheet?.rowBands?.[String(row)] ?? sheet?.rowBands?.[row];
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
  if (isSynProjHeaderGreenCol(row, col)) {
    style.background = SYN_SP2_TARGET_BG;
    style.backgroundColor = SYN_SP2_TARGET_BG;
    style.color = '#000';
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
      sheet?.matrixColors?.[row] ?? sheet?.matrixColors?.[String(row)];
    if (
      rowColor &&
      (rowCls === 'syn-row-section' || rowCls === 'syn-row-subsection')
    ) {
      style.backgroundColor = rowColor;
    }
    return style;
  }
  if (cell?.b) style.fontWeight = '700';
  const disp = synDisplayValue(cell, map, row, col, sheet, pillarColumns);
  Object.assign(style, synAdaptCjBoldFontStyle(row, col, disp, sheet) || {});
  return style;
}

function isSynHeaderPanelVehicleCol(col) {
  const n = colToNum(col);
  return (
    n >= colToNum(SYN_HDR_PANEL_COL_START) && n <= colToNum(SYN_HDR_PANEL_COL_END)
  );
}

/** Rows 3–19, display columns C–J (Excel H–O). */
export function isSynHeaderPanelBoldCol(row, col) {
  const r = Number(row);
  if (!Number.isFinite(r) || r < SYN_GRID_FIRST_ROW || r > SYN_HDR_PANEL_BOLD_LAST_ROW) {
    return false;
  }
  if (col === SYN_LABEL_COL) return false;
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
  const last = sheet?.effectiveLastRow ?? sheet?.lastRow ?? SYN_MAX_EXCEL_ROW;
  return r <= last;
}

export function isSynAdaptCjValueBold(row, col, displayText, sheet) {
  if (!isSynAdaptCjBoldRow(row, sheet)) return false;
  if (!isSynHeaderPanelVehicleCol(col)) return false;
  if (isSynSpacerDisplayExcelCol(col)) return false;
  const t = String(displayText ?? '').trim();
  if (!t || t === '0,00') return false;
  return true;
}

export function synAdaptCjBoldFontStyle(row, col, displayText, sheet) {
  if (!isSynAdaptCjValueBold(row, col, displayText, sheet)) return null;
  return { fontWeight: '700' };
}

/** Avenger like (#d8e4bc) / P1X (#c0504d) — exact cell value (row 5 Silhouette). */
export function synCellAccentClass(displayText) {
  const u = String(displayText ?? '')
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
  const accent = synCellAccentClass(displayText);
  if (accent) return accent;
  const v = String(displayText ?? '')
    .trim()
    .toUpperCase();
  if (v.includes('MHEVP2')) return 'syn-hdr-val-mhevp2';
  if (v.includes('P1H')) return 'syn-hdr-val-p1h';
  if (v.includes('HEV')) return 'syn-hdr-val-hev';
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
  if (v === 'TARGET') return 'cell-proj-target';
  if (v === 'STATUS') return 'cell-proj-status';
  if (v === 'SPC') return 'cell-proj-spc';
  if (['BEV', 'HEV', 'MHEVP2', 'PHEV', 'ICE', 'MHEV', 'PHEV2'].includes(v)) {
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
  if (isSynZeroFillDataCol(row, col, pillarColumns)) {
    if (cell?.userEdited) {
      const rawEdited = displayValue(cell);
      if (isSynNumericRaw(rawEdited)) {
        return synTranslateText(formatSynNumericDisplay(rawEdited), col);
      }
      return synTranslateText(rawEdited, col);
    }
    if (isSynRow26ZeroCol(row, col)) {
      return '0,00';
    }
    if (isSynMaaPresetExcelCol(col)) {
      const maaPreset = synRowMaaPresetRaw(row, col);
      if (maaPreset !== undefined) {
        if (maaPreset == null || maaPreset === '') return '';
        return synTranslateText(formatSynNumericDisplay(String(maaPreset)), col);
      }
      if (isSynMaaGreySpacerExcelRow(row)) return '';
    }
    if (row === SYN_ZERO_FILL_FIRST_ROW) {
      const raw25 = cell ? displayValue(cell) : '';
      if (raw25 && String(raw25).trim() !== '') {
        if (isSynNumericRaw(raw25)) {
          return synTranslateText(formatSynNumericDisplay(raw25), col);
        }
        return synTranslateText(raw25, col);
      }
      return '0,00';
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
    if (isSynSectionLabelRow(map, row, sheet) && label) {
      return synTranslateText(label, SYN_LABEL_COL);
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

/** Yellow section / blue subsection / separator rows (outline / eye view). */
export function isSynOutlineRow(map, row, sheet) {
  const cls = synRowStyleClass(map, row, sheet);
  return (
    cls === 'syn-row-section' ||
    cls === 'syn-row-subsection' ||
    cls === 'syn-row-separator'
  );
}

/** Last Excel row with label, structure band, or vehicle-column data. */
export function computeSynEffectiveLastRow(sheet, cellMap) {
  const map = cellMap || new Map();
  let last = SYN_HEADER_PANEL_LAST_ROW;
  for (const cell of sheet?.cells || []) {
    if (cell.r <= SYN_HEADER_PANEL_LAST_ROW) continue;
    if (synRowHasContent(map, cell.r, sheet)) last = Math.max(last, cell.r);
  }
  for (const rowStr of Object.keys(sheet?.rowBands || {})) {
    const r = parseInt(rowStr, 10);
    if (r > SYN_HEADER_PANEL_LAST_ROW && synRowHasContent(map, r, sheet)) {
      last = Math.max(last, r);
    }
  }
  return Math.min(last, SYN_MAX_EXCEL_ROW);
}

/** Skip blank lines below the grid; keep filter band rows 3–14. */
export function synRowHasContent(map, row, sheet) {
  if (isSynPanelGapRow(row)) return true;
  if (isSynHeaderPanelRow(row)) return true;
  if (synLabel(map, row)) return true;
  const cls = synRowStyleClass(map, row, sheet);
  if (cls !== 'syn-row-data') return true;
  for (const col of sheet?.columns || []) {
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
  return Boolean(entry?.gapBeforePanel || entry?.gapAfterPanel || entry?.gapBetween);
}

export function computeSynBodyRows(sheet, cellMap, outlineOnly = false) {
  const map = cellMap || new Map();
  const lastRow =
    sheet?.effectiveLastRow ??
    computeSynEffectiveLastRow(sheet, map);
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
    // Insert an extra blank display row between Excel rows 18 and 19.
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
  const ht = sheet?.rowHeights?.[String(excelRow)] ?? sheet?.rowHeights?.[excelRow];
  return ht || defaultPx;
}
