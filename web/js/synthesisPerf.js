/** Excel columns A–E hidden in UI; F→A, G→B, … */
export const SYN_HIDDEN_COLS = new Set(['A', 'B', 'C', 'D', 'E']);
export const SYN_STICKY_COL = 'F';
/** Display column A (Excel F) — sub-system labels. */
export const SYN_LABEL_COL_MIN_W = 240;
/** SP1 / SP2 TARGET pillar — minimum width; Excel width wins when wider. */
export const SYN_PILLAR_COL_WIDTH = 72;

/** Excel columns that are always pillar bands (display B, K, CG). */
export const SYN_BUILTIN_PILLAR_EXCEL_COLS = new Set(['G', 'P', 'CL']);

export function isSynBuiltinPillarExcelCol(excelCol) {
  return SYN_BUILTIN_PILLAR_EXCEL_COLS.has(excelCol);
}

export function synPillarColWidth(col, sheet, pillarColumns) {
  if (!pillarColumns?.has(col) && !isSynBuiltinPillarExcelCol(col)) return null;
  const fromSheet = sheet?.colWidths?.find((w) => w.col === col)?.width;
  return Math.max(fromSheet || 0, SYN_PILLAR_COL_WIDTH);
}

export function synStickyColWidth(sheet) {
  const fromSheet = sheet?.colWidths?.find((w) => w.col === SYN_STICKY_COL)?.width;
  return Math.max(fromSheet || SYN_LABEL_COL_MIN_W, SYN_LABEL_COL_MIN_W);
}

const HIDDEN_OFFSET = 5;

export function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Excel column letter shown in the grid (F → A, G → B, …). */
export function excelToDisplayCol(excelCol) {
  const n = colToNum(excelCol) - HIDDEN_OFFSET;
  return n > 0 ? numToCol(n) : excelCol;
}

/** Display column → Excel column (C → H, H → M, …). */
export function displayToExcelCol(displayCol) {
  return numToCol(colToNum(displayCol) + HIDDEN_OFFSET);
}

/** SP2 TARGET pillar — display column K (Excel P). */
export const SYN_SP2_PILLAR_DISPLAY_COL = 'K';
/** SP2 RESTART pillar — display column CG (Excel CL). */
export const SYN_SP2_RESTART_PILLAR_DISPLAY_COL = 'CG';
/** White gutter between SP2 and project columns — display L (Excel Q). */
export const SYN_SPACER_DISPLAY_COL = 'L';
/** Rows 3–4 project header band — display columns M…AA (Excel R…AF). */
export const SYN_PROJ_HDR_GREEN_DISPLAY_START = 'M';
export const SYN_PROJ_HDR_GREEN_DISPLAY_END = 'AA';
/** Bold vertical frame for the project table (display M…AA), rows 3–22. */
export const SYN_PROJ_TABLE_EDGE_DISPLAY_COLS = {
  left: SYN_PROJ_HDR_GREEN_DISPLAY_START,
  right: SYN_PROJ_HDR_GREEN_DISPLAY_END,
};
/** Summary table frame — display columns AC…AN, Excel rows 3–22 (display rows 3–23). */
export const SYN_AC_AN_TABLE_DISPLAY_START = 'AC';
export const SYN_AC_AN_TABLE_DISPLAY_END = 'AN';
/** Summary table frame — display columns AP…BB, Excel rows 3–22 (display rows 3–23). */
export const SYN_AP_BB_TABLE_DISPLAY_START = 'AP';
export const SYN_AP_BB_TABLE_DISPLAY_END = 'BB';
/** Summary table frame — display columns BS…CE, Excel rows 3–22 (display rows 3–23). */
export const SYN_BS_CE_TABLE_DISPLAY_START = 'BS';
export const SYN_BS_CE_TABLE_DISPLAY_END = 'CE';
/** Summary table frame — display columns BD…BO, Excel rows 3–22 (display rows 3–23). */
export const SYN_BD_BO_TABLE_DISPLAY_START = 'BD';
export const SYN_BD_BO_TABLE_DISPLAY_END = 'BO';

/** Filter band rows 3–14: fixed grey cells in display columns C and H. */
export const SYN_FILTER_GREY_DISPLAY_COLS = ['C', 'H'];

/** ADAPTATION band row 25+: display C & H — same grey as SP1 pillar (B). */
export const SYN_ADAPT_GREY_DISPLAY_COLS = ['C', 'H'];

/** ADAPTATION band rows 25–41: display D–G & I–J — fluorescent yellow. */
export const SYN_ADAPT_FLUO_DISPLAY_COLS = ['D', 'E', 'F', 'G', 'I', 'J'];
/** Last Excel row with fluo yellow on D–G & I–J (row 42+ → grey like C/H). */
export const SYN_ADAPT_FLUO_LAST_ROW = 41;
/** Extra rows (display C–J band): D–G & I–J fluo outside the 25–41 block. */
export const SYN_ADAPT_FLUO_EXTRA_ROWS = new Set([
  44, 51, 56, 60, 62, 63, 64, 75, 83, 88, 164, 289, 297, 306, 311, 312, 318, 321, 360,
  365, 366, 393, 394, 395, 396, 397, 402,
]);
/** Extra rows: display I & J only — fluorescent yellow (not D–G). */
export const SYN_ADAPT_FLUO_IJ_ONLY_ROWS = new Set([422]);
export const SYN_ADAPT_FLUO_IJ_ONLY_DISPLAY_COLS = ['I', 'J'];

export function isSynFilterGreyExcelCol(excelCol) {
  for (const d of SYN_FILTER_GREY_DISPLAY_COLS) {
    if (excelCol === displayToExcelCol(d)) return true;
  }
  return false;
}

export function isSynAdaptGreyExcelCol(excelCol) {
  for (const d of SYN_ADAPT_GREY_DISPLAY_COLS) {
    if (excelCol === displayToExcelCol(d)) return true;
  }
  return false;
}

export function isSynAdaptFluoExcelCol(excelCol) {
  for (const d of SYN_ADAPT_FLUO_DISPLAY_COLS) {
    if (excelCol === displayToExcelCol(d)) return true;
  }
  return false;
}

/** Spot blue — display F & G on listed Excel rows. */
export const SYN_SPOT_BLUE_FG_DISPLAY_COLS = ['F', 'G'];
export const SYN_SPOT_BLUE_FG_ROWS = new Set([
  53, 71, 72, 73, 89, 91, 98, 164, 204, 212, 278, 288, 343, 345, 360, 364, 383, 384,
  388, 389, 393, 394, 395, 396,
]);
/** Spot blue — display I & J on listed Excel rows. */
export const SYN_SPOT_BLUE_IJ_DISPLAY_COLS = ['I', 'J'];
export const SYN_SPOT_BLUE_IJ_ROWS = new Set([278, 393]);

export function isSynSpotBlueFgExcelCol(excelCol) {
  for (const d of SYN_SPOT_BLUE_FG_DISPLAY_COLS) {
    if (excelCol === displayToExcelCol(d)) return true;
  }
  return false;
}

export function isSynSpotBlueIjExcelCol(excelCol) {
  for (const d of SYN_SPOT_BLUE_IJ_DISPLAY_COLS) {
    if (excelCol === displayToExcelCol(d)) return true;
  }
  return false;
}

/** Excel row + column — fixed blue highlight (overrides ADAPTATION band colours). */
export function isSynSpotBlueCell(row, excelCol) {
  const r = Number(row);
  if (!Number.isFinite(r)) return false;
  if (SYN_SPOT_BLUE_FG_ROWS.has(r) && isSynSpotBlueFgExcelCol(excelCol)) return true;
  if (SYN_SPOT_BLUE_IJ_ROWS.has(r) && isSynSpotBlueIjExcelCol(excelCol)) return true;
  return false;
}

/** Fluorescent yellow on D–G & I–J — rows 25–41 plus SYN_ADAPT_FLUO_EXTRA_ROWS. */
export function isSynAdaptFluoBandRow(row) {
  const r = Number(row);
  if (!Number.isFinite(r)) return false;
  if (SYN_ADAPT_FLUO_EXTRA_ROWS.has(r)) return true;
  return r >= 25 && r <= SYN_ADAPT_FLUO_LAST_ROW;
}

export function isSynAdaptFluoIjOnlyExcelCol(excelCol) {
  for (const d of SYN_ADAPT_FLUO_IJ_ONLY_DISPLAY_COLS) {
    if (excelCol === displayToExcelCol(d)) return true;
  }
  return false;
}

/** Excel row + column — I & J fluo only (outside full D–G / I–J band rows). */
export function isSynAdaptFluoIjOnlyCell(row, excelCol) {
  const r = Number(row);
  if (!Number.isFinite(r)) return false;
  return SYN_ADAPT_FLUO_IJ_ONLY_ROWS.has(r) && isSynAdaptFluoIjOnlyExcelCol(excelCol);
}

export function isSynProjHeaderGreenExcelCol(excelCol) {
  const n = colToNum(excelCol);
  const lo = colToNum(displayToExcelCol(SYN_PROJ_HDR_GREEN_DISPLAY_START));
  const hi = colToNum(displayToExcelCol(SYN_PROJ_HDR_GREEN_DISPLAY_END));
  return n >= lo && n <= hi;
}

export function isSynSpacerDisplayExcelCol(excelCol) {
  return excelCol === displayToExcelCol(SYN_SPACER_DISPLAY_COL);
}

/** Entire display column K (Excel P) — SP2 TARGET band colour. */
export function isSynSp2DisplayExcelCol(excelCol) {
  return excelCol === displayToExcelCol(SYN_SP2_PILLAR_DISPLAY_COL);
}

/** Entire display column CG (Excel CL) — SP2 RESTART band colour. */
export function isSynSp2RestartDisplayExcelCol(excelCol) {
  return excelCol === displayToExcelCol(SYN_SP2_RESTART_PILLAR_DISPLAY_COL);
}

/** CSS class for coloured pillar columns (SP1 grey has no extra class). */
export function synPillarAccentClass(excelCol) {
  if (isSynSp2DisplayExcelCol(excelCol)) return 'syn-pillar-k';
  if (isSynSp2RestartDisplayExcelCol(excelCol)) return 'syn-pillar-cg';
  return '';
}

export function filterSynDisplayColumns(columns = []) {
  return columns.filter((c) => !SYN_HIDDEN_COLS.has(c));
}

export function isSynFilterEdit(row, col) {
  return row >= 3 && row <= 14;
}

/** Header panel vehicle band — display columns C…J (Excel H…O). */
export const SYN_HDR_PANEL_COL_START = 'H';
export const SYN_HDR_PANEL_COL_END = 'O';

export function isSynHeaderPanelVehicleCol(col) {
  const n = colToNum(col);
  return (
    n >= colToNum(SYN_HDR_PANEL_COL_START) && n <= colToNum(SYN_HDR_PANEL_COL_END)
  );
}
