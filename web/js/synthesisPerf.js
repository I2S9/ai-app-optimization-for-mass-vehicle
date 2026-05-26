/** Excel columns A–E hidden in UI; F→A, G→B, … */
export const SYN_HIDDEN_COLS = new Set(['A', 'B', 'C', 'D', 'E']);
export const SYN_STICKY_COL = 'F';
/** Display column A (Excel F) — sub-system labels. */
export const SYN_LABEL_COL_MIN_W = 240;
/** SP1 / SP2 TARGET pillar — minimum width; Excel width wins when wider. */
export const SYN_PILLAR_COL_WIDTH = 72;

export function synPillarColWidth(col, sheet, pillarColumns) {
  if (!pillarColumns?.has(col)) return null;
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

/** Filter band rows 3–14: fixed grey cells in display columns C and H. */
export const SYN_FILTER_GREY_DISPLAY_COLS = ['C', 'H'];

/** ADAPTATION band row 25+: display C & H — same grey as SP1 pillar (B). */
export const SYN_ADAPT_GREY_DISPLAY_COLS = ['C', 'H'];

/** ADAPTATION band rows 25–41: display D–G & I–J — fluorescent yellow (42+ grey). */
export const SYN_ADAPT_FLUO_DISPLAY_COLS = ['D', 'E', 'F', 'G', 'I', 'J'];

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
