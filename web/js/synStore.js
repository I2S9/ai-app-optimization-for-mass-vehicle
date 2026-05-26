/** SYNTHESIS sheet display helpers (filter band, labels, merges). */
import { displayValue, getCell, isSectionLabel } from './bdStore.js';
import { translateValue, translateSubsystemLabel } from './bdTranslate.js';
import {
  displayToExcelCol,
  isSynFilterGreyExcelCol,
  isSynAdaptGreyExcelCol,
  isSynAdaptFluoExcelCol,
  isSynProjHeaderGreenExcelCol,
  SYN_PROJ_HDR_GREEN_DISPLAY_START,
  SYN_PROJ_HDR_GREEN_DISPLAY_END,
  isSynSpacerDisplayExcelCol,
  isSynSp2DisplayExcelCol,
} from './synthesisPerf.js';

export {
  isSynSpacerDisplayExcelCol,
  isSynSp2DisplayExcelCol,
} from './synthesisPerf.js';

export const SYN_FILTER_ROWS = new Set([3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]);
export const SYN_LABEL_COL = 'F';
export const SYN_VEHICLE_COL_START = 'G';
/** Header panel colour band — display C…J (Excel H…O), rows 3–22. */
export const SYN_HDR_PANEL_COL_START = 'H';
export const SYN_HDR_PANEL_COL_END = 'O';
/** Rows 18–19, display columns C–J (Excel H–O). */
export const SYN_HDR_METRIC_ROW_BG = '#ebf1de';
export const SYN_HDR_METRIC_BG_ROWS = new Set([18, 19]);
export const SYN_HDR_PANEL_GAP_COUNT = 2;
/** Blank rows before Date (display 1–2); pillars B/K keep pillar fill. */
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
/** SP1 / SP2 pillars — vertical title rendered in a grid overlay (Excel G & P). */
export const SYN_PILLAR_OVERLAY_COLS = new Set(['G', 'P']);
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
/** Last Excel row with fluo yellow in display D–G & I–J (row 42+ = grey like C & H). */
export const SYN_ADAPT_FLUO_LAST_ROW = 41;
/** Display column K (Excel P) + rows 3–4 project header (display M…AA). */
export const SYN_SP2_TARGET_BG = '#92d050';
export const SYN_COL_K_BG = SYN_SP2_TARGET_BG;
export const SYN_PROJ_HDR_GREEN_ROWS = new Set([3, 4]);
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

/** "SP1 Target" → "SP1 TARGET" (space between SPn and TARGET). */
export function normalizeSynPillarTitle(raw) {
  const t = String(raw ?? '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
  const m = t.match(/^(SP\d+)\s*TARGET$/i);
  if (m) return `${m[1]} TARGET`;
  return t;
}

/** One character per row; keeps the space between SP1 and TARGET. */
export function synPillarLettersFromTitle(title) {
  return [...normalizeSynPillarTitle(title)];
}

/** One letter every SYN_PILLAR_LETTER_ROW_STEP rows from Échappement downward. */
export function synPillarLetterForRow(row, col, pillarColumns, map, sheet) {
  const p = pillarColumns?.get(col);
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
  for (const m of sheet?.merges || []) {
    if (m.colspan !== 1 || m.endRow - m.startRow < 50) continue;
    const cell = getCell(cellMap, m.startRow, m.startCol);
    const raw = String(cell?.v ?? '').trim();
    if (!raw || !/target/i.test(raw)) continue;
    pillars.set(m.startCol, {
      title: normalizeSynPillarTitle(raw),
      startRow: m.startRow,
      endRow: m.endRow,
    });
  }
  return pillars;
}

export function isSynPillarCol(col, pillarColumns) {
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

export function isSynHdrLmDividerRightEntry(entry, col) {
  if (!entry || entry.excelRow == null) return false;
  return isSynHdrLmDividerRightCol(entry.excelRow, col);
}

/** @deprecated use isSynHdrLmDividerRightCol */
export function isSynHdrLmDividerCol(row, col) {
  return isSynHdrLmDividerRightCol(row, col);
}

/** Display column L (Excel Q) — white gutter, all body rows. */
export function synSpacerColClass(col) {
  return isSynSpacerDisplayExcelCol(col) ? 'syn-spacer-col-l' : '';
}

const SYN_SPACER_COL_BORDER = '1px solid #fff';

export function synSpacerColStyle(col) {
  if (synSpacerColClass(col)) {
    return {
      background: '#fff',
      backgroundColor: '#fff',
      color: '#000',
      border: SYN_SPACER_COL_BORDER,
      borderTop: SYN_SPACER_COL_BORDER,
      borderBottom: SYN_SPACER_COL_BORDER,
      borderLeft: SYN_SPACER_COL_BORDER,
      borderRight: SYN_SPACER_COL_BORDER,
    };
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
  const spacerStyle = synSpacerColStyle(col);
  if (spacerStyle) {
    Object.assign(style, spacerStyle);
    return style;
  }
  if (isSynPillarColAtRow(col, row, pillarColumns)) {
    const bg = isSynSp2DisplayExcelCol(col) ? SYN_SP2_TARGET_BG : SYN_PILLAR_BG;
    style.background = bg;
    style.backgroundColor = bg;
    style.color = '#000';
    style.border = 'none';
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
    if (isSynHeaderPanelBoldCol(row, col)) style.fontWeight = '700';
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
  const greyStyle = synFilterGreyColStyle(row, col);
  if (greyStyle) {
    Object.assign(style, greyStyle);
    if (isSynHeaderPanelBoldCol(row, col)) style.fontWeight = '700';
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
    return style;
  }
  if (isSynHeaderPanelBoldCol(row, col)) {
    style.fontWeight = '700';
    style.fontSize = '12px';
  }
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
  return style;
}

/** Display columns C–J (Excel H–O). */
export function isSynHeaderPanelVehicleCol(col) {
  const n = colToNum(col);
  return (
    n >= colToNum(SYN_HDR_PANEL_COL_START) && n <= colToNum(SYN_HDR_PANEL_COL_END)
  );
}

/** Rows 3–22, display columns C–J only (not L / project columns). */
export function isSynHeaderPanelBoldCol(row, col) {
  if (!isSynHeaderPanelRow(row)) return false;
  if (col === SYN_LABEL_COL) return false;
  return isSynHeaderPanelVehicleCol(col);
}

/** Avenger like (light green) / P1X (#c0504d) — exact cell value. */
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
    return { backgroundColor: '#c0504d', color: '#fff' };
  }
  if (cls === 'syn-val-avenger-like') {
    return { backgroundColor: '#c6efce', color: '#000' };
  }
  return null;
}

/** Rows 3–14, display columns C & H (Excel H & M) — #a6a6a6. */
export function synFilterGreyColClass(row, col) {
  if (row >= 3 && row <= 14 && isSynFilterGreyExcelCol(col)) {
    return 'syn-filter-col-grey';
  }
  return '';
}

export function synFilterGreyColStyle(row, col) {
  if (synFilterGreyColClass(row, col)) {
    return { backgroundColor: '#a6a6a6', color: '#000' };
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

/** Row 25+ — display C/H grey; rows 25–41 D–G & I–J fluo; row 42+ D–G & I–J grey. */
export function synAdaptBandColClass(row, col, pillarColumns) {
  if (row < SYN_ZERO_FILL_FIRST_ROW) return '';
  if (col === SYN_LABEL_COL) return '';
  if (isSynPillarColAtRow(col, row, pillarColumns)) return '';
  if (isSynAdaptGreyExcelCol(col)) return 'syn-adapt-col-grey';
  if (isSynAdaptFluoExcelCol(col)) {
    return row <= SYN_ADAPT_FLUO_LAST_ROW ? 'syn-adapt-col-fluo' : 'syn-adapt-col-grey';
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
  if (row === 18 || row === 19) return 'syn-hdr-row-metric-bg';
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
  if (isSynPillarColAtRow(col, row, pillarColumns)) return false;
  return colToNum(col) >= colToNum(SYN_HDR_PANEL_COL_START);
}

export function synDisplayValue(cell, map, row, col, sheet, pillarColumns) {
  if (isSynPillarColAtRow(col, row, pillarColumns)) {
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
  return Boolean(entry?.gapBeforePanel || entry?.gapAfterPanel);
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
