/** Prepare BD sheet JSON for the web grid (columns, English labels & values). */
import {
  buildCellMap,
  computeBodyDisplayRows,
  computeOutlineRows,
  computeSectionHeaderRows,
  measureFreeFieldWidth,
  shouldDisplayBodyRow,
  BD_BODY_DISPLAY_ROW_START,
} from './bdStore.js';
import {
  HEADER_FR_EN,
  translateValue,
  translateSubsystemLabel,
} from './bdTranslate.js';
import { colToIndex, indexToCol } from './formulaUtil.js';
import {
  BD_DESIGN_DEPT_COL_RAW,
  BD_MASS_AV_AR_COLS,
  BD_TITLE_COL,
  BD_POSITION_COLS,
  BD_SUBSYSTEM_L1_COL_RAW,
  BD_SUBSYSTEM_L2_COL_RAW,
} from './bdColumnConfig.js';
import {
  buildSynPillarColumns,
  computeSynEffectiveLastRow,
  SYN_MAX_EXCEL_ROW,
  SYN_BUILTIN_PILLAR_META,
  applySynRow25PresetCells,
  applySynRowsCjPresetCells,
  applySynRowsMaaPresetCells,
  applySynRowsMaaPresetHeaderRows,
  applySynRowsAcanPresetCells,
  applySynRowsAcanPresetHeaderRows,
  applySynRowsApbbPresetCells,
  applySynRowsApbbPresetHeaderRows,
} from './synStore.js?v=syn-apbb8';
import { runInChunks, yieldToMain } from './yieldMain.js?v=2';
import { filterSynDisplayColumns } from './synthesisPerf.js';
import { sanitizeSynAdaptationSumCells, sanitizeSynLiveMassCells } from './synthesisCalc.js?v=syn-apbb8';
import { findSynAdaptationRow } from './synStore.js';

const POSITION_INSERT_AFTER = 'AD';
const POSITION_INSERT_COUNT = 2;
const POSITION_HEADER_FR = [
  'Positionnement en Y',
  'Positionnement en Z',
];
const HIDDEN_COLUMNS = new Set([
  'AI',
  'AJ',
  'AK',
  'AL',
  'AM',
  'AN',
  'AQ', // L1 lookup (header "AP" in Excel — not shown in UI)
  'AR', // Row index for INDIRECT → "Line #"
  'AT', // L2 lookup (header "AS" in Excel — not shown in UI)
  'AV',
  'Z', // Source
  'AA', // Pack
  'AB', // Reference
]);
/** Visible subsystem columns after Y/Z insert at AD. */
const SUBSYSTEM_HEADERS_RAW = {
  AP: 'Sub-system L1',
  AS: 'Sub-system L2',
  AU: 'Sub-System Design Dpt',
};
function remapColAfterYZ(col) {
  const afterIdx = colToIndex(POSITION_INSERT_AFTER);
  return remapSheetCol(col, afterIdx, POSITION_INSERT_COUNT);
}
const SUBSYSTEM_RAW_COLS = new Set([
  BD_SUBSYSTEM_L1_COL_RAW,
  BD_SUBSYSTEM_L2_COL_RAW,
  BD_DESIGN_DEPT_COL_RAW,
]);

function isSubsystemDataCol(col) {
  for (const raw of SUBSYSTEM_RAW_COLS) {
    if (col === remapColAfterYZ(raw)) return true;
  }
  return SUBSYSTEM_RAW_COLS.has(col);
}

function translateCellValue(v, col) {
  if (v == null || v === '') return v;
  if (isSubsystemDataCol(col)) return String(v).trim();
  return translateValue(String(v));
}

function stripExcelErrors(v) {
  if (v == null || v === '') return v;
  const t = String(v).trim();
  if (t.startsWith('#') && t.endsWith('!')) return '';
  return v;
}
/** First row where column A is FIN (end of usable data). */
export function findFinRow(sheet) {
  let finRow = null;
  const match = (r, col, v) => {
    if (col !== 'A' || v == null) return;
    if (String(v).trim().toUpperCase() !== 'FIN') return;
    if (finRow == null || r < finRow) finRow = r;
  };
  for (const cell of sheet.cells || []) match(cell.r, cell.c, cell.v);
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const r = parseInt(rowStr, 10);
    if (cols.A) match(r, 'A', cols.A.v);
  }
  return finRow;
}
/** Drop empty tail rows after the FIN marker. */
export function trimSheetAfterFin(sheet) {
  const finRow = findFinRow(sheet);
  if (finRow == null) return sheet;
  const cells = (sheet.cells || []).filter((c) => c.r <= finRow);
  const headerRows = {};
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const r = parseInt(rowStr, 10);
    if (r <= finRow) headerRows[rowStr] = cols;
  }
  return { ...sheet, lastRow: finRow, finRow, cells, headerRows };
}

function shiftFormulaColumns(formula, firstShiftIdx, delta) {
  if (!formula || delta === 0) return formula;
  return String(formula).replace(
    /(\$?)([A-Z]{1,3})(\$?)(\d+)/gi,
    (m, d1, col, d2, row) => {
      const idx = colToIndex(col.toUpperCase());
      if (idx < firstShiftIdx) return m;
      return `${d1}${indexToCol(idx + delta)}${d2}${row}`;
    }
  );
}

function remapSheetCol(col, afterIdx, delta) {
  const idx = colToIndex(col);
  if (idx <= afterIdx) return col;
  return indexToCol(idx + delta);
}

/** Insert Y/Z columns after AD; shift following columns; clear X/Y/Z cells. */
function insertPositionColumns(sheet) {
  const afterIdx = colToIndex(POSITION_INSERT_AFTER);
  const delta = POSITION_INSERT_COUNT;
  const firstShiftIdx = afterIdx + 1;
  const newCols = POSITION_HEADER_FR.map((_, i) => indexToCol(afterIdx + 1 + i));

  const columns = [];
  for (const col of sheet.columns || []) {
    if (colToIndex(col) <= afterIdx) {
      columns.push(col);
      if (col === POSITION_INSERT_AFTER) columns.push(...newCols);
    } else {
      columns.push(remapSheetCol(col, afterIdx, delta));
    }
  }

  const remapKey = (obj) => {
    const out = {};
    for (const [col, val] of Object.entries(obj || {})) {
      const nc = remapSheetCol(col, afterIdx, delta);
      out[nc] = val;
    }
    for (let i = 0; i < newCols.length; i++) {
      out[newCols[i]] = POSITION_HEADER_FR[i];
    }
    return out;
  };

  const headers = remapKey(sheet.headers);
  const headersRaw = remapKey(sheet.headersRaw || sheet.headers);
  for (let i = 0; i < newCols.length; i++) {
    headers[newCols[i]] = HEADER_FR_EN[POSITION_HEADER_FR[i]];
  }

  const colWidths = (sheet.colWidths || []).map((w) => ({
    ...w,
    col: remapSheetCol(w.col, afterIdx, delta),
  }));

  const headerRows = {};
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const row = {};
    for (const [col, cell] of Object.entries(cols)) {
      const nc = remapSheetCol(col, afterIdx, delta);
      const entry = { ...cell };
      if (entry.f) entry.f = shiftFormulaColumns(entry.f, firstShiftIdx, delta);
      row[nc] = entry;
    }
    headerRows[rowStr] = row;
  }

  const cells = (sheet.cells || []).map((cell) => {
    const nc = remapSheetCol(cell.c, afterIdx, delta);
    const entry = { ...cell, c: nc };
    if (entry.f) entry.f = shiftFormulaColumns(entry.f, firstShiftIdx, delta);
    if (BD_POSITION_COLS.has(nc)) {
      delete entry.f;
      delete entry.v;
    } else if (entry.v === '#REF!') {
      delete entry.v;
    }
    return entry;
  });

  return {
    ...sheet,
    columns,
    headers,
    headersRaw,
    colWidths,
    headerRows,
    cells,
  };
}

function clearColumnCells(sheet, cols) {
  const colSet = cols instanceof Set ? cols : new Set(cols);
  const cells = (sheet.cells || []).map((cell) => {
    if (!colSet.has(cell.c)) return cell;
    const entry = { ...cell };
    delete entry.f;
    delete entry.v;
    return entry;
  });
  const headerRows = {};
  for (const [rowStr, colsByRow] of Object.entries(sheet.headerRows || {})) {
    headerRows[rowStr] = { ...colsByRow };
    for (const col of colSet) {
      if (!headerRows[rowStr][col]) continue;
      const entry = { ...headerRows[rowStr][col] };
      delete entry.f;
      delete entry.v;
      headerRows[rowStr][col] = entry;
    }
  }
  return { ...sheet, cells, headerRows };
}

/** Resolve Design Dpt inherited values (anchor rows + formula pointer to parent). */
function fillDesignDeptInherited(sheet, col) {
  const byRow = new Map();
  for (const cell of sheet.cells || []) {
    if (cell.c !== col) continue;
    byRow.set(cell.r, cell);
  }
  const anchors = [];
  for (const [row, cell] of byRow) {
    const literal = cell.v;
    if (
      literal != null &&
      literal !== '' &&
      !String(literal).startsWith('#')
    ) {
      anchors.push({ row, v: String(literal).trim() });
    }
  }
  if (!anchors.length) return;
  anchors.sort((a, b) => a.row - b.row);

  const resolveFormulaParent = (row, visiting = new Set()) => {
    if (visiting.has(row)) return '';
    visiting.add(row);
    const cell = byRow.get(row);
    if (!cell) return '';
    const literal = cell.v;
    if (
      literal != null &&
      literal !== '' &&
      !String(literal).startsWith('#')
    ) {
      return String(literal).trim();
    }
    const f = String(cell.f || '');
    if (!f.includes('#REF!')) {
      const ref = f.match(/(?:AU|AW)(\d+)/i);
      if (ref) return resolveFormulaParent(parseInt(ref[1], 10), visiting);
    }
    return '';
  };

  const start = sheet.dataStartRow || 6;
  let anchorIdx = 0;
  let inherited = '';
  for (let r = start; r <= sheet.lastRow; r++) {
    while (
      anchorIdx + 1 < anchors.length &&
      anchors[anchorIdx + 1].row <= r
    ) {
      anchorIdx += 1;
    }
    if (anchors[anchorIdx].row <= r) inherited = anchors[anchorIdx].v;
    let v = inherited;
    const cell = byRow.get(r);
    if (cell && cell.f && !cell.v) v = resolveFormulaParent(r) || inherited;
    else if (!cell || cell.v == null || cell.v === '') v = inherited;
    else continue;
    if (!v) continue;
    let target = cell;
    if (!target) {
      target = { r, c: col };
      sheet.cells.push(target);
      byRow.set(r, target);
    }
    target.v = v;
  }
}

export function transformBdSheet(sheet) {
  sheet = trimSheetAfterFin(sheet);
  sheet = insertPositionColumns(sheet);
  sheet = clearColumnCells(sheet, BD_MASS_AV_AR_COLS);
  sheet = clearColumnCells(sheet, new Set([BD_TITLE_COL]));
  const afterIdx = colToIndex(POSITION_INSERT_AFTER);
  const hiddenCols = new Set(
    [...HIDDEN_COLUMNS].map((c) => remapSheetCol(c, afterIdx, POSITION_INSERT_COUNT))
  );
  const columns = sheet.columns.filter((c) => !hiddenCols.has(c));
  const headers = { ...sheet.headers };
  for (const col of hiddenCols) {
    delete headers[col];
  }
  for (const [rawCol, label] of Object.entries(SUBSYSTEM_HEADERS_RAW)) {
    const col = remapColAfterYZ(rawCol);
    if (columns.includes(col)) headers[col] = label;
  }
  for (const col of columns) {
    if (Object.values(SUBSYSTEM_HEADERS_RAW).includes(headers[col])) continue;
    const raw = headers[col];
    if (raw && HEADER_FR_EN[raw]) headers[col] = HEADER_FR_EN[raw];
    else if (raw) headers[col] = translateValue(raw);
  }
  const cells = sheet.cells
    .filter((c) => !hiddenCols.has(c.c))
    .map((c) => {
      let entry = { ...c };
      if (isSubsystemDataCol(c.c) && entry.v != null) {
        entry.v = stripExcelErrors(String(entry.v));
        if (entry.v === '') delete entry.v;
        else entry.v = translateSubsystemLabel(String(entry.v));
      } else if (entry.v != null && entry.v !== '') {
        const v = translateCellValue(String(entry.v), c.c);
        if (v !== entry.v) entry = { ...entry, v };
      }
      return entry;
    });
  const headerRows = {};
  for (const [row, cols] of Object.entries(sheet.headerRows || {})) {
    headerRows[row] = {};
    for (const [col, cell] of Object.entries(cols)) {
      if (hiddenCols.has(col)) continue;
      let v = cell.v;
      if (v != null && v !== '') {
        v = isSubsystemDataCol(col)
          ? stripExcelErrors(String(v))
          : translateCellValue(String(v), col);
      }
      headerRows[row][col] = { ...cell, v };
    }
  }
  const colWidths = (sheet.colWidths || []).filter((w) => !hiddenCols.has(w.col));
  const designCol = remapColAfterYZ(BD_DESIGN_DEPT_COL_RAW);
  fillDesignDeptInherited(sheet, designCol);
  const prepared = {
    ...sheet,
    columns,
    headers,
    cells,
    headerRows,
    colWidths,
  };
  return finalizeBdPrepared(prepared);
}

function finalizeBdPrepared(prepared) {
  prepared.cellMap = buildCellMap(prepared.cells, prepared.headerRows);
  const { rows, canonicalByLabel } = computeSectionHeaderRows(
    prepared,
    prepared.cellMap
  );
  prepared.sectionHeaderRows = rows;
  prepared.canonicalSectionByLabel = Object.fromEntries(canonicalByLabel);
  prepared.canonicalSectionMap = canonicalByLabel;
  prepared.outlineRows = computeOutlineRows(prepared, prepared.cellMap).filter(
    (r) => r <= prepared.lastRow
  );
  prepared.bodyDisplayRows = computeBodyDisplayRows(prepared);
  let outlineDisplayRow = BD_BODY_DISPLAY_ROW_START;
  prepared.outlineBodyDisplayRows = (prepared.outlineRows || [])
    .filter((r) => shouldDisplayBodyRow(prepared.cellMap, r, prepared))
    .map((excelRow) => ({ excelRow, displayRow: outlineDisplayRow++ }));
  prepared.freeFieldWidth = measureFreeFieldWidth(prepared.cells);
  return prepared;
}

function mapBdCell(c) {
  let entry = { ...c };
  if (isSubsystemDataCol(c.c) && entry.v != null) {
    entry.v = stripExcelErrors(String(entry.v));
    if (entry.v === '') delete entry.v;
    else entry.v = translateSubsystemLabel(String(entry.v));
  } else if (entry.v != null && entry.v !== '') {
    const v = translateCellValue(String(entry.v), c.c);
    if (v !== entry.v) entry = { ...entry, v };
  }
  return entry;
}

/** Same as transformBdSheet but yields during the heaviest cell pass. */
export async function transformBdSheetAsync(sheet) {
  sheet = trimSheetAfterFin(sheet);
  sheet = insertPositionColumns(sheet);
  sheet = clearColumnCells(sheet, BD_MASS_AV_AR_COLS);
  sheet = clearColumnCells(sheet, new Set([BD_TITLE_COL]));
  const afterIdx = colToIndex(POSITION_INSERT_AFTER);
  const hiddenCols = new Set(
    [...HIDDEN_COLUMNS].map((c) => remapSheetCol(c, afterIdx, POSITION_INSERT_COUNT))
  );
  const columns = sheet.columns.filter((c) => !hiddenCols.has(c));
  const headers = { ...sheet.headers };
  for (const col of hiddenCols) {
    delete headers[col];
  }
  for (const [rawCol, label] of Object.entries(SUBSYSTEM_HEADERS_RAW)) {
    const col = remapColAfterYZ(rawCol);
    if (columns.includes(col)) headers[col] = label;
  }
  for (const col of columns) {
    if (Object.values(SUBSYSTEM_HEADERS_RAW).includes(headers[col])) continue;
    const raw = headers[col];
    if (raw && HEADER_FR_EN[raw]) headers[col] = HEADER_FR_EN[raw];
    else if (raw) headers[col] = translateValue(raw);
  }
  const filtered = sheet.cells.filter((c) => !hiddenCols.has(c.c));
  const cells = await runInChunks(filtered, (slice) => slice.map(mapBdCell));
  await yieldToMain();
  const headerRows = {};
  for (const [row, cols] of Object.entries(sheet.headerRows || {})) {
    headerRows[row] = {};
    for (const [col, cell] of Object.entries(cols)) {
      if (hiddenCols.has(col)) continue;
      let v = cell.v;
      if (v != null && v !== '') {
        v = isSubsystemDataCol(col)
          ? stripExcelErrors(String(v))
          : translateCellValue(String(v), col);
      }
      headerRows[row][col] = { ...cell, v };
    }
  }
  const colWidths = (sheet.colWidths || []).filter((w) => !hiddenCols.has(w.col));
  const designCol = remapColAfterYZ(BD_DESIGN_DEPT_COL_RAW);
  fillDesignDeptInherited(sheet, designCol);
  const prepared = {
    ...sheet,
    columns,
    headers,
    cells,
    headerRows,
    colWidths,
  };
  await yieldToMain();
  return finalizeBdPrepared(prepared);
}

/** Fix legacy exports that stored shared-string indices as plain numbers. */
function resolveSynSheetValues(sheet) {
  const shared = sheet.sharedStrings;
  if (!Array.isArray(shared) || !shared.length) return sheet;
  const resolve = (v) => {
    if (v == null || v === '') return v;
    const s = String(v).trim();
    if (!/^\d+$/.test(s)) return v;
    const hit = shared[parseInt(s, 10)];
    if (
      hit != null &&
      hit !== '' &&
      !/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(String(hit).trim())
    ) {
      return hit;
    }
    return v;
  };
  const cells = (sheet.cells || []).map((c) => {
    if (c.v == null || c.v === '') return c;
    const v = resolve(c.v);
    return v === c.v ? c : { ...c, v: String(v) };
  });
  const headerRows = {};
  for (const [row, cols] of Object.entries(sheet.headerRows || {})) {
    headerRows[row] = {};
    for (const [col, cell] of Object.entries(cols)) {
      const v = cell.v != null && cell.v !== '' ? resolve(cell.v) : cell.v;
      headerRows[row][col] = v === cell.v ? cell : { ...cell, v: String(v) };
    }
  }
  return { ...sheet, cells, headerRows };
}

/** Synthesis grid: columns F… only in UI (A–E hidden), cell map keeps all cols for filters. */
export async function transformSynthesisSheetAsync(sheet) {
  sheet = resolveSynSheetValues(sheet);
  await yieldToMain();
  const columns = filterSynDisplayColumns(sheet.columns || []);
  const headers = { ...(sheet.headers || {}) };
  for (const col of columns) {
    const raw = headers[col];
    if (raw && HEADER_FR_EN[raw]) headers[col] = HEADER_FR_EN[raw];
    else if (raw) headers[col] = translateValue(String(raw));
  }
  const cells = [];
  for (const c of sheet.cells || []) {
    if (c.r <= SYN_MAX_EXCEL_ROW) cells.push(c);
  }
  applySynRow25PresetCells(cells);
  applySynRowsCjPresetCells(cells);
  applySynRowsMaaPresetCells(cells);
  applySynRowsAcanPresetCells(cells);
  applySynRowsApbbPresetCells(cells);
  sanitizeSynLiveMassCells(cells);
  await yieldToMain();
  const headerRows = applySynRowsApbbPresetHeaderRows(
    applySynRowsAcanPresetHeaderRows(
      applySynRowsMaaPresetHeaderRows(
        Object.fromEntries(
          Object.entries(sheet.headerRows || {}).map(([row, cols]) => [
            row,
            Object.fromEntries(
              Object.entries(cols).map(([col, cell]) => [col, { ...cell }])
            ),
          ])
        )
      )
    )
  );
  const cellMap = buildCellMap(cells, headerRows);
  const pillarMap = buildSynPillarColumns(sheet, cellMap);
  for (const [col, meta] of Object.entries(SYN_BUILTIN_PILLAR_META)) {
    pillarMap.set(col, { ...meta, ...pillarMap.get(col) });
  }
  const pillarColumns = Object.fromEntries(pillarMap);
  const effectiveLastRow = Math.min(
    computeSynEffectiveLastRow({ ...sheet, cells }, cellMap),
    SYN_MAX_EXCEL_ROW
  );
  const adaptationHeaderRow = findSynAdaptationRow(cellMap, sheet);
  /** Totals on “-ADAPTATION” row; row below (_ADDBLUE) keeps its own C..J values. */
  const adaptationSumRow = adaptationHeaderRow;
  const adaptationSumFromRow = adaptationHeaderRow + 2;
  const adaptationSumToRow = Math.min(41, effectiveLastRow);
  const adaptationMeta = {
    adaptationHeaderRow,
    adaptationSumRow,
    adaptationSumFromRow,
    adaptationSumToRow,
  };
  sanitizeSynAdaptationSumCells(cells, adaptationMeta);
  const cellMapFinal = buildCellMap(cells, headerRows);
  return {
    ...sheet,
    columns,
    headers,
    cells,
    headerRows,
    cellMap: cellMapFinal,
    pillarColumns,
    effectiveLastRow,
    lastRow: effectiveLastRow,
    adaptationHeaderRow,
    adaptationSumRow,
    adaptationSumFromRow,
    adaptationSumToRow,
    dataStartRow: sheet.dataStartRow || 15,
    filterRows: sheet.filterRows || [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
    rowBands: sheet.rowBands || {},
    sectionHeaderRows: new Set(),
    outlineRows: [],
  };
}

export function transformSynthesisSheet(sheet) {
  sheet = resolveSynSheetValues(sheet);
  const columns = filterSynDisplayColumns(sheet.columns || []);
  const headers = { ...(sheet.headers || {}) };
  for (const col of columns) {
    const raw = headers[col];
    if (raw && HEADER_FR_EN[raw]) headers[col] = HEADER_FR_EN[raw];
    else if (raw) headers[col] = translateValue(String(raw));
  }
  const cells = [];
  for (const c of sheet.cells || []) {
    if (c.r <= SYN_MAX_EXCEL_ROW) cells.push(c);
  }
  applySynRow25PresetCells(cells);
  applySynRowsCjPresetCells(cells);
  applySynRowsMaaPresetCells(cells);
  applySynRowsAcanPresetCells(cells);
  applySynRowsApbbPresetCells(cells);
  sanitizeSynLiveMassCells(cells);
  const headerRows = applySynRowsApbbPresetHeaderRows(
    applySynRowsAcanPresetHeaderRows(
      applySynRowsMaaPresetHeaderRows(
        Object.fromEntries(
          Object.entries(sheet.headerRows || {}).map(([row, cols]) => [
            row,
            Object.fromEntries(
              Object.entries(cols).map(([col, cell]) => [col, { ...cell }])
            ),
          ])
        )
      )
    )
  );
  const cellMap = buildCellMap(cells, headerRows);
  const pillarMap = buildSynPillarColumns(sheet, cellMap);
  for (const [col, meta] of Object.entries(SYN_BUILTIN_PILLAR_META)) {
    pillarMap.set(col, { ...meta, ...pillarMap.get(col) });
  }
  const pillarColumns = Object.fromEntries(pillarMap);
  const effectiveLastRow = Math.min(
    computeSynEffectiveLastRow({ ...sheet, cells }, cellMap),
    SYN_MAX_EXCEL_ROW
  );
  const adaptationHeaderRow = findSynAdaptationRow(cellMap, sheet);
  /** Totals on “-ADAPTATION” row; row below (_ADDBLUE) keeps its own C..J values. */
  const adaptationSumRow = adaptationHeaderRow;
  const adaptationSumFromRow = adaptationHeaderRow + 2;
  const adaptationSumToRow = Math.min(41, effectiveLastRow);
  const adaptationMeta = {
    adaptationHeaderRow,
    adaptationSumRow,
    adaptationSumFromRow,
    adaptationSumToRow,
  };
  sanitizeSynAdaptationSumCells(cells, adaptationMeta);
  const cellMapFinal = buildCellMap(cells, headerRows);
  return {
    ...sheet,
    columns,
    headers,
    cells,
    headerRows,
    cellMap: cellMapFinal,
    pillarColumns,
    effectiveLastRow,
    lastRow: effectiveLastRow,
    adaptationHeaderRow,
    adaptationSumRow,
    adaptationSumFromRow,
    adaptationSumToRow,
    dataStartRow: sheet.dataStartRow || 15,
    filterRows: sheet.filterRows || [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
    rowBands: sheet.rowBands || {},
    sectionHeaderRows: new Set(),
    outlineRows: [],
  };
}
