/**
 * Server-side Synthesis mass materialization (SUMPRODUCT, section sums, AB diff).
 * Used at build time and by web/server.mjs PATCH /cells (P2).
 */
import {
  buildCellMap,
  getCell,
  displayCellValue,
  displayValue,
  bdSubsystemL1Col,
  bdSubsystemL2Col,
} from './bdStore.js';
import {
  buildBdColumnIndex,
  buildSumproductL2Index,
  buildVehColFilterIndex,
  buildSynSectionRowList,
  computeSumproduct,
  computeSynSectionSum,
  computeSynAbDiff,
  computeAdaptationRowSum,
  isSynAdaptationSumCell,
  isSynAbDiffCell,
  isSynBlueSubsectionRow,
  isSynSectionSumDataCell,
  isSynSumproductDataCell,
  isSynCalculatedMassCell,
  isSynVehicleMassCol,
  isSynAbDiffExcelCol,
  synLiveMassExcelCols,
  SYN_CALC_FIRST_ROW,
  synVehicleMassExcelCol,
} from './synthesisCalc.js';
import { getSynAdaptBandNumeric } from './synStore.js';
import { synLabel, synRowStyleClass } from './synStore.js';
import { formatSynNumericDisplay } from './synStore.js';
import { BD_SUBSYSTEM_L2_COL } from './bdColumnConfig.js';

function parseSynNum(raw) {
  if (raw == null || raw === '') return 0;
  const n = parseFloat(String(raw).replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
}

function formatCalc(n) {
  return formatSynNumericDisplay(String(Math.round(n * 10000) / 10000));
}

/**
 * @param {object} bdSheet transformed BD grid sheet
 * @param {object} synSheet transformed Synthesis grid sheet
 */
export function createSynCalcContext(bdSheet, synSheet) {
  const bdCols = buildBdColumnIndex(bdSheet);
  const bdMap =
    bdSheet.cellMap instanceof Map
      ? bdSheet.cellMap
      : buildCellMap(bdSheet.cells, bdSheet.headerRows);
  const synMap =
    synSheet.cellMap instanceof Map
      ? synSheet.cellMap
      : buildCellMap(synSheet.cells, synSheet.headerRows);
  const bdL1 = bdSubsystemL1Col(bdSheet);
  const bdL2 = bdSubsystemL2Col(bdSheet);

  const sumproductCache = new Map();
  const vehColFilterIndex = new Map();

  function getBdValue(r, c) {
    const raw = bdCols[c] ? bdCols[c][r] : undefined;
    if (raw != null && raw !== '') return String(raw);
    const displayed = displayCellValue(
      bdMap,
      r,
      c,
      bdSheet.sectionHeaderRows,
      bdSheet.canonicalSectionByLabel,
      bdL1,
      bdL2
    );
    if (displayed !== '') return displayed;
    return '';
  }

  function getSynCell(row, col) {
    if (col === 'F') {
      const label = synLabel(synMap, row);
      if (label != null && String(label).trim() !== '') return String(label).trim();
    }
    const cell = getCell(synMap, row, col);
    return cell && cell.v != null && cell.v !== '' ? String(cell.v) : '';
  }

  function getSynLabel(row) {
    return synLabel(synMap, row);
  }

  function getSynRowClass(row) {
    return synRowStyleClass(synMap, row, synSheet);
  }

  function isBlueRow(row) {
    return isSynBlueSubsectionRow(row, synSheet, getSynLabel(row), getSynRowClass(row));
  }

  const sumproductL2Index = buildSumproductL2Index(bdCols, getBdValue);

  function getVehColFilterRows(excelCol) {
    if (!vehColFilterIndex.has(excelCol)) {
      vehColFilterIndex.set(
        excelCol,
        buildVehColFilterIndex(bdCols, excelCol, getSynCell, getBdValue)
      );
    }
    return vehColFilterIndex.get(excelCol);
  }

  function getBlueMaaValue(row, col) {
    const excelCol = synVehicleMassExcelCol(col);
    const cacheKey = `b:${row}:${excelCol}`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const n = computeSumproduct(
      bdCols,
      row,
      excelCol,
      getSynCell,
      getBdValue,
      sumproductL2Index,
      getVehColFilterRows(excelCol)
    );
    const rounded = Math.round(n * 10000) / 10000;
    sumproductCache.set(cacheKey, String(rounded));
    return rounded;
  }

  let synSectionRows = buildSynSectionRowList(synSheet, getSynLabel, getSynRowClass);

  function rebuildSectionRows() {
    synSectionRows = buildSynSectionRowList(synSheet, getSynLabel, getSynRowClass);
  }

  function getSectionMaaValue(sectionRow, col) {
    const excelCol = synVehicleMassExcelCol(col);
    const cacheKey = `g:${sectionRow}:${excelCol}`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const n = computeSynSectionSum(
      sectionRow,
      excelCol,
      synSectionRows,
      isBlueRow,
      (r, c) => getBlueMaaValue(r, c),
      synSheet.effectiveLastRow != null
        ? synSheet.effectiveLastRow
        : synSheet.lastRow != null
          ? synSheet.lastRow
          : 422
    );
    sumproductCache.set(cacheKey, String(n));
    return n;
  }

  function getAdaptationBlockNumeric(row, col) {
    if (isSynAdaptationSumCell(row, col)) return 0;
    if (isSynVehicleMassCol(col) || isSynAbDiffExcelCol(col)) {
      return resolveNumericAtInner(row, col);
    }
    return getSynAdaptBandNumeric((r, c) => getCell(synMap, r, c), row, col);
  }

  function resolveNumericAt(row, col) {
    if (isSynAdaptationSumCell(row, col)) {
      return computeAdaptationRowSum(getAdaptationBlockNumeric, col);
    }
    return resolveNumericAtInner(row, col);
  }

  function resolveNumericAtInner(row, col) {
    const cell = getCell(synMap, row, col);
    if (cell && cell.userEdited) return parseSynNum(displayValue(cell));
    const label = getSynLabel(row);
    const rowClass = getSynRowClass(row);
    if (isSynSectionSumDataCell(row, col, synSheet, cell, label, rowClass)) {
      return getSectionMaaValue(row, col);
    }
    if (isSynSumproductDataCell(row, col, synSheet, cell, label, rowClass)) {
      return getBlueMaaValue(row, col);
    }
    return parseSynNum(displayValue(cell));
  }

  function getAbDiffValue(row) {
    const cacheKey = `d:${row}:AB`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const n = computeSynAbDiff(resolveNumericAt, row);
    sumproductCache.set(cacheKey, String(n));
    return n;
  }

  function getDisplayValue(row, col) {
    const cell = getCell(synMap, row, col);
    if (cell && cell.userEdited && !isSynAdaptationSumCell(row, col)) {
      return displayValue(cell);
    }
    if (isSynAdaptationSumCell(row, col)) {
      return formatCalc(computeAdaptationRowSum(getAdaptationBlockNumeric, col));
    }
    const label = getSynLabel(row);
    const rowClass = getSynRowClass(row);
    if (isSynAbDiffCell(row, col, synSheet)) {
      return formatCalc(getAbDiffValue(row));
    }
    if (isSynSectionSumDataCell(row, col, synSheet, cell, label, rowClass)) {
      return formatCalc(getSectionMaaValue(row, col));
    }
    if (isSynSumproductDataCell(row, col, synSheet, cell, label, rowClass)) {
      return formatCalc(getBlueMaaValue(row, col));
    }
    return displayValue(cell);
  }

  function invalidateCaches(vehCol = null) {
    for (const k of sumproductCache.keys()) {
      if (k.startsWith('b:') || k.startsWith('g:') || k.startsWith('d:')) {
        sumproductCache.delete(k);
      }
    }
    if (vehCol) {
      vehColFilterIndex.delete(synVehicleMassExcelCol(vehCol));
    } else {
      vehColFilterIndex.clear();
    }
    rebuildSectionRows();
  }

  function patchBdCell(row, col, value) {
    const key = `${row}:${col}`;
    let cell = bdMap.get(key);
    if (!cell) {
      cell = { r: row, c: col, v: value, userEdited: true };
      bdSheet.cells.push(cell);
      bdMap.set(key, cell);
    } else {
      cell.v = value;
      cell.userEdited = true;
      delete cell.f;
    }
    if (!bdCols[col]) bdCols[col] = [];
    bdCols[col][row] = value;
    // A filter / gate / L2 / mass change can affect every vehicle column's filter
    // set, so fully invalidate. (recalcAfterBdEdit also does this, but keeping
    // patchBdCell self-consistent means a standalone patch is never stale.)
    invalidateCaches(null);
  }

  /** @returns {{ r: number, c: string, v: string, mat: true }[]} */
  function materializeAll() {
    const patches = [];
    const last =
      synSheet.effectiveLastRow != null
        ? synSheet.effectiveLastRow
        : synSheet.lastRow != null
          ? synSheet.lastRow
          : 422;
    const cols = synLiveMassExcelCols();
    for (const excelCol of cols) {
      getVehColFilterRows(excelCol);
    }
    for (let row = SYN_CALC_FIRST_ROW; row <= last; row++) {
      for (const col of cols) {
        const cell = getCell(synMap, row, col);
        const label = getSynLabel(row);
        const rowClass = getSynRowClass(row);
        if (
          !isSynCalculatedMassCell(cell, row, col, synSheet, label, rowClass) &&
          !isSynAdaptationSumCell(row, col)
        ) {
          continue;
        }
        if (cell && cell.userEdited) continue;
        const v = getDisplayValue(row, col);
        if (v == null || String(v).trim() === '') continue;
        patches.push({ r: row, c: col, v, mat: true });
      }
    }
    return patches;
  }

  function applyPatches(patches) {
    for (const p of patches) {
      const key = `${p.r}:${p.c}`;
      let cell = synMap.get(key);
      if (!cell) {
        cell = { r: p.r, c: p.c, v: p.v, mat: true };
        synSheet.cells.push(cell);
        synMap.set(key, cell);
      } else if (!cell.userEdited) {
        cell.v = p.v;
        cell.mat = true;
        delete cell.f;
      }
    }
    synSheet.materializedCalc = true;
  }

  /** Recalc all live mass cells after a BD edit (returns patches only). */
  function recalcAfterBdEdit() {
    invalidateCaches();
    return materializeAll();
  }

  return {
    getDisplayValue,
    materializeAll,
    applyPatches,
    patchBdCell,
    recalcAfterBdEdit,
    invalidateCaches,
  };
}

/** Apply materialized values onto a transformed syn sheet (build + server). */
export function materializeSynSheet(bdSheet, synSheet) {
  const ctx = createSynCalcContext(bdSheet, synSheet);
  const patches = ctx.materializeAll();
  ctx.applyPatches(patches);
  return { patchCount: patches.length, synSheet };
}
