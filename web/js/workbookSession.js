import { ref } from 'vue';
import { WorkbookEngineClient } from './workbookEngineClient.js?v=hf-worker2';
import {
  displayValue,
  formatMassDisplayValue,
} from './bdStore.js';
import { BD_MASS_COL } from './bdColumnConfig.js';
import {
  buildBdColumnIndex,
  buildBdL2Registry,
  buildSumproductL2Index,
  buildSynSectionRowList,
  buildVehColFilterIndex,
  computeSumproduct,
  computeSynSectionSum,
  computeSynAbDiff,
  computeAdaptationRowSum,
  isSumproductCell,
  isSynAdaptationSumCell,
  isSynAbDiffCell,
  isSynBlueSubsectionRow,
  isSynCalculatedMassCell,
  isSynSectionSumDataCell,
  isSynSumproductDataCell,
  affectsAdaptationSum,
  synCalcExcelCols,
} from './synthesisCalc.js?v=sumprod-ab1';
import { getSynAdaptBandNumeric, synRowMaaPresetRaw } from './synStore.js?v=syn-cicy1';
import { isSynFilterEdit, isSynProjHeaderGreenExcelCol } from './synthesisPerf.js';

export function createWorkbookSession() {
  const revision = ref(0);
  const ready = ref(false);
  const loading = ref(false);
  const error = ref(null);
  const engine = new WorkbookEngineClient();
  let bdCols = null;
  let sumproductL2Index = null;
  let vehColFilterIndex = null;
  let synGridGetter = null;
  let synSheetMeta = null;
  let synLabelGetter = null;
  let synRowClassGetter = null;
  let synSectionRows = [];
  const sumproductCache = new Map();

  function getBdValue(r, c) {
    if (ready.value && engine.hasFormula('BD', r, c)) {
      const v = engine.getCellValue('BD', r, c);
      if (v !== '') return v;
    }
    return bdCols?.[c]?.[r] ?? '';
  }

  function getSynLabel(row) {
    if (synLabelGetter) return synLabelGetter(row);
    return getSynCell(row, 'F');
  }

  function getSynRowClass(row) {
    return synRowClassGetter ? synRowClassGetter(row) : '';
  }

  function isBlueSubsectionRow(row) {
    return isSynBlueSubsectionRow(
      row,
      synSheetMeta,
      getSynLabel(row),
      getSynRowClass(row)
    );
  }

  function rebuildSumproductIndex() {
    if (!bdCols) {
      sumproductL2Index = null;
      return;
    }
    sumproductL2Index = buildSumproductL2Index(bdCols, getBdValue);
  }

  function rebuildVehColFilterIndex() {
    vehColFilterIndex = null;
    if (!bdCols || !synGridGetter) return;
    const map = new Map();
    for (const col of synCalcExcelCols(synSheetMeta)) {
      map.set(
        col,
        buildVehColFilterIndex(bdCols, col, getSynCell, getBdValue)
      );
    }
    vehColFilterIndex = map;
  }

  function rebuildSynSectionRows() {
    synSectionRows = [];
    if (!synSheetMeta || !synGridGetter) return;
    synSectionRows = buildSynSectionRowList(
      synSheetMeta,
      getSynLabel,
      getSynRowClass
    );
  }

  function invalidateSumproductCaches() {
    sumproductCache.clear();
    rebuildVehColFilterIndex();
    rebuildSynSectionRows();
  }

  function getBlueMaaValue(row, col) {
    const cacheKey = `b:${row}:${col}`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const filterRows = vehColFilterIndex?.get(col) ?? null;
    const n = computeSumproduct(
      bdCols,
      row,
      col,
      getSynCell,
      getBdValue,
      sumproductL2Index,
      filterRows
    );
    const rounded = Math.round(n * 10000) / 10000;
    sumproductCache.set(cacheKey, String(rounded));
    return rounded;
  }

  function getSectionMaaValue(sectionRow, col) {
    const cacheKey = `g:${sectionRow}:${col}`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const n = computeSynSectionSum(
      sectionRow,
      col,
      synSectionRows,
      isBlueSubsectionRow,
      getBlueMaaValue
    );
    sumproductCache.set(cacheKey, String(n));
    return n;
  }

  function parseSynNum(raw) {
    if (raw == null || raw === '') return 0;
    const n = parseFloat(String(raw).replace(',', '.'));
    return Number.isFinite(n) ? n : 0;
  }

  function resolveSynNumericAt(row, col) {
    const cell = synGridGetter?.(row, col) ?? null;
    if (isSynAdaptationSumCell(row, col)) {
      return computeAdaptationRowSum(getSynNumeric, col);
    }
    if (cell?.userEdited) {
      return parseSynNum(displayValue(cell));
    }
    if (
      isSynProjHeaderGreenExcelCol(col) &&
      row >= 3 &&
      row <= 22
    ) {
      const preset = synRowMaaPresetRaw(row, col);
      if (preset !== undefined) {
        return parseSynNum(preset);
      }
    }
    const synLabel = getSynLabel(row);
    const rowClass = getSynRowClass(row);
    if (
      bdCols &&
      synSectionRows.length &&
      isSynSectionSumDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    ) {
      return getSectionMaaValue(row, col);
    }
    if (
      bdCols &&
      isSynSumproductDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    ) {
      return getBlueMaaValue(row, col);
    }
    return parseSynNum(displayValue(cell));
  }

  function getAbDiffValue(row) {
    const cacheKey = `d:${row}:AB`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const n = computeSynAbDiff(resolveSynNumericAt, row);
    sumproductCache.set(cacheKey, String(n));
    return n;
  }

  async function loadSheets(sheets) {
    loading.value = true;
    error.value = null;
    ready.value = false;
    try {
      const bdEntry = sheets.find((s) => s.name === 'BD');
      if (bdEntry) {
        await engine.loadSheetData('BD', bdEntry.data);
        bdCols = buildBdColumnIndex(bdEntry.data);
        rebuildSumproductIndex();
        rebuildVehColFilterIndex();
      }
      ready.value = Boolean(bdEntry);
    } catch (e) {
      error.value = e?.message || String(e);
      console.error('Workbook engine load failed:', e);
    } finally {
      loading.value = false;
    }
  }

  function bindSynthesisGrid(getCell, sheet = null, getLabel = null, getRowClass = null) {
    synGridGetter = getCell;
    synSheetMeta = sheet;
    synLabelGetter = getLabel;
    synRowClassGetter = getRowClass;
    invalidateSumproductCaches();
  }

  function getSynCell(row, col) {
    if (!synGridGetter) return '';
    const cell = synGridGetter(row, col);
    return cell?.v != null && cell.v !== '' ? String(cell.v) : '';
  }

  function getSynNumeric(row, col) {
    if (!synGridGetter) return 0;
    return getSynAdaptBandNumeric(synGridGetter, row, col);
  }

  function getBdL2Registry() {
    if (!bdCols) return [];
    return buildBdL2Registry(bdCols, getBdValue);
  }

  async function setCellValue(sheetName, row, col, value) {
    if (sheetName === 'BD') {
      if (bdCols) {
        if (!bdCols[col]) bdCols[col] = [];
        bdCols[col][row] = value;
        rebuildSumproductIndex();
      }
      invalidateSumproductCaches();
      if (ready.value && engine.hasFormula(sheetName, row, col)) {
        try {
          await engine.setCellValue(sheetName, row, col, value);
        } catch (e) {
          console.warn('HyperFormula setCellValue failed:', e);
        }
      }
      revision.value += 1;
      return;
    } else if (sheetName === 'SYNTHESIS' && synGridGetter) {
      const cell = synGridGetter(row, col);
      if (cell) {
        cell.v = value;
        delete cell.f;
      }
      if (affectsAdaptationSum(row, col)) {
        sumproductCache.clear();
      }
      if (isSynFilterEdit(row, col)) {
        invalidateSumproductCaches();
      } else if (
        col === 'F' &&
        row >= (synSheetMeta?.dataStartRow ?? 15)
      ) {
        invalidateSumproductCaches();
      }
    }
    revision.value += 1;
  }

  function getDisplayValue(sheetName, row, col, cell) {
    if (sheetName === 'SYNTHESIS' && isSynAdaptationSumCell(row, col)) {
      const n = computeAdaptationRowSum(getSynNumeric, col);
      return String(n);
    }
    if (cell?.userEdited && !isSynAdaptationSumCell(row, col)) {
      return displayValue(cell);
    }
    if (
      sheetName === 'SYNTHESIS' &&
      isSynProjHeaderGreenExcelCol(col) &&
      row >= 3 &&
      row <= 22
    ) {
      const preset = synRowMaaPresetRaw(row, col);
      if (preset !== undefined) {
        if (preset == null || preset === '') return '';
        return String(preset);
      }
    }
    const synLabel = getSynLabel(row);
    const rowClass = getSynRowClass(row);
    if (sheetName === 'SYNTHESIS' && isSynAbDiffCell(row, col, synSheetMeta)) {
      return String(getAbDiffValue(row));
    }
    if (sheetName === 'SYNTHESIS' && bdCols && synSectionRows.length) {
      if (
        isSynSectionSumDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
      ) {
        return String(getSectionMaaValue(row, col));
      }
    }
    if (
      sheetName === 'SYNTHESIS' &&
      bdCols &&
      isSynSumproductDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    ) {
      if (!String(synLabel).trim()) {
        return displayValue(cell);
      }
      return String(getBlueMaaValue(row, col));
    }
    if (sheetName === 'BD' && col === BD_MASS_COL && cell) {
      const cached =
        cell.v != null && cell.v !== '' ? formatMassDisplayValue(cell.v) : '';
      if (cached !== '') return cached;
    }
    if (sheetName === 'BD' && ready.value && engine.hasFormula(sheetName, row, col)) {
      const computed = engine.getCellValue(sheetName, row, col);
      if (computed !== '' && computed !== '#REF!') return computed;
    }
    return displayValue(cell);
  }

  function isFormulaCell(sheetName, row, col, cell) {
    if (sheetName === 'SYNTHESIS' && isSynAdaptationSumCell(row, col)) {
      return true;
    }
    const synLabel = sheetName === 'SYNTHESIS' ? getSynLabel(row) : '';
    const rowClass = sheetName === 'SYNTHESIS' ? getSynRowClass(row) : '';
    if (isSynCalculatedMassCell(cell, row, col, synSheetMeta, synLabel, rowClass)) {
      return true;
    }
    if (sheetName === 'BD') {
      if (!cell?.f) return false;
      if (ready.value) return engine.hasFormula(sheetName, row, col);
      return true;
    }
    return false;
  }

  function isReadonlySynCell(row, col, cell) {
    const synLabel = getSynLabel(row);
    const rowClass = getSynRowClass(row);
    return (
      isSynAbDiffCell(row, col, synSheetMeta) ||
      isSynSectionSumDataCell(row, col, synSheetMeta, cell, synLabel, rowClass) ||
      isSynSumproductDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    );
  }

  function destroy() {
    engine.destroy();
    bdCols = null;
    sumproductL2Index = null;
    vehColFilterIndex = null;
    synGridGetter = null;
    synSheetMeta = null;
    synLabelGetter = null;
    synRowClassGetter = null;
    synSectionRows = [];
    sumproductCache.clear();
    ready.value = false;
  }

  return {
    revision,
    ready,
    loading,
    error,
    loadSheets,
    bindSynthesisGrid,
    setCellValue,
    getDisplayValue,
    isFormulaCell,
    isReadonlySynCell,
    getBdL2Registry,
    destroy,
  };
}
