import { ref } from 'vue';
import { WorkbookEngineClient } from './workbookEngineClient.js?v=hf-worker2';
import {
  displayValue,
  stripExcelErrorValue,
  formatMassDisplayValue,
} from './bdStore.js';
import { BD_MASS_COL } from './bdColumnConfig.js';
import {
  buildBdColumnIndex,
  computeSumproduct,
  computeAdaptationRowSum,
  isSumproductCell,
  isSynAdaptationSumCell,
  affectsAdaptationSum,
} from './synthesisCalc.js?v=adapt-sum1';
import { getSynAdaptBandNumeric } from './synStore.js?v=adapt-sum1';
import { isSynFilterEdit } from './synthesisPerf.js';

export function createWorkbookSession() {
  const revision = ref(0);
  const ready = ref(false);
  const loading = ref(false);
  const error = ref(null);
  const engine = new WorkbookEngineClient();
  let bdCols = null;
  let synGridGetter = null;
  let synFiltersDirty = false;
  const sumproductCache = new Map();
  const adaptationSumCache = new Map();
  let adaptationSumDirty = true;

  async function loadSheets(sheets) {
    loading.value = true;
    error.value = null;
    ready.value = false;
    try {
      const bdEntry = sheets.find((s) => s.name === 'BD');
      if (bdEntry) {
        await engine.loadSheetData('BD', bdEntry.data);
        bdCols = buildBdColumnIndex(bdEntry.data);
      }
      ready.value = Boolean(bdEntry);
    } catch (e) {
      error.value = e?.message || String(e);
      console.error('Workbook engine load failed:', e);
    } finally {
      loading.value = false;
    }
  }

  function bindSynthesisGrid(getCell) {
    synGridGetter = getCell;
    clearAdaptationSumCache();
  }

  function getSynCell(row, col) {
    if (!synGridGetter) return '';
    const cell = synGridGetter(row, col);
    return cell?.v != null && cell.v !== '' ? String(cell.v) : '';
  }

  function clearSumproductCache() {
    sumproductCache.clear();
  }

  function clearAdaptationSumCache() {
    adaptationSumCache.clear();
    adaptationSumDirty = true;
  }

  function getSynNumeric(row, col) {
    if (!synGridGetter) return 0;
    return getSynAdaptBandNumeric(synGridGetter, row, col);
  }

  async function setCellValue(sheetName, row, col, value) {
    if (sheetName === 'BD') {
      if (bdCols) {
        if (!bdCols[col]) bdCols[col] = [];
        bdCols[col][row] = value;
      }
      synFiltersDirty = true;
      clearSumproductCache();
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
        clearAdaptationSumCache();
      }
      if (isSynFilterEdit(row, col)) {
        synFiltersDirty = true;
        clearSumproductCache();
      }
    }
    revision.value += 1;
  }

  function getDisplayValue(sheetName, row, col, cell) {
    if (cell?.userEdited && !isSynAdaptationSumCell(row, col)) {
      return displayValue(cell);
    }
    if (sheetName === 'SYNTHESIS' && isSynAdaptationSumCell(row, col)) {
      const cacheKey = `adapt:${col}`;
      if (!adaptationSumDirty && adaptationSumCache.has(cacheKey)) {
        return adaptationSumCache.get(cacheKey);
      }
      const n = computeAdaptationRowSum(getSynNumeric, col);
      const out = String(n);
      adaptationSumCache.set(cacheKey, out);
      adaptationSumDirty = false;
      return out;
    }
    if (sheetName === 'SYNTHESIS' && isSumproductCell(cell) && bdCols) {
      if (!synFiltersDirty && cell?.v != null && cell.v !== '') {
        return String(cell.v);
      }
      const cacheKey = `${row}:${col}`;
      if (sumproductCache.has(cacheKey)) {
        return sumproductCache.get(cacheKey);
      }
      const getBdValue = (r, c) => {
        if (ready.value && engine.hasFormula('BD', r, c)) {
          const v = engine.getCellValue('BD', r, c);
          if (v !== '') return v;
        }
        return bdCols[c]?.[r] ?? '';
      };
      const n = computeSumproduct(
        bdCols,
        row,
        col,
        getSynCell,
        getBdValue
      );
      const rounded = Math.round(n * 10000) / 10000;
      const out = String(rounded);
      sumproductCache.set(cacheKey, out);
      return out;
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
    if (isSumproductCell(cell)) return true;
    if (sheetName === 'BD') {
      if (!cell?.f) return false;
      if (ready.value) return engine.hasFormula(sheetName, row, col);
      return true;
    }
    return false;
  }

  function destroy() {
    engine.destroy();
    bdCols = null;
    synGridGetter = null;
    synFiltersDirty = false;
    clearSumproductCache();
    clearAdaptationSumCache();
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
    destroy,
  };
}
