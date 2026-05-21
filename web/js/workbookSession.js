import { ref } from 'vue';
import { WorkbookEngine } from './workbookEngine.js';
import { displayValue, stripExcelErrorValue } from './bdStore.js';
import { BD_MASS_COL } from './bdColumnConfig.js';
import {
  buildBdColumnIndex,
  computeSumproduct,
  isSumproductCell,
} from './synthesisCalc.js';
import { isSynFilterEdit } from './synthesisPerf.js';

export function createWorkbookSession() {
  const revision = ref(0);
  const ready = ref(false);
  const loading = ref(false);
  const error = ref(null);
  const engine = new WorkbookEngine();
  let bdCols = null;
  let synGridGetter = null;
  let synFiltersDirty = false;
  const sumproductCache = new Map();
  const hfDisplayCache = new Map();

  function clearHfDisplayCache() {
    hfDisplayCache.clear();
  }

  async function loadSheets(sheets) {
    loading.value = true;
    error.value = null;
    ready.value = false;
    try {
      await new Promise((r) => setTimeout(r, 0));
      const bdEntry = sheets.find((s) => s.name === 'BD');
      if (bdEntry) {
        engine.loadSheetData('BD', bdEntry.data);
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
  }

  function getSynCell(row, col) {
    if (!synGridGetter) return '';
    const cell = synGridGetter(row, col);
    return cell?.v != null && cell.v !== '' ? String(cell.v) : '';
  }

  function clearSumproductCache() {
    sumproductCache.clear();
  }

  function setCellValue(sheetName, row, col, value) {
    if (sheetName === 'BD') {
      engine.setCellValue(sheetName, row, col, value);
      if (bdCols) {
        if (!bdCols[col]) bdCols[col] = [];
        bdCols[col][row] = value;
      }
      synFiltersDirty = true;
      clearSumproductCache();
      clearHfDisplayCache();
    } else if (sheetName === 'SYNTHESIS' && synGridGetter) {
      const cell = synGridGetter(row, col);
      if (cell) {
        cell.v = value;
        delete cell.f;
      }
      if (isSynFilterEdit(row, col)) {
        synFiltersDirty = true;
        clearSumproductCache();
      }
    }
    revision.value += 1;
  }

  function getDisplayValue(sheetName, row, col, cell) {
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
      const cached = stripExcelErrorValue(
        cell.v != null && cell.v !== '' ? String(cell.v) : ''
      );
      if (cached !== '') return cached;
    }
    if (sheetName === 'BD' && ready.value && engine.hasFormula(sheetName, row, col)) {
      const hfKey = `${row}:${col}`;
      if (hfDisplayCache.has(hfKey)) return hfDisplayCache.get(hfKey);
      const computed = engine.getCellValue(sheetName, row, col);
      if (computed !== '' && computed !== '#REF!') {
        hfDisplayCache.set(hfKey, computed);
        return computed;
      }
    }
    return displayValue(cell);
  }

  function isFormulaCell(sheetName, row, col, cell) {
    if (isSumproductCell(cell)) return true;
    if (sheetName === 'BD' && cell?.f) return true;
    if (ready.value) return engine.hasFormula(sheetName, row, col);
    return false;
  }

  function destroy() {
    engine.destroy();
    bdCols = null;
    synGridGetter = null;
    synFiltersDirty = false;
    clearSumproductCache();
    clearHfDisplayCache();
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
