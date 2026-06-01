import { ref } from 'vue';
import { WorkbookEngineClient } from './workbookEngineClient.js?v=hf-worker2';
import {
  displayValue,
  formatMassDisplayValue,
  buildCellMap,
  displayCellValue,
  bdSubsystemL1Col,
  bdSubsystemL2Col,
} from './bdStore.js';
import { BD_MASS_COL } from './bdColumnConfig.js';
import {
  buildBdColumnIndex,
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
  isSynFilterEdit,
  isSynAcanMassCol,
  isSynApbbMassCol,
  isSynVehicleMassCol,
  isSynAbDiffExcelCol,
  describeSynCellFormula,
  synVehicleMassExcelCol,
  affectsAdaptationSum,
} from './synthesisCalc.js?v=grid-perf2';
import { getSynAdaptBandNumeric, synRowMaaPresetRaw, synRowAcanPresetRaw, synRowApbbPresetRaw } from './synStore.js?v=grid-perf2';
import { isSynProjHeaderGreenExcelCol } from './synthesisPerf.js';

export function createWorkbookSession() {
  const revision = ref(0);
  const ready = ref(false);
  const loading = ref(false);
  const error = ref(null);
  const engine = new WorkbookEngineClient();
  let bdCols = null;
  let bdSheetMeta = null;
  let bdCellMap = null;
  let bdL1Col = null;
  let bdL2Col = null;
  let sumproductL2Index = null;
  /** @type {Map<string, number[]>} lazy BD row filter per vehicle column */
  let vehColFilterIndex = new Map();
  let synGridGetter = null;
  let synSheetMeta = null;
  let synLabelGetter = null;
  let synRowClassGetter = null;
  let synSectionRows = [];
  const sumproductCache = new Map();

  function rebuildBdDisplayContext(data) {
    if (!data) {
      bdSheetMeta = null;
      bdCellMap = null;
      bdL1Col = null;
      bdL2Col = null;
      return;
    }
    bdSheetMeta = data;
    bdCellMap = buildCellMap(data.cells, data.headerRows);
    bdL1Col = bdSubsystemL1Col(data);
    bdL2Col = bdSubsystemL2Col(data);
  }

  function getBdValue(r, c) {
    const cell = bdCellMap?.get(`${r}:${c}`);
    if (cell?.userEdited) {
      return cell.v != null && cell.v !== '' ? String(cell.v) : '';
    }
    if (ready.value && engine.hasFormula('BD', r, c)) {
      const v = engine.getCellValue('BD', r, c);
      if (v !== '' && v !== '#REF!') return v;
    }
    if (bdCellMap && bdSheetMeta) {
      const displayed = displayCellValue(
        bdCellMap,
        r,
        c,
        bdSheetMeta.sectionHeaderRows,
        bdSheetMeta.canonicalSectionByLabel,
        bdL1Col,
        bdL2Col
      );
      if (displayed !== '') return displayed;
    }
    const raw = bdCols?.[c]?.[r];
    return raw != null && raw !== '' ? String(raw) : '';
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

  function getVehColFilterRows(col) {
    if (!bdCols || !synGridGetter) return null;
    const excelCol = synVehicleMassExcelCol(col);
    if (!vehColFilterIndex.has(excelCol)) {
      vehColFilterIndex.set(
        excelCol,
        buildVehColFilterIndex(bdCols, excelCol, getSynCell, getBdValue)
      );
    }
    return vehColFilterIndex.get(excelCol);
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

  function invalidateSumproductCaches(vehCol = null) {
    sumproductCache.clear();
    if (vehCol) {
      vehColFilterIndex.delete(synVehicleMassExcelCol(vehCol));
    } else {
      vehColFilterIndex.clear();
    }
    rebuildSynSectionRows();
  }

  function getBlueMaaValue(row, col) {
    const excelCol = synVehicleMassExcelCol(col);
    const cacheKey = `b:${row}:${excelCol}`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const filterRows = getVehColFilterRows(excelCol);
    const n = computeSumproduct(
      bdCols,
      row,
      excelCol,
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
    const excelCol = synVehicleMassExcelCol(col);
    const cacheKey = `g:${sectionRow}:${excelCol}`;
    if (sumproductCache.has(cacheKey)) {
      return parseFloat(sumproductCache.get(cacheKey)) || 0;
    }
    const n = computeSynSectionSum(
      sectionRow,
      excelCol,
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
      return computeAdaptationRowSum(getAdaptationBlockNumeric, col);
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
    if (
      isSynAcanMassCol(col) &&
      row >= 3 &&
      row <= 22
    ) {
      const preset = synRowAcanPresetRaw(row, col);
      if (preset !== undefined) {
        return parseSynNum(preset);
      }
    }
    if (
      isSynApbbMassCol(col) &&
      row >= 3 &&
      row <= 22
    ) {
      const preset = synRowApbbPresetRaw(row, col);
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
    try {
      const bdEntry = sheets.find((s) => s.name === 'BD');
      if (bdEntry) {
        // Live SOMMEPROD — index BD first; never block UI on HyperFormula worker load.
        bdCols = buildBdColumnIndex(bdEntry.data);
        rebuildBdDisplayContext(bdEntry.data);
        rebuildSumproductIndex();
        vehColFilterIndex.clear();
        invalidateSumproductCaches();
        ready.value = true;
        revision.value += 1;
        loading.value = false;
        void engine
          .loadSheetData('BD', bdEntry.data)
          .catch((engineErr) => {
            console.warn(
              'HyperFormula BD load failed — Synthesis mass uses BD index:',
              engineErr
            );
          })
          .finally(() => {
            revision.value += 1;
          });
        return;
      }
      ready.value = false;
    } catch (e) {
      error.value = e?.message || String(e);
      console.error('Workbook session load failed:', e);
    } finally {
      loading.value = false;
      revision.value += 1;
    }
  }

  let boundSynSheet = null;

  function bindSynthesisGrid(getCell, sheet = null, getLabel = null, getRowClass = null) {
    synGridGetter = getCell;
    synLabelGetter = getLabel;
    synRowClassGetter = getRowClass;
    if (boundSynSheet === sheet && synSheetMeta === sheet) {
      return;
    }
    boundSynSheet = sheet;
    synSheetMeta = sheet;
    invalidateSumproductCaches();
    revision.value += 1;
  }

  function getSynCell(row, col) {
    if (!synGridGetter) return '';
    if (col === 'F' && synLabelGetter) {
      const label = synLabelGetter(row);
      if (label != null && String(label).trim() !== '') {
        return String(label).trim();
      }
    }
    const cell = synGridGetter(row, col);
    return cell?.v != null && cell.v !== '' ? String(cell.v) : '';
  }

  function getAdaptationBlockNumeric(row, col) {
    if (isSynVehicleMassCol(col) || isSynAbDiffExcelCol(col)) {
      return resolveSynNumericAt(row, col);
    }
    if (!synGridGetter) return 0;
    return getSynAdaptBandNumeric(synGridGetter, row, col);
  }

  function getSynFormula(row, col) {
    if (!synSheetMeta) return '';
    const cell = synGridGetter?.(row, col) ?? null;
    const label = getSynLabel(row);
    const rowClass = getSynRowClass(row);
    return describeSynCellFormula(
      row,
      col,
      synSheetMeta,
      label,
      rowClass,
      getSynCell
    );
  }

  async function setCellValue(sheetName, row, col, value) {
    if (sheetName === 'BD') {
      if (bdCols) {
        if (!bdCols[col]) bdCols[col] = [];
        bdCols[col][row] = value;
        if (bdCellMap) {
          const key = `${row}:${col}`;
          let cell = bdCellMap.get(key);
          if (!cell) {
            cell = { r: row, c: col };
            bdCellMap.set(key, cell);
          }
          cell.v = value;
          cell.userEdited = true;
          delete cell.f;
        }
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
        invalidateSumproductCaches(col);
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
      if (!synGridGetter) return '0';
      const n = computeAdaptationRowSum(getAdaptationBlockNumeric, col);
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
    if (
      sheetName === 'SYNTHESIS' &&
      isSynAcanMassCol(col) &&
      row >= 3 &&
      row <= 22
    ) {
      const preset = synRowAcanPresetRaw(row, col);
      if (preset !== undefined) {
        if (preset == null || preset === '') return '';
        return String(preset);
      }
    }
    if (
      sheetName === 'SYNTHESIS' &&
      isSynApbbMassCol(col) &&
      row >= 3 &&
      row <= 22
    ) {
      const preset = synRowApbbPresetRaw(row, col);
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
      isSynSectionSumDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    ) {
      return '0';
    }
    if (
      sheetName === 'SYNTHESIS' &&
      bdCols &&
      isSynSumproductDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    ) {
      if (!String(synLabel).trim()) {
        return '0';
      }
      return String(getBlueMaaValue(row, col));
    }
    if (
      sheetName === 'SYNTHESIS' &&
      isSynSumproductDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    ) {
      return '0';
    }
    if (
      sheetName === 'SYNTHESIS' &&
      isSynCalculatedMassCell(cell, row, col, synSheetMeta, synLabel, rowClass)
    ) {
      return '0';
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
    bdSheetMeta = null;
    bdCellMap = null;
    bdL1Col = null;
    bdL2Col = null;
    sumproductL2Index = null;
    vehColFilterIndex.clear();
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
    getSynFormula,
    destroy,
  };
}
