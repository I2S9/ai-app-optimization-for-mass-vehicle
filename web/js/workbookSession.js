import { ref } from 'vue';
import { WorkbookEngineClient } from './workbookEngineClient.js?v=hf-worker3';
import {
  displayValue,
  formatMassDisplayValue,
  buildCellMap,
  displayCellValue,
  bdSubsystemL1Col,
  bdSubsystemL2Col,
} from './bdStore.js';
import { BD_MASS_COL, BD_SUBSYSTEM_L2_COL } from './bdColumnConfig.js';
import { canonicalL2MatchKey } from './bdTranslate.js';
import {
  buildBdColumnIndex,
  buildSumproductL2Index,
  buildSynSectionRowList,
  buildVehColFilterIndex,
  resolveSynSumproductValue,
  SYN_SUMPRODUCT_REF,
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
  classifyBdEditForSynMass,
  describeSynCellFormula,
  synVehicleMassExcelCol,
  affectsAdaptationSum,
  synAdaptationSumRow,
  synAdaptationSumFromRow,
  synAdaptationSumToRow,
} from './synthesisCalc.js?v=grid-perf9';
import {
  getSynAdaptBandNumeric,
  synRowMaaPresetRaw,
  synRowAcanPresetRaw,
  synRowApbbPresetRaw,
  synRowBdboPresetRaw,
  synRowBscePresetRaw,
  isSynBodyEmptyFromRow27Cell,
  synFilterStaleProjectDisplay,
} from './synStore.js?v=syn-form4';
import { isSynProjHeaderGreenExcelCol } from './synthesisPerf.js';

export function createWorkbookSession() {
  const revision = ref(0);
  /** Grid display invalidation — decoupled from full revision bumps. */
  const displayTick = ref(0);
  /** Synthesis live-calc invalidation (BD edit → Syn SUMPRODUCT only). */
  const synCalcTick = ref(0);
  /**
   * True once a BD edit has touched a calc input (mass / filter / L2 index).
   * Flips Synthesis out of the build-time "materialized" snapshot so impacted
   * SUMPRODUCT / section-sum / AB-diff cells recompute live. Only the cells whose
   * cache entry was invalidated recompute — untouched cells stay cached.
   */
  const liveBdEdited = ref(false);
  /** ADAPTATION row-26 SUM only (cols C..J) — lightweight, no SUMPRODUCT. */
  const adaptationSumTick = ref(0);
  let displayDirtyAll = true;
  const displayDirtyKeys = new Set();
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
  /** @type {Map<number, string>} BD row → canonical L2 key for O(1) index patch */
  let bdRowToL2Key = new Map();
  /** @type {Map<string, number[]>} lazy BD row filter per vehicle column */
  let vehColFilterIndex = new Map();
  let synGridGetter = null;
  let synSheetMeta = null;
  let synLabelGetter = null;
  let synRowClassGetter = null;
  let synSectionRows = [];
  const sumproductCache = new Map();
  /** Last BD payload passed to loadSheets — lazy index if an edit arrives early. */
  let lastBdPayload = null;

  function bumpDisplay(all = true, keys = []) {
    if (all) {
      displayDirtyAll = true;
      displayDirtyKeys.clear();
    } else {
      if (displayDirtyAll) return;
      for (const k of keys) displayDirtyKeys.add(k);
    }
    displayTick.value += 1;
  }

  function takeDisplayDirty() {
    const all = displayDirtyAll;
    const keys = displayDirtyAll ? null : new Set(displayDirtyKeys);
    displayDirtyAll = false;
    displayDirtyKeys.clear();
    return { all, keys };
  }

  function rebuildBdDisplayContext(data) {
    if (!data) {
      bdSheetMeta = null;
      bdCellMap = null;
      bdL1Col = null;
      bdL2Col = null;
      return;
    }
    bdSheetMeta = data;
    bdCellMap =
      data.cellMap instanceof Map
        ? data.cellMap
        : buildCellMap(data.cells, data.headerRows);
    bdL1Col = bdSubsystemL1Col(data);
    bdL2Col = bdSubsystemL2Col(data);
  }

  function getBdValue(r, c) {
    const cell = bdCellMap ? bdCellMap.get(`${r}:${c}`) : null;
    if (cell && cell.userEdited) {
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
    const raw = bdCols && bdCols[c] ? bdCols[c][r] : undefined;
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
      bdRowToL2Key = new Map();
      return;
    }
    sumproductL2Index = buildSumproductL2Index(bdCols, getBdValue);
    bdRowToL2Key = new Map();
    for (const [key, rows] of sumproductL2Index) {
      for (const r of rows) bdRowToL2Key.set(r, key);
    }
  }

  function patchSumproductL2IndexForRow(row) {
    if (!sumproductL2Index) return;
    const oldKey = bdRowToL2Key.get(row);
    if (oldKey) {
      const arr = sumproductL2Index.get(oldKey);
      if (arr) {
        const i = arr.indexOf(row);
        if (i >= 0) arr.splice(i, 1);
      }
    }
    const qRaw =
      getBdValue(row, 'Q') != null
        ? getBdValue(row, 'Q')
        : bdCols && bdCols.Q
          ? bdCols.Q[row]
          : '';
    const q = String(qRaw != null ? qRaw : '').trim();
    if (q !== 'S') {
      bdRowToL2Key.delete(row);
      return;
    }
    const l2Raw =
      getBdValue(row, BD_SUBSYSTEM_L2_COL) != null
        ? getBdValue(row, BD_SUBSYSTEM_L2_COL)
        : bdCols && bdCols[BD_SUBSYSTEM_L2_COL]
          ? bdCols[BD_SUBSYSTEM_L2_COL][row]
          : '';
    const newKey = canonicalL2MatchKey(l2Raw);
    if (!newKey) {
      bdRowToL2Key.delete(row);
      return;
    }
    if (!sumproductL2Index.has(newKey)) sumproductL2Index.set(newKey, []);
    sumproductL2Index.get(newKey).push(row);
    bdRowToL2Key.set(row, newKey);
  }

  function bdL2KeyAtRow(row) {
    const l2Raw =
      getBdValue(row, BD_SUBSYSTEM_L2_COL) != null
        ? getBdValue(row, BD_SUBSYSTEM_L2_COL)
        : bdCols && bdCols[BD_SUBSYSTEM_L2_COL]
          ? bdCols[BD_SUBSYSTEM_L2_COL][row]
          : '';
    return canonicalL2MatchKey(l2Raw);
  }

  /** Drop SUMPRODUCT cache entries for one L2 label (O(cache), not O(all Syn cells)). */
  function invalidateSumproductForL2(l2Key, { sectionSums = false } = {}) {
    if (!l2Key) return;
    for (const k of [...sumproductCache.keys()]) {
      if (!k.startsWith('b:')) continue;
      const synRow = parseInt(k.split(':')[1], 10);
      if (!Number.isFinite(synRow)) continue;
      const labelKey = canonicalL2MatchKey(getSynCell(synRow, 'F'));
      if (labelKey === l2Key) sumproductCache.delete(k);
    }
    if (sectionSums) {
      for (const k of [...sumproductCache.keys()]) {
        if (k.startsWith('g:')) sumproductCache.delete(k);
      }
    }
    // AB diff (display AB = V − Y) is derived from blue mass values, so it must
    // be recomputed whenever any blue mass changes. There are only a few of these
    // rows, so always dropping them is cheap and avoids stale Δ values.
    for (const k of [...sumproductCache.keys()]) {
      if (k.startsWith('d:')) sumproductCache.delete(k);
    }
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
    rebuildSynSectionRows();
  }

  function invalidateAfterBdEdit(row, col) {
    revision.value += 1;
    bumpDisplay(false, [`${row}:${col}`]);
  }

  function ensureBdColsFromPayload() {
    if (bdCols || !lastBdPayload) return false;
    bdCols = buildBdColumnIndex(lastBdPayload);
    rebuildBdDisplayContext(lastBdPayload);
    rebuildSumproductIndex();
    vehColFilterIndex.clear();
    ready.value = true;
    return true;
  }

  /**
   * A Database edit was committed. Invalidate exactly the Synthesis live-mass
   * caches that the edited BD column feeds, so the dependent SOMMEPROD / section
   * sums / Δ cells recompute on the next render.
   *
   * The set of relevant columns (mass V, sub-system L2 AU, gate Q, and every
   * filter column A,B,C,D,E,F,G,I,K,L,O) is defined once in
   * {@link classifyBdEditForSynMass} — the previous logic only watched O/P/Q/AU/V
   * and therefore silently ignored edits to the actual filter columns A–L.
   */
  function applyBdEditInvalidation(row, col) {
    const kind = classifyBdEditForSynMass(col);
    if (!kind) return;

    // From now on Synthesis must recompute impacted cells live rather than
    // serving the build-time materialized snapshot.
    liveBdEdited.value = true;

    if (kind === 'l2') {
      // The row may move from one Synthesis label to another: invalidate both the
      // previous and the new L2 group so neither keeps a stale total.
      const oldKey = bdRowToL2Key.get(row);
      patchSumproductL2IndexForRow(row);
      const newKey = bdL2KeyAtRow(row);
      invalidateSumproductForL2(oldKey, { sectionSums: true });
      invalidateSumproductForL2(newKey, { sectionSums: true });
    } else if (kind === 'gate') {
      // Q="S" gates a row in/out of every vehicle column's filter set and the L2
      // index, so the per-column filter caches and all blue/green/Δ totals reset.
      patchSumproductL2IndexForRow(row);
      invalidateSumproductCaches();
    } else if (kind === 'mass') {
      const l2 = bdL2KeyAtRow(row);
      if (synGridGetter && l2) {
        invalidateSumproductForL2(l2, { sectionSums: true });
      } else {
        invalidateSumproductCaches();
      }
    } else if (kind === 'filter') {
      // A filter column changed which BD rows match — for every vehicle column —
      // so drop the cached per-column filter sets and all dependent totals.
      invalidateSumproductCaches();
    }
    synCalcTick.value += 1;
  }

  function getBlueMaaValue(row, col) {
    const excelCol = synVehicleMassExcelCol(col);
    const cacheKey = `b:${row}:${excelCol}`;
    if (sumproductCache.has(cacheKey)) {
      const cached = sumproductCache.get(cacheKey);
      if (cached === SYN_SUMPRODUCT_REF) return SYN_SUMPRODUCT_REF;
      return parseFloat(cached) || 0;
    }
    const filterRows = getVehColFilterRows(excelCol);
    const result = resolveSynSumproductValue(
      bdCols,
      row,
      excelCol,
      getSynCell,
      getBdValue,
      sumproductL2Index,
      filterRows
    );
    if (result === SYN_SUMPRODUCT_REF) {
      sumproductCache.set(cacheKey, SYN_SUMPRODUCT_REF);
      return SYN_SUMPRODUCT_REF;
    }
    sumproductCache.set(cacheKey, String(result));
    return result;
  }

  function getSectionMaaValue(sectionRow, col) {
    const excelCol = synVehicleMassExcelCol(col);
    const cacheKey = `g:${sectionRow}:${excelCol}`;
    if (sumproductCache.has(cacheKey)) {
      const cached = sumproductCache.get(cacheKey);
      if (cached === SYN_SUMPRODUCT_REF) return SYN_SUMPRODUCT_REF;
      return parseFloat(cached) || 0;
    }
    const n = computeSynSectionSum(
      sectionRow,
      excelCol,
      synSectionRows,
      isBlueSubsectionRow,
      getBlueMaaValue,
      synSheetMeta && synSheetMeta.effectiveLastRow != null
        ? synSheetMeta.effectiveLastRow
        : synSheetMeta && synSheetMeta.lastRow != null
          ? synSheetMeta.lastRow
          : 422
    );
    if (n === SYN_SUMPRODUCT_REF) {
      sumproductCache.set(cacheKey, SYN_SUMPRODUCT_REF);
      return SYN_SUMPRODUCT_REF;
    }
    sumproductCache.set(cacheKey, String(n));
    return n;
  }

  function parseSynNum(raw) {
    if (raw == null || raw === '') return 0;
    const n = parseFloat(String(raw).replace(',', '.'));
    return Number.isFinite(n) ? n : 0;
  }

  function resolveSynNumericAt(row, col) {
    const cell = synGridGetter ? synGridGetter(row, col) : null;
    if (isSynAdaptationSumCell(row, col, synSheetMeta)) {
      return computeAdaptationRowSum(
        getAdaptationBlockNumeric,
        col,
        synAdaptationSumFromRow(synSheetMeta),
        synAdaptationSumToRow(synSheetMeta)
      );
    }
    if (cell && cell.userEdited) {
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
    if (isSynVehicleMassCol(col) && row >= 3 && row <= 22) {
      const preset = synRowBdboPresetRaw(row, col);
      if (preset !== undefined) {
        return parseSynNum(preset);
      }
      const bscePreset = synRowBscePresetRaw(row, col);
      if (bscePreset !== undefined) {
        return parseSynNum(bscePreset);
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
        lastBdPayload = bdEntry.data;
        // Live SOMMEPROD — index BD first; never block UI on HyperFormula worker load.
        bdCols = buildBdColumnIndex(bdEntry.data);
        rebuildBdDisplayContext(bdEntry.data);
        rebuildSumproductIndex();
        vehColFilterIndex.clear();
        invalidateSumproductCaches();
        ready.value = true;
        synCalcTick.value += 1;
        revision.value += 1;
        bumpDisplay(true);
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
            bumpDisplay(true);
          });
        return;
      }
      ready.value = false;
    } catch (e) {
      error.value = (e && e.message) || String(e);
      console.error('Workbook session load failed:', e);
    } finally {
      loading.value = false;
      revision.value += 1;
      bumpDisplay(true);
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
    rebuildSynSectionRows();
    invalidateSumproductCaches();
    synCalcTick.value += 1;
    revision.value += 1;
    bumpDisplay(true);
  }

  function synFilterPresetRaw(row, col) {
    if (row < 3 || row > 22) return undefined;
    const excelCol = synVehicleMassExcelCol(col);
    const maa = synRowMaaPresetRaw(row, excelCol);
    if (maa !== undefined && maa != null && maa !== '') return String(maa);
    const acan = synRowAcanPresetRaw(row, excelCol);
    if (acan !== undefined && acan != null && acan !== '') return String(acan);
    const apbb = synRowApbbPresetRaw(row, excelCol);
    if (apbb !== undefined && apbb != null && apbb !== '') return String(apbb);
    const bdbo = synRowBdboPresetRaw(row, excelCol);
    if (bdbo !== undefined && bdbo != null && bdbo !== '') return String(bdbo);
    const bsce = synRowBscePresetRaw(row, excelCol);
    if (bsce !== undefined && bsce != null && bsce !== '') return String(bsce);
    return undefined;
  }

  function getSynCell(row, col) {
    if (!synGridGetter) return '';
    if (col === 'F' && synLabelGetter) {
      const label = synLabelGetter(row);
      if (label != null && String(label).trim() !== '') {
        return String(label).trim();
      }
    }
    const excelCol = synVehicleMassExcelCol(col);
    const cell = synGridGetter(row, excelCol);
    if (cell && cell.v != null && cell.v !== '') return String(cell.v);
    const preset = synFilterPresetRaw(row, excelCol);
    return preset != null ? preset : '';
  }

  function getAdaptationBlockNumeric(row, col) {
    if (!synGridGetter) return 0;
    return getSynAdaptBandNumeric(synGridGetter, row, col);
  }

  function getSynFormula(row, col) {
    if (!synSheetMeta) return '';
    const cell = synGridGetter ? synGridGetter(row, col) : null;
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
    const perfOn =
      typeof window !== 'undefined' && Boolean(window.__perfEdits);
    const t0 = perfOn ? performance.now() : 0;
    if (sheetName === 'BD') {
      ensureBdColsFromPayload();
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
        applyBdEditInvalidation(row, col);
      }
      const tInv = perfOn ? performance.now() : 0;
      if (ready.value && engine.hasFormula(sheetName, row, col)) {
        try {
          await engine.setCellValue(sheetName, row, col, value);
        } catch (e) {
          console.warn('HyperFormula setCellValue failed:', e);
        }
      }
      const tEng = perfOn ? performance.now() : 0;
      invalidateAfterBdEdit(row, col);
      if (perfOn) {
        const t1 = performance.now();
        console.log(
          `[perf] edit BD ${col}${row}: total ${Math.round(t1 - t0)}ms | ` +
            `invalidate ${Math.round(tInv - t0)}ms | ` +
            `engine ${Math.round(tEng - tInv)}ms | ` +
            `post ${Math.round(t1 - tEng)}ms`
        );
      }
      return;
    } else if (sheetName === 'SYNTHESIS' && synGridGetter) {
      const cell = synGridGetter(row, col);
      if (cell) {
        cell.v = value;
        cell.userEdited = true;
        delete cell.f;
        delete cell.mat;
      }
      if (affectsAdaptationSum(row, col, synSheetMeta)) {
        adaptationSumTick.value += 1;
        bumpDisplay(false, [
          `${synAdaptationSumRow(synSheetMeta)}:${col}`,
        ]);
        revision.value += 1;
        return;
      }
      if (isSynFilterEdit(row, col)) {
        invalidateSumproductCaches(col);
        synCalcTick.value += 1;
      } else if (
        col === 'F' &&
        row >= (synSheetMeta && synSheetMeta.dataStartRow != null ? synSheetMeta.dataStartRow : 15)
      ) {
        invalidateSumproductCaches();
        synCalcTick.value += 1;
      }
    }
    revision.value += 1;
    bumpDisplay(false, [`${row}:${col}`]);
    if (perfOn) {
      const t1 = performance.now();
      console.log(
        `[perf] edit ${sheetName} ${col}${row}: ${Math.round(t1 - t0)}ms`
      );
    }
  }

  function getDisplayValue(sheetName, row, col, cell) {
    if (sheetName === 'SYNTHESIS' && isSynBodyEmptyFromRow27Cell(row, col)) {
      return '';
    }
    if (sheetName === 'SYNTHESIS' && isSynAdaptationSumCell(row, col, synSheetMeta)) {
      if (!synGridGetter) return '';
      const n = computeAdaptationRowSum(
        getAdaptationBlockNumeric,
        col,
        synAdaptationSumFromRow(synSheetMeta),
        synAdaptationSumToRow(synSheetMeta)
      );
      return String(n);
    }
    if (sheetName === 'SYNTHESIS' && cell && cell.mat && !cell.userEdited) {
      const synLabel = getSynLabel(row);
      const rowClass = getSynRowClass(row);
      const isLiveMass = isSynCalculatedMassCell(
        cell,
        row,
        col,
        synSheetMeta,
        synLabel,
        rowClass
      );
      if (!isLiveMass) {
        return synFilterStaleProjectDisplay(
          displayValue(cell),
          cell,
          row,
          col,
          synSheetMeta,
          null,
          getSynLabel(row),
          getSynRowClass(row)
        );
      }
    }
    if (cell && cell.userEdited && !isSynAdaptationSumCell(row, col, synSheetMeta)) {
      const edited = displayValue(cell);
      return synFilterStaleProjectDisplay(
        edited,
        cell,
        row,
        col,
        synSheetMeta,
        null,
        getSynLabel(row),
        getSynRowClass(row)
      );
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
    if (
      sheetName === 'SYNTHESIS' &&
      isSynVehicleMassCol(col) &&
      row >= 3 &&
      row <= 22
    ) {
      const preset = synRowBdboPresetRaw(row, col);
      if (preset !== undefined) {
        if (preset == null || preset === '') return '';
        return String(preset);
      }
      const bscePreset = synRowBscePresetRaw(row, col);
      if (bscePreset !== undefined) {
        if (bscePreset == null || bscePreset === '') return '';
        return String(bscePreset);
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
      return SYN_SUMPRODUCT_REF;
    }
    if (
      sheetName === 'SYNTHESIS' &&
      isSynSumproductDataCell(row, col, synSheetMeta, cell, synLabel, rowClass)
    ) {
      if (!bdCols) return SYN_SUMPRODUCT_REF;
      if (!String(synLabel).trim()) return SYN_SUMPRODUCT_REF;
      const blue = getBlueMaaValue(row, col);
      return String(blue);
    }
    if (
      sheetName === 'SYNTHESIS' &&
      isSynCalculatedMassCell(cell, row, col, synSheetMeta, synLabel, rowClass)
    ) {
      return SYN_SUMPRODUCT_REF;
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
    const raw = displayValue(cell);
    if (sheetName !== 'SYNTHESIS') return raw;
    return synFilterStaleProjectDisplay(
      raw,
      cell,
      row,
      col,
      synSheetMeta,
      null,
      getSynLabel(row),
      getSynRowClass(row)
    );
  }

  function isFormulaCell(sheetName, row, col, cell) {
    if (sheetName === 'SYNTHESIS' && isSynAdaptationSumCell(row, col, synSheetMeta)) {
      return true;
    }
    const synLabel = sheetName === 'SYNTHESIS' ? getSynLabel(row) : '';
    const rowClass = sheetName === 'SYNTHESIS' ? getSynRowClass(row) : '';
    if (isSynCalculatedMassCell(cell, row, col, synSheetMeta, synLabel, rowClass)) {
      return true;
    }
    if (sheetName === 'BD') {
      if (!cell || !cell.f) return false;
      if (ready.value) return engine.hasFormula(sheetName, row, col);
      return true;
    }
    return false;
  }

  function isReadonlySynCell(row, col, cell) {
    const synLabel = getSynLabel(row);
    const rowClass = getSynRowClass(row);
    return (
      isSynAdaptationSumCell(row, col, synSheetMeta) ||
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
    lastBdPayload = null;
    ready.value = false;
  }

  return {
    revision,
    displayTick,
    synCalcTick,
    liveBdEdited,
    adaptationSumTick,
    takeDisplayDirty,
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
