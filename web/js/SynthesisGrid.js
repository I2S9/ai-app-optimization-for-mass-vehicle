/**
 * Synthesis grid — Excel A–E hidden; display letters F→A, G→B, …
 */
import {
  ref,
  computed,
  shallowRef,
  onMounted,
  onUnmounted,
  watch,
  watchEffect,
  nextTick,
} from 'vue';
import { getCell, buildCellMap } from './bdStore.js?v=input-fix3';
import { upsertRawCell } from './sessionPersistence.js?v=edit-fix2';
import {
  isSynAdaptationSumCell,
  affectsAdaptationSum,
  synAdaptationSumRow,
  computeAdaptationRowSum,
  isSynCalculatedMassCell,
} from './synthesisCalc.js?v=grid-perf2';
import { createGridCellEditor } from './gridCellEdit.js?v=grid-nav4';
import { createGridCellNavigation } from './gridCellNavigation.js?v=grid-search4';
import {
  buildSearchIndex,
  createGridSearchController,
} from './gridSearch.js?v=grid-search4';
import {
  computeSynBodyRows,
  isSynPanelGapEntry,
  synDisplayValue,
  buildSynPillarColumns,
  synCellInlineStyle,
  synProjectCellClass,
  synMetricCellClass,
  synHeaderPanelVehicleClass,
  synCellAccentClass,
  synHdrEnergyValueClass,
  synHdrEnergyValueStyle,
  synHdrBevTextClass,
  synHdrBevTextStyle,
  synHdrBevDisplayHtml,
  synFilterGreyColClass,
  synAdaptBandColClass,
  synSpotBlueColClass,
  synMetricCjWhiteColClass,
  formatSynNumericDisplay,
  synRowStyleClass,
  isSynPillarCol,
  isSynPillarColAtRow,
  synPillarLetterForRow,
  synPillarLettersFromTitle,
  findSynEchappementRow,
  SYN_PILLAR_LETTER_ROW_STEP,
  SYN_PILLAR_OVERLAY_COLS,
  isSynMetricRow,
  isSynHeaderPanelRow,
  isSynHeaderPanelBoldCol,
  isSynAdaptCjValueBold,
  isSynSp2DisplayExcelCol,
  synPillarAccentClass,
  isSynProjHeaderGreenCol,
  isSynStlaSlashGreenCell,
  isSynTargetTextGreenCell,
  synProjHeaderGreenStyle,
  isSynProjHeaderYellowCol,
  synProjHeaderYellowStyle,
  isSynProjHeaderGreyCol,
  synProjHeaderGreyStyle,
  isSynProjHeaderRedCol,
  synProjHeaderRedStyle,
  isSynRow19MoGreenCol,
  isSynRow19PaaRedCol,
  synRow19MoGreenStyle,
  synRow19PaaRedStyle,
  isSynRow20PortfolioRedCol,
  isSynRow20PortfolioYellowCol,
  synRow20PortfolioRedStyle,
  synRow20PortfolioYellowStyle,
  isSynRow21Step3Col,
  isSynRow17BlueEvery3Col,
  synRow17MaaBlueStyle,
  isSynRow18GreyCol,
  synRow18MaaGreyStyle,
  isSynRow25MaGreenCol,
  synRow25MaGreenStyle,
  isSynRow16FluoEvery3Col,
  isSynApbbRow16FluoCol,
  isSynApbbRow17BlueCol,
  isSynApbbRow18GreyCol,
  isSynApbbP3sBlackCol,
  synApbbP3sBlackStyle,
  isSynRow17FluoEvery3FromMCol,
  synRow16FluoStyle,
  SYN_DISPLAY_GREEN_ROWS,
  SYN_DISPLAY_GREEN_BG,
  isSynDisplayRowGreyMaaCol,
  synDisplayRowGreyMaaStyle,
  isSynDisplayRowGreenMaaCol,
  synDisplayRowGreenMaaStyle,
  isSynDisplayRowGreenAcanCol,
  synDisplayRowGreenAcanStyle,
  isSynYellowFluoGreenFromMCol,
  synYellowFluoGreenFromMStyle,
  isSynHdrLmDividerRightCol,
  isSynHdrLmDividerLeftCol,
  isSynHdrAaDividerRightCol,
  isSynHdrCjDividerRightEntry,
  isSynHdrMaDividerRightEntry,
  isSynHdrLmDividerRightEntry,
  isSynHdrLmDividerLeftEntry,
  isSynHdrAaDividerRightEntry,
  isSynAcAnTableCellEntry,
  isSynHdrAcAnDividerRightEntry,
  isSynHdrAcAnDividerLeftEntry,
  isSynHdrAcAnDividerRightEdgeEntry,
  isSynApBbTableCellEntry,
  isSynHdrApBbDividerRightEntry,
  isSynHdrApBbDividerLeftEntry,
  isSynHdrApBbDividerRightEdgeEntry,
  isSynBsCeTableCellEntry,
  isSynHdrBsCeDividerRightEntry,
  isSynHdrBsCeDividerLeftEntry,
  isSynHdrBsCeDividerRightEdgeEntry,
  isSynBdBoTableCellEntry,
  isSynHdrBdBoDividerRightEntry,
  isSynHdrBdBoDividerLeftEntry,
  isSynHdrBdBoDividerRightEdgeEntry,
  isSynCiCyTableCellEntry,
  isSynHdrCiCyDividerRightEntry,
  isSynHdrCiCyDividerLeftEntry,
  isSynHdrCiCyDividerRightEdgeEntry,
  isSynDaDpTableCellEntry,
  isSynHdrDaDpDividerRightEntry,
  isSynHdrDaDpDividerLeftEntry,
  isSynHdrDaDpDividerRightEdgeEntry,
  isSynDrEdTableCellEntry,
  isSynHdrDrEdDividerRightEntry,
  isSynHdrDrEdDividerLeftEntry,
  isSynHdrDrEdDividerRightEdgeEntry,
  isSynEfEqTableCellEntry,
  isSynHdrEfEqDividerRightEntry,
  isSynHdrEfEqDividerLeftEntry,
  isSynHdrEfEqDividerRightEdgeEntry,
  isSynEsFeTableCellEntry,
  isSynHdrEsFeDividerRightEntry,
  isSynHdrEsFeDividerLeftEntry,
  isSynHdrEsFeDividerRightEdgeEntry,
  isSynFjFzTableCellEntry,
  isSynHdrFjFzDividerRightEntry,
  isSynHdrFjFzDividerLeftEntry,
  isSynHdrFjFzDividerRightEdgeEntry,
  SYN_HEADER_PANEL_LAST_ROW,
  isSynSpacerDisplayExcelCol,
  isSynForceWhiteExcelCol,
  synSpacerColClass,
  SYN_GRID_FIRST_ROW,
  isSynNumericEntryCell,
  SYN_BUILTIN_PILLAR_META,
  SYN_SP2_RESTART_BG,
  isSynSp2RestartDisplayExcelCol,
  synLabel,
  getSynAdaptBandNumeric,
} from './synStore.js?v=grid-perf2';
import {
  SYN_STICKY_COL,
  excelToDisplayCol,
  synStickyColWidth,
  synPillarColWidth,
  isSynBuiltinPillarExcelCol,
} from './synthesisPerf.js?v=syn-pillar-cg2';
import {
  ROW_H,
  synColOverscanPx,
  rowOverscan,
  shouldVirtualizeCols,
  shouldVirtualizeRows,
  createScrollRafSync,
  createColScrollCache,
  createRowScrollCache,
  SYN_MAX_RENDERED_COLS,
  SYN_MAX_RENDERED_ROWS,
} from './gridScroll.js?v=syn-fix2';
const SYN_HEAD_ROW_H = 22;
const ROW_NUM_W = 56;

export default {
  name: 'SynthesisGrid',
  props: {
    sheet: { type: Object, required: true },
    session: { type: Object, default: null },
    rawSyn: { type: Object, default: null },
    outlineOnly: { type: Boolean, default: false },
    paneVisible: { type: Boolean, default: true },
    externalEditTick: { type: Number, default: 0 },
    searchCmd: { type: Object, default: null },
  },
  emits: ['cell-change', 'cell-select', 'search-row-hidden', 'search-navigated'],
  setup(props, { emit }) {
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const scrollLeft = ref(0);
    const viewportH = ref(600);
    const viewportW = ref(1000);
    const cellMap = shallowRef((props.sheet && props.sheet.cellMap) || new Map());

    watch(
      () => props.sheet,
      (sheet) => {
        if (sheet && sheet.cellMap instanceof Map) {
          cellMap.value = sheet.cellMap;
        } else if (sheet) {
          cellMap.value = buildCellMap(sheet.cells, sheet.headerRows);
        }
      },
      { immediate: true }
    );

    const pillarColumns = computed(() => {
      const preset = props.sheet && props.sheet.pillarColumns;
      const map =
        preset && typeof preset === 'object'
          ? new Map(Object.entries(preset))
          : buildSynPillarColumns(props.sheet, cellMap.value);
      for (const [col, meta] of Object.entries(SYN_BUILTIN_PILLAR_META)) {
        map.set(col, { ...meta, ...map.get(col) });
      }
      return map;
    });

    const displayColumns = computed(() => props.sheet.columns || []);

    const widthByCol = computed(() => {
      const m = new Map();
      const labelW = synStickyColWidth(props.sheet);
      for (const w of props.sheet.colWidths || []) {
        if (w.col === SYN_STICKY_COL) continue;
        const pillarW = synPillarColWidth(w.col, props.sheet, pillarColumns.value);
        const wPx =
          pillarW != null ? pillarW : Math.min(w.width || 64, 100);
        m.set(w.col, wPx);
      }
      for (const col of displayColumns.value) {
        if (!m.has(col) && col !== SYN_STICKY_COL) {
          m.set(
            col,
            synPillarColWidth(col, props.sheet, pillarColumns.value) != null
              ? synPillarColWidth(col, props.sheet, pillarColumns.value)
              : 54
          );
        }
      }
      m.set(SYN_STICKY_COL, labelW);
      return m;
    });

    const columnLayout = computed(() => {
      let left = ROW_NUM_W;
      return displayColumns.value.map((col) => {
        const width = widthByCol.value.get(col) || 54;
        const entry = {
          col,
          letter: excelToDisplayCol(col),
          left,
          width,
        };
        left += width;
        return entry;
      });
    });

    const tableWidth = computed(() => {
      const layout = columnLayout.value;
      if (!layout.length) return ROW_NUM_W;
      const last = layout[layout.length - 1];
      return last.left + last.width;
    });

    const pinnedCols = computed(() =>
      columnLayout.value.filter((c) => c.col === SYN_STICKY_COL)
    );

    const scrollableCols = computed(() =>
      columnLayout.value.filter((c) => c.col !== SYN_STICKY_COL)
    );

    const stickyLabelLeft = computed(() => ROW_NUM_W);

    const virtualizeCols = computed(() =>
      shouldVirtualizeCols(tableWidth.value, viewportW.value)
    );
    const colBufferPx = computed(() => synColOverscanPx(viewportW.value));
    /** Column window — viewport-first (monotonic off avoids 60k+ DOM nodes). */
    const colScrollCache = createColScrollCache(SYN_MAX_RENDERED_COLS, {
      monotonic: false,
    });
    const colRangeStart = ref(0);
    const colRangeEnd = ref(0);

    watchEffect(() => {
      if (!props.paneVisible) return;
      const cols = scrollableCols.value;
      if (!virtualizeCols.value) {
        colRangeStart.value = 0;
        colRangeEnd.value = cols.length;
        return;
      }
      const range = colScrollCache.resolve(
        cols,
        scrollLeft.value,
        viewportW.value,
        colBufferPx.value
      );
      colRangeStart.value = range.start;
      colRangeEnd.value = range.end;
    });

    const visibleScrollCols = computed(() =>
      scrollableCols.value.slice(colRangeStart.value, colRangeEnd.value)
    );

    const scrollColsLen = computed(() => visibleScrollCols.value.length);

    const leftPad = computed(() => {
      if (!virtualizeCols.value) return 0;
      const first = visibleScrollCols.value[0];
      const pin = pinnedCols.value[0];
      if (!first || !pin) return 0;
      return Math.max(0, first.left - (pin.left + pin.width));
    });

    const rightPad = computed(() => {
      if (!virtualizeCols.value) return 0;
      const last = visibleScrollCols.value[visibleScrollCols.value.length - 1];
      if (!last) return 0;
      return Math.max(0, tableWidth.value - (last.left + last.width));
    });

    const bodyRows = computed(() => {
      void props.outlineOnly;
      void (props.sheet && props.sheet.effectiveLastRow);
      return computeSynBodyRows(props.sheet, cellMap.value, props.outlineOnly);
    });
    const rowCount = computed(() => bodyRows.value.length);

    const displayRowByExcel = computed(() => {
      const m = new Map();
      for (const e of bodyRows.value) {
        if (e.excelRow != null) m.set(e.excelRow, e.displayRow);
      }
      return m;
    });

    const excelRowBodyIndex = computed(() => {
      const m = new Map();
      const rows = bodyRows.value;
      for (let i = 0; i < rows.length; i++) {
        if (rows[i].excelRow != null) m.set(rows[i].excelRow, i);
      }
      return m;
    });

    const rowScrollCache = createRowScrollCache(SYN_MAX_RENDERED_ROWS, {
      monotonic: false,
    });
    const virtualizeRows = computed(() =>
      shouldVirtualizeRows(rowCount.value, viewportH.value)
    );
    const rowOverscanPx = computed(() => rowOverscan(viewportH.value));
    const visibleRowStart = ref(0);
    const visibleRowEnd = ref(0);

    watchEffect(() => {
      if (!props.paneVisible) return;
      const count = rowCount.value;
      if (!virtualizeRows.value) {
        visibleRowStart.value = 0;
        visibleRowEnd.value = count;
        return;
      }
      const range = rowScrollCache.resolve(
        scrollTop.value,
        viewportH.value,
        count,
        rowOverscanPx.value
      );
      visibleRowStart.value = range.start;
      visibleRowEnd.value = range.end;
    });

    const visibleRows = computed(() =>
      bodyRows.value.slice(visibleRowStart.value, visibleRowEnd.value)
    );
    const topSpacer = computed(() =>
      virtualizeRows.value ? visibleRowStart.value * ROW_H : 0
    );
    const bottomSpacer = computed(() => {
      if (!virtualizeRows.value) return 0;
      const shown = visibleRowEnd.value - visibleRowStart.value;
      const remaining = rowCount.value - visibleRowStart.value - shown;
      return Math.max(0, remaining * ROW_H);
    });

    function bodyTopForExcelRow(excelRow) {
      const i = excelRowBodyIndex.value.get(excelRow);
      if (i == null) return null;
      return SYN_HEAD_ROW_H + i * ROW_H;
    }

    function usesPillarLetterOverlay(col) {
      return (
        SYN_PILLAR_OVERLAY_COLS.has(col) &&
        (pillarColumns.value.has(col) || isSynBuiltinPillarExcelCol(col))
      );
    }

    function synExcelColTraceClass(col) {
      return col === 'CL' ? 'syn-col-cl' : '';
    }

    const pillarLetterOverlays = computed(() => {
      const start = findSynEchappementRow(cellMap.value, props.sheet);
      const visibleCols = new Set([
        ...visibleScrollCols.value.map((c) => c.col),
        SYN_STICKY_COL,
      ]);
      const overlays = [];
      for (const layout of columnLayout.value) {
        if (!visibleCols.has(layout.col)) continue;
        if (!SYN_PILLAR_OVERLAY_COLS.has(layout.col)) continue;
        if (
          !pillarColumns.value.has(layout.col) &&
          !isSynBuiltinPillarExcelCol(layout.col)
        ) {
          continue;
        }
        const pillarMeta = pillarColumns.value.get(layout.col);
        const builtin = SYN_BUILTIN_PILLAR_META[layout.col];
        const title =
          (pillarMeta && pillarMeta.title != null ? pillarMeta.title : null) != null
            ? pillarMeta.title
            : (builtin && builtin.title != null ? builtin.title : '');
        const letters = synPillarLettersFromTitle(title);
        const chars = [];
        for (let i = 0; i < letters.length; i++) {
          const excelRow = start + i * SYN_PILLAR_LETTER_ROW_STEP;
          const top = bodyTopForExcelRow(excelRow);
          if (top == null) continue;
          const ch = letters[i];
          chars.push({
            ch: ch === ' ' ? '\u00a0' : ch,
            top,
          });
        }
        overlays.push({
          col: layout.col,
          left: layout.left,
          width: layout.width,
          chars,
        });
      }
      return overlays;
    });

    const colspan = computed(() => {
      let n = 1 + pinnedCols.value.length + visibleScrollCols.value.length;
      if (leftPad.value > 0) n += 1;
      if (rightPad.value > 0) n += 1;
      return n;
    });

    function colStyle(col, w) {
      const width =
        w != null
          ? w
          : widthByCol.value.has(col)
            ? widthByCol.value.get(col)
            : 54;
      return { width: `${width}px`, minWidth: `${width}px`, maxWidth: `${width}px` };
    }

    function stickyStyle(col) {
      if (col !== SYN_STICKY_COL) return {};
      return { left: `${stickyLabelLeft.value}px` };
    }

    const scrollSync = createScrollRafSync({
      scrollTop,
      scrollLeft,
      getScrollEl: () => scrollEl.value,
    });

    const editEpoch = ref(0);
    /** Defer SUMPRODUCT until after first paint. */
    const liveCalcReady = ref(false);
    /** Static values while scrolling — live calc refreshes when scroll stops. */
    const scrollActive = ref(false);
    let scrollEndTimer = null;

    function onGridScroll(e) {
      scrollSync.onScroll(e);
      if (!scrollActive.value) scrollActive.value = true;
      if (scrollEndTimer) clearTimeout(scrollEndTimer);
      scrollEndTimer = setTimeout(() => {
        scrollEndTimer = null;
        scrollActive.value = false;
        editEpoch.value += 1;
      }, 80);
    }
    const selectedCell = ref(null);
    const calcRevision = computed(() =>
      props.session && props.session.revision && props.session.revision.value != null
        ? props.session.revision.value
        : 0
    );
    const synCalcTick = computed(() =>
      props.session && props.session.synCalcTick && props.session.synCalcTick.value != null
        ? props.session.synCalcTick.value
        : 0
    );
    const displayCache = new Map();
    const scrollStyleCache = new Map();
    const rowClassesCache = new Map();
    let displayCacheKey = '';

    function invalidateDisplayCache() {
      displayCache.clear();
      scrollStyleCache.clear();
      rowClassesCache.clear();
      displayCacheKey = '';
    }

    function cachedEntryRowClasses(entry) {
      const key = isGapEntry(entry)
        ? `gap-${entry.gapKey || ((entry.gapBeforePanel ? 'top-' : 'bot-') + entry.gapIndex)}`
        : String(entry.excelRow);
      if (rowClassesCache.has(key)) return rowClassesCache.get(key);
      const cls = entryRowClasses(entry);
      rowClassesCache.set(key, cls);
      return cls;
    }

    watch(
      () => props.externalEditTick,
      () => {
        if (cellEditor && typeof cellEditor.clearIdleDisplays === 'function') {
          cellEditor.clearIdleDisplays();
        }
        editEpoch.value += 1;
        invalidateDisplayCache();
      }
    );

    watch(
      () =>
        props.session && props.session.synCalcTick && props.session.synCalcTick.value != null
          ? props.session.synCalcTick.value
          : 0,
      () => {
        scrollActive.value = false;
        editEpoch.value += 1;
        invalidateDisplayCache();
      }
    );

    watch(
      () =>
        props.session && props.session.ready && props.session.ready.value
          ? props.session.ready.value
          : false,
      (ready) => {
        if (!ready || liveCalcReady.value) return;
        const enable = () => {
          // Avoid kicking off heavy SUMPRODUCT while the user is actively scrolling.
          if (!scrollActive.value) liveCalcReady.value = true;
        };
        if (typeof requestIdleCallback === 'function') {
          requestIdleCallback(enable, { timeout: 1200 });
        } else {
          setTimeout(enable, 300);
        }
      },
      { immediate: true }
    );

    watch(
      () =>
        props.session && props.session.displayTick && props.session.displayTick.value != null
          ? props.session.displayTick.value
          : 0,
      () => {
        const dirty =
          props.session && props.session.takeDisplayDirty
            ? props.session.takeDisplayDirty()
            : null;
        if (!dirty || dirty.all) {
          editEpoch.value += 1;
          invalidateDisplayCache();
          return;
        }
        if (dirty.keys && dirty.keys.size) {
          for (const k of dirty.keys) {
            displayCache.delete(k);
            const [row, col] = k.split(':');
            const needle = `:${row}:${col}:`;
            for (const sk of scrollStyleCache.keys()) {
              if (sk.includes(needle)) scrollStyleCache.delete(sk);
            }
          }
        }
      }
    );

    watch(
      () => Boolean(props.session && props.session.ready && props.session.ready.value),
      (ready, wasReady) => {
        if (ready && !wasReady) {
          editEpoch.value += 1;
          invalidateDisplayCache();
        }
      }
    );

    function invalidateAdaptationSumCol(col) {
      const sumRow = synAdaptationSumRow(props.sheet);
      displayCache.delete(`${sumRow}:${col}`);
      const needle = `:${sumRow}:${col}:`;
      for (const sk of scrollStyleCache.keys()) {
        if (sk.includes(needle)) scrollStyleCache.delete(sk);
      }
      bumpCellModelCache();
    }

    /** SUM(ADAPTATION) cols C..J — local, instant (rows adaptationSumFrom..To on sheet). */
    function adaptationSumLocal(col) {
      const sheet = props.sheet;
      const from =
        sheet && sheet.adaptationSumFromRow != null ? sheet.adaptationSumFromRow : 27;
      const to =
        sheet && sheet.adaptationSumToRow != null ? sheet.adaptationSumToRow : 40;
      const map = cellMap.value;
      const getAt = (r, c) => getCell(map, r, c);
      const n = computeAdaptationRowSum(
        (r, c) => getSynAdaptBandNumeric(getAt, r, c),
        col,
        from,
        to
      );
      return formatVal(String(n));
    }

    function previewAdaptationBandInput(row, col, raw) {
      const key = `${row}:${col}`;
      let cell = cellMap.value.get(key);
      if (!cell) {
        cell = { r: row, c: col, v: raw, userEdited: true };
        props.sheet.cells.push(cell);
        cellMap.value.set(key, cell);
      } else {
        cell.v = raw;
        cell.userEdited = true;
      }
      invalidateAdaptationSumCol(col);
      editEpoch.value += 1;
    }

    function commitCellInput(row, col, value, previousValue) {
      const key = `${row}:${col}`;
      let cell = cellMap.value.get(key);
      if (!cell) {
        cell = { r: row, c: col, v: value, userEdited: true };
        props.sheet.cells.push(cell);
        cellMap.value.set(key, cell);
      } else {
        cell.v = value;
        cell.userEdited = true;
        delete cell.f;
      }
      if (props.rawSyn) upsertRawCell(props.rawSyn, row, col, value);
      if (affectsAdaptationSum(row, col, props.sheet)) {
        invalidateAdaptationSumCol(col);
        editEpoch.value += 1;
      } else {
        editEpoch.value += 1;
        invalidateDisplayCache();
      }
      emit('cell-change', {
        row,
        col,
        value,
        sheet: 'SYNTHESIS',
        previousValue: previousValue != null ? previousValue : '',
      });
      if (props.session && props.session.setCellValue) {
        void props.session
          .setCellValue('SYNTHESIS', row, col, value)
          .catch((e) => console.warn('Synthesis setCellValue:', e));
      }
    }

    const cellEditor = createGridCellEditor({
      isNumericAt: (row, col) =>
        isSynNumericEntryCell(row, col, pillarColumns.value),
      displayAt: (row, col) => cellDisplay(row, col),
      commitAt: commitCellInput,
    });
    const {
      isCellActive,
      cellShowValue,
      onCellMouseDown,
      onCellSpanMouseDown: onCellSpanMouseDownBase,
      onCellFocus: onCellFocusBase,
      onCellInput: onCellInputBase,
      onCellBlur,
      onCellKeydown: onCellKeydownBase,
      prepareNavigate,
      beginNavigationTo,
      activateCell,
      setNavigationLock,
    } = cellEditor;

    function onCellSpanMouseDown(row, col, event) {
      onReadonlyCellSelect(row, col, event);
      onCellSpanMouseDownBase(row, col, event);
    }

    function onCellFocus(row, col, event) {
      selectCell(row, col);
      onCellFocusBase(row, col, event);
    }

    /**
     * Only the ADAPTATION total row (=SOMME(...)) stays locked. Every other cell —
     * including blue SUMPRODUCT, green section sums and the AB diff — is editable in
     * real time; typing a value overrides the live formula (userEdited path).
     */
    function cellReadonly(row, col) {
      if (row == null) return true;
      return isSynAdaptationSumCell(row, col, props.sheet);
    }

    const cellNavigation = createGridCellNavigation({
      getColumns: () => displayColumns.value,
      getRows: () =>
        bodyRows.value
          .filter((e) => !isSynPanelGapEntry(e))
          .map((e) => e.excelRow)
          .filter((r) => r != null),
      isNavigable: (row, col) => !cellReadonly(row, col),
      getScrollEl: () => scrollEl.value,
      flushScroll: () => scrollSync.flush(),
      getRowTop: (_rowIndex, excelRow) => bodyTopForExcelRow(excelRow),
      getColLeft: (col) => {
        const entry = columnLayout.value.find((c) => c.col === col);
        return entry && entry.left != null ? entry.left : null;
      },
      getColWidth: (col) => widthByCol.value.get(col) || 54,
      rowHeight: ROW_H,
    });
    const onCellKeydown = cellNavigation.wrapKeydown(
      onCellKeydownBase,
      prepareNavigate,
      beginNavigationTo,
      activateCell,
      setNavigationLock
    );

    function onCellInputLive(row, col, event) {
      onCellInputBase(row, col, event);
      if (!affectsAdaptationSum(row, col, props.sheet)) return;
      const el = event.target;
      if (!el) return;
      previewAdaptationBandInput(row, col, el.value);
    }

    function selectCell(row, col) {
      if (row == null || !col) return;
      selectedCell.value = { row, col };
      emit('cell-select', { row, col });
    }

    const formulaText = computed(() => {
      void calcRevision.value;
      const sel = selectedCell.value;
      if (!sel) return '';
      if (props.session && props.session.getSynFormula) {
        const v = props.session.getSynFormula(sel.row, sel.col);
        return v != null ? v : '';
      }
      return '';
    });

    const formulaCellRef = computed(() => {
      const sel = selectedCell.value;
      if (!sel) return '';
      return `${excelToDisplayCol(sel.col)}${sel.row}`;
    });

    function onReadonlyCellSelect(row, col, event) {
      if (event.button !== 0) return;
      selectCell(row, col);
    }

    function formatVal(v) {
      const formatted = formatSynNumericDisplay(v);
      return formatted === '' && v != null && String(v).trim() !== ''
        ? String(v)
        : formatted;
    }

    function cellDisplayStatic(row, col, cell, label, rowClass) {
      if (isSynAdaptationSumCell(row, col, props.sheet)) {
        return adaptationSumLocal(col);
      }
      const shown = synDisplayValue(
        cell,
        cellMap.value,
        row,
        col,
        props.sheet,
        pillarColumns.value
      );
      return isSynMetricRow(row) ? shown : formatVal(shown);
    }

    const adaptationSumTick = computed(() =>
      props.session &&
      props.session.adaptationSumTick &&
      props.session.adaptationSumTick.value != null
        ? props.session.adaptationSumTick.value
        : 0
    );

    function cellDisplay(row, col) {
      const cacheKey = `${synCalcTick.value}:${adaptationSumTick.value}:${editEpoch.value}:${liveCalcReady.value ? 1 : 0}`;
      if (cacheKey !== displayCacheKey) {
        invalidateDisplayCache();
        displayCacheKey = cacheKey;
      }
      if (row == null) return '';
      const hitKey = `${row}:${col}`;
      if (displayCache.has(hitKey)) return displayCache.get(hitKey);
      // ADAPTATION total row — SUM(rows adaptationSumFrom..adaptationSumTo), cols C..J.
      if (isSynAdaptationSumCell(row, col, props.sheet)) {
        const out = adaptationSumLocal(col);
        displayCache.set(hitKey, out);
        return out;
      }
      if (isSynPillarColAtRow(col, row, pillarColumns.value)) {
        if (usesPillarLetterOverlay(col)) {
          displayCache.set(hitKey, '');
          return '';
        }
        const pillar = synPillarLetterForRow(
          row,
          col,
          pillarColumns.value,
          cellMap.value,
          props.sheet
        );
        displayCache.set(hitKey, pillar);
        return pillar;
      }
      const cell = getCell(cellMap.value, row, col);
      const label = synLabel(cellMap.value, row);
      const rowClass = synRowStyleClass(cellMap.value, row, props.sheet);
      const isLiveCalc = isSynCalculatedMassCell(
        cell,
        row,
        col,
        props.sheet,
        label,
        rowClass
      );
      const materialized =
        (props.sheet && props.sheet.materializedCalc) ||
        (cell &&
          cell.mat &&
          !cell.userEdited &&
          !isLiveCalc);

      if (
        isLiveCalc &&
        !materialized &&
        liveCalcReady.value &&
        !scrollActive.value &&
        props.session &&
        props.session.getDisplayValue &&
        props.session.ready &&
        props.session.ready.value
      ) {
        const fromSession = props.session.getDisplayValue(
          'SYNTHESIS',
          row,
          col,
          cell
        );
        if (fromSession != null && !(cell && cell.userEdited)) {
          const out = isSynMetricRow(row)
            ? fromSession
            : formatVal(fromSession);
          displayCache.set(hitKey, out);
          return out;
        }
      }

      const out = cellDisplayStatic(row, col, cell, label, rowClass);
      displayCache.set(hitKey, out);
      return out;
    }

    let searchIndex = null;
    let searchIndexEpoch = -1;
    let searchIndexPromise = null;

    function getSearchIndexForGrid() {
      const epoch = editEpoch.value;
      if (searchIndex && searchIndexEpoch === epoch) {
        return Promise.resolve(searchIndex);
      }
      if (searchIndexPromise && searchIndexEpoch === epoch) {
        return searchIndexPromise;
      }
      searchIndexEpoch = epoch;
      searchIndex = null;
      searchIndexPromise = buildSearchIndex(
        cellMap.value,
        props.sheet,
        displayColumns.value
      ).then((idx) => {
        searchIndex = idx;
        searchIndexPromise = null;
        return idx;
      });
      return searchIndexPromise;
    }

    const gridSearch = createGridSearchController({
      getSearchIndex: getSearchIndexForGrid,
      getBodyExcelRows: () =>
        bodyRows.value
          .filter((e) => !isSynPanelGapEntry(e))
          .map((e) => e.excelRow)
          .filter((r) => r != null),
      getScrollEl: () => scrollEl.value,
      scrollCellIntoView: (row, col) =>
        cellNavigation.scrollCellIntoViewForced(row, col),
      flushScroll: () => scrollSync.flush(),
      getRowTop: (_rowIndex, excelRow) => bodyTopForExcelRow(excelRow),
      getViewportH: () => viewportH.value,
      rowHeight: ROW_H,
      onRowHidden: (match) => emit('search-row-hidden', match),
    });

    watch(
      () => props.searchCmd,
      (cmd) => {
        if (!props.paneVisible || !cmd) return;
        if (!cmd.q) {
          gridSearch.clearSearch();
          searchIndex = null;
          searchIndexEpoch = -1;
          searchIndexPromise = null;
          emit('search-navigated', { count: 0, index: 0 });
          return;
        }
        if (cmd.step === 0) emit('search-navigated', { searching: true });
        void gridSearch
          .searchAndScroll(cmd.q, { step: cmd.step != null ? cmd.step : 0 })
          .then((result) => {
            if (!result || !result.cancelled) emit('search-navigated', result);
          });
      }
    );

    function headerEdgeRight(row, colIdx, colsLen) {
      return isSynHeaderPanelRow(row) && colsLen > 0 && colIdx === colsLen - 1;
    }

    function headerLabelEdgeRight(row) {
      return isSynHeaderPanelRow(row) && visibleScrollCols.value.length === 0;
    }

    function synHeaderTableFrameRowClasses() {
      return [
        'syn-header-block',
        'syn-proj-table-frame',
        'syn-hdr-panel-grid',
        'syn-ac-an-table-frame',
        'syn-ap-bb-table-frame',
        'syn-bs-ce-table-frame',
        'syn-bd-bo-table-frame',
        'syn-ci-cy-table-frame',
        'syn-da-dp-table-frame',
        'syn-dr-ed-table-frame',
        'syn-ef-eq-table-frame',
        'syn-es-fe-table-frame',
        'syn-fj-fz-table-frame',
      ];
    }

    function entryRowClasses(entry) {
      if (isSynPanelGapEntry(entry)) {
        const gap = ['syn-panel-gap-row', 'syn-panel-gap'];
        if (entry.gapBetween) {
          gap.push('syn-header-spacer-row', 'syn-header-spacer-white', 'syn-header-table-gap-row');
          gap.push(...synHeaderTableFrameRowClasses());
        }
        if (entry.gapBeforePanel) gap.push('syn-panel-gap-top');
        if (entry.gapAfterPanel && entry.gapIndex === 1) gap.push('syn-panel-gap-first');
        if (entry.gapAfterPanel && entry.gapIndex === 2) gap.push('syn-panel-gap-last');
        return gap;
      }
      const row = entry.excelRow;
      const cls = synRowStyleClass(cellMap.value, row, props.sheet);
      const list = [cls];
      if (isSynHeaderPanelRow(row)) {
        list.push(...synHeaderTableFrameRowClasses());
        if (row === SYN_GRID_FIRST_ROW) {
          list.push('syn-ac-an-edge-top');
          list.push('syn-ap-bb-edge-top');
          list.push('syn-bs-ce-edge-top');
          list.push('syn-bd-bo-edge-top');
          list.push('syn-ci-cy-edge-top');
          list.push('syn-da-dp-edge-top');
          list.push('syn-dr-ed-edge-top');
          list.push('syn-ef-eq-edge-top');
          list.push('syn-es-fe-edge-top');
          list.push('syn-fj-fz-edge-top');
        }
        if (row === SYN_HEADER_PANEL_LAST_ROW) {
          list.push('syn-ac-an-edge-bottom');
          list.push('syn-ap-bb-edge-bottom');
          list.push('syn-bs-ce-edge-bottom');
          list.push('syn-bd-bo-edge-bottom');
          list.push('syn-ci-cy-edge-bottom');
          list.push('syn-da-dp-edge-bottom');
          list.push('syn-dr-ed-edge-bottom');
          list.push('syn-ef-eq-edge-bottom');
          list.push('syn-es-fe-edge-bottom');
          list.push('syn-fj-fz-edge-bottom');
        }
      }
      if (row >= 3 && row <= 14) list.push('syn-filter-band');
      if (row >= 15 && row <= 22) list.push('syn-metric-band');
      if (row === 16 || row === 18) list.push('syn-metric-curb');
      if (row === 10 || row === 15) list.push('syn-header-spacer-row', 'syn-header-spacer-white');
      if (row === 10) list.push('syn-hdr-row-10-label');
      if (row === 9) list.push('syn-header-edge-below-spec');
      if (row === 11) list.push('syn-header-edge-above-pole');
      if (row === 14) list.push('syn-header-edge-below-finition');
      if (row === 16) list.push('syn-header-edge-above-curb');
      if (row === 22) list.push('syn-header-panel-end');
      if (row === SYN_GRID_FIRST_ROW) list.push('syn-header-edge-above-date');
      if (row >= 3 && row <= 8) list.push('syn-header-edge-sep');
      if (row === 11 || row === 12 || row === 13) list.push('syn-header-edge-sep');
      if (row >= 16 && row <= 21) list.push('syn-header-edge-sep');
      return list;
    }

    function isGapEntry(entry) {
      return isSynPanelGapEntry(entry);
    }

    function isPillarColForEntry(entry, col) {
      if (isGapEntry(entry)) return false;
      const row = entry.excelRow;
      if (row >= SYN_GRID_FIRST_ROW && isSynBuiltinPillarExcelCol(col)) return true;
      return isSynPillarColAtRow(col, row, pillarColumns.value);
    }

    /** Gap rows 21–22: display B/K/CG pillars keep pillar fill. */
    function isGapGreenPillarCol(col) {
      const letter = excelToDisplayCol(col);
      return letter === 'B' || letter === 'K' || letter === 'CG';
    }

    function pillarTitle(col) {
      const meta = pillarColumns.value.get(col);
      return meta && meta.title != null ? meta.title : '';
    }

    function cellInlineStyle(row, col) {
      const cell = getCell(cellMap.value, row, col);
      return synCellInlineStyle(
        cell,
        cellMap.value,
        row,
        col,
        props.sheet,
        pillarColumns.value
      );
    }

    function scrollDataCellStyle(entry, col, width) {
      const base = colStyle(col, width);
      if (isGapEntry(entry)) return base;
      const row = entry.excelRow;
      const styleKey = `${displayCacheKey}:${row}:${col}:${width}`;
      if (scrollStyleCache.has(styleKey)) {
        return scrollStyleCache.get(styleKey);
      }
      if (isSynForceWhiteExcelCol(col)) return base;
      if (isSynSpacerDisplayExcelCol(col)) return base;
      // Body rows 26+ (non header panel): skip heavy header style chain.
      if (row >= SYN_GRID_FIRST_ROW && !isSynHeaderPanelRow(row)) {
        const out = { ...base, ...cellInlineStyle(row, col) };
        if (isSynDisplayRowGreyMaaCol(entry.displayRow, col)) {
          Object.assign(out, synDisplayRowGreyMaaStyle());
        } else if (isSynDisplayRowGreenMaaCol(entry.displayRow, col)) {
          Object.assign(out, synDisplayRowGreenMaaStyle());
        } else if (isSynDisplayRowGreenAcanCol(entry.displayRow, col)) {
          Object.assign(out, synDisplayRowGreenAcanStyle());
        }
        scrollStyleCache.set(styleKey, out);
        return out;
      }
      const display = cellDisplay(row, col);
      const bevStyle = synHdrBevTextStyle(row, display);
      if (bevStyle) {
        return { ...base, ...bevStyle };
      }
      const energyStyle = synHdrEnergyValueStyle(row, col, display);
      if (energyStyle) {
        return { ...base, ...energyStyle };
      }
      if (row >= SYN_GRID_FIRST_ROW && isSynSp2RestartDisplayExcelCol(col)) {
        return {
          ...base,
          background: SYN_SP2_RESTART_BG,
          backgroundColor: SYN_SP2_RESTART_BG,
          color: '#000',
        };
      }
      if (
        isSynStlaSlashGreenCell(row, col, display) ||
        isSynProjHeaderGreenCol(row, col) ||
        isSynTargetTextGreenCell(col, display)
      ) {
        return { ...base, ...synProjHeaderGreenStyle() };
      }
      if (isSynApbbP3sBlackCol(row, col)) {
        return { ...base, ...synApbbP3sBlackStyle() };
      }
      if (isSynProjHeaderYellowCol(row, col)) {
        return { ...base, ...synProjHeaderYellowStyle() };
      }
      if (isSynProjHeaderGreyCol(row, col)) {
        return { ...base, ...synProjHeaderGreyStyle() };
      }
      if (isSynProjHeaderRedCol(row, col)) {
        return { ...base, ...synProjHeaderRedStyle() };
      }
      if (isSynRow19MoGreenCol(row, col)) {
        return { ...base, ...synRow19MoGreenStyle() };
      }
      if (isSynRow19PaaRedCol(row, col)) {
        return { ...base, ...synRow19PaaRedStyle() };
      }
      if (isSynRow20PortfolioRedCol(row, col)) {
        const style = { ...base, ...synRow20PortfolioRedStyle() };
        if (isSynRow20PortfolioYellowCol(row, col)) {
          Object.assign(style, synRow20PortfolioYellowStyle());
        }
        return style;
      }
      if (isSynRow17BlueEvery3Col(row, col)) {
        return { ...base, ...synRow17MaaBlueStyle() };
      }
      if (isSynApbbRow17BlueCol(row, col)) {
        return { ...base, ...synRow17MaaBlueStyle() };
      }
      if (isSynRow18GreyCol(row, col)) {
        return { ...base, ...synRow18MaaGreyStyle() };
      }
      if (isSynApbbRow18GreyCol(row, col)) {
        return { ...base, ...synRow18MaaGreyStyle() };
      }
      if (isSynRow25MaGreenCol(row, col)) {
        return { ...base, ...synRow25MaGreenStyle() };
      }
      if (isSynYellowFluoGreenFromMCol(row, col, cellMap.value, props.sheet)) {
        return { ...base, ...synYellowFluoGreenFromMStyle() };
      }
      if (isSynRow16FluoEvery3Col(row, col)) {
        return { ...base, ...synRow16FluoStyle() };
      }
      if (isSynApbbRow16FluoCol(row, col)) {
        return { ...base, ...synRow16FluoStyle() };
      }
      if (isSynRow17FluoEvery3FromMCol(row, col)) {
        return { ...base, ...synRow16FluoStyle() };
      }
      if (isSynDisplayRowGreyMaaCol(entry.displayRow, col)) {
        return { ...base, ...synDisplayRowGreyMaaStyle() };
      }
      if (isSynDisplayRowGreenMaaCol(entry.displayRow, col)) {
        return { ...base, ...synDisplayRowGreenMaaStyle() };
      }
      if (isSynDisplayRowGreenAcanCol(entry.displayRow, col)) {
        return { ...base, ...synDisplayRowGreenAcanStyle() };
      }
      const out = { ...base, ...cellInlineStyle(row, col) };
      scrollStyleCache.set(styleKey, out);
      return out;
    }

    function withHdrPanelBold(row, col, cls, display = '') {
      const parts = [cls];
      if (isSynHeaderPanelBoldCol(row, col)) parts.push('syn-hdr-panel-bold');
      if (isSynAdaptCjValueBold(row, col, display, props.sheet)) {
        parts.push('syn-adapt-cj-bold');
      }
      return parts.filter(Boolean).join(' ');
    }

    function cellExtraClass(row, col, display) {
      if (
        isSynPillarColAtRow(col, row, pillarColumns.value) ||
        (row >= SYN_GRID_FIRST_ROW && isSynSp2RestartDisplayExcelCol(col))
      ) {
        return display ? 'syn-pillar-has-char' : '';
      }
      if (isSynForceWhiteExcelCol(col)) return 'syn-force-white-col';
      const spacerCol = synSpacerColClass(col);
      if (spacerCol) return spacerCol;
      const bevCls = synHdrBevTextClass(row, display);
      if (bevCls) return withHdrPanelBold(row, col, bevCls, display);
      const energyCls = synHdrEnergyValueClass(row, col, display);
      if (energyCls) return withHdrPanelBold(row, col, energyCls, display);
      // Display-row based styling (after spacer/pillar checks).
      const displayRow = displayRowByExcel.value.get(row);
      if (isSynDisplayRowGreyMaaCol(displayRow, col)) {
        return withHdrPanelBold(row, col, 'syn-displayrow-grey-maa', display);
      }
      if (isSynDisplayRowGreenMaaCol(displayRow, col)) {
        return withHdrPanelBold(row, col, 'syn-displayrow-green-maa', display);
      }
      if (isSynDisplayRowGreenAcanCol(displayRow, col)) {
        return withHdrPanelBold(row, col, 'syn-displayrow-green-acan', display);
      }
      const spotBlue = synSpotBlueColClass(row, col);
      if (spotBlue) return withHdrPanelBold(row, col, spotBlue, display);
      const greyCol = synFilterGreyColClass(row, col);
      if (greyCol) return withHdrPanelBold(row, col, greyCol, display);
      const adaptCol = synAdaptBandColClass(row, col, pillarColumns.value);
      if (adaptCol) return withHdrPanelBold(row, col, adaptCol, display);
      const metricWhiteCol = synMetricCjWhiteColClass(row, col);
      if (metricWhiteCol) return withHdrPanelBold(row, col, metricWhiteCol, display);
      const accent = synCellAccentClass(display);
      if (accent) return withHdrPanelBold(row, col, accent, display);
      if (
        isSynStlaSlashGreenCell(row, col, display) ||
        isSynProjHeaderGreenCol(row, col) ||
        isSynTargetTextGreenCell(col, display)
      ) {
        return withHdrPanelBold(row, col, 'syn-proj-hdr-green', display);
      }
      if (isSynApbbP3sBlackCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-apbb-p3s-black', display);
      }
      if (isSynProjHeaderYellowCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-proj-hdr-yellow', display);
      }
      if (isSynProjHeaderGreyCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-proj-hdr-grey', display);
      }
      if (isSynProjHeaderRedCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-proj-hdr-red', display);
      }
      if (isSynHeaderPanelRow(row)) {
        const hdrCls = synHeaderPanelVehicleClass(row, col, display);
        const combined = withHdrPanelBold(row, col, hdrCls, display);
        if (combined) return combined;
      }
      if (isSynRow19MoGreenCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row19-mo-green', display);
      }
      if (isSynRow19PaaRedCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row19-paa-red', display);
      }
      if (isSynRow20PortfolioRedCol(row, col)) {
        const cls = isSynRow20PortfolioYellowCol(row, col)
          ? 'syn-row20-portfolio-red syn-row20-portfolio-yellow'
          : 'syn-row20-portfolio-red';
        return withHdrPanelBold(row, col, cls, display);
      }
      if (isSynRow21Step3Col(row, col, display)) {
        return withHdrPanelBold(row, col, 'syn-row21-step3', display);
      }
      if (isSynRow17BlueEvery3Col(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row17-maa-blue', display);
      }
      if (isSynApbbRow17BlueCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row17-maa-blue', display);
      }
      if (isSynRow18GreyCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row18-maa-grey', display);
      }
      if (isSynApbbRow18GreyCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row18-maa-grey', display);
      }
      if (isSynRow25MaGreenCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row25-ma-green', display);
      }
      if (isSynYellowFluoGreenFromMCol(row, col, cellMap.value, props.sheet)) {
        return withHdrPanelBold(row, col, 'syn-fluo-green-from-m', display);
      }
      if (isSynRow16FluoEvery3Col(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row16-fluo-every3', display);
      }
      if (isSynApbbRow16FluoCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row16-fluo-every3', display);
      }
      if (isSynRow17FluoEvery3FromMCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row17-fluo-every3', display);
      }
      const rc = synRowStyleClass(cellMap.value, row, props.sheet);
      if (
        rc === 'syn-row-section' ||
        rc === 'syn-row-subsection' ||
        rc === 'syn-row-separator'
      ) {
        return withHdrPanelBold(row, col, '', display);
      }
      if (isSynMetricRow(row)) {
        return withHdrPanelBold(row, col, synMetricCellClass(row, col, display) || '', display);
      }
      return withHdrPanelBold(row, col, synProjectCellClass(display, col), display);
    }

    /** One pass for all scroll-column CSS classes (avoids 30+ template checks per cell). */
    function scrollCellClassNames(entry, col, colIdx, colsLen, display) {
      if (isGapEntry(entry)) {
        return isGapGreenPillarCol(col) ? 'syn-panel-gap-pillar syn-pillar-col' : '';
      }
      const row = entry.excelRow;
      const parts = [
        cellExtraClass(row, col, display),
        synPillarAccentClass(col),
        synExcelColTraceClass(col),
      ];
      if (isPillarColForEntry(entry, col)) parts.push('syn-pillar-col');
      if (usesPillarLetterOverlay(col)) parts.push('syn-pillar-overlay-host');
      if (isSynProjHeaderGreenCol(row, col) || isSynTargetTextGreenCell(col, display)) {
        parts.push('syn-proj-hdr-green');
      }
      if (isSynProjHeaderRedCol(row, col)) parts.push('syn-proj-hdr-red');
      if (headerEdgeRight(row, colIdx, colsLen)) parts.push('syn-header-edge-right');
      if (isSynSpacerDisplayExcelCol(col)) parts.push('syn-spacer-col-l');
      if (isSynHdrCjDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-cj-right');
      if (isSynHdrMaDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-ma-right');
      if (isSynHdrLmDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-lm-right');
      if (isSynHdrLmDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-lm-left');
      if (isSynHdrAaDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-aa-right');
      if (isSynAcAnTableCellEntry(entry, col)) parts.push('syn-ac-an-cell');
      if (isSynHdrAcAnDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-acan-right');
      if (isSynHdrAcAnDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-ac-left');
      if (isSynHdrAcAnDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-an-right');
      if (isSynApBbTableCellEntry(entry, col)) parts.push('syn-ap-bb-cell');
      if (isSynHdrApBbDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-apbb-right');
      if (isSynHdrApBbDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-ap-left');
      if (isSynHdrApBbDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-bb-right');
      if (isSynBsCeTableCellEntry(entry, col)) parts.push('syn-bs-ce-cell');
      if (isSynHdrBsCeDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-bsce-right');
      if (isSynHdrBsCeDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-bs-left');
      if (isSynHdrBsCeDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-ce-right');
      if (isSynBdBoTableCellEntry(entry, col)) parts.push('syn-bd-bo-cell');
      if (isSynHdrBdBoDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-bdbo-right');
      if (isSynHdrBdBoDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-bd-left');
      if (isSynHdrBdBoDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-bo-right');
      if (isSynCiCyTableCellEntry(entry, col)) parts.push('syn-ci-cy-cell');
      if (isSynHdrCiCyDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-cicy-right');
      if (isSynHdrCiCyDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-ci-left');
      if (isSynHdrCiCyDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-cy-right');
      if (isSynDaDpTableCellEntry(entry, col)) parts.push('syn-da-dp-cell');
      if (isSynHdrDaDpDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-dadp-right');
      if (isSynHdrDaDpDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-da-left');
      if (isSynHdrDaDpDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-dp-right');
      if (isSynDrEdTableCellEntry(entry, col)) parts.push('syn-dr-ed-cell');
      if (isSynHdrDrEdDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-dred-right');
      if (isSynHdrDrEdDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-dr-left');
      if (isSynHdrDrEdDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-ed-right');
      if (isSynEfEqTableCellEntry(entry, col)) parts.push('syn-ef-eq-cell');
      if (isSynHdrEfEqDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-efeq-right');
      if (isSynHdrEfEqDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-ef-left');
      if (isSynHdrEfEqDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-eq-right');
      if (isSynEsFeTableCellEntry(entry, col)) parts.push('syn-es-fe-cell');
      if (isSynHdrEsFeDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-esfe-right');
      if (isSynHdrEsFeDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-es-left');
      if (isSynHdrEsFeDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-fe-right');
      if (isSynFjFzTableCellEntry(entry, col)) parts.push('syn-fj-fz-cell');
      if (isSynHdrFjFzDividerRightEntry(entry, col)) parts.push('syn-hdr-edge-fjfz-right');
      if (isSynHdrFjFzDividerLeftEntry(entry, col)) parts.push('syn-hdr-edge-fj-left');
      if (isSynHdrFjFzDividerRightEdgeEntry(entry, col)) parts.push('syn-hdr-edge-fz-right');
      return parts.filter(Boolean).join(' ');
    }

    function buildPinnedCellModel(entry, colEntry) {
      const col = colEntry.col;
      const width = colEntry.width;
      if (isGapEntry(entry)) {
        return {
          col,
          width,
          gap: true,
          readonly: true,
          display: '',
          html: '',
          classes: '',
          style: [colStyle(col, width), stickyStyle(col)],
        };
      }
      const row = entry.excelRow;
      const display = cellDisplay(row, col);
      const readonly = cellReadonly(row, col);
      return {
        col,
        width,
        row,
        gap: false,
        display,
        html: cellDisplayHtml(row, col, display),
        readonly,
        classes: [
          cellExtraClass(row, col, display),
          { 'syn-header-edge-right': headerLabelEdgeRight(row) },
        ],
        style: [colStyle(col, width), stickyStyle(col), cellInlineStyle(row, col)],
      };
    }

    function buildScrollCellModel(entry, colEntry, colIdx, colsLen) {
      const col = colEntry.col;
      const width = colEntry.width;
      if (isGapEntry(entry)) {
        return {
          col,
          width,
          gap: true,
          readonly: true,
          display: '',
          html: '',
          classes: scrollCellClassNames(entry, col, colIdx, colsLen, ''),
          style: scrollDataCellStyle(entry, col, width),
        };
      }
      const row = entry.excelRow;
      const display = cellDisplay(row, col);
      const readonly = cellReadonly(row, col);
      return {
        col,
        width,
        row,
        gap: false,
        display,
        html: cellDisplayHtml(row, col, display),
        readonly,
        classes: scrollCellClassNames(entry, col, colIdx, colsLen, display),
        style: scrollDataCellStyle(entry, col, width),
      };
    }

    function rowEntryKey(entry) {
      if (isGapEntry(entry)) {
        return `panel-gap-${entry.gapKey || ((entry.gapBeforePanel ? 'top-' : 'bot-') + entry.gapIndex)}`;
      }
      return String(entry.excelRow);
    }

    const pinnedCellCache = new Map();
    const scrollCellCache = new Map();
    let cellModelCacheGen = 0;

    function bumpCellModelCache() {
      cellModelCacheGen += 1;
      if (cellModelCacheGen > 1000) {
        pinnedCellCache.clear();
        scrollCellCache.clear();
        cellModelCacheGen = 0;
      } else {
        pinnedCellCache.clear();
        scrollCellCache.clear();
      }
    }

    watch(
      () =>
        [
          editEpoch.value,
          synCalcTick.value,
          adaptationSumTick.value,
          liveCalcReady.value,
        ],
      () => bumpCellModelCache()
    );

    watch(
      () => adaptationSumTick.value,
      () => {
        editEpoch.value += 1;
        invalidateDisplayCache();
      }
    );

    function scrollCell(entry, colEntry, colIdx) {
      return getCachedScrollCellModel(entry, colEntry, colIdx);
    }

    function getCachedPinnedCellModel(entry, colEntry) {
      const key = `${cellModelCacheGen}:${rowEntryKey(entry)}:${colEntry.col}`;
      if (pinnedCellCache.has(key)) return pinnedCellCache.get(key);
      const model = buildPinnedCellModel(entry, colEntry);
      pinnedCellCache.set(key, model);
      return model;
    }

    function getCachedScrollCellModel(entry, colEntry, colIdx) {
      const colsLen = scrollColsLen.value;
      const key = `${cellModelCacheGen}:${rowEntryKey(entry)}:${colEntry.col}`;
      if (scrollCellCache.has(key)) return scrollCellCache.get(key);
      const model = buildScrollCellModel(entry, colEntry, colIdx, colsLen);
      scrollCellCache.set(key, model);
      return model;
    }

    const leftPadStyle = computed(() => {
      const lp = leftPad.value;
      return lp > 0 ? { width: `${lp}px`, minWidth: `${lp}px` } : null;
    });

    const rightPadStyle = computed(() => {
      const rp = rightPad.value;
      return rp > 0 ? { width: `${rp}px`, minWidth: `${rp}px` } : null;
    });

    /**
     * Displayed row number. The blank "between-18-19" gap takes number 19, so
     * every Excel row from 19 onward is shown shifted by +1 (old 19 → 20, …).
     */
    function synRowNumberLabel(entry) {
      if (!entry) return '';
      if (entry.gapKey === 'between-18-19') return 19;
      const r = entry.excelRow;
      if (r == null) return '';
      return r >= 19 ? r + 1 : r;
    }

    /** Row metadata — stable when only horizontal scroll changes. */
    const visibleRowMeta = computed(() => {
      if (!props.paneVisible) return [];
      const map = cellMap.value;
      const sheet = props.sheet;
      return visibleRows.value.map((entry) => ({
        key: rowEntryKey(entry),
        entry,
        excelRow: entry.excelRow != null ? entry.excelRow : null,
        rowNum: synRowNumberLabel(entry),
        displayRow: entry.displayRow,
        rowLabel:
          entry.excelRow != null ? synLabel(map, entry.excelRow) : '',
        isAdaptationSumRow:
          entry.excelRow != null &&
          entry.excelRow === synAdaptationSumRow(sheet),
        rowClasses: cachedEntryRowClasses(entry),
      }));
    });

    const pinnedCellsByRowKey = computed(() => {
      if (!props.paneVisible) return {};
      void editEpoch.value;
      void synCalcTick.value;
      const pinned = pinnedCols.value;
      const out = {};
      for (const entry of visibleRows.value) {
        const key = rowEntryKey(entry);
        out[key] = pinned.map((p) => getCachedPinnedCellModel(entry, p));
      }
      return out;
    });

    const scrollCellsByRowKey = computed(() => {
      if (!props.paneVisible) return {};
      void editEpoch.value;
      void synCalcTick.value;
      void visibleScrollCols.value;
      const cols = visibleScrollCols.value;
      const out = {};
      for (const entry of visibleRows.value) {
        const key = rowEntryKey(entry);
        out[key] = cols.map((c, i) => getCachedScrollCellModel(entry, c, i));
      }
      return out;
    });

    function cellDisplayHtml(row, col, displayed) {
      const value = cellShowValue(row, col, displayed);
      const bevHtml = synHdrBevDisplayHtml(row, value);
      if (bevHtml) return bevHtml;
      return String(value != null ? value : '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    function updateViewport() {
      if (!scrollEl.value) return;
      viewportH.value = scrollEl.value.clientHeight;
      viewportW.value = scrollEl.value.clientWidth;
    }

    watch(
      () => rowCount.value,
      () => {
        rowScrollCache.reset();
        colScrollCache.reset();
        nextTick(() => scrollSync.flush());
      }
    );

    watch(
      () => props.outlineOnly,
      () => {
        scrollTop.value = 0;
        scrollLeft.value = 0;
        if (scrollEl.value) {
          scrollEl.value.scrollTop = 0;
          scrollEl.value.scrollLeft = 0;
        }
        colScrollCache.reset();
        rowScrollCache.reset();
        invalidateDisplayCache();
      }
    );

    watch(
      () => props.paneVisible,
      (visible) => {
        if (!visible) return;
        nextTick(() => {
          updateViewport();
          scrollSync.flush();
        });
      },
      { immediate: true }
    );

    // Bind once per sheet — avoids full display-cache flush on every render pass.
    watch(
      () => props.sheet,
      (sheet) => {
        if (!sheet || !props.session || !props.session.bindSynthesisGrid) return;
        const map = cellMap.value;
        props.session.bindSynthesisGrid(
          (row, col) => getCell(map, row, col),
          sheet,
          (row) => synLabel(map, row),
          (row) => synRowStyleClass(map, row, sheet)
        );
      },
      { immediate: true }
    );

    onMounted(() => {
      updateViewport();
      scrollSync.flush();
      window.addEventListener('resize', updateViewport);
      if (props.sheet && props.sheet.materializedCalc) return;
      requestAnimationFrame(() => {
        const enable = () => {
          liveCalcReady.value = true;
        };
        if (typeof requestIdleCallback === 'function') {
          requestIdleCallback(enable, { timeout: 1500 });
        } else {
          setTimeout(enable, 350);
        }
      });
    });
    onUnmounted(() => {
      if (scrollEndTimer) clearTimeout(scrollEndTimer);
      window.removeEventListener('resize', updateViewport);
      scrollSync.dispose();
    });

    // Vue template only sees setup() return values — keep all template helpers here.
    /** No-op fallback when a cached synStore.js omits a table helper (avoids render crash). */
    function synTemplateFn(name, fn) {
      if (typeof fn === 'function') return fn;
      console.error(
        `[SynthesisGrid] ${name} is not a function — hard-refresh (Ctrl+F5) to reload synStore.js`
      );
      return () => false;
    }

    const synGridTemplateHelpers = {
      excelToDisplayCol,
      isSynPillarCol: (col) => isSynPillarCol(col, pillarColumns.value),
      isSynPillarColAtRow: (col, row) =>
        isSynPillarColAtRow(col, row, pillarColumns.value),
      isSynSp2DisplayExcelCol,
      synPillarAccentClass,
      isSynSpacerDisplayExcelCol,
      isSynForceWhiteExcelCol,
      synSpacerColClass,
      isSynHdrLmDividerRightCol,
      isSynHdrLmDividerLeftCol,
      isSynHdrAaDividerRightCol,
      isSynHdrLmDividerRightEntry: synTemplateFn(
        'isSynHdrLmDividerRightEntry',
        isSynHdrLmDividerRightEntry
      ),
      isSynHdrLmDividerLeftEntry: synTemplateFn(
        'isSynHdrLmDividerLeftEntry',
        isSynHdrLmDividerLeftEntry
      ),
      isSynHdrAaDividerRightEntry: synTemplateFn(
        'isSynHdrAaDividerRightEntry',
        isSynHdrAaDividerRightEntry
      ),
      isSynHdrCjDividerRightEntry: synTemplateFn(
        'isSynHdrCjDividerRightEntry',
        isSynHdrCjDividerRightEntry
      ),
      isSynHdrMaDividerRightEntry: synTemplateFn(
        'isSynHdrMaDividerRightEntry',
        isSynHdrMaDividerRightEntry
      ),
      isSynAcAnTableCellEntry: synTemplateFn(
        'isSynAcAnTableCellEntry',
        isSynAcAnTableCellEntry
      ),
      isSynHdrAcAnDividerRightEntry: synTemplateFn(
        'isSynHdrAcAnDividerRightEntry',
        isSynHdrAcAnDividerRightEntry
      ),
      isSynHdrAcAnDividerLeftEntry: synTemplateFn(
        'isSynHdrAcAnDividerLeftEntry',
        isSynHdrAcAnDividerLeftEntry
      ),
      isSynHdrAcAnDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrAcAnDividerRightEdgeEntry',
        isSynHdrAcAnDividerRightEdgeEntry
      ),
      isSynApBbTableCellEntry: synTemplateFn(
        'isSynApBbTableCellEntry',
        isSynApBbTableCellEntry
      ),
      isSynHdrApBbDividerRightEntry: synTemplateFn(
        'isSynHdrApBbDividerRightEntry',
        isSynHdrApBbDividerRightEntry
      ),
      isSynHdrApBbDividerLeftEntry: synTemplateFn(
        'isSynHdrApBbDividerLeftEntry',
        isSynHdrApBbDividerLeftEntry
      ),
      isSynHdrApBbDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrApBbDividerRightEdgeEntry',
        isSynHdrApBbDividerRightEdgeEntry
      ),
      isSynBsCeTableCellEntry: synTemplateFn(
        'isSynBsCeTableCellEntry',
        isSynBsCeTableCellEntry
      ),
      isSynHdrBsCeDividerRightEntry: synTemplateFn(
        'isSynHdrBsCeDividerRightEntry',
        isSynHdrBsCeDividerRightEntry
      ),
      isSynHdrBsCeDividerLeftEntry: synTemplateFn(
        'isSynHdrBsCeDividerLeftEntry',
        isSynHdrBsCeDividerLeftEntry
      ),
      isSynHdrBsCeDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrBsCeDividerRightEdgeEntry',
        isSynHdrBsCeDividerRightEdgeEntry
      ),
      isSynBdBoTableCellEntry: synTemplateFn(
        'isSynBdBoTableCellEntry',
        isSynBdBoTableCellEntry
      ),
      isSynHdrBdBoDividerRightEntry: synTemplateFn(
        'isSynHdrBdBoDividerRightEntry',
        isSynHdrBdBoDividerRightEntry
      ),
      isSynHdrBdBoDividerLeftEntry: synTemplateFn(
        'isSynHdrBdBoDividerLeftEntry',
        isSynHdrBdBoDividerLeftEntry
      ),
      isSynHdrBdBoDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrBdBoDividerRightEdgeEntry',
        isSynHdrBdBoDividerRightEdgeEntry
      ),
      isSynCiCyTableCellEntry: synTemplateFn(
        'isSynCiCyTableCellEntry',
        isSynCiCyTableCellEntry
      ),
      isSynHdrCiCyDividerRightEntry: synTemplateFn(
        'isSynHdrCiCyDividerRightEntry',
        isSynHdrCiCyDividerRightEntry
      ),
      isSynHdrCiCyDividerLeftEntry: synTemplateFn(
        'isSynHdrCiCyDividerLeftEntry',
        isSynHdrCiCyDividerLeftEntry
      ),
      isSynHdrCiCyDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrCiCyDividerRightEdgeEntry',
        isSynHdrCiCyDividerRightEdgeEntry
      ),
      isSynDaDpTableCellEntry: synTemplateFn(
        'isSynDaDpTableCellEntry',
        isSynDaDpTableCellEntry
      ),
      isSynHdrDaDpDividerRightEntry: synTemplateFn(
        'isSynHdrDaDpDividerRightEntry',
        isSynHdrDaDpDividerRightEntry
      ),
      isSynHdrDaDpDividerLeftEntry: synTemplateFn(
        'isSynHdrDaDpDividerLeftEntry',
        isSynHdrDaDpDividerLeftEntry
      ),
      isSynHdrDaDpDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrDaDpDividerRightEdgeEntry',
        isSynHdrDaDpDividerRightEdgeEntry
      ),
      isSynDrEdTableCellEntry: synTemplateFn(
        'isSynDrEdTableCellEntry',
        isSynDrEdTableCellEntry
      ),
      isSynHdrDrEdDividerRightEntry: synTemplateFn(
        'isSynHdrDrEdDividerRightEntry',
        isSynHdrDrEdDividerRightEntry
      ),
      isSynHdrDrEdDividerLeftEntry: synTemplateFn(
        'isSynHdrDrEdDividerLeftEntry',
        isSynHdrDrEdDividerLeftEntry
      ),
      isSynHdrDrEdDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrDrEdDividerRightEdgeEntry',
        isSynHdrDrEdDividerRightEdgeEntry
      ),
      isSynEfEqTableCellEntry: synTemplateFn(
        'isSynEfEqTableCellEntry',
        isSynEfEqTableCellEntry
      ),
      isSynHdrEfEqDividerRightEntry: synTemplateFn(
        'isSynHdrEfEqDividerRightEntry',
        isSynHdrEfEqDividerRightEntry
      ),
      isSynHdrEfEqDividerLeftEntry: synTemplateFn(
        'isSynHdrEfEqDividerLeftEntry',
        isSynHdrEfEqDividerLeftEntry
      ),
      isSynHdrEfEqDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrEfEqDividerRightEdgeEntry',
        isSynHdrEfEqDividerRightEdgeEntry
      ),
      isSynEsFeTableCellEntry: synTemplateFn(
        'isSynEsFeTableCellEntry',
        isSynEsFeTableCellEntry
      ),
      isSynHdrEsFeDividerRightEntry: synTemplateFn(
        'isSynHdrEsFeDividerRightEntry',
        isSynHdrEsFeDividerRightEntry
      ),
      isSynHdrEsFeDividerLeftEntry: synTemplateFn(
        'isSynHdrEsFeDividerLeftEntry',
        isSynHdrEsFeDividerLeftEntry
      ),
      isSynHdrEsFeDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrEsFeDividerRightEdgeEntry',
        isSynHdrEsFeDividerRightEdgeEntry
      ),
      isSynFjFzTableCellEntry: synTemplateFn(
        'isSynFjFzTableCellEntry',
        isSynFjFzTableCellEntry
      ),
      isSynHdrFjFzDividerRightEntry: synTemplateFn(
        'isSynHdrFjFzDividerRightEntry',
        isSynHdrFjFzDividerRightEntry
      ),
      isSynHdrFjFzDividerLeftEntry: synTemplateFn(
        'isSynHdrFjFzDividerLeftEntry',
        isSynHdrFjFzDividerLeftEntry
      ),
      isSynHdrFjFzDividerRightEdgeEntry: synTemplateFn(
        'isSynHdrFjFzDividerRightEdgeEntry',
        isSynHdrFjFzDividerRightEdgeEntry
      ),
      isSynProjHeaderGreenCol,
      isSynTargetTextGreenCell,
      synProjHeaderGreenStyle,
      isSynProjHeaderRedCol,
      synProjHeaderRedStyle,
    };

    return {
      scrollEl,
      pinnedCols,
      visibleScrollCols,
      tableWidth,
      visibleRowMeta,
      pinnedCellsByRowKey,
      scrollCellsByRowKey,
      topSpacer,
      bottomSpacer,
      leftPad,
      rightPad,
      leftPadStyle,
      rightPadStyle,
      colspan,
      SYN_STICKY_COL,
      colStyle,
      stickyStyle,
      cellDisplay,
      cellReadonly,
      entryRowClasses,
      isGapEntry,
      isPillarColForEntry,
      isGapGreenPillarCol,
      pillarTitle,
      pillarColumns,
      ...synGridTemplateHelpers,
      cellInlineStyle,
      scrollDataCellStyle,
      cellExtraClass,
      headerEdgeRight,
      headerLabelEdgeRight,
      pillarLetterOverlays,
      usesPillarLetterOverlay,
      synExcelColTraceClass,
      onScroll: onGridScroll,
      isSearchHit: gridSearch.isSearchHit,
      isSearchFocus: gridSearch.isSearchFocus,
      isCellActive,
      cellShowValue,
      cellDisplayHtml,
      onCellMouseDown,
      onCellSpanMouseDown,
      onCellFocus,
      onCellInput: onCellInputLive,
      onCellBlur,
      onCellKeydown,
      selectedCell,
      formulaText,
      formulaCellRef,
      onReadonlyCellSelect,
    };
  },
  template: `
    <div class="bd-grid-root synthesis-grid">
      <div v-if="selectedCell" class="syn-formula-bar" role="status">
        <span class="syn-formula-bar-ref">{{ formulaCellRef }}</span>
        <span class="syn-formula-bar-text">{{ formulaText || '—' }}</span>
      </div>
      <div class="bd-grid-scroll syn-scroll" ref="scrollEl" @scroll.passive="onScroll">
        <div
          class="syn-table-wrap"
          :style="{ width: tableWidth + 'px', minWidth: tableWidth + 'px' }"
        >
        <div
          v-if="pillarLetterOverlays.length"
          class="syn-pillar-overlays"
          aria-hidden="true"
        >
          <div
            v-for="pillar in pillarLetterOverlays"
            :key="'pillar-overlay-' + pillar.col"
            class="syn-pillar-overlay-col"
            :class="synPillarAccentClass(pillar.col)"
            :style="{
              left: pillar.left + 'px',
              width: pillar.width + 'px',
            }"
          >
            <span
              v-for="(item, idx) in pillar.chars"
              :key="pillar.col + '-' + idx"
              class="syn-pillar-overlay-char"
              :style="{ top: item.top + 'px' }"
            >{{ item.ch }}</span>
          </div>
        </div>
        <table
          class="bd-table syn-table"
          role="grid"
          :style="{ width: tableWidth + 'px', minWidth: tableWidth + 'px' }"
        >
          <thead>
            <tr class="hdr-row-letters">
              <th class="corner syn-row-num-hdr"></th>
              <th
                v-for="entry in pinnedCols"
                :key="'L-' + entry.col"
                class="col-letter col-sticky-label"
                :style="[colStyle(entry.col, entry.width), stickyStyle(entry.col)]"
              >{{ entry.letter }}</th>
              <th v-if="leftPad > 0" class="syn-pad" :style="{ width: leftPad + 'px', minWidth: leftPad + 'px' }"></th>
              <th
                v-for="entry in visibleScrollCols"
                :key="'L-' + entry.col"
                class="col-letter"
                :class="{
                  'syn-spacer-col-l-hdr': isSynSpacerDisplayExcelCol(entry.col),
                  'syn-force-white-col': isSynForceWhiteExcelCol(entry.col),
                }"
                :style="colStyle(entry.col, entry.width)"
              >{{ entry.letter }}</th>
              <th v-if="rightPad > 0" class="syn-pad" :style="{ width: rightPad + 'px', minWidth: rightPad + 'px' }"></th>
            </tr>
          </thead>
          <tbody>
            <tr v-if="topSpacer > 0" class="bd-spacer-top">
              <td :colspan="colspan" :style="{ height: topSpacer + 'px', border: 'none', padding: 0 }"></td>
            </tr>
            <tr
              v-for="row in visibleRowMeta"
              :key="row.key"
              class="grid-row-cv"
              :class="row.rowClasses"
            >
              <td
                class="row-num syn-row-num"
                :class="{ 'syn-adaptation-sum-row-num': row.isAdaptationSumRow }"
                :title="row.rowLabel || ''"
              >
                {{ row.rowNum }}
              </td>
              <td
                v-for="cell in pinnedCellsByRowKey[row.key]"
                :key="row.key + '-p-' + cell.col"
                class="data-cell col-sticky-label syn-label-col"
                :class="[
                  { readonly: cell.readonly },
                  cell.classes,
                  {
                    'grid-search-hit': isSearchHit(cell.row, cell.col),
                    'grid-search-focus': isSearchFocus(cell.row, cell.col),
                  },
                ]"
                :style="cell.style"
              >
                <template v-if="cell.gap || cell.readonly">
                  <span @mousedown="onReadonlyCellSelect(cell.row, cell.col, $event)">{{ cell.display }}</span>
                </template>
                <input
                  v-else-if="isCellActive(cell.row, cell.col)"
                  type="text"
                  class="grid-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  :data-grid-row="cell.row"
                  :data-grid-col="cell.col"
                  @mousedown="onCellMouseDown(cell.row, cell.col, $event)"
                  @focus="onCellFocus(cell.row, cell.col, $event)"
                  @input="onCellInputLive(cell.row, cell.col, $event)"
                  @blur="onCellBlur(cell.row, cell.col, $event)"
                  @keydown="onCellKeydown(cell.row, cell.col, $event)"
                />
                <span
                  v-else
                  class="grid-cell-display"
                  @mousedown="onCellSpanMouseDown(cell.row, cell.col, $event)"
                  v-html="cell.html"
                ></span>
              </td>
              <td v-if="leftPad > 0" class="syn-pad" :style="leftPadStyle"></td>
              <td
                v-for="cell in scrollCellsByRowKey[row.key]"
                :key="row.key + '-' + cell.col"
                class="data-cell"
                :class="[
                  { readonly: cell.readonly },
                  cell.classes,
                  {
                    'grid-search-hit': isSearchHit(cell.row, cell.col),
                    'grid-search-focus': isSearchFocus(cell.row, cell.col),
                  },
                ]"
                :style="cell.style"
              >
                <template v-if="cell.gap || cell.readonly">
                  <span @mousedown="onReadonlyCellSelect(cell.row, cell.col, $event)">{{ cell.display }}</span>
                </template>
                <input
                  v-else-if="isCellActive(cell.row, cell.col)"
                  type="text"
                  class="grid-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  :data-grid-row="cell.row"
                  :data-grid-col="cell.col"
                  @mousedown="onCellMouseDown(cell.row, cell.col, $event)"
                  @focus="onCellFocus(cell.row, cell.col, $event)"
                  @input="onCellInputLive(cell.row, cell.col, $event)"
                  @blur="onCellBlur(cell.row, cell.col, $event)"
                  @keydown="onCellKeydown(cell.row, cell.col, $event)"
                />
                <span
                  v-else
                  class="grid-cell-display"
                  @mousedown="onCellSpanMouseDown(cell.row, cell.col, $event)"
                  v-html="cell.html"
                ></span>
              </td>
              <td v-if="rightPad > 0" class="syn-pad" :style="rightPadStyle"></td>
            </tr>
            <tr v-if="bottomSpacer > 0" class="bd-spacer-bottom">
              <td :colspan="colspan" :style="{ height: bottomSpacer + 'px', border: 'none', padding: 0 }"></td>
            </tr>
          </tbody>
        </table>
        </div>
      </div>
    </div>
  `,
};
