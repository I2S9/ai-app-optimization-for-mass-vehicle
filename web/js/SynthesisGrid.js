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
import { isSynAdaptationSumCell } from './synthesisCalc.js?v=adapt-sum1';
import { createGridCellEditor } from './gridCellEdit.js?v=grid-nav4';
import { createGridCellNavigation } from './gridCellNavigation.js?v=grid-nav4';
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
  isSynRow20MopGreenCol,
  isSynRow20PaaRedCol,
  synRow20MopGreenStyle,
  synRow20PaaRedStyle,
  isSynRow21MpsvyRedCol,
  isSynRow21MaaWhiteCol,
  synRow21MpsvyRedStyle,
  synRow21MaaWhiteStyle,
  isSynRow25MaGreenCol,
  synRow25MaGreenStyle,
  isSynRow16FluoEvery3FromMCol,
  isSynRow17FluoEvery3FromMCol,
  synRow16FluoStyle,
  SYN_DISPLAY_GREEN_ROWS,
  SYN_DISPLAY_GREEN_BG,
  isSynDisplayRowGreyMaaCol,
  synDisplayRowGreyMaaStyle,
  isSynDisplayRowGreenMaaCol,
  synDisplayRowGreenMaaStyle,
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
} from './synStore.js?v=syn-bands3';
import {
  SYN_STICKY_COL,
  excelToDisplayCol,
  synStickyColWidth,
  synPillarColWidth,
  isSynBuiltinPillarExcelCol,
} from './synthesisPerf.js?v=syn-pillar-cg2';
import {
  ROW_H,
  colOverscanPx,
  rowOverscanForColCount,
  shouldVirtualizeRows,
  shouldVirtualizeCols,
  createRowScrollCache,
  createScrollRafSync,
  SYN_MAX_RENDERED_ROWS,
} from './gridScroll.js?v=syn-nav-perf1';
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
  },
  emits: ['cell-change'],
  setup(props, { emit }) {
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const scrollLeft = ref(0);
    const viewportH = ref(600);
    const viewportW = ref(1000);
    const cellMap = shallowRef(props.sheet?.cellMap || new Map());

    watch(
      () => props.sheet,
      (sheet) => {
        if (sheet?.cellMap instanceof Map) {
          cellMap.value = sheet.cellMap;
        } else if (sheet) {
          cellMap.value = buildCellMap(sheet.cells, sheet.headerRows);
        }
      },
      { immediate: true }
    );

    const pillarColumns = computed(() => {
      const map = buildSynPillarColumns(props.sheet, cellMap.value);
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
            synPillarColWidth(col, props.sheet, pillarColumns.value) ?? 54
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
    const colBufferPx = computed(() => colOverscanPx(viewportW.value));

    const visibleScrollCols = computed(() => {
      if (!virtualizeCols.value) return scrollableCols.value;
      const buf = colBufferPx.value;
      const min = scrollLeft.value - buf;
      const max = scrollLeft.value + viewportW.value + buf;
      return scrollableCols.value.filter(
        (c) => c.left + c.width >= min && c.left <= max
      );
    });

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

    const bodyRows = computed(() =>
      computeSynBodyRows(props.sheet, cellMap.value, props.outlineOnly)
    );
    const rowCount = computed(() => bodyRows.value.length);

    const virtualizeRows = computed(() =>
      shouldVirtualizeRows(rowCount.value, viewportH.value)
    );
    const rowBuffer = computed(() =>
      rowOverscanForColCount(viewportH.value, displayColumns.value.length)
    );
    const rowScrollCache = createRowScrollCache(SYN_MAX_RENDERED_ROWS);
    const visibleStart = ref(0);
    const visibleEnd = ref(0);

    watchEffect(() => {
      const count = rowCount.value;
      if (!virtualizeRows.value) {
        visibleStart.value = 0;
        visibleEnd.value = count;
        return;
      }
      const range = rowScrollCache.resolve(
        scrollTop.value,
        viewportH.value,
        count,
        rowBuffer.value
      );
      visibleStart.value = range.start;
      visibleEnd.value = range.end;
    });

    const visibleRows = computed(() =>
      bodyRows.value.slice(visibleStart.value, visibleEnd.value)
    );
    const topSpacer = computed(() =>
      virtualizeRows.value ? visibleStart.value * ROW_H : 0
    );
    const bottomSpacer = computed(() => {
      if (!virtualizeRows.value) return 0;
      const shown = visibleEnd.value - visibleStart.value;
      return Math.max(0, (rowCount.value - visibleStart.value - shown) * ROW_H);
    });

    function bodyTopForExcelRow(excelRow) {
      const rows = bodyRows.value;
      for (let i = 0; i < rows.length; i++) {
        if (rows[i].excelRow === excelRow) {
          return SYN_HEAD_ROW_H + i * ROW_H;
        }
      }
      return null;
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
      const overlays = [];
      for (const layout of columnLayout.value) {
        if (!SYN_PILLAR_OVERLAY_COLS.has(layout.col)) continue;
        if (
          !pillarColumns.value.has(layout.col) &&
          !isSynBuiltinPillarExcelCol(layout.col)
        ) {
          continue;
        }
        const title =
          pillarColumns.value.get(layout.col)?.title ??
          SYN_BUILTIN_PILLAR_META[layout.col]?.title ??
          '';
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
      const width = w ?? widthByCol.value.get(col) ?? 54;
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
    const calcRevision = computed(() => props.session?.revision?.value ?? 0);

    watch(
      () => props.externalEditTick,
      () => {
        editEpoch.value += 1;
      }
    );

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
      editEpoch.value += 1;
      emit('cell-change', {
        row,
        col,
        value,
        sheet: 'SYNTHESIS',
        previousValue: previousValue ?? '',
      });
      void props.session
        ?.setCellValue?.('SYNTHESIS', row, col, value)
        ?.catch((e) => console.warn('Synthesis setCellValue:', e));
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
      onCellSpanMouseDown,
      onCellFocus,
      onCellInput,
      onCellBlur,
      onCellKeydown: onCellKeydownBase,
      prepareNavigate,
      beginNavigationTo,
      activateCell,
      setNavigationLock,
    } = cellEditor;

    /** Only virtual gap rows (no Excel row) stay display-only. */
    function cellReadonly(row, col) {
      if (row == null) return true;
      return isSynAdaptationSumCell(row, col);
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
        return entry?.left ?? null;
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

    function formatVal(v) {
      const formatted = formatSynNumericDisplay(v);
      return formatted === '' && v != null && String(v).trim() !== ''
        ? String(v)
        : formatted;
    }

    function cellDisplay(row, col) {
      void calcRevision.value;
      void editEpoch.value;
      if (row == null) return '';
      if (isSynPillarColAtRow(col, row, pillarColumns.value)) {
        if (usesPillarLetterOverlay(col)) return '';
        return synPillarLetterForRow(
          row,
          col,
          pillarColumns.value,
          cellMap.value,
          props.sheet
        );
      }
      const cell = getCell(cellMap.value, row, col);
      let displayCell = cell;
      if (cell?.userEdited) {
        displayCell = cell;
      } else if (props.session?.getDisplayValue) {
        const fromSession = props.session.getDisplayValue(
          'SYNTHESIS',
          row,
          col,
          cell
        );
        if (fromSession != null && !cell?.userEdited) {
          displayCell = { ...(cell || {}), v: fromSession };
        }
      }
      const shown = synDisplayValue(
        displayCell,
        cellMap.value,
        row,
        col,
        props.sheet,
        pillarColumns.value
      );
      if (isSynMetricRow(row)) return shown;
      return formatVal(shown);
    }

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
      return pillarColumns.value.get(col)?.title ?? '';
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
      if (isSynForceWhiteExcelCol(col)) return base;
      if (isSynSpacerDisplayExcelCol(col)) return base;
      const display = cellDisplay(row, col);
      const energyStyle = synHdrEnergyValueStyle(display);
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
      if (isSynProjHeaderGreenCol(row, col)) {
        return { ...base, ...synProjHeaderGreenStyle() };
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
      if (isSynRow20MopGreenCol(row, col)) {
        return { ...base, ...synRow20MopGreenStyle() };
      }
      if (isSynRow20PaaRedCol(row, col)) {
        return { ...base, ...synRow20PaaRedStyle() };
      }
      if (isSynRow21MpsvyRedCol(row, col)) {
        return { ...base, ...synRow21MpsvyRedStyle() };
      }
      if (isSynRow21MaaWhiteCol(row, col)) {
        return { ...base, ...synRow21MaaWhiteStyle() };
      }
      if (isSynRow25MaGreenCol(row, col)) {
        return { ...base, ...synRow25MaGreenStyle() };
      }
      if (isSynRow16FluoEvery3FromMCol(row, col)) {
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
      return { ...base, ...cellInlineStyle(row, col) };
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
      const energyCls = synHdrEnergyValueClass(display);
      if (energyCls) return withHdrPanelBold(row, col, energyCls, display);
      // Display-row based styling (after spacer/pillar checks).
      const displayRow = bodyRows.value.find((e) => e.excelRow === row)?.displayRow;
      if (isSynDisplayRowGreyMaaCol(displayRow, col)) {
        return withHdrPanelBold(row, col, 'syn-displayrow-grey-maa', display);
      }
      if (isSynDisplayRowGreenMaaCol(displayRow, col)) {
        return withHdrPanelBold(row, col, 'syn-displayrow-green-maa', display);
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
      if (isSynProjHeaderGreenCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-proj-hdr-green', display);
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
      if (isSynRow20MopGreenCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row20-mop-green', display);
      }
      if (isSynRow20PaaRedCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row20-paa-red', display);
      }
      if (isSynRow21MpsvyRedCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row21-mpsvy-red', display);
      }
      if (isSynRow21MaaWhiteCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row21-maa-white', display);
      }
      if (isSynRow25MaGreenCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-row25-ma-green', display);
      }
      if (isSynRow16FluoEvery3FromMCol(row, col)) {
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

    function updateViewport() {
      if (!scrollEl.value) return;
      viewportH.value = scrollEl.value.clientHeight;
      viewportW.value = scrollEl.value.clientWidth;
    }

    watch(
      () => props.outlineOnly,
      () => {
        rowScrollCache.reset();
        scrollTop.value = 0;
        if (scrollEl.value) scrollEl.value.scrollTop = 0;
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

    onMounted(() => {
      props.session?.bindSynthesisGrid?.((row, col) =>
        getCell(cellMap.value, row, col)
      );
      updateViewport();
      scrollSync.flush();
      window.addEventListener('resize', updateViewport);
    });
    onUnmounted(() => {
      window.removeEventListener('resize', updateViewport);
      scrollSync.dispose();
    });

    // Vue template only sees setup() return values — keep all template helpers here.
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
      isSynHdrLmDividerRightEntry,
      isSynHdrLmDividerLeftEntry,
      isSynHdrAaDividerRightEntry,
      isSynHdrCjDividerRightEntry,
      isSynHdrMaDividerRightEntry,
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
      isSynProjHeaderGreenCol,
      synProjHeaderGreenStyle,
      isSynProjHeaderRedCol,
      synProjHeaderRedStyle,
    };

    return {
      scrollEl,
      pinnedCols,
      visibleScrollCols,
      tableWidth,
      visibleRows,
      topSpacer,
      bottomSpacer,
      leftPad,
      rightPad,
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
      onScroll: scrollSync.onScroll,
      isCellActive,
      cellShowValue,
      onCellMouseDown,
      onCellSpanMouseDown,
      onCellFocus,
      onCellInput,
      onCellBlur,
      onCellKeydown,
    };
  },
  template: `
    <div class="bd-grid-root synthesis-grid">
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
              v-for="entry in visibleRows"
              :key="isGapEntry(entry) ? ('panel-gap-' + (entry.gapKey || ((entry.gapBeforePanel ? 'top-' : 'bot-') + entry.gapIndex))) : entry.excelRow"
              :class="entryRowClasses(entry)"
            >
              <td class="row-num syn-row-num">{{ entry.displayRow }}</td>
              <td
                v-for="p in pinnedCols"
                :key="(isGapEntry(entry) ? 'gap' : entry.excelRow) + '-p-' + p.col"
                class="data-cell col-sticky-label syn-label-col"
                :class="[
                  { readonly: cellReadonly(entry.excelRow, p.col) },
                  isGapEntry(entry)
                    ? ''
                    : cellExtraClass(entry.excelRow, p.col, cellDisplay(entry.excelRow, p.col)),
                  {
                    'syn-header-edge-right':
                      !isGapEntry(entry) && headerLabelEdgeRight(entry.excelRow),
                  },
                ]"
                :style="[
                  colStyle(p.col, p.width),
                  stickyStyle(p.col),
                  isGapEntry(entry) ? {} : cellInlineStyle(entry.excelRow, p.col),
                ]"
              >
                <template v-if="isGapEntry(entry) || cellReadonly(entry.excelRow, p.col)">
                  <span>{{ isGapEntry(entry) ? '' : cellDisplay(entry.excelRow, p.col) }}</span>
                </template>
                <input
                  v-else-if="isCellActive(entry.excelRow, p.col)"
                  type="text"
                  class="grid-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  :data-grid-row="entry.excelRow"
                  :data-grid-col="p.col"
                  @mousedown="onCellMouseDown(entry.excelRow, p.col, $event)"
                  @focus="onCellFocus(entry.excelRow, p.col, $event)"
                  @input="onCellInput(entry.excelRow, p.col, $event)"
                  @blur="onCellBlur(entry.excelRow, p.col, $event)"
                  @keydown="onCellKeydown(entry.excelRow, p.col, $event)"
                />
                <span
                  v-else
                  class="grid-cell-display"
                  @mousedown="onCellSpanMouseDown(entry.excelRow, p.col, $event)"
                >{{ cellShowValue(entry.excelRow, p.col, cellDisplay(entry.excelRow, p.col)) }}</span>
              </td>
              <td v-if="leftPad > 0" class="syn-pad" :style="{ width: leftPad + 'px', minWidth: leftPad + 'px' }"></td>
              <td
                v-for="(colEntry, colIdx) in visibleScrollCols"
                :key="(isGapEntry(entry) ? 'gap' : entry.excelRow) + '-' + colEntry.col"
                class="data-cell"
                :class="[
                  { readonly: cellReadonly(entry.excelRow, colEntry.col) },
                  isGapEntry(entry)
                    ? ''
                    : cellExtraClass(
                        entry.excelRow,
                        colEntry.col,
                        cellDisplay(entry.excelRow, colEntry.col)
                      ),
                  synPillarAccentClass(colEntry.col),
                  synExcelColTraceClass(colEntry.col),
                  {
                    'syn-pillar-col': isGapEntry(entry)
                      ? isGapGreenPillarCol(colEntry.col)
                      : isPillarColForEntry(entry, colEntry.col),
                    'syn-pillar-overlay-host':
                      !isGapEntry(entry) &&
                      usesPillarLetterOverlay(colEntry.col),
                    'syn-panel-gap-pillar':
                      isGapEntry(entry) && isGapGreenPillarCol(colEntry.col),
                    'syn-proj-hdr-green':
                      !isGapEntry(entry) &&
                      isSynProjHeaderGreenCol(entry.excelRow, colEntry.col),
                    'syn-proj-hdr-red':
                      !isGapEntry(entry) &&
                      isSynProjHeaderRedCol(entry.excelRow, colEntry.col),
                    'syn-header-edge-right':
                      !isGapEntry(entry) &&
                      headerEdgeRight(entry.excelRow, colIdx, visibleScrollCols.length),
                    'syn-spacer-col-l': isSynSpacerDisplayExcelCol(colEntry.col),
                    'syn-hdr-edge-cj-right': isSynHdrCjDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-ma-right': isSynHdrMaDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-lm-right': isSynHdrLmDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-lm-left': isSynHdrLmDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-aa-right': isSynHdrAaDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-ac-an-cell': isSynAcAnTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-acan-right': isSynHdrAcAnDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-ac-left': isSynHdrAcAnDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-an-right': isSynHdrAcAnDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-ap-bb-cell': isSynApBbTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-apbb-right': isSynHdrApBbDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-ap-left': isSynHdrApBbDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-bb-right': isSynHdrApBbDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-bs-ce-cell': isSynBsCeTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-bsce-right': isSynHdrBsCeDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-bs-left': isSynHdrBsCeDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-ce-right': isSynHdrBsCeDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-bd-bo-cell': isSynBdBoTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-bdbo-right': isSynHdrBdBoDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-bd-left': isSynHdrBdBoDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-bo-right': isSynHdrBdBoDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-ci-cy-cell': isSynCiCyTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-cicy-right': isSynHdrCiCyDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-ci-left': isSynHdrCiCyDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-cy-right': isSynHdrCiCyDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-da-dp-cell': isSynDaDpTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-dadp-right': isSynHdrDaDpDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-da-left': isSynHdrDaDpDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-dp-right': isSynHdrDaDpDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-dr-ed-cell': isSynDrEdTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-dred-right': isSynHdrDrEdDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-dr-left': isSynHdrDrEdDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-ed-right': isSynHdrDrEdDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-ef-eq-cell': isSynEfEqTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-efeq-right': isSynHdrEfEqDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-ef-left': isSynHdrEfEqDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-eq-right': isSynHdrEfEqDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-es-fe-cell': isSynEsFeTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-esfe-right': isSynHdrEsFeDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-es-left': isSynHdrEsFeDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-fe-right': isSynHdrEsFeDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-fj-fz-cell': isSynFjFzTableCellEntry(entry, colEntry.col),
                    'syn-hdr-edge-fjfz-right': isSynHdrFjFzDividerRightEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-fj-left': isSynHdrFjFzDividerLeftEntry(
                      entry,
                      colEntry.col
                    ),
                    'syn-hdr-edge-fz-right': isSynHdrFjFzDividerRightEdgeEntry(
                      entry,
                      colEntry.col
                    ),
                  },
                ]"
                :style="scrollDataCellStyle(entry, colEntry.col, colEntry.width)"
              >
                <template v-if="isGapEntry(entry) || cellReadonly(entry.excelRow, colEntry.col)">
                  <span>{{
                    isGapEntry(entry) ? '' : cellDisplay(entry.excelRow, colEntry.col)
                  }}</span>
                </template>
                <input
                  v-else-if="isCellActive(entry.excelRow, colEntry.col)"
                  type="text"
                  class="grid-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  :data-grid-row="entry.excelRow"
                  :data-grid-col="colEntry.col"
                  @mousedown="onCellMouseDown(entry.excelRow, colEntry.col, $event)"
                  @focus="onCellFocus(entry.excelRow, colEntry.col, $event)"
                  @input="onCellInput(entry.excelRow, colEntry.col, $event)"
                  @blur="onCellBlur(entry.excelRow, colEntry.col, $event)"
                  @keydown="onCellKeydown(entry.excelRow, colEntry.col, $event)"
                />
                <span
                  v-else
                  class="grid-cell-display"
                  @mousedown="onCellSpanMouseDown(entry.excelRow, colEntry.col, $event)"
                >{{ cellShowValue(entry.excelRow, colEntry.col, cellDisplay(entry.excelRow, colEntry.col)) }}</span>
              </td>
              <td v-if="rightPad > 0" class="syn-pad" :style="{ width: rightPad + 'px', minWidth: rightPad + 'px' }"></td>
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
