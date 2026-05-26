/**
 * Synthesis grid — Excel A–E hidden; display letters F→A, G→B, …
 */
import { ref, computed, shallowRef, onMounted, onUnmounted, watch } from 'vue';
import { getCell } from './bdStore.js?v=syn-perf30';
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
  isSynSp2DisplayExcelCol,
  isSynProjHeaderGreenCol,
  synProjHeaderGreenStyle,
  isSynHdrLmDividerRightCol,
  isSynHdrLmDividerLeftCol,
  isSynHdrAaDividerRightCol,
  isSynHdrCjDividerRightEntry,
  isSynHdrLmDividerRightEntry,
  isSynHdrLmDividerLeftEntry,
  isSynHdrAaDividerRightEntry,
  isSynSpacerDisplayExcelCol,
  synSpacerColClass,
  SYN_GRID_FIRST_ROW,
} from './synStore.js?v=syn-perf66';
import {
  SYN_STICKY_COL,
  excelToDisplayCol,
  synStickyColWidth,
  synPillarColWidth,
} from './synthesisPerf.js?v=syn-perf37';
import {
  ROW_H,
  visibleRowRange,
  rowOverscan,
  colOverscanPx,
  shouldVirtualizeRows,
  shouldVirtualizeCols,
  createScrollRafSync,
} from './gridScroll.js?v=syn-scroll3';
const SYN_HEAD_ROW_H = 22;
const ROW_NUM_W = 56;

export default {
  name: 'SynthesisGrid',
  props: {
    sheet: { type: Object, required: true },
    session: { type: Object, default: null },
    outlineOnly: { type: Boolean, default: false },
  },
  emits: ['cell-change'],
  setup(props, { emit }) {
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const scrollLeft = ref(0);
    const viewportH = ref(600);
    const viewportW = ref(1000);
    const cellMap = shallowRef(props.sheet?.cellMap || new Map());

    const pillarColumns = computed(() => {
      const raw = props.sheet?.pillarColumns;
      if (raw && Object.keys(raw).length) {
        return new Map(Object.entries(raw).map(([col, meta]) => [col, meta]));
      }
      return buildSynPillarColumns(props.sheet, cellMap.value);
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
    const rowBuffer = computed(() => rowOverscan(viewportH.value));
    const visibleRange = computed(() => {
      if (!virtualizeRows.value) {
        return { start: 0, end: rowCount.value };
      }
      return visibleRowRange(
        scrollTop.value,
        viewportH.value,
        rowCount.value,
        rowBuffer.value
      );
    });
    const visibleStart = computed(() => visibleRange.value.start);
    const visibleEnd = computed(() => visibleRange.value.end);
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
      return SYN_PILLAR_OVERLAY_COLS.has(col) && pillarColumns.value.has(col);
    }

    const pillarLetterOverlays = computed(() => {
      const start = findSynEchappementRow(cellMap.value, props.sheet);
      const overlays = [];
      for (const layout of columnLayout.value) {
        if (!SYN_PILLAR_OVERLAY_COLS.has(layout.col)) continue;
        if (!pillarColumns.value.has(layout.col)) continue;
        const title = pillarColumns.value.get(layout.col)?.title ?? '';
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

    const scrollSync = createScrollRafSync({ scrollTop, scrollLeft });

    function onCellInput(row, col, value) {
      const key = `${row}:${col}`;
      let cell = cellMap.value.get(key);
      if (!cell) {
        cell = { r: row, c: col, v: value };
        props.sheet.cells.push(cell);
        cellMap.value.set(key, cell);
      } else {
        cell.v = value;
        delete cell.f;
      }
      props.session?.setCellValue?.('SYNTHESIS', row, col, value);
      emit('cell-change', { row, col, value });
    }

    /** Only virtual gap rows (no Excel row) stay display-only. */
    function cellReadonly(row, col) {
      return row == null;
    }

    function formatVal(v) {
      const formatted = formatSynNumericDisplay(v);
      return formatted === '' && v != null && String(v).trim() !== ''
        ? String(v)
        : formatted;
    }

    function cellDisplay(row, col) {
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
      if (props.session?.getDisplayValue) {
        const fromSession = props.session.getDisplayValue(
          'SYNTHESIS',
          row,
          col,
          cell
        );
        if (fromSession != null && fromSession !== '') {
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

    function entryRowClasses(entry) {
      if (isSynPanelGapEntry(entry)) {
        const gap = ['syn-panel-gap-row', 'syn-panel-gap'];
        if (entry.gapBeforePanel) gap.push('syn-panel-gap-top');
        if (entry.gapAfterPanel && entry.gapIndex === 1) gap.push('syn-panel-gap-first');
        if (entry.gapAfterPanel && entry.gapIndex === 2) gap.push('syn-panel-gap-last');
        return gap;
      }
      const row = entry.excelRow;
      const cls = synRowStyleClass(cellMap.value, row, props.sheet);
      const list = [cls];
      if (isSynHeaderPanelRow(row)) {
        list.push('syn-header-block', 'syn-proj-table-frame', 'syn-hdr-panel-grid');
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
      return isSynPillarColAtRow(col, entry.excelRow, pillarColumns.value);
    }

    /** Gap rows 21–22: display B (SP1 grey) & K (SP2 green) keep pillar fill. */
    function isGapGreenPillarCol(col) {
      const letter = excelToDisplayCol(col);
      return letter === 'B' || letter === 'K';
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
      if (isSynProjHeaderGreenCol(row, col)) {
        return { ...base, ...synProjHeaderGreenStyle() };
      }
      return { ...base, ...cellInlineStyle(row, col) };
    }

    function withHdrPanelBold(row, col, cls) {
      const bold = isSynHeaderPanelBoldCol(row, col) ? 'syn-hdr-panel-bold' : '';
      const parts = [cls, bold].filter(Boolean);
      return parts.join(' ');
    }

    function cellExtraClass(row, col, display) {
      if (isSynPillarColAtRow(col, row, pillarColumns.value)) {
        return display ? 'syn-pillar-has-char' : '';
      }
      const spacerCol = synSpacerColClass(col);
      if (spacerCol) return spacerCol;
      const spotBlue = synSpotBlueColClass(row, col);
      if (spotBlue) return withHdrPanelBold(row, col, spotBlue);
      const greyCol = synFilterGreyColClass(row, col);
      if (greyCol) return withHdrPanelBold(row, col, greyCol);
      const adaptCol = synAdaptBandColClass(row, col, pillarColumns.value);
      if (adaptCol) return withHdrPanelBold(row, col, adaptCol);
      const metricWhiteCol = synMetricCjWhiteColClass(row, col);
      if (metricWhiteCol) return withHdrPanelBold(row, col, metricWhiteCol);
      const accent = synCellAccentClass(display);
      if (accent) return withHdrPanelBold(row, col, accent);
      if (isSynHeaderPanelRow(row)) {
        const hdrCls = synHeaderPanelVehicleClass(row, col, display);
        const combined = withHdrPanelBold(row, col, hdrCls);
        if (combined) return combined;
      }
      if (isSynProjHeaderGreenCol(row, col)) {
        return withHdrPanelBold(row, col, 'syn-proj-hdr-green');
      }
      const rc = synRowStyleClass(cellMap.value, row, props.sheet);
      if (
        rc === 'syn-row-section' ||
        rc === 'syn-row-subsection' ||
        rc === 'syn-row-separator'
      ) {
        return withHdrPanelBold(row, col, '');
      }
      if (isSynMetricRow(row)) {
        return withHdrPanelBold(row, col, synMetricCellClass(row, col, display) || '');
      }
      return withHdrPanelBold(row, col, synProjectCellClass(display, col));
    }

    function updateViewport() {
      if (!scrollEl.value) return;
      viewportH.value = scrollEl.value.clientHeight;
      viewportW.value = scrollEl.value.clientWidth;
    }

    watch(
      () => props.outlineOnly,
      () => {
        scrollTop.value = 0;
        if (scrollEl.value) scrollEl.value.scrollTop = 0;
      }
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
      excelToDisplayCol,
      isSynPillarCol: (col) => isSynPillarCol(col, pillarColumns.value),
      isSynPillarColAtRow: (col, row) =>
        isSynPillarColAtRow(col, row, pillarColumns.value),
      isSynSp2DisplayExcelCol,
      isSynSpacerDisplayExcelCol,
      synSpacerColClass,
      isSynHdrLmDividerRightCol,
      isSynHdrLmDividerLeftCol,
      isSynHdrAaDividerRightCol,
      isSynHdrLmDividerRightEntry,
      isSynHdrLmDividerLeftEntry,
      isSynHdrAaDividerRightEntry,
      isSynHdrCjDividerRightEntry,
      isSynProjHeaderGreenCol,
      synProjHeaderGreenStyle,
      cellInlineStyle,
      scrollDataCellStyle,
      cellExtraClass,
      headerEdgeRight,
      headerLabelEdgeRight,
      pillarLetterOverlays,
      usesPillarLetterOverlay,
      onScroll: scrollSync.onScroll,
      onCellInput,
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
            :class="{ 'syn-pillar-k': isSynSp2DisplayExcelCol(pillar.col) }"
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
                :class="{ 'syn-spacer-col-l-hdr': isSynSpacerDisplayExcelCol(entry.col) }"
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
              :key="isGapEntry(entry) ? 'panel-gap-' + (entry.gapBeforePanel ? 'top-' : 'bot-') + entry.gapIndex : entry.excelRow"
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
                  v-else
                  type="text"
                  :value="cellDisplay(entry.excelRow, p.col)"
                  @change="onCellInput(entry.excelRow, p.col, $event.target.value)"
                />
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
                  {
                    'syn-pillar-col': isGapEntry(entry)
                      ? isGapGreenPillarCol(colEntry.col)
                      : isPillarColForEntry(entry, colEntry.col),
                    'syn-pillar-overlay-host':
                      !isGapEntry(entry) &&
                      usesPillarLetterOverlay(colEntry.col),
                    'syn-panel-gap-pillar':
                      isGapEntry(entry) && isGapGreenPillarCol(colEntry.col),
                    'syn-pillar-k': isSynSp2DisplayExcelCol(colEntry.col),
                    'syn-proj-hdr-green':
                      !isGapEntry(entry) &&
                      isSynProjHeaderGreenCol(entry.excelRow, colEntry.col),
                    'syn-header-edge-right':
                      !isGapEntry(entry) &&
                      headerEdgeRight(entry.excelRow, colIdx, visibleScrollCols.length),
                    'syn-spacer-col-l': isSynSpacerDisplayExcelCol(colEntry.col),
                    'syn-hdr-edge-cj-right': isSynHdrCjDividerRightEntry(
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
                  v-else
                  type="text"
                  :value="cellDisplay(entry.excelRow, colEntry.col)"
                  @change="onCellInput(entry.excelRow, colEntry.col, $event.target.value)"
                />
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
