/**
 * Synthesis grid — Excel A–E hidden; display letters F→A, G→B, …
 */
import { ref, computed, shallowRef, onMounted, onUnmounted, watch } from 'vue';
import { getCell } from './bdStore.js?v=syn-perf30';
import {
  computeSynBodyRows,
  synDisplayValue,
  buildSynPillarColumns,
  synCellInlineStyle,
  synProjectCellClass,
  synMetricCellClass,
  synHeaderPanelVehicleClass,
  synRowStyleClass,
  isSynPillarCol,
  isSynPillarColAtRow,
  synPillarLetterForRow,
  isSynMetricRow,
  isSynHeaderPanelRow,
} from './synStore.js?v=syn-perf30';
import {
  SYN_STICKY_COL,
  excelToDisplayCol,
  synStickyColWidth,
  synPillarColWidth,
} from './synthesisPerf.js?v=syn-perf35';
const ROW_H = 21;
const ROW_NUM_W = 56;
const BUFFER_ROWS = 6;
const BUFFER_PX = 280;

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

    const visibleScrollCols = computed(() => {
      const min = scrollLeft.value - BUFFER_PX;
      const max = scrollLeft.value + viewportW.value + BUFFER_PX;
      return scrollableCols.value.filter(
        (c) => c.left + c.width >= min && c.left <= max
      );
    });

    const leftPad = computed(() => {
      const first = visibleScrollCols.value[0];
      const pin = pinnedCols.value[0];
      if (!first || !pin) return 0;
      return Math.max(0, first.left - (pin.left + pin.width));
    });

    const rightPad = computed(() => {
      const last = visibleScrollCols.value[visibleScrollCols.value.length - 1];
      if (!last) return 0;
      return Math.max(0, tableWidth.value - (last.left + last.width));
    });

    const bodyRows = computed(() =>
      computeSynBodyRows(props.sheet, cellMap.value, props.outlineOnly)
    );
    const rowCount = computed(() => bodyRows.value.length);

    const visibleStart = computed(() =>
      Math.max(0, Math.floor(scrollTop.value / ROW_H) - BUFFER_ROWS)
    );
    const visibleEnd = computed(() => {
      const n = Math.ceil(viewportH.value / ROW_H) + BUFFER_ROWS * 2;
      return Math.min(rowCount.value, visibleStart.value + n);
    });
    const visibleRows = computed(() =>
      bodyRows.value.slice(visibleStart.value, visibleEnd.value)
    );
    const topSpacer = computed(() => visibleStart.value * ROW_H);
    const bottomSpacer = computed(() => {
      const shown = visibleEnd.value - visibleStart.value;
      return Math.max(0, (rowCount.value - visibleStart.value - shown) * ROW_H);
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

    function onScroll(e) {
      scrollTop.value = e.target.scrollTop;
      scrollLeft.value = e.target.scrollLeft;
    }

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
      const s = String(v);
      if (!/^-?\d+(\.\d+)?$/.test(s.trim())) return s;
      const n = parseFloat(s);
      if (!Number.isFinite(n)) return s;
      if (Number.isInteger(n)) return String(n);
      return String(Math.round(n * 1000) / 1000);
    }

    function cellDisplay(row, col) {
      if (row == null) return '';
      if (isSynPillarColAtRow(col, row, pillarColumns.value)) {
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
      if (entry.gapAfterPanel) {
        const gap = ['syn-panel-gap-row', 'syn-panel-gap', 'syn-header-spacer-white'];
        if (entry.gapIndex === 1) gap.push('syn-panel-gap-first');
        if (entry.gapIndex === 2) gap.push('syn-panel-gap-last');
        return gap;
      }
      const row = entry.excelRow;
      const cls = synRowStyleClass(cellMap.value, row, props.sheet);
      const list = [cls];
      if (isSynHeaderPanelRow(row)) list.push('syn-header-block');
      if (row >= 3 && row <= 14) list.push('syn-filter-band');
      if (row >= 15 && row <= 22) list.push('syn-metric-band');
      if (row === 16 || row === 18) list.push('syn-metric-curb');
      if (row === 10 || row === 15) list.push('syn-header-spacer-row', 'syn-header-spacer-white');
      if (row === 9) list.push('syn-header-edge-below-spec');
      if (row === 11) list.push('syn-header-edge-above-pole');
      if (row === 14) list.push('syn-header-edge-below-finition');
      if (row === 16) list.push('syn-header-edge-above-curb');
      if (row === 22) list.push('syn-header-panel-end');
      const dr = entry.displayRow;
      if (dr >= 1 && dr <= 6) list.push('syn-header-edge-sep');
      if (dr === 9 || dr === 10 || dr === 11) list.push('syn-header-edge-sep');
      return list;
    }

    function isGapEntry(entry) {
      return Boolean(entry.gapAfterPanel);
    }

    function isPillarColForEntry(entry, col) {
      if (isGapEntry(entry)) return false;
      return isSynPillarColAtRow(col, entry.excelRow, pillarColumns.value);
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

    function cellExtraClass(row, col, display) {
      if (isSynPillarColAtRow(col, row, pillarColumns.value)) {
        return display ? 'syn-pillar-has-char' : '';
      }
      if (isSynHeaderPanelRow(row)) {
        const hdrCls = synHeaderPanelVehicleClass(row, col, display);
        if (hdrCls) return hdrCls;
      }
      const rc = synRowStyleClass(cellMap.value, row, props.sheet);
      if (
        rc === 'syn-row-section' ||
        rc === 'syn-row-subsection' ||
        rc === 'syn-row-separator'
      ) {
        return '';
      }
      if (isSynMetricRow(row)) {
        return synMetricCellClass(row, col, display) || '';
      }
      return synProjectCellClass(display, col);
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
      window.addEventListener('resize', updateViewport);
    });
    onUnmounted(() => {
      window.removeEventListener('resize', updateViewport);
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
      pillarTitle,
      isSynPillarCol: (col) => isSynPillarCol(col, pillarColumns.value),
      isSynPillarColAtRow: (col, row) =>
        isSynPillarColAtRow(col, row, pillarColumns.value),
      cellInlineStyle,
      cellExtraClass,
      headerEdgeRight,
      headerLabelEdgeRight,
      onScroll,
      onCellInput,
    };
  },
  template: `
    <div class="bd-grid-root synthesis-grid">
      <div class="bd-grid-scroll syn-scroll" ref="scrollEl" @scroll.passive="onScroll">
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
              :key="entry.gapAfterPanel ? 'panel-gap-' + entry.gapIndex : entry.excelRow"
              :class="entryRowClasses(entry)"
            >
              <td class="row-num syn-row-num">{{ entry.displayRow }}</td>
              <td
                v-for="p in pinnedCols"
                :key="(entry.gapAfterPanel ? 'gap' : entry.excelRow) + '-p-' + p.col"
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
                :key="(entry.gapAfterPanel ? 'gap' : entry.excelRow) + '-' + colEntry.col"
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
                      ? isSynPillarCol(colEntry.col)
                      : isPillarColForEntry(entry, colEntry.col),
                    'syn-header-edge-right':
                      !isGapEntry(entry) &&
                      headerEdgeRight(entry.excelRow, colIdx, visibleScrollCols.length),
                  },
                ]"
                :style="[
                  colStyle(colEntry.col, colEntry.width),
                  isGapEntry(entry) ? {} : cellInlineStyle(entry.excelRow, colEntry.col),
                ]"
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
  `,
};
