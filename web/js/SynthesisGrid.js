/**
 * Synthesis grid — Excel A–E hidden; display letters F→A, G→B, …
 */
import { ref, computed, shallowRef, onMounted, onUnmounted, watch } from 'vue';
import { getCell } from './bdStore.js?v=syn-perf5';
import {
  computeSynBodyRows,
  synDisplayValue,
  synIsReadonly,
  synProjectCellClass,
  synRowStyleClass,
} from './synStore.js?v=syn-perf5';
import {
  SYN_STICKY_COL,
  excelToDisplayCol,
} from './synthesisPerf.js?v=syn-perf5';

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

    const displayColumns = computed(() => props.sheet.columns || []);

    const widthByCol = computed(() => {
      const m = new Map();
      for (const w of props.sheet.colWidths || []) {
        m.set(w.col, Math.min(w.width || 64, 100));
      }
      for (const col of displayColumns.value) {
        if (!m.has(col)) {
          m.set(col, col === SYN_STICKY_COL ? 160 : 54);
        }
      }
      m.set(SYN_STICKY_COL, Math.max(m.get(SYN_STICKY_COL) || 160, 160));
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

    function cellReadonly(row, col) {
      const cell = getCell(cellMap.value, row, col);
      return Boolean(cell?.f) || synIsReadonly(cell, row, props.sheet);
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
      const cell = getCell(cellMap.value, row, col);
      if (!cell) return '';
      if (cell.v != null && cell.v !== '') return formatVal(cell.v);
      if (cell.f === 'SUMPRODUCT') return '0';
      return formatVal(synDisplayValue(cell, cellMap.value, row, col));
    }

    function rowClass(row) {
      return synRowStyleClass(cellMap.value, row, props.sheet);
    }

    function cellExtraClass(row, col, display) {
      const rc = rowClass(row);
      if (rc === 'syn-row-section' || rc === 'syn-row-subsection') return '';
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
      rowClass,
      cellExtraClass,
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
            <tr class="hdr-row-1">
              <th class="corner syn-row-num-hdr">1</th>
              <th
                v-for="entry in pinnedCols"
                :key="'H-' + entry.col"
                class="col-hdr col-sticky-label"
                :style="[colStyle(entry.col, entry.width), stickyStyle(entry.col)]"
                :title="sheet.headers[entry.col] || entry.letter"
              >{{ sheet.headers[entry.col] || entry.letter }}</th>
              <th v-if="leftPad > 0" class="syn-pad" :style="{ width: leftPad + 'px', minWidth: leftPad + 'px' }"></th>
              <th
                v-for="entry in visibleScrollCols"
                :key="'H-' + entry.col"
                class="col-hdr"
                :style="colStyle(entry.col, entry.width)"
                :title="sheet.headers[entry.col] || entry.letter"
              >{{ sheet.headers[entry.col] || entry.letter }}</th>
              <th v-if="rightPad > 0" class="syn-pad" :style="{ width: rightPad + 'px', minWidth: rightPad + 'px' }"></th>
            </tr>
          </thead>
          <tbody>
            <tr v-if="topSpacer > 0" class="bd-spacer-top">
              <td :colspan="colspan" :style="{ height: topSpacer + 'px', border: 'none', padding: 0 }"></td>
            </tr>
            <tr
              v-for="entry in visibleRows"
              :key="entry.excelRow"
              :class="rowClass(entry.excelRow)"
            >
              <td class="row-num syn-row-num">{{ entry.displayRow }}</td>
              <td
                v-for="p in pinnedCols"
                :key="entry.excelRow + '-p-' + p.col"
                class="data-cell col-sticky-label syn-label-col"
                :class="[
                  { readonly: cellReadonly(entry.excelRow, p.col) },
                  cellExtraClass(entry.excelRow, p.col, cellDisplay(entry.excelRow, p.col)),
                ]"
                :style="[colStyle(p.col, p.width), stickyStyle(p.col)]"
              >
                <template v-if="cellReadonly(entry.excelRow, p.col)">
                  <span>{{ cellDisplay(entry.excelRow, p.col) }}</span>
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
                v-for="colEntry in visibleScrollCols"
                :key="entry.excelRow + '-' + colEntry.col"
                class="data-cell"
                :class="[
                  { readonly: cellReadonly(entry.excelRow, colEntry.col) },
                  cellExtraClass(entry.excelRow, colEntry.col, cellDisplay(entry.excelRow, colEntry.col)),
                ]"
                :style="colStyle(colEntry.col, colEntry.width)"
              >
                <template v-if="cellReadonly(entry.excelRow, colEntry.col)">
                  <span>{{ cellDisplay(entry.excelRow, colEntry.col) }}</span>
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
