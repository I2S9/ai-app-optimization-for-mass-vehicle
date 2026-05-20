import { ref, computed, onMounted, onUnmounted } from 'vue';
import {
  buildCellMap,
  buildWidthMap,
  getCell,
  displayValue,
  displayCellValue,
  isReadonlyCell,
  rowStyleClass,
  cellInlineStyle,
} from './bdStore.js?v=20260520-4';

const ROW_H = 21;
const BUFFER = 12;
const DISPLAY_START = 2;

export default {
  name: 'BdGrid',
  props: {
    sheet: { type: Object, required: true },
  },
  emits: ['cell-change'],
  setup(props, { emit }) {
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const viewportH = ref(600);

    const cellMap = computed(() =>
      buildCellMap(props.sheet.cells, props.sheet.headerRows)
    );
    const widthMap = computed(() =>
      buildWidthMap(props.sheet.colWidths, props.sheet.columns, props.sheet.headers)
    );

    const displayStart = DISPLAY_START;
    const lastRow = computed(() => props.sheet.lastRow);
    const columns = computed(() => props.sheet.columns);

    const rowCount = computed(() => lastRow.value - displayStart + 1);

    const visibleStart = computed(() => {
      const start = Math.floor(scrollTop.value / ROW_H) - BUFFER;
      return Math.max(0, start);
    });

    const visibleEnd = computed(() => {
      const count = Math.ceil(viewportH.value / ROW_H) + BUFFER * 2;
      return Math.min(rowCount.value, visibleStart.value + count);
    });

    const visibleRows = computed(() => {
      const rows = [];
      for (let i = visibleStart.value; i < visibleEnd.value; i++) {
        rows.push(displayStart + i);
      }
      return rows;
    });

    const topSpacer = computed(() => visibleStart.value * ROW_H);
    const bottomSpacer = computed(() => {
      const shown = visibleEnd.value - visibleStart.value;
      const remaining = rowCount.value - visibleStart.value - shown;
      return Math.max(0, remaining * ROW_H);
    });

    function colStyle(col) {
      const w = widthMap.value.get(col) || 72;
      return { width: `${w}px`, minWidth: `${w}px`, maxWidth: `${w}px` };
    }

    function onScroll(e) {
      scrollTop.value = e.target.scrollTop;
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
      emit('cell-change', { row, col, value });
    }

    function updateViewport() {
      if (scrollEl.value) viewportH.value = scrollEl.value.clientHeight;
    }

    onMounted(() => {
      updateViewport();
      window.addEventListener('resize', updateViewport);
    });
    onUnmounted(() => {
      window.removeEventListener('resize', updateViewport);
    });

    return {
      scrollEl,
      columns,
      displayStart,
      visibleRows,
      topSpacer,
      bottomSpacer,
      colStyle,
      getCell: (r, c) => getCell(cellMap.value, r, c),
      displayValue,
      displayCellValue: (r, c) => displayCellValue(cellMap.value, r, c),
      isReadonlyCell: (cell, row) => isReadonlyCell(cell, row, props.sheet.dataStartRow),
      rowStyleClass: (row) => rowStyleClass(cellMap.value, row),
      cellInlineStyle: (row, col) =>
        cellInlineStyle(getCell(cellMap.value, row, col), cellMap.value, row, col),
      onScroll,
      onCellInput,
    };
  },
  template: `
    <div class="bd-grid-root">
      <div class="bd-grid-scroll" ref="scrollEl" @scroll="onScroll">
        <table class="bd-table" role="grid">
          <thead>
            <tr class="hdr-row-letters">
              <th class="corner"></th>
              <th
                v-for="col in columns"
                :key="'letter-' + col"
                class="col-letter"
                :style="colStyle(col)"
              >{{ col }}</th>
            </tr>
            <tr class="hdr-row-1">
              <th class="corner">1</th>
              <th
                v-for="col in columns"
                :key="'h1-' + col"
                class="col-hdr"
                :style="colStyle(col)"
                :title="sheet.headers[col] || col"
              >{{ sheet.headers[col] || col }}</th>
            </tr>
          </thead>
          <tbody>
            <tr v-if="topSpacer > 0" class="bd-spacer-top">
              <td :colspan="columns.length + 1" :style="{ height: topSpacer + 'px', padding: 0, border: 'none' }"></td>
            </tr>
            <tr
              v-for="row in visibleRows"
              :key="row"
              :class="rowStyleClass(row)"
            >
              <td class="row-num">{{ row }}</td>
              <td
                v-for="col in columns"
                :key="row + '-' + col"
                class="data-cell"
                :class="{ readonly: isReadonlyCell(getCell(row, col), row) }"
                :style="[colStyle(col), cellInlineStyle(row, col)]"
              >
                <template v-if="isReadonlyCell(getCell(row, col), row)">
                  <span :title="getCell(row, col)?.f || ''">{{ displayCellValue(row, col) }}</span>
                </template>
                <input
                  v-else
                  type="text"
                  :value="displayCellValue(row, col)"
                  @change="onCellInput(row, col, $event.target.value)"
                />
              </td>
            </tr>
            <tr v-if="bottomSpacer > 0" class="bd-spacer-bottom">
              <td :colspan="columns.length + 1" :style="{ height: bottomSpacer + 'px', padding: 0, border: 'none' }"></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `,
};
