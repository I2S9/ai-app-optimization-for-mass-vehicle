import { ref, computed, onMounted, onUnmounted } from 'vue';
import {
  buildCellMap,
  buildWidthMap,
  getCell,
  displayValue,
  isReadonlyCell,
  isAddBlueRow,
} from './bdStore.js';

const ROW_H = 21;
const BUFFER = 8;

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
    const search = ref('');

    const cellMap = computed(() => buildCellMap(props.sheet.cells));
    const widthMap = computed(() =>
      buildWidthMap(props.sheet.colWidths, props.sheet.columns)
    );

    const dataStart = computed(() => props.sheet.dataStartRow);
    const lastRow = computed(() => props.sheet.lastRow);
    const columns = computed(() => props.sheet.columns);

    const dataRowCount = computed(() => lastRow.value - dataStart.value + 1);
    const totalBodyHeight = computed(() => dataRowCount.value * ROW_H);

    const visibleStart = computed(() => {
      const start = Math.floor(scrollTop.value / ROW_H) - BUFFER;
      return Math.max(0, start);
    });

    const visibleEnd = computed(() => {
      const count = Math.ceil(viewportH.value / ROW_H) + BUFFER * 2;
      return Math.min(dataRowCount.value, visibleStart.value + count);
    });

    const visibleRows = computed(() => {
      const rows = [];
      for (let i = visibleStart.value; i < visibleEnd.value; i++) {
        rows.push(dataStart.value + i);
      }
      return rows;
    });

    const topSpacer = computed(() => visibleStart.value * ROW_H);
    const bottomSpacer = computed(() => {
      const shown = visibleEnd.value - visibleStart.value;
      const remaining = dataRowCount.value - visibleStart.value - shown;
      return Math.max(0, remaining * ROW_H);
    });

    function colStyle(col) {
      const w = widthMap.value.get(col) || 72;
      return { width: `${w}px`, minWidth: `${w}px`, maxWidth: `${w}px` };
    }

    function headerCell(row, col) {
      const hr = props.sheet.headerRows?.[row];
      if (!hr?.[col]) {
        if (row === 1) return props.sheet.headers[col] || '';
        return '';
      }
      const c = hr[col];
      return c.v ?? c.f ?? '';
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
      search,
      columns,
      dataStart,
      visibleRows,
      topSpacer,
      bottomSpacer,
      totalBodyHeight,
      colStyle,
      headerCell,
      getCell: (r, c) => getCell(cellMap.value, r, c),
      displayValue,
      isReadonlyCell: (cell, row) => isReadonlyCell(cell, row, dataStart.value),
      isAddBlueRow: (row) => isAddBlueRow(cellMap.value, row),
      onScroll,
      onCellInput,
      ROW_H,
    };
  },
  template: `
    <div class="bd-grid-root">
      <div class="bd-grid-toolbar">
        <strong>Sheet: BD</strong>
        <span>Rows {{ dataStart }}–{{ sheet.lastRow }} ({{ sheet.cells.length }} populated cells)</span>
        <label>
          <span>Go to row</span>
          <input type="number" :min="dataStart" :max="sheet.lastRow" style="width:72px" />
        </label>
      </div>
      <div class="bd-grid-scroll" ref="scrollEl" @scroll="onScroll">
        <table class="bd-table" role="grid">
          <thead>
            <tr class="hdr-row-1">
              <th class="corner"></th>
              <th
                v-for="col in columns"
                :key="'h1-' + col"
                class="col-hdr"
                :style="colStyle(col)"
              >{{ sheet.headers[col] || col }}</th>
            </tr>
            <tr v-for="hr in [2,3,4,5]" :key="'hr-' + hr" :class="'hdr-row-' + hr">
              <th class="corner">{{ hr }}</th>
              <th
                v-for="col in columns"
                :key="hr + '-' + col"
                :style="colStyle(col)"
              >{{ headerCell(hr, col) }}</th>
            </tr>
          </thead>
          <tbody>
            <tr v-if="topSpacer > 0" class="bd-spacer-top">
              <td :colspan="columns.length + 1" :style="{ height: topSpacer + 'px', padding: 0, border: 'none' }"></td>
            </tr>
            <tr
              v-for="row in visibleRows"
              :key="row"
              :class="{ 'row-addblue': isAddBlueRow(row) }"
            >
              <td class="row-num">{{ row }}</td>
              <td
                v-for="col in columns"
                :key="row + '-' + col"
                class="data-cell"
                :class="{ readonly: isReadonlyCell(getCell(row, col), row) }"
                :style="colStyle(col)"
              >
                <template v-if="isReadonlyCell(getCell(row, col), row)">
                  <span :title="getCell(row, col)?.f || ''">{{ displayValue(getCell(row, col)) }}</span>
                </template>
                <input
                  v-else
                  type="text"
                  :value="displayValue(getCell(row, col))"
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
