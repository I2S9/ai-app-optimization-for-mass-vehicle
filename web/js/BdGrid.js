import { ref, computed, onMounted, onUnmounted, watch } from '/vendor/vue.esm-browser.js';
import {
  buildCellMap,
  buildWidthMap,
  computeBodyDisplayRows,
  getCell,
  displayValue,
  displayCellValue,
  isReadonlyCell,
  rowStyleClass,
  cellInlineStyle,
  shouldDisplayBodyRow,
} from './bdStore.js?v=20260520-edge3';
const ROW_H = 21;
const BUFFER = 12;
export default {
  name: 'BdGrid',
  props: {
    sheet: { type: Object, required: true },
    outlineOnly: { type: Boolean, default: false },
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
    const bodyRows = computed(() => {
      if (props.outlineOnly) {
        const map = cellMap.value;
        let displayRow = 1;
        return (props.sheet.outlineRows || [])
          .filter((r) => shouldDisplayBodyRow(map, r, props.sheet))
          .map((excelRow) => ({ excelRow, displayRow: displayRow++ }));
      }
      return computeBodyDisplayRows(props.sheet);
    });
    const rowCount = computed(() => bodyRows.value.length);
    const visibleStart = computed(() => {
      const start = Math.floor(scrollTop.value / ROW_H) - BUFFER;
      return Math.max(0, start);
    });
    const visibleEnd = computed(() => {
      const count = Math.ceil(viewportH.value / ROW_H) + BUFFER * 2;
      return Math.min(rowCount.value, visibleStart.value + count);
    });
    const visibleRows = computed(() =>
      bodyRows.value.slice(visibleStart.value, visibleEnd.value)
    );
    const topSpacer = computed(() => visibleStart.value * ROW_H);
    const bottomSpacer = computed(() => {
      const shown = visibleEnd.value - visibleStart.value;
      const remaining = rowCount.value - visibleStart.value - shown;
      return Math.max(0, remaining * ROW_H);
    });
    const columns = computed(() => props.sheet.columns);
    const sectionHeaderRows = computed(
      () => props.sheet.sectionHeaderRows
    );
    function colStyle(col) {
      const w = widthMap.value.get(col) || 72;
      return { width: `${w}px`, minWidth: `${w}px`, maxWidth: `${w}px` };
    }
    function isStickyDateCol(col) {
      return col === 'A';
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
    watch(
      () => props.outlineOnly,
      () => {
        scrollTop.value = 0;
        if (scrollEl.value) scrollEl.value.scrollTop = 0;
      }
    );
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
      visibleRows,
      topSpacer,
      bottomSpacer,
      colStyle,
      isStickyDateCol,
      getCell: (r, c) => getCell(cellMap.value, r, c),
      displayValue,
      displayCellValue: (r, c) =>
        displayCellValue(
          cellMap.value,
          r,
          c,
          sectionHeaderRows.value,
          props.sheet.canonicalSectionByLabel
        ),
      isReadonlyCell: (cell, row) =>
        isReadonlyCell(cell, row, props.sheet.dataStartRow),
      rowStyleClass: (row) =>
        rowStyleClass(cellMap.value, row, sectionHeaderRows.value),
      cellInlineStyle: (row, col) =>
        cellInlineStyle(
          getCell(cellMap.value, row, col),
          cellMap.value,
          row,
          col,
          sectionHeaderRows.value
        ),
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
                :class="{ 'col-sticky-date': isStickyDateCol(col) }"
                :style="colStyle(col)"
              >{{ col }}</th>
            </tr>
            <tr class="hdr-row-1">
              <th class="corner">1</th>
              <th
                v-for="col in columns"
                :key="'h1-' + col"
                class="col-hdr"
                :class="{ 'col-sticky-date': isStickyDateCol(col) }"
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
              v-for="entry in visibleRows"
              :key="entry.excelRow"
              :class="rowStyleClass(entry.excelRow)"
            >
              <td class="row-num">{{ entry.displayRow }}</td>
              <td
                v-for="col in columns"
                :key="entry.excelRow + '-' + col"
                class="data-cell"
                :class="{
                  readonly: isReadonlyCell(getCell(entry.excelRow, col), entry.excelRow),
                  'col-sticky-date': isStickyDateCol(col),
                }"
                :style="[colStyle(col), cellInlineStyle(entry.excelRow, col)]"
              >
                <template v-if="isReadonlyCell(getCell(entry.excelRow, col), entry.excelRow)">
                  <span :title="getCell(entry.excelRow, col)?.f || ''">{{ displayCellValue(entry.excelRow, col) }}</span>
                </template>
                <input
                  v-else
                  type="text"
                  :value="displayCellValue(entry.excelRow, col)"
                  @change="onCellInput(entry.excelRow, col, $event.target.value)"
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
