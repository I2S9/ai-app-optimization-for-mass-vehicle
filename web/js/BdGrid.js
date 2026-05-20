import { ref, computed, onMounted, onUnmounted, watch } from 'vue';
import {
  buildCellMap,
  buildWidthMap,
  computeBodyDisplayRows,
  getCell,
  displayValue,
  displayCellValue,
  isReadonlyCell,
  isStructureRow,
  isTitleMarkerRow,
  rowStyleClass,
  cellInlineStyle,
  projectCellClass,
  shouldDisplayBodyRow,
} from './bdStore.js?v=calc-syn6';
const ROW_H = 21;
const BUFFER = 12;
export default {
  name: 'BdGrid',
  props: {
    sheet: { type: Object, required: true },
    sheetName: { type: String, default: 'BD' },
    session: { type: Object, default: null },
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
      if (props.sheetName === 'SYNTHESIS') {
        const last = props.sheet.lastRow || 1;
        const rows = [];
        for (let r = 1; r <= last; r++) {
          rows.push({ excelRow: r, displayRow: r });
        }
        return rows;
      }
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
    const calcRevision = computed(
      () => props.session?.revision?.value ?? 0
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
      if (props.session?.ready?.value) {
        props.session.setCellValue(props.sheetName, row, col, value);
      }
      emit('cell-change', { row, col, value });
    }
    function cellReadonly(row, col) {
      const cell = getCell(cellMap.value, row, col);
      if (props.session?.ready?.value) {
        return props.session.isFormulaCell(
          props.sheetName,
          row,
          col,
          cell
        );
      }
      return isReadonlyCell(cell, row, props.sheet.dataStartRow);
    }
    function cellDisplay(row, col) {
      void calcRevision.value;
      const map = cellMap.value;
      const sh = sectionHeaderRows.value;
      const canon = props.sheet.canonicalSectionByLabel;
      const cell = getCell(map, row, col);

      if (props.sheetName === 'BD') {
        const masked = displayCellValue(map, row, col, sh, canon);
        const bookmark =
          isStructureRow(map, row, sh) || isTitleMarkerRow(map, row, sh);
        const useFormula =
          !bookmark &&
          props.session?.ready?.value &&
          props.session.isFormulaCell(props.sheetName, row, col, cell);
        if (useFormula) {
          const computed = props.session.getDisplayValue(
            props.sheetName,
            row,
            col,
            cell
          );
          if (computed !== '' && computed !== '#REF!') return computed;
        }
        return masked;
      }

      if (props.session?.ready?.value) {
        return props.session.getDisplayValue(
          props.sheetName,
          row,
          col,
          cell
        );
      }
      return displayValue(cell);
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
      if (props.sheetName === 'SYNTHESIS' && props.session?.bindSynthesisGrid) {
        props.session.bindSynthesisGrid((row, col) =>
          getCell(cellMap.value, row, col)
        );
      }
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
      calcRevision,
      colStyle,
      isStickyDateCol,
      getCell: (r, c) => getCell(cellMap.value, r, c),
      displayValue,
      cellDisplay,
      cellReadonly,
      rowStyleClass: (row) =>
        rowStyleClass(cellMap.value, row, sectionHeaderRows.value),
      projectCellClass,
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
                :class="[
                  {
                    readonly: cellReadonly(entry.excelRow, col),
                    'col-sticky-date': isStickyDateCol(col),
                  },
                  projectCellClass(cellDisplay(entry.excelRow, col), col),
                ]"
                :style="[colStyle(col), cellInlineStyle(entry.excelRow, col)]"
              >
                <template v-if="cellReadonly(entry.excelRow, col)">
                  <span :title="getCell(entry.excelRow, col)?.f || ''">{{ cellDisplay(entry.excelRow, col) }}</span>
                </template>
                <input
                  v-else
                  type="text"
                  :value="cellDisplay(entry.excelRow, col)"
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
