import { ref, computed, shallowRef, onMounted, onUnmounted, watch } from 'vue';
import {
  buildCellMap,
  buildWidthMap,
  bdSubsystemL1Col,
  bdSubsystemL2Col,
  bdDesignDeptCol,
  computeBodyDisplayRows,
  BD_BODY_DISPLAY_ROW_START,
  getCell,
  displayValue,
  displayCellValue,
  isReadonlyCell,
  isStructureRow,
  isTitleMarkerRow,
  rowStyleClass,
  rowDataStripeClass,
  cellInlineStyle,
  projectCellClass,
  bdColMetaClass,
  bdMassCol,
  bdTitleCol,
  shouldDisplayBodyRow,
} from './bdStore.js?v=20260521-row-num-blue';
import { ROW_H, visibleRowRange } from './gridScroll.js?v=syn-scroll2';
import {
  BD_FREE_FIELD_COL,
  BD_MASS_AV_AR_COLS,
  BD_POSITION_COLS,
} from './bdColumnConfig.js';

const ROW_OVERSCAN = 8;

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

    const cellMap = shallowRef(new Map());
    watch(
      () => props.sheet,
      (sheet) => {
        cellMap.value =
          sheet?.cellMap instanceof Map
            ? sheet.cellMap
            : buildCellMap(sheet?.cells, sheet?.headerRows);
      },
      { immediate: true }
    );

    const widthMap = computed(() =>
      buildWidthMap(
        props.sheet.colWidths,
        props.sheet.columns,
        props.sheet.headers,
        props.sheet.cells,
        props.sheet.freeFieldWidth
      )
    );

    const columns = computed(() => props.sheet.columns || []);

    const bodyRows = computed(() => {
      if (props.outlineOnly) {
        const map = cellMap.value;
        let displayRow = BD_BODY_DISPLAY_ROW_START;
        return (props.sheet.outlineRows || [])
          .filter((r) => shouldDisplayBodyRow(map, r, props.sheet))
          .map((excelRow) => ({ excelRow, displayRow: displayRow++ }));
      }
      return props.sheet.bodyDisplayRows ?? computeBodyDisplayRows(props.sheet);
    });

    const rowCount = computed(() => bodyRows.value.length);
    const visibleRange = computed(() =>
      visibleRowRange(
        scrollTop.value,
        viewportH.value,
        rowCount.value,
        ROW_OVERSCAN
      )
    );
    const visibleStart = computed(() => visibleRange.value.start);
    const visibleEnd = computed(() => visibleRange.value.end);
    const visibleRows = computed(() =>
      bodyRows.value.slice(visibleStart.value, visibleEnd.value)
    );
    const topSpacer = computed(() => visibleStart.value * ROW_H);
    const bottomSpacer = computed(() => {
      const shown = visibleEnd.value - visibleStart.value;
      const remaining = rowCount.value - visibleStart.value - shown;
      return Math.max(0, remaining * ROW_H);
    });

    const sectionHeaderRows = computed(() => props.sheet.sectionHeaderRows);
    const subsystemL1Col = computed(() => bdSubsystemL1Col(props.sheet));
    const subsystemL2Col = computed(() => bdSubsystemL2Col(props.sheet));
    const designDeptCol = computed(() => bdDesignDeptCol(props.sheet));
    const titleCol = computed(() => bdTitleCol(props.sheet));
    const massCol = computed(() => bdMassCol(props.sheet));
    const calcRevision = computed(() => props.session?.revision?.value ?? 0);
    const engineReady = computed(() => props.session?.ready?.value ?? false);
    const displayCache = new Map();
    let displayCacheKey = '';

    function invalidateDisplayCache() {
      displayCache.clear();
      displayCacheKey = '';
    }

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
      const cacheKey = `${calcRevision.value}:${engineReady.value}:${props.outlineOnly}:${props.sheet === null ? 0 : 1}`;
      if (cacheKey !== displayCacheKey) {
        invalidateDisplayCache();
        displayCacheKey = cacheKey;
      }
      const hitKey = `${row}:${col}`;
      if (displayCache.has(hitKey)) return displayCache.get(hitKey);

      const map = cellMap.value;
      const sh = sectionHeaderRows.value;
      const canon =
        props.sheet.canonicalSectionMap ??
        props.sheet.canonicalSectionByLabel;
      const cell = getCell(map, row, col);

      if (props.sheetName === 'BD') {
        const masked = displayCellValue(
          map,
          row,
          col,
          sh,
          canon,
          subsystemL1Col.value,
          subsystemL2Col.value
        );
        const bookmark =
          isStructureRow(map, row, sh) || isTitleMarkerRow(map, row, sh);
        const isSubsystemCol =
          col === subsystemL1Col.value ||
          col === subsystemL2Col.value ||
          col === designDeptCol.value;
        const excelSnapshotCol =
          col === titleCol.value || col === massCol.value;
        const useFormula =
          !excelSnapshotCol &&
          !isSubsystemCol &&
          !BD_POSITION_COLS.has(col) &&
          !BD_MASS_AV_AR_COLS.has(col) &&
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
          if (computed !== '' && computed !== '#REF!') {
            displayCache.set(hitKey, computed);
            return computed;
          }
        }
        displayCache.set(hitKey, masked);
        return masked;
      }

      const plain = displayValue(cell);
      displayCache.set(hitKey, plain);
      return plain;
    }

    function updateViewport() {
      if (scrollEl.value) viewportH.value = scrollEl.value.clientHeight;
    }

    watch(
      () => props.sheet,
      () => invalidateDisplayCache()
    );
    watch(
      () => props.outlineOnly,
      () => {
        scrollTop.value = 0;
        if (scrollEl.value) scrollEl.value.scrollTop = 0;
        invalidateDisplayCache();
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
      calcRevision,
      colStyle,
      isStickyDateCol,
      getCell: (r, c) => getCell(cellMap.value, r, c),
      cellDisplay,
      cellReadonly,
      rowStyleClass: (row) => {
        const map = cellMap.value;
        const sh = sectionHeaderRows.value;
        const base = rowStyleClass(map, row, sh);
        const stripe = props.outlineOnly
          ? ''
          : rowDataStripeClass(map, row, sh, props.sheet.dataStartRow);
        return [base, stripe].filter(Boolean).join(' ');
      },
      bdColMetaClass: (col) => bdColMetaClass(col, props.sheet),
      projectCellClass,
      cellInlineStyle: (row, col) =>
        cellInlineStyle(
          getCell(cellMap.value, row, col),
          cellMap.value,
          row,
          col,
          sectionHeaderRows.value,
          props.sheet.matrixColors
        ),
      onScroll,
      onCellInput,
    };
  },
  template: `
    <div class="bd-grid-root">
      <div class="bd-grid-scroll" ref="scrollEl" @scroll.passive="onScroll">
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
                    'col-free-field': col === BD_FREE_FIELD_COL,
                  },
                  bdColMetaClass(col),
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
