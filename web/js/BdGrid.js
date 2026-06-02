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
  isStructureRow,
  isTitleMarkerRow,
  rowStyleClass,
  rowDataStripeClass,
  cellInlineStyle,
  projectCellClass,
  bdColMetaClass,
  bdMassCol,
  bdTitleCol,
} from './bdStore.js?v=edit-fix2';
import { upsertRawCell } from './sessionPersistence.js?v=edit-fix2';
import {
  ROW_H,
  rowOverscan,
  shouldVirtualizeRows,
  createScrollRafSync,
  createRowScrollCache,
  MAX_RENDERED_ROWS,
} from './gridScroll.js?v=syn-fix1';
import {
  BD_FREE_FIELD_COL,
  BD_MASS_AV_AR_COLS,
  BD_POSITION_COLS,
  isBdNumericEntryCell,
} from './bdColumnConfig.js?v=input-fix1';
import { createGridCellEditor } from './gridCellEdit.js?v=grid-nav4';
import { createGridCellNavigation } from './gridCellNavigation.js?v=grid-search4';
import {
  buildSearchIndex,
  createGridSearchController,
} from './gridSearch.js?v=grid-search4';

export default {
  name: 'BdGrid',
  props: {
    sheet: { type: Object, required: true },
    sheetName: { type: String, default: 'BD' },
    session: { type: Object, default: null },
    rawBd: { type: Object, default: null },
    outlineOnly: { type: Boolean, default: false },
    paneVisible: { type: Boolean, default: true },
    externalEditTick: { type: Number, default: 0 },
    searchCmd: { type: Object, default: null },
  },
  emits: ['cell-change', 'search-row-hidden', 'search-navigated'],
  setup(props, { emit }) {
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const viewportH = ref(600);
    const editEpoch = ref(0);

    const cellMap = shallowRef(new Map());

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
        return (
          (props.sheet.outlineBodyDisplayRows != null
            ? props.sheet.outlineBodyDisplayRows
            : props.sheet.bodyDisplayRows != null
              ? props.sheet.bodyDisplayRows
              : null) ||
          computeBodyDisplayRows(props.sheet)
        );
      }
      return props.sheet.bodyDisplayRows != null
        ? props.sheet.bodyDisplayRows
        : computeBodyDisplayRows(props.sheet);
    });

    const rowCount = computed(() => bodyRows.value.length);
    const virtualizeRows = computed(() =>
      shouldVirtualizeRows(rowCount.value, viewportH.value)
    );
    const rowOverscanPx = computed(() => rowOverscan(viewportH.value));
    const rowScrollCache = createRowScrollCache(MAX_RENDERED_ROWS, { monotonic: false });
    const visibleStart = ref(0);
    const visibleEnd = ref(0);

    watchEffect(() => {
      if (!props.paneVisible) return;
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
        rowOverscanPx.value
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
      const remaining = rowCount.value - visibleStart.value - shown;
      return Math.max(0, remaining * ROW_H);
    });

    const sectionHeaderRows = computed(() => props.sheet.sectionHeaderRows);
    const subsystemL1Col = computed(() => bdSubsystemL1Col(props.sheet));
    const subsystemL2Col = computed(() => bdSubsystemL2Col(props.sheet));
    const designDeptCol = computed(() => bdDesignDeptCol(props.sheet));
    const titleCol = computed(() => bdTitleCol(props.sheet));
    const massCol = computed(() => bdMassCol(props.sheet));
    const calcRevision = computed(() =>
      props.session && props.session.revision && props.session.revision.value != null
        ? props.session.revision.value
        : 0
    );
    const engineReady = computed(() =>
      Boolean(props.session && props.session.ready && props.session.ready.value)
    );
    const displayCache = new Map();
    let displayCacheKey = '';

    function invalidateDisplayCache() {
      displayCache.clear();
      displayCacheKey = '';
    }

    watch(
      () => props.sheet,
      (sheet) => {
        cellMap.value =
          sheet && sheet.cellMap instanceof Map
            ? sheet.cellMap
            : buildCellMap(sheet && sheet.cells, sheet && sheet.headerRows);
        invalidateDisplayCache();
      },
      { immediate: true }
    );

    function colStyle(col) {
      const w = widthMap.value.get(col) || 72;
      return { width: `${w}px`, minWidth: `${w}px`, maxWidth: `${w}px` };
    }

    function isStickyDateCol(col) {
      return col === 'A';
    }

    const scrollSync = createScrollRafSync({
      scrollTop,
      getScrollEl: () => scrollEl.value,
    });

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
          for (const k of dirty.keys) displayCache.delete(k);
        }
      }
    );

    watch(
      () => props.externalEditTick,
      () => {
        editEpoch.value += 1;
        invalidateDisplayCache();
      }
    );

    function commitCellInput(row, col, value, previousValue) {
      const key = `${row}:${col}`;
      let cell = cellMap.value.get(key);
      const hadFormula = Boolean(cell && cell.f);
      if (!cell) {
        cell = { r: row, c: col, v: value, userEdited: true };
        props.sheet.cells.push(cell);
        cellMap.value.set(key, cell);
      } else {
        cell.v = value;
        cell.userEdited = true;
        delete cell.f;
      }
      if (props.rawBd) upsertRawCell(props.rawBd, row, col, value);
      editEpoch.value += 1;
      invalidateDisplayCache();
      emit('cell-change', {
        row,
        col,
        value,
        sheet: props.sheetName || 'BD',
        previousValue: previousValue != null ? previousValue : '',
      });
      if (props.session && props.session.ready && props.session.ready.value) {
        void props.session
          .setCellValue(props.sheetName, row, col, value)
          .catch((e) => console.warn('Engine setCellValue:', e));
      }
    }

    const BD_ROW_NUM_W = 42;
    const BD_HEADER_H = ROW_H * 2;

    const cellEditor = createGridCellEditor({
      isNumericAt: (row, col) => isBdNumericEntryCell(row, col),
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

    const cellNavigation = createGridCellNavigation({
      getColumns: () => columns.value,
      getRows: () => bodyRows.value.map((e) => e.excelRow),
      isNavigable: (row, col) => !cellReadonly(row, col),
      getScrollEl: () => scrollEl.value,
      flushScroll: () => scrollSync.flush(),
      getRowTop: (rowIndex) => BD_HEADER_H + rowIndex * ROW_H,
      getColLeft: (col) => {
        let left = BD_ROW_NUM_W;
        for (const c of columns.value) {
          if (c === col) return left;
          left += widthMap.value.get(c) || 72;
        }
        return null;
      },
      getColWidth: (col) => widthMap.value.get(col) || 72,
      rowHeight: ROW_H,
    });
    const onCellKeydown = cellNavigation.wrapKeydown(
      onCellKeydownBase,
      prepareNavigate,
      beginNavigationTo,
      activateCell,
      setNavigationLock
    );

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
        columns.value
      ).then((idx) => {
        searchIndex = idx;
        searchIndexPromise = null;
        return idx;
      });
      return searchIndexPromise;
    }

    const gridSearch = createGridSearchController({
      getSearchIndex: getSearchIndexForGrid,
      getBodyExcelRows: () => bodyRows.value.map((e) => e.excelRow),
      getScrollEl: () => scrollEl.value,
      scrollCellIntoView: (row, col) =>
        cellNavigation.scrollCellIntoViewForced(row, col),
      flushScroll: () => scrollSync.flush(),
      getRowTop: (rowIndex) => BD_HEADER_H + rowIndex * ROW_H,
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

    /** Row 1 = Excel headers only; all body rows editable. */
    function cellReadonly(row, col) {
      return row < 2;
    }

    function cellDisplay(row, col) {
      const cacheKey = `${calcRevision.value}:${editEpoch.value}:${engineReady.value}:${props.outlineOnly}:${props.sheet === null ? 0 : 1}`;
      if (cacheKey !== displayCacheKey) {
        invalidateDisplayCache();
        displayCacheKey = cacheKey;
      }
      const map = cellMap.value;
      const hitKey = `${row}:${col}`;
      if (displayCache.has(hitKey)) {
        const cached = displayCache.get(hitKey);
        if (cached !== '' || !getCell(map, row, col)) return cached;
        displayCache.delete(hitKey);
      }
      const sh = sectionHeaderRows.value;
      const canon =
        props.sheet.canonicalSectionMap != null
          ? props.sheet.canonicalSectionMap
          : props.sheet.canonicalSectionByLabel;
      const cell = getCell(map, row, col);

      if (props.sheetName === 'BD') {
        if (cell && cell.userEdited) {
          const plain = cell.v != null ? String(cell.v) : '';
          displayCache.set(hitKey, plain);
          return plain;
        }
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
          props.session &&
          props.session.ready &&
          props.session.ready.value &&
          props.session.isFormulaCell(props.sheetName, row, col, cell);
        if (useFormula && cell && cell.f) {
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

    function cellTitle(row, col) {
      const cell = getCell(cellMap.value, row, col);
      return cell && cell.f ? String(cell.f) : '';
    }

    function updateViewport() {
      if (scrollEl.value) viewportH.value = scrollEl.value.clientHeight;
    }

    watch(
      () => rowCount.value,
      () => {
        rowScrollCache.reset();
        nextTick(() => scrollSync.flush());
      }
    );

    watch(
      () => props.outlineOnly,
      () => {
        rowScrollCache.reset();
        scrollTop.value = 0;
        if (scrollEl.value) scrollEl.value.scrollTop = 0;
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

    onMounted(() => {
      updateViewport();
      window.addEventListener('resize', updateViewport);
    });
    onUnmounted(() => {
      window.removeEventListener('resize', updateViewport);
      scrollSync.dispose();
    });

    return {
      scrollEl,
      columns,
      visibleRows,
      topSpacer,
      bottomSpacer,
      visibleStart,
      outlineOnly: computed(() => props.outlineOnly),
      calcRevision,
      colStyle,
      isStickyDateCol,
      getCell: (r, c) => getCell(cellMap.value, r, c),
      cellTitle,
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
      isSearchHit: gridSearch.isSearchHit,
      isSearchFocus: gridSearch.isSearchFocus,
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
              class="grid-row-cv"
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
                    'grid-search-hit': isSearchHit(entry.excelRow, col),
                    'grid-search-focus': isSearchFocus(entry.excelRow, col),
                  },
                  bdColMetaClass(col),
                  projectCellClass(cellDisplay(entry.excelRow, col), col),
                ]"
                :style="[colStyle(col), cellInlineStyle(entry.excelRow, col)]"
              >
                <template v-if="cellReadonly(entry.excelRow, col)">
                  <span :title="cellTitle(entry.excelRow, col)">{{ cellDisplay(entry.excelRow, col) }}</span>
                </template>
                <input
                  v-else-if="isCellActive(entry.excelRow, col)"
                  type="text"
                  class="grid-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  :data-grid-row="entry.excelRow"
                  :data-grid-col="col"
                  @mousedown="onCellMouseDown(entry.excelRow, col, $event)"
                  @focus="onCellFocus(entry.excelRow, col, $event)"
                  @input="onCellInput(entry.excelRow, col, $event)"
                  @blur="onCellBlur(entry.excelRow, col, $event)"
                  @keydown="onCellKeydown(entry.excelRow, col, $event)"
                />
                <span
                  v-else
                  class="grid-cell-display"
                  @mousedown="onCellSpanMouseDown(entry.excelRow, col, $event)"
                >{{ cellShowValue(entry.excelRow, col, cellDisplay(entry.excelRow, col)) }}</span>
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
