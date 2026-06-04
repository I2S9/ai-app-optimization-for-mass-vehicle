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
  BD_HIDDEN_DISPLAY_ROWS,
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
} from './bdStore.js?v=edit-fix3';
import { upsertRawCell } from './sessionPersistence.js?v=edit-fix2';
import {
  ROW_H,
  rowOverscan,
  shouldVirtualizeRows,
  shouldVirtualizeCols,
  createScrollRafSync,
  createRowScrollCache,
  createColScrollCache,
  MAX_RENDERED_ROWS,
} from './gridScroll.js?v=scroll-perf1';
import {
  BD_FREE_FIELD_COL,
  BD_MASS_AV_AR_COLS,
  BD_POSITION_COLS,
  isBdNumericEntryCell,
} from './bdColumnConfig.js?v=input-fix1';
import { createGridCellEditor } from './gridCellEdit.js?v=grid-nav4';
import { createGridCellNavigation } from './gridCellNavigation.js?v=nav-cache1';
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
    outlineMode: { type: Number, default: 0 },
    paneVisible: { type: Boolean, default: true },
    externalEditTick: { type: Number, default: 0 },
    searchCmd: { type: Object, default: null },
  },
  emits: ['cell-change', 'row-delete', 'search-row-hidden', 'search-navigated'],
  setup(props, { emit }) {
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const scrollLeft = ref(0);
    const viewportH = ref(600);
    const viewportW = ref(1200);
    const editEpoch = ref(0);

    const cellMap = shallowRef(new Map());

    /** Rows the user removed via the right-click menu (persisted in raw.deletedRows). */
    function readDeletedRows() {
      const src =
        (props.rawBd && props.rawBd.deletedRows) ||
        (props.sheet && props.sheet.deletedRows) ||
        [];
      return new Set(src.map(Number).filter(Number.isFinite));
    }
    const deletedRows = ref(readDeletedRows());

    const ctxMenu = ref({ visible: false, x: 0, y: 0, row: null, displayRow: null });

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

    /** Outline view: 0 = full, 1 = sections + sub-sections, 2 = sections only. */
    const effectiveOutlineMode = computed(() => {
      const m = Number(props.outlineMode);
      if (Number.isFinite(m) && m >= 0 && m <= 2) return m;
      return props.outlineOnly ? 1 : 0;
    });

    /** Excel rows of yellow sections folded via the row-1 fold column. */
    const collapsedSections = ref(new Set());

    function toggleSection(excelRow) {
      if (excelRow == null) return;
      const next = new Set(collapsedSections.value);
      if (next.has(excelRow)) next.delete(excelRow);
      else next.add(excelRow);
      collapsedSections.value = next;
    }

    const sectionHeaderRows = computed(() => props.sheet.sectionHeaderRows);

    const baseBodyRows = computed(() =>
      props.sheet.bodyDisplayRows != null
        ? props.sheet.bodyDisplayRows
        : computeBodyDisplayRows(props.sheet)
    );

    /**
     * Apply the outline mode (eye) and per-section fold. A folded yellow section
     * hides everything under it up to the next yellow section (the header stays).
     */
    const bodyRows = computed(() => {
      const base = baseBodyRows.value;
      const mode = effectiveOutlineMode.value;
      const collapsed = collapsedSections.value;
      const map = cellMap.value;
      const sh = sectionHeaderRows.value;
      const out = [];
      let displayRow = BD_BODY_DISPLAY_ROW_START;
      let sectionCollapsed = false;
      for (const e of base) {
        const r = e.excelRow;
        if (BD_HIDDEN_DISPLAY_ROWS.has(r)) continue;
        if (deletedRows.value.has(r)) continue;
        const cls = rowStyleClass(map, r, sh);
        if (cls === 'row-section') {
          sectionCollapsed = collapsed.has(r);
          out.push({
            excelRow: r,
            displayRow: displayRow++,
            isSection: true,
            sectionCollapsed,
          });
          continue;
        }
        if (sectionCollapsed) continue;
        if (mode === 2) continue;
        if (mode === 1 && cls !== 'row-subsection') continue;
        out.push({
          excelRow: r,
          displayRow: displayRow++,
          isSection: false,
          sectionCollapsed: false,
        });
      }
      return out;
    });

    const rowCount = computed(() => bodyRows.value.length);
    /** Stable excel-row list — memoized so cell-nav/search keep a constant reference. */
    const bodyExcelRows = computed(() => bodyRows.value.map((e) => e.excelRow));
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
    /**
     * Per-cell style/class memo. Building these inline in the template (colStyle +
     * cellInlineStyle + bdColMetaClass + projectCellClass) allocated a fresh object
     * and ran several functions for every one of the ~200 mounted rows × every column
     * on each scroll-window shift — and the new object identity forced Vue to re-patch
     * the DOM even when nothing changed. Caching by a generation counter returns the
     * SAME reference between renders, so unchanged cells are skipped entirely; only
     * newly scrolled-in rows compute. Mirrors the SynthesisGrid cell-model cache.
     */
    const cellStyleCache = new Map();
    const cellClassCache = new Map();
    let cellModelGen = 0;
    let displayCacheKey = '';

    function invalidateDisplayCache() {
      displayCache.clear();
      cellStyleCache.clear();
      cellClassCache.clear();
      cellModelGen += 1;
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

    /** Merged width + inline style for a data cell — cached (stable ref) per generation. */
    function cellStyleFor(row, col) {
      const key = `${cellModelGen}:${row}:${col}`;
      const cached = cellStyleCache.get(key);
      if (cached !== undefined) return cached;
      const w = widthMap.value.get(col) || 72;
      const base = { width: `${w}px`, minWidth: `${w}px`, maxWidth: `${w}px` };
      const inline = cellInlineStyle(
        getCell(cellMap.value, row, col),
        cellMap.value,
        row,
        col,
        sectionHeaderRows.value,
        props.sheet.matrixColors
      );
      const merged =
        inline && typeof inline === 'object' && Object.keys(inline).length
          ? { ...base, ...inline }
          : base;
      cellStyleCache.set(key, merged);
      return merged;
    }

    /** Static (non-search) class string for a data cell — cached per generation. */
    function cellStaticClass(row, col) {
      const key = `${cellModelGen}:${row}:${col}`;
      const cached = cellClassCache.get(key);
      if (cached !== undefined) return cached;
      const meta = bdColMetaClass(col, props.sheet);
      const proj = projectCellClass(cellDisplay(row, col), col);
      const cls = [meta, proj].filter(Boolean).join(' ');
      cellClassCache.set(key, cls);
      return cls;
    }

    function isStickyDateCol(col) {
      return col === 'A';
    }

    const scrollSync = createScrollRafSync({
      scrollTop,
      scrollLeft,
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
          // A changed display value can change projectCellClass, so the style/class
          // memo must roll its generation too (cheap: rebuilt lazily on next render).
          cellStyleCache.clear();
          cellClassCache.clear();
          cellModelGen += 1;
        }
      }
    );

    watch(
      () => props.externalEditTick,
      () => {
        if (cellEditor && typeof cellEditor.clearIdleDisplays === 'function') {
          cellEditor.clearIdleDisplays();
        }
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
      if (props.session && typeof props.session.setCellValue === 'function') {
        void props.session
          .setCellValue(props.sheetName || 'BD', row, col, value)
          .catch((e) => console.warn('Engine setCellValue:', e));
      }
    }

    const BD_ROW_NUM_W = 42;
    /** Width of the section fold/unfold column inserted before the row numbers. */
    const BD_FOLD_W = 24;
    const BD_HEADER_H = ROW_H * 2;

    // ── Column virtualization (only the sticky col A + the on-screen scrollable
    // columns are rendered). BD draws ~38 columns but ~10 are visible; rendering
    // them all is the main cause of slow re-renders / blank rows on scroll.
    const STICKY_COL = 'A';
    const scrollColumns = computed(() =>
      (columns.value || []).filter((c) => c !== STICKY_COL)
    );
    const colLayout = computed(() => {
      const wm = widthMap.value;
      let left = 0;
      const out = [];
      for (const col of scrollColumns.value) {
        const width = wm.get(col) || 72;
        out.push({ col, left, width });
        left += width;
      }
      return out;
    });
    const scrollableWidth = computed(() => {
      const l = colLayout.value;
      return l.length ? l[l.length - 1].left + l[l.length - 1].width : 0;
    });
    const stickyWidth = computed(
      () => BD_FOLD_W + BD_ROW_NUM_W + (widthMap.value.get(STICKY_COL) || 72)
    );
    const virtualizeCols = computed(() =>
      shouldVirtualizeCols(stickyWidth.value + scrollableWidth.value, viewportW.value)
    );
    const colBufferPx = computed(() =>
      Math.max(400, Math.round(viewportW.value * 0.5))
    );
    const colScrollCache = createColScrollCache(64, { monotonic: false });
    const colRangeStart = ref(0);
    const colRangeEnd = ref(0);

    watchEffect(() => {
      if (!props.paneVisible) return;
      const layout = colLayout.value;
      if (!virtualizeCols.value) {
        colRangeStart.value = 0;
        colRangeEnd.value = layout.length;
        return;
      }
      const vpScrollW = Math.max(100, viewportW.value - stickyWidth.value);
      const range = colScrollCache.resolve(
        layout,
        scrollLeft.value,
        vpScrollW,
        colBufferPx.value
      );
      colRangeStart.value = range.start;
      colRangeEnd.value = range.end;
    });

    /** [A, ...on-screen scrollable cols] — what each row actually renders. */
    const renderColumns = computed(() => {
      if (!virtualizeCols.value) return columns.value || [];
      const out = [STICKY_COL];
      const layout = colLayout.value;
      for (let i = colRangeStart.value; i < colRangeEnd.value; i++) {
        if (layout[i]) out.push(layout[i].col);
      }
      return out;
    });
    /** Width of skipped scrollable columns left of the window (placed after col A). */
    const leftPad = computed(() => {
      if (!virtualizeCols.value) return 0;
      const first = colLayout.value[colRangeStart.value];
      return first ? first.left : 0;
    });
    /** Width of skipped scrollable columns right of the window. */
    const rightPad = computed(() => {
      if (!virtualizeCols.value) return 0;
      const last = colLayout.value[colRangeEnd.value - 1];
      if (!last) return 0;
      return Math.max(0, scrollableWidth.value - (last.left + last.width));
    });

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
      getRows: () => bodyExcelRows.value,
      isNavigable: (row, col) => !cellReadonly(row, col),
      getScrollEl: () => scrollEl.value,
      flushScroll: () => scrollSync.flush(),
      getRowTop: (rowIndex) => BD_HEADER_H + rowIndex * ROW_H,
      getColLeft: (col) => {
        let left = BD_FOLD_W + BD_ROW_NUM_W;
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
      getBodyExcelRows: () => bodyExcelRows.value,
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

    watch(
      () => [props.sheet, props.rawBd],
      () => {
        deletedRows.value = readDeletedRows();
      }
    );

    function onRowContextMenu(entry, event) {
      if (!entry || entry.excelRow == null || entry.excelRow < 2) return;
      event.preventDefault();
      ctxMenu.value = {
        visible: true,
        x: event.clientX,
        y: event.clientY,
        row: entry.excelRow,
        displayRow: entry.displayRow,
      };
    }

    function closeCtxMenu() {
      if (ctxMenu.value.visible) {
        ctxMenu.value = { visible: false, x: 0, y: 0, row: null, displayRow: null };
      }
    }

    function confirmDeleteRow() {
      const row = ctxMenu.value.row;
      closeCtxMenu();
      if (row == null) return;
      const next = new Set(deletedRows.value);
      next.add(row);
      deletedRows.value = next;
      invalidateDisplayCache();
      editEpoch.value += 1;
      emit('row-delete', { sheet: props.sheetName || 'BD', excelRow: row });
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
      if (scrollEl.value) {
        viewportH.value = scrollEl.value.clientHeight;
        viewportW.value = scrollEl.value.clientWidth;
      }
    }

    watch(
      () => rowCount.value,
      () => {
        rowScrollCache.reset();
        nextTick(() => scrollSync.flush());
      }
    );

    watch(
      () => effectiveOutlineMode.value,
      () => {
        rowScrollCache.reset();
        scrollTop.value = 0;
        if (scrollEl.value) scrollEl.value.scrollTop = 0;
        invalidateDisplayCache();
      }
    );

    watch(
      () => collapsedSections.value,
      () => {
        rowScrollCache.reset();
        invalidateDisplayCache();
        nextTick(() => scrollSync.flush());
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
      BD_FREE_FIELD_COL,
      scrollEl,
      columns,
      renderColumns,
      leftPad,
      rightPad,
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
        const stripe = effectiveOutlineMode.value !== 0
          ? ''
          : rowDataStripeClass(map, row, sh, props.sheet.dataStartRow);
        return [base, stripe].filter(Boolean).join(' ');
      },
      bdColMetaClass: (col) => bdColMetaClass(col, props.sheet),
      projectCellClass,
      cellStyleFor,
      cellStaticClass,
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
      ctxMenu,
      onRowContextMenu,
      closeCtxMenu,
      confirmDeleteRow,
      isCellActive,
      cellShowValue,
      onCellMouseDown,
      onCellSpanMouseDown,
      onCellFocus,
      onCellInput,
      onCellBlur,
      onCellKeydown,
      toggleSection,
    };
  },
  template: `
    <div class="bd-grid-root">
      <div class="bd-grid-scroll" ref="scrollEl" @scroll.passive="onScroll">
        <table class="bd-table" role="grid">
          <thead>
            <tr class="hdr-row-letters">
              <th class="bd-fold-col bd-fold-col-hdr"></th>
              <th class="corner"></th>
              <template v-for="(col, idx) in renderColumns" :key="'letter-' + col">
                <th v-if="idx === 1 && leftPad > 0" class="col-spacer" :style="{ width: leftPad + 'px', minWidth: leftPad + 'px', maxWidth: leftPad + 'px', padding: 0, border: 'none' }"></th>
                <th
                  class="col-letter"
                  :class="{ 'col-sticky-date': isStickyDateCol(col) }"
                  :style="colStyle(col)"
                >{{ col }}</th>
              </template>
              <th v-if="rightPad > 0" class="col-spacer" :style="{ width: rightPad + 'px', minWidth: rightPad + 'px', maxWidth: rightPad + 'px', padding: 0, border: 'none' }"></th>
            </tr>
            <tr class="hdr-row-1">
              <th class="bd-fold-col bd-fold-col-hdr"></th>
              <th class="corner">1</th>
              <template v-for="(col, idx) in renderColumns" :key="'h1-' + col">
                <th v-if="idx === 1 && leftPad > 0" class="col-spacer" :style="{ width: leftPad + 'px', minWidth: leftPad + 'px', maxWidth: leftPad + 'px', padding: 0, border: 'none' }"></th>
                <th
                  class="col-hdr"
                  :class="{ 'col-sticky-date': isStickyDateCol(col) }"
                  :style="colStyle(col)"
                  :title="sheet.headers[col] || col"
                >{{ sheet.headers[col] || col }}</th>
              </template>
              <th v-if="rightPad > 0" class="col-spacer" :style="{ width: rightPad + 'px', minWidth: rightPad + 'px', maxWidth: rightPad + 'px', padding: 0, border: 'none' }"></th>
            </tr>
          </thead>
          <tbody>
            <tr v-if="topSpacer > 0" class="bd-spacer-top">
              <td :colspan="columns.length + 2" :style="{ height: topSpacer + 'px', padding: 0, border: 'none' }"></td>
            </tr>
            <tr
              v-for="entry in visibleRows"
              :key="entry.excelRow"
              class="grid-row-cv"
              :class="rowStyleClass(entry.excelRow)"
              @contextmenu="onRowContextMenu(entry, $event)"
            >
              <td
                class="bd-fold-col"
                :class="{ 'bd-fold-col-section': entry.isSection }"
              >
                <button
                  v-if="entry.isSection"
                  type="button"
                  class="bd-fold-btn"
                  :title="(entry.sectionCollapsed ? 'Déplier' : 'Plier') + ' la section'"
                  @click="toggleSection(entry.excelRow)"
                >{{ entry.sectionCollapsed ? '+' : '\u2212' }}</button>
              </td>
              <td class="row-num">{{ entry.displayRow }}</td>
              <template v-for="(col, idx) in renderColumns" :key="entry.excelRow + '-' + col">
                <td v-if="idx === 1 && leftPad > 0" class="col-spacer" :style="{ width: leftPad + 'px', minWidth: leftPad + 'px', maxWidth: leftPad + 'px', padding: 0, border: 'none' }"></td>
                <td
                  class="data-cell"
                  :class="[
                    {
                      readonly: cellReadonly(entry.excelRow, col),
                      'col-sticky-date': isStickyDateCol(col),
                      'col-free-field': col === BD_FREE_FIELD_COL,
                      'grid-search-hit': isSearchHit(entry.excelRow, col),
                      'grid-search-focus': isSearchFocus(entry.excelRow, col),
                    },
                    cellStaticClass(entry.excelRow, col),
                  ]"
                  :style="cellStyleFor(entry.excelRow, col)"
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
              </template>
              <td v-if="rightPad > 0" class="col-spacer" :style="{ width: rightPad + 'px', minWidth: rightPad + 'px', maxWidth: rightPad + 'px', padding: 0, border: 'none' }"></td>
            </tr>
            <tr v-if="bottomSpacer > 0" class="bd-spacer-bottom">
              <td :colspan="columns.length + 2" :style="{ height: bottomSpacer + 'px', padding: 0, border: 'none' }"></td>
            </tr>
          </tbody>
        </table>
      </div>
      <template v-if="ctxMenu.visible">
        <div
          class="grid-ctx-overlay"
          @mousedown="closeCtxMenu"
          @contextmenu.prevent="closeCtxMenu"
          @wheel="closeCtxMenu"
        ></div>
        <div
          class="grid-ctx-menu"
          :style="{ left: ctxMenu.x + 'px', top: ctxMenu.y + 'px' }"
          @contextmenu.prevent
        >
          <button type="button" class="grid-ctx-item grid-ctx-danger" @click="confirmDeleteRow">
            Supprimer la ligne {{ ctxMenu.displayRow }}
          </button>
        </div>
      </template>
    </div>
  `,
};
