/** Fast in-grid search (raw cell text) + yellow highlights. */
import { ref, shallowRef, nextTick } from 'vue';
import { ROW_H } from './gridScroll.js?v=grid-search4';
import { yieldToMain } from './yieldMain.js?v=2';

export function normalizeSearchQuery(q) {
  return String(q != null ? q : '').trim().toLowerCase();
}

function cellSearchText(cell) {
  if (!cell) return '';
  const parts = [];
  if (cell.v != null) parts.push(String(cell.v));
  if (cell.w != null) parts.push(String(cell.w));
  if (cell.f) parts.push(String(cell.f));
  return parts.join(' ').toLowerCase();
}

function cellKey(row, col) {
  return `${row}:${col}`;
}

/**
 * @returns {{ row: number, col: string, text: string }[]}
 */
export async function buildSearchIndex(cellMap, sheet, columns) {
  if (!(cellMap instanceof Map) || cellMap.size === 0) return [];

  const cols = Array.isArray(columns) ? columns : null;
  const colSet = cols && cols.length ? new Set(cols) : null;
  const colOrder = cols && cols.length
    ? new Map(cols.map((c, i) => [c, i]))
    : null;
  const minRow = sheet && sheet.dataStartRow != null ? sheet.dataStartRow : 1;
  const maxRow = sheet && sheet.lastRow != null ? sheet.lastRow : 999999;
  const entries = [];
  let n = 0;

  for (const [key, cell] of cellMap) {
    const sep = key.indexOf(':');
    if (sep < 0) continue;
    const row = Number(key.slice(0, sep));
    const col = key.slice(sep + 1);
    if (!Number.isFinite(row) || !col) continue;
    if (row < minRow || row > maxRow) continue;
    if (colSet && !colSet.has(col)) continue;
    const text = cellSearchText(cell);
    if (!text) continue;
    entries.push({
      row,
      col,
      // Store numeric order once to make sorting cheap.
      _o: colOrder ? (colOrder.has(col) ? colOrder.get(col) : 999999) : 0,
      text,
    });
    n += 1;
    if (n % 12000 === 0) await yieldToMain();
  }

  entries.sort(
    (a, b) =>
      a.row - b.row || (a._o != null ? a._o : 0) - (b._o != null ? b._o : 0)
  );
  // Remove internal sort key to keep the index compact.
  for (let i = 0; i < entries.length; i += 1) delete entries[i]._o;
  return entries;
}

/** Substring match — cells in row/col order. */
export function findInSearchIndex(index, query) {
  const q = normalizeSearchQuery(query);
  if (!q || !index || !index.length) return [];
  const hits = [];
  for (let i = 0; i < index.length; i += 1) {
    if (index[i].text.includes(q)) {
      hits.push({ row: index[i].row, col: index[i].col });
    }
  }
  return hits;
}

export function createGridSearchController(opts) {
  const {
    getSearchIndex,
    getBodyExcelRows,
    getScrollEl,
    scrollCellIntoView,
    flushScroll,
    getRowTop,
    getViewportH,
    rowHeight = ROW_H,
    onRowHidden,
  } = opts;

  const searchHitKeys = shallowRef(new Set());
  const searchFocusKey = ref('');
  let searchMatches = [];
  let searchMatchIndex = 0;
  let lastQuery = '';
  let searchGen = 0;

  function isSearchHit(row, col) {
    return searchHitKeys.value.has(cellKey(row, col));
  }

  function isSearchFocus(row, col) {
    return searchFocusKey.value === cellKey(row, col);
  }

  function applyHighlights() {
    searchHitKeys.value = new Set(
      searchMatches.map((m) => cellKey(m.row, m.col))
    );
    const cur = searchMatches[searchMatchIndex];
    searchFocusKey.value = cur ? cellKey(cur.row, cur.col) : '';
  }

  function clearSearch() {
    searchMatches = [];
    searchMatchIndex = 0;
    lastQuery = '';
    searchHitKeys.value = new Set();
    searchFocusKey.value = '';
    searchGen += 1;
  }

  function centerRowInView(rowIndex, excelRow) {
    const scrollEl = getScrollEl();
    if (!scrollEl) return;
    const top = getRowTop(rowIndex, excelRow);
    if (top == null) return;
    const vh = getViewportH();
    scrollEl.scrollTop = Math.max(0, top - vh / 2 + rowHeight / 2);
    if (flushScroll) flushScroll();
  }

  async function searchAndScroll(query, { step = 0 } = {}) {
    const q = normalizeSearchQuery(query);
    const gen = searchGen;
    if (!q) {
      clearSearch();
      return { count: 0, index: 0 };
    }

    const queryChanged = lastQuery !== q;

    if (step === 0 || queryChanged) {
      const index = await getSearchIndex();
      if (gen !== searchGen) return { count: 0, index: 0, cancelled: true };
      searchMatches = findInSearchIndex(index, q);
      lastQuery = q;
      searchMatchIndex = 0;
    } else if (searchMatches.length) {
      searchMatchIndex =
        (searchMatchIndex + step + searchMatches.length) %
        searchMatches.length;
    }

    if (!searchMatches.length) {
      searchHitKeys.value = new Set();
      searchFocusKey.value = '';
      return { count: 0, index: 0 };
    }

    applyHighlights();

    const { row, col } = searchMatches[searchMatchIndex];
    const rows = getBodyExcelRows();
    const ri = rows.indexOf(row);
    if (ri < 0) {
      if (onRowHidden) onRowHidden({ row, col });
      return {
        count: searchMatches.length,
        index: searchMatchIndex,
        hidden: true,
      };
    }

    scrollCellIntoView(row, col);
    centerRowInView(ri, row);
    await nextTick();
    if (flushScroll) flushScroll();

    return {
      count: searchMatches.length,
      index: searchMatchIndex,
    };
  }

  return {
    isSearchHit,
    isSearchFocus,
    clearSearch,
    searchAndScroll,
    bumpGeneration: () => {
      searchGen += 1;
    },
  };
}
