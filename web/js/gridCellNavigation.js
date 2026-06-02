/**
 * Arrow-key navigation between editable grid cells (Excel-style, low latency).
 */
import { nextTick } from 'vue';
import { ROW_H } from './gridScroll.js?v=syn-scroll5';

const ARROW_KEYS = new Set([
  'ArrowUp',
  'ArrowDown',
  'ArrowLeft',
  'ArrowRight',
]);

const ARROW_DELTA = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

/**
 * @param {{
 *   getColumns: () => string[],
 *   getRows: () => number[],
 *   isNavigable: (row: number, col: string) => boolean,
 *   getScrollEl: () => HTMLElement | null,
 *   flushScroll?: () => void,
 *   getRowTop?: (rowIndex: number, excelRow: number) => number | null,
 *   getColLeft?: (col: string) => number | null,
 *   getColWidth?: (col: string) => number,
 *   rowHeight?: number,
 * }} opts
 */
export function createGridCellNavigation(opts) {
  const {
    getColumns,
    getRows,
    isNavigable,
    getScrollEl,
    flushScroll,
    getRowTop,
    getColLeft,
    getColWidth = () => 72,
    rowHeight = ROW_H,
  } = opts;

  let colIndex = new Map();
  let rowIndex = new Map();
  let cachedCols = [];
  let cachedRows = [];

  function rebuildIndex() {
    const cols = getColumns();
    const rows = getRows();
    cachedCols = cols;
    cachedRows = rows;
    colIndex = new Map(cols.map((c, i) => [c, i]));
    rowIndex = new Map(rows.map((r, i) => [r, i]));
  }

  function move(row, col, key) {
    const delta = ARROW_DELTA[key];
    if (!delta) return null;

    rebuildIndex();

    let ri = rowIndex.get(row);
    let ci = colIndex.get(col);
    if (ri == null || ci == null) return null;

    const [dr, dc] = delta;
    const maxSteps = cachedRows.length + cachedCols.length;
    for (let step = 0; step < maxSteps; step += 1) {
      ri += dr;
      ci += dc;
      if (ri < 0 || ci < 0 || ri >= cachedRows.length || ci >= cachedCols.length) {
        return null;
      }
      const nextRow = cachedRows[ri];
      const nextCol = cachedCols[ci];
      if (isNavigable(nextRow, nextCol)) {
        return { row: nextRow, col: nextCol };
      }
    }
    return null;
  }

  function findInput(scrollRoot, row, col) {
    if (!scrollRoot) return null;
    return scrollRoot.querySelector(
      `input.grid-cell-input[data-grid-row="${row}"][data-grid-col="${col}"]`
    );
  }

  function scrollCellIntoView(row, col, { force = false } = {}) {
    const scrollEl = getScrollEl();
    if (!scrollEl) return;

    if (!force && findInput(scrollEl, row, col)) return;

    const rows = getRows();
    const ri = rows.indexOf(row);
    if (ri >= 0 && getRowTop) {
      const top = getRowTop(ri, row);
      if (top != null) {
        const bottom = top + rowHeight;
        const st = scrollEl.scrollTop;
        const vh = scrollEl.clientHeight;
        if (top < st) scrollEl.scrollTop = top;
        else if (bottom > st + vh) scrollEl.scrollTop = bottom - vh;
      }
    }

    if (getColLeft) {
      const left = getColLeft(col);
      if (left != null) {
        const right = left + getColWidth(col);
        const sl = scrollEl.scrollLeft;
        const vw = scrollEl.clientWidth;
        if (left < sl) scrollEl.scrollLeft = left;
        else if (right > sl + vw) scrollEl.scrollLeft = right - vw;
      }
    }

    if (flushScroll) flushScroll();
  }

  function focusCell(row, col, beginNavigationTo, activateCell, onDone) {
    beginNavigationTo(row, col);

    const attempt = () => {
      const input = findInput(getScrollEl(), row, col);
      if (input instanceof HTMLInputElement) {
        activateCell(row, col, input);
        return true;
      }
      return false;
    };

    nextTick(() => {
      if (attempt()) {
        onDone();
        return;
      }
      scrollCellIntoView(row, col, { force: false });
      nextTick(() => {
        if (attempt()) {
          onDone();
          return;
        }
        requestAnimationFrame(() => {
          attempt();
          onDone();
        });
      });
    });
  }

  function scrollCellIntoViewForced(row, col) {
    scrollCellIntoView(row, col, { force: true });
  }

  function wrapKeydown(
    onCellKeydown,
    prepareNavigate,
    beginNavigationTo,
    activateCell,
    setNavigationLock
  ) {
    return (row, col, event) => {
      if (!ARROW_KEYS.has(event.key)) {
        onCellKeydown(row, col, event);
        return;
      }
      event.preventDefault();
      if (!prepareNavigate(row, col, event)) return;
      const next = move(row, col, event.key);
      if (!next) return;

      setNavigationLock(true);
      focusCell(next.row, next.col, beginNavigationTo, activateCell, () => {
        setNavigationLock(false);
      });
    };
  }

  rebuildIndex();

  return {
    move,
    focusCell,
    wrapKeydown,
    rebuildIndex,
    scrollCellIntoView,
    scrollCellIntoViewForced,
  };
}
