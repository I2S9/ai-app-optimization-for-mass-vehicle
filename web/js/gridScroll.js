/** Shared virtual-scroll helpers for BD / Synthesis grids. */

export const ROW_H = 21;

/** BD row window — small cap keeps UI + menu responsive. */
export const MAX_RENDERED_ROWS = 64;

/** Synthesis horizontal window — bounded to limit DOM / memory (was unbounded monotonic). */
export const SYN_MAX_RENDERED_COLS = 56;

/** Rows to render above/below the viewport. */
export function rowOverscan(viewportH, minRows = 24, maxRows = 120) {
  const visible = Math.max(1, Math.ceil(viewportH / ROW_H));
  return Math.min(maxRows, Math.max(minRows, Math.ceil(visible * 1.5)));
}

/** Horizontal pixels to render left/right of the viewport. */
export function colOverscanPx(viewportW, minPx = 480) {
  return Math.min(3200, Math.max(minPx, Math.floor(viewportW * 1.5)));
}

/** Synthesis horizontal buffer — wide enough to avoid pop-in while scrolling. */
export function synColOverscanPx(viewportW) {
  return Math.min(1600, Math.max(400, Math.floor(viewportW * 1.2)));
}

/** Binary search on precomputed column layout (`left`, `width`). */
export function findVisibleColIndexRange(layout, minPx, maxPx) {
  if (!layout.length) return { start: 0, end: 0 };

  let lo = 0;
  let hi = layout.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (layout[mid].left + layout[mid].width < minPx) lo = mid + 1;
    else hi = mid;
  }
  const start = lo;

  lo = 0;
  hi = layout.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (layout[mid].left <= maxPx) lo = mid + 1;
    else hi = mid;
  }

  return { start, end: Math.min(lo, layout.length) };
}

export function visibleRowRange(
  scrollTop,
  viewportH,
  rowCount,
  overscan,
  maxRendered = MAX_RENDERED_ROWS
) {
  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - overscan);
  let count = Math.ceil(viewportH / ROW_H) + overscan * 2;
  count = Math.min(count, maxRendered);
  const end = Math.min(rowCount, start + count);
  return { start, end };
}

/**
 * Expanding range cache — once a row/col was visited it stays mounted.
 * monotonic=true: never trim (scroll back keeps everything).
 */
export function createRangeScrollCache(maxSpan = MAX_RENDERED_ROWS, { monotonic = false } = {}) {
  let lo = 0;
  let hi = 0;
  let primed = false;

  function reset() {
    lo = 0;
    hi = 0;
    primed = false;
  }

  function resolveVisible(start, end, total) {
    if (!primed) {
      lo = start;
      hi = end;
      primed = true;
      return { start: lo, end: hi };
    }

    lo = Math.min(lo, start);
    hi = Math.max(hi, end);
    hi = Math.min(hi, total);

    if (!monotonic) {
      const span = hi - lo;
      const cap = maxSpan;
      if (span > cap) {
        const viewMid = (start + end) >> 1;
        const half = Math.floor(cap / 2);
        lo = Math.max(0, viewMid - half);
        hi = Math.min(total, lo + cap);
        if (hi - lo < cap) lo = Math.max(0, hi - cap);
      }
    }

    return { start: lo, end: hi };
  }

  function resolveRows(scrollTop, viewportH, rowCount, overscan) {
    if (rowCount <= 0) return { start: 0, end: 0 };
    const { start, end } = visibleRowRange(
      scrollTop,
      viewportH,
      rowCount,
      overscan,
      monotonic ? rowCount : maxSpan
    );
    return resolveVisible(start, end, rowCount);
  }

  function resolveCols(layout, scrollLeft, viewportW, bufferPx) {
    if (!layout.length) return { start: 0, end: 0 };
    const min = scrollLeft - bufferPx;
    const max = scrollLeft + viewportW + bufferPx;
    const { start, end } = findVisibleColIndexRange(layout, min, max);
    return resolveVisible(start, end, layout.length);
  }

  return { resolveRows, resolveCols, reset };
}

/** @deprecated alias */
export function createRowScrollCache(maxSpan, opts) {
  const cache = createRangeScrollCache(maxSpan, opts);
  return {
    resolve: (scrollTop, viewportH, rowCount, overscan) =>
      cache.resolveRows(scrollTop, viewportH, rowCount, overscan),
    reset: cache.reset,
  };
}

/** @deprecated alias */
export function createColScrollCache(maxSpan, opts) {
  const cache = createRangeScrollCache(maxSpan, opts);
  return {
    resolve: (layout, scrollLeft, viewportW, bufferPx) =>
      cache.resolveCols(layout, scrollLeft, viewportW, bufferPx),
    reset: cache.reset,
  };
}

/** BD uses row windowing; Synthesis renders all body rows (native scroll). */
export function shouldVirtualizeRows(rowCount, viewportH, { forceOff = false } = {}) {
  if (forceOff) return false;
  const visible = Math.ceil(viewportH / ROW_H);
  return rowCount > visible + 24;
}

export function shouldVirtualizeCols(tableWidth, viewportW) {
  return tableWidth > viewportW + 120;
}

/** Sync scroll — virtual window must match native scroll position (no grey bands). */
export function createScrollRafSync(refs) {
  const { scrollTop, scrollLeft = null, getScrollEl } = refs;

  function apply(el) {
    scrollTop.value = el.scrollTop;
    if (scrollLeft) scrollLeft.value = el.scrollLeft;
  }

  function onScroll(e) {
    apply(e.target);
  }

  function flush() {
    const el = getScrollEl?.() ?? null;
    if (el) apply(el);
  }

  function dispose() {}

  return { onScroll, flush, dispose };
}
