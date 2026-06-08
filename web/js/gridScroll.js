/** Shared virtual-scroll helpers for BD / Synthesis grids. */

export const ROW_H = 21;

/**
 * Render windows. Sized so that a fast fling stays inside the already-mounted
 * DOM for ~1.5 viewports in every direction — this is what kills the grey/white
 * "loading" gaps: the browser can scroll a long way before Vue re-renders, so
 * the buffer must extend well past the visible area. Cell models are cached
 * (keyed on edit/calc generation, not scroll position), so a larger window only
 * costs DOM nodes, not recomputation, once a region has been visited.
 */
export const MAX_RENDERED_ROWS = 120;

/** Synthesis row/column windows — viewport always covered, no full-sheet mount. */
export const SYN_MAX_RENDERED_COLS = 104;
export const SYN_MAX_RENDERED_ROWS = 132;

/** Rows above/below viewport (BD + Synthesis) — ~1.5 screens of buffer. */
export function rowOverscan(viewportH, minRows = 28, maxRows = 64) {
  const visible = Math.max(1, Math.ceil(viewportH / ROW_H));
  return Math.min(maxRows, Math.max(minRows, Math.ceil(visible * 1.5)));
}

/** Horizontal pixels left/right of viewport. */
export function colOverscanPx(viewportW, minPx = 640) {
  return Math.min(4000, Math.max(minPx, Math.floor(viewportW * 2)));
}

/**
 * Synthesis horizontal buffer (~1.2 screens each side). Wider than the old 0.6×
 * window so fast flings stay inside already-mounted DOM; monotonic col cache caps
 * total nodes at SYN_MAX_RENDERED_COLS.
 */
export function synColOverscanPx(viewportW) {
  return Math.min(2800, Math.max(900, Math.floor(viewportW * 1.2)));
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

/**
 * Row slice that always includes the viewport within [start, end).
 */
export function visibleRowRange(
  scrollTop,
  viewportH,
  rowCount,
  overscanRows,
  maxRendered = MAX_RENDERED_ROWS
) {
  if (rowCount <= 0) return { start: 0, end: 0, viewportStart: 0, viewportEnd: 0 };

  const visible = Math.max(1, Math.ceil(viewportH / ROW_H));
  const span = Math.min(maxRendered, rowCount);
  const viewportStart = Math.min(
    Math.max(0, Math.floor(scrollTop / ROW_H)),
    Math.max(0, rowCount - 1)
  );
  const viewportEnd = Math.min(rowCount, viewportStart + visible);

  const maxOverscan = Math.max(0, Math.floor((span - visible) / 2));
  const overscan = Math.min(
    typeof overscanRows === 'number' ? overscanRows : 0,
    maxOverscan
  );

  let start = Math.max(0, viewportStart - overscan);
  let end = Math.min(rowCount, viewportEnd + overscan);

  if (end - start > span) {
    const mid = Math.floor((viewportStart + viewportEnd - 1) / 2);
    start = Math.max(0, mid - Math.floor(span / 2));
    end = Math.min(rowCount, start + span);
    if (end - start < span) start = Math.max(0, end - span);
  } else if (end - start < span) {
    const deficit = span - (end - start);
    start = Math.max(0, start - Math.floor(deficit / 2));
    end = Math.min(rowCount, start + span);
    start = Math.max(0, end - span);
  }

  start = Math.min(start, viewportStart);
  end = Math.max(end, viewportEnd);

  return { start, end, viewportStart, viewportEnd };
}

/**
 * Merge viewport range into cache; never leave [vpStart, vpEnd) uncovered.
 */
function mergeViewportRange(lo, hi, start, end, vpStart, vpEnd, total, maxSpan) {
  lo = Math.min(lo, start, vpStart);
  hi = Math.max(hi, end, vpEnd);
  lo = Math.max(0, lo);
  hi = Math.min(total, hi);

  if (hi - lo <= maxSpan) return { start: lo, end: hi };

  let nLo = Math.max(0, vpStart - Math.floor((maxSpan - (vpEnd - vpStart)) / 2));
  let nHi = Math.min(total, nLo + maxSpan);
  if (nHi < vpEnd) {
    nHi = vpEnd;
    nLo = Math.max(0, nHi - maxSpan);
  }
  if (nLo > vpStart) {
    nLo = vpStart;
    nHi = Math.min(total, nLo + maxSpan);
  }
  return { start: nLo, end: nHi };
}

/**
 * monotonic=true: visited rows/cols stay mounted (scroll back = nothing vanishes).
 * Viewport is always fully covered even when trimming to maxSpan.
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

  function resolveVisible(start, end, total, vpStart = start, vpEnd = end) {
    if (!monotonic) {
      return mergeViewportRange(start, end, start, end, vpStart, vpEnd, total, maxSpan);
    }

    if (!primed) {
      lo = Math.min(start, vpStart);
      hi = Math.max(end, vpEnd);
      primed = true;
      return mergeViewportRange(lo, hi, start, end, vpStart, vpEnd, total, maxSpan);
    }

    lo = Math.min(lo, start, vpStart);
    hi = Math.max(hi, end, vpEnd);
    return mergeViewportRange(lo, hi, start, end, vpStart, vpEnd, total, maxSpan);
  }

  function resolveRows(scrollTop, viewportH, rowCount, overscan) {
    if (rowCount <= 0) return { start: 0, end: 0 };
    const cap = monotonic ? maxSpan : maxSpan;
    const { start, end, viewportStart, viewportEnd } = visibleRowRange(
      scrollTop,
      viewportH,
      rowCount,
      overscan,
      cap
    );
    return resolveVisible(start, end, rowCount, viewportStart, viewportEnd);
  }

  function resolveCols(layout, scrollLeft, viewportW, bufferPx) {
    if (!layout.length) return { start: 0, end: 0 };
    const min = scrollLeft - bufferPx;
    const max = scrollLeft + viewportW + bufferPx;
    const { start, end } = findVisibleColIndexRange(layout, min, max);
    return resolveVisible(start, end, layout.length, start, end);
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

/** BD row windowing; Synthesis uses forceOff (all rows in DOM). */
export function shouldVirtualizeRows(rowCount, viewportH, { forceOff = false } = {}) {
  if (forceOff) return false;
  const visible = Math.ceil(viewportH / ROW_H);
  return rowCount > visible + 24;
}

export function shouldVirtualizeCols(tableWidth, viewportW) {
  return tableWidth > viewportW + 120;
}

/** Sync scroll position immediately — virtual window must track scroll (no 1-frame lag). */
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
    const el = getScrollEl ? getScrollEl() : null;
    if (el) apply(el);
  }

  function dispose() {}

  return { onScroll, flush, dispose };
}
