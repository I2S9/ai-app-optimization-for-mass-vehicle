/** Shared virtual-scroll helpers for BD / Synthesis grids. */

export const ROW_H = 21;

/**
 * Hard cap on mounted rows when row virtualization is active (Database sheet).
 */
export const MAX_RENDERED_ROWS = 220;
export const SYN_MAX_RENDERED_ROWS = 120;

/** Rows to render above/below the viewport. */
export function rowOverscan(viewportH, minRows = 20, maxRows = 56) {
  const visible = Math.max(1, Math.ceil(viewportH / ROW_H));
  return Math.min(maxRows, Math.max(minRows, Math.ceil(visible * 0.85)));
}

/** Tighter row buffer when many columns are on screen (Synthesis). */
export function rowOverscanForColCount(viewportH, colCount) {
  const base = rowOverscan(viewportH);
  if (colCount > 120) return Math.min(base, 20);
  if (colCount > 60) return Math.min(base, 28);
  return base;
}

/** Horizontal pixels to render left/right of the viewport. */
export function colOverscanPx(viewportW, minPx = 480) {
  return Math.min(2400, Math.max(minPx, Math.floor(viewportW * 1.25)));
}

/** Tighter horizontal buffer for wide Synthesis sheets. */
export function synColOverscanPx(viewportW) {
  return Math.min(960, Math.max(280, Math.floor(viewportW * 0.6)));
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
 * Expanding row window: once a row range was visited it stays mounted (until the
 * cap forces a trim around the viewport). Fixes "grey band + reload" when
 * scrolling back up on Database.
 */
export function createRowScrollCache(maxSpan = MAX_RENDERED_ROWS) {
  let lo = 0;
  let hi = 0;
  let primed = false;

  function reset() {
    lo = 0;
    hi = 0;
    primed = false;
  }

  function resolve(scrollTop, viewportH, rowCount, overscan) {
    if (rowCount <= 0) return { start: 0, end: 0 };

    const { start, end } = visibleRowRange(
      scrollTop,
      viewportH,
      rowCount,
      overscan
    );
    if (!primed) {
      lo = start;
      hi = end;
      primed = true;
      return { start: lo, end: hi };
    }

    lo = Math.min(lo, start);
    hi = Math.max(hi, end);
    hi = Math.min(hi, rowCount);

    const span = hi - lo;
    if (span > maxSpan) {
      const viewStart = Math.max(0, Math.floor(scrollTop / ROW_H) - overscan);
      const viewEnd = Math.min(
        rowCount,
        viewStart + Math.ceil(viewportH / ROW_H) + overscan * 2
      );
      const viewMid = (viewStart + viewEnd) >> 1;
      const half = Math.floor(maxSpan / 2);
      lo = Math.max(0, viewMid - half);
      hi = Math.min(rowCount, lo + maxSpan);
      if (hi - lo < maxSpan) lo = Math.max(0, hi - maxSpan);
    }

    return { start: lo, end: hi };
  }

  return { resolve, reset };
}

/** Row windowing when the body is taller than the viewport buffer. */
export function shouldVirtualizeRows(rowCount, viewportH) {
  const visible = Math.ceil(viewportH / ROW_H);
  return rowCount > visible + 24;
}

/** Skip column windowing when the table fits in the viewport. */
export function shouldVirtualizeCols(tableWidth, viewportW) {
  return tableWidth > viewportW + 120;
}

/**
 * Scroll position is applied synchronously so the virtual row window matches
 * the native scroll position (no grey spacer flash between frames).
 */
export function createScrollRafSync(refs) {
  const { scrollTop, scrollLeft = null } = refs;

  function onScroll(e) {
    scrollTop.value = e.target.scrollTop;
    if (scrollLeft) scrollLeft.value = e.target.scrollLeft;
  }

  function flush() {
    const el = refs.getScrollEl?.() ?? null;
    if (!el) return;
    scrollTop.value = el.scrollTop;
    if (scrollLeft) scrollLeft.value = el.scrollLeft;
  }

  function dispose() {}

  return { onScroll, flush, dispose };
}
