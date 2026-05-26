/** Shared virtual-scroll helpers for BD / Synthesis grids. */

export const ROW_H = 21;

/** Hard cap so tall viewports never mount the whole sheet (Synthesis ≈ 400 rows × 300+ cols). */
export const MAX_RENDERED_ROWS = 72;

/** Rows to render above/below the viewport. */
export function rowOverscan(viewportH, minRows = 10, maxRows = 24) {
  const visible = Math.max(1, Math.ceil(viewportH / ROW_H));
  return Math.min(maxRows, Math.max(minRows, Math.ceil(visible * 0.75)));
}

/** Tighter row buffer when many columns are on screen (Synthesis). */
export function rowOverscanForColCount(viewportH, colCount) {
  const base = rowOverscan(viewportH);
  if (colCount > 120) return Math.min(base, 12);
  if (colCount > 60) return Math.min(base, 16);
  return base;
}

/** Horizontal pixels to render left/right of the viewport. */
export function colOverscanPx(viewportW, minPx = 480) {
  return Math.min(2400, Math.max(minPx, Math.floor(viewportW * 1.25)));
}

export function visibleRowRange(scrollTop, viewportH, rowCount, overscan) {
  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - overscan);
  let count = Math.ceil(viewportH / ROW_H) + overscan * 2;
  count = Math.min(count, MAX_RENDERED_ROWS);
  const end = Math.min(rowCount, start + count);
  return { start, end };
}

/** Skip row virtualization when the sheet is small enough to paint in one pass. */
export function shouldVirtualizeRows(rowCount, viewportH) {
  const visible = Math.ceil(viewportH / ROW_H);
  return rowCount > visible + 60;
}

/** Skip column windowing when the table fits in the viewport. */
export function shouldVirtualizeCols(tableWidth, viewportW) {
  return tableWidth > viewportW + 120;
}

/**
 * Coalesce scroll events to one reactive update per animation frame
 * (reduces Vue re-renders and row recycle flicker while scrolling).
 */
export function createScrollRafSync(refs) {
  const { scrollTop, scrollLeft = null } = refs;
  let rafId = 0;
  let pendingTop = 0;
  let pendingLeft = 0;

  function apply() {
    rafId = 0;
    scrollTop.value = pendingTop;
    if (scrollLeft) scrollLeft.value = pendingLeft;
  }

  function onScroll(e) {
    pendingTop = e.target.scrollTop;
    if (scrollLeft) pendingLeft = e.target.scrollLeft;
    if (!rafId) rafId = requestAnimationFrame(apply);
  }

  function flush() {
    if (rafId) {
      cancelAnimationFrame(rafId);
      rafId = 0;
    }
    apply();
  }

  function dispose() {
    if (rafId) {
      cancelAnimationFrame(rafId);
      rafId = 0;
    }
  }

  return { onScroll, flush, dispose };
}
