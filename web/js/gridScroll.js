/** Shared virtual-scroll helpers for BD / Synthesis grids. */

export const ROW_H = 21;

/** Rows to render above/below the viewport (≈2× visible height). */
export function rowOverscan(viewportH, minRows = 30) {
  const visible = Math.max(1, Math.ceil(viewportH / ROW_H));
  return Math.max(minRows, visible * 2);
}

/** Horizontal pixels to render left/right of the viewport. */
export function colOverscanPx(viewportW, minPx = 960) {
  return Math.max(minPx, Math.floor(viewportW * 2));
}

export function visibleRowRange(scrollTop, viewportH, rowCount, overscan) {
  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - overscan);
  const count = Math.ceil(viewportH / ROW_H) + overscan * 2;
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
