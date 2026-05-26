/**
 * In-cell edit: one active input; inactive cells are display spans (no :value overwrite).
 */

import { ref } from 'vue';

export const NUMERIC_VALUE_ALERT_MSG =
  'The entered value must be a numeric value.';

export function isAllowedNumericTyping(value) {
  const s = String(value ?? '');
  if (s === '' || s === '-') return true;
  if (!/^-?\d*([.,]?\d*)?$/.test(s)) return false;
  const body = s.startsWith('-') ? s.slice(1) : s;
  return (body.match(/[.,]/g) || []).length <= 1;
}

export function isCompleteNumericValue(value) {
  const s = String(value ?? '').trim();
  if (s === '') return true;
  if (s === '-') return false;
  if (!/^-?\d+([.,]\d+)?$/.test(s)) return false;
  const body = s.replace(/^-/, '');
  return (body.match(/[.,]/g) || []).length <= 1;
}

function cellKey(row, col) {
  return `${row}:${col}`;
}

function placeCaretAtEnd(el) {
  const len = el.value.length;
  try {
    el.setSelectionRange(len, len);
  } catch {
    /* ignore */
  }
}

/**
 * @param {{
 *   isNumericAt: (row: number, col: string) => boolean,
 *   displayAt: (row: number, col: string) => string,
 *   commitAt: (row: number, col: string, value: string, previousValue: string) => void,
 * }} opts
 */
export function createGridCellEditor(opts) {
  const { isNumericAt, displayAt, commitAt } = opts;
  const activeKey = ref(null);
  const lastGoodByEl = new WeakMap();
  const dirtyByEl = new WeakMap();
  const startSeedByEl = new WeakMap();
  const idleDisplayByKey = new Map();

  let navigationLock = false;
  let suppressInputDirty = 0;

  function isCellActive(row, col) {
    return activeKey.value === cellKey(row, col);
  }

  function rememberIdleDisplay(row, col, value) {
    idleDisplayByKey.set(cellKey(row, col), value);
  }

  function forgetIdleDisplay(row, col) {
    idleDisplayByKey.delete(cellKey(row, col));
  }

  function cellShowValue(row, col, displayed) {
    const k = cellKey(row, col);
    if (idleDisplayByKey.has(k)) return idleDisplayByKey.get(k);
    return displayed;
  }

  function clearEditState(el) {
    lastGoodByEl.delete(el);
    dirtyByEl.delete(el);
    startSeedByEl.delete(el);
  }

  function applyEditSeed(el, row, col) {
    suppressInputDirty += 1;
    try {
      forgetIdleDisplay(row, col);
      const seed = displayAt(row, col);
      el.value = seed;
      lastGoodByEl.set(el, seed);
      startSeedByEl.set(el, seed);
      dirtyByEl.set(el, false);
      placeCaretAtEnd(el);
      return seed;
    } finally {
      suppressInputDirty -= 1;
    }
  }

  function onCellMouseDown(row, col, event) {
    if (event.button !== 0) return;
    applyEditSeed(event.target, row, col);
  }

  function onCellSpanMouseDown(row, col, event) {
    if (event.button !== 0) return;
    event.preventDefault();
    const td = event.currentTarget.closest('td');
    if (!td) return;
    activeKey.value = cellKey(row, col);
    queueMicrotask(() => {
      const input = td.querySelector('input.grid-cell-input');
      if (input instanceof HTMLInputElement) {
        applyEditSeed(input, row, col);
        input.focus({ preventScroll: true });
      }
    });
  }

  function onCellFocus(row, col, event) {
    const el = event.target;
    activeKey.value = cellKey(row, col);
    dirtyByEl.set(el, false);
    if (navigationLock) return;
    applyEditSeed(el, row, col);
    queueMicrotask(() => {
      if (navigationLock) return;
      if (activeKey.value !== cellKey(row, col)) return;
      applyEditSeed(el, row, col);
    });
  }

  function onCellInput(row, col, event) {
    if (suppressInputDirty > 0 || navigationLock) return;
    const el = event.target;
    dirtyByEl.set(el, true);
    const next = el.value;
    if (!isNumericAt(row, col)) {
      lastGoodByEl.set(el, next);
      return;
    }
    if (!isAllowedNumericTyping(next)) {
      window.alert(NUMERIC_VALUE_ALERT_MSG);
      el.value = lastGoodByEl.get(el) ?? '';
      placeCaretAtEnd(el);
      return;
    }
    lastGoodByEl.set(el, next);
  }

  function normalizeCommit(row, col, raw) {
    if (!isNumericAt(row, col)) {
      return { ok: true, value: raw };
    }
    const v = String(raw ?? '').trim();
    if (!isCompleteNumericValue(v)) {
      return { ok: false };
    }
    return { ok: true, value: v };
  }

  function finishEdit(row, col, event, options = {}) {
    const { deferDeactivate = false } = options;
    const el = event.target;
    const parsed = normalizeCommit(row, col, el.value);
    const wasDirty = dirtyByEl.get(el) === true;

    const deactivate = () => {
      if (!deferDeactivate) activeKey.value = null;
    };

    if (!wasDirty) {
      rememberIdleDisplay(row, col, el.value);
      clearEditState(el);
      deactivate();
      return true;
    }

    if (!parsed.ok) {
      window.alert(NUMERIC_VALUE_ALERT_MSG);
      const fallback = displayAt(row, col);
      el.value = fallback;
      rememberIdleDisplay(row, col, fallback);
      clearEditState(el);
      deactivate();
      return false;
    }

    const startSeed = startSeedByEl.get(el) ?? '';
    if (String(parsed.value).trim() === String(startSeed).trim()) {
      rememberIdleDisplay(row, col, el.value);
      clearEditState(el);
      deactivate();
      return true;
    }

    commitAt(row, col, parsed.value, startSeed);
    const shown = String(parsed.value);
    rememberIdleDisplay(row, col, shown);
    clearEditState(el);
    deactivate();
    return true;
  }

  function prepareNavigate(row, col, event) {
    const el = event.target;
    if (dirtyByEl.get(el) === true) {
      return finishEdit(row, col, event, { deferDeactivate: true });
    }
    const preserved = startSeedByEl.get(el) ?? el.value;
    rememberIdleDisplay(row, col, preserved);
    clearEditState(el);
    return true;
  }

  function beginNavigationTo(row, col) {
    activeKey.value = cellKey(row, col);
  }

  function activateCell(row, col, el) {
    if (!(el instanceof HTMLInputElement)) return;
    activeKey.value = cellKey(row, col);
    applyEditSeed(el, row, col);
    el.focus({ preventScroll: true });
  }

  function onCellBlur(row, col, event) {
    if (navigationLock) return;
    if (activeKey.value !== cellKey(row, col)) return;
    finishEdit(row, col, event);
  }

  function onCellKeydown(row, col, event) {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    if (finishEdit(row, col, event)) {
      event.target.blur();
    }
  }

  return {
    activeKey,
    isCellActive,
    cellShowValue,
    onCellMouseDown,
    onCellSpanMouseDown,
    onCellFocus,
    onCellInput,
    onCellBlur,
    onCellKeydown,
    finishEdit,
    prepareNavigate,
    beginNavigationTo,
    activateCell,
    setNavigationLock: (v) => {
      navigationLock = v;
    },
  };
}
