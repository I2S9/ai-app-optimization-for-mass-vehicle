/**
 * In-cell edit: display value on focus, caret at end, commit only when changed.
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
 *   commitAt: (row: number, col: string, value: string) => void,
 * }} opts
 */
export function createGridCellEditor(opts) {
  const { isNumericAt, displayAt, commitAt } = opts;
  const activeKey = ref(null);
  const lastGoodByEl = new WeakMap();
  const dirtyByEl = new WeakMap();
  const startSeedByEl = new WeakMap();

  function isCellActive(row, col) {
    return activeKey.value === cellKey(row, col);
  }

  /** Uncontrolled while editing so typing stays fast. */
  function boundValue(row, col, displayed) {
    if (isCellActive(row, col)) return undefined;
    return displayed;
  }

  function clearEditState(el) {
    lastGoodByEl.delete(el);
    dirtyByEl.delete(el);
    startSeedByEl.delete(el);
  }

  function applyEditSeed(el, row, col) {
    const seed = displayAt(row, col);
    el.value = seed;
    lastGoodByEl.set(el, seed);
    startSeedByEl.set(el, seed);
    placeCaretAtEnd(el);
    return seed;
  }

  /** Before focus — keeps visible text in the input when Vue drops :value. */
  function onCellMouseDown(row, col, event) {
    if (event.button !== 0) return;
    applyEditSeed(event.target, row, col);
  }

  function onCellFocus(row, col, event) {
    const el = event.target;
    activeKey.value = cellKey(row, col);
    dirtyByEl.set(el, false);
    applyEditSeed(el, row, col);
    queueMicrotask(() => {
      if (activeKey.value !== cellKey(row, col)) return;
      applyEditSeed(el, row, col);
    });
  }

  function onCellInput(row, col, event) {
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

  function finishEdit(row, col, event) {
    const el = event.target;
    const parsed = normalizeCommit(row, col, el.value);
    activeKey.value = null;
    const wasDirty = dirtyByEl.get(el) === true;

    if (!wasDirty) {
      el.value = displayAt(row, col);
      clearEditState(el);
      return true;
    }

    if (!parsed.ok) {
      window.alert(NUMERIC_VALUE_ALERT_MSG);
      el.value = displayAt(row, col);
      clearEditState(el);
      return false;
    }

    const startSeed = startSeedByEl.get(el) ?? '';
    if (String(parsed.value).trim() === String(startSeed).trim()) {
      el.value = displayAt(row, col);
      clearEditState(el);
      return true;
    }

    commitAt(row, col, parsed.value);
    clearEditState(el);
    el.value = displayAt(row, col);
    return true;
  }

  function onCellBlur(row, col, event) {
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
    boundValue,
    onCellMouseDown,
    onCellFocus,
    onCellInput,
    onCellBlur,
    onCellKeydown,
  };
}
