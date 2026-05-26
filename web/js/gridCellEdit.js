/**
 * In-cell edit: show sheet value via :value; while focused, uncontrolled (fast typing).
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

/**
 * @param {{
 *   isNumericAt: (row: number, col: string) => boolean,
 *   seedAt: (row: number, col: string) => string,
 *   displayAt: (row: number, col: string) => string,
 *   commitAt: (row: number, col: string, value: string) => void,
 * }} opts
 */
export function createGridCellEditor(opts) {
  const { isNumericAt, seedAt, displayAt, commitAt } = opts;
  const activeKey = ref(null);
  const lastGoodByEl = new WeakMap();

  function isCellActive(row, col) {
    return activeKey.value === cellKey(row, col);
  }

  /** Vue-bound display value — skip while this cell is being edited (uncontrolled). */
  function boundValue(row, col, displayed) {
    if (isCellActive(row, col)) return undefined;
    return displayed;
  }

  function onCellFocus(row, col, event) {
    const el = event.target;
    activeKey.value = cellKey(row, col);
    const seed = seedAt(row, col);
    el.value = seed;
    lastGoodByEl.set(el, seed);
    try {
      el.select();
    } catch {
      /* ignore */
    }
  }

  function onCellInput(row, col, event) {
    const el = event.target;
    const next = el.value;
    if (!isNumericAt(row, col)) {
      lastGoodByEl.set(el, next);
      return;
    }
    if (!isAllowedNumericTyping(next)) {
      window.alert(NUMERIC_VALUE_ALERT_MSG);
      el.value = lastGoodByEl.get(el) ?? '';
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
    if (!parsed.ok) {
      window.alert(NUMERIC_VALUE_ALERT_MSG);
      el.value = displayAt(row, col);
      lastGoodByEl.delete(el);
      return false;
    }
    commitAt(row, col, parsed.value);
    lastGoodByEl.delete(el);
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
    onCellFocus,
    onCellInput,
    onCellBlur,
    onCellKeydown,
  };
}
