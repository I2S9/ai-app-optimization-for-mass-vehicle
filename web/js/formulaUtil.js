/** Excel address helpers and formula normalization for HyperFormula. */

export function colToIndex(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n - 1;
}

export function indexToCol(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function toEngineAddress(sheetId, row, col) {
  return { sheet: sheetId, row: row - 1, col: colToIndex(col) };
}

/** Excel OOXML often stores formulas without a leading '='. */
export function normalizeFormula(raw) {
  let f = String(raw || '').trim().replace(/#REF!/g, '0');
  if (!f) return '';
  if (f.startsWith('=')) return f;
  return `=${f}`;
}

/** Shared-formula anchor (e.g. f: "AG6") → real formula from master cell. */
export function resolveFormula(ref, byRef, vis = new Set()) {
  const cell = byRef.get(ref);
  if (!cell || !cell.f) return null;
  const anchor = String(cell.f).trim();
  if (!/^[A-Z]+\d+$/i.test(anchor)) return normalizeFormula(anchor);
  if (vis.has(anchor)) return null;
  vis.add(anchor);
  return resolveFormula(anchor, byRef, vis);
}

export function parseLiteral(value) {
  if (value == null || value === '') return '';
  const s = String(value).trim();
  if (s === 'TRUE' || s === 'true') return true;
  if (s === 'FALSE' || s === 'false') return false;
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  return s;
}

export function formatEngineValue(value) {
  if (value == null || value === '') return '';
  if (typeof value === 'object' && value !== null) {
    if (value.type === 'ERROR') return value.value || '#ERROR!';
    if (value.type === 'NUMBER') return String(value.value);
  }
  if (typeof value === 'number') {
    if (Number.isInteger(value)) return String(value);
    const rounded = Math.round(value * 1e6) / 1e6;
    return String(rounded);
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}
