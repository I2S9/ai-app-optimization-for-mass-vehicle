/**
 * HyperFormula workbook logic (main thread or Web Worker).
 */
import {
  colToIndex,
  formatEngineValue,
  parseLiteral,
  resolveFormula,
  toEngineAddress,
} from './formulaUtil.js';

const HF_OPTIONS = {
  licenseKey: 'gpl-v3',
  useArrayArithmetic: true,
};

export function sheetToGrid(sheetJson) {
  const colNums = (sheetJson.columns || []).map((c) => colToIndex(c) + 1);
  const maxCol = Math.max(0, ...colNums);
  const maxRow = sheetJson.lastRow || 1;
  const grid = Array.from({ length: maxRow }, () =>
    Array(maxCol).fill(null)
  );
  const byRef = new Map();
  for (const cell of sheetJson.cells || []) {
    byRef.set(`${cell.c}${cell.r}`, cell);
  }
  for (const [rowStr, cols] of Object.entries(sheetJson.headerRows || {})) {
    const row = parseInt(rowStr, 10);
    if (Number.isNaN(row)) continue;
    for (const [col, cell] of Object.entries(cols)) {
      byRef.set(`${col}${row}`, { r: row, c: col, v: cell.v, f: cell.f });
    }
  }
  for (const cell of byRef.values()) {
    const ci = colToIndex(cell.c);
    if (cell.r < 1 || cell.r > maxRow || ci < 0 || ci >= maxCol) continue;
    if (cell.f) continue;
    const v = cell.v;
    grid[cell.r - 1][ci] =
      v === '#REF!' || v === '' ? null : parseLiteral(v);
  }
  for (const cell of byRef.values()) {
    if (!cell.f) continue;
    const ci = colToIndex(cell.c);
    if (cell.r < 1 || cell.r > maxRow || ci < 0 || ci >= maxCol) continue;
    const f = resolveFormula(`${cell.c}${cell.r}`, byRef);
    if (f) grid[cell.r - 1][ci] = f;
  }
  return grid;
}

function indexFormulas(sheetJson) {
  const keys = [];
  const add = (r, c, f) => {
    if (f) keys.push(`${r}:${c}`);
  };
  for (const cell of sheetJson.cells || []) add(cell.r, cell.c, cell.f);
  for (const [rowStr, cols] of Object.entries(sheetJson.headerRows || {})) {
    const row = parseInt(rowStr, 10);
    for (const [col, cell] of Object.entries(cols)) add(row, col, cell.f);
  }
  return keys;
}

export class WorkbookEngine {
  constructor(HyperFormula) {
    if (!HyperFormula) {
      throw new Error('HyperFormula constructor required');
    }
    this.HyperFormula = HyperFormula;
    this.hf = null;
    this.sheetIds = new Map();
    this.formulaKeys = new Map();
  }

  loadSheetData(name, sheetJson) {
    if (name !== 'BD') return { cellCount: 0, formulaKeys: [], values: {} };
    const grid = sheetToGrid(sheetJson);
    this.hf = this.HyperFormula.buildFromSheets({ BD: grid }, HF_OPTIONS);
    const sheetId = this.hf.getSheetId('BD');
    if (sheetId == null) {
      throw new Error('HyperFormula did not create the BD sheet');
    }
    this.sheetIds.set('BD', sheetId);
    const keys = indexFormulas(sheetJson);
    this.formulaKeys.set('BD', new Set(keys));
    return {
      cellCount: sheetJson.cells && sheetJson.cells.length ? sheetJson.cells.length : 0,
      formulaKeys: keys,
      values: {},
    };
  }

  dumpFormulaValues(name) {
    const sheetId = this.sheetIds.get(name);
    const keys = this.formulaKeys.get(name);
    if (sheetId == null || !keys || !keys.size) return {};
    const values = {};
    try {
      const serialized = this.hf.getSheetValues(sheetId);
      for (const key of keys) {
        const colon = key.indexOf(':');
        const row = parseInt(key.slice(0, colon), 10);
        const col = key.slice(colon + 1);
        const ci = colToIndex(col);
        const raw = serialized[row - 1] ? serialized[row - 1][ci] : undefined;
        const v = formatEngineValue(raw);
        if (v !== '' && v !== '#REF!') values[key] = v;
      }
      return values;
    } catch (e) {
      for (const key of keys) {
        const colon = key.indexOf(':');
        const row = parseInt(key.slice(0, colon), 10);
        const col = key.slice(colon + 1);
        const v = this.getCellValue(name, row, col);
        if (v !== '' && v !== '#REF!') values[key] = v;
      }
      return values;
    }
  }

  setCellValue(name, row, col, value) {
    const sheetId = this.sheetIds.get(name);
    if (sheetId == null) return { values: {} };
    const addr = toEngineAddress(sheetId, row, col);
    this.hf.setCellContents(addr, [[parseLiteral(value)]]);
    return { values: this.dumpFormulaValues(name) };
  }

  getCellValue(name, row, col) {
    const sheetId = this.sheetIds.get(name);
    if (sheetId == null) return '';
    const raw = this.hf.getCellValue(toEngineAddress(sheetId, row, col));
    return formatEngineValue(raw);
  }

  hasFormula(name, row, col) {
    const set = this.formulaKeys.get(name);
    if (!set) return false;
    return set.has(`${row}:${col}`);
  }

  destroy() {
    if (this.hf) {
      this.hf.destroy();
      this.hf = null;
    }
    this.sheetIds.clear();
    this.formulaKeys.clear();
  }
}
