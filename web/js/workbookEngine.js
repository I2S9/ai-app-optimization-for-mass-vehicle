import { HyperFormula } from 'hyperformula';
import {
  colToIndex,
  formatEngineValue,
  normalizeFormula,
  parseLiteral,
  resolveFormula,
  toEngineAddress,
} from './formulaUtil.js';

const HF_OPTIONS = {
  licenseKey: 'gpl-v3',
  useArrayArithmetic: true,
};

function sheetToGrid(sheetJson) {
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

export class WorkbookEngine {
  constructor() {
    this.hf = null;
    this.sheetIds = new Map();
    this.formulaKeys = new Map();
  }

  loadSheetData(name, sheetJson) {
    if (name !== 'BD') return 0;
    if (!HyperFormula) {
      throw new Error(
        'HyperFormula not loaded. Run scripts/setup-web-vendor.ps1 then restart the server.'
      );
    }
    const grid = sheetToGrid(sheetJson);
    this.hf = HyperFormula.buildFromSheets({ BD: grid }, HF_OPTIONS);
    const sheetId = this.hf.getSheetId('BD');
    if (sheetId == null) {
      throw new Error('HyperFormula did not create the BD sheet');
    }
    this.sheetIds.set('BD', sheetId);
    this._indexFormulas('BD', sheetJson);
    return sheetJson.cells?.length ?? 0;
  }

  registerSheet(name) {
    if (this.sheetIds.has(name)) return this.sheetIds.get(name);
    const id = this.hf.getSheetId(name);
    if (id != null) {
      this.sheetIds.set(name, id);
      return id;
    }
    const newId = this.hf.addSheet(name);
    this.sheetIds.set(name, newId);
    return newId;
  }

  _indexFormulas(name, sheetJson) {
    const keys = new Set();
    const add = (r, c, f) => {
      if (f) keys.add(`${r}:${c}`);
    };
    for (const cell of sheetJson.cells || []) add(cell.r, cell.c, cell.f);
    for (const [rowStr, cols] of Object.entries(sheetJson.headerRows || {})) {
      const row = parseInt(rowStr, 10);
      for (const [col, cell] of Object.entries(cols)) add(row, col, cell.f);
    }
    this.formulaKeys.set(name, keys);
  }

  setCellValue(name, row, col, value) {
    const sheetId = this.sheetIds.get(name);
    if (sheetId == null) return;
    const addr = toEngineAddress(sheetId, row, col);
    this.hf.setCellContents(addr, [[parseLiteral(value)]]);
  }

  getCellValue(name, row, col) {
    const sheetId = this.sheetIds.get(name);
    if (sheetId == null) return '';
    const raw = this.hf.getCellValue(toEngineAddress(sheetId, row, col));
    return formatEngineValue(raw);
  }

  hasFormula(name, row, col) {
    return this.formulaKeys.get(name)?.has(`${row}:${col}`) ?? false;
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
