/**
 * Main-thread facade: HyperFormula runs in workbookEngine.worker.js.
 */
import { HyperFormula } from 'hyperformula';
import { WorkbookEngine } from './workbookEngineCore.js?v=hf-worker2';

const WORKER_URL = '/js/workbookEngine.worker.js?v=hf-worker2';

export class WorkbookEngineClient {
  constructor() {
    this._useWorker = true;
    this._worker = null;
    this._seq = 0;
    this._pending = new Map();
    /** @type {Map<string, string>} */
    this.valueCache = new Map();
    /** @type {Map<string, Set<string>>} */
    this.formulaKeysBySheet = new Map();
    this._fallback = null;
    this._initWorker();
  }

  _initWorker() {
    try {
      this._worker = new Worker(WORKER_URL, { type: 'module' });
      this._worker.onmessage = (e) => this._onMessage(e.data);
      this._worker.onerror = (e) => {
        console.error('Workbook engine worker error:', e);
        this._failPending(e.message || 'Worker error');
      };
    } catch (err) {
      console.warn('Web Worker unavailable, using main-thread engine:', err);
      this._useWorker = false;
      this._fallback = new WorkbookEngine(HyperFormula);
    }
  }

  _failPending(msg) {
    for (const [, { reject }] of this._pending) {
      reject(new Error(msg));
    }
    this._pending.clear();
  }

  _onMessage(data) {
    const { id, ok, result, values, error } = data;
    const pending = this._pending.get(id);
    if (!pending) return;
    this._pending.delete(id);
    if (!ok) {
      pending.reject(new Error(error || 'Worker failed'));
      return;
    }
    if (result) {
      this._applyLoadResult(result);
      pending.resolve(result);
      return;
    }
    if (values) {
      this._applyValues(values);
    }
    pending.resolve(values ?? {});
  }

  _call(type, payload) {
    if (!this._useWorker) {
      return this._callFallback(type, payload);
    }
    return new Promise((resolve, reject) => {
      const id = ++this._seq;
      this._pending.set(id, { resolve, reject });
      this._worker.postMessage({ id, type, ...payload });
    });
  }

  _callFallback(type, payload) {
    const eng = this._fallback;
    if (!eng) return Promise.reject(new Error('No engine'));
    switch (type) {
      case 'load': {
        const result = eng.loadSheetData(payload.name, payload.sheetJson);
        this._applyLoadResult(result);
        return Promise.resolve(result);
      }
      case 'setCell': {
        const { values } = eng.setCellValue(
          payload.name,
          payload.row,
          payload.col,
          payload.value
        );
        this._applyValues(values);
        return Promise.resolve({ values });
      }
      case 'destroy':
        eng.destroy();
        this.valueCache.clear();
        this.formulaKeysBySheet.clear();
        return Promise.resolve({});
      default:
        return Promise.reject(new Error(`Unknown: ${type}`));
    }
  }

  _applyLoadResult(result) {
    const keys = result.formulaKeys || [];
    this.formulaKeysBySheet.set('BD', new Set(keys));
    this._applyValues(result.values || {});
  }

  _applyValues(values) {
    for (const [key, v] of Object.entries(values)) {
      this.valueCache.set(key, v);
    }
  }

  async loadSheetData(name, sheetJson) {
    const result = await this._call('load', { name, sheetJson });
    return result.cellCount ?? 0;
  }

  async setCellValue(name, row, col, value) {
    await this._call('setCell', { name, row, col, value });
  }

  getCellValue(name, row, col) {
    if (this._useWorker || this._fallback) {
      const cached = this.valueCache.get(`${row}:${col}`);
      if (cached !== undefined) return cached;
      if (this._fallback) {
        return this._fallback.getCellValue(name, row, col);
      }
    }
    return '';
  }

  hasFormula(name, row, col) {
    return (
      this.formulaKeysBySheet.get(name)?.has(`${row}:${col}`) ?? false
    );
  }

  destroy() {
    if (this._worker) {
      void this._call('destroy', {}).finally(() => {
        this._worker.terminate();
        this._worker = null;
      });
    } else if (this._fallback) {
      this._fallback.destroy();
      this._fallback = null;
    }
    this.valueCache.clear();
    this.formulaKeysBySheet.clear();
    this._failPending('destroyed');
  }
}
