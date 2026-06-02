/**
 * Web Worker: HyperFormula recalc off the UI thread.
 */
import { ensureHyperFormula } from './hyperformulaLoad.js?v=hf-worker2';
import { WorkbookEngine } from './workbookEngineCore.js?v=hf-worker2';

let engine = null;

function reply(id, ok, payload) {
  self.postMessage({ id, ok, ...payload });
}

self.onmessage = async (event) => {
  const { id, type } = event.data;
  try {
    switch (type) {
      case 'load': {
        const HF = await ensureHyperFormula();
        if (engine) engine.destroy();
        engine = new WorkbookEngine(HF);
        const result = engine.loadSheetData(
          event.data.name,
          event.data.sheetJson
        );
        reply(id, true, { result });
        break;
      }
      case 'setCell': {
        if (!engine) throw new Error('Engine not loaded');
        const { values } = engine.setCellValue(
          event.data.name,
          event.data.row,
          event.data.col,
          event.data.value
        );
        reply(id, true, { values });
        break;
      }
      case 'destroy': {
        if (engine) {
          engine.destroy();
          engine = null;
        }
        reply(id, true, {});
        break;
      }
      default:
        throw new Error(`Unknown worker message: ${type}`);
    }
  } catch (err) {
    reply(id, false, {
      error: (err && err.message) || String(err),
    });
  }
};
