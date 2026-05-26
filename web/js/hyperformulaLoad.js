/** Load HyperFormula UMD build in contexts without index.html script tag (Web Worker). */

const HF_SCRIPT_URL = '/vendor/hyperformula.full.min.js';

let loadPromise = null;

export function ensureHyperFormula() {
  if (globalThis.HyperFormula) {
    return Promise.resolve(globalThis.HyperFormula);
  }
  if (!loadPromise) {
    loadPromise = fetch(HF_SCRIPT_URL)
      .then((res) => {
        if (!res.ok) {
          throw new Error(
            `HyperFormula vendor missing (${res.status}). Run scripts/setup-web-vendor.ps1`
          );
        }
        return res.text();
      })
      .then((code) => {
        // Same-origin vendor bundle; runs in Worker or main.
        // eslint-disable-next-line no-new-func
        const fn = new Function(
          `${code}\nreturn typeof HyperFormula !== "undefined" ? HyperFormula : null;`
        );
        const HF = fn();
        if (!HF) {
          throw new Error('HyperFormula did not initialize');
        }
        globalThis.HyperFormula = HF;
        return HF;
      });
  }
  return loadPromise;
}
