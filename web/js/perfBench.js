/**
 * Console perf bench — run: window.__runPerfBench?.()
 * Requires app mounted with session + grids loaded.
 */

export function createPerfBench(getState) {
  async function time(label, fn) {
    const t0 = performance.now();
    const result = await fn();
    const ms = Math.round(performance.now() - t0);
    console.log(`[perf] ${label}: ${ms} ms`);
    return { ms, result };
  }

  async function run() {
    const { session, bdSheet, synthesisSheet, applyBdFromRaw, applySynFromRaw } =
      getState();
    console.group('[perf] Vehicle Mass Platform bench');
    await time('getDisplayValue (Syn C15)', () => {
      session?.getDisplayValue?.('SYNTHESIS', 15, 'H', null);
    });
    if (applyBdFromRaw) {
      await time('transformBdSheet (cached miss path)', async () => {
        await applyBdFromRaw(true);
      });
    }
    if (applySynFromRaw && synthesisSheet) {
      await time('transformSynthesisSheet (force)', async () => {
        await applySynFromRaw(true);
      });
    }
    console.groupEnd();
    return 'Done — see golden cells in docs/golden-cells.md';
  }

  return { run };
}
