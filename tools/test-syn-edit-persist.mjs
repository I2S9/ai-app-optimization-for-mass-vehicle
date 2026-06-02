/**
 * Regression: Synthesis label edits must land on raw JSON before snapshot build.
 */
import {
  buildPersistRecord,
  extractSheetEdits,
  syncGridEditsToRaw,
} from '../web/js/sessionPersistence.js';

let failed = 0;
function ok(msg) {
  console.log(`  ok: ${msg}`);
}
function fail(msg) {
  failed++;
  console.error(`  FAIL: ${msg}`);
}

const raw = {
  cells: [{ r: 27, c: 'A', v: 'OLD' }],
  columns: ['A'],
  lastRow: 50,
};
const grid = {
  cells: [{ r: 27, c: 'A', v: 'NEW LABEL', userEdited: true }],
};

syncGridEditsToRaw(grid, raw);
const hit = raw.cells.find((c) => c.r === 27 && c.c === 'A');
if (!hit || hit.v !== 'NEW LABEL' || !hit.userEdited) {
  fail('syncGridEditsToRaw did not copy userEdited grid cell to raw');
} else {
  ok('syncGridEditsToRaw copies grid edits to raw');
}

const rec = buildPersistRecord({ bd: null, syn: raw, revision: 1 });
const ed = rec && rec.synEdits;
if (!ed || !ed.cells.some((c) => c.r === 27 && c.c === 'A' && c.v === 'NEW LABEL')) {
  fail('buildPersistRecord missing syn cell edit after sync');
} else {
  ok('snapshot includes renamed synthesis cell');
}

if (failed) {
  console.error(`\n${failed} check(s) failed`);
  process.exit(1);
}
console.log('\nALL SYN-EDIT PERSIST CHECKS PASSED');
