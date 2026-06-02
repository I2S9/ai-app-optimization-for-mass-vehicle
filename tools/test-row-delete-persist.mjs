/**
 * Regression: row delete-only edits must count as persisted edits (F5 restore).
 */
import { hasSheetEdits } from '../web/js/gridBoot.js';
import {
  buildPersistRecord,
  applySheetEdits,
  extractSheetEdits,
  rawFingerprint,
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
  cells: [{ r: 10, c: 'A', v: 'x' }],
  columns: ['A'],
  lastRow: 20,
};

const editsOnlyDelete = { cells: [], headerRows: {}, deletedRows: [10] };
if (!hasSheetEdits(editsOnlyDelete)) {
  fail('hasSheetEdits ignores deletedRows-only snapshot');
} else {
  ok('hasSheetEdits sees deletedRows-only snapshot');
}

raw.deletedRows = [10];
const rec = buildPersistRecord({ bd: raw, syn: null, revision: 1 });
if (!rec || !rec.bdEdits || !rec.bdEdits.deletedRows.includes(10)) {
  fail('buildPersistRecord omits deletedRows');
} else {
  ok('buildPersistRecord stores deletedRows');
}

const fresh = {
  cells: [{ r: 10, c: 'A', v: 'x' }],
  columns: ['A'],
  lastRow: 20,
};
applySheetEdits(fresh, editsOnlyDelete);
if (!fresh.deletedRows || !fresh.deletedRows.includes(10)) {
  fail('applySheetEdits does not restore deletedRows');
} else {
  ok('applySheetEdits restores deletedRows');
}

const fp0 = rawFingerprint({ ...raw, deletedRows: [] });
const fp1 = rawFingerprint({ ...raw, deletedRows: [10] });
if (fp0 === fp1) {
  fail('rawFingerprint unchanged when only deletedRows change');
} else {
  ok('rawFingerprint includes deletedRows');
}

if (failed) {
  console.error(`\n${failed} check(s) failed`);
  process.exit(1);
}
console.log('\nALL ROW-DELETE PERSIST CHECKS PASSED');
