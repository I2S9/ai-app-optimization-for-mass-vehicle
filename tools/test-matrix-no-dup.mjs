/**
 * Guarantee test: rename / archive / reorder must UPDATE existing rows, never
 * duplicate sections or sub-sections in BD or Synthesis. Adding is the only
 * operation allowed to create rows. Mirrors the main.js apply path.
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const url = (p) => `file://${join(root, p).replace(/\\/g, '/')}`;
const bdRaw0 = JSON.parse(readFileSync(join(root, 'web/public/data/bd-sheet.json'), 'utf8'));
const synRaw0 = JSON.parse(readFileSync(join(root, 'web/public/data/synthesis-sheet.json'), 'utf8'));
const { transformBdSheet, transformSynthesisSheet } = await import(url('web/js/sheetTransform.js'));
const sm = await import(url('web/js/structureModel.js'));
const bdStore = await import(url('web/js/bdStore.js'));
const { synLabel } = await import(url('web/js/synStore.js'));
const clone = (x) => JSON.parse(JSON.stringify(x));

const bdSheet = transformBdSheet(bdRaw0);
const synSheet = transformSynthesisSheet(synRaw0);
const baseState = sm.buildMatrixState(bdSheet, synSheet);

const baseBdSecCount = baseState.bd.sections.length;
const baseSynSecCount = baseState.syn.sections.length;
let baseBdSubCount = 0; for (const s of baseState.bd.sections) baseBdSubCount += s.subsections.length;
let baseSynSubCount = 0; for (const s of baseState.syn.sections) baseSynSubCount += s.subsections.length;

let failed = 0;
function check(name, cond, detail) {
  if (!cond) { console.error(`  FAIL: ${name}${detail ? ' — ' + detail : ''}`); failed++; }
  else console.log(`  ok: ${name}`);
}

function rawStats(raw) {
  const rows = new Set();
  for (const c of raw.cells) rows.add(c.r);
  return { cells: raw.cells.length, rows: rows.size };
}

function bandOverlaps(model) {
  const bands = model.sections
    .filter((s) => s.headerRow != null)
    .map((s) => [s.headerRow, s.endRow != null ? s.endRow : s.headerRow])
    .sort((a, b) => a[0] - b[0]);
  let n = 0;
  for (let i = 1; i < bands.length; i++) if (bands[i][0] <= bands[i - 1][1]) n++;
  return n;
}

function labelCounts(model) {
  const sec = new Map();
  const sub = new Map();
  for (const s of model.sections) {
    sec.set(s.label, (sec.get(s.label) || 0) + 1);
    for (const ss of s.subsections) sub.set(`${s.label}|${ss.label}`, (sub.get(`${s.label}|${ss.label}`) || 0) + 1);
  }
  return { sec, sub };
}
function maxCount(map) { let m = 0; for (const v of map.values()) if (v > m) m = v; return m; }

// Run the real apply path: edited BD model -> aligned syn model -> applyMatrixSave
function runApply(editFn) {
  const bd = clone(baseState.bd);
  editFn(bd, sm);
  const synModel = sm.alignSynModelToBd(clone(baseState.syn), bd);
  return sm.applyMatrixSave(clone(bdRaw0), clone(synRaw0), clone(bd), synModel);
}

function assertNoDup(label, result, { expectBdSec, expectSynSec, expectBdSub, expectSynSub } = {}) {
  console.log(`\n[${label}]`);
  const bdL = labelCounts(result.bdModel);
  const synL = labelCounts(result.synModel);
  check('BD no duplicate section label', maxCount(bdL.sec) <= 1, `max=${maxCount(bdL.sec)}`);
  check('BD no duplicate subsection label', maxCount(bdL.sub) <= 1, `max=${maxCount(bdL.sub)}`);
  check('SYN no duplicate section label', maxCount(synL.sec) <= 1, `max=${maxCount(synL.sec)}`);
  check('SYN no duplicate subsection label', maxCount(synL.sub) <= 1, `max=${maxCount(synL.sub)}`);
  check('BD no band overlap', bandOverlaps(result.bdModel) === 0);
  check('SYN no band overlap', bandOverlaps(result.synModel) === 0);
  if (expectBdSec != null) check('BD section count', result.bdModel.sections.length === expectBdSec, `got ${result.bdModel.sections.length} want ${expectBdSec}`);
  if (expectSynSec != null) check('SYN section count', result.synModel.sections.filter(s=>s.headerRow!=null).length === expectSynSec, `got ${result.synModel.sections.filter(s=>s.headerRow!=null).length} want ${expectSynSec}`);
  const bdSub = result.bdModel.sections.reduce((a, s) => a + s.subsections.length, 0);
  if (expectBdSub != null) check('BD subsection count', bdSub === expectBdSub, `got ${bdSub} want ${expectBdSub}`);
  return result;
}

// pick a BD section that has at least one subsection
const secWithSubIdx = baseState.bd.sections.findIndex((s) => s.subsections.length > 0);

// 1) RENAME a section — counts unchanged, no dup
assertNoDup('rename section', runApply((bd) => { bd.sections[2].label = 'RENAMED-SECTION-X'; }),
  { expectBdSec: baseBdSecCount, expectSynSec: baseSynSecCount, expectBdSub: baseBdSubCount });

// 2) RENAME a subsection — counts unchanged, no dup
assertNoDup('rename subsection', runApply((bd) => { bd.sections[secWithSubIdx].subsections[0].label = '_RENAMED_SUB_X'; }),
  { expectBdSec: baseBdSecCount, expectSynSec: baseSynSecCount, expectBdSub: baseBdSubCount });

// 3) ARCHIVE a subsection
assertNoDup('archive subsection', runApply((bd, m) => { m.archiveSubsection(bd, bd.sections[secWithSubIdx].subsections[0].id); }),
  { expectBdSec: baseBdSecCount, expectSynSec: baseSynSecCount, expectBdSub: baseBdSubCount });

// 4) ARCHIVE a section
assertNoDup('archive section', runApply((bd, m) => { m.archiveSection(bd, bd.sections[3].id); }),
  { expectBdSec: baseBdSecCount, expectSynSec: baseSynSecCount, expectBdSub: baseBdSubCount });

// 5) REORDER two sections (swap)
assertNoDup('reorder sections', runApply((bd) => { const t = bd.sections[2]; bd.sections[2] = bd.sections[5]; bd.sections[5] = t; }),
  { expectBdSec: baseBdSecCount, expectSynSec: baseSynSecCount, expectBdSub: baseBdSubCount });

// 6) COMBINED: rename + archive sub + reorder, all at once
const combined = assertNoDup('combined rename+archive+reorder', runApply((bd, m) => {
  bd.sections[2].label = 'COMBO-RENAME';
  m.archiveSubsection(bd, bd.sections[secWithSubIdx].subsections[0].id);
  const t = bd.sections[1]; bd.sections[1] = bd.sections[4]; bd.sections[4] = t;
}), { expectBdSec: baseBdSecCount, expectSynSec: baseSynSecCount, expectBdSub: baseBdSubCount });

// 7) DOUBLE APPLY (same session re-apply) must NOT grow rows or duplicate
console.log('\n[double apply idempotency]');
const r1 = combined;
const synModel2 = sm.alignSynModelToBd(clone(r1.synModel), r1.bdModel);
const r2 = sm.applyMatrixSave(clone(r1.bdRaw), clone(r1.synRaw), clone(r1.bdModel), synModel2);
const s1 = rawStats(r1.synRaw), s2 = rawStats(r2.synRaw);
const b1 = rawStats(r1.bdRaw), b2 = rawStats(r2.bdRaw);
check('BD rows stable on re-apply', b2.rows === b1.rows, `${b1.rows} -> ${b2.rows}`);
check('SYN rows stable on re-apply', s2.rows === s1.rows, `${s1.rows} -> ${s2.rows}`);
check('BD no dup label after re-apply', maxCount(labelCounts(r2.bdModel).sec) <= 1);
check('SYN no dup label after re-apply', maxCount(labelCounts(r2.synModel).sec) <= 1);

// 8) data preservation baseline (no-op rename of nonexistent => essentially identity)
const baseStats = rawStats(synRaw0);
const renameStats = rawStats(runApply((bd) => { bd.sections[2].label = 'RENAMED-SECTION-X'; }).synRaw);
check('SYN keeps ~all rows (only skip rows 23/24 removed)', renameStats.rows >= baseStats.rows - 2, `${baseStats.rows} -> ${renameStats.rows}`);

// 9) ADD a section — exactly ONE new section, no duplication, idempotent re-apply
console.log('\n[add section]');
const added = runApply((bd) => {
  bd.sections.push({ id: 'sec-added', label: 'ADDED-SECTION-Z', headerRow: null, endRow: null, color: '#ffff00',
    subsections: [{ id: 'sub-added', label: '_ADDED_SUB_Z', color: '#00b0f0', isNew: true, startRow: null, endRow: null }],
    customLines: [], isNew: true });
});
check('BD has exactly +1 section', added.bdModel.sections.length === baseBdSecCount + 1, `got ${added.bdModel.sections.length}`);
check('BD added section appears once', labelCounts(added.bdModel).sec.get('ADDED-SECTION-Z') === 1);
check('SYN added section appears once', (labelCounts(added.synModel).sec.get('ADDED-SECTION-Z') || 0) === 1);
check('BD no band overlap after add', bandOverlaps(added.bdModel) === 0);
check('SYN no band overlap after add', bandOverlaps(added.synModel) === 0);
const addReSyn = sm.alignSynModelToBd(clone(added.synModel), added.bdModel);
const addR2 = sm.applyMatrixSave(clone(added.bdRaw), clone(added.synRaw), clone(added.bdModel), addReSyn);
check('BD add re-apply no growth', rawStats(addR2.bdRaw).rows === rawStats(added.bdRaw).rows, `${rawStats(added.bdRaw).rows} -> ${rawStats(addR2.bdRaw).rows}`);
check('SYN add re-apply no growth', rawStats(addR2.synRaw).rows === rawStats(added.synRaw).rows, `${rawStats(added.synRaw).rows} -> ${rawStats(addR2.synRaw).rows}`);
check('BD add section still single after re-apply', labelCounts(addR2.bdModel).sec.get('ADDED-SECTION-Z') === 1);

// 10) RENAME the special CA chapter "-ADAPTATION" — must update in ALL THREE
//     views (bookmark matrix + database grid + synthesis), never duplicate.
console.log('\n[rename -ADAPTATION (CA chapter, special row 5)]');
const NEW = '-ADAPTATION RENAMED';
const adaptRes = runApply((bd) => {
  const sec = bd.sections.find((s) => s.label === '-ADAPTATION');
  if (!sec) throw new Error('no -ADAPTATION section in base model');
  sec.label = NEW;
});
// (a) synthesis model carries the new label
check('SYN adaptation renamed', (labelCounts(adaptRes.synModel).sec.get(NEW) || 0) === 1);
// (b) re-transform the persisted raw and re-extract the matrix (bookmark view)
const bdSheet2 = transformBdSheet(adaptRes.bdRaw);
const reState = sm.buildMatrixState(bdSheet2, transformSynthesisSheet(adaptRes.synRaw));
const reAdapt = reState.bd.sections.find((s) => s.headerRow === 5);
check('MATRIX adaptation renamed after refresh', reAdapt && reAdapt.label === NEW, reAdapt && reAdapt.label);
check('MATRIX no old -ADAPTATION left', (labelCounts(reState.bd).sec.get('-ADAPTATION') || 0) === 0);
check('MATRIX no duplicate adaptation', (labelCounts(reState.bd).sec.get(NEW) || 0) === 1);
// (c) database grid label (getRowLabel) reflects the rename on its single anchor cell
const map2 = bdStore.buildCellMap(bdSheet2.cells, bdSheet2.headerRows);
check('DATABASE adaptation renamed', bdStore.getRowLabel(map2, 5, bdSheet2.sectionHeaderRows) === NEW);
// (d) no duplicate adaptation cell in the raw grid (the original bug)
const adaptCells = adaptRes.bdRaw.cells.filter((c) => c.r === 5 && c.v === NEW);
check('DATABASE single adaptation cell (no dup)', adaptCells.length === 1, `got ${adaptCells.length}`);
check('BD no band overlap after CA rename', bandOverlaps(adaptRes.bdModel) === 0);

// 11) Rename CA "-ADAPTATION" to "ADAPTATION TEST" — synthesis band must stay intact
console.log('\n[rename -ADAPTATION → ADAPTATION TEST (syn band preserved)]');
const ADAPT_TEST = 'ADAPTATION TEST';
const adaptTestRes = runApply((bd) => {
  const sec = bd.sections.find((s) => s.label === '-ADAPTATION');
  if (!sec) throw new Error('no -ADAPTATION section in base model');
  sec.label = ADAPT_TEST;
});
const synMap = bdStore.buildCellMap(adaptTestRes.synRaw.cells, adaptTestRes.synRaw.headerRows);
check('SYN row 25 renamed', synLabel(synMap, 25) === ADAPT_TEST, synLabel(synMap, 25));
check('SYN row 26 _ADDBLUE kept', String(synLabel(synMap, 26) || '').includes('ADDBLUE'), synLabel(synMap, 26));
check(
  'SYN row 41 CABIN band kept',
  /CABIN|ADTH/i.test(String(synLabel(synMap, 41) || '')),
  synLabel(synMap, 41)
);
check(
  'SYN row 46 FENDERS not shifted to 41',
  /FENDER/i.test(String(synLabel(synMap, 46) || '')),
  synLabel(synMap, 41)
);
check('SYN adaptation section count', (labelCounts(adaptTestRes.synModel).sec.get(ADAPT_TEST) || 0) === 1);
check('SYN no orphan -ADAPTATION label', (labelCounts(adaptTestRes.synModel).sec.get('-ADAPTATION') || 0) === 0);

console.log(`\n${failed === 0 ? 'ALL NO-DUP CHECKS PASSED' : failed + ' CHECK(S) FAILED'}`);
process.exit(failed === 0 ? 0 : 1);
