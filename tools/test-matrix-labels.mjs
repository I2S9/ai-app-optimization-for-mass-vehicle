import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const url = (p) => `file://${join(root, p).replace(/\\/g, '/')}`;
const clone = (x) => JSON.parse(JSON.stringify(x));
const bdRaw0 = JSON.parse(readFileSync(join(root, 'web/public/data/bd-sheet.json'), 'utf8'));
const synRaw0 = JSON.parse(readFileSync(join(root, 'web/public/data/synthesis-sheet.json'), 'utf8'));
const { transformBdSheet, transformSynthesisSheet } = await import(url('web/js/sheetTransform.js'));
const sm = await import(url('web/js/structureModel.js'));
const { buildCellMap, cellLabelValue, getCell, getRowLabel } = await import(url('web/js/bdStore.js'));

let failed = 0;
function ok(m) {
  console.log(`  ok: ${m}`);
}
function fail(m, d) {
  failed++;
  console.error(`  FAIL: ${m}${d ? ' — ' + d : ''}`);
}

const base = sm.buildMatrixState(transformBdSheet(bdRaw0), transformSynthesisSheet(synRaw0));
const res = sm.applyMatrixSave(
  clone(bdRaw0),
  clone(synRaw0),
  (() => {
    const bd = clone(base.bd);
    const sec = bd.sections.find((s) => s.caBand && s.headerRow === 5);
    sec.label = 'ADAPTATION';
    return bd;
  })(),
  sm.alignSynModelToBd(clone(base.syn), (() => {
    const bd = clone(base.bd);
    const sec = bd.sections.find((s) => s.caBand && s.headerRow === 5);
    sec.label = 'ADAPTATION';
    return bd;
  })())
);

const bdSheet = transformBdSheet(res.bdRaw);
const map = buildCellMap(bdSheet.cells, bdSheet.headerRows);
const w5 = cellLabelValue(getCell(map, 5, 'W'));
if (w5 !== 'ADAPTATION') fail('BD row 5 W after CA rename', w5);
else ok('BD row 5 W = ADAPTATION after matrix rename');

const rowLabel = getRowLabel(map, 5, bdSheet.sectionHeaderRows);
if (rowLabel !== 'ADAPTATION') fail('Database getRowLabel row 5', rowLabel);
else ok('Database row 5 label matches matrix');

for (const s of res.bdModel.sections) {
  if (s.label.includes('IF(') || s.label.startsWith('=')) {
    fail('formula in bdModel section', s.label.slice(0, 60));
  }
}
ok('no formula text in bdModel section labels');

console.log(failed ? `\n${failed} failed` : '\nALL MATRIX LABEL CHECKS PASSED');
process.exit(failed ? 1 : 0);
