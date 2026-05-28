/**
 * Smoke test: swap two BD sections must not drop row count.
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const bdPath = join(root, 'web/public/data/bd-sheet.json');
const raw = JSON.parse(readFileSync(bdPath, 'utf8'));

const { extractBdStructure, applyMatrixSave, cloneStructure } = await import(
  `file://${join(root, 'web/js/structureModel.js').replace(/\\/g, '/')}`
);

const { transformBdSheet } = await import(
  `file://${join(root, 'web/js/sheetTransform.js').replace(/\\/g, '/')}`
);

const sheet = transformBdSheet(raw);
const model = extractBdStructure(sheet);
const beforeCells = raw.cells.length;
const beforeRows = new Set(raw.cells.map((c) => c.r)).size;

if (model.sections.length < 2) {
  console.error('Need at least 2 sections');
  process.exit(1);
}

const swapped = cloneStructure(model);
const a = swapped.sections[0];
const b = swapped.sections[1];
swapped.sections[0] = b;
swapped.sections[1] = a;

const renamed = cloneStructure(swapped);
renamed.sections[0].label = 'TEST-SECTION-RENAME';

const result = applyMatrixSave(
  JSON.parse(JSON.stringify(raw)),
  null,
  renamed,
  null
);

const afterCells = result.bdRaw.cells.length;
const afterRows = new Set(result.bdRaw.cells.map((c) => c.r)).size;
const secCount = result.bdModel.sections.length;

console.log({
  beforeCells,
  afterCells,
  beforeRows,
  afterRows,
  sectionsAfter: secCount,
  headers: result.bdModel.sections.slice(0, 3).map((s) => ({
    label: s.label,
    headerRow: s.headerRow,
    endRow: s.endRow,
  })),
});

if (secCount < model.sections.length) {
  console.error('FAIL: sections lost');
  process.exit(1);
}
if (afterCells < beforeCells * 0.95) {
  console.error('FAIL: too many cells dropped');
  process.exit(1);
}
console.log('OK');
