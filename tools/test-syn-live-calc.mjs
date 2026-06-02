/**
 * Live SYNTHESIS ↔ DATABASE coherence test.
 *
 * Proves two things end to end against the real project data:
 *  1. The shared classifier {@link classifyBdEditForSynMass} flags every BD column
 *     the SOMMEPROD depends on (mass V, sub-system L2 AU, gate Q and the filter
 *     columns A,B,C,D,E,F,G,I,K,L,O) and nothing else.
 *  2. Editing each kind of Database cell changes the dependent blue SOMMEPROD cell
 *     on Synthesis (via the same engine the app uses), so a Database change really
 *     propagates to Synthesis.
 *
 * Run: node tools/test-syn-live-calc.mjs
 */
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const url = (p) => `file://${join(root, p).replace(/\\/g, '/')}`;

const bdRaw = JSON.parse(readFileSync(join(root, 'web/public/data/bd-sheet.json'), 'utf8'));
const synRaw = JSON.parse(readFileSync(join(root, 'web/public/data/synthesis-sheet.json'), 'utf8'));

const { transformBdSheet, transformSynthesisSheet } = await import(url('web/js/sheetTransform.js'));
const { createSynCalcContext } = await import(url('web/js/synMaterialize.js'));
const bdStore = await import(url('web/js/bdStore.js'));
const synStore = await import(url('web/js/synStore.js'));
const calc = await import(url('web/js/synthesisCalc.js'));
const { canonicalL2MatchKey } = await import(url('web/js/bdTranslate.js'));
const { BD_MASS_COL, BD_SUBSYSTEM_L2_COL } = await import(url('web/js/bdColumnConfig.js'));

const {
  classifyBdEditForSynMass,
  bdEditAffectsSynMass,
  SYN_FILTER_BD_ROWS,
  SYN_CALC_FIRST_ROW,
  buildBdColumnIndex,
  buildSumproductL2Index,
  buildVehColFilterIndex,
  synMaaExcelCols,
  isSynBlueSubsectionRow,
} = calc;

let failed = 0;
function check(name, cond, detail) {
  if (!cond) {
    console.error(`  FAIL: ${name}${detail ? ' — ' + detail : ''}`);
    failed++;
  } else {
    console.log(`  ok: ${name}`);
  }
}
const num = (v) => {
  const n = parseFloat(String(v == null ? '' : v).replace(',', '.'));
  return Number.isFinite(n) ? n : 0;
};

// ---------------------------------------------------------------------------
// 1) Classifier covers every SOMMEPROD input column.
// ---------------------------------------------------------------------------
console.log('\n[classifier]');
const filterCols = SYN_FILTER_BD_ROWS.map(([c]) => c);
check('mass column V classified as mass', classifyBdEditForSynMass('V') === 'mass');
check('L2 column AU classified as l2', classifyBdEditForSynMass('AU') === 'l2');
check('gate column Q classified as gate', classifyBdEditForSynMass('Q') === 'gate');
for (const c of filterCols) {
  check(`filter column ${c} classified as filter`, classifyBdEditForSynMass(c) === 'filter');
}
// Columns the SOMMEPROD does NOT read must not trigger a recompute.
check('codification column R is not a SUMPRODUCT input', classifyBdEditForSynMass('R') === null);
check('title column T is not a SUMPRODUCT input', classifyBdEditForSynMass('T') === null);
check('bdEditAffectsSynMass(C) true', bdEditAffectsSynMass('C') === true);
check('bdEditAffectsSynMass(T) false', bdEditAffectsSynMass('T') === false);

// Regression guard: the old logic only watched O/P/Q/AU/V and missed A–L filters.
const oldLogic = (col) => col === BD_MASS_COL || col === 'O' || col === 'P' || col === 'Q' || col === BD_SUBSYSTEM_L2_COL;
const missedByOld = filterCols.filter((c) => !oldLogic(c) && bdEditAffectsSynMass(c));
check(
  'fix now covers filter columns the old logic ignored',
  missedByOld.length > 0 && missedByOld.every((c) => bdEditAffectsSynMass(c)),
  `recovered: ${missedByOld.join(',')}`
);

// ---------------------------------------------------------------------------
// 2) End-to-end: editing each BD column kind moves the dependent blue cell.
// ---------------------------------------------------------------------------
console.log('\n[live recompute]');

const bdSheet = transformBdSheet(bdRaw);
const synSheet = transformSynthesisSheet(synRaw);

// Independent mirror (read-only) to locate a blue Synthesis cell and the BD rows
// that feed it, before we mutate anything through the calc context.
const bdCols = buildBdColumnIndex(bdSheet);
const bdMap =
  bdSheet.cellMap instanceof Map
    ? bdSheet.cellMap
    : bdStore.buildCellMap(bdSheet.cells, bdSheet.headerRows);
const synMap =
  synSheet.cellMap instanceof Map
    ? synSheet.cellMap
    : bdStore.buildCellMap(synSheet.cells, synSheet.headerRows);
const bdL1 = bdStore.bdSubsystemL1Col(bdSheet);
const bdL2 = bdStore.bdSubsystemL2Col(bdSheet);

function getBdValue(r, c) {
  const raw = bdCols[c] ? bdCols[c][r] : undefined;
  if (raw != null && raw !== '') return String(raw);
  const displayed = bdStore.displayCellValue(
    bdMap,
    r,
    c,
    bdSheet.sectionHeaderRows,
    bdSheet.canonicalSectionByLabel,
    bdL1,
    bdL2
  );
  return displayed !== '' ? displayed : '';
}
function getSynCell(row, col) {
  if (col === 'F') {
    const label = synStore.synLabel(synMap, row);
    if (label != null && String(label).trim() !== '') return String(label).trim();
  }
  const cell = bdStore.getCell(synMap, row, col);
  return cell && cell.v != null && cell.v !== '' ? String(cell.v) : '';
}

const l2Index = buildSumproductL2Index(bdCols, getBdValue);
const lastRow =
  synSheet.effectiveLastRow != null
    ? synSheet.effectiveLastRow
    : synSheet.lastRow != null
      ? synSheet.lastRow
      : 422;
const vehCol = synMaaExcelCols()[0]; // first M…AA vehicle column (Excel R)

// Find a blue SOMMEPROD cell with a non-zero value and at least one contributing BD row.
let target = null;
const vehFilterRows = buildVehColFilterIndex(bdCols, vehCol, getSynCell, getBdValue);
const filterSet = new Set(vehFilterRows);
for (let row = SYN_CALC_FIRST_ROW; row <= lastRow; row++) {
  const label = synStore.synLabel(synMap, row);
  const labelKey = canonicalL2MatchKey(label);
  if (!labelKey) continue;
  const rowClass = synStore.synRowStyleClass(synMap, row, synSheet);
  if (!isSynBlueSubsectionRow(row, synSheet, label, rowClass)) continue;
  const l2Rows = l2Index.get(labelKey) || [];
  const contributing = l2Rows.filter((r) => filterSet.has(r) && num(getBdValue(r, BD_MASS_COL)) !== 0);
  if (!contributing.length) continue;
  target = { row, label, labelKey, contributing };
  break;
}

check('found a blue SOMMEPROD cell with contributing Database rows', target != null,
  target ? `Syn row ${target.row} "${target.label}" via col ${vehCol}` : 'none');

if (target) {
  const { row: synRow, contributing } = target;
  const bdRow = contributing[0];

  // Fresh calc context per scenario so each edit starts from the same baseline.
  function freshCtx() {
    const bd = transformBdSheet(JSON.parse(JSON.stringify(bdRaw)));
    const syn = transformSynthesisSheet(JSON.parse(JSON.stringify(synRaw)));
    return createSynCalcContext(bd, syn);
  }
  const baseVal = num(freshCtx().getDisplayValue(synRow, vehCol));
  console.log(`  baseline ${target.label} @ col ${vehCol} = ${baseVal}`);
  const origMass = num(getBdValue(bdRow, BD_MASS_COL));

  // (a) MASS edit: +1000 kg on a contributing row → cell grows by exactly 1000.
  {
    const ctx = freshCtx();
    ctx.patchBdCell(bdRow, BD_MASS_COL, String(origMass + 1000));
    ctx.recalcAfterBdEdit();
    const after = num(ctx.getDisplayValue(synRow, vehCol));
    // Tolerance 1 kg: displayed values are rounded for presentation, so the +1000
    // delta can land within ±0.1 of 1000 — the point is that the mass propagated.
    check('mass edit (V) updates blue cell by ~+1000', Math.abs(after - (baseVal + 1000)) < 1,
      `${baseVal} -> ${after}`);
  }

  // (b) FILTER edit: break the Silhouette (C) match on the contributing row → it
  //     drops out of the sum (cell decreases by that row's mass).
  {
    const ctx = freshCtx();
    ctx.patchBdCell(bdRow, 'C', '___NO_MATCH___');
    ctx.recalcAfterBdEdit();
    const after = num(ctx.getDisplayValue(synRow, vehCol));
    check('filter edit (C) removes the row from the SOMMEPROD', after < baseVal - 1e-6,
      `${baseVal} -> ${after} (removed mass ~${origMass})`);
  }

  // (c) L2 edit: move the row to a different sub-system → it no longer feeds this label.
  {
    const ctx = freshCtx();
    ctx.patchBdCell(bdRow, BD_SUBSYSTEM_L2_COL, '___OTHER_SUBSYSTEM___');
    ctx.recalcAfterBdEdit();
    const after = num(ctx.getDisplayValue(synRow, vehCol));
    check('L2 edit (AU) removes the row from the label total', after < baseVal - 1e-6,
      `${baseVal} -> ${after}`);
  }

  // (d) GATE edit: flip Q away from "S" → the row is excluded everywhere.
  {
    const ctx = freshCtx();
    ctx.patchBdCell(bdRow, 'Q', 'X');
    ctx.recalcAfterBdEdit();
    const after = num(ctx.getDisplayValue(synRow, vehCol));
    check('gate edit (Q≠S) removes the row from the SOMMEPROD', after < baseVal - 1e-6,
      `${baseVal} -> ${after}`);
  }
}

console.log(`\n${failed === 0 ? 'ALL LIVE-CALC CHECKS PASSED' : failed + ' CHECK(S) FAILED'}`);
process.exit(failed === 0 ? 0 : 1);
