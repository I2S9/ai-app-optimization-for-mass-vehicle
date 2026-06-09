/**
 * Options SP2 row 14 (Curb mass) ↔ Synthesis row 16.
 * SP2 F…T = Synthesis display M…AA (Excel R…AF).
 * SP2 V…AG = Synthesis display AC…AN (Excel AH…AS).
 */
import {
  formatSynNumericDisplay,
  isSynEfEqTableCol,
  synRowMaaPresetRaw,
  synRowAcanPresetRaw,
  synRowApbbPresetRaw,
  synRowBdboPresetRaw,
  synRowBscePresetRaw,
  synRowEfeqPresetRaw,
} from './synStore.js';
import {
  synMaaExcelCols,
  synAcanExcelCols,
} from './synthesisCalc.js?v=grid-perf2';

export const OPTIONS_SP2_SYN_COL_OFFSET = 12;

/** Options SP2 row 14 ↔ Synthesis display M…AA. */
export const SP2_ROW14_MAA_COLS = [
  'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
];

/** Options SP2 row 14 ↔ Synthesis display AC…AN. */
export const SP2_ROW14_ACAN_COLS = [
  'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',
];

const SP2_ROW14_CURB_COL_SET = new Set([
  ...SP2_ROW14_MAA_COLS,
  ...SP2_ROW14_ACAN_COLS,
]);

function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function numToCol(n) {
  let s = '';
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

/** Excel columns for Synthesis row 16 that feed Options SP2 row 14 (M…AA + AC…AN). */
export function synRow16CurbSourceExcelCols() {
  return [...synMaaExcelCols(), ...synAcanExcelCols()];
}

/** Map SP2 row-14 column → matching Synthesis Excel column on row 16. */
export function sp2ColToSynExcelCol(sp2Col) {
  const iMaa = SP2_ROW14_MAA_COLS.indexOf(sp2Col);
  if (iMaa >= 0) {
    const maa = synMaaExcelCols();
    return maa[iMaa] || null;
  }
  const iAcan = SP2_ROW14_ACAN_COLS.indexOf(sp2Col);
  if (iAcan >= 0) {
    const acan = synAcanExcelCols();
    return acan[iAcan] || null;
  }
  return null;
}

export function isSp2CurbLinkedCol(sp2Col) {
  return SP2_ROW14_CURB_COL_SET.has(sp2Col);
}

export function isSp2Row14MaaCurbCol(sp2Col) {
  return SP2_ROW14_MAA_COLS.includes(sp2Col);
}

export function formatSp2CurbMassDisplay(raw) {
  if (raw == null) return '';
  const s = String(raw).trim();
  if (s === '') return '';
  if (/\s*kg\s*$/i.test(s)) return s;
  const formatted = formatSynNumericDisplay(s);
  return formatted ? `${formatted} kg` : '';
}

function row16PresetRaw(excelCol) {
  if (isSynEfEqTableCol(excelCol)) {
    const efeq = synRowEfeqPresetRaw(16, excelCol);
    if (efeq !== undefined) return efeq;
  }
  const chain = [
    synRowMaaPresetRaw,
    synRowAcanPresetRaw,
    synRowApbbPresetRaw,
    synRowBdboPresetRaw,
    synRowBscePresetRaw,
  ];
  for (const fn of chain) {
    const v = fn(16, excelCol);
    if (v !== undefined && v != null && String(v).trim() !== '') return v;
  }
  return undefined;
}

function rawCellValue(cells, row, col) {
  const hit = cells.find(
    (c) => Number(c.r) === Number(row) && String(c.c) === String(col)
  );
  return hit && hit.v != null && String(hit.v).trim() !== '' ? String(hit.v) : '';
}

/**
 * Best-effort row 16 display from synRaw when the Synthesis grid is not mounted.
 * Live values (Σ green totals) come from SynthesisGrid via synthesisCellLink.
 */
export function resolveSynRow16DisplayFromRaw(cells, excelCol) {
  if (!cells || !Array.isArray(cells) || !excelCol) return '';

  const direct = rawCellValue(cells, 16, excelCol);
  if (direct) {
    if (/\s*kg\s*$/i.test(direct)) return direct;
    return formatSp2CurbMassDisplay(direct);
  }

  const preset = row16PresetRaw(excelCol);
  if (preset !== undefined) return formatSp2CurbMassDisplay(String(preset));

  return '';
}
