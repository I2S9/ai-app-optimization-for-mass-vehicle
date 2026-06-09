/**
 * Options SP2 row 14 (Curb mass) ↔ Synthesis row 16.
 * SP2 F…T = Synthesis display M…AA (Excel R…AF).
 * SP2 V…AG = Synthesis display AC…AN (Excel AH…AS).
 * SP2 AI…AU = Synthesis display AP…BB (Excel AU…BG).
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
  synApbbExcelCols,
} from './synthesisCalc.js?v=grid-perf2';

export const OPTIONS_SP2_SYN_COL_OFFSET = 12;

/** Options SP2 row 14 ↔ Synthesis display M…AA. */
export const SP2_ROW14_MAA_COLS = [
  'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
];

/** MNS row 14 (Y…AM) ↔ Synthesis display M…AA (same Excel source as SP2 F…T). */
export const MNS_SYN_YAM_COLS = [
  'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM',
];

function mnsColsFromTo(fromCol, toCol) {
  const out = [];
  for (let n = colToNum(fromCol); n <= colToNum(toCol); n++) out.push(numToCol(n));
  return out;
}

/**
 * MNS Synthesis copy — display band AD…AZ (AC…AN + AP…AZ).
 * Column AN = blank spacer (24px, like X); tables AO…AZ + BB…BL (BA = inner spacer).
 */
export const MNS_SYN_TABLE_SPACER_COL = 'AN';
export const MNS_SYN_ADAO_COLS = mnsColsFromTo('AO', 'AZ');
/** Blank spacer between AC…AN (AO…AZ) and AP…AZ (BB…BL) blocks — same width as X. */
export const MNS_SYN_ADAZ_SPACER_COL = 'BA';
export const MNS_SYN_APAZ_COLS = mnsColsFromTo('BB', 'BL');
export const MNS_SYN_ADAZ_COLS = [...MNS_SYN_ADAO_COLS, ...MNS_SYN_APAZ_COLS];

/** Options SP2 row 14 ↔ Synthesis display AC…AN. */
export const SP2_ROW14_ACAN_COLS = [
  'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',
];

/** Options SP2 row 14 ↔ Synthesis display AP…BB. */
export const SP2_ROW14_APBB_COLS = [
  'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU',
];

const SP2_ROW14_CURB_COL_SET = new Set([
  ...SP2_ROW14_MAA_COLS,
  ...SP2_ROW14_ACAN_COLS,
  ...SP2_ROW14_APBB_COLS,
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

/** Excel columns for Synthesis row 16 that feed Options SP2 row 14 (M…AA + AC…AN + AP…BB). */
export function synRow16CurbSourceExcelCols() {
  return [...synMaaExcelCols(), ...synAcanExcelCols(), ...synApbbExcelCols()];
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
  const iApbb = SP2_ROW14_APBB_COLS.indexOf(sp2Col);
  if (iApbb >= 0) {
    const apbb = synApbbExcelCols();
    return apbb[iApbb] || null;
  }
  return null;
}

export function isSp2CurbLinkedCol(sp2Col) {
  return SP2_ROW14_CURB_COL_SET.has(sp2Col);
}

/** Map MNS Y…AM column → matching Synthesis Excel column on row 16. */
export function mnsYamColToSynExcelCol(mnsCol) {
  const i = MNS_SYN_YAM_COLS.indexOf(mnsCol);
  if (i < 0) return null;
  const maa = synMaaExcelCols();
  return maa[i] || null;
}

export function isMnsSynRow16LinkedCol(mnsCol) {
  return MNS_SYN_YAM_COLS.includes(mnsCol);
}

/** Map MNS AO…BL column → matching Synthesis Excel column on row 16. */
export function mnsAdazColToSynExcelCol(mnsCol) {
  const iAcan = MNS_SYN_ADAO_COLS.indexOf(mnsCol);
  if (iAcan >= 0) {
    const acan = synAcanExcelCols();
    return acan[iAcan] || null;
  }
  const iApbb = MNS_SYN_APAZ_COLS.indexOf(mnsCol);
  if (iApbb >= 0) {
    const apbb = synApbbExcelCols();
    return apbb[iApbb] || null;
  }
  return null;
}

export function isMnsSynRow16LinkedColAdaz(mnsCol) {
  return MNS_SYN_ADAZ_COLS.includes(mnsCol);
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
