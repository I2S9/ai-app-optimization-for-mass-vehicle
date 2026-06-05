/**
 * CDC grid — a lightweight editable spreadsheet shared by the CDC ▸ MNS and
 * CDC ▸ Options SP2 pages.
 *
 * Mirrors the Database / Synthesis convention: lettered column headers
 * (A, B, C, …) across the top and numbered row headers (1, 2, 3, …) down the
 * side. Column A is intentionally left blank (filled in later); column B is
 * seeded with the field names, one per row. Each page passes its own
 * `storageKey` so the two grids keep independent edits.
 */
import { ref, computed, inject, onMounted, watch } from 'vue';
import { formatSynNumericDisplay } from './synStore.js';

const MNS_COLUMN_COUNT = 26;
const MNS_MIN_ROWS = 60;

/**
 * Column B is intentionally left blank on the MNS page (the field labels now
 * live in column E, right-aligned, over the peach paint block).
 */
const MNS_FIELDS_COL_B = [];

/**
 * Field labels seeded into column E (rows 1–11), right-aligned over the peach
 * paint block. One label per row starting at row 1.
 */
const MNS_FIELDS_COL_E = [
  'Design Plate',
  'Trim',
  'Seats',
  'Hybridization',
  'Technical Package',
  'Technical Specification',
  'Payload',
  'Wheels',
  'Pole',
  'Engine',
  'Transmission',
];

/** Peach paint block on the MNS page: columns A–E, rows 1–11, no inner gridlines. */
const MNS_PAINT_COLS = ['A', 'B', 'C', 'D', 'E'];
const MNS_PAINT_MAX_ROW = 11;

/**
 * Red paint bands on the MNS page (columns A–E, no inner gridlines). Rows 12–14
 * are intentionally left blank between the peach block and these red bands.
 * Each band carries a centred white label (English).
 */
const MNS_RED_ROWS = [15, 16];
const MNS_RED_COLS = ['A', 'B', 'C', 'D', 'E'];
const MNS_RED_LABELS = {
  '15:C': 'CURB MASS',
  '16:C': 'OPTIONAL EQUIPMENT',
};

/**
 * Dark-green paint bands on the MNS page (columns A–E, no inner gridlines).
 * Rows 23–24 head the front unsprung-masses block, row 42 the rear one.
 * Labels (English) are placed in specific columns; A labels are left-aligned and
 * may overflow into column B, the short C/D/E headers are centred.
 */
const MNS_GREEN_ROWS = [23, 24, 42];
const MNS_GREEN_COLS = ['A', 'B', 'C', 'D', 'E'];
const MNS_GREEN_LABELS = {
  '23:A': { text: 'UNSPRUNG MASSES DETAIL (MNS)', align: 'left' },
  '24:A': { text: 'FRONT UNSPRUNG MASSES', align: 'left' },
  '24:C': { text: 'Coeff', align: 'center' },
  '42:A': { text: 'REAR UNSPRUNG MASSES', align: 'left' },
  '42:C': { text: 'Coeff', align: 'center' },
  '42:D': { text: 'TD', align: 'center' },
  '42:E': { text: 'MB', align: 'center' },
};

/**
 * Front unsprung-masses table (rows 25–38). Columns: A = family code,
 * B = description, C = code (under the "Coeff" header), D = coefficient (yellow).
 */
const MNS_FRONT_START_ROW = 25;
const MNS_FRONT_ROWS = [
  ['FRN', 'TRANSMISSION AV', 'C81', '0.70'],
  ['FRN', 'TRANSMISSION A JOINTS AV', 'C71', '0.70'],
  ['FRN', 'ELEMENTS COMMUNS 1C7', 'C7Z', '0.70'],
  ['FRN', 'MOYEU - ROULEMENT ROUE AV', 'E13', '1.00'],
  ['FRN', 'ROUE PRINCIPALE', 'E31', '0.50'],
  ['FRN', 'ENJOLIVEUR ROUE', 'E33', '0.50'],
  ['FRN', 'INDICATEUR PERTE PRESSION', 'E34', '0.50'],
  ['FRN', 'DISQUE - TAMBOUR AV', 'F11', '1.00'],
  ['FRN', 'ETRIER - PLATEAU AV (SANS CABLE)', 'F21', '1.00'],
  ['HYD', 'ELEMENT SUSPENSION AV HYDROPNEUM', 'E16', '0.70'],
  ['LAS', 'BERCEAU - TRAVERSE TRAIN AV', 'E11', '0.00'],
  ['LAS', 'TRIANGLE BRAS ESSIEU RIGIDE', 'E12', '0.40'],
  ['LAS', 'ELEMENT SUSPENSION AV METALLIQUE', 'E14', '0.80'],
  ['LAS', 'ANTIDEVERS AV', 'E15', '0.20'],
];

/**
 * Rear unsprung-masses table (rows 43–57). Columns: A = family code,
 * B = description, C = code ("Coeff"), D = TD, E = MB (D & E beige).
 */
const MNS_REAR_START_ROW = 43;
const MNS_REAR_ROWS = [
  ['FRN', 'TRANSMISSION AR ROUES INDEPENDANT', 'C62', '0.70', '0.70'],
  ['FRN', 'TRANSMISSION AR TRAIN RIGIDE', 'C63', '1.00', '1.00'],
  ['FRN', 'TRANSMISSION A JOINTS AR', 'C73', '0.70', '0.70'],
  ['FRN', 'MOYEU - ROULEMENT ROUE AR', 'E24', '1.00', '1.00'],
  ['FRN', 'ROUE PRINCIPALE', 'E31', '0.50', '0.50'],
  ['FRN', 'ENJOLIVEUR ROUE', 'E33', '0.50', '0.50'],
  ['FRN', 'INDICATEUR PERTE PRESSION', 'E34', '0.50', '0.50'],
  ['FRN', 'DISQUE - TAMBOUR AR', 'F31', '1.00', '1.00'],
  ['FRN', 'ETRIER - PLATEAU AR (SANS CABLE)', 'F41', '1.00', '1.00'],
  ['HYD', 'ELEMENT SUSPENSION AR HYDROPNEUM', 'E28', '0.40', '0.40'],
  ['LAS', 'BERCEAU - TRAVERSE TRAIN AR', 'E21', '0.70', '0.70'],
  ['LAS', 'BRAS SUSPENSION AR - ENS. BRAS - FUSEE', 'E22', '1.00', '1.00'],
  ['LAS', 'ESSIEU RIGIDE - GUIDAGE', 'E23', '0.50', '0.50'],
  ['LAS', 'ELEMENT SUSPENSION AR METALLIQUE', 'E25', '1.00', '0.75'],
  ['LAS', 'ANTIDEVERS AR', 'E26', '0.20', '0.20'],
];

/** Coloured table bodies. Yellow front coeff column, beige rear TD/MB columns. */
const MNS_FRONT_YELLOW_COL = 'D';
const MNS_FRONT_YELLOW_RANGE = [25, 40];
const MNS_REAR_BEIGE_COLS = ['D', 'E'];
const MNS_REAR_BEIGE_RANGE = [43, 59];

/** Values displayed in red on the rear table (matches the source). */
const MNS_REAR_RED_VALUES = new Set(['53:D', '56:E']);

/**
 * Column B labels for the Options SP2 page: adds BATTERY before POLE and an
 * empty row between TRIM and CURB MASS. Empty strings leave that row blank.
 */
const OPTIONS_SP2_FIELDS_COL_B = [
  'PLATEFORM',
  'PROJECT',
  'SILHOUETTE',
  'HYBRIDIZATION',
  'DESIGN PLATE',
  'SEATS',
  'TECHNICAL SPECIFICATIONS',
  'BATTERY',
  'POLE',
  'ENERGY',
  'TECHNICAL PACK',
  'TRIM',
  '',
  'CURB MASS',
  '', // row 15 — white block extends to here
  '', // row 16 — red band
  '', // row 17 — red band
  '', // row 18 — red band
  // Options list (from image), one per row starting at row 19
  'OBC/IDCM 11kw',
  'heat pump',
  '2 position trunk floor',
  'Anti-lift, Volumetric, Perimeter Alarm',
  'DàD',
  'Wireless charger 15W',
  'Heated windscreen',
  'Head-Up Display',
  'glass roof',
  'sun roof pano (J4U)',
  'Heated steering wheel',
  'heated front seat',
  'heated rear seat',
  'Powered trunk opening',
  'Clean cabin',
  'BTA 3S',
  'Branded HIFI',
  'Roof Bar',
  'Roof Bar (dacia)',
  'attelage',
  'Spare wheel',
  'KDPP',
  'Tire 235/50R19 alloy',
  'Rear & front parking + Flankguard + WPP + VP1',
  "'REAR + FRONT  PARK + WPP + VP3 + FULL PARK ASSIST (CP",
  'Garnissage mi-TEP/Alcantara ald mi-TEP',
  'Siège cond électrique',
  'Peinture biton',
  'Projecteurs MATRIX LED',
  'MECHANICAL, REMOTE, APPROACHING HANDS FREE',
  'ecran sous caisse (prise en compte export)',
];

/**
 * Options SP2 numeric data for columns C, D, E, one row per option starting at
 * row 19 (aligned with OPTIONS_SP2_FIELDS_COL_B). Stored as display strings so
 * the one-decimal formatting from the source is preserved.
 */
const OPTIONS_SP2_CDE_START_ROW = 19;
const OPTIONS_SP2_CDE = [
  ['5.0', '4.0', '1.0'],
  ['7.6', '6.6', '1.0'],
  ['2.0', '0.9', '1.1'],
  ['0.3', '0.3', '0.0'],
  ['5.0', '4.0', '1.0'],
  ['0.8', '0.7', '0.1'],
  ['0.3', '0.2', '0.0'],
  ['1.8', '1.3', '0.5'],
  ['13.0', '6.0', '7.0'],
  ['24.0', '8.0', '16.0'],
  ['0.5', '0.4', '0.1'],
  ['0.5', '0.3', '0.2'],
  ['0.4', '0.2', '0.2'],
  ['3.0', '-0.6', '3.6'],
  ['0.1', '0.1', '0.0'],
  ['0.2', '0.1', '0.1'],
  ['7.1', '1.7', '5.4'],
  ['2.7', '1.0', '1.7'],
  ['6.3', '3.0', '3.3'],
  ['25.0', '-6.0', '31.0'],
  ['14.7', '-6.0', '20.7'],
  ['-1.2', '-0.2', '-1.0'],
  ['4.0', '2.0', '2.0'],
  ['1.0', '0.8', '0.2'],
  ['1.8', '1.4', '0.4'],
  ['1.8', '0.8', '1.0'],
  ['2.3', '1.0', '1.3'],
  ['0.5', '0.2', '0.3'],
  ['2.9', '3.5', '-0.6'],
  ['0.6', '0.3', '0.3'],
  ['8.5', '8.0', '0.5'],
];

/**
 * Synthesis header table copied into Options SP2 columns F–T, rows 1–12.
 * The Synthesis page hides Excel columns A–E, so the *displayed* columns M–AA
 * correspond to Excel columns R–AF — that is the actual source window here.
 * Each inner array holds the 15 values for columns F…T of one row (Excel 3–14).
 */
const OPTIONS_SP2_FT_COLS = ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T'];
/** Synthesis source columns (Excel R–AF = displayed M–AA) feeding Options SP2 F–T. */
const OPTIONS_SP2_FT_SRC_COLS = ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'];
/** Options SP2 row 14 (Curb mass) is mirrored live from Synthesis row 16. */
const OPTIONS_SP2_CURB_ROW = 14;
const SYN_CURB_MASS_ROW = 16;
const OPTIONS_SP2_FT = [
  ['STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S'],
  ['SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2'],
  ['O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H', 'O3H'],
  ['BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'MHEVP2', 'MHEVP2', 'MHEVP2', 'HEV', 'HEV', 'HEV'],
  ['EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA'],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'AWD', 'AWD', 'AWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD'],
  ['HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'XR', 'XR', 'XR', 'TT', 'TT', 'TT', 'TT', 'TT', 'TT'],
  ['HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'X_Range', 'X_Range', 'X_Range', 'X_Range', 'X_Range', 'X_Range', '', '', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET'],
  ['S', 'M', 'L', 'S', 'M', 'L', 'M', 'L', 'GSE', 'S', 'M', 'L', 'S', 'M', 'L'],
];

const MNS_DEFAULT_STORAGE_KEY = 'mns-grid-cells-v3';

/** Rows that make up the red bands on the Options SP2 page (columns A–E). */
const SP2_RED_ROWS = new Set([16, 17, 18, 53, 54, 55, 57, 58, 59]);
const SP2_RED_COLS = ['A', 'B', 'C', 'D', 'E'];

/** Fixed labels keyed by "row:col". Red bands + the "Montée en gamme" header. */
const SP2_STATIC_LABELS = {
  // Options band (16–18)
  '17:A': { text: 'Options', cls: 'sp2-red-title' },
  '16:C': { text: 'Optional Equipment', cls: 'sp2-red-text' },
  '17:C': { text: 'FRONT', cls: 'sp2-red-text' },
  '18:C': { text: 'BACK', cls: 'sp2-red-text' },
  // Standalone red header straddling D–E
  '52:D': { text: 'Montée en gamme', cls: 'sp2-upgrade-text' },
  // Equipment band (53–55)
  '54:A': { text: 'Equipment', cls: 'sp2-red-title' },
  '53:C': { text: 'Optional Equipment', cls: 'sp2-red-text' },
  '54:C': { text: 'FRONT', cls: 'sp2-red-text' },
  '55:C': { text: 'BACK', cls: 'sp2-red-text' },
  // Options EXPORT band (57–59)
  '58:A': { text: 'Options EXPORT', cls: 'sp2-red-title' },
  '57:C': { text: 'Optional Equipment', cls: 'sp2-red-text' },
  '58:C': { text: 'FRONT', cls: 'sp2-red-text' },
  '59:C': { text: 'BACK', cls: 'sp2-red-text' },
};

function fieldsForKey(storageKey) {
  return String(storageKey || '').includes('options-sp2')
    ? OPTIONS_SP2_FIELDS_COL_B
    : MNS_FIELDS_COL_B;
}

/** Excel-style column letter for a 1-based column index (1 → A, 27 → AA). */
function numToCol(n) {
  let s = '';
  let x = n;
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

function buildSeedCells(storageKey) {
  const cells = {};
  fieldsForKey(storageKey).forEach((label, idx) => {
    if (label) cells[`${idx + 1}:B`] = label;
  });
  if (String(storageKey || '').includes('options-sp2')) {
    OPTIONS_SP2_CDE.forEach(([c, d, e], i) => {
      const row = OPTIONS_SP2_CDE_START_ROW + i;
      if (c !== '') cells[`${row}:C`] = c;
      if (d !== '') cells[`${row}:D`] = d;
      if (e !== '') cells[`${row}:E`] = e;
    });
    // Synthesis header table → columns F–T, rows 1–12.
    OPTIONS_SP2_FT.forEach((rowVals, i) => {
      const row = i + 1;
      rowVals.forEach((v, j) => {
        if (v !== '') cells[`${row}:${OPTIONS_SP2_FT_COLS[j]}`] = v;
      });
    });
  } else {
    // MNS page: seed the right-aligned field labels into column E (rows 1–11).
    MNS_FIELDS_COL_E.forEach((label, idx) => {
      if (label) cells[`${idx + 1}:E`] = label;
    });
    // Front unsprung-masses table (A=family, B=desc, C=code, D=coeff).
    MNS_FRONT_ROWS.forEach(([a, b, c, d], i) => {
      const row = MNS_FRONT_START_ROW + i;
      cells[`${row}:A`] = a;
      cells[`${row}:B`] = b;
      cells[`${row}:C`] = c;
      cells[`${row}:D`] = d;
    });
    // Rear unsprung-masses table (A=family, B=desc, C=code, D=TD, E=MB).
    MNS_REAR_ROWS.forEach(([a, b, c, d, e], i) => {
      const row = MNS_REAR_START_ROW + i;
      cells[`${row}:A`] = a;
      cells[`${row}:B`] = b;
      cells[`${row}:C`] = c;
      cells[`${row}:D`] = d;
      cells[`${row}:E`] = e;
    });
  }
  return cells;
}

function loadStoredCells(storageKey) {
  if (typeof localStorage === 'undefined') return null;
  try {
    const raw = localStorage.getItem(storageKey);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch {
    return null;
  }
}

export default {
  name: 'MnsGrid',
  props: {
    storageKey: { type: String, default: MNS_DEFAULT_STORAGE_KEY },
  },
  setup(props) {
    const cells = ref(buildSeedCells(props.storageKey));

    // The Options SP2 page reuses this grid but brands column A (rows 1–12)
    // with a white "STLA S / SP2" header block instead of editable cells.
    const isOptionsSp2 = computed(() =>
      String(props.storageKey || '').includes('options-sp2')
    );

    // Options SP2 needs columns out to T (20) for the copied synthesis table.
    const columns = computed(() => {
      const count = isOptionsSp2.value ? 20 : MNS_COLUMN_COUNT;
      return Array.from({ length: count }, (_, i) => numToCol(i + 1));
    });

    // Live link to the Synthesis sheet (provided by the app root). Used to mirror
    // the Curb mass row (Synthesis 16 → Options SP2 row 14) in real time.
    const synLink = inject('synthesisCellLink', null);

    // Map of Options SP2 column (F–T) → formatted Curb mass value, recomputed
    // whenever the synthesis revision bumps (edit / recalc / Supabase reload).
    const curbMassRow = computed(() => {
      try {
      if (!isOptionsSp2.value || !synLink || !synLink.synRaw) return {};
      // Touch both signals so this recomputes on any synthesis change. synRaw is a
      // shallowRef mutated in place on edits, so the edit tick is what actually fires.
      if (synLink.synRevision) void synLink.synRevision.value;
      if (synLink.synEditTick) void synLink.synEditTick.value;
      const raw = synLink.synRaw.value;
        if (!raw || !Array.isArray(raw.cells)) return {};
        // One pass to gather the Curb mass row (Synthesis row 16) source columns.
        const srcVals = {};
        const wanted = new Set(OPTIONS_SP2_FT_SRC_COLS);
        for (const cell of raw.cells) {
          if (Number(cell.r) === SYN_CURB_MASS_ROW && wanted.has(cell.c)) {
            srcVals[cell.c] = cell.v;
          }
        }
        const out = {};
        OPTIONS_SP2_FT_COLS.forEach((destCol, i) => {
          const v = srcVals[OPTIONS_SP2_FT_SRC_COLS[i]];
          if (v === undefined || v === null || String(v).trim() === '') return;
          const formatted = formatSynNumericDisplay(String(v));
          out[destCol] = formatted ? `${formatted} kg` : '';
        });
        return out;
      } catch (e) {
        console.warn('Options SP2 curb-mass link:', e);
        return {};
      }
    });

    function isCurbLinkedCell(row, col) {
      return (
        isOptionsSp2.value &&
        row === OPTIONS_SP2_CURB_ROW &&
        OPTIONS_SP2_FT_COLS.includes(col)
      );
    }

    // Any cell of the copied Synthesis table block (F–T, rows 1–12) — used for the
    // bold weight that matches the Synthesis sheet.
    function isSynTableCell(row, col) {
      return (
        isOptionsSp2.value &&
        row >= 1 &&
        row <= 12 &&
        OPTIONS_SP2_FT_COLS.includes(col)
      );
    }

    // Colours for the copied Synthesis table (Options F–T, rows 1–12), matching
    // the Synthesis display rules: energy row (BEV/MHEVP2/HEV), drivetrain AWD,
    // TARGET band and the grey TT / empty-range columns.
    function synTableClass(row, col) {
      if (!isOptionsSp2.value || row < 1 || row > 12) return '';
      const i = OPTIONS_SP2_FT_COLS.indexOf(col);
      if (i < 0) return '';
      if (row === 4) {
        if (i <= 8) return 'sp2-syn-bev';
        if (i <= 11) return 'sp2-syn-mhevp2';
        return 'sp2-syn-hev';
      }
      if (row === 7 && i >= 6 && i <= 8) return 'sp2-syn-awd';
      if ((row === 8 || row === 9) && i >= 9) return 'sp2-syn-grey';
      if (row === 11) return 'sp2-syn-target';
      return '';
    }

    function curbLinkedValue(col) {
      const map = curbMassRow.value;
      return (map && map[col]) || '';
    }

    function isBrandCell(row, col) {
      return isOptionsSp2.value && col === 'A' && row >= 1 && row <= 15;
    }

    // MNS page: peach paint block over columns A–E, rows 1–11. Seamless fill
    // (no inner gridlines) for the "painted" look.
    function isMnsPaintCell(row, col) {
      return (
        !isOptionsSp2.value &&
        row >= 1 &&
        row <= MNS_PAINT_MAX_ROW &&
        MNS_PAINT_COLS.includes(col)
      );
    }

    // MNS page: column E field labels (rows 1–11), right-aligned.
    function isMnsLabelCell(row, col) {
      return !isOptionsSp2.value && col === 'E' && row >= 1 && row <= MNS_PAINT_MAX_ROW;
    }

    // MNS page: red paint bands (rows 15–16, columns A–E), no inner gridlines.
    function isMnsRedCell(row, col) {
      return (
        !isOptionsSp2.value &&
        MNS_RED_ROWS.includes(row) &&
        MNS_RED_COLS.includes(col)
      );
    }

    // Centred white label sitting on a red band (English).
    function mnsRedLabel(row, col) {
      if (isOptionsSp2.value) return null;
      return MNS_RED_LABELS[`${row}:${col}`] || null;
    }

    // MNS page: dark-green paint bands (rows 23, 24, 42, columns A–E).
    function isMnsGreenCell(row, col) {
      return (
        !isOptionsSp2.value &&
        MNS_GREEN_ROWS.includes(row) &&
        MNS_GREEN_COLS.includes(col)
      );
    }

    // White label sitting on a green band (English): { text, align }.
    function mnsGreenLabel(row, col) {
      if (isOptionsSp2.value) return null;
      return MNS_GREEN_LABELS[`${row}:${col}`] || null;
    }

    // Front coeff column (D) painted yellow over the front table body.
    function isMnsYellowCell(row, col) {
      return (
        !isOptionsSp2.value &&
        col === MNS_FRONT_YELLOW_COL &&
        row >= MNS_FRONT_YELLOW_RANGE[0] &&
        row <= MNS_FRONT_YELLOW_RANGE[1]
      );
    }

    // Rear TD/MB columns (D, E) painted beige over the rear table body.
    function isMnsBeigeCell(row, col) {
      return (
        !isOptionsSp2.value &&
        MNS_REAR_BEIGE_COLS.includes(col) &&
        row >= MNS_REAR_BEIGE_RANGE[0] &&
        row <= MNS_REAR_BEIGE_RANGE[1]
      );
    }

    // Rear values shown in red (matches the source table).
    function isMnsRedValue(row, col) {
      return !isOptionsSp2.value && MNS_REAR_RED_VALUES.has(`${row}:${col}`);
    }

    // Options SP2 red bands (rows in SP2_RED_ROWS, columns A–E), no inner gridlines.
    function isRedCell(row, col) {
      return isOptionsSp2.value && SP2_RED_ROWS.has(row) && SP2_RED_COLS.includes(col);
    }
    function isUpgradeCell(row, col) {
      return isOptionsSp2.value && row === 52 && col === 'D';
    }

    // Highlighted data cells on the Options SP2 page.
    const SP2_GREEN_C_ROWS = [22, 28, 32, 33, 40, 41, 42, 43];
    function isGreenCell(row, col) {
      return isOptionsSp2.value && col === 'C' && SP2_GREEN_C_ROWS.includes(row);
    }
    function isBlueCell(row, col) {
      return isOptionsSp2.value && col === 'B' && row === 49;
    }

    // Fixed labels (red bands + the standalone "Montée en gamme" header).
    function redLabel(row, col) {
      if (!isOptionsSp2.value) return null;
      return SP2_STATIC_LABELS[`${row}:${col}`] || null;
    }

    const rowCount = computed(() => {
      // Options SP2 needs room for the red bands / labels down to row 59.
      let maxRow = isOptionsSp2.value ? 60 : MNS_MIN_ROWS;
      for (const key of Object.keys(cells.value)) {
        const r = Number(key.split(':')[0]);
        if (Number.isFinite(r) && r > maxRow) maxRow = r;
      }
      return maxRow;
    });

    const rows = computed(() =>
      Array.from({ length: rowCount.value }, (_, i) => i + 1)
    );

    function cellKey(row, col) {
      return `${row}:${col}`;
    }

    function cellValue(row, col) {
      return cells.value[cellKey(row, col)] || '';
    }

    function persist() {
      if (typeof localStorage === 'undefined') return;
      try {
        localStorage.setItem(props.storageKey, JSON.stringify(cells.value));
      } catch {
        /* storage full / unavailable — keep working in-memory */
      }
    }

    function loadForKey(storageKey) {
      const stored = loadStoredCells(storageKey);
      // Stored edits win, but re-seed any field label the user never touched.
      cells.value = stored
        ? { ...buildSeedCells(storageKey), ...stored }
        : buildSeedCells(storageKey);
    }

    function onCellInput(row, col, event) {
      const value = event && event.target ? event.target.value : '';
      const key = cellKey(row, col);
      const next = { ...cells.value };
      if (value === '') delete next[key];
      else next[key] = value;
      cells.value = next;
      persist();
    }

    onMounted(() => {
      loadForKey(props.storageKey);
      // Ensure Synthesis data is loaded so the linked Curb mass row resolves.
      if (isOptionsSp2.value && synLink && typeof synLink.ensureSyn === 'function') {
        Promise.resolve(synLink.ensureSyn()).catch(() => {});
      }
    });

    // Reload the right dataset if the same component instance is reused for
    // another page (different storageKey).
    watch(
      () => props.storageKey,
      (key) => loadForKey(key)
    );

    return {
      columns,
      rows,
      cellValue,
      onCellInput,
      isOptionsSp2,
      isBrandCell,
      isMnsPaintCell,
      isMnsLabelCell,
      isMnsRedCell,
      mnsRedLabel,
      isMnsGreenCell,
      mnsGreenLabel,
      isMnsYellowCell,
      isMnsBeigeCell,
      isMnsRedValue,
      isRedCell,
      isUpgradeCell,
      isGreenCell,
      isBlueCell,
      redLabel,
      isCurbLinkedCell,
      curbLinkedValue,
      synTableClass,
      isSynTableCell,
    };
  },
  template: `
    <div class="bd-grid-root mns-grid" :class="{ 'options-sp2-grid': isOptionsSp2 }">
      <div class="bd-grid-scroll">
        <table class="bd-table mns-table" role="grid">
          <thead>
            <tr class="hdr-row-letters">
              <th class="corner"></th>
              <th
                v-for="col in columns"
                :key="'mns-letter-' + col"
                class="col-letter"
                :data-col="col"
              >{{ col }}</th>
            </tr>
          </thead>
          <tbody>
            <tr
              v-for="row in rows"
              :key="'mns-row-' + row"
              class="grid-row-cv"
            >
              <td class="row-num">{{ row }}</td>
              <td
                v-for="col in columns"
                :key="'mns-' + row + '-' + col"
                class="data-cell"
                :data-col="col"
                :class="[{
                  'sp2-brand-cell': isBrandCell(row, col),
                  'sp2-red-cell': isRedCell(row, col),
                  'sp2-upgrade-cell': isUpgradeCell(row, col),
                  'sp2-green-cell': isGreenCell(row, col),
                  'sp2-blue-cell': isBlueCell(row, col),
                  'mns-paint-cell': isMnsPaintCell(row, col),
                  'mns-e-label': isMnsLabelCell(row, col),
                  'mns-red-cell': isMnsRedCell(row, col),
                  'mns-green-cell': isMnsGreenCell(row, col),
                  'mns-yellow-cell': isMnsYellowCell(row, col),
                  'mns-beige-cell': isMnsBeigeCell(row, col),
                  'mns-red-value': isMnsRedValue(row, col),
                  'sp2-syn-cell': isSynTableCell(row, col),
                }, synTableClass(row, col)]"
              >
                <template v-if="isBrandCell(row, col)">
                  <span v-if="row === 6" class="sp2-brand-title">STLA S</span>
                  <span v-else-if="row === 7" class="sp2-brand-sub">SP2</span>
                </template>
                <span
                  v-else-if="redLabel(row, col)"
                  :class="redLabel(row, col).cls"
                >{{ redLabel(row, col).text }}</span>
                <span
                  v-else-if="mnsRedLabel(row, col)"
                  class="mns-red-label"
                >{{ mnsRedLabel(row, col) }}</span>
                <span
                  v-else-if="mnsGreenLabel(row, col)"
                  class="mns-green-label"
                  :class="'mns-green-' + mnsGreenLabel(row, col).align"
                >{{ mnsGreenLabel(row, col).text }}</span>
                <span
                  v-else-if="isCurbLinkedCell(row, col)"
                  class="sp2-curb-linked"
                  title="Lié à Synthesis (Curb mass) — mis à jour automatiquement"
                >{{ curbLinkedValue(col) }}</span>
                <input
                  v-else
                  type="text"
                  class="grid-cell-input mns-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  :value="cellValue(row, col)"
                  @input="onCellInput(row, col, $event)"
                />
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `,
};
