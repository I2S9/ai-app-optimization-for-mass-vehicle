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
import {
  ref,
  shallowRef,
  computed,
  inject,
  onMounted,
  onUnmounted,
  watch,
  watchEffect,
  nextTick,
} from 'vue';
import {
  isSp2CurbLinkedCol,
  sp2ColToSynExcelCol,
} from './sp2CurbLink.js?v=2';
import { createGridAxisHighlight } from './gridAxisHighlight.js?v=axis2';
import {
  ROW_H,
  rowOverscan,
  colOverscanPx,
  createRangeScrollCache,
  createScrollRafSync,
  shouldVirtualizeRows,
  shouldVirtualizeCols,
} from './gridScroll.js?v=scroll-perf1';

const MNS_COLUMN_COUNT = 26;
const MNS_MIN_ROWS = 60;
const OPTIONS_SP2_MIN_ROWS = 69;
/** Frozen label columns while horizontally scrolling the wide SP2 sheet. */
const SP2_STICKY_COLS = ['A', 'B'];
const SP2_ROW_NUM_W = 42;
const SP2_DEFAULT_COL_W = 72;
const SP2_MAX_RENDERED_COLS = 52;
const SP2_MAX_RENDERED_ROWS = 64;

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

function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function sp2ColsFromTo(fromCol, toCol) {
  const out = [];
  for (let n = colToNum(fromCol); n <= colToNum(toCol); n++) out.push(numToCol(n));
  return out;
}

/** Options SP2 grid width — A…HC (211 cols). */
const OPTIONS_SP2_MAX_COL = 'HC';
const OPTIONS_SP2_COLUMN_COUNT = colToNum(OPTIONS_SP2_MAX_COL);
/** Blank spacer between the M…AA block (F–T) and the AC…AN block (V–AG). */
const OPTIONS_SP2_BLANK_COL = 'U';
/** Blank spacer before the AP…BB copy (AI…AU). */
const OPTIONS_SP2_BLANK_COL_AH = 'AH';
/** Green options band — rows 19–39 & 41–47 (row 40 = white), cols V–AG + AV–HC. */
const OPTIONS_SP2_EXTENDED_GREEN_COLS = sp2ColsFromTo('AV', OPTIONS_SP2_MAX_COL);
/** Rows 44–47 pre-filled with NO on extended green cols (AV–HC). */
const OPTIONS_SP2_NO_SEED_ROWS = [44, 45, 46, 47];
const OPTIONS_SP2_NO_SEED_VALUE = 'NO';
const OPTIONS_SP2_OPTION_WHITE_ROW = 40;
const OPTIONS_SP2_FT_OPTION_START_ROW = 19;
const OPTIONS_SP2_FT_OPTION_END_ROW = 47;

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

/** Blue spot labels on row 15 (columns H–I), variant headers on the curb-mass band. */
const MNS_BLUE_LABELS = {
  '15:H': 'P64 € Bev FWD',
  '15:I': 'P64 € Bev AWD',
};

/** Curb-mass mini table — columns G–I, rows 17–21 (bold black grid). */
const MNS_HI_TABLE_ROWS = [17, 18, 19, 20, 21];
const MNS_HI_TABLE_COLS = ['G', 'H', 'I'];
/** Row labels in column G (rows 18–21); H/I stay editable under the row-15 headers. */
const MNS_HI_TABLE_LABELS = {
  '18:G': 'Front Axle',
  '19:G': 'Rear Axle',
  '20:G': 'Front Brake',
  '21:G': 'Rear Brake',
};

/** Green variant headers on row 15 (columns K–N). */
const MNS_LIME_LABELS = {
  '15:K': 'P1X FWD L1',
  '15:L': 'P1X AWD L2',
  '15:M': 'O2X FWD L1',
  '15:N': 'O2X AWD L3',
};

/** Curb-mass mini table — columns K–N, rows 18–21 (row labels shared with col G). */
const MNS_KN_TABLE_ROWS = [18, 19, 20, 21];
const MNS_KN_TABLE_COLS = ['K', 'L', 'M', 'N'];

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
  '24:G': { text: 'MNS AV', align: 'center' },
  '42:A': { text: 'REAR UNSPRUNG MASSES', align: 'left' },
  '42:C': { text: 'Coeff', align: 'center' },
  '42:D': { text: 'TD', align: 'center' },
  '42:E': { text: 'MB', align: 'center' },
  '42:G': { text: 'MNS AR', align: 'center' },
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
  // Rows 50–59 — blank
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  'Engine protection shield BEV',
  'Half spare wheel (after market)',
  '', // row 62 — bold title via SP2_STATIC_LABELS
  '',
  '',
  '',
  'Jeep 235/50 R19 AWD upgrade',
  '',
  '',
  'Weight master file Avanger',
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
/** Dropdown values on the green/white options band (rows 19–47, cols F–T + V–AG + AH–HC). */
const SP2_OPTION_CHOICES = ['O', 'NO', 'REF'];
const SP2_OPTION_CYCLE = ['', 'O', 'NO', 'REF'];
/** Equipment red band (rows 53–55): columns F–T seeded with 0.00. */
const OPTIONS_SP2_EQUIPMENT_ROWS = [53, 54, 55];
const OPTIONS_SP2_EQUIPMENT_FT_VALUE = '0.00';
/** Synthesis source columns (Excel R–AF = displayed M–AA) feeding Options SP2 F–T. */
const OPTIONS_SP2_FT_SRC_COLS = ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF'];
/** Options SP2 row 14 (Curb mass) is mirrored live from Synthesis row 16. */
const OPTIONS_SP2_CURB_ROW = 14;
/** Synthesis-style table frame (F–T): header rows 1–12 + body rows 14–18. */
const OPTIONS_SP2_SYN_TABLE_END_ROW = 18;
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

/** Numeric block columns F–T, rows 16–18. Row 15 = live sum (row 16 + row 53). */
const OPTIONS_SP2_SUM_ROW = 15;
const OPTIONS_SP2_SUM_SRC_ROWS = [16, 53];
const OPTIONS_SP2_FT_16_18 = [
  ['35.17', '64.47', '48.67', '30.17', '59.47', '43.67', '58.97', '43.67', '43.67', '22.57', '59.47', '43.67', '22.57', '59.47', '43.67'],
  ['11.38', '4.04', '-4.06', '7.38', '0.04', '-8.06', '-0.26', '-8.06', '-8.06', '0.78', '0.04', '-8.06', '0.78', '0.04', '-8.06'],
  ['23.82', '60.46', '52.77', '22.82', '59.46', '51.77', '59.27', '51.77', '51.77', '21.82', '59.46', '51.77', '21.82', '59.46', '51.77'],
];
const OPTIONS_SP2_FT_16_18_START_ROW = 16;

/**
 * Rows 16–18 (F…T + V…AG) — per column, SOMME.SI on the option band (l.19–47):
 * row 16 ← sum C, row 17 ← sum D, row 18 ← sum E, when that column = "O".
 * Row 40 (white band) is skipped.
 */
const OPTIONS_SP2_FT_SUMIF_ROWS = [
  { row: 16, sumCol: 'C' },
  { row: 17, sumCol: 'D' },
  { row: 18, sumCol: 'E' },
];
const OPTIONS_SP2_ROW16_SUMIF_ROW = 16;
const OPTIONS_SP2_SUMIF_START_ROW = OPTIONS_SP2_FT_OPTION_START_ROW;
const OPTIONS_SP2_SUMIF_END_ROW = 49;
const OPTIONS_SP2_ROW16_SUMIF_START_ROW = OPTIONS_SP2_SUMIF_START_ROW;
const OPTIONS_SP2_ROW16_SUMIF_END_ROW = OPTIONS_SP2_SUMIF_END_ROW;
const OPTIONS_SP2_ROW16_SUMIF_CRITERIA = 'O';
const OPTIONS_SP2_FT_SUMIF_SUM_COLS = new Set(['C', 'D', 'E']);

/**
 * Options picker values — columns F–T, rows 19–47 (source spreadsheet).
 * Row 40 = white band; rows 19–39 & 41–47 = green band.
 */
const OPTIONS_SP2_FT_OPTIONS = [
  ['O', 'REF', 'REF', 'O', 'REF', 'REF', 'REF', 'REF', 'REF', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['O', 'REF', 'REF', 'O', 'REF', 'REF', 'REF', 'REF', 'REF', 'NO', 'REF', 'NO', 'NO', 'REF', 'REF'],
  ['NO', 'NO', 'REF', 'NO', 'NO', 'REF', 'NO', 'REF', 'REF', 'NO', 'NO', 'REF', 'NO', 'NO', 'REF'],
  ['NO', 'NO', 'REF', 'NO', 'NO', 'REF', 'NO', 'REF', 'REF', 'NO', 'NO', 'REF', 'NO', 'NO', 'REF'],
  ['O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O'],
  ['NO', 'REF', 'REF', 'NO', 'REF', 'REF', 'REF', 'REF', 'REF', 'NO', 'NO', 'REF', 'NO', 'REF', 'REF'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['O', 'REF', 'REF', 'O', 'REF', 'REF', 'REF', 'REF', 'REF', 'O', 'REF', 'REF', 'O', 'REF', 'REF'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['O', 'O', 'REF', 'O', 'O', 'REF', 'O', 'REF', 'REF', 'O', 'O', 'REF', 'O', 'O', 'REF'],
  ['O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O'],
  ['NO', 'NO', 'REF', 'NO', 'NO', 'REF', 'NO', 'REF', 'REF', 'NO', 'NO', 'REF', 'NO', 'NO', 'REF'],
  ['NO', 'REF', 'REF', 'NO', 'REF', 'REF', 'REF', 'REF', 'REF', 'NO', 'REF', 'REF', 'NO', 'REF', 'REF'],
  ['O', 'REF', 'REF', 'O', 'REF', 'REF', 'REF', 'REF', 'REF', 'O', 'REF', 'REF', 'O', 'REF', 'REF'],
  ['O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O'],
  ['NO', 'NO', 'REF', 'NO', 'NO', 'REF', 'NO', 'REF', 'REF', 'NO', 'NO', 'REF', 'NO', 'NO', 'REF'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'O', 'O', 'NO', 'O', 'O', 'O', 'O', 'O', 'NO', 'O', 'O', 'NO', 'O', 'O'],
  ['O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O'],
  ['O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O', 'O'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'NO', 'NO', 'O', 'NO', 'NO', 'NO', 'NO', 'NO', 'O', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'O', 'REF', 'NO', 'O', 'REF', 'O', 'REF', 'REF', 'NO', 'O', 'REF', 'NO', 'O', 'REF'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
  ['NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO', 'NO'],
];

/**
 * Synthesis AC…AN header table → Options SP2 columns V–AG, rows 1–12.
 * Source: Synthesis display AC…AN (Excel AH…AS), rows 3–14.
 */
const OPTIONS_SP2_VAG_COLS = [
  'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',
];
/** Synthesis source columns (Excel AH…AS = displayed AC…AN) feeding Options SP2 V–AG. */
const OPTIONS_SP2_VAG_SRC_COLS = [
  'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS',
];
const OPTIONS_SP2_VAG = [
  ['STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S'],
  ['SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2'],
  ['O3W', 'O3W', 'O3W', 'O3W', 'O3W', 'O3W', 'O3W', 'O3W', 'O3W', 'O3W', 'O3W', 'O3W'],
  ['BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'MHEVP2', 'MHEVP2', 'MHEVP2', 'HEV', 'HEV', 'HEV'],
  ['EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA'],
  ['', '', '', '', '', '', '', '', '', '', '', ''],
  ['FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD'],
  ['HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'TT', 'TT', 'TT', 'TT', 'TT', 'TT'],
  ['HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'X_Range', 'X_Range', 'X_Range', '', '', '', '', '', ''],
  ['SW', 'SW', 'SW', 'SW', 'SW', 'SW', 'SW', 'SW', 'SW', 'SW', 'SW', 'SW'],
  ['TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET'],
  ['S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L', 'S', 'M', 'L'],
];

/** Map 15-col F…T seed row → 12-col V…AG (AC…AN vehicle block). */
function sp2PickVagFromFtRow(ftRow) {
  return ftRow.slice(3, 15);
}

/** Options picker — columns V–AG, rows 19–47 (same pattern as F–T, AC…AN vehicles). */
const OPTIONS_SP2_VAG_OPTIONS = OPTIONS_SP2_FT_OPTIONS.map(sp2PickVagFromFtRow);

/**
 * Synthesis AP…BB header table → Options SP2 columns AI…AU, rows 1–12.
 * Source: Synthesis display AP…BB (Excel AU…BG), rows 3–14.
 */
const OPTIONS_SP2_AU_COLS = sp2ColsFromTo('AI', 'AU');
const OPTIONS_SP2_AU = [
  ['STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S', 'STLA/S'],
  ['SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2', 'SP2'],
  ['P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S', 'P3S'],
  ['BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'BEV', 'MHEVP2', 'MHEVP2', 'HEV', 'HEV'],
  ['EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA', 'EMEA'],
  ['', '', '', 'SBW', '', 'SBW', '', 'SBW', 'SBW', 'SBW', '', '', ''],
  ['FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'FWD', 'AWD', 'FWD', 'FWD', 'FWD', 'FWD'],
  ['HR', 'HR', 'HR', 'HR', 'HR', 'XR', 'XR', 'XR', 'XR', 'TT', 'TT', 'TT', 'TT'],
  ['HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'HIGH_Range', 'X_Range', 'X_Range', 'X_Range', 'X_Range', '', '', '', ''],
  ['', '', '', '', '', '', '', '', '', '', '', '', ''],
  ['TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET', 'TARGET'],
  ['L2', 'L3', 'L3', 'GT', 'GT', 'L3', 'L3', 'GT', 'GTI', 'L2', 'L3', 'L3', 'GT'],
];

const MNS_DEFAULT_STORAGE_KEY = 'mns-grid-cells-v3';

/** Rows that make up the red bands on the Options SP2 page (columns A–E). */
const SP2_RED_ROWS = new Set([16, 17, 18, 53, 54, 55, 57, 58, 59]);
const SP2_RED_COLS = ['A', 'B', 'C', 'D', 'E'];
/** Options descriptor table — bold black grid on B–E, rows 19–49 (like F–T). */
const OPTIONS_SP2_BDE_COLS = ['B', 'C', 'D', 'E'];
const OPTIONS_SP2_BDE_TABLE_START_ROW = 19;
const OPTIONS_SP2_BDE_TABLE_END_ROW = 49;
const SP2_BDE_COL_SET = new Set(OPTIONS_SP2_BDE_COLS);

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
  // Column B — lower block
  '62:B': { text: 'Maximum mass transfer', cls: 'sp2-b-bold-title' },
  // Column A — rows 44–47 (derived from D41)
  '44:A': { text: 'issu D41', cls: 'sp2-a-note' },
  '45:A': { text: 'issu D41', cls: 'sp2-a-note' },
  '46:A': { text: 'issu D41', cls: 'sp2-a-note' },
  '47:A': { text: 'issu D41', cls: 'sp2-a-note' },
};

const SP2_GREEN_C_ROWS = new Set([22, 28, 32, 33, 40, 41, 42, 43]);
const OPTIONS_SP2_OPTION_GREEN_COLS = [
  ...OPTIONS_SP2_FT_COLS,
  ...OPTIONS_SP2_VAG_COLS,
  ...OPTIONS_SP2_EXTENDED_GREEN_COLS,
];
const SP2_FT_COL_SET = new Set(OPTIONS_SP2_FT_COLS);
const SP2_VAG_COL_SET = new Set(OPTIONS_SP2_VAG_COLS);
const SP2_AU_COL_SET = new Set(OPTIONS_SP2_AU_COLS);
const SP2_OPTION_GREEN_COL_SET = new Set(OPTIONS_SP2_OPTION_GREEN_COLS);
const SP2_BLANK_COL_SET = new Set([OPTIONS_SP2_BLANK_COL, OPTIONS_SP2_BLANK_COL_AH]);
const SP2_COL_WIDTH_OVERRIDES = {
  B: 300,
  U: 24,
  AH: 24,
  F: 100,
  G: 100,
  H: 100,
  I: 74,
  J: 65,
  K: 77,
  L: 77,
  M: 77,
  N: 77,
  O: 77,
  P: 77,
  Q: 83,
  R: 83,
  S: 83,
  T: 77,
};

/** Mirrors Options SP2 column widths from bd-grid.css for virtual-scroll layout. */
function sp2ColWidth(col) {
  if (SP2_COL_WIDTH_OVERRIDES[col] != null) return SP2_COL_WIDTH_OVERRIDES[col];
  if (SP2_VAG_COL_SET.has(col) || SP2_FT_COL_SET.has(col)) return 77;
  if (SP2_AU_COL_SET.has(col)) return 77;
  if (SP2_OPTION_GREEN_COL_SET.has(col)) return 77;
  return SP2_DEFAULT_COL_W;
}

/** Rows 16–18 live SUMIF — columns F…T + V…AG (vehicle blocks). */
const SP2_SUMIF_COLS = [...OPTIONS_SP2_FT_COLS, ...OPTIONS_SP2_VAG_COLS];
const SP2_SUMIF_COL_SET = new Set(SP2_SUMIF_COLS);
const SP2_FT_SUMIF_COL_SET = new Set(OPTIONS_SP2_FT_COLS);
const SP2_ROW16_SUMIF_COL_SET = SP2_SUMIF_COL_SET;
/** Row 15 live sum (row 16 + row 53) — F…T + V…AG. */
const SP2_SUM_COL_SET = new Set([...OPTIONS_SP2_FT_COLS, ...OPTIONS_SP2_VAG_COLS]);
const OPTIONS_SP2_ALL_COLS = Array.from(
  { length: OPTIONS_SP2_COLUMN_COUNT },
  (_, i) => numToCol(i + 1)
);

function isSp2OptionGreenRow(row) {
  return (
    (row >= OPTIONS_SP2_FT_OPTION_START_ROW && row <= 39) ||
    (row >= 41 && row <= OPTIONS_SP2_FT_OPTION_END_ROW)
  );
}

function isSp2OptionBandRow(row) {
  return row >= OPTIONS_SP2_FT_OPTION_START_ROW && row <= OPTIONS_SP2_FT_OPTION_END_ROW;
}

function synTableClassStatic(row, col) {
  if (row < 1 || row > 12) return '';
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

function acanTableClassStatic(row, col) {
  if (row < 1 || row > 12) return '';
  const i = OPTIONS_SP2_VAG_COLS.indexOf(col);
  if (i < 0) return '';
  if (row === 3) return 'sp2-syn-o3w';
  if (row === 4) {
    if (i <= 5) return 'sp2-syn-bev';
    if (i <= 8) return 'sp2-syn-mhevp2';
    return 'sp2-syn-hev';
  }
  if ((row === 8 || row === 9) && i >= 6) return 'sp2-syn-grey';
  if (row === 11) return 'sp2-syn-target';
  return '';
}

/** Synthesis AP…BB colour rules on Options SP2 AI…AU (rows 1–12 = Syn rows 3–14). */
function auTableClassStatic(row, col) {
  if (row < 1 || row > 12) return '';
  const i = OPTIONS_SP2_AU_COLS.indexOf(col);
  if (i < 0) return '';
  if (row === 3) return 'sp2-syn-p3s';
  if (row === 4) {
    if (i <= 8) return 'sp2-syn-bev';
    if (i <= 10) return 'sp2-syn-mhevp2';
    return 'sp2-syn-hev';
  }
  if (row === 7 && i === 8) return 'sp2-syn-awd';
  if ((row === 8 || row === 9) && i >= 9) return 'sp2-syn-grey';
  if (row === 11) return 'sp2-syn-target';
  return '';
}

/** Bold black grid (AI–AU): Synthesis AP…BB header rows 1–12 only. */
function isSp2AuSynFramedRow(row) {
  return row >= 1 && row <= 12;
}

/** Bold black grid (V–AG): header rows 1–12 + numeric body rows 14–18 (like F–T). */
function isSp2VagSynFramedRow(row) {
  return (
    (row >= 1 && row <= 12) ||
    (row >= OPTIONS_SP2_CURB_ROW && row <= OPTIONS_SP2_SYN_TABLE_END_ROW)
  );
}

/** Bold black grid (F–T): header rows 1–12 + numeric body rows 14–18. */
function isSp2SynFramedRow(row) {
  return (
    (row >= 1 && row <= 12) ||
    (row >= OPTIONS_SP2_CURB_ROW && row <= OPTIONS_SP2_SYN_TABLE_END_ROW)
  );
}

/** Bold black grid (B–E): options labels + C/D/E masses, rows 19–49. */
function isSp2BdeOptionsTableRow(row) {
  return (
    row >= OPTIONS_SP2_BDE_TABLE_START_ROW &&
    row <= OPTIONS_SP2_BDE_TABLE_END_ROW
  );
}

function buildSp2StaticClassMap() {
  const map = new Map();
  for (let row = 1; row <= OPTIONS_SP2_MIN_ROWS; row++) {
    for (const col of OPTIONS_SP2_ALL_COLS) {
      const classes = [];
      if (col === 'A' && row >= 1 && row <= 15) classes.push('sp2-brand-cell');
      if (SP2_RED_ROWS.has(row) && SP2_RED_COLS.includes(col)) classes.push('sp2-red-cell');
      if (row === 52 && col === 'D') classes.push('sp2-upgrade-cell');
      if (col === 'C' && SP2_GREEN_C_ROWS.has(row)) classes.push('sp2-green-cell');
      if (col === 'B' && row === 49) classes.push('sp2-blue-cell');
      if (SP2_OPTION_GREEN_COL_SET.has(col) && isSp2OptionBandRow(row)) {
        if (isSp2OptionGreenRow(row)) classes.push('sp2-option-green-band');
        else if (row === OPTIONS_SP2_OPTION_WHITE_ROW) classes.push('sp2-option-white-band');
      }
      if (SP2_BDE_COL_SET.has(col) && isSp2BdeOptionsTableRow(row)) {
        classes.push('sp2-bde-cell');
      }
      if (
        SP2_FT_COL_SET.has(col) &&
        isSp2SynFramedRow(row)
      ) {
        classes.push('sp2-syn-cell');
      }
      if (
        SP2_VAG_COL_SET.has(col) &&
        isSp2VagSynFramedRow(row)
      ) {
        classes.push('sp2-syn-cell', 'sp2-vag-syn-cell');
      }
      if (SP2_AU_COL_SET.has(col) && isSp2AuSynFramedRow(row)) {
        classes.push('sp2-syn-cell', 'sp2-au-syn-cell');
      }
      if (SP2_BLANK_COL_SET.has(col)) classes.push('sp2-blank-col');
      if (row === OPTIONS_SP2_CURB_ROW && isSp2CurbLinkedCol(col)) {
        classes.push('sp2-curb-cell');
      }
      if (row === 15 && SP2_SUM_COL_SET.has(col)) classes.push('sp2-ft-row15-magenta');
      const syn = synTableClassStatic(row, col);
      if (syn) classes.push(syn);
      const acan = acanTableClassStatic(row, col);
      if (acan) classes.push(acan);
      const au = auTableClassStatic(row, col);
      if (au) classes.push(au);
      map.set(`${row}:${col}`, classes);
    }
  }
  return map;
}

const SP2_STATIC_CLASS_MAP = buildSp2StaticClassMap();
const SP2_RENDER_FALLBACK = { kind: 'input', classes: [], label: null };

function buildSp2RenderMap() {
  const map = new Map();
  for (let row = 1; row <= OPTIONS_SP2_MIN_ROWS; row++) {
    for (const col of OPTIONS_SP2_ALL_COLS) {
      const classes = SP2_STATIC_CLASS_MAP.get(`${row}:${col}`) || [];
      const label = SP2_STATIC_LABELS[`${row}:${col}`] || null;
      let kind = 'input';
      if (col === 'A' && row >= 1 && row <= 15) {
        kind = row === 6 ? 'brand-title' : row === 7 ? 'brand-sub' : 'brand-empty';
      } else if (label) {
        kind = 'red-label';
      } else if (row === OPTIONS_SP2_CURB_ROW && isSp2CurbLinkedCol(col)) {
        kind = 'curb';
      } else if (isSp2FtSumifCell(row, col)) {
        kind = 'row-ft-sumif';
      } else if (row === OPTIONS_SP2_SUM_ROW && SP2_SUM_COL_SET.has(col)) {
        kind = 'sum';
      } else if (SP2_BLANK_COL_SET.has(col)) {
        kind = 'blank';
      } else if (
        classes.includes('sp2-option-green-band') ||
        classes.includes('sp2-option-white-band')
      ) {
        kind = 'option';
      }
      map.set(`${row}:${col}`, { kind, classes, label });
    }
  }
  return map;
}

const SP2_RENDER_MAP = buildSp2RenderMap();

function parseSp2Num(raw) {
  const s = String(raw ?? '')
    .trim()
    .replace(',', '.');
  if (s === '') return null;
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
}

function isSp2SumifDataRow(row) {
  if (row === OPTIONS_SP2_OPTION_WHITE_ROW) return false;
  return row >= OPTIONS_SP2_SUMIF_START_ROW && row <= OPTIONS_SP2_SUMIF_END_ROW;
}

/** One SUMIF output row: SOMME.SI(col19:col47,"O",sumCol19:sumCol47). */
function computeSp2FtSumifRow(cellsObj, sumCol) {
  const out = {};
  for (const col of SP2_SUMIF_COLS) {
    let sum = 0;
    for (let row = OPTIONS_SP2_SUMIF_START_ROW; row <= OPTIONS_SP2_SUMIF_END_ROW; row++) {
      if (!isSp2SumifDataRow(row)) continue;
      const choice = String(cellsObj[`${row}:${col}`] ?? '').trim().toUpperCase();
      if (choice !== OPTIONS_SP2_ROW16_SUMIF_CRITERIA) continue;
      const n = parseSp2Num(cellsObj[`${row}:${sumCol}`]);
      if (n != null) sum += n;
    }
    out[col] = sum.toFixed(2);
  }
  return out;
}

function computeSp2FtSumifRows(cellsObj) {
  const out = {};
  for (const { row, sumCol } of OPTIONS_SP2_FT_SUMIF_ROWS) {
    out[row] = computeSp2FtSumifRow(cellsObj, sumCol);
  }
  return out;
}

function computeSp2Row16Sumifs(cellsObj) {
  return computeSp2FtSumifRow(cellsObj, 'C');
}

function computeSp2Row15Sums(cellsObj, ftSumifs) {
  const all = ftSumifs || computeSp2FtSumifRows(cellsObj);
  const r16 = all[OPTIONS_SP2_ROW16_SUMIF_ROW] || computeSp2Row16Sumifs(cellsObj);
  const out = {};
  for (const col of SP2_SUM_COL_SET) {
    const a = parseSp2Num(r16[col]);
    const b = parseSp2Num(cellsObj[`${OPTIONS_SP2_SUM_SRC_ROWS[1]}:${col}`]);
    if (a == null && b == null) out[col] = '';
    else out[col] = ((a ?? 0) + (b ?? 0)).toFixed(2);
  }
  return out;
}

function affectsSp2FtSumif(row, col) {
  if (!isSp2SumifDataRow(row)) return false;
  return OPTIONS_SP2_FT_SUMIF_SUM_COLS.has(col) || SP2_SUMIF_COL_SET.has(col);
}

function affectsSp2Row16Sumif(row, col) {
  return affectsSp2FtSumif(row, col);
}

function isSp2FtSumifCell(row, col) {
  if (!SP2_SUMIF_COL_SET.has(col)) return false;
  return OPTIONS_SP2_FT_SUMIF_ROWS.some((e) => e.row === row);
}

function isSp2Row16SumifCell(row, col) {
  return isSp2FtSumifCell(row, col);
}

function fieldsForKey(storageKey) {
  return String(storageKey || '').includes('options-sp2')
    ? OPTIONS_SP2_FIELDS_COL_B
    : MNS_FIELDS_COL_B;
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
    // Equipment band → columns F–T & V–AG, rows 53–55.
    OPTIONS_SP2_EQUIPMENT_ROWS.forEach((row) => {
      for (const col of [...OPTIONS_SP2_FT_COLS, ...OPTIONS_SP2_VAG_COLS]) {
        cells[`${row}:${col}`] = OPTIONS_SP2_EQUIPMENT_FT_VALUE;
      }
    });
    // Rows 16–18 F…T = live SUMIF (C / D / E); no static seed on F…T.
    // Synthesis AC…AN header table → columns V–AG, rows 1–12.
    OPTIONS_SP2_VAG.forEach((rowVals, i) => {
      const row = i + 1;
      rowVals.forEach((v, j) => {
        if (v !== '') cells[`${row}:${OPTIONS_SP2_VAG_COLS[j]}`] = v;
      });
    });
    // Rows 16–18 V…AG = live SUMIF (FRONT D + BACK E → row 16); no static seed.
    // Options picker — columns V–AG, rows 19–47 (green / white bands).
    OPTIONS_SP2_VAG_OPTIONS.forEach((rowVals, i) => {
      const row = OPTIONS_SP2_FT_OPTION_START_ROW + i;
      rowVals.forEach((v, j) => {
        cells[`${row}:${OPTIONS_SP2_VAG_COLS[j]}`] = v;
      });
    });
    // Options picker — columns F–T, rows 19–47 (O / NO / REF from source sheet).
    OPTIONS_SP2_FT_OPTIONS.forEach((rowVals, i) => {
      const row = OPTIONS_SP2_FT_OPTION_START_ROW + i;
      rowVals.forEach((v, j) => {
        cells[`${row}:${OPTIONS_SP2_FT_COLS[j]}`] = v;
      });
    });
    // Synthesis AP…BB header table → columns AI…AU, rows 1–12.
    OPTIONS_SP2_AU.forEach((rowVals, i) => {
      const row = i + 1;
      rowVals.forEach((v, j) => {
        if (v !== '') cells[`${row}:${OPTIONS_SP2_AU_COLS[j]}`] = v;
      });
    });
    // Extended green cols AV–HC — rows 44–47 pre-filled NO.
    OPTIONS_SP2_NO_SEED_ROWS.forEach((row) => {
      OPTIONS_SP2_EXTENDED_GREEN_COLS.forEach((col) => {
        cells[`${row}:${col}`] = OPTIONS_SP2_NO_SEED_VALUE;
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
    const cells = shallowRef(buildSeedCells(props.storageKey));
    const editTick = ref(0);
    let persistTimer = 0;

    const {
      syncFromCell: syncAxisFromCell,
      onRowNumClick: onAxisRowNumClick,
      onColHeaderClick: onAxisColHeaderClick,
      isAxisRow,
      isAxisCol,
    } = createGridAxisHighlight();

    function onCellAxisSelect(row, col, event) {
      if (event && event.button != null && event.button !== 0) return;
      syncAxisFromCell(row, col);
    }

    const ftSumifRows = ref({});
    const row15Sums = ref({});

    // The Options SP2 page reuses this grid but brands column A (rows 1–12)
    // with a white "STLA S / SP2" header block instead of editable cells.
    const isOptionsSp2 = computed(() =>
      String(props.storageKey || '').includes('options-sp2')
    );

    // Options SP2: A…AG — F–T = M…AA copy, U blank, V–AG = AC…AN copy (rows 3–14).
    const columns = computed(() => {
      const count = isOptionsSp2.value ? OPTIONS_SP2_COLUMN_COUNT : MNS_COLUMN_COUNT;
      return Array.from({ length: count }, (_, i) => numToCol(i + 1));
    });

    // Live link to the Synthesis sheet (provided by the app root). Used to mirror
    // the Curb mass row (Synthesis 16 → Options SP2 row 14) in real time.
    const synLink = inject('synthesisCellLink', null);

    // Options SP2 row 14 ← Synthesis row 16 display M…AA (SP2 F…T) + AC…AN (SP2 V…AG).
    const curbMassRow = computed(() => {
      try {
        if (!isOptionsSp2.value || !synLink) return {};
        if (synLink.synRevision) void synLink.synRevision.value;
        if (synLink.synEditTick) void synLink.synEditTick.value;
        if (synLink.session?.synCalcTick) void synLink.session.synCalcTick.value;
        if (synLink.session?.displayTick) void synLink.session.displayTick.value;
        if (synLink.session?.liveBdEdited) void synLink.session.liveBdEdited.value;
        const getRow16 = synLink.getSynRow16Display;
        if (typeof getRow16 !== 'function') return {};

        const out = {};
        for (const sp2Col of [...OPTIONS_SP2_FT_COLS, ...OPTIONS_SP2_VAG_COLS]) {
          if (!isSp2CurbLinkedCol(sp2Col)) continue;
          const synExcelCol = sp2ColToSynExcelCol(sp2Col);
          if (!synExcelCol) continue;
          const shown = getRow16(synExcelCol);
          if (shown != null && String(shown).trim() !== '') {
            out[sp2Col] = String(shown);
          }
        }
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
        isSp2CurbLinkedCol(col)
      );
    }

    // Copied Synthesis table blocks: F–T & V–AG (rows 1–12 & 14–18).
    function isSynTableCell(row, col) {
      if (!isOptionsSp2.value) return false;
      if (OPTIONS_SP2_FT_COLS.includes(col) && isSp2SynFramedRow(row)) return true;
      if (OPTIONS_SP2_VAG_COLS.includes(col) && isSp2VagSynFramedRow(row)) return true;
      return OPTIONS_SP2_AU_COLS.includes(col) && isSp2AuSynFramedRow(row);
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

    function sp2Render(row, col) {
      return SP2_RENDER_MAP.get(`${row}:${col}`) || SP2_RENDER_FALLBACK;
    }

    function recomputeSp2DerivedSums() {
      if (!isOptionsSp2.value) return;
      const sums = computeSp2FtSumifRows(cells.value);
      ftSumifRows.value = sums;
      row15Sums.value = computeSp2Row15Sums(cells.value, sums);
      touchCells();
    }

    function rowFtSumifValue(row, col) {
      const byRow = ftSumifRows.value[row];
      return (byRow && byRow[col]) || '';
    }

    function rowFtSumifTitle(row) {
      const entry = OPTIONS_SP2_FT_SUMIF_ROWS.find((e) => e.row === row);
      const sumCol = entry ? entry.sumCol : 'C';
      return `SOMME.SI F…AG : somme colonne ${sumCol} (l.19–49) si cette colonne = O`;
    }

    function row16SumifValue(col) {
      return rowFtSumifValue(OPTIONS_SP2_ROW16_SUMIF_ROW, col);
    }

    /** Row 15 = row 16 + row 53 (F–T columns only), read-only. */
    function isRow15SumCell(row, col) {
      return isOptionsSp2.value && row === OPTIONS_SP2_SUM_ROW && SP2_SUM_COL_SET.has(col);
    }

    function isRow15MagentaCol(col) {
      return SP2_SUM_COL_SET.has(col);
    }

    function isSp2BlankCol(col) {
      return SP2_BLANK_COL_SET.has(col);
    }

    function bindSp2Input(el, row, col) {
      if (!el) return;
      const key = cellKey(row, col);
      if (el.dataset.sp2Key === key) return;
      el.dataset.sp2Key = key;
      el.value = cells.value[key] || '';
    }

    function onSp2CellInput(row, col, event) {
      const value = event && event.target ? event.target.value : '';
      const key = cellKey(row, col);
      const cur = cells.value[key];
      let changed = false;
      if (value === '') {
        if (cur !== undefined) {
          delete cells.value[key];
          changed = true;
        }
      } else if (cur !== value) {
        cells.value[key] = value;
        changed = true;
      }
      if (!changed) return;
      schedulePersist();
      if (
        affectsSp2FtSumif(row, col) ||
        (row === OPTIONS_SP2_SUM_SRC_ROWS[1] && SP2_SUM_COL_SET.has(col))
      ) {
        recomputeSp2DerivedSums();
      } else {
        touchCells();
      }
    }

    function syncSp2OptionButton(el, row, col) {
      if (!el) return;
      const val = cells.value[cellKey(row, col)] || '';
      el.textContent = val;
      el.classList.toggle('sp2-option-choice-no', val === 'NO');
      el.classList.toggle('sp2-option-choice-default', val !== 'NO');
    }

    function bindSp2Option(el, row, col) {
      syncSp2OptionButton(el, row, col);
    }

    function cycleSp2Option(row, col, event) {
      syncAxisFromCell(row, col);
      const key = cellKey(row, col);
      const cur = cells.value[key] || '';
      const idx = SP2_OPTION_CYCLE.indexOf(cur);
      const next = SP2_OPTION_CYCLE[(idx + 1) % SP2_OPTION_CYCLE.length];
      if (next === '') delete cells.value[key];
      else cells.value[key] = next;
      syncSp2OptionButton(event.currentTarget, row, col);
      schedulePersist();
      if (affectsSp2FtSumif(row, col)) {
        recomputeSp2DerivedSums();
      }
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

    // Blue spot cells (row 15, columns H–I) with centred labels.
    function isMnsBlueCell(row, col) {
      if (isOptionsSp2.value) return false;
      return Object.prototype.hasOwnProperty.call(MNS_BLUE_LABELS, `${row}:${col}`);
    }

    function mnsBlueLabel(row, col) {
      if (isOptionsSp2.value) return null;
      return MNS_BLUE_LABELS[`${row}:${col}`] || null;
    }

    // Curb-mass mini table (G–I, rows 17–21) — bold black grid.
    function isMnsHiTableCell(row, col) {
      if (isOptionsSp2.value) return false;
      return MNS_HI_TABLE_ROWS.includes(row) && MNS_HI_TABLE_COLS.includes(col);
    }

    function mnsHiTableLabel(row, col) {
      if (isOptionsSp2.value) return null;
      return MNS_HI_TABLE_LABELS[`${row}:${col}`] || null;
    }

    // Green variant headers (row 15, columns K–N).
    function isMnsLimeCell(row, col) {
      if (isOptionsSp2.value) return false;
      return Object.prototype.hasOwnProperty.call(MNS_LIME_LABELS, `${row}:${col}`);
    }

    function mnsLimeLabel(row, col) {
      if (isOptionsSp2.value) return null;
      return MNS_LIME_LABELS[`${row}:${col}`] || null;
    }

    // Curb-mass mini table (K–N, rows 18–21) — bold black grid.
    function isMnsKnTableCell(row, col) {
      if (isOptionsSp2.value) return false;
      return MNS_KN_TABLE_ROWS.includes(row) && MNS_KN_TABLE_COLS.includes(col);
    }

    // MNS page: dark-green paint bands (rows 23, 24, 42, columns A–E; G on 24 & 42).
    function isMnsGreenCell(row, col) {
      if (isOptionsSp2.value) return false;
      if (!MNS_GREEN_ROWS.includes(row)) return false;
      if (MNS_GREEN_COLS.includes(col)) return true;
      return col === 'G' && (row === 24 || row === 42);
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
    function isGreenCell(row, col) {
      return isOptionsSp2.value && col === 'C' && SP2_GREEN_C_ROWS.has(row);
    }
    function isBlueCell(row, col) {
      return isOptionsSp2.value && col === 'B' && row === 49;
    }

    /** Options band rows 19–47 — green (#92d050) + white row 40, cols F–T + V–AG + AV–HC. */
    function isOptionGreenBandCell(row, col) {
      if (!isOptionsSp2.value || !SP2_OPTION_GREEN_COL_SET.has(col)) return false;
      return isSp2OptionGreenRow(row);
    }

    function isSp2OptionCol(col) {
      return SP2_OPTION_GREEN_COL_SET.has(col);
    }

    // Fixed labels (red bands + the standalone "Montée en gamme" header).
    function redLabel(row, col) {
      if (!isOptionsSp2.value) return null;
      return SP2_STATIC_LABELS[`${row}:${col}`] || null;
    }

    const rowCount = computed(() => {
      // Options SP2 needs room for the red bands / labels down to row 69.
      let maxRow = isOptionsSp2.value ? OPTIONS_SP2_MIN_ROWS : MNS_MIN_ROWS;
      for (const key of Object.keys(cells.value)) {
        const r = Number(key.split(':')[0]);
        if (Number.isFinite(r) && r > maxRow) maxRow = r;
      }
      return maxRow;
    });

    const rows = computed(() =>
      Array.from({ length: rowCount.value }, (_, i) => i + 1)
    );

    // ── Options SP2 virtual scroll (211 cols × ~70 rows — render only the viewport)
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const scrollLeft = ref(0);
    const viewportH = ref(600);
    const viewportW = ref(1200);
    let sp2ResizeObs = null;

    const scrollSync = createScrollRafSync({
      scrollTop,
      scrollLeft,
      getScrollEl: () => scrollEl.value,
    });

    function updateSp2Viewport() {
      if (!scrollEl.value) return;
      viewportH.value = scrollEl.value.clientHeight;
      viewportW.value = scrollEl.value.clientWidth;
    }

    const sp2ScrollColumns = computed(() =>
      columns.value.filter((c) => !SP2_STICKY_COLS.includes(c))
    );

    const sp2ColLayout = computed(() => {
      let left = 0;
      const out = [];
      for (const col of sp2ScrollColumns.value) {
        const width = sp2ColWidth(col);
        out.push({ col, left, width });
        left += width;
      }
      return out;
    });

    const sp2ScrollableWidth = computed(() => {
      const layout = sp2ColLayout.value;
      return layout.length
        ? layout[layout.length - 1].left + layout[layout.length - 1].width
        : 0;
    });

    const sp2StickyWidth = computed(
      () => SP2_ROW_NUM_W + sp2ColWidth('A') + sp2ColWidth('B')
    );

    const sp2VirtualizeCols = computed(
      () =>
        isOptionsSp2.value &&
        shouldVirtualizeCols(sp2StickyWidth.value + sp2ScrollableWidth.value, viewportW.value)
    );

    const sp2ColScrollCache = createRangeScrollCache(SP2_MAX_RENDERED_COLS, {
      monotonic: false,
    });
    const sp2ColRangeStart = ref(0);
    const sp2ColRangeEnd = ref(0);

    watchEffect(() => {
      if (!isOptionsSp2.value) return;
      const layout = sp2ColLayout.value;
      if (!sp2VirtualizeCols.value) {
        sp2ColRangeStart.value = 0;
        sp2ColRangeEnd.value = layout.length;
        return;
      }
      const vpScrollW = Math.max(100, viewportW.value - sp2StickyWidth.value);
      const range = sp2ColScrollCache.resolveCols(
        layout,
        scrollLeft.value,
        vpScrollW,
        colOverscanPx(viewportW.value, 480)
      );
      sp2ColRangeStart.value = range.start;
      sp2ColRangeEnd.value = range.end;
    });

    const sp2RenderColumns = computed(() => {
      if (!isOptionsSp2.value) return columns.value;
      if (!sp2VirtualizeCols.value) return columns.value;
      const out = [...SP2_STICKY_COLS];
      const layout = sp2ColLayout.value;
      for (let i = sp2ColRangeStart.value; i < sp2ColRangeEnd.value; i++) {
        if (layout[i]) out.push(layout[i].col);
      }
      return out;
    });

    const headerColumns = computed(() =>
      isOptionsSp2.value ? sp2RenderColumns.value : columns.value
    );

    const sp2LeftPad = computed(() => {
      if (!isOptionsSp2.value || !sp2VirtualizeCols.value) return 0;
      const first = sp2ColLayout.value[sp2ColRangeStart.value];
      return first ? first.left : 0;
    });

    const sp2RightPad = computed(() => {
      if (!isOptionsSp2.value || !sp2VirtualizeCols.value) return 0;
      const last = sp2ColLayout.value[sp2ColRangeEnd.value - 1];
      if (!last) return 0;
      return Math.max(0, sp2ScrollableWidth.value - (last.left + last.width));
    });

    const sp2TableColSpan = computed(() => {
      if (!isOptionsSp2.value) return columns.value.length + 1;
      let n = 1 + headerColumns.value.length;
      if (sp2VirtualizeCols.value && sp2LeftPad.value > 0) n += 1;
      if (sp2VirtualizeCols.value && sp2RightPad.value > 0) n += 1;
      return n;
    });

    const sp2VirtualizeRows = computed(
      () =>
        isOptionsSp2.value &&
        shouldVirtualizeRows(rowCount.value, viewportH.value)
    );

    const sp2RowScrollCache = createRangeScrollCache(SP2_MAX_RENDERED_ROWS, {
      monotonic: true,
    });
    const sp2VisibleStart = ref(0);
    const sp2VisibleEnd = ref(0);

    watchEffect(() => {
      if (!isOptionsSp2.value) return;
      const count = rowCount.value;
      if (!sp2VirtualizeRows.value) {
        sp2VisibleStart.value = 0;
        sp2VisibleEnd.value = count;
        return;
      }
      const range = sp2RowScrollCache.resolveRows(
        scrollTop.value,
        viewportH.value,
        count,
        rowOverscan(viewportH.value, 20, 40)
      );
      sp2VisibleStart.value = range.start;
      sp2VisibleEnd.value = range.end;
    });

    const sp2VisibleRows = computed(() => {
      if (!isOptionsSp2.value || !sp2VirtualizeRows.value) return rows.value;
      return rows.value.slice(sp2VisibleStart.value, sp2VisibleEnd.value);
    });

    const sp2TopSpacer = computed(() =>
      isOptionsSp2.value && sp2VirtualizeRows.value
        ? sp2VisibleStart.value * ROW_H
        : 0
    );

    const sp2BottomSpacer = computed(() => {
      if (!isOptionsSp2.value || !sp2VirtualizeRows.value) return 0;
      const shown = sp2VisibleEnd.value - sp2VisibleStart.value;
      const remaining = rowCount.value - sp2VisibleStart.value - shown;
      return Math.max(0, remaining * ROW_H);
    });

    function cellKey(row, col) {
      return `${row}:${col}`;
    }

    function cellValue(row, col) {
      void editTick.value;
      return cells.value[cellKey(row, col)] || '';
    }

    function touchCells() {
      editTick.value += 1;
    }

    function schedulePersist() {
      if (typeof localStorage === 'undefined') return;
      clearTimeout(persistTimer);
      persistTimer = setTimeout(persist, 300);
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
      let merged = stored
        ? { ...buildSeedCells(storageKey), ...stored }
        : buildSeedCells(storageKey);
      if (String(storageKey || '').includes('options-sp2')) {
        for (const key of Object.keys(merged)) {
          if (key.startsWith(`${OPTIONS_SP2_SUM_ROW}:`)) delete merged[key];
          for (const { row } of OPTIONS_SP2_FT_SUMIF_ROWS) {
            if (!key.startsWith(`${row}:`)) continue;
            const col = key.split(':')[1];
            if (SP2_SUMIF_COL_SET.has(col)) delete merged[key];
          }
        }
      }
      cells.value = merged;
      if (String(storageKey || '').includes('options-sp2')) {
        recomputeSp2DerivedSums();
      } else {
        touchCells();
      }
    }

    function onCellInput(row, col, event) {
      const value = event && event.target ? event.target.value : '';
      const key = cellKey(row, col);
      const cur = cells.value[key];
      if (value === '') {
        if (cur !== undefined) {
          delete cells.value[key];
          touchCells();
          schedulePersist();
        }
      } else if (cur !== value) {
        cells.value[key] = value;
        touchCells();
        schedulePersist();
      }
    }

    onMounted(() => {
      loadForKey(props.storageKey);
      updateSp2Viewport();
      window.addEventListener('resize', updateSp2Viewport);
      if (scrollEl.value && typeof ResizeObserver !== 'undefined') {
        sp2ResizeObs = new ResizeObserver(updateSp2Viewport);
        sp2ResizeObs.observe(scrollEl.value);
      }
      nextTick(() => scrollSync.flush());
      if (!isOptionsSp2.value || !synLink) return;
      const boot =
        typeof synLink.ensureSynGrid === 'function'
          ? synLink.ensureSynGrid()
          : typeof synLink.ensureSyn === 'function'
            ? synLink.ensureSyn()
            : null;
      if (boot) Promise.resolve(boot).catch(() => {});
    });

    onUnmounted(() => {
      clearTimeout(persistTimer);
      window.removeEventListener('resize', updateSp2Viewport);
      if (sp2ResizeObs) {
        sp2ResizeObs.disconnect();
        sp2ResizeObs = null;
      }
      scrollSync.dispose();
    });

    // Reload the right dataset if the same component instance is reused for
    // another page (different storageKey).
    watch(
      () => props.storageKey,
      (key) => loadForKey(key)
    );

    watch(rowCount, () => {
      sp2RowScrollCache.reset();
      nextTick(() => scrollSync.flush());
    });

    return {
      columns,
      headerColumns,
      rows,
      sp2VisibleRows,
      sp2VirtualizeCols,
      sp2LeftPad,
      sp2RightPad,
      sp2TopSpacer,
      sp2BottomSpacer,
      sp2TableColSpan,
      scrollEl,
      onScroll: scrollSync.onScroll,
      cellValue,
      onCellInput,
      onCellAxisSelect,
      onAxisRowNumClick,
      onAxisColHeaderClick,
      isAxisRow,
      isAxisCol,
      isOptionsSp2,
      sp2Render,
      bindSp2Input,
      onSp2CellInput,
      bindSp2Option,
      cycleSp2Option,
      ftSumifRows,
      rowFtSumifValue,
      rowFtSumifTitle,
      row16SumifValue,
      row15Sums,
      isBrandCell,
      isMnsPaintCell,
      isMnsLabelCell,
      isMnsRedCell,
      mnsRedLabel,
      isMnsBlueCell,
      mnsBlueLabel,
      isMnsHiTableCell,
      mnsHiTableLabel,
      isMnsLimeCell,
      mnsLimeLabel,
      isMnsKnTableCell,
      isMnsGreenCell,
      mnsGreenLabel,
      isMnsYellowCell,
      isMnsBeigeCell,
      isMnsRedValue,
      isRedCell,
      isUpgradeCell,
      isGreenCell,
      isBlueCell,
      isOptionGreenBandCell,
      isSp2OptionCol,
      redLabel,
      isCurbLinkedCell,
      curbLinkedValue,
      isRow15SumCell,
      isRow15MagentaCol,
      synTableClass,
      isSynTableCell,
    };
  },
  template: `
    <div class="bd-grid-root mns-grid" :class="{ 'options-sp2-grid': isOptionsSp2 }">
      <div class="bd-grid-scroll" ref="scrollEl" @scroll.passive="onScroll">
        <table class="bd-table mns-table" role="grid">
          <thead>
            <tr class="hdr-row-letters">
              <th class="corner"></th>
              <template v-for="(col, idx) in headerColumns" :key="'mns-letter-' + col">
                <th
                  v-if="isOptionsSp2 && sp2VirtualizeCols && idx === 2 && sp2LeftPad > 0"
                  class="col-spacer"
                  :style="{ width: sp2LeftPad + 'px', minWidth: sp2LeftPad + 'px', maxWidth: sp2LeftPad + 'px', padding: 0, border: 'none' }"
                ></th>
                <th
                  class="col-letter grid-axis-hdr"
                  :class="{
                    'sp2-option-col-letter': isOptionsSp2 && isSp2OptionCol(col),
                    'sp2-sticky-col-b': isOptionsSp2 && col === 'B',
                    'grid-axis-hdr-focus': isAxisCol(col),
                  }"
                  :data-col="col"
                  @mousedown="onAxisColHeaderClick(col, $event)"
                >{{ col }}</th>
              </template>
              <th
                v-if="isOptionsSp2 && sp2VirtualizeCols && sp2RightPad > 0"
                class="col-spacer"
                :style="{ width: sp2RightPad + 'px', minWidth: sp2RightPad + 'px', maxWidth: sp2RightPad + 'px', padding: 0, border: 'none' }"
              ></th>
            </tr>
          </thead>
          <tbody v-if="isOptionsSp2">
            <tr v-if="sp2TopSpacer > 0" class="bd-spacer-top">
              <td :colspan="sp2TableColSpan" :style="{ height: sp2TopSpacer + 'px', padding: 0, border: 'none' }"></td>
            </tr>
            <tr
              v-for="row in sp2VisibleRows"
              :key="'sp2-row-' + row"
              class="grid-row-cv"
              :class="{ 'grid-axis-row-focus': isAxisRow(row) }"
            >
              <td
                class="row-num grid-axis-hdr"
                :class="{ 'grid-axis-hdr-focus': isAxisRow(row) }"
                @mousedown="onAxisRowNumClick(row, $event)"
              >{{ row }}</td>
              <template v-for="(col, idx) in headerColumns" :key="'sp2-' + row + '-' + col">
                <td
                  v-if="sp2VirtualizeCols && idx === 2 && sp2LeftPad > 0"
                  class="col-spacer"
                  :style="{ width: sp2LeftPad + 'px', minWidth: sp2LeftPad + 'px', maxWidth: sp2LeftPad + 'px', padding: 0, border: 'none' }"
                ></td>
                <td
                  class="data-cell"
                  :data-col="col"
                  :class="[
                    sp2Render(row, col).classes,
                    {
                      'sp2-sticky-col-b': col === 'B' && !isRedCell(row, col),
                      'grid-axis-col-focus': isAxisCol(col),
                    },
                  ]"
                >
                <span
                  v-if="sp2Render(row, col).kind === 'brand-title'"
                  class="sp2-brand-title"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >STLA S</span>
                <span
                  v-else-if="sp2Render(row, col).kind === 'brand-sub'"
                  class="sp2-brand-sub"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >SP2</span>
                <span
                  v-else-if="sp2Render(row, col).kind === 'red-label'"
                  :class="sp2Render(row, col).label.cls"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ sp2Render(row, col).label.text }}</span>
                <span
                  v-else-if="sp2Render(row, col).kind === 'curb'"
                  class="sp2-curb-linked"
                  title="Lié à Synthesis (Curb mass) — mis à jour automatiquement"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ curbLinkedValue(col) }}</span>
                <span
                  v-else-if="sp2Render(row, col).kind === 'row-ft-sumif'"
                  class="sp2-row16-sumif"
                  :title="rowFtSumifTitle(row)"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ rowFtSumifValue(row, col) }}</span>
                <span
                  v-else-if="sp2Render(row, col).kind === 'sum'"
                  class="sp2-row15-sum"
                  title="Somme ligne 16 + ligne 53 (même colonne)"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ row15Sums[col] }}</span>
                <span
                  v-else-if="sp2Render(row, col).kind === 'blank'"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                ></span>
                <button
                  v-else-if="sp2Render(row, col).kind === 'option'"
                  type="button"
                  class="grid-cell-input sp2-option-btn sp2-option-choice-default"
                  :ref="el => bindSp2Option(el, row, col)"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                  @click="cycleSp2Option(row, col, $event)"
                ></button>
                <input
                  v-else
                  :ref="el => bindSp2Input(el, row, col)"
                  type="text"
                  class="grid-cell-input mns-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                  @focus="onCellAxisSelect(row, col, $event)"
                  @input="onSp2CellInput(row, col, $event)"
                />
              </td>
              </template>
              <td
                v-if="sp2VirtualizeCols && sp2RightPad > 0"
                class="col-spacer"
                :style="{ width: sp2RightPad + 'px', minWidth: sp2RightPad + 'px', maxWidth: sp2RightPad + 'px', padding: 0, border: 'none' }"
              ></td>
            </tr>
            <tr v-if="sp2BottomSpacer > 0" class="bd-spacer-bottom">
              <td :colspan="sp2TableColSpan" :style="{ height: sp2BottomSpacer + 'px', padding: 0, border: 'none' }"></td>
            </tr>
          </tbody>
          <tbody v-else>
            <tr
              v-for="row in rows"
              :key="'mns-row-' + row"
              class="grid-row-cv"
              :class="{ 'grid-axis-row-focus': isAxisRow(row) }"
            >
              <td
                class="row-num grid-axis-hdr"
                :class="{ 'grid-axis-hdr-focus': isAxisRow(row) }"
                @mousedown="onAxisRowNumClick(row, $event)"
              >{{ row }}</td>
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
                  'sp2-option-green-band': isOptionGreenBandCell(row, col),
                  'mns-paint-cell': isMnsPaintCell(row, col),
                  'mns-e-label': isMnsLabelCell(row, col),
                  'mns-red-cell': isMnsRedCell(row, col),
                  'mns-blue-cell': isMnsBlueCell(row, col),
                  'mns-lime-cell': isMnsLimeCell(row, col),
                  'mns-hi-table-cell': isMnsHiTableCell(row, col),
                  'mns-kn-table-cell': isMnsKnTableCell(row, col),
                  'mns-green-cell': isMnsGreenCell(row, col),
                  'mns-yellow-cell': isMnsYellowCell(row, col),
                  'mns-beige-cell': isMnsBeigeCell(row, col),
                  'mns-red-value': isMnsRedValue(row, col),
                  'sp2-syn-cell': isSynTableCell(row, col),
                  'grid-axis-col-focus': isAxisCol(col),
                }, synTableClass(row, col)]"
              >
                <template v-if="isBrandCell(row, col)">
                  <span
                    v-if="row === 6"
                    class="sp2-brand-title"
                    @mousedown="onCellAxisSelect(row, col, $event)"
                  >STLA S</span>
                  <span
                    v-else-if="row === 7"
                    class="sp2-brand-sub"
                    @mousedown="onCellAxisSelect(row, col, $event)"
                  >SP2</span>
                </template>
                <span
                  v-else-if="redLabel(row, col)"
                  :class="redLabel(row, col).cls"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ redLabel(row, col).text }}</span>
                <span
                  v-else-if="mnsRedLabel(row, col)"
                  class="mns-red-label"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ mnsRedLabel(row, col) }}</span>
                <span
                  v-else-if="mnsBlueLabel(row, col)"
                  class="mns-blue-label"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ mnsBlueLabel(row, col) }}</span>
                <span
                  v-else-if="mnsLimeLabel(row, col)"
                  class="mns-lime-label"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ mnsLimeLabel(row, col) }}</span>
                <span
                  v-else-if="mnsHiTableLabel(row, col)"
                  class="mns-hi-table-label"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ mnsHiTableLabel(row, col) }}</span>
                <span
                  v-else-if="mnsGreenLabel(row, col)"
                  class="mns-green-label"
                  :class="'mns-green-' + mnsGreenLabel(row, col).align"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ mnsGreenLabel(row, col).text }}</span>
                <span
                  v-else-if="isCurbLinkedCell(row, col)"
                  class="sp2-curb-linked"
                  title="Lié à Synthesis (Curb mass) — mis à jour automatiquement"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ curbLinkedValue(col) }}</span>
                <input
                  v-else
                  type="text"
                  class="grid-cell-input mns-cell-input"
                  autocomplete="off"
                  spellcheck="false"
                  :value="cellValue(row, col)"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                  @focus="onCellAxisSelect(row, col, $event)"
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
