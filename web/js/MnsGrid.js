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
import { ref, computed, onMounted, watch } from 'vue';

const MNS_COLUMN_COUNT = 12;
const MNS_MIN_ROWS = 24;

/** Field labels seeded into column B, one per row starting at row 1. */
const MNS_FIELDS_COL_B = [
  'PLATEFORM',
  'PROJECT',
  'SILHOUETTE',
  'HYBRIDIZATION',
  'DESIGN PLATE',
  'SEATS',
  'TECHNICAL SPECIFICATIONS',
  'POLE',
  'ENERGY',
  'TECHNICAL PACK',
  'TRIM',
  'CURB MASS',
];

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

const MNS_DEFAULT_STORAGE_KEY = 'mns-grid-cells-v1';

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

    const columns = computed(() =>
      Array.from({ length: MNS_COLUMN_COUNT }, (_, i) => numToCol(i + 1))
    );

    // The Options SP2 page reuses this grid but brands column A (rows 1–12)
    // with a white "STLA S / SP2" header block instead of editable cells.
    const isOptionsSp2 = computed(() =>
      String(props.storageKey || '').includes('options-sp2')
    );

    function isBrandCell(row, col) {
      return isOptionsSp2.value && col === 'A' && row >= 1 && row <= 15;
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

    onMounted(() => loadForKey(props.storageKey));

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
      isRedCell,
      isUpgradeCell,
      isGreenCell,
      isBlueCell,
      redLabel,
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
                :class="{
                  'sp2-brand-cell': isBrandCell(row, col),
                  'sp2-red-cell': isRedCell(row, col),
                  'sp2-upgrade-cell': isUpgradeCell(row, col),
                  'sp2-green-cell': isGreenCell(row, col),
                  'sp2-blue-cell': isBlueCell(row, col),
                }"
              >
                <template v-if="isBrandCell(row, col)">
                  <span v-if="row === 6" class="sp2-brand-title">STLA S</span>
                  <span v-else-if="row === 7" class="sp2-brand-sub">SP2</span>
                </template>
                <span
                  v-else-if="redLabel(row, col)"
                  :class="redLabel(row, col).cls"
                >{{ redLabel(row, col).text }}</span>
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
