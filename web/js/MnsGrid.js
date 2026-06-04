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

const MNS_DEFAULT_STORAGE_KEY = 'mns-grid-cells-v1';

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

function buildSeedCells() {
  const cells = {};
  MNS_FIELDS_COL_B.forEach((label, idx) => {
    cells[`${idx + 1}:B`] = label;
  });
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
    const cells = ref(buildSeedCells());

    const columns = computed(() =>
      Array.from({ length: MNS_COLUMN_COUNT }, (_, i) => numToCol(i + 1))
    );

    const rowCount = computed(() => {
      let maxRow = MNS_MIN_ROWS;
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
        ? { ...buildSeedCells(), ...stored }
        : buildSeedCells();
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
    };
  },
  template: `
    <div class="bd-grid-root mns-grid">
      <div class="bd-grid-scroll">
        <table class="bd-table mns-table" role="grid">
          <thead>
            <tr class="hdr-row-letters">
              <th class="corner"></th>
              <th
                v-for="col in columns"
                :key="'mns-letter-' + col"
                class="col-letter"
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
              >
                <input
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
