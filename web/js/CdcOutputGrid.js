/**
 * CDC ▸ Output for CDC — blank editable spreadsheet (A…CA, 200 rows).
 * Column B is intentionally wide; all other columns use the default width.
 */
import { ref, shallowRef, computed, onMounted, onUnmounted } from 'vue';
import { createGridAxisHighlight } from './gridAxisHighlight.js?v=axis2';

const STORAGE_KEY = 'cdc-output-grid-cells-v1';
const ROW_COUNT = 200;
const MAX_COL = 'CA';

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

const COLUMN_COUNT = colToNum(MAX_COL);
const ROW1_HEADER = 1;
const ROW1_BLUE_COLS = new Set(['A', 'B']);
const ROW1_FINISH_COL = 'C';
const ROW1_FLUO_FROM = colToNum('D');
const ROW1_FLUO_TO = colToNum('N');
const ROW1_BLUE_TAIL_FROM = colToNum('W');
/** Columns O…V collapse behind the +/- on column N (hidden by default). */
const COL_OV_TOGGLE = 'N';
const COL_OV_FROM = colToNum('O');
const COL_OV_TO = colToNum('V');
const ROW5_BAND = 5;
const ROW121_BAND = 121;
const FIXED_LABELS = {
  '1:B': "'IN-OUT SYNTHESIS'!",
  '1:C': 'Trim',
  '1:D': 'Design Plate',
  '1:E': 'Seats',
  '1:F': 'Hybridization',
  '1:G': 'Technical Package',
  '1:H': 'Technical Specification',
  '1:I': 'Payload',
  '1:J': 'Wheels',
  '1:K': 'Pole',
  '1:L': 'Front Engine',
  '1:M': 'Rear Engine',
  '5:B': 'SPC',
  '121:B': 'SP2',
};

function isRow1FixedCell(row, col) {
  if (row !== ROW1_HEADER) return false;
  if (ROW1_BLUE_COLS.has(col)) return true;
  if (col === ROW1_FINISH_COL) return true;
  const n = colToNum(col);
  if (n >= ROW1_FLUO_FROM && n <= ROW1_FLUO_TO) return true;
  return n >= ROW1_BLUE_TAIL_FROM && n <= COLUMN_COUNT;
}

function isOvHiddenCol(col) {
  const n = colToNum(col);
  return n >= COL_OV_FROM && n <= COL_OV_TO;
}

function isFullRowBandCell(row, col) {
  return (
    (row === ROW5_BAND || row === ROW121_BAND) &&
    colToNum(col) >= 1 &&
    colToNum(col) <= COLUMN_COUNT
  );
}

function isFixedCell(row, col) {
  return isRow1FixedCell(row, col) || isFullRowBandCell(row, col);
}

function fixedCellClass(row, col) {
  if (isRow1FixedCell(row, col)) {
    if (ROW1_BLUE_COLS.has(col)) return 'cdc-r1-blue';
    if (col === ROW1_FINISH_COL) return 'cdc-r1-finish';
    const n = colToNum(col);
    if (n >= ROW1_FLUO_FROM && n <= ROW1_FLUO_TO) return 'cdc-r1-fluo';
    return 'cdc-r1-blue';
  }
  if (row === ROW5_BAND && isFullRowBandCell(row, col)) return 'cdc-r5-green';
  if (row === ROW121_BAND && isFullRowBandCell(row, col)) return 'cdc-r121-black';
  return '';
}

function fixedCellLabel(row, col) {
  return FIXED_LABELS[`${row}:${col}`] || '';
}

function cellKey(row, col) {
  return `${row}:${col}`;
}

function loadStoredCells() {
  if (typeof localStorage === 'undefined') return null;
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch {
    return null;
  }
}

export default {
  name: 'CdcOutputGrid',
  setup() {
    const cells = shallowRef(loadStoredCells() || {});
    const editTick = ref(0);
    let persistTimer = 0;

    const {
      syncFromCell: syncAxisFromCell,
      onRowNumClick: onAxisRowNumClick,
      onColHeaderClick: onAxisColHeaderClick,
      isAxisRow,
      isAxisCol,
    } = createGridAxisHighlight();

    const allColumns = computed(() =>
      Array.from({ length: COLUMN_COUNT }, (_, i) => numToCol(i + 1))
    );

    /** O…V hidden by default; +/- on column N toggles the block. */
    const ovGroupCollapsed = ref(true);

    const visibleColumns = computed(() => {
      if (!ovGroupCollapsed.value) return allColumns.value;
      return allColumns.value.filter((col) => !isOvHiddenCol(col));
    });

    const rows = computed(() =>
      Array.from({ length: ROW_COUNT }, (_, i) => i + 1)
    );

    function isColGroupToggle(col) {
      return col === COL_OV_TOGGLE;
    }

    function toggleOvGroup() {
      ovGroupCollapsed.value = !ovGroupCollapsed.value;
    }

    function onCellAxisSelect(row, col, event) {
      if (event && event.button != null && event.button !== 0) return;
      syncAxisFromCell(row, col);
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
        localStorage.setItem(STORAGE_KEY, JSON.stringify(cells.value));
      } catch {
        /* storage full / unavailable — keep working in-memory */
      }
    }

    function cellValue(row, col) {
      void editTick.value;
      return cells.value[cellKey(row, col)] || '';
    }

    function onCellInput(row, col, event) {
      if (isFixedCell(row, col)) return;
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
      const stored = loadStoredCells();
      if (stored) cells.value = stored;
    });

    onUnmounted(() => {
      clearTimeout(persistTimer);
    });

    return {
      visibleColumns,
      rows,
      cellValue,
      onCellInput,
      onCellAxisSelect,
      onAxisRowNumClick,
      onAxisColHeaderClick,
      isAxisRow,
      isAxisCol,
      isFixedCell,
      fixedCellClass,
      fixedCellLabel,
      isColGroupToggle,
      ovGroupCollapsed,
      toggleOvGroup,
    };
  },
  template: `
    <div class="bd-grid-root cdc-output-grid">
      <div class="bd-grid-scroll">
        <table class="bd-table cdc-output-table" role="grid">
          <thead>
            <tr class="hdr-row-toggle">
              <th class="corner"></th>
              <th
                v-for="col in visibleColumns"
                :key="'cdc-toggle-' + col"
                class="cdc-col-toggle-cell"
                :data-col="col"
              >
                <button
                  v-if="isColGroupToggle(col)"
                  type="button"
                  class="cdc-col-toggle-btn"
                  :title="(ovGroupCollapsed ? 'Déplier' : 'Plier') + ' colonnes O…V'"
                  @click="toggleOvGroup"
                >{{ ovGroupCollapsed ? '+' : '\u2212' }}</button>
              </th>
            </tr>
            <tr class="hdr-row-band">
              <th class="corner"></th>
              <th
                v-for="col in visibleColumns"
                :key="'cdc-band-' + col"
                class="col-letter cdc-col-band-hdr"
                :data-col="col"
              ></th>
            </tr>
            <tr class="hdr-row-letters">
              <th class="corner"></th>
              <th
                v-for="col in visibleColumns"
                :key="'cdc-letter-' + col"
                class="col-letter grid-axis-hdr"
                :class="{ 'grid-axis-hdr-focus': isAxisCol(col) }"
                :data-col="col"
                @mousedown="onAxisColHeaderClick(col, $event)"
              >{{ col }}</th>
            </tr>
          </thead>
          <tbody>
            <tr
              v-for="row in rows"
              :key="'cdc-row-' + row"
              class="grid-row-cv"
              :class="{ 'grid-axis-row-focus': isAxisRow(row) }"
            >
              <td
                class="row-num grid-axis-hdr"
                :class="{ 'grid-axis-hdr-focus': isAxisRow(row) }"
                @mousedown="onAxisRowNumClick(row, $event)"
              >{{ row }}</td>
              <td
                v-for="col in visibleColumns"
                :key="'cdc-' + row + '-' + col"
                class="data-cell"
                :data-col="col"
                :class="[
                  fixedCellClass(row, col),
                  { 'grid-axis-col-focus': isAxisCol(col) },
                ]"
              >
                <span
                  v-if="isFixedCell(row, col)"
                  class="cdc-fixed-label"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                >{{ fixedCellLabel(row, col) }}</span>
                <input
                  v-else
                  type="text"
                  class="grid-cell-input cdc-cell-input"
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
