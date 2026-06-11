/**
 * CDC ▸ Output for CDC — blank editable spreadsheet (A…CA, 200 rows).
 * Column B is intentionally wide; all other columns use the default width.
 */
import { ref, shallowRef, computed, onMounted, onUnmounted, nextTick } from 'vue';
import { createGridAxisHighlight } from './gridAxisHighlight.js?v=axis2';
import {
  ROW_H,
  MAX_RENDERED_ROWS,
  rowOverscan,
  createScrollRafSync,
} from './gridScroll.js?v=sp2-scroll-fix1';

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
const ROW1_FLUO_TO = colToNum('V');
const ROW1_BLUE_TAIL_FROM = colToNum('W');
/** Columns O…V collapse behind the +/- on column N (visible by default). */
const COL_OV_TOGGLE = 'N';
const COL_OV_FROM = colToNum('O');
const COL_OV_TO = colToNum('V');
const ROW5_BAND = 5;
const ROW121_BAND = 121;
/** Rows kept fully blank: never receive default text nor colour. */
const NO_FILL_ROWS = new Set([2, 3, 4, 6]);
/** Rows that keep their colour band but receive no default text (except listed labels). */
const NO_TEXT_ROWS = new Set([ROW5_BAND]);
/** Text-only labels still shown on blank / no-text rows (nothing else is written). */
const NO_FILL_TEXT_EXCEPTIONS = { '5:B': 'SPC' };
/** Column F ("Hybridization") edits via a dropdown with colour-coded choices. */
const COL_HYBRID = 'F';
const HYBRID_OPTIONS = ['', 'BEV', 'HEV', 'MHEV P2'];
const HYBRID_CLASS = {
  BEV: 'cdc-hyb-bev',
  HEV: 'cdc-hyb-hev',
  'MHEV P2': 'cdc-hyb-mhev',
};
/** Default hybridization per row (column F); the rest defaults to "MHEV P2". */
const HYBRID_DEFAULT_FROM = 7;
const HYBRID_DEFAULT_TO = 194;
const HYBRID_BEV_RANGES = [
  [7, 15], [22, 27], [34, 43], [51, 58], [63, 72], [93, 95],
  [105, 110], [124, 133], [141, 149], [157, 162], [170, 174],
  [182, 187],
];
const HYBRID_HEV_RANGES = [
  [19, 21], [31, 33], [47, 50], [61, 62], [76, 79], [102, 104],
  [114, 119], [137, 140], [153, 156], [166, 168], [178, 181],
  [191, 194],
];

function rowInRanges(row, ranges) {
  return ranges.some(([from, to]) => row >= from && row <= to);
}

function hybridDefault(row) {
  if (row < HYBRID_DEFAULT_FROM || row > HYBRID_DEFAULT_TO) return '';
  if (rowInRanges(row, HYBRID_BEV_RANGES)) return 'BEV';
  if (rowInRanges(row, HYBRID_HEV_RANGES)) return 'HEV';
  return 'MHEV P2';
}
/** Column J ("Wheels"): same value on every data row from 7, skipping 120…123. */
const COL_WHEELS = 'J';
const WHEELS_DEFAULT = '225/50 R19 (struc)';
const WHEELS_FROM = 7;
const WHEELS_SKIP_ROWS = new Set([120, 121, 122, 123]);

function wheelsDefault(row) {
  if (row < WHEELS_FROM || WHEELS_SKIP_ROWS.has(row)) return '';
  return WHEELS_DEFAULT;
}
/** Column G ("Technical Package"): "TB" on every data row from 7, skipping 120…123. */
const COL_TECH_PACKAGE = 'G';
const TECH_PACKAGE_DEFAULT = 'TB';
const TECH_PACKAGE_FROM = 7;
const TECH_PACKAGE_SKIP_ROWS = new Set([120, 121, 122, 123]);

function techPackageDefault(row) {
  if (row < TECH_PACKAGE_FROM || TECH_PACKAGE_SKIP_ROWS.has(row)) return '';
  return TECH_PACKAGE_DEFAULT;
}
/** Column K ("Pole"): same value on every data row from 7, skipping 120…123. */
const COL_POLE = 'K';
const POLE_DEFAULT = 'X_Range 80kWhU NMX';
const POLE_FROM = 7;
const POLE_SKIP_ROWS = new Set([120, 121, 122, 123]);

function poleDefault(row) {
  if (row < POLE_FROM || POLE_SKIP_ROWS.has(row)) return '';
  return POLE_DEFAULT;
}
/** Column H ("Technical Specification"): "FWD" on every data row from 7, skipping 120…123. */
const COL_TECH_SPEC = 'H';
const TECH_SPEC_DEFAULT = 'FWD';
const TECH_SPEC_FROM = 7;
const TECH_SPEC_SKIP_ROWS = new Set([120, 121, 122, 123]);
const TECH_SPEC_AWD_RANGES = [[12, 15], [43, 43], [88, 88]];

function techSpecDefault(row) {
  if (row < TECH_SPEC_FROM || TECH_SPEC_SKIP_ROWS.has(row)) return '';
  if (rowInRanges(row, TECH_SPEC_AWD_RANGES)) return 'AWD Performance';
  return TECH_SPEC_DEFAULT;
}
/** Column M ("Rear Engine"): "w/o" on every data row from 7, with overrides. */
const COL_REAR_ENGINE = 'M';
const REAR_ENGINE_DEFAULT = 'w/o';
const REAR_ENGINE_FROM = 7;
const REAR_ENGINE_ERAD4_140_RANGES = [[13, 15], [43, 43], [88, 88], [99, 101]];
const REAR_ENGINE_ERAD4_RANGES = [[133, 133], [162, 162], [187, 187]];
const REAR_ENGINE_EAWD_RANGES = [[117, 119]];

function rearEngineDefault(row) {
  if (row < REAR_ENGINE_FROM) return '';
  if (isFixedCell(row, COL_REAR_ENGINE)) return '';
  if (rowInRanges(row, REAR_ENGINE_ERAD4_140_RANGES)) return 'eRAD4 (140kW)';
  if (rowInRanges(row, REAR_ENGINE_ERAD4_RANGES)) return 'eRAD4';
  if (rowInRanges(row, REAR_ENGINE_EAWD_RANGES)) return 'eAWD';
  return REAR_ENGINE_DEFAULT;
}
/** Column tints. Yellow (#ffff99) from row 7; orange (#ffc000) on every row. */
const TINT_ROW_FROM = 7;
const YELLOW_COLS = new Set(['C', 'D', 'G', 'I', 'K', 'L', 'M', 'N']);
const ORANGE_COLS = new Set(['J', 'AC']);
/**
 * Column B is a read-only concatenation of the row's other columns, in the
 * order of the Excel formula =D&C&E&F&G&H&I&J&K&L&M&N, joined by a space.
 */
const COL_B_CONCAT = 'B';
const COL_B_CONCAT_ORDER = ['D', 'C', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
/** Column B row bands: orange (#ffc000) on rows 22…33, black on rows 34…50. */
const COL_B_ORANGE_FROM = 22;
const COL_B_ORANGE_TO = 33;
const COL_B_BLACK_FROM = 34;
const COL_B_BLACK_TO = 50;

function colBBandClass(row, col) {
  if (col !== COL_B_CONCAT) return '';
  if (row >= COL_B_ORANGE_FROM && row <= COL_B_ORANGE_TO) return 'cdc-b-orange';
  if (row >= COL_B_BLACK_FROM && row <= COL_B_BLACK_TO) return 'cdc-b-black';
  return '';
}
/** Default text written into whole columns (overridable by the user). */
const COL_DEFAULT_TEXT = { E: '5', I: 'TT', N: 'No SBW' };
/** Column D ("Design Plate"): EMEA labels per row range. */
const COL_D = 'D';
const COL_D_RANGES = [
  [7, 21, 'EMEA O3H'],
  [22, 33, 'EMEA O3W'],
  [34, 50, 'EMEA P3S'],
  [51, 62, 'EMEA P3W'],
  [63, 79, 'EMEA P3U'],
  [80, 92, 'EMEA P3H'],
  [93, 104, 'EMEA J1X'],
  [105, 119, 'EMEA J2U'],
  [124, 140, 'EMEA P3H'],
  [141, 156, 'EMEA P3W'],
  [157, 169, 'EMEA O3H'],
  [170, 181, 'EMEA O3W'],
  [182, 194, 'EMEA A3H'],
];

function colDDefault(row) {
  const match = COL_D_RANGES.find(([from, to]) => row >= from && row <= to);
  return match ? match[2] : '';
}
const COL_I_SW_FROM = 22;
const COL_I_SW_TO = 33;
/** Column L ("Front Engine"): default "S13F4.9" from row 7, blank on rows 120…123. */
const COL_L_DEFAULT = 'S13F4.9';
const COL_L_DEFAULT_FROM = 7;
const COL_L_BLANK_ROWS = new Set([120, 121, 122, 123]);
/** Column L override: "EB2LTDH2" on these row ranges instead of "S13F4.9". */
const COL_L_EB2_DEFAULT = 'EB2LTDH2';
const COL_L_EB2_RANGES = [
  [16, 21], [28, 33], [44, 50], [59, 62], [73, 79], [89, 92],
  [96, 104], [111, 119], [134, 140], [150, 156], [163, 169],
  [175, 181], [188, 194],
];
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

/**
 * Row 1 titles for columns G ("Technical Package") and H ("Technical
 * Specification") are shown as wrapping labels so each word stacks on its own
 * line inside the (taller) header row, keeping the columns narrow.
 */
const ROW1_WRAP_COLS = new Set(['G', 'H']);

function isRow1WrapCell(row, col) {
  return row === ROW1_HEADER && ROW1_WRAP_COLS.has(col);
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

function isHybridSelectCell(row, col) {
  if (NO_FILL_ROWS.has(row)) return false;
  return col === COL_HYBRID && !isFixedCell(row, col);
}

function isBConcatCell(row, col) {
  if (NO_FILL_ROWS.has(row)) return false;
  return col === COL_B_CONCAT && !isFixedCell(row, col);
}

function cellColorClass(row, col) {
  if (NO_FILL_ROWS.has(row)) return '';
  return fixedCellClass(row, col);
}

/** Column B fluo-yellow highlight on specific row ranges. */
const COL_B_FLUO_RANGES = [[7, 21]];

function columnTintClass(row, col) {
  if (NO_FILL_ROWS.has(row)) return '';
  if (isFixedCell(row, col)) return '';
  if (col === COL_B_CONCAT && rowInRanges(row, COL_B_FLUO_RANGES)) return 'cdc-tint-fluo';
  if (ORANGE_COLS.has(col)) return 'cdc-tint-orange';
  if (YELLOW_COLS.has(col) && row >= TINT_ROW_FROM) return 'cdc-tint-yellow';
  return '';
}

function columnDefaultValue(row, col) {
  if (NO_FILL_ROWS.has(row) || NO_TEXT_ROWS.has(row)) {
    return NO_FILL_TEXT_EXCEPTIONS[`${row}:${col}`] || '';
  }
  const fixed = fixedCellLabel(row, col);
  if (fixed) return fixed;
  if (col === COL_D) return colDDefault(row);
  if (col === COL_HYBRID) return isHybridSelectCell(row, col) ? hybridDefault(row) : '';
  if (col === COL_WHEELS) return wheelsDefault(row);
  if (col === COL_TECH_PACKAGE) return techPackageDefault(row);
  if (col === COL_POLE) return poleDefault(row);
  if (col === COL_TECH_SPEC) return techSpecDefault(row);
  if (col === COL_REAR_ENGINE) return rearEngineDefault(row);
  if (col === 'I' && row >= COL_I_SW_FROM && row <= COL_I_SW_TO) return 'SW';
  if (col === 'L' && row >= COL_L_DEFAULT_FROM && !COL_L_BLANK_ROWS.has(row)) {
    return rowInRanges(row, COL_L_EB2_RANGES) ? COL_L_EB2_DEFAULT : COL_L_DEFAULT;
  }
  return COL_DEFAULT_TEXT[col] || '';
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

    /** O…V visible by default; +/- on column N toggles the block. */
    const ovGroupCollapsed = ref(false);

    const visibleColumns = computed(() => {
      if (!ovGroupCollapsed.value) return allColumns.value;
      return allColumns.value.filter((col) => !isOvHiddenCol(col));
    });

    const rows = computed(() =>
      Array.from({ length: ROW_COUNT }, (_, i) => i + 1)
    );

    // ── Row virtualization ────────────────────────────────────────────────
    // Rendering all 200 rows × ~79 columns means ~15 800 live <input>/<select>
    // elements, which made horizontal/vertical scrolling stutter badly. We keep
    // only the rows inside (and just around) the viewport mounted; top/bottom
    // spacer rows preserve the full scroll height so the scrollbar stays exact.
    // The row pitch is measured live (a CDC body row is 21px tall + a 1px
    // collapsed border = 22px) so the window never drifts away from the
    // scrollbar near the bottom of the sheet.
    const scrollEl = ref(null);
    const scrollTop = ref(0);
    const viewportH = ref(900);
    const rowPitch = ref(22);

    const scrollSync = createScrollRafSync({
      scrollTop,
      getScrollEl: () => scrollEl.value,
    });

    function measureViewport() {
      const el = scrollEl.value;
      if (!el) return;
      viewportH.value = el.clientHeight || viewportH.value;
      const rs = el.querySelectorAll('tbody tr.grid-row-cv');
      if (rs.length >= 2) {
        // Measure the last two rendered rows so the (taller) row-1 header band
        // never skews the uniform body pitch used by the scroll math.
        const a = rs[rs.length - 2];
        const b = rs[rs.length - 1];
        const pitch =
          b.getBoundingClientRect().top - a.getBoundingClientRect().top;
        if (pitch > 10 && pitch < 60) rowPitch.value = pitch;
      }
    }

    const rowWindow = computed(() => {
      const pitch = rowPitch.value || ROW_H;
      const visible = Math.max(1, Math.ceil(viewportH.value / pitch));
      const overscan = rowOverscan(viewportH.value);
      const span = Math.min(MAX_RENDERED_ROWS, ROW_COUNT);
      const vpStart = Math.min(
        Math.max(0, Math.floor(scrollTop.value / pitch)),
        Math.max(0, ROW_COUNT - 1)
      );
      let start = Math.max(0, vpStart - overscan);
      let end = Math.min(ROW_COUNT, vpStart + visible + overscan);
      if (end - start > span) end = start + span;
      return { start, end };
    });

    const visibleRows = computed(() => {
      const { start, end } = rowWindow.value;
      const out = [];
      for (let r = start + 1; r <= end; r++) out.push(r);
      return out;
    });

    const topSpacerH = computed(() => rowWindow.value.start * rowPitch.value);
    const bottomSpacerH = computed(
      () => (ROW_COUNT - rowWindow.value.end) * rowPitch.value
    );
    const spacerColspan = computed(() => visibleColumns.value.length + 1);

    function onScroll(event) {
      scrollSync.onScroll(event);
    }

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
      const stored = cells.value[cellKey(row, col)];
      if (stored !== undefined) return stored;
      if (isBConcatCell(row, col)) return concatBValue(row);
      return columnDefaultValue(row, col);
    }

    function hybridCellClass(row, col) {
      return HYBRID_CLASS[cellValue(row, col)] || '';
    }

    function concatBValue(row) {
      return COL_B_CONCAT_ORDER.map((col) => String(cellValue(row, col)).trim())
        .filter((part) => part !== '')
        .join(' ');
    }

    // ── Edit on demand ────────────────────────────────────────────────────
    // Cells render as cheap static text and only the cell being edited mounts a
    // real <input>. Live <input>s are far heavier than text to paint AND to
    // mount/unmount, so keeping one at a time is what makes scrolling smooth.
    const activeKey = ref(null);

    function isEditing(row, col) {
      return activeKey.value === cellKey(row, col);
    }

    function startEdit(row, col, event) {
      onCellAxisSelect(row, col, event);
      activeKey.value = cellKey(row, col);
      nextTick(() => {
        const el = scrollEl.value
          ? scrollEl.value.querySelector('.cdc-cell-input-active')
          : null;
        if (el && typeof el.focus === 'function') {
          el.focus();
          if (typeof el.select === 'function') el.select();
        }
      });
    }

    function stopEdit() {
      activeKey.value = null;
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
      const stored = loadStoredCells();
      if (stored) cells.value = stored;
      measureViewport();
      scrollSync.flush();
      if (typeof window !== 'undefined') {
        window.addEventListener('resize', measureViewport);
      }
    });

    onUnmounted(() => {
      clearTimeout(persistTimer);
      if (typeof window !== 'undefined') {
        window.removeEventListener('resize', measureViewport);
      }
    });

    return {
      visibleColumns,
      rows,
      scrollEl,
      onScroll,
      visibleRows,
      topSpacerH,
      bottomSpacerH,
      spacerColspan,
      cellValue,
      onCellInput,
      isEditing,
      startEdit,
      stopEdit,
      onCellAxisSelect,
      onAxisRowNumClick,
      onAxisColHeaderClick,
      isAxisRow,
      isAxisCol,
      isFixedCell,
      isRow1WrapCell,
      colBBandClass,
      cellColorClass,
      fixedCellClass,
      fixedCellLabel,
      isColGroupToggle,
      ovGroupCollapsed,
      toggleOvGroup,
      isHybridSelectCell,
      hybridCellClass,
      isBConcatCell,
      concatBValue,
      columnTintClass,
      hybridOptions: HYBRID_OPTIONS,
    };
  },
  template: `
    <div class="bd-grid-root cdc-output-grid">
      <div class="bd-grid-scroll" ref="scrollEl" @scroll.passive="onScroll">
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
            <tr v-if="topSpacerH > 0" class="cdc-spacer-row" aria-hidden="true">
              <td
                :colspan="spacerColspan"
                :style="{ height: topSpacerH + 'px', padding: 0, border: 'none' }"
              ></td>
            </tr>
            <tr
              v-for="row in visibleRows"
              :key="'cdc-row-' + row"
              class="grid-row-cv"
              :class="{ 'grid-axis-row-focus': isAxisRow(row), 'cdc-row1': row === 1 }"
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
                  cellColorClass(row, col),
                  colBBandClass(row, col),
                  isHybridSelectCell(row, col) ? hybridCellClass(row, col) : columnTintClass(row, col),
                  { 'grid-axis-col-focus': isAxisCol(col) },
                ]"
              >
                <div
                  v-if="isRow1WrapCell(row, col)"
                  class="cdc-r1-wrap-label"
                >{{ fixedCellLabel(row, col) }}</div>
                <select
                  v-else-if="isHybridSelectCell(row, col)"
                  class="grid-cell-input cdc-cell-input cdc-hyb-select"
                  :value="cellValue(row, col)"
                  @mousedown="onCellAxisSelect(row, col, $event)"
                  @focus="onCellAxisSelect(row, col, $event)"
                  @change="onCellInput(row, col, $event)"
                >
                  <option v-for="opt in hybridOptions" :key="'hyb-' + opt" :value="opt">{{ opt }}</option>
                </select>
                <input
                  v-else-if="isEditing(row, col)"
                  type="text"
                  class="grid-cell-input cdc-cell-input cdc-cell-input-active"
                  autocomplete="off"
                  spellcheck="false"
                  :value="cellValue(row, col)"
                  @input="onCellInput(row, col, $event)"
                  @blur="stopEdit"
                  @keydown.enter.prevent="stopEdit"
                  @keydown.esc.prevent="stopEdit"
                />
                <span
                  v-else
                  class="cdc-cell-text"
                  @mousedown="startEdit(row, col, $event)"
                >{{ cellValue(row, col) }}</span>
              </td>
            </tr>
            <tr v-if="bottomSpacerH > 0" class="cdc-spacer-row" aria-hidden="true">
              <td
                :colspan="spacerColspan"
                :style="{ height: bottomSpacerH + 'px', padding: 0, border: 'none' }"
              ></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `,
};
