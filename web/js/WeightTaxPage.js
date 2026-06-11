/**
 * Weight Tax — première planche (mise en page type tableur).
 *
 * Repères Excel reproduits :
 *  - lignes 1 et 2 : bandeau bleu plein (#0070c0) couvrant B1:V2, cellules fusionnées
 *    (rowspan/colspan) donc aucune ligne / colonne ne le traverse,
 *  - au centre du bandeau : le texte « BEV » en vert,
 *  - lignes 3 → 22, colonnes B → V : tableau au contour gras, lignes et colonnes
 *    internes en gras.
 *
 * Page volontairement statique pour l'instant : c'est le squelette visuel sur
 * lequel les données et calculs (côté backend) viendront se greffer ensuite.
 */

import { reactive, ref, computed, onMounted, onUnmounted, watch, nextTick } from 'vue';
import {
  listCdcVehiclesForWeightTaxVariant,
  loadCdcOutputCells,
  CDC_OUTPUT_STORAGE_KEY,
} from './CdcOutputGrid.js?v=cdc-wt-link1';

/** En-têtes de colonnes Excel B → V (21 colonnes). */
function makeColumns() {
  const out = [];
  for (let i = 1; i < 22; i++) out.push(String.fromCharCode(65 + i));
  return out;
}

const HEADER_ROW = 3;
const DATA_ROW_START = 4;
const BANNER_LABELS = {
  bev: 'BEV',
  mhev: 'MHEV P2  Weight Tax',
  hev: 'HEV Weight Tax',
};
const DEFAULT_CHART_TITLE = 'Fr Mass + CO2 Malus (€ estimation) versus Curb mass';
const DATA_COLS = makeColumns();
const DEFAULT_LAUNCH_YEAR = 2028;
const LAUNCH_YEAR_END = 2050;

const BEV_CELL_DEFAULTS = {
  G: '1399 kg',
  H: '600 kg',
};

function createVariantHeaderLabels() {
  const labels = {};
  for (const col of DATA_COLS) {
    labels[col] = ROW5_LABELS[col] || '';
  }
  return labels;
}

function createEmptyRowCols(defaultsByCol = {}) {
  const row = {};
  for (const col of DATA_COLS) {
    row[col] = defaultsByCol[col] ?? '';
  }
  return row;
}

function makeLaunchYearOptions() {
  const startYear = new Date().getFullYear();
  const options = [];
  for (let y = startYear; y <= LAUNCH_YEAR_END; y++) options.push(y);
  return options;
}

/** Libellés d'en-tête de la ligne 3, par colonne Excel. */
const ROW5_LABELS = {
  B: 'Vehicle',
  C: 'Years of\nlaunch',
  D: 'Curbweight',
  E: 'Mass in\nrunning order',
  F: 'Mass in running order\nafter deduction',
  G: 'Treeshold',
  H: 'Deduction',
  I: '200 kg',
  J: '10 €/kg',
  K: '100 kg',
  L: '15€/kg',
  M: '100 kg',
  N: '20 €/kg',
  O: '100 kg',
  P: '25 €/kg',
  Q: '100 kg',
  R: '30€/kg',
  S: 'Gap for\nweight tax',
  T: 'Weight Tax\n(€)',
  U: 'CO2 Malus\n(€)',
  V: 'Mass + CO2\nMalus (€)',
};

/**
 * Couleur de fond de la ligne 3 par colonne (dégradé de bleus sur I→R).
 * Les autres colonnes du tableau (B→H, S→V) restent en bleu foncé #002060.
 */
const ROW5_COLORS = {
  I: 'c-daeef3',
  J: 'c-daeef3',
  K: 'c-b7dee8',
  L: 'c-b7dee8',
  M: 'c-92cddc',
  N: 'c-92cddc',
  O: 'c-31869b',
  P: 'c-31869b',
  Q: 'c-215967',
  R: 'c-215967',
};

const STORAGE_KEY = 'weight-tax-state-v1';
let persistTimer = 0;

function loadPersistedState() {
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

function hydrateVehiclesForVariant(variant, savedList) {
  if (!Array.isArray(savedList) || savedList.length === 0) return [];
  const catalog = listCdcVehiclesForWeightTaxVariant(variant, loadCdcOutputCells());
  const byCdcRow = new Map(catalog.map((v) => [v.cdcRow, v]));
  return savedList
    .map((item) => {
      const cdcRow = typeof item === 'number' ? item : item?.cdcRow;
      if (typeof cdcRow !== 'number') return null;
      const fresh = byCdcRow.get(cdcRow);
      if (fresh) return fresh;
      if (typeof item === 'object' && item) return item;
      return { cdcRow, trim: `Row ${cdcRow}`, hybrid: '', colBClasses: [] };
    })
    .filter(Boolean);
}

function createHeaderLabelsFromSaved(savedVariant) {
  const labels = createVariantHeaderLabels();
  if (savedVariant && typeof savedVariant === 'object') {
    for (const col of DATA_COLS) {
      if (savedVariant[col] != null) labels[col] = savedVariant[col];
    }
  }
  return labels;
}

export default {
  name: 'WeightTaxPage',
  props: {
    variant: {
      type: String,
      default: 'bev',
      validator: (v) => ['bev', 'mhev', 'hev'].includes(v),
    },
  },
  setup(props) {
    const cols = makeColumns();
    const launchYearOptions = makeLaunchYearOptions();
    const saved = loadPersistedState();

    const selectedVehicles = reactive({
      bev: hydrateVehiclesForVariant('bev', saved?.selectedVehicles?.bev ?? []),
      mhev: hydrateVehiclesForVariant('mhev', saved?.selectedVehicles?.mhev ?? []),
      hev: hydrateVehiclesForVariant('hev', saved?.selectedVehicles?.hev ?? []),
    });
    const cellData = reactive({
      bev: { ...(saved?.cellData?.bev ?? {}) },
      mhev: { ...(saved?.cellData?.mhev ?? {}) },
      hev: { ...(saved?.cellData?.hev ?? {}) },
    });
    const headerLabels = reactive({
      bev: createHeaderLabelsFromSaved(saved?.headerLabels?.bev),
      mhev: createHeaderLabelsFromSaved(saved?.headerLabels?.mhev),
      hev: createHeaderLabelsFromSaved(saved?.headerLabels?.hev),
    });
    const launchYears = reactive({
      bev: { ...(saved?.launchYears?.bev ?? {}) },
      mhev: { ...(saved?.launchYears?.mhev ?? {}) },
      hev: { ...(saved?.launchYears?.hev ?? {}) },
    });
    const chartTitles = reactive({
      bev: saved?.chartTitles?.bev ?? DEFAULT_CHART_TITLE,
      mhev: saved?.chartTitles?.mhev ?? DEFAULT_CHART_TITLE,
      hev: saved?.chartTitles?.hev ?? DEFAULT_CHART_TITLE,
    });
    const chartTitle = computed({
      get: () => chartTitles[props.variant] ?? DEFAULT_CHART_TITLE,
      set: (value) => {
        chartTitles[props.variant] = value;
      },
    });
    const bannerText = computed(() => BANNER_LABELS[props.variant] || BANNER_LABELS.bev);
    const bannerLong = computed(() => props.variant !== 'bev');

    const dataRows = computed(() => selectedVehicles[props.variant] || []);

    const rows = computed(() => {
      const count = HEADER_ROW + dataRows.value.length;
      return Array.from({ length: Math.max(count, HEADER_ROW) }, (_, i) => i + 1);
    });

    function vehicleAtRow(r) {
      if (r < DATA_ROW_START) return null;
      return dataRows.value[r - DATA_ROW_START] ?? null;
    }

    function vehicleRowKey(cdcRow) {
      return String(cdcRow);
    }

    function ensureVehicleData(variant, cdcRow) {
      const key = vehicleRowKey(cdcRow);
      if (!cellData[variant][key]) {
        const defaults = variant === 'bev' ? BEV_CELL_DEFAULTS : {};
        cellData[variant][key] = createEmptyRowCols(defaults);
      }
      if (launchYears[variant][key] == null) {
        launchYears[variant][key] = DEFAULT_LAUNCH_YEAR;
      }
    }

    for (const variant of ['bev', 'mhev', 'hev']) {
      for (const vehicle of selectedVehicles[variant]) {
        ensureVehicleData(variant, vehicle.cdcRow);
      }
    }

    function persistState() {
      if (typeof localStorage === 'undefined') return;
      try {
        localStorage.setItem(
          STORAGE_KEY,
          JSON.stringify({
            selectedVehicles: {
              bev: selectedVehicles.bev.map((v) => v.cdcRow),
              mhev: selectedVehicles.mhev.map((v) => v.cdcRow),
              hev: selectedVehicles.hev.map((v) => v.cdcRow),
            },
            cellData: {
              bev: cellData.bev,
              mhev: cellData.mhev,
              hev: cellData.hev,
            },
            launchYears: {
              bev: launchYears.bev,
              mhev: launchYears.mhev,
              hev: launchYears.hev,
            },
            headerLabels: {
              bev: headerLabels.bev,
              mhev: headerLabels.mhev,
              hev: headerLabels.hev,
            },
            chartTitles: {
              bev: chartTitles.bev,
              mhev: chartTitles.mhev,
              hev: chartTitles.hev,
            },
          })
        );
      } catch {
        /* quota / private mode */
      }
    }

    function schedulePersist() {
      clearTimeout(persistTimer);
      persistTimer = setTimeout(persistState, 300);
    }

    function rehydrateSelectedVehiclesFromCdc() {
      for (const variant of ['bev', 'mhev', 'hev']) {
        const rows = selectedVehicles[variant].map((v) => v.cdcRow);
        selectedVehicles[variant] = hydrateVehiclesForVariant(variant, rows);
      }
    }

    const vehiclePickerOpen = ref(false);
    const vehiclePickerOptions = ref([]);
    const vehiclePickerDraft = ref([]);

    function refreshVehiclePickerOptions() {
      vehiclePickerOptions.value = listCdcVehiclesForWeightTaxVariant(
        props.variant,
        loadCdcOutputCells()
      );
    }

    function openVehiclePicker() {
      headerEditingCol.value = null;
      refreshVehiclePickerOptions();
      vehiclePickerDraft.value = dataRows.value.map((v) => v.cdcRow);
      vehiclePickerOpen.value = true;
    }

    function closeVehiclePicker() {
      vehiclePickerOpen.value = false;
    }

    function isDraftVehicleSelected(cdcRow) {
      return vehiclePickerDraft.value.includes(cdcRow);
    }

    function toggleVehicleDraft(cdcRow) {
      const draft = vehiclePickerDraft.value.slice();
      const idx = draft.indexOf(cdcRow);
      if (idx >= 0) draft.splice(idx, 1);
      else draft.push(cdcRow);
      vehiclePickerDraft.value = draft;
    }

    function applyVehiclePicker() {
      const picked = vehiclePickerOptions.value.filter((v) =>
        vehiclePickerDraft.value.includes(v.cdcRow)
      );
      selectedVehicles[props.variant] = picked;
      for (const vehicle of picked) ensureVehicleData(props.variant, vehicle.cdcRow);
      vehiclePickerOpen.value = false;
      nextTick(syncChartWidth);
    }

    function onHeaderCellClick(col) {
      if (col === 'B') {
        openVehiclePicker();
        return;
      }
      startHeaderEdit(col);
    }

    function vehicleColBClasses(vehicle) {
      if (!vehicle?.colBClasses?.length) return [];
      return vehicle.colBClasses;
    }

    function row5ColorClass(col) {
      return ROW5_COLORS[col] || 'c-002060';
    }

    const headerEditingCol = ref(null);
    const headerInputRef = ref(null);

    function isHeaderEditing(col) {
      return headerEditingCol.value === col;
    }

    function startHeaderEdit(col) {
      if (col === 'B') {
        openVehiclePicker();
        return;
      }
      headerEditingCol.value = col;
      nextTick(() => {
        const el = headerInputRef.value;
        const input = Array.isArray(el)
          ? el.find((node) => node && typeof node.focus === 'function')
          : el;
        if (!input || typeof input.focus !== 'function') return;
        input.focus();
        if (typeof input.select === 'function') input.select();
      });
    }

    function stopHeaderEdit() {
      headerEditingCol.value = null;
    }

    function onHeaderKeydown(event) {
      if (event.key === 'Escape') {
        event.preventDefault();
        stopHeaderEdit();
      }
    }

    const tableWrapRef = ref(null);
    const chartZoneRef = ref(null);
    const chartCanvasRef = ref(null);
    const chartTitleInputRef = ref(null);
    const chartTitleEditing = ref(false);
    const chartTitleInputWidth = computed(() => `${Math.max(chartTitle.value.length + 2, 20)}ch`);
    let tableWidthObserver = null;

    function syncChartWidth() {
      const table = tableWrapRef.value?.querySelector('.wt-sheet');
      const chart = chartZoneRef.value;
      if (!table || !chart) return;
      chart.style.width = `${table.getBoundingClientRect().width}px`;
    }

    onMounted(() => {
      syncChartWidth();
      const table = tableWrapRef.value?.querySelector('.wt-sheet');
      if (table && typeof ResizeObserver !== 'undefined') {
        tableWidthObserver = new ResizeObserver(syncChartWidth);
        tableWidthObserver.observe(table);
      }
      window.addEventListener('resize', syncChartWidth);
      window.addEventListener('storage', onCdcStorageChange);
    });

    onUnmounted(() => {
      clearTimeout(persistTimer);
      tableWidthObserver?.disconnect();
      window.removeEventListener('resize', syncChartWidth);
      window.removeEventListener('storage', onCdcStorageChange);
    });

    function onCdcStorageChange(event) {
      if (event.key !== CDC_OUTPUT_STORAGE_KEY) return;
      rehydrateSelectedVehiclesFromCdc();
      if (vehiclePickerOpen.value) refreshVehiclePickerOptions();
    }

    watch(
      () => ({
        selectedVehicles: {
          bev: selectedVehicles.bev.map((v) => v.cdcRow),
          mhev: selectedVehicles.mhev.map((v) => v.cdcRow),
          hev: selectedVehicles.hev.map((v) => v.cdcRow),
        },
        cellData,
        launchYears,
        headerLabels,
        chartTitles: { ...chartTitles },
      }),
      schedulePersist,
      { deep: true }
    );

    watch(() => props.variant, () => {
      headerEditingCol.value = null;
      vehiclePickerOpen.value = false;
      requestAnimationFrame(syncChartWidth);
    });

    function startChartTitleEdit() {
      chartTitleEditing.value = true;
      nextTick(() => {
        const input = chartTitleInputRef.value;
        if (!input) return;
        input.focus();
        input.select();
      });
    }

    function stopChartTitleEdit() {
      chartTitleEditing.value = false;
    }

    function onChartTitleKeydown(event) {
      if (event.key === 'Enter') {
        event.preventDefault();
        stopChartTitleEdit();
      }
      if (event.key === 'Escape') {
        event.preventDefault();
        stopChartTitleEdit();
      }
    }

    function exportChartPng() {
      const zone = chartZoneRef.value;
      const canvasEl = chartCanvasRef.value;
      if (!zone || !canvasEl) return;

      const headerH = 44;
      const width = Math.max(Math.round(zone.getBoundingClientRect().width), 1);
      const chartH = Math.max(Math.round(canvasEl.getBoundingClientRect().height), 1);
      const height = headerH + chartH;

      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      ctx.fillStyle = '#f3f3f3';
      ctx.fillRect(0, 0, width, headerH);
      ctx.strokeStyle = '#d9d9d9';
      ctx.beginPath();
      ctx.moveTo(0, headerH);
      ctx.lineTo(width, headerH);
      ctx.stroke();

      ctx.fillStyle = '#000000';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.font = '700 14px "Segoe UI", Arial, sans-serif';
      ctx.fillText(chartTitle.value || DEFAULT_CHART_TITLE, width / 2, headerH / 2);

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, headerH, width, chartH);

      ctx.strokeStyle = '#000000';
      ctx.lineWidth = 2;
      ctx.strokeRect(1, 1, width - 2, height - 2);

      const titlePart = (chartTitle.value || DEFAULT_CHART_TITLE)
        .replace(/\s+/g, '-')
        .replace(/[^\w-]+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '');
      const filename = `weight-tax-${props.variant}${titlePart ? `-${titlePart}` : ''}.png`;

      const link = document.createElement('a');
      link.href = canvas.toDataURL('image/png');
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
    }

    return {
      cols,
      rows,
      dataRows,
      vehicleAtRow,
      vehicleRowKey,
      row5ColorClass,
      headerLabels,
      isHeaderEditing,
      onHeaderCellClick,
      startHeaderEdit,
      stopHeaderEdit,
      onHeaderKeydown,
      headerInputRef,
      vehiclePickerOpen,
      vehiclePickerOptions,
      vehiclePickerDraft,
      isDraftVehicleSelected,
      toggleVehicleDraft,
      applyVehiclePicker,
      closeVehiclePicker,
      vehicleColBClasses,
      launchYearOptions,
      launchYears,
      cellData,
      chartTitle,
      bannerText,
      bannerLong,
      tableWrapRef,
      chartZoneRef,
      chartCanvasRef,
      chartTitleInputRef,
      chartTitleEditing,
      chartTitleInputWidth,
      startChartTitleEdit,
      stopChartTitleEdit,
      onChartTitleKeydown,
      exportChartPng,
    };
  },
  template: `
    <div class="weight-tax-page">
      <div class="wt-sheet-column">
        <div class="wt-table-wrap" ref="tableWrapRef">
          <table class="wt-sheet">
        <thead>
          <tr>
            <th class="wt-corner"></th>
            <th
              v-for="col in cols"
              :key="col"
              class="wt-colhead"
              :class="{ 'wt-col-b': col === 'B', 'wt-col-c': col === 'C', 'wt-col-ch': ['D','E','F'].includes(col), 'wt-col-gh': ['G','H'].includes(col) }"
            >{{ col }}</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="r in rows" :key="r">
            <th class="wt-rownum">{{ r }}</th>
            <template v-for="(col, ci) in cols" :key="col">
              <!-- Bandeau bleu : ancré en B1, fusionné sur B1:V2 (2 lignes × 21 colonnes). -->
              <td
                v-if="r === 1 && ci === 0"
                class="wt-banner"
                rowspan="2"
                colspan="21"
              >
                <span class="wt-bev" :class="{ 'wt-banner-text-long': bannerLong }">{{ bannerText }}</span>
              </td>
              <!-- Cellules B..V des lignes 1 et 2 : couvertes par le bandeau fusionné. -->
              <template v-else-if="r === 1 || r === 2"></template>
              <!-- Ligne 3 : en-têtes centrés ; colonne B ouvre le sélecteur véhicules. -->
              <td
                v-else-if="r === 3"
                class="wt-cell wt-row5"
                :class="row5ColorClass(col)"
              >
                <div class="wt-row5-label-wrap">
                  <span
                    v-if="!isHeaderEditing(col)"
                    class="wt-row5-label"
                    :class="{ 'wt-row5-label-pick': col === 'B' }"
                    :title="col === 'B' ? 'Cliquer pour choisir les véhicules (Output for CDC)' : 'Cliquer pour modifier'"
                    @click.stop="onHeaderCellClick(col)"
                  >{{ headerLabels[variant][col] }}</span>
                  <textarea
                    v-else-if="isHeaderEditing(col)"
                    :ref="(el) => { if (isHeaderEditing(col)) headerInputRef = el; }"
                    v-model="headerLabels[variant][col]"
                    class="wt-row5-input"
                    spellcheck="false"
                    @blur="stopHeaderEdit"
                    @keydown="onHeaderKeydown"
                  ></textarea>
                </div>
              </td>
              <!-- Colonne B : Trim CDC (couleur colonne B Output for CDC). -->
              <td
                v-else-if="vehicleAtRow(r) && col === 'B'"
                class="wt-cell wt-cell-vehicle"
                :class="vehicleColBClasses(vehicleAtRow(r))"
              >
                <span class="wt-vehicle-label">{{ vehicleAtRow(r).trim }}</span>
              </td>
              <!-- Colonne C (Years of launch) : sélecteur d'année, défaut 2028. -->
              <td v-else-if="vehicleAtRow(r) && col === 'C'" class="wt-cell wt-cell-year">
                <select
                  class="wt-year-select"
                  v-model.number="launchYears[variant][vehicleRowKey(vehicleAtRow(r).cdcRow)]"
                >
                  <option v-for="y in launchYearOptions" :key="y" :value="y">{{ y }}</option>
                </select>
              </td>
              <!-- Lignes véhicules — cellules éditables. -->
              <td v-else-if="vehicleAtRow(r)" class="wt-cell wt-cell-editable">
                <input
                  v-model="cellData[variant][vehicleRowKey(vehicleAtRow(r).cdcRow)][col]"
                  class="wt-cell-input"
                  type="text"
                  spellcheck="false"
                />
              </td>
              <td v-else class="wt-empty"></td>
            </template>
          </tr>
        </tbody>
        </table>
        </div>

        <section class="wt-chart-zone" ref="chartZoneRef" aria-label="Weight Tax chart">
          <div class="wt-chart-title-wrap">
            <button
              type="button"
              class="wt-chart-export-btn"
              title="Exporter le graphe en PNG"
              aria-label="Exporter le graphe en PNG"
              @click="exportChartPng"
            >
              <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true">
                <path
                  fill="none"
                  stroke="currentColor"
                  stroke-width="1.35"
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  d="M8 2v7.5M5.5 7 8 9.5 10.5 7M3 12.5h10"
                />
              </svg>
            </button>
            <div class="wt-chart-title-center">
              <span
                v-if="!chartTitleEditing"
                class="wt-chart-title-text"
                title="Cliquer pour modifier le titre"
                @click="startChartTitleEdit"
              >{{ chartTitle }}</span>
              <input
                v-else
                ref="chartTitleInputRef"
                v-model="chartTitle"
                class="wt-chart-title-input"
                type="text"
                aria-label="Graph title"
                spellcheck="false"
                :style="{ width: chartTitleInputWidth }"
                @blur="stopChartTitleEdit"
                @keydown="onChartTitleKeydown"
              />
            </div>
          </div>
          <div class="wt-chart-canvas" ref="chartCanvasRef"></div>
        </section>
      </div>

      <div
        v-if="vehiclePickerOpen"
        class="wt-modal-backdrop"
        @click.self="closeVehiclePicker"
      >
        <div class="wt-modal" role="dialog" aria-labelledby="wt-vehicle-picker-title">
          <h3 id="wt-vehicle-picker-title" class="wt-modal-title">Sélection des véhicules</h3>
          <p class="wt-modal-hint">
            Source : Output for CDC (Trim), filtré par hybridation
            <strong>{{ variant === 'mhev' ? 'MHEV P2' : variant.toUpperCase() }}</strong>.
          </p>
          <div v-if="vehiclePickerOptions.length" class="wt-vehicle-list">
            <label
              v-for="opt in vehiclePickerOptions"
              :key="opt.cdcRow"
              class="wt-vehicle-option"
            >
              <input
                type="checkbox"
                :checked="isDraftVehicleSelected(opt.cdcRow)"
                @change="toggleVehicleDraft(opt.cdcRow)"
              />
              <span class="wt-vehicle-option-swatch" :class="opt.colBClasses"></span>
              <span class="wt-vehicle-option-label">{{ opt.trim }}</span>
            </label>
          </div>
          <p v-else class="wt-modal-empty">Aucun véhicule trouvé pour cette hybridation dans Output for CDC.</p>
          <div class="wt-modal-actions">
            <button type="button" class="wt-modal-btn" @click="closeVehiclePicker">Annuler</button>
            <button type="button" class="wt-modal-btn wt-modal-btn-primary" @click="applyVehiclePicker">
              Appliquer
            </button>
          </div>
        </div>
      </div>
    </div>
  `,
};
