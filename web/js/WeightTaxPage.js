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

import { reactive, ref, computed } from 'vue';

/** En-têtes de colonnes Excel B → V (21 colonnes). */
function makeColumns() {
  const out = [];
  for (let i = 1; i < 22; i++) out.push(String.fromCharCode(65 + i));
  return out;
}

const ROW_COUNT = 22;
const BANNER_LABELS = {
  bev: 'BEV',
  mhev: 'MHEV P2  Weight Tax',
  hev: 'HEV Weight Tax',
};
const DEFAULT_CHART_TITLE = 'Fr Mass + CO2 Malus (€ estimation) versus Curb mass';
const DATA_ROW_START = 4;
const DATA_ROW_END = 22;
const DEFAULT_LAUNCH_YEAR = 2028;
const LAUNCH_YEAR_END = 2050;

function makeLaunchYearOptions() {
  const startYear = new Date().getFullYear();
  const options = [];
  for (let y = startYear; y <= LAUNCH_YEAR_END; y++) options.push(y);
  return options;
}

/** Libellés d'en-tête de la ligne 3, par colonne Excel. */
const ROW5_LABELS = {
  B: 'Vehicle',
  C: 'Years of launch',
  D: 'Curbweight',
  E: 'Mass in running order',
  F: 'Mass in running order after deduction',
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
  S: 'Gap for weight tax',
  T: 'Weight Tax (€)',
  U: 'CO2 Malus (€)',
  V: 'Mass + CO2 Malus (€)',
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
    const rows = Array.from({ length: ROW_COUNT }, (_, i) => i + 1);
    const launchYearOptions = makeLaunchYearOptions();
    const launchYears = reactive({});
    const chartTitles = reactive({
      bev: DEFAULT_CHART_TITLE,
      mhev: DEFAULT_CHART_TITLE,
      hev: DEFAULT_CHART_TITLE,
    });
    const chartTitle = computed({
      get: () => chartTitles[props.variant] ?? DEFAULT_CHART_TITLE,
      set: (value) => {
        chartTitles[props.variant] = value;
      },
    });
    const bannerText = computed(() => BANNER_LABELS[props.variant] || BANNER_LABELS.bev);
    const bannerLong = computed(() => props.variant !== 'bev');

    for (let r = DATA_ROW_START; r <= DATA_ROW_END; r++) {
      launchYears[r] = DEFAULT_LAUNCH_YEAR;
    }

    function row5ColorClass(col) {
      return ROW5_COLORS[col] || 'c-002060';
    }

    function row5Label(col) {
      return ROW5_LABELS[col] || '';
    }

    return { cols, rows, row5ColorClass, row5Label, launchYearOptions, launchYears, chartTitle, bannerText, bannerLong };
  },
  template: `
    <div class="weight-tax-page">
      <div class="wt-table-wrap">
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
              <!-- Ligne 3 : en-tête (haute), fond bleu foncé + dégradé I→R. -->
              <td
                v-else-if="r === 3"
                class="wt-cell wt-row5"
                :class="row5ColorClass(col)"
              >{{ row5Label(col) }}</td>
              <!-- Colonne C (Years of launch) : sélecteur d'année, défaut 2028. -->
              <td v-else-if="r >= 4 && r <= 22 && col === 'C'" class="wt-cell wt-cell-year">
                <select class="wt-year-select" v-model.number="launchYears[r]">
                  <option v-for="y in launchYearOptions" :key="y" :value="y">{{ y }}</option>
                </select>
              </td>
              <!-- Tableau gras B4:V22. -->
              <td v-else-if="r >= 4 && r <= 22" class="wt-cell"></td>
              <td v-else class="wt-empty"></td>
            </template>
          </tr>
        </tbody>
        </table>
      </div>

      <section class="wt-chart-zone" aria-label="Weight Tax chart">
        <div class="wt-chart-title-wrap">
          <input
            v-model="chartTitle"
            class="wt-chart-title-input"
            type="text"
            title="Cliquer pour modifier le titre"
            aria-label="Graph title"
            spellcheck="false"
          />
        </div>
        <div class="wt-chart-canvas"></div>
      </section>
    </div>
  `,
};
