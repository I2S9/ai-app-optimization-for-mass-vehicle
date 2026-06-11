/**
 * Waterline — page de visualisation type "planche" :
 * fond noir + grands rectangles blancs disposés en 2x2.
 * Chaque rectangle contient :
 *  - un titre éditable (noir, majuscules par défaut) en haut à gauche,
 *  - un libellé "WATERLINE <MOIS> <ANNEE>" en haut à droite (mois/année via menus déroulants),
 *  - une zone image où l'on peut importer depuis le PC, glisser-déposer ou coller (Ctrl+V).
 *
 * L'état (titre, mois, année, image) est conservé localement (localStorage) afin de
 * rester léger côté client — aucune donnée n'est chargée depuis la base pour cette page.
 */
import { ref, computed, onMounted, onUnmounted, watch } from 'vue';

const STORAGE_KEY = 'waterline-panels-v1';
/** Disposition : 2 lignes × 20 colonnes (chaque colonne = 1 image en haut + 1 en bas). */
const PANEL_ROWS = 2;
const PANEL_COLS = 20;
const PANEL_COUNT = PANEL_ROWS * PANEL_COLS;

/** Mois en abrégé FR (majuscules) — index 0 = janvier. "WATERLINE FEV 2026" => février. */
const MONTHS = [
  'JAN', 'FEV', 'MAR', 'AVR', 'MAI', 'JUN',
  'JUL', 'AOU', 'SEP', 'OCT', 'NOV', 'DEC',
];

function yearChoices() {
  const now = new Date().getFullYear();
  const out = [];
  for (let y = now - 3; y <= now + 4; y++) out.push(y);
  return out;
}

function makeDefaultPanels() {
  const now = new Date();
  const m = now.getMonth();
  const y = now.getFullYear();
  return Array.from({ length: PANEL_COUNT }, () => ({
    title: '',
    month: m,
    year: y,
    image: null,
  }));
}

function loadPanels() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return makeDefaultPanels();
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return makeDefaultPanels();
    const base = makeDefaultPanels();
    for (let i = 0; i < PANEL_COUNT; i++) {
      if (parsed[i] && typeof parsed[i] === 'object') {
        base[i] = {
          title: typeof parsed[i].title === 'string' ? parsed[i].title : '',
          month: Number.isInteger(parsed[i].month) ? parsed[i].month : base[i].month,
          year: Number.isInteger(parsed[i].year) ? parsed[i].year : base[i].year,
          image: typeof parsed[i].image === 'string' ? parsed[i].image : null,
        };
      }
    }
    return base;
  } catch {
    return makeDefaultPanels();
  }
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error || new Error('read failed'));
    reader.readAsDataURL(file);
  });
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('image load failed'));
    img.src = src;
  });
}

export default {
  name: 'WaterlinePage',
  setup() {
    const panels = ref(loadPanels());
    const years = yearChoices();
    /** Dernier panneau cliqué : cible du collage clavier (Ctrl+V) global. */
    const activePanel = ref(0);
    const dragOver = ref(-1);
    /** Refs des <input type=file> cachés, un par panneau. */
    const fileInputs = ref([]);

    let saveTimer = 0;
    function persist() {
      if (saveTimer) clearTimeout(saveTimer);
      saveTimer = setTimeout(() => {
        saveTimer = 0;
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(panels.value));
        } catch (e) {
          // Quota dépassé (images volumineuses) : on n'interrompt pas l'utilisateur.
          console.warn('[waterline] sauvegarde locale impossible:', e);
        }
      }, 300);
    }

    watch(panels, persist, { deep: true });

    function labelFor(panel) {
      return `WATERLINE ${MONTHS[panel.month] || ''} ${panel.year}`.trim();
    }

    function setTitle(i, value) {
      panels.value[i].title = String(value || '').toUpperCase();
    }

    async function setImageFromFile(i, file) {
      if (!file || !/^image\//.test(file.type)) return;
      try {
        panels.value[i].image = await fileToDataUrl(file);
      } catch (e) {
        console.warn('[waterline] import image échoué:', e);
      }
    }

    function openFileDialog(i) {
      activePanel.value = i;
      const input = fileInputs.value[i];
      if (input) input.click();
    }

    function onFileChange(i, event) {
      const file = event.target.files && event.target.files[0];
      if (file) void setImageFromFile(i, file);
      event.target.value = '';
    }

    function onDrop(i, event) {
      dragOver.value = -1;
      activePanel.value = i;
      const dt = event.dataTransfer;
      if (!dt) return;
      const file = dt.files && dt.files[0];
      if (file) void setImageFromFile(i, file);
    }

    function onDragOver(i) {
      dragOver.value = i;
    }

    function onDragLeave(i) {
      if (dragOver.value === i) dragOver.value = -1;
    }

    function clearImage(i) {
      panels.value[i].image = null;
    }

    /**
     * Exporte un panneau en image PNG : entête blanc (date WATERLINE à gauche,
     * titre centré) puis l'image en dessous. Tout est dessiné sur un canvas afin
     * de ne dépendre d'aucune librairie externe.
     */
    async function exportPanel(i) {
      const panel = panels.value[i];
      const dateText = labelFor(panel);
      const titleText = String(panel.title || '').toUpperCase();

      let img = null;
      if (panel.image) {
        try {
          img = await loadImage(panel.image);
        } catch (e) {
          console.warn('[waterline] export: image illisible:', e);
        }
      }

      const baseWidth = Math.max(img ? img.naturalWidth : 1000, 900);
      const imgH = img
        ? Math.round(img.naturalHeight * (baseWidth / img.naturalWidth))
        : 0;
      const headerH = 70;
      const width = baseWidth;
      const height = headerH + imgH;

      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const ctx = canvas.getContext('2d');
      if (!ctx) return;

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, width, height);

      const pad = 24;
      ctx.textBaseline = 'middle';

      ctx.fillStyle = '#000000';
      ctx.textAlign = 'left';
      ctx.font = '700 22px "Segoe UI", Arial, sans-serif';
      ctx.fillText(dateText, pad, headerH / 2);

      if (titleText) {
        ctx.textAlign = 'center';
        ctx.font = '800 28px "Segoe UI", Arial, sans-serif';
        ctx.fillText(titleText, width / 2, headerH / 2);
      }

      if (img) {
        ctx.drawImage(img, 0, headerH, width, imgH);
      }

      ctx.strokeStyle = '#000000';
      ctx.lineWidth = 3;
      ctx.strokeRect(1.5, 1.5, width - 3, height - 3);
      ctx.beginPath();
      ctx.moveTo(0, headerH);
      ctx.lineTo(width, headerH);
      ctx.lineWidth = 1;
      ctx.strokeStyle = '#e2e2e2';
      ctx.stroke();

      const safe = (titleText || dateText).replace(/[^\w-]+/g, '_').slice(0, 60);
      const a = document.createElement('a');
      a.href = canvas.toDataURL('image/png');
      a.download = `${safe || 'waterline'}.png`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }

    function focusPanel(i) {
      activePanel.value = i;
    }

    async function onGlobalPaste(event) {
      const items = event.clipboardData && event.clipboardData.items;
      if (!items) return;
      for (const item of items) {
        if (item.kind === 'file' && /^image\//.test(item.type)) {
          const file = item.getAsFile();
          if (file) {
            event.preventDefault();
            await setImageFromFile(activePanel.value, file);
            return;
          }
        }
      }
    }

    onMounted(() => {
      document.addEventListener('paste', onGlobalPaste);
    });
    onUnmounted(() => {
      document.removeEventListener('paste', onGlobalPaste);
      if (saveTimer) clearTimeout(saveTimer);
    });

    return {
      panels,
      years,
      MONTHS,
      activePanel,
      dragOver,
      fileInputs,
      labelFor,
      setTitle,
      openFileDialog,
      onFileChange,
      onDrop,
      onDragOver,
      onDragLeave,
      clearImage,
      focusPanel,
      exportPanel,
    };
  },
  template: `
    <div class="waterline-page">
      <div class="waterline-grid">
        <section
          v-for="(panel, i) in panels"
          :key="i"
          class="waterline-panel"
          :class="{ active: activePanel === i, dragover: dragOver === i }"
          tabindex="0"
          @mousedown="focusPanel(i)"
          @focus="focusPanel(i)"
        >
          <header class="waterline-head">
            <div class="waterline-left">
              <div class="waterline-tag">
                <span class="waterline-tag-word">WATERLINE</span>
                <select class="waterline-select" v-model.number="panel.month" aria-label="Month">
                  <option v-for="(m, mi) in MONTHS" :key="mi" :value="mi">{{ m }}</option>
                </select>
                <select class="waterline-select" v-model.number="panel.year" aria-label="Year">
                  <option v-for="y in years" :key="y" :value="y">{{ y }}</option>
                </select>
              </div>
            </div>
            <input
              class="waterline-title"
              type="text"
              :value="panel.title"
              placeholder="TITLE"
              spellcheck="false"
              @input="setTitle(i, $event.target.value)"
            />
            <div class="waterline-right">
              <button
                type="button"
                class="waterline-export-btn"
                title="Export this panel as an image (title + date + image)"
                @click="exportPanel(i)"
              >
                Export
              </button>
            </div>
          </header>

          <div
            class="waterline-body"
            @dragover.prevent="onDragOver(i)"
            @dragleave="onDragLeave(i)"
            @drop.prevent="onDrop(i, $event)"
          >
            <template v-if="panel.image">
              <img class="waterline-image" :src="panel.image" alt="" />
              <div class="waterline-img-actions">
                <button type="button" class="waterline-mini-btn" @click="openFileDialog(i)">
                  Replace
                </button>
                <button type="button" class="waterline-mini-btn" @click="clearImage(i)">
                  Remove
                </button>
              </div>
            </template>
            <button
              v-else
              type="button"
              class="waterline-drop"
              @click="openFileDialog(i)"
            >
              <span class="waterline-drop-icon" aria-hidden="true">+</span>
              <span class="waterline-drop-text">Import an image</span>
              <span class="waterline-drop-hint">Click · drag &amp; drop · paste (Ctrl+V)</span>
            </button>

            <input
              :ref="el => { if (el) fileInputs[i] = el; }"
              type="file"
              accept="image/*"
              class="waterline-file"
              @change="onFileChange(i, $event)"
            />
          </div>
        </section>
      </div>
    </div>
  `,
};
