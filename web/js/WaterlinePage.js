/**
 * Waterline — planche fond noir + grands rectangles blancs (2 lignes × 20 colonnes,
 * 4 visibles à l'écran, le reste accessible par l'ascenseur horizontal).
 *
 * Chaque rectangle :
 *  - titre éditable (noir, majuscules par défaut) centré,
 *  - libellé "WATERLINE <MOIS> <ANNEE>" à gauche (menus déroulants mois / année),
 *  - bouton véhicule (à gauche d'Export) ouvrant un menu à cases à cocher listant
 *    TOUS les trios silhouette / Hybridization / curb mass de Synthesis dont le
 *    curb mass est en jaune fluo (ligne 16). Le curb mass est lu en direct et se
 *    met à jour si Synthesis change.
 *  - cocher un véhicule crée une "bulle" bleue (silhouette + hybridization + curb
 *    mass, rond + ligne pointillée) déplaçable sur l'image ; les bulles sont
 *    reproduites à l'export PNG.
 *  - bouton Export (à droite) : génère un PNG (date à gauche, titre centré, image
 *    et bulles).
 *
 * État (titre, mois, année, image, bulles) conservé localement (localStorage).
 * Données véhicules lues via le lien injecté `synthesisCellLink`.
 */
import { ref, computed, inject, onMounted, onUnmounted, watch } from 'vue';
import { synCalcExcelCols } from './synthesisCalc.js';
import { isSynRow16FluoCol } from './synStore.js';
import { excelToDisplayCol, colToNum, numToCol } from './synthesisPerf.js';

const STORAGE_KEY = 'waterline-panels-v1';
/** 2 lignes × 20 colonnes (chaque colonne = 1 image en haut + 1 en bas). */
const PANEL_ROWS = 2;
const PANEL_COLS = 20;
const PANEL_COUNT = PANEL_ROWS * PANEL_COLS;

/** Mois abrégés FR (majuscules). "WATERLINE FEV 2026" => février. */
const MONTHS = [
  'JAN', 'FEV', 'MAR', 'AVR', 'MAI', 'JUN',
  'JUL', 'AOU', 'SEP', 'OCT', 'NOV', 'DEC',
];

/** Géométrie d'une bulle (px à l'échelle 1, réutilisée pour le DOM et l'export). */
const BUBBLE = {
  dotR: 6,
  lineLen: 40,
  gap: 4,
  font: 13,
  padX: 10,
  padY: 5,
  radius: 7,
  border: 1.5,
};
const BUBBLE_HEIGHT = BUBBLE.font + BUBBLE.padY * 2;

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
    bubbles: [],
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
      const p = parsed[i];
      if (p && typeof p === 'object') {
        base[i] = {
          title: typeof p.title === 'string' ? p.title : '',
          month: Number.isInteger(p.month) ? p.month : base[i].month,
          year: Number.isInteger(p.year) ? p.year : base[i].year,
          image: typeof p.image === 'string' ? p.image : null,
          bubbles: Array.isArray(p.bubbles)
            ? p.bubbles
                .filter((b) => b && typeof b.id === 'string')
                .map((b) => ({
                  id: b.id,
                  x: Number.isFinite(b.x) ? b.x : 0.05,
                  y: Number.isFinite(b.y) ? b.y : 0.1,
                  snap: b.snap && typeof b.snap === 'object' ? b.snap : null,
                }))
            : [],
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

function formatCurb(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  const num = s.replace(/kg/i, '').trim();
  return num ? `${num}kg` : '';
}

function roundRect(ctx, x, y, w, h, r) {
  const rr = Math.min(r, w / 2, h / 2);
  ctx.beginPath();
  ctx.moveTo(x + rr, y);
  ctx.arcTo(x + w, y, x + w, y + h, rr);
  ctx.arcTo(x + w, y + h, x, y + h, rr);
  ctx.arcTo(x, y + h, x, y, rr);
  ctx.arcTo(x, y, x + w, y, rr);
  ctx.closePath();
}

export default {
  name: 'WaterlinePage',
  setup() {
    const panels = ref(loadPanels());
    const years = yearChoices();
    const activePanel = ref(0);
    const dragOver = ref(-1);
    const fileInputs = ref([]);
    const imageEls = ref([]);
    const canvasEls = ref([]);
    /** Index du panneau dont le menu véhicules est ouvert (-1 = aucun). */
    const menuPanel = ref(-1);
    /** Position fixe du menu (téléporté dans <body> pour rester au premier plan). */
    const menuPos = ref({ top: 0, left: 0 });

    const synLink = inject('synthesisCellLink', null);

    let saveTimer = 0;
    function persist() {
      if (saveTimer) clearTimeout(saveTimer);
      saveTimer = setTimeout(() => {
        saveTimer = 0;
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(panels.value));
        } catch (e) {
          console.warn('[waterline] sauvegarde locale impossible:', e);
        }
      }, 300);
    }
    watch(panels, persist, { deep: true });

    // ── Données véhicules (Synthesis) ────────────────────────────────────────
    function rawCellAt(raw, row, excelCol) {
      if (!raw) return '';
      if (Array.isArray(raw.cells)) {
        const hit = raw.cells.find(
          (c) => Number(c.r) === row && String(c.c) === excelCol
        );
        if (hit && hit.v != null && String(hit.v).trim() !== '') {
          return String(hit.v).trim();
        }
      }
      const hr = raw.headerRows && raw.headerRows[row];
      if (hr && hr[excelCol] && hr[excelCol].v != null) {
        const v = String(hr[excelCol].v).trim();
        if (v !== '') return v;
      }
      return '';
    }

    /** Première valeur non vide sur la ligne parmi la colonne et son triplet. */
    function rowValueTriplet(raw, row, excelCol) {
      const direct = rawCellAt(raw, row, excelCol);
      if (direct) return direct;
      const base = colToNum(excelCol);
      for (let k = 1; k <= 2; k++) {
        const v = rawCellAt(raw, row, numToCol(base + k));
        if (v) return v;
      }
      return '';
    }

    /**
     * Liste réactive des véhicules : trios silhouette/hybridization/curb mass des
     * colonnes dont le curb mass (ligne 16) est en jaune fluo.
     */
    const vehicles = computed(() => {
      if (!synLink) return [];
      // Dépendances réactives : recalcul à chaque édition / recalcul Synthesis.
      if (synLink.synEditTick) void synLink.synEditTick.value;
      if (synLink.synRevision) void synLink.synRevision.value;
      if (synLink.session && synLink.session.synCalcTick) {
        void synLink.session.synCalcTick.value;
      }
      const raw = synLink.synRaw && synLink.synRaw.value;
      const cols = synCalcExcelCols();
      const out = [];
      const seen = new Set();
      for (const col of cols) {
        if (!isSynRow16FluoCol(16, col)) continue;
        if (seen.has(col)) continue;
        const silhouette = rowValueTriplet(raw, 5, col);
        const hybridization = rowValueTriplet(raw, 6, col);
        const curbMass = formatCurb(
          synLink.getSynRow16Display ? synLink.getSynRow16Display(col) : ''
        );
        if (!silhouette && !hybridization && !curbMass) continue;
        seen.add(col);
        out.push({
          id: col,
          displayCol: excelToDisplayCol(col),
          silhouette,
          hybridization,
          curbMass,
        });
      }
      return out;
    });

    const vehiclesById = computed(() => {
      const map = {};
      for (const v of vehicles.value) map[v.id] = v;
      return map;
    });

    function bubbleParts(bubble) {
      const live = vehiclesById.value[bubble.id];
      const data = live || bubble.snap || {};
      return {
        silhouette: data.silhouette || '',
        hybridization: data.hybridization || '',
        curbMass: data.curbMass || '',
      };
    }

    // Garde l'instantané (snap) à jour quand une valeur live change (curb mass…).
    watch(vehiclesById, (map) => {
      let touched = false;
      for (const panel of panels.value) {
        for (const b of panel.bubbles || []) {
          const v = map[b.id];
          if (v) {
            const snap = {
              silhouette: v.silhouette,
              hybridization: v.hybridization,
              curbMass: v.curbMass,
            };
            if (JSON.stringify(snap) !== JSON.stringify(b.snap)) {
              b.snap = snap;
              touched = true;
            }
          }
        }
      }
      if (touched) persist();
    });

    // ── Menu véhicules ───────────────────────────────────────────────────────
    function toggleMenu(i, ev) {
      if (menuPanel.value === i) {
        menuPanel.value = -1;
        return;
      }
      menuPanel.value = i;
      const MENU_W = 300;
      const MENU_H = 360;
      const btn =
        ev && ev.currentTarget && ev.currentTarget.getBoundingClientRect
          ? ev.currentTarget.getBoundingClientRect()
          : { top: 80, bottom: 104, left: 200, right: 228 };
      const vw = window.innerWidth || 1280;
      const vh = window.innerHeight || 800;
      let top = btn.bottom + 6;
      if (top + MENU_H > vh) top = Math.max(8, btn.top - MENU_H - 6);
      let left = btn.right - MENU_W;
      left = Math.min(Math.max(8, left), vw - MENU_W - 8);
      menuPos.value = { top, left };
      if (synLink && synLink.ensureSyn) {
        void synLink.ensureSyn();
        if (synLink.ensureSynGrid) void synLink.ensureSynGrid();
      }
    }

    function isSelected(i, vehicleId) {
      return (panels.value[i].bubbles || []).some((b) => b.id === vehicleId);
    }

    function toggleVehicle(i, vehicle) {
      const panel = panels.value[i];
      if (!Array.isArray(panel.bubbles)) panel.bubbles = [];
      const idx = panel.bubbles.findIndex((b) => b.id === vehicle.id);
      if (idx >= 0) {
        panel.bubbles.splice(idx, 1);
        return;
      }
      const n = panel.bubbles.length;
      panel.bubbles.push({
        id: vehicle.id,
        x: 0.06,
        y: Math.min(0.12 + n * 0.12, 0.9),
        snap: {
          silhouette: vehicle.silhouette,
          hybridization: vehicle.hybridization,
          curbMass: vehicle.curbMass,
        },
      });
    }

    function removeBubble(i, bubbleIdx) {
      panels.value[i].bubbles.splice(bubbleIdx, 1);
    }

    function onDocPointerDown(e) {
      if (menuPanel.value < 0) return;
      const t = e.target;
      if (
        t &&
        t.closest &&
        (t.closest('.waterline-menu') || t.closest('.waterline-veh-btn'))
      ) {
        return;
      }
      menuPanel.value = -1;
    }

    // ── Déplacement des bulles ───────────────────────────────────────────────
    let drag = null;
    function startBubbleDrag(i, bubbleIdx, e) {
      const canvas = e.currentTarget.closest('.waterline-canvas');
      if (!canvas) return;
      const rect = canvas.getBoundingClientRect();
      const bubble = panels.value[i].bubbles[bubbleIdx];
      drag = {
        i,
        bubbleIdx,
        rect,
        startClientX: e.clientX,
        startClientY: e.clientY,
        startX: bubble.x,
        startY: bubble.y,
      };
      window.addEventListener('pointermove', onBubbleDragMove);
      window.addEventListener('pointerup', onBubbleDragEnd);
      e.preventDefault();
      e.stopPropagation();
    }

    function onBubbleDragMove(e) {
      if (!drag) return;
      const dx = (e.clientX - drag.startClientX) / Math.max(1, drag.rect.width);
      const dy = (e.clientY - drag.startClientY) / Math.max(1, drag.rect.height);
      const bubble = panels.value[drag.i].bubbles[drag.bubbleIdx];
      bubble.x = Math.min(0.97, Math.max(0, drag.startX + dx));
      bubble.y = Math.min(0.97, Math.max(0.02, drag.startY + dy));
    }

    function onBubbleDragEnd() {
      drag = null;
      window.removeEventListener('pointermove', onBubbleDragMove);
      window.removeEventListener('pointerup', onBubbleDragEnd);
    }

    // ── Image : import / drop / paste ────────────────────────────────────────
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

    // ── Export PNG (image + bulles) ──────────────────────────────────────────
    function drawBubbleGroup(ctx, leftX, centerY, scale, parts) {
      const dotR = BUBBLE.dotR * scale;
      const lineLen = BUBBLE.lineLen * scale;
      const gap = BUBBLE.gap * scale;
      const fontPx = BUBBLE.font * scale;
      const padX = BUBBLE.padX * scale;
      const padY = BUBBLE.padY * scale;

      ctx.save();
      ctx.lineCap = 'round';

      // Longue ligne pointillée à gauche du rond (7× la ligne de droite).
      const leftLineLen = 7 * lineLen;
      ctx.beginPath();
      ctx.setLineDash([3 * scale, 3 * scale]);
      ctx.moveTo(leftX - leftLineLen, centerY);
      ctx.lineTo(leftX, centerY);
      ctx.strokeStyle = '#1b3f72';
      ctx.lineWidth = 1.5 * scale;
      ctx.stroke();
      ctx.setLineDash([]);

      // Rond d'ancrage.
      ctx.beginPath();
      ctx.arc(leftX + dotR, centerY, dotR, 0, Math.PI * 2);
      ctx.fillStyle = '#1b3f72';
      ctx.fill();
      ctx.lineWidth = 1.5 * scale;
      ctx.strokeStyle = '#0d2747';
      ctx.stroke();

      // Ligne pointillée.
      const lineStart = leftX + 2 * dotR;
      ctx.beginPath();
      ctx.setLineDash([3 * scale, 3 * scale]);
      ctx.moveTo(lineStart, centerY);
      ctx.lineTo(lineStart + lineLen, centerY);
      ctx.strokeStyle = '#1b3f72';
      ctx.lineWidth = 1.5 * scale;
      ctx.stroke();
      ctx.setLineDash([]);

      // Mesure du texte.
      ctx.font = `700 ${fontPx}px "Segoe UI", Arial, sans-serif`;
      const segs = [
        { t: parts.silhouette, c: '#0a1f3f' },
        { t: parts.hybridization, c: '#b07a00' },
        { t: parts.curbMass, c: '#0a1f3f' },
      ].filter((s) => s.t);
      const spaceW = ctx.measureText(' ').width;
      let textW = 0;
      segs.forEach((s, idx) => {
        textW += ctx.measureText(s.t).width;
        if (idx < segs.length - 1) textW += spaceW;
      });

      const bubbleH = fontPx + padY * 2;
      const bubbleW = textW + padX * 2;
      const bx = lineStart + lineLen + gap;
      const by = centerY - bubbleH / 2;

      const grad = ctx.createLinearGradient(0, by, 0, by + bubbleH);
      grad.addColorStop(0, '#d4e4f7');
      grad.addColorStop(1, '#9fc1e8');
      roundRect(ctx, bx, by, bubbleW, bubbleH, BUBBLE.radius * scale);
      ctx.fillStyle = grad;
      ctx.fill();
      ctx.lineWidth = BUBBLE.border * scale;
      ctx.strokeStyle = '#1b3f72';
      ctx.stroke();

      ctx.textBaseline = 'middle';
      ctx.textAlign = 'left';
      let tx = bx + padX;
      for (let idx = 0; idx < segs.length; idx++) {
        ctx.fillStyle = segs[idx].c;
        ctx.fillText(segs[idx].t, tx, centerY + 0.5 * scale);
        tx += ctx.measureText(segs[idx].t).width;
        if (idx < segs.length - 1) tx += spaceW;
      }
      ctx.restore();
    }

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

        // Bulles : échelle = largeur naturelle / largeur affichée.
        const imgEl = imageEls.value[i];
        const displayedW = imgEl && imgEl.clientWidth ? imgEl.clientWidth : width;
        const scale = width / displayedW;
        for (const b of panel.bubbles || []) {
          const leftX = b.x * width;
          const centerY = headerH + b.y * imgH;
          drawBubbleGroup(ctx, leftX, centerY, scale, bubbleParts(b));
        }
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

      // Nom : waterline-<mois>-<annee>-<titre> (espaces du titre remplacés par "-").
      const monthStr = MONTHS[panel.month] || '';
      const titlePart = titleText ? titleText.replace(/\s+/g, '-') : '';
      const name =
        ['waterline', monthStr, String(panel.year), titlePart]
          .filter(Boolean)
          .join('-')
          .replace(/[^\w-]+/g, '-')
          .replace(/-+/g, '-')
          .replace(/^-|-$/g, '') || 'waterline';
      const a = document.createElement('a');
      a.href = canvas.toDataURL('image/png');
      a.download = `${name}.png`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }

    // ── Divers ───────────────────────────────────────────────────────────────
    function labelFor(panel) {
      return `WATERLINE ${MONTHS[panel.month] || ''} ${panel.year}`.trim();
    }
    function setTitle(i, value) {
      panels.value[i].title = String(value || '').toUpperCase();
    }

    onMounted(() => {
      document.addEventListener('paste', onGlobalPaste);
      document.addEventListener('pointerdown', onDocPointerDown, true);
      if (synLink && synLink.ensureSyn) void synLink.ensureSyn();
    });
    onUnmounted(() => {
      document.removeEventListener('paste', onGlobalPaste);
      document.removeEventListener('pointerdown', onDocPointerDown, true);
      window.removeEventListener('pointermove', onBubbleDragMove);
      window.removeEventListener('pointerup', onBubbleDragEnd);
      if (saveTimer) clearTimeout(saveTimer);
    });

    return {
      panels,
      years,
      MONTHS,
      activePanel,
      dragOver,
      fileInputs,
      imageEls,
      canvasEls,
      menuPanel,
      menuPos,
      vehicles,
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
      toggleMenu,
      isSelected,
      toggleVehicle,
      removeBubble,
      bubbleParts,
      startBubbleDrag,
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
                class="waterline-veh-btn"
                :class="{ active: menuPanel === i }"
                title="Add vehicles (fluo curb mass) from Synthesis"
                @click.stop="toggleMenu(i, $event)"
              >
                <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
                  <path
                    fill="currentColor"
                    d="M5 11l1.5-4.5A2 2 0 0 1 8.4 5h7.2a2 2 0 0 1 1.9 1.5L19 11h1a1 1 0 0 1 1 1v4a1 1 0 0 1-1 1h-1v1a1 1 0 0 1-1 1h-1a1 1 0 0 1-1-1v-1H8v1a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1v-1H4a1 1 0 0 1-1-1v-4a1 1 0 0 1 1-1h1zm1.8-.5h10.4l-1-3a.5.5 0 0 0-.5-.4H8.3a.5.5 0 0 0-.5.4l-1 3zM7 15a1 1 0 1 0 0-2 1 1 0 0 0 0 2zm10 0a1 1 0 1 0 0-2 1 1 0 0 0 0 2z"
                  />
                </svg>
              </button>
              <button
                type="button"
                class="waterline-export-btn"
                title="Export this panel as an image (title + date + image + bubbles)"
                aria-label="Export"
                @click="exportPanel(i)"
              >
                <svg viewBox="0 0 24 24" width="16" height="16" aria-hidden="true">
                  <path
                    fill="none"
                    stroke="currentColor"
                    stroke-width="2"
                    stroke-linecap="round"
                    stroke-linejoin="round"
                    d="M12 15V4M8 8l4-4 4 4M5 14v4a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4"
                  />
                </svg>
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
              <div class="waterline-canvas">
                <img
                  class="waterline-image"
                  :ref="el => { if (el) imageEls[i] = el; }"
                  :src="panel.image"
                  alt=""
                />
                <div
                  v-for="(bubble, bi) in panel.bubbles"
                  :key="bubble.id"
                  class="wl-bubble-group"
                  :style="{ left: (bubble.x * 100) + '%', top: (bubble.y * 100) + '%' }"
                  @pointerdown="startBubbleDrag(i, bi, $event)"
                >
                  <span class="wl-line wl-line-left"></span>
                  <span class="wl-dot"></span>
                  <span class="wl-line"></span>
                  <span class="wl-bubble">
                    <span class="wl-sil">{{ bubbleParts(bubble).silhouette }}</span>
                    <span class="wl-hyb">{{ bubbleParts(bubble).hybridization }}</span>
                    <span class="wl-curb">{{ bubbleParts(bubble).curbMass }}</span>
                    <button
                      type="button"
                      class="wl-bubble-x"
                      title="Remove"
                      @pointerdown.stop
                      @click.stop="removeBubble(i, bi)"
                    >×</button>
                  </span>
                </div>
              </div>
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
              <svg viewBox="0 0 24 24" width="36" height="36" aria-hidden="true">
                <path
                  fill="none"
                  stroke="currentColor"
                  stroke-width="2"
                  stroke-linecap="round"
                  stroke-linejoin="round"
                  d="M12 4v11M8 11l4 4 4-4M5 16v2a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-2"
                />
              </svg>
              <span class="waterline-drop-text">Import an image</span>
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

      <Teleport to="body">
        <div
          v-if="menuPanel >= 0"
          class="waterline-menu waterline-menu-fixed"
          :style="{ top: menuPos.top + 'px', left: menuPos.left + 'px' }"
          @click.stop
          @mousedown.stop
        >
          <div class="waterline-menu-head">
            Vehicles · fluo curb mass
            <span class="waterline-menu-count">{{ vehicles.length }}</span>
          </div>
          <div class="waterline-menu-list">
            <label v-for="v in vehicles" :key="v.id" class="waterline-menu-item">
              <input
                type="checkbox"
                :checked="isSelected(menuPanel, v.id)"
                @change="toggleVehicle(menuPanel, v)"
              />
              <span class="wl-sil">{{ v.silhouette }}</span>
              <span class="wl-hyb">{{ v.hybridization }}</span>
              <span class="wl-curb">{{ v.curbMass }}</span>
            </label>
            <div v-if="!vehicles.length" class="waterline-menu-empty">
              No fluo curb-mass vehicle found in Synthesis.
            </div>
          </div>
        </div>
      </Teleport>
    </div>
  `,
};
