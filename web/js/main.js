import { createApp, ref, computed, onMounted, onUnmounted, onErrorCaptured, KeepAlive } from 'vue';
import BdGrid from './BdGrid.js?v=grid-nav4';
import SynthesisGrid from './SynthesisGrid.js?v=syn-acan-green1';
import { createEditHistory } from './editHistory.js?v=undo2';
import AppSidebar from './AppSidebar.js?v=syn-perf32';
import EmptyPage from './EmptyPage.js?v=syn-perf32';
import MatrixModal from './MatrixModal.js?v=matrix12';
import { NAV_ITEMS, DEFAULT_ROUTE } from './navConfig.js?v=syn-perf32';
import { transformBdSheet, transformSynthesisSheet } from './sheetTransform.js?v=syn-cicy1';
import { createWorkbookSession } from './workbookSession.js?v=sumprod-ab1';
import {
  SYN_CALC_FIRST_ROW,
  SYN_FILTER_BD_LABELS,
} from './synthesisCalc.js?v=sumprod-ab1';
import {
  buildMatrixState,
  applyMatrixSave,
  alignSynModelToBd,
  cloneStructure,
} from './structureModel.js?v=matrix13';
import {
  upsertRawCell,
  loadLocalSnapshot,
  clearLocalSnapshot,
  createAutoSave,
  applySheetEdits,
} from './sessionPersistence.js?v=persist-v3';

const App = {
  components: { BdGrid, SynthesisGrid, AppSidebar, EmptyPage, MatrixModal, KeepAlive },
  setup() {
    const bdLoading = ref(true);
    const synthesisLoading = ref(false);
    const error = ref(null);
    const bdSheet = ref(null);
    const synthesisSheet = ref(null);
    const bdRaw = ref(null);
    const synRaw = ref(null);
    const matrixOpen = ref(false);
    const matrixState = ref(null);
    const matrixSaving = ref(false);
    const dirty = ref(0);
    const saveStatus = ref('idle');
    const saveError = ref('');
    const loadedFromLocal = ref(false);
    const route = ref(DEFAULT_ROUTE);
    const menuOpen = ref(false);
    const outlineOnly = ref(false);
    const session = createWorkbookSession();
    const editHistory = createEditHistory();
    const historyTick = ref(0);
    const externalEditTick = ref(0);
    /** >0 after Bookmark Matrix save — triggers full-sheet local persistence. */
    const structureRevision = ref(0);
    let applyingHistory = false;
    let engineStarted = false;
    let synthesisLoadPromise = null;
    let synthesisPreparePromise = null;
    let matrixApplyTimer = null;
    let matrixPendingModel = null;
    let matrixSessionBefore = null;
    let matrixSessionDirty = false;
    let matrixApplyQueued = false;
    const MATRIX_APPLY_MS = 200;
    /** Keep last known synthesis raw when BD-only saves run before Syn is opened. */
    let preservedSynRaw = null;
    /** Skip re-running transformBdSheet / transformSynthesisSheet when raw unchanged. */
    const bdSheetRevision = ref(0);
    const synSheetRevision = ref(0);
    let bdSheetBuiltAt = -1;
    let synSheetBuiltAt = -1;

    const autoSave = createAutoSave(
      () => ({
        bd: bdRaw.value,
        syn: synRaw.value ?? preservedSynRaw,
        revision: dirty.value,
        structureRevision: structureRevision.value,
      }),
      {
        debounceMs: 400,
        onStatus(status, err) {
          saveStatus.value = status;
          if (err) saveError.value = err?.message || String(err);
          else if (status !== 'error') saveError.value = '';
        },
      }
    );

    const currentNav = computed(
      () => NAV_ITEMS.find((n) => n.id === route.value) || NAV_ITEMS[0]
    );

    const isDatabase = computed(() => route.value === 'database');
    const isSynthesis = computed(() => route.value === 'synthesis');
    const isGridPage = computed(
      () => route.value === 'database' || route.value === 'synthesis'
    );

    const canUndo = computed(() => {
      void historyTick.value;
      return editHistory.canUndo;
    });
    const canRedo = computed(() => {
      void historyTick.value;
      return editHistory.canRedo;
    });

    const overlayMessage = computed(() => {
      if (isSynthesis.value && synthesisLoading.value) return 'Loading synthesis…';
      if (isSynthesis.value && bdLoading.value) return 'Loading…';
      if (bdLoading.value) return 'Loading database…';
      return 'Loading…';
    });

    const showContentOverlay = computed(() => {
      if (error.value) return false;
      if (isDatabase.value) {
        return bdLoading.value || !bdSheet.value;
      }
      if (isSynthesis.value) {
        if (synthesisSheet.value) return false;
        return synthesisLoading.value || bdLoading.value;
      }
      return false;
    });

    function slimBdEnginePayload(sheet) {
      return {
        columns: sheet.columns,
        lastRow: sheet.lastRow,
        cells: sheet.cells,
        headerRows: sheet.headerRows,
      };
    }

    function scheduleEngine() {
      if (engineStarted || !bdSheet.value || !isGridPage.value) return;
      engineStarted = true;
      let engineLoaded = false;
      const run = () => {
        if (engineLoaded) return;
        engineLoaded = true;
        window.removeEventListener('pointerdown', onInteract, true);
        window.removeEventListener('keydown', onInteract, true);
        void session.loadSheets([
          { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
        ]);
      };
      const onInteract = () => run();
      setTimeout(run, 3000);
      window.addEventListener('pointerdown', onInteract, true);
      window.addEventListener('keydown', onInteract, true);
    }

    function applyBdFromRaw(force = false) {
      if (!bdRaw.value) return;
      if (!force && bdSheet.value && bdSheetBuiltAt === bdSheetRevision.value) {
        return;
      }
      bdSheet.value = transformBdSheet(bdRaw.value);
      bdSheetBuiltAt = bdSheetRevision.value;
    }

    function applySynFromRaw(force = false) {
      if (!synRaw.value) return;
      if (!force && synthesisSheet.value && synSheetBuiltAt === synSheetRevision.value) {
        return;
      }
      synthesisSheet.value = transformSynthesisSheet(synRaw.value);
      synSheetBuiltAt = synSheetRevision.value;
    }

    function bumpBdSheetRevision() {
      bdSheetRevision.value += 1;
    }

    function bumpSynSheetRevision() {
      synSheetRevision.value += 1;
    }

    /** Defer 7MB synthesis fetch/transform until the browser is idle (keeps BD snappy). */
    function scheduleSynBackgroundPrepare() {
      if (synthesisSheet.value || synthesisPreparePromise) return;
      const run = () => {
        if (synthesisSheet.value || synthesisPreparePromise) return;
        void startSynBackgroundPrepare();
      };
      if (typeof requestIdleCallback === 'function') {
        requestIdleCallback(run, { timeout: 12000 });
      } else {
        setTimeout(run, 3000);
      }
    }

    /** Fetch + transform Synthesis in background (no grid mount until user opens the page). */
    function startSynBackgroundPrepare() {
      if (synthesisSheet.value) return Promise.resolve();
      if (synthesisPreparePromise) return synthesisPreparePromise;
      synthesisPreparePromise = (async () => {
        if (!synRaw.value) await fetchSynFromServer();
        bumpSynSheetRevision();
        applySynFromRaw(true);
      })().catch((e) => {
        synthesisPreparePromise = null;
        console.warn('Synthesis background prepare failed:', e);
        throw e;
      });
      return synthesisPreparePromise;
    }

    async function fetchBdFromServer() {
      const bdRes = await fetch('/public/data/bd-sheet.json');
      if (!bdRes.ok) throw new Error(`Failed to load BD data (${bdRes.status})`);
      bdRaw.value = await bdRes.json();
    }

    async function fetchSynFromServer() {
      const res = await fetch('/public/data/synthesis-sheet.json');
      if (!res.ok) throw new Error(`Synthesis data (${res.status})`);
      synRaw.value = await res.json();
      preservedSynRaw = synRaw.value;
    }

    function loadSynthesis() {
      if (synthesisSheet.value) return Promise.resolve();
      if (synthesisLoadPromise) return synthesisLoadPromise;
      synthesisLoading.value = true;
      synthesisLoadPromise = startSynBackgroundPrepare()
        .catch((e) => {
          error.value = e?.message || String(e);
          synthesisLoadPromise = null;
        })
        .finally(() => {
          synthesisLoading.value = false;
        });
      return synthesisLoadPromise;
    }

    onErrorCaptured((err) => {
      error.value = err?.message || String(err);
      synthesisLoading.value = false;
      bdLoading.value = false;
      console.error('Grid render failed:', err);
      return false;
    });

    onMounted(async () => {
      document.title = 'WGHT Dashboard';
      const routeParam = new URLSearchParams(location.search).get('route');
      if (routeParam && NAV_ITEMS.some((n) => n.id === routeParam)) {
        route.value = routeParam;
      }
      const wantSynthesis = route.value === 'synthesis';

      try {
        const snapshot = await loadLocalSnapshot({ timeoutMs: 15000 });
        if (snapshot?.version === 3) {
          if (snapshot.bdFull) {
            bdRaw.value = snapshot.bdFull;
          } else {
            await fetchBdFromServer();
            if (snapshot.bdEdits) {
              applySheetEdits(bdRaw.value, snapshot.bdEdits);
            }
          }
          if (snapshot.synFull) {
            synRaw.value = snapshot.synFull;
            preservedSynRaw = snapshot.synFull;
          } else if (snapshot.synEdits) {
            await fetchSynFromServer();
            applySheetEdits(synRaw.value, snapshot.synEdits);
          }
          structureRevision.value = snapshot.structureRevision ?? 0;
          loadedFromLocal.value = true;
          saveStatus.value = 'saved';
        } else if (snapshot?.version === 2) {
          await fetchBdFromServer();
          if (snapshot.bdEdits) {
            applySheetEdits(bdRaw.value, snapshot.bdEdits);
          }
          if (snapshot.synEdits) {
            await fetchSynFromServer();
            applySheetEdits(synRaw.value, snapshot.synEdits);
          }
          loadedFromLocal.value = true;
          saveStatus.value = 'saved';
        } else if (snapshot?.bd) {
          bdRaw.value = snapshot.bd;
          if (snapshot.syn) {
            synRaw.value = snapshot.syn;
            preservedSynRaw = snapshot.syn;
          }
          loadedFromLocal.value = true;
          saveStatus.value = 'saved';
        } else {
          await fetchBdFromServer();
        }

        bumpBdSheetRevision();
        applyBdFromRaw(true);

        if (wantSynthesis) {
          if (synRaw.value) {
            bumpSynSheetRevision();
            applySynFromRaw(true);
          } else {
            await loadSynthesis();
          }
        } else {
          synthesisLoading.value = false;
        }

        scheduleEngine();
        void startSynBackgroundPrepare();
      } catch (e) {
        error.value = e.message;
        console.error('Startup failed:', e);
      } finally {
        bdLoading.value = false;
      }

      window.addEventListener('beforeunload', onBeforeUnload);
      window.addEventListener('pagehide', onBeforeUnload);
      window.addEventListener('keydown', onGlobalKeydown, true);
    });

    function onBeforeUnload() {
      void autoSave.saveNow();
    }

    onUnmounted(() => {
      window.removeEventListener('beforeunload', onBeforeUnload);
      window.removeEventListener('pagehide', onBeforeUnload);
      window.removeEventListener('keydown', onGlobalKeydown, true);
      autoSave.destroy();
      session.destroy();
    });

    function syncSheetCell(gridSheet, row, col, value) {
      if (!gridSheet) return;
      const key = `${row}:${col}`;
      const map = gridSheet.cellMap instanceof Map ? gridSheet.cellMap : null;
      if (!map) return;
      let cell = map.get(key);
      if (!cell) {
        cell = { r: row, c: col, v: value, userEdited: true };
        gridSheet.cells.push(cell);
        map.set(key, cell);
      } else {
        cell.v = value;
        cell.userEdited = true;
        delete cell.f;
      }
    }

    function cloneRawSheet(sheet) {
      if (!sheet) return null;
      return JSON.parse(JSON.stringify(sheet));
    }

    async function applyMatrixSnapshot(bd, syn) {
      if (!bd) return;
      bdRaw.value = cloneRawSheet(bd);
      bumpBdSheetRevision();
      applyBdFromRaw(true);
      if (syn) {
        synRaw.value = cloneRawSheet(syn);
        preservedSynRaw = synRaw.value;
        bumpSynSheetRevision();
        applySynFromRaw(true);
      }
      externalEditTick.value += 1;
      if (engineStarted && bdSheet.value) {
        await session.loadSheets([
          { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
        ]);
      }
      dirty.value += 1;
      autoSave.schedule();
      void autoSave.saveNow();
    }

    function applyHistoryCell({ sheet, row, col, value }) {
      const raw =
        sheet === 'SYNTHESIS'
          ? synRaw.value
          : sheet === 'BD'
            ? bdRaw.value
            : null;
      const gridSheet =
        sheet === 'SYNTHESIS' ? synthesisSheet.value : bdSheet.value;
      if (!raw || !gridSheet) return;

      upsertRawCell(raw, row, col, value);
      syncSheetCell(gridSheet, row, col, value);
      externalEditTick.value += 1;

      const engineSheet = sheet === 'SYNTHESIS' ? 'SYNTHESIS' : 'BD';
      void session
        .setCellValue(engineSheet, row, col, value)
        .catch((e) => console.warn('History setCellValue:', e));

      dirty.value += 1;
      autoSave.schedule();
      void autoSave.saveNow();
    }

    async function performUndo() {
      const entry = editHistory.undo();
      if (!entry) return;
      applyingHistory = true;
      try {
        if (entry.type === 'matrix') {
          await applyMatrixSnapshot(entry.bdBefore, entry.synBefore);
        } else {
          applyHistoryCell({ ...entry, value: entry.oldValue });
        }
      } finally {
        applyingHistory = false;
        historyTick.value = editHistory.revision;
      }
    }

    async function performRedo() {
      const entry = editHistory.redo();
      if (!entry) return;
      applyingHistory = true;
      try {
        if (entry.type === 'matrix') {
          await applyMatrixSnapshot(entry.bdAfter, entry.synAfter);
        } else {
          applyHistoryCell({ ...entry, value: entry.newValue });
        }
      } finally {
        applyingHistory = false;
        historyTick.value = editHistory.revision;
      }
    }

    function isEditingGridInput() {
      const el = document.activeElement;
      return el?.tagName === 'INPUT' && el.classList?.contains('grid-cell-input');
    }

    function onGlobalKeydown(e) {
      if (!isGridPage.value) return;
      if (!(e.ctrlKey || e.metaKey)) return;
      if (isEditingGridInput()) return;
      const key = e.key.toLowerCase();
      if (key === 'z') {
        e.preventDefault();
        if (e.shiftKey) performRedo();
        else performUndo();
      } else if (key === 'y') {
        e.preventDefault();
        performRedo();
      }
    }

    function onCellChange({ row, col, value, sheet, previousValue }) {
      if (!applyingHistory) {
        editHistory.push({
          sheet,
          row,
          col,
          oldValue: previousValue ?? '',
          newValue: value,
        });
        historyTick.value = editHistory.revision;
      }
      const raw =
        sheet === 'SYNTHESIS'
          ? synRaw.value
          : sheet === 'BD'
            ? bdRaw.value
            : null;
      if (raw) upsertRawCell(raw, row, col, value);
      dirty.value += 1;
      autoSave.schedule();
      void autoSave.saveNow();
    }

    async function saveNow() {
      await autoSave.saveNow();
    }

    async function resetFromServerFiles() {
      if (
        !confirm(
          'Effacer la copie locale et recharger les fichiers JSON du projet ? Les modifications non exportées ailleurs seront perdues.'
        )
      ) {
        return;
      }
      await clearLocalSnapshot();
      editHistory.clear();
      historyTick.value = editHistory.revision;
      structureRevision.value = 0;
      loadedFromLocal.value = false;
      engineStarted = false;
      synthesisLoadPromise = null;
      synthesisPreparePromise = null;
      synthesisSheet.value = null;
      bdLoading.value = true;
      try {
        await fetchBdFromServer();
        bumpBdSheetRevision();
        applyBdFromRaw(true);
        synRaw.value = null;
        synthesisSheet.value = null;
        synSheetBuiltAt = -1;
        if (route.value === 'synthesis') {
          await fetchSynFromServer();
          bumpSynSheetRevision();
          applySynFromRaw(true);
        }
        await session.loadSheets([
          { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
        ]);
        engineStarted = true;
        dirty.value = 0;
        saveStatus.value = 'saved';
      } catch (e) {
        error.value = e.message;
      } finally {
        bdLoading.value = false;
      }
    }

    function navigate(id) {
      route.value = id;
      menuOpen.value = false;
      if (id === 'synthesis') {
        void loadSynthesis();
      } else if (id === 'database' || id === 'synthesis') {
        scheduleEngine();
      }
    }

    function toggleOutline() {
      outlineOnly.value = !outlineOnly.value;
    }

    async function openMatrix() {
      if (!bdSheet.value) return;
      await loadSynthesis();
      matrixState.value = buildMatrixState(bdSheet.value, synRaw.value);
      matrixSessionBefore = {
        bd: cloneRawSheet(bdRaw.value),
        syn: synRaw.value ? cloneRawSheet(synRaw.value) : null,
      };
      matrixSessionDirty = false;
      matrixPendingModel = null;
      matrixOpen.value = true;
    }

    function commitMatrixSessionHistory() {
      if (!matrixSessionBefore || applyingHistory || !matrixSessionDirty) {
        matrixSessionBefore = null;
        matrixSessionDirty = false;
        return;
      }
      editHistory.push({
        type: 'matrix',
        bdBefore: matrixSessionBefore.bd,
        synBefore: matrixSessionBefore.syn,
        bdAfter: cloneRawSheet(bdRaw.value),
        synAfter: synRaw.value ? cloneRawSheet(synRaw.value) : null,
      });
      historyTick.value = editHistory.revision;
      matrixSessionBefore = null;
      matrixSessionDirty = false;
    }

    async function applyMatrixFromModel(bdModel) {
      if (!bdRaw.value || !bdModel?.sections?.length) return;
      if (bdModel.sections.length < 2) return;
      if (matrixSaving.value) {
        matrixApplyQueued = true;
        matrixPendingModel = bdModel;
        return;
      }
      matrixSaving.value = true;
      try {
        await loadSynthesis();
        const synBase = matrixState.value?.syn ?? null;
        const synModel = synBase
          ? alignSynModelToBd(cloneStructure(synBase), bdModel)
          : null;
        const bdModelClone = cloneStructure(bdModel);
        const result = applyMatrixSave(
          bdRaw.value,
          synRaw.value,
          bdModelClone,
          synModel
        );
        bdRaw.value = result.bdRaw;
        bumpBdSheetRevision();
        applyBdFromRaw(true);
        if (result.synRaw) {
          synRaw.value = result.synRaw;
          preservedSynRaw = result.synRaw;
          bumpSynSheetRevision();
          applySynFromRaw(true);
        }
        externalEditTick.value += 1;
        if (engineStarted && bdSheet.value) {
          await session.loadSheets([
            { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
          ]);
        }
        matrixState.value = {
          bd: cloneStructure(result.bdModel),
          syn: result.synModel
            ? cloneStructure(result.synModel)
            : matrixState.value?.syn ?? null,
        };
        matrixSessionDirty = true;
        structureRevision.value = Math.max(structureRevision.value, 1);
        dirty.value += 1;
        autoSave.schedule();
        void autoSave.saveNow();
      } catch (e) {
        error.value = e?.message || String(e);
        console.error('Matrix apply failed:', e);
      } finally {
        matrixSaving.value = false;
        if (matrixApplyQueued && matrixPendingModel) {
          matrixApplyQueued = false;
          const next = matrixPendingModel;
          void applyMatrixFromModel(next);
        }
      }
    }

    function onMatrixChange({ bd }) {
      if (!matrixOpen.value || !bd) return;
      matrixPendingModel = bd;
      if (matrixApplyTimer) clearTimeout(matrixApplyTimer);
      matrixApplyTimer = setTimeout(() => {
        matrixApplyTimer = null;
        void applyMatrixFromModel(matrixPendingModel);
      }, MATRIX_APPLY_MS);
    }

    async function closeMatrix() {
      if (matrixApplyTimer) {
        clearTimeout(matrixApplyTimer);
        matrixApplyTimer = null;
      }
      if (matrixPendingModel) {
        await applyMatrixFromModel(matrixPendingModel);
        matrixPendingModel = null;
      }
      commitMatrixSessionHistory();
      matrixOpen.value = false;
    }

    const bdL2Registry = computed(() => {
      void session.revision.value;
      return session.getBdL2Registry?.() ?? [];
    });

    const synFilterBdLabels = SYN_FILTER_BD_LABELS;

    return {
      bdLoading,
      synthesisLoading,
      overlayMessage,
      showContentOverlay,
      error,
      bdSheet,
      synthesisSheet,
      isDatabase,
      isSynthesis,
      isGridPage,
      dirty,
      saveStatus,
      saveError,
      loadedFromLocal,
      saveNow,
      resetFromServerFiles,
      route,
      menuOpen,
      outlineOnly,
      currentNav,
      session,
      onCellChange,
      navigate,
      toggleOutline,
      matrixOpen,
      matrixState,
      matrixSaving,
      openMatrix,
      closeMatrix,
      onMatrixChange,
      canUndo,
      canRedo,
      performUndo,
      performRedo,
      externalEditTick,
      bdL2Registry,
      synFilterBdLabels,
      SYN_CALC_FIRST_ROW,
    };
  },
  template: `
    <div class="app-shell" :class="{ 'menu-open': menuOpen }">
      <AppSidebar
        :current="route"
        :open="menuOpen"
        @navigate="navigate"
        @close="menuOpen = false"
      />
      <header class="app-topbar">
        <div class="topbar-left">
          <button
            type="button"
            class="burger"
            aria-label="Open menu"
            @click="menuOpen = !menuOpen"
          >
            <span></span><span></span><span></span>
          </button>
          <span v-if="isDatabase" class="page-title">Database</span>
          <span v-else-if="isSynthesis" class="page-title">Synthesis</span>
          <span v-else class="page-title">{{ currentNav.label }}</span>
          <template v-if="isDatabase || isSynthesis">
            <div class="topbar-history" role="group" aria-label="Historique des modifications">
              <button
                type="button"
                class="icon-btn icon-btn-sm"
                :disabled="!canUndo"
                title="Annuler (Ctrl+Z)"
                aria-label="Annuler"
                @click="performUndo"
              >
                <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                  <path
                    fill="none"
                    stroke="currentColor"
                    stroke-width="1.35"
                    stroke-linecap="round"
                    stroke-linejoin="round"
                    d="M6 4 2 8l4 4V9h3a3 3 0 1 1 0 6H6"
                  />
                </svg>
              </button>
              <button
                type="button"
                class="icon-btn icon-btn-sm"
                :disabled="!canRedo"
                title="Rétablir (Ctrl+Y)"
                aria-label="Rétablir"
                @click="performRedo"
              >
                <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                  <path
                    fill="none"
                    stroke="currentColor"
                    stroke-width="1.35"
                    stroke-linecap="round"
                    stroke-linejoin="round"
                    d="M10 4l4 4-4 4V9H7a3 3 0 1 0 0 6h3"
                  />
                </svg>
              </button>
            </div>
            <button
              type="button"
              class="icon-btn icon-btn-sm"
              title="Bookmark Matrix — reorder sections and sub-sections"
              aria-label="Open Bookmark Matrix"
              @click="openMatrix"
            >
              <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                <rect x="1" y="1" width="6" height="6" fill="currentColor"/>
                <rect x="9" y="1" width="6" height="6" fill="currentColor"/>
                <rect x="1" y="9" width="6" height="6" fill="currentColor"/>
                <rect x="9" y="9" width="6" height="6" fill="currentColor"/>
              </svg>
            </button>
            <button
              type="button"
              class="icon-btn icon-btn-sm"
              :class="{ active: outlineOnly }"
              title="Show yellow sections and blue bands only"
              aria-label="Toggle outline view"
              @click="toggleOutline"
            >
              <svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true">
                <ellipse cx="8" cy="8" rx="6.5" ry="4" fill="none" stroke="currentColor" stroke-width="1.2"/>
                <circle cx="8" cy="8" r="2" fill="currentColor"/>
              </svg>
            </button>
          </template>
        </div>
        <div class="topbar-right">
          <span class="status" v-if="bdLoading && isDatabase">Loading…</span>
          <span class="status" v-else-if="isSynthesis && (synthesisLoading || (bdLoading && !synthesisSheet))">Loading synthesis…</span>
          <span class="status error-text" v-else-if="error">{{ error }}</span>
          <span class="status" v-else-if="isGridPage && session.loading">Calculating…</span>
          <span class="status" v-else-if="isGridPage && saveStatus === 'saving'">Enregistrement…</span>
          <span class="status saved-status" v-else-if="isGridPage && saveStatus === 'saved'">
            Modifications enregistrées (ce navigateur)
          </span>
          <span class="status" v-else-if="isGridPage && saveStatus === 'dirty'">Enregistrement automatique…</span>
          <span class="status error-text" v-else-if="isGridPage && saveStatus === 'error'">
            Erreur — {{ saveError }}
          </span>
          <template v-if="isGridPage && !session.loading">
            <button type="button" class="topbar-save-btn" @click="saveNow">Enregistrer</button>
            <button
              type="button"
              class="topbar-save-btn topbar-save-btn-muted"
              title="Recharger les JSON du projet (efface la copie locale)"
              @click="resetFromServerFiles"
            >
              Fichiers projet
            </button>
          </template>
        </div>
      </header>
      <div v-if="isDatabase && bdSheet" class="bd-syn-engine-panel">
        <details class="bd-syn-engine-details">
          <summary>
            Moteur Synthesis — lignes {{ SYN_CALC_FIRST_ROW }}+ (bleues + vertes), toutes colonnes G…NI
            <span v-if="session.ready" class="bd-syn-engine-meta">
              · {{ bdL2Registry.length }} sous-systèmes répertoriés (BD col. AU, statut Q=S)
            </span>
            <span v-else class="bd-syn-engine-meta"> · moteur en chargement…</span>
          </summary>
          <div class="bd-syn-engine-body">
            <p class="bd-syn-engine-lead">
              Colonnes G…NI (ligne {{ SYN_CALC_FIRST_ROW }} → fin) — calcul instantané depuis la Database&nbsp;:
              <strong>lignes bleues</strong> = masse filtrée (col.&nbsp;V, sous-système AU, critères lignes 3–14 de la colonne)&nbsp;;
              <strong>lignes vertes/jaunes</strong> (sections L1) = somme des lignes bleues jusqu'à la section suivante.
              Ligne 26 colonnes H–O = total adaptation (somme lignes 27–41). Modifier Database ou filtres → Synthesis à jour.
            </p>
            <div class="bd-syn-engine-grid">
              <section>
                <h3>Filtres Synthesis → BD (par colonne M…AA)</h3>
                <table class="bd-syn-engine-table">
                  <thead>
                    <tr><th>Ligne Syn</th><th>Col. BD</th><th>Champ</th></tr>
                  </thead>
                  <tbody>
                    <tr v-for="f in synFilterBdLabels" :key="f.synRow">
                      <td>{{ f.synRow }}</td>
                      <td>{{ f.bdCol }}</td>
                      <td>{{ f.label }}</td>
                    </tr>
                  </tbody>
                </table>
              </section>
              <section>
                <h3>Sous-systèmes L2 (col. AU) — masse totale Q=S</h3>
                <div class="bd-syn-l2-scroll">
                  <table class="bd-syn-engine-table bd-syn-l2-table">
                    <thead>
                      <tr><th>Label L2 (AU)</th><th>Lignes</th><th>Masse ΣV</th></tr>
                    </thead>
                    <tbody>
                      <tr v-for="entry in bdL2Registry" :key="entry.label">
                        <td>{{ entry.label }}</td>
                        <td>{{ entry.rows }}</td>
                        <td>{{ entry.mass }}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </section>
            </div>
          </div>
        </details>
      </div>
      <div class="app-body">
        <main class="app-content">
          <KeepAlive>
            <BdGrid
              v-if="bdSheet && isDatabase"
              key="bd-grid"
              :sheet="bdSheet"
              sheet-name="BD"
              :session="session"
              :raw-bd="bdRaw"
              :outline-only="outlineOnly"
              :pane-visible="isDatabase"
              :external-edit-tick="externalEditTick"
              @cell-change="onCellChange"
            />
          </KeepAlive>
          <KeepAlive>
            <SynthesisGrid
              v-if="synthesisSheet && isSynthesis"
              key="syn-grid"
              :sheet="synthesisSheet"
              :session="session"
              :raw-syn="synRaw"
              :outline-only="outlineOnly"
              :pane-visible="isSynthesis"
              :external-edit-tick="externalEditTick"
              @cell-change="onCellChange"
            />
          </KeepAlive>
          <div v-if="error" class="loading-overlay error-text">{{ error }}</div>
          <div v-else-if="showContentOverlay" class="loading-overlay">{{ overlayMessage }}</div>
          <div
            v-else-if="isSynthesis && !synthesisSheet"
            class="loading-overlay error-text"
          >
            Missing synthesis-sheet.json
          </div>
          <EmptyPage
            v-else-if="!isDatabase && !isSynthesis"
            :title="currentNav.label"
          />
        </main>
      </div>
      <MatrixModal
        :open="matrixOpen"
        :state="matrixState"
        :saving="matrixSaving"
        @close="closeMatrix"
        @change="onMatrixChange"
      />
    </div>
  `,
};

createApp(App).mount('#app');
