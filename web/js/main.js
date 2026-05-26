import { createApp, ref, computed, onMounted, onUnmounted } from 'vue';
import BdGrid from './BdGrid.js?v=input-fix2';
import SynthesisGrid from './SynthesisGrid.js?v=input-fix2';
import AppSidebar from './AppSidebar.js?v=syn-perf32';
import EmptyPage from './EmptyPage.js?v=syn-perf32';
import MatrixModal from './MatrixModal.js?v=matrix11';
import { NAV_ITEMS, DEFAULT_ROUTE } from './navConfig.js?v=syn-perf32';
import { transformBdSheet, transformSynthesisSheet } from './sheetTransform.js?v=20260526-syn-max-422';
import { createWorkbookSession } from './workbookSession.js?v=adapt-sum1';
import { buildMatrixState, applyMatrixSave } from './structureModel.js?v=matrix10';
import {
  upsertRawCell,
  loadLocalSnapshot,
  clearLocalSnapshot,
  createAutoSave,
  applySheetEdits,
} from './sessionPersistence.js?v=persist-v2';

const App = {
  components: { BdGrid, SynthesisGrid, AppSidebar, EmptyPage, MatrixModal },
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
    let engineStarted = false;
    let synthesisLoadPromise = null;
    /** Keep last known synthesis raw when BD-only saves run before Syn is opened. */
    let preservedSynRaw = null;

    const autoSave = createAutoSave(
      () => ({
        bd: bdRaw.value,
        syn: synRaw.value ?? preservedSynRaw,
        revision: dirty.value,
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

    const overlayMessage = computed(() => {
      if (isSynthesis.value && synthesisLoading.value) return 'Loading synthesis…';
      if (isSynthesis.value && bdLoading.value) return 'Loading…';
      if (bdLoading.value) return 'Loading database…';
      return 'Loading…';
    });

    const showContentOverlay = computed(() => {
      if (error.value) return false;
      if (isSynthesis.value) {
        if (synthesisSheet.value) return false;
        return synthesisLoading.value || bdLoading.value;
      }
      if (isDatabase.value) {
        if (bdSheet.value) return false;
        return bdLoading.value;
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

    function cloneRaw(raw) {
      return JSON.parse(JSON.stringify(raw));
    }

    function applyBdFromRaw() {
      bdSheet.value = transformBdSheet(cloneRaw(bdRaw.value));
    }

    function applySynFromRaw() {
      if (!synRaw.value) return;
      synthesisSheet.value = transformSynthesisSheet(cloneRaw(synRaw.value));
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
      synthesisLoadPromise = (async () => {
        if (!synRaw.value) await fetchSynFromServer();
        applySynFromRaw();
      })()
        .catch((e) => {
          error.value = e.message;
          synthesisLoadPromise = null;
        })
        .finally(() => {
          synthesisLoading.value = false;
        });
      return synthesisLoadPromise;
    }

    onMounted(async () => {
      document.title = 'WGHT Dashboard';
      const routeParam = new URLSearchParams(location.search).get('route');
      if (routeParam && NAV_ITEMS.some((n) => n.id === routeParam)) {
        route.value = routeParam;
      }
      const wantSynthesis = route.value === 'synthesis';

      try {
        const snapshot = await loadLocalSnapshot({ timeoutMs: 15000 });
        if (snapshot?.version === 2) {
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

        applyBdFromRaw();

        if (wantSynthesis) {
          if (synRaw.value) {
            applySynFromRaw();
          } else {
            await loadSynthesis();
          }
        } else {
          synthesisLoading.value = false;
        }

        scheduleEngine();
      } catch (e) {
        error.value = e.message;
        console.error('Startup failed:', e);
      } finally {
        bdLoading.value = false;
      }

      window.addEventListener('beforeunload', onBeforeUnload);
      window.addEventListener('pagehide', onBeforeUnload);
    });

    function onBeforeUnload() {
      void autoSave.saveNow();
    }

    onUnmounted(() => {
      window.removeEventListener('beforeunload', onBeforeUnload);
      window.removeEventListener('pagehide', onBeforeUnload);
      autoSave.destroy();
      session.destroy();
    });

    function onCellChange({ row, col, value, sheet }) {
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
      loadedFromLocal.value = false;
      engineStarted = false;
      synthesisLoadPromise = null;
      synthesisSheet.value = null;
      bdLoading.value = true;
      try {
        await fetchBdFromServer();
        applyBdFromRaw();
        synRaw.value = null;
        if (route.value === 'synthesis') {
          await fetchSynFromServer();
          applySynFromRaw();
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
      if (id === 'synthesis') loadSynthesis();
      if (id === 'database' || id === 'synthesis') scheduleEngine();
    }

    function toggleOutline() {
      outlineOnly.value = !outlineOnly.value;
    }

    async function openMatrix() {
      if (!bdSheet.value) return;
      await loadSynthesis();
      matrixState.value = buildMatrixState(bdSheet.value, synRaw.value);
      matrixOpen.value = true;
    }

    function closeMatrix() {
      matrixOpen.value = false;
    }

    async function onMatrixSave({ bd: bdModel }) {
      if (!bdRaw.value || matrixSaving.value) return;
      matrixSaving.value = true;
      try {
        const result = applyMatrixSave(
          bdRaw.value,
          synRaw.value,
          bdModel,
          matrixState.value?.syn ?? null
        );
        bdRaw.value = result.bdRaw;
        applyBdFromRaw();
        if (result.synRaw) {
          synRaw.value = result.synRaw;
          preservedSynRaw = result.synRaw;
          applySynFromRaw();
        }
        if (engineStarted) {
          await session.loadSheets([
          { name: 'BD', data: slimBdEnginePayload(bdSheet.value) },
        ]);
        }
        matrixOpen.value = false;
        dirty.value += 1;
        await autoSave.saveNow();
      } catch (e) {
        error.value = e?.message || String(e);
        console.error('Matrix save failed:', e);
      } finally {
        matrixSaving.value = false;
      }
    }

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
      onMatrixSave,
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
      <div class="app-body">
        <main class="app-content">
          <SynthesisGrid
            v-if="isSynthesis && synthesisSheet"
            :sheet="synthesisSheet"
            :session="session"
            :raw-syn="synRaw"
            :outline-only="outlineOnly"
            @cell-change="onCellChange"
          />
          <BdGrid
            v-else-if="isDatabase && bdSheet"
            :sheet="bdSheet"
            sheet-name="BD"
            :session="session"
            :raw-bd="bdRaw"
            :outline-only="outlineOnly"
            @cell-change="onCellChange"
          />
          <div v-else-if="error" class="loading-overlay error-text">{{ error }}</div>
          <div v-else-if="showContentOverlay" class="loading-overlay">{{ overlayMessage }}</div>
          <div v-else-if="isSynthesis && !synthesisSheet && error" class="loading-overlay error-text">{{ error }}</div>
          <div v-else-if="isSynthesis && !synthesisSheet" class="loading-overlay error-text">
            Missing synthesis-sheet.json
          </div>
          <EmptyPage v-else :title="currentNav.label" />
        </main>
      </div>
      <MatrixModal
        :open="matrixOpen"
        :state="matrixState"
        :saving="matrixSaving"
        @close="closeMatrix"
        @save="onMatrixSave"
      />
    </div>
  `,
};

createApp(App).mount('#app');
