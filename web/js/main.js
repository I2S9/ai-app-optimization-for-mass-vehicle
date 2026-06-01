import { createApp, ref, computed, onMounted, onUnmounted, onErrorCaptured, nextTick } from 'vue';
import BdGrid from './BdGrid.js?v=grid-perf9';
import SynthesisGrid from './SynthesisGrid.js?v=syn-p2';
import { createEditHistory } from './editHistory.js?v=undo2';
import AppSidebar from './AppSidebar.js?v=syn-perf32';
import EmptyPage from './EmptyPage.js?v=syn-perf32';
import MatrixModal from './MatrixModal.js?v=matrix13';
import { NAV_ITEMS, DEFAULT_ROUTE } from './navConfig.js?v=syn-perf32';
import { transformBdSheetAsync, transformSynthesisSheetAsync } from './sheetTransform.js?v=grid-perf7';
import { createWorkbookSession } from './workbookSession.js?v=grid-perf9';
import { createPerfBench } from './perfBench.js?v=1';
import { yieldToMain, deferHeavy } from './yieldMain.js?v=2';
import {
  fetchPrecomputedPack,
  resolveGridSheet,
  hasSheetEdits,
} from './gridBoot.js?v=1';
import { loadSheetRaw, probeSheetApi, patchSheetCells } from './sheetDataApi.js?v=p2';
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
  saveSheetTransform,
  rawFingerprint,
} from './sessionPersistence.js?v=persist-v5';

const App = {
  components: { BdGrid, SynthesisGrid, AppSidebar, EmptyPage, MatrixModal },
  setup() {
    const bdLoading = ref(false);
    const gridPreparing = ref(false);
    const synthesisLoading = ref(false);
    /** @type {import('vue').Ref<{ sheetId: string, pct: number } | null>} */
    const dataLoadProgress = ref(null);
    const chunkedApi = ref(false);
    const serverCalc = ref(false);
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
    let matrixPendingModel = null;
    let matrixSessionBefore = null;
    let matrixSessionDirty = false;
    let matrixApplyQueued = false;
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
      if (isDatabase.value) return false;
      if (isSynthesis.value) {
        return !synthesisSheet.value && synthesisLoading.value;
      }
      return false;
    });

    function slimBdEnginePayload(sheet) {
      return {
        columns: sheet.columns,
        lastRow: sheet.lastRow,
        cells: sheet.cells,
        headerRows: sheet.headerRows,
        sectionHeaderRows: sheet.sectionHeaderRows,
        canonicalSectionByLabel: sheet.canonicalSectionByLabel,
        cellMap: sheet.cellMap,
      };
    }

    function scheduleEngine() {
      if (engineStarted || !bdSheet.value || !isGridPage.value) return;
      let engineLoaded = false;
      const run = () => {
        if (engineLoaded || engineStarted) return;
        engineLoaded = true;
        engineStarted = true;
        window.removeEventListener('pointerdown', onInteract, true);
        window.removeEventListener('keydown', onInteract, true);
        void session
          .loadSheets([{ name: 'BD', data: slimBdEnginePayload(bdSheet.value) }])
          .then(() => {
            externalEditTick.value += 1;
          });
      };
      const onInteract = () => run();
      window.addEventListener('pointerdown', onInteract, true);
      window.addEventListener('keydown', onInteract, true);
      if (typeof requestIdleCallback === 'function') {
        requestIdleCallback(() => run(), { timeout: 4000 });
      } else {
        setTimeout(() => run(), 3000);
      }
    }

    async function applyBdFromRaw(force = false) {
      if (!bdRaw.value) {
        if (!force) {
          const pre = await fetchPrecomputedPack('bd');
          if (pre?.sheet) {
            await yieldToMain();
            bdSheet.value = pre.sheet;
            bdSheetBuiltAt = bdSheetRevision.value;
            return;
          }
        }
        return;
      }
      if (!force && bdSheet.value && bdSheetBuiltAt === bdSheetRevision.value) {
        return;
      }
      const resolved = await resolveGridSheet('bd', bdRaw.value, { force });
      if (resolved) {
        await yieldToMain();
        bdSheet.value = resolved;
        bdSheetBuiltAt = bdSheetRevision.value;
        return;
      }
      await yieldToMain();
      const transformed = await transformBdSheetAsync(bdRaw.value);
      await yieldToMain();
      bdSheet.value = transformed;
      bdSheetBuiltAt = bdSheetRevision.value;
      const fp = rawFingerprint(bdRaw.value);
      const pre = await fetchPrecomputedPack('bd');
      void saveSheetTransform('bd', fp, bdSheet.value, {
        skipIfPrecomputed: pre?.fingerprint === fp,
      });
    }

    async function applySynFromRaw(force = false) {
      if (!synRaw.value) {
        if (!force) {
          const pre = await fetchPrecomputedPack('syn');
          if (pre?.sheet) {
            await yieldToMain();
            synthesisSheet.value = pre.sheet;
            synSheetBuiltAt = synSheetRevision.value;
            return;
          }
        }
        return;
      }
      if (!force && synthesisSheet.value && synSheetBuiltAt === synSheetRevision.value) {
        return;
      }
      const resolved = await resolveGridSheet('syn', synRaw.value, { force });
      if (resolved) {
        await yieldToMain();
        synthesisSheet.value = resolved;
        synSheetBuiltAt = synSheetRevision.value;
        return;
      }
      await yieldToMain();
      const transformed = await transformSynthesisSheetAsync(synRaw.value);
      await yieldToMain();
      synthesisSheet.value = transformed;
      synSheetBuiltAt = synSheetRevision.value;
      const fp = rawFingerprint(synRaw.value);
      const pre = await fetchPrecomputedPack('syn');
      void saveSheetTransform('syn', fp, synthesisSheet.value, {
        skipIfPrecomputed: pre?.fingerprint === fp,
      });
    }

    function bumpBdSheetRevision() {
      bdSheetRevision.value += 1;
    }

    function bumpSynSheetRevision() {
      synSheetRevision.value += 1;
    }

    /** Fetch + transform Synthesis when user opens the page (never while on Database). */
    function startSynBackgroundPrepare() {
      if (synthesisSheet.value) return Promise.resolve();
      if (synthesisPreparePromise) return synthesisPreparePromise;
      synthesisPreparePromise = (async () => {
        const pre = await fetchPrecomputedPack('syn');
        if (!synRaw.value) {
          if (pre?.sheet) {
            await yieldToMain();
            synthesisSheet.value = pre.sheet;
            synSheetBuiltAt = synSheetRevision.value;
            void fetchSynFromServer();
            return;
          }
          await fetchSynFromServer();
        }
        if (pre?.sheet) {
          const fp = rawFingerprint(synRaw.value);
          if (pre.fingerprint === fp) {
            await yieldToMain();
            synthesisSheet.value = pre.sheet;
            synSheetBuiltAt = synSheetRevision.value;
            return;
          }
        }
        bumpSynSheetRevision();
        await applySynFromRaw(false);
      })().catch((e) => {
        synthesisPreparePromise = null;
        console.warn('Synthesis prepare failed:', e);
        throw e;
      });
      return synthesisPreparePromise;
    }

    function mergeExtraCellsIntoBd(fullRaw) {
      const sheet = bdSheet.value;
      if (!sheet || !fullRaw?.cells?.length) return;
      const map = sheet.cellMap instanceof Map ? sheet.cellMap : null;
      if (!map) {
        bumpBdSheetRevision();
        return applyBdFromRaw(false);
      }
      let added = 0;
      for (const c of fullRaw.cells) {
        const k = `${c.r}:${c.c}`;
        if (!map.has(k)) {
          sheet.cells.push(c);
          map.set(k, c);
          added++;
        }
      }
      if (added > 0) externalEditTick.value += 1;
    }

    /**
     * Load BD raw — chunked API when available (meta + row windows), else full JSON.
     * @param {{ progressive?: boolean }} opts progressive=true → show grid after 1st chunk
     */
    async function loadBdFromServer({ progressive = true } = {}) {
      bdLoading.value = true;
      let earlyRender = false;
      try {
        const full = await loadSheetRaw('bd', {
          onProgress(info) {
            if (info.lastRow > 0) {
              dataLoadProgress.value = {
                sheetId: 'bd',
                pct: Math.min(
                  100,
                  Math.round((info.rowMax / info.lastRow) * 100)
                ),
              };
            }
          },
          async onFirstChunk(partial) {
            if (!progressive || earlyRender) return;
            earlyRender = true;
            bdRaw.value = partial;
            bumpBdSheetRevision();
            await applyBdFromRaw(false);
            gridPreparing.value = false;
            scheduleEngine();
          },
        });
        bdRaw.value = full;
        if (!earlyRender) {
          bumpBdSheetRevision();
          await applyBdFromRaw(false);
          gridPreparing.value = false;
          scheduleEngine();
        } else {
          mergeExtraCellsIntoBd(full);
        }
      } finally {
        bdLoading.value = false;
        dataLoadProgress.value = null;
      }
    }

    async function fetchBdFromServer() {
      await loadBdFromServer({ progressive: false });
    }

    async function fetchSynFromServer() {
      synthesisLoading.value = true;
      try {
        synRaw.value = await loadSheetRaw('synthesis', {
          onProgress(info) {
            if (info.lastRow > 0) {
              dataLoadProgress.value = {
                sheetId: 'syn',
                pct: Math.min(
                  100,
                  Math.round((info.rowMax / info.lastRow) * 100)
                ),
              };
            }
          },
        });
        preservedSynRaw = synRaw.value;
      } finally {
        dataLoadProgress.value = null;
        synthesisLoading.value = false;
      }
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

    onMounted(() => {
      document.title = 'WGHT Dashboard';
      const routeParam = new URLSearchParams(location.search).get('route');
      if (routeParam && NAV_ITEMS.some((n) => n.id === routeParam)) {
        route.value = routeParam;
      }
      const bench = createPerfBench(() => ({
        session,
        bdSheet: bdSheet.value,
        synthesisSheet: synthesisSheet.value,
        applyBdFromRaw,
        applySynFromRaw,
      }));
      if (typeof window !== 'undefined') {
        window.__runPerfBench = () => bench.run();
      }
      void bootApp();
      window.addEventListener('beforeunload', onBeforeUnload);
      window.addEventListener('pagehide', onBeforeUnload);
      window.addEventListener('keydown', onGlobalKeydown, true);
    });

    async function bootApp() {
      const wantSynthesis = route.value === 'synthesis';
      void fetchPrecomputedPack('syn');

      try {
        await nextTick();
        const [snapshot, apiCfg, preBd] = await Promise.all([
          loadLocalSnapshot({ timeoutMs: 1500 }),
          probeSheetApi(),
          fetchPrecomputedPack('bd'),
        ]);
        chunkedApi.value = Boolean(apiCfg?.chunkedLoad);
        serverCalc.value = Boolean(apiCfg?.serverCalc);

        const bdHasEdits =
          (snapshot?.version === 3 && hasSheetEdits(snapshot.bdEdits)) ||
          (snapshot?.version === 2 && hasSheetEdits(snapshot.bdEdits));
        const synHasEdits =
          (snapshot?.version === 3 && hasSheetEdits(snapshot.synEdits)) ||
          (snapshot?.version === 2 && hasSheetEdits(snapshot.synEdits));
        const hasBdFull = snapshot?.version === 3 && snapshot.bdFull;
        const hasSynFull = snapshot?.version === 3 && snapshot.synFull;

        if (snapshot?.version === 3) {
          if (snapshot.bdFull) {
            bdRaw.value = snapshot.bdFull;
          } else if (bdHasEdits) {
            await fetchBdFromServer();
            applySheetEdits(bdRaw.value, snapshot.bdEdits);
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
          if (snapshot.bdEdits) {
            await fetchBdFromServer();
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
        }

        gridPreparing.value = true;

        if (!bdRaw.value && preBd?.sheet) {
          await yieldToMain();
          bdSheet.value = preBd.sheet;
          bdSheetBuiltAt = 0;
          gridPreparing.value = false;
          scheduleEngine();
          void loadSheetRaw('bd').then((raw) => {
            if (raw && !bdRaw.value) bdRaw.value = raw;
          });
        } else if (bdRaw.value) {
          bumpBdSheetRevision();
          await applyBdFromRaw(false);
          gridPreparing.value = false;
          scheduleEngine();
        } else {
          await loadBdFromServer({ progressive: chunkedApi.value });
        }

        if (wantSynthesis) {
          synthesisLoading.value = true;
          try {
            await loadSynthesis();
          } finally {
            synthesisLoading.value = false;
          }
        } else {
          void deferHeavy(() => startSynBackgroundPrepare());
        }
      } catch (e) {
        error.value = e.message;
        gridPreparing.value = false;
        console.error('Startup failed:', e);
      }
    }

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
      await applyBdFromRaw(true);
      if (syn) {
        synRaw.value = cloneRawSheet(syn);
        preservedSynRaw = synRaw.value;
        bumpSynSheetRevision();
        await applySynFromRaw(true);
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

    function applySynPatches(patches) {
      if (!patches?.length) return;
      const raw = synRaw.value;
      const grid = synthesisSheet.value;
      if (!raw) return;
      for (const p of patches) {
        upsertRawCell(raw, p.r, p.c, p.v);
        if (grid) {
          const key = `${p.r}:${p.c}`;
          const map = grid.cellMap instanceof Map ? grid.cellMap : null;
          let cell = map?.get(key);
          if (!cell) {
            cell = { r: p.r, c: p.c, v: p.v, mat: true };
            grid.cells.push(cell);
            map?.set(key, cell);
          } else if (!cell.userEdited) {
            cell.v = p.v;
            cell.mat = true;
          }
        }
      }
      externalEditTick.value += 1;
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
      if (sheet === 'BD' && serverCalc.value && chunkedApi.value) {
        void patchSheetCells('bd', [{ r: row, c: col, v: value }])
          .then((res) => applySynPatches(res.synPatches))
          .catch((e) => console.warn('Server Syn recalc:', e));
      }
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
      bdLoading.value = false;
      try {
        await fetchBdFromServer();
        bumpBdSheetRevision();
        await applyBdFromRaw(true);
        synRaw.value = null;
        synthesisSheet.value = null;
        synSheetBuiltAt = -1;
        if (route.value === 'synthesis') {
          await fetchSynFromServer();
          bumpSynSheetRevision();
          await applySynFromRaw(true);
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

    function closeMenu() {
      menuOpen.value = false;
    }

    function toggleMenu() {
      menuOpen.value = !menuOpen.value;
    }

    function navigate(id) {
      closeMenu();
      if (id === route.value) return;
      route.value = id;
      requestAnimationFrame(() => {
        if (id === 'synthesis' && !synthesisSheet.value) {
          void loadSynthesis();
        }
        scheduleEngine();
      });
    }

    function toggleOutline() {
      requestAnimationFrame(() => {
        outlineOnly.value = !outlineOnly.value;
      });
    }

    function deferClick(fn) {
      requestAnimationFrame(() => {
        void deferHeavy(fn);
      });
    }

    async function openMatrix() {
      if (!bdSheet.value) return;
      matrixPendingModel = null;
      matrixState.value = null;
      matrixOpen.value = true;
      try {
        if (!synRaw.value) {
          void startSynBackgroundPrepare();
          try {
            const raw = await loadSheetRaw('synthesis');
            if (raw) synRaw.value = raw;
          } catch {
            /* BD-only matrix still usable */
          }
        }
        const synSource = synRaw.value ?? null;
        matrixState.value = buildMatrixState(bdSheet.value, synSource);
        matrixSessionBefore = {
          bd: cloneRawSheet(bdRaw.value),
          syn: synRaw.value ? cloneRawSheet(synRaw.value) : null,
        };
        matrixSessionDirty = false;
      } catch (e) {
        error.value = e?.message || String(e);
        console.error('Matrix open failed:', e);
        matrixOpen.value = false;
      }
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
        if (!synRaw.value) {
          try {
            const raw = await loadSheetRaw('synthesis');
            if (raw) synRaw.value = raw;
          } catch {
            /* apply BD structure only */
          }
        }
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
        await applyBdFromRaw(false);
        if (result.synRaw) {
          synRaw.value = result.synRaw;
          preservedSynRaw = result.synRaw;
          bumpSynSheetRevision();
          await applySynFromRaw(false);
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
      if (!bd) return;
      matrixPendingModel = bd;
    }

    async function closeMatrix() {
      const pending = matrixPendingModel;
      matrixPendingModel = null;
      if (pending) {
        await applyMatrixFromModel(pending);
      }
      commitMatrixSessionHistory();
      matrixOpen.value = false;
      matrixState.value = null;
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
      isGridPage,
      dirty,
      saveStatus,
      saveError,
      loadedFromLocal,
      saveNow,
      resetFromServerFiles,
      route,
      gridPreparing,
      dataLoadProgress,
      chunkedApi,
      menuOpen,
      toggleMenu,
      closeMenu,
      outlineOnly,
      currentNav,
      session,
      onCellChange,
      navigate,
      toggleOutline,
      deferClick,
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
    };
  },
  template: `
    <div class="app-shell" :class="{ 'menu-open': menuOpen }">
      <AppSidebar
        :current="route"
        :open="menuOpen"
        @navigate="navigate"
        @close="closeMenu"
      />
      <header class="app-topbar">
        <div class="topbar-left">
          <button
            type="button"
            class="burger"
            aria-label="Open menu"
            @click="toggleMenu"
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
          <span class="status" v-if="dataLoadProgress">
            Données {{ dataLoadProgress.sheetId === 'syn' ? 'Synthesis' : 'Database' }}
            {{ dataLoadProgress.pct }}%
          </span>
          <span class="status" v-else-if="isSynthesis && synthesisLoading && !synthesisSheet">Préparation Synthesis…</span>
          <span class="status error-text" v-else-if="error">{{ error }}</span>
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
          <div
            v-if="gridPreparing && isDatabase && !bdSheet"
            class="loading-overlay loading-overlay-subtle"
          >
            Préparation de la grille…
          </div>
          <div
            v-if="isSynthesis && synthesisLoading && !synthesisSheet"
            class="loading-overlay"
          >
            Chargement Synthesis…
          </div>
          <div v-if="error" class="loading-overlay error-text">{{ error }}</div>
          <div v-else-if="showContentOverlay" class="loading-overlay">{{ overlayMessage }}</div>
          <div
            v-else-if="isSynthesis && !synthesisSheet && !synthesisLoading"
            class="loading-overlay error-text"
          >
            Données Synthesis indisponibles
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
