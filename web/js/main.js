import { createApp, ref, computed, onMounted, onUnmounted } from 'vue';
import BdGrid from './BdGrid.js?v=syn-perf32';
import SynthesisGrid from './SynthesisGrid.js?v=syn-perf32';
import AppSidebar from './AppSidebar.js?v=syn-perf32';
import EmptyPage from './EmptyPage.js?v=syn-perf32';
import { NAV_ITEMS, DEFAULT_ROUTE } from './navConfig.js?v=syn-perf32';
import { transformBdSheet, transformSynthesisSheet } from './sheetTransform.js?v=syn-perf32';
import { createWorkbookSession } from './workbookSession.js?v=syn-perf32';

const App = {
  components: { BdGrid, SynthesisGrid, AppSidebar, EmptyPage },
  setup() {
    const loading = ref(true);
    const synthesisLoading = ref(false);
    const error = ref(null);
    const bdSheet = ref(null);
    const synthesisSheet = ref(null);
    const bdRaw = ref(null);
    const dirty = ref(0);
    const route = ref(DEFAULT_ROUTE);
    const menuOpen = ref(false);
    const outlineOnly = ref(false);
    const session = createWorkbookSession();
    let engineStarted = false;
    let synthesisLoadPromise = null;

    const currentNav = computed(
      () => NAV_ITEMS.find((n) => n.id === route.value) || NAV_ITEMS[0]
    );

    const isDatabase = computed(() => route.value === 'database');
    const isSynthesis = computed(() => route.value === 'synthesis');
    const isGridPage = computed(
      () => route.value === 'database' || route.value === 'synthesis'
    );

    function scheduleEngine() {
      if (engineStarted || !bdRaw.value || !isGridPage.value) return;
      engineStarted = true;
      const run = () => session.loadSheets([{ name: 'BD', data: bdRaw.value }]);
      if (typeof requestIdleCallback === 'function') {
        requestIdleCallback(run, { timeout: 3000 });
      } else {
        setTimeout(run, 100);
      }
    }

    function loadSynthesis() {
      if (synthesisSheet.value) return Promise.resolve();
      if (synthesisLoadPromise) return synthesisLoadPromise;
      synthesisLoading.value = true;
      synthesisLoadPromise = fetch('/public/data/synthesis-sheet.json')
        .then((res) => {
          if (!res.ok) throw new Error(`Synthesis data (${res.status})`);
          return res.json();
        })
        .then((raw) => {
          synthesisSheet.value = transformSynthesisSheet(raw);
        })
        .catch((e) => {
          error.value = e.message;
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
      try {
        const bdRes = await fetch('/public/data/bd-sheet.json');
        if (!bdRes.ok) throw new Error(`Failed to load BD data (${bdRes.status})`);
        bdRaw.value = await bdRes.json();
        bdSheet.value = transformBdSheet(bdRaw.value);
        if (route.value === 'synthesis') {
          await loadSynthesis();
        }
        scheduleEngine();
      } catch (e) {
        error.value = e.message;
      } finally {
        loading.value = false;
      }
    });

    onUnmounted(() => session.destroy());

    function onCellChange() {
      dirty.value += 1;
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

    return {
      loading,
      synthesisLoading,
      error,
      bdSheet,
      synthesisSheet,
      isDatabase,
      isSynthesis,
      dirty,
      route,
      menuOpen,
      outlineOnly,
      currentNav,
      session,
      onCellChange,
      navigate,
      toggleOutline,
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
            <button type="button" class="icon-btn icon-btn-sm" title="Matrix view" aria-label="Matrix view">
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
        <span class="status" v-if="loading">Loading…</span>
        <span class="status" v-else-if="isSynthesis && synthesisLoading">Loading synthesis…</span>
        <span class="status error-text" v-else-if="error">{{ error }}</span>
        <span class="status" v-else-if="isGridPage && dirty">Unsaved changes</span>
      </header>
      <div class="app-body">
        <main class="app-content">
          <div v-if="loading" class="loading-overlay">Loading database…</div>
          <div v-else-if="isSynthesis && synthesisLoading" class="loading-overlay">Loading synthesis…</div>
          <div v-else-if="isSynthesis && !synthesisSheet && error" class="loading-overlay error-text">{{ error }}</div>
          <div v-else-if="isSynthesis && !synthesisSheet" class="loading-overlay error-text">
            Missing synthesis-sheet.json
          </div>
          <BdGrid
            v-else-if="isDatabase && bdSheet"
            :sheet="bdSheet"
            sheet-name="BD"
            :session="session"
            :outline-only="outlineOnly"
            @cell-change="onCellChange"
          />
          <SynthesisGrid
            v-else-if="isSynthesis && synthesisSheet"
            :sheet="synthesisSheet"
            :session="session"
            :outline-only="outlineOnly"
            @cell-change="onCellChange"
          />
          <EmptyPage v-else :title="currentNav.label" />
        </main>
      </div>
    </div>
  `,
};

createApp(App).mount('#app');
