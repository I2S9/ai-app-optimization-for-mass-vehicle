import { createApp, ref, computed, onMounted, onUnmounted, watch } from 'vue';
import BdGrid from './BdGrid.js?v=calc-syn6';
import AppSidebar from './AppSidebar.js?v=calc-syn6';
import EmptyPage from './EmptyPage.js?v=calc-syn6';
import { NAV_ITEMS, DEFAULT_ROUTE } from './navConfig.js?v=calc-syn6';
import { transformBdSheet, transformSynthesisSheet } from './sheetTransform.js?v=calc-syn6';
import { createWorkbookSession } from './workbookSession.js?v=calc-syn6';

const App = {
  components: { BdGrid, AppSidebar, EmptyPage },
  setup() {
    const loading = ref(true);
    const calcLoading = ref(false);
    const error = ref(null);
    const calcError = ref(null);
    const bdSheet = ref(null);
    const synthesisSheet = ref(null);
    const bdRaw = ref(null);
    const synthesisRaw = ref(null);
    const dirty = ref(0);
    const route = ref(DEFAULT_ROUTE);
    const menuOpen = ref(false);
    const outlineOnly = ref(false);
    const session = createWorkbookSession();

    const currentNav = computed(
      () => NAV_ITEMS.find((n) => n.id === route.value) || NAV_ITEMS[0]
    );

    const activeSheet = computed(() => {
      if (route.value === 'synthesis') return synthesisSheet.value;
      if (route.value === 'database') return bdSheet.value;
      return null;
    });

    const activeSheetName = computed(() =>
      route.value === 'synthesis' ? 'SYNTHESIS' : 'BD'
    );

    async function ensureEngine() {
      if (session.ready.value || session.loading.value) return;
      if (!bdRaw.value) return;
      const sheets = [{ name: 'BD', data: bdRaw.value }];
      calcLoading.value = true;
      calcError.value = null;
      try {
        await session.loadSheets(sheets);
        if (session.error.value) calcError.value = session.error.value;
      } finally {
        calcLoading.value = false;
      }
    }

    onMounted(async () => {
      document.title = 'WGHT Dashboard';
      try {
        const bdRes = await fetch('/public/data/bd-sheet.json');
        if (!bdRes.ok) throw new Error(`Failed to load BD data (${bdRes.status})`);
        const rawBd = await bdRes.json();
        bdRaw.value = rawBd;
        bdSheet.value = transformBdSheet(rawBd);

        const synRes = await fetch('/public/data/synthesis-sheet.json');
        if (synRes.ok) {
          const rawSyn = await synRes.json();
          synthesisRaw.value = rawSyn;
          synthesisSheet.value = transformSynthesisSheet(rawSyn);
        }
        await ensureEngine();
      } catch (e) {
        error.value = e.message;
      } finally {
        loading.value = false;
      }
    });

    watch(route, async (id) => {
      if (id === 'database' || id === 'synthesis') await ensureEngine();
    });

    onUnmounted(() => session.destroy());

    function onCellChange() {
      dirty.value += 1;
    }

    function navigate(id) {
      route.value = id;
      menuOpen.value = false;
    }

    function toggleOutline() {
      outlineOnly.value = !outlineOnly.value;
    }

    return {
      loading,
      calcLoading,
      error,
      calcError,
      bdSheet,
      synthesisSheet,
      activeSheet,
      activeSheetName,
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
    <div class="app-shell">
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
          <span v-if="route === 'database'" class="page-title">Database</span>
          <span v-else-if="route === 'synthesis'" class="page-title">Synthesis</span>
          <span v-else class="page-title">{{ currentNav.label }}</span>
          <template v-if="route === 'database'">
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
        <span class="status" v-if="calcLoading">Building calculation engine…</span>
        <span class="status error-text" v-else-if="calcError">{{ calcError }}</span>
        <span class="status" v-else-if="(route === 'database' || route === 'synthesis') && activeSheet && dirty">
          Unsaved changes
        </span>
        <span class="status" v-else-if="session.ready && (route === 'database' || route === 'synthesis')">
          Live formulas
        </span>
      </header>
      <div class="app-body">
        <main class="app-content">
          <div v-if="loading" class="loading-overlay">Loading…</div>
          <div v-else-if="error" class="loading-overlay error-text">{{ error }}</div>
          <div v-else-if="route === 'synthesis' && !synthesisSheet" class="loading-overlay error-text">
            Missing synthesis-sheet.json — run: node tools/export-synthesis-sheet.mjs
          </div>
          <template v-else>
            <BdGrid
              v-if="(route === 'database' && bdSheet) || (route === 'synthesis' && synthesisSheet)"
              :sheet="activeSheet"
              :sheet-name="activeSheetName"
              :session="session"
              :outline-only="outlineOnly && route === 'database'"
              @cell-change="onCellChange"
            />
            <EmptyPage v-else :title="currentNav.label" />
          </template>
        </main>
      </div>
    </div>
  `,
};

createApp(App).mount('#app');
