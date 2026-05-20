import { createApp, ref, computed, onMounted } from 'vue';
import BdGrid from './BdGrid.js?v=20260520-18';
import AppSidebar from './AppSidebar.js?v=20260520-18';
import EmptyPage from './EmptyPage.js?v=20260520-18';
import { NAV_ITEMS, DEFAULT_ROUTE } from './navConfig.js?v=20260520-18';
import { transformBdSheet } from './sheetTransform.js?v=20260520-18';

const App = {
  components: { BdGrid, AppSidebar, EmptyPage },
  setup() {
    const loading = ref(true);
    const error = ref(null);
    const sheet = ref(null);
    const dirty = ref(0);
    const route = ref(DEFAULT_ROUTE);
    const menuOpen = ref(false);
    const outlineOnly = ref(false);

    const currentNav = computed(
      () => NAV_ITEMS.find((n) => n.id === route.value) || NAV_ITEMS[0]
    );

    onMounted(async () => {
      document.title = 'WGHT Dashboard';
      try {
        const res = await fetch('/public/data/bd-sheet.json');
        if (!res.ok) throw new Error(`Failed to load BD data (${res.status})`);
        sheet.value = transformBdSheet(await res.json());
      } catch (e) {
        error.value = e.message;
      } finally {
        loading.value = false;
      }
    });

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
      error,
      sheet,
      dirty,
      route,
      menuOpen,
      outlineOnly,
      currentNav,
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
              title="Show sections and sub-sections only"
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
        <span class="status" v-if="route === 'database' && sheet && dirty">
          Unsaved changes
        </span>
      </header>
      <div class="app-body">
        <main class="app-content">
          <div v-if="loading" class="loading-overlay">Loading…</div>
          <div v-else-if="error" class="loading-overlay error-text">{{ error }}</div>
          <template v-else>
            <BdGrid
              v-if="route === 'database' && sheet"
              :sheet="sheet"
              :outline-only="outlineOnly"
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
