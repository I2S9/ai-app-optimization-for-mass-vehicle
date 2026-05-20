import { createApp, ref, computed, onMounted } from 'vue';
import BdGrid from './BdGrid.js?v=20260520-4';
import AppSidebar from './AppSidebar.js?v=20260520-4';
import EmptyPage from './EmptyPage.js?v=20260520-4';
import { NAV_ITEMS, DEFAULT_ROUTE } from './navConfig.js?v=20260520-4';

const App = {
  components: { BdGrid, AppSidebar, EmptyPage },
  setup() {
    const loading = ref(true);
    const error = ref(null);
    const sheet = ref(null);
    const dirty = ref(0);
    const route = ref(DEFAULT_ROUTE);
    const menuOpen = ref(false);

    const currentNav = computed(
      () => NAV_ITEMS.find((n) => n.id === route.value) || NAV_ITEMS[0]
    );

    onMounted(async () => {
      try {
        const res = await fetch('/public/data/bd-sheet.json');
        if (!res.ok) throw new Error(`Failed to load BD data (${res.status})`);
        sheet.value = await res.json();
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
    }

    return {
      loading,
      error,
      sheet,
      dirty,
      route,
      menuOpen,
      currentNav,
      NAV_ITEMS,
      onCellChange,
      navigate,
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
        <button
          type="button"
          class="burger"
          aria-label="Open menu"
          @click="menuOpen = !menuOpen"
        >
          <span></span><span></span><span></span>
        </button>
        <h1>WGHT Dashboard</h1>
        <span class="meta">{{ currentNav.label }}</span>
        <span class="status" v-if="route === 'database' && sheet">
          {{ dirty ? 'Unsaved changes' : '' }}
        </span>
      </header>
      <div class="app-body">
        <main class="app-content">
          <div v-if="loading" class="loading-overlay">Loading…</div>
          <div v-else-if="error" class="loading-overlay" style="color:#b00020">{{ error }}</div>
          <template v-else>
            <BdGrid
              v-if="route === 'database' && sheet"
              :sheet="sheet"
              @cell-change="onCellChange"
            />
            <EmptyPage
              v-else
              :title="currentNav.label"
            />
          </template>
        </main>
      </div>
    </div>
  `,
};

createApp(App).mount('#app');
