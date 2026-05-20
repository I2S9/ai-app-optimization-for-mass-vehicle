import { createApp, ref, onMounted } from 'vue';
import BdGrid from './BdGrid.js';

const App = {
  components: { BdGrid },
  setup() {
    const loading = ref(true);
    const error = ref(null);
    const sheet = ref(null);
    const dirty = ref(0);

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

    return { loading, error, sheet, dirty, onCellChange };
  },
  template: `
    <div class="app-shell">
      <header class="app-topbar">
        <h1>WGHT Dashboard</h1>
        <span class="meta">Database (BD)</span>
        <span class="status" v-if="sheet">
          {{ dirty ? 'Unsaved changes' : 'Loaded from workbook reference' }}
        </span>
      </header>
      <main class="app-main">
        <div v-if="loading" class="loading-overlay">Loading Database (BD)…</div>
        <div v-else-if="error" class="loading-overlay" style="color:#b00020">{{ error }}</div>
        <BdGrid v-else-if="sheet" :sheet="sheet" @cell-change="onCellChange" />
      </main>
    </div>
  `,
};

createApp(App).mount('#app');
