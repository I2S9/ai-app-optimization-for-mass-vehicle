import { NAV_ITEMS } from './navConfig.js?v=20260520-edge';
export default {
  name: 'AppSidebar',
  props: {
    current: { type: String, required: true },
    open: { type: Boolean, default: false },
  },
  emits: ['navigate', 'close'],
  setup(props, { emit }) {
    function go(id) {
      emit('navigate', id);
      emit('close');
    }
    return { NAV_ITEMS, go };
  },
  template: `
    <div class="sidebar-backdrop" :class="{ open }" @click="$emit('close')"></div>
    <aside class="app-sidebar" :class="{ open }">
      <div class="sidebar-head">
        <span class="sidebar-menu-label">Menu</span>
        <button type="button" class="sidebar-close" aria-label="Close menu" @click="$emit('close')">
          <svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true">
            <path d="M3 3l10 10M13 3L3 13" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
          </svg>
        </button>
      </div>
      <nav class="sidebar-nav">
        <button
          v-for="item in NAV_ITEMS"
          :key="item.id"
          type="button"
          class="sidebar-link"
          :class="{ active: current === item.id }"
          @click="go(item.id)"
        >
          {{ item.label }}
        </button>
      </nav>
    </aside>
  `,
};
