import { ref } from 'vue';
import { NAV_ITEMS } from './navConfig.js?v=20260520-3';

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
        <strong>WGHT</strong>
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
