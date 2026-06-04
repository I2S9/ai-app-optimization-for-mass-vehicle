import { ref, watch } from 'vue';
import { NAV_ITEMS } from './navConfig.js?v=syn-perf2';
export default {
  name: 'AppSidebar',
  props: {
    current: { type: String, required: true },
    open: { type: Boolean, default: false },
  },
  emits: ['navigate', 'close'],
  setup(props, { emit }) {
    // Track which collapsible groups are expanded.
    const expanded = ref({});

    function isChildActive(item) {
      return !!(item.children && item.children.some((c) => c.id === props.current));
    }

    // Auto-expand a group when one of its children becomes the active route.
    watch(
      () => props.current,
      () => {
        NAV_ITEMS.forEach((item) => {
          if (isChildActive(item)) expanded.value[item.id] = true;
        });
      },
      { immediate: true }
    );

    function toggle(id) {
      expanded.value[id] = !expanded.value[id];
    }

    function go(id, event) {
      if (event && event.currentTarget && typeof event.currentTarget.blur === 'function') {
        event.currentTarget.blur();
      }
      emit('navigate', id);
    }
    return { NAV_ITEMS, expanded, isChildActive, toggle, go };
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
        <template v-for="item in NAV_ITEMS" :key="item.id">
          <template v-if="item.children">
            <button
              type="button"
              class="sidebar-link sidebar-group-toggle"
              :class="{ active: current === item.id || isChildActive(item) }"
              :aria-expanded="!!expanded[item.id]"
              @click="toggle(item.id)"
            >
              <span>{{ item.label }}</span>
              <svg
                class="sidebar-caret"
                :class="{ open: expanded[item.id] }"
                viewBox="0 0 16 16"
                width="12"
                height="12"
                aria-hidden="true"
              >
                <path d="M5 6l3 3 3-3" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"/>
              </svg>
            </button>
            <div v-show="expanded[item.id]" class="sidebar-subnav">
              <button
                v-for="child in item.children"
                :key="child.id"
                type="button"
                class="sidebar-link sidebar-sublink"
                :class="{ active: current === child.id }"
                @click="go(child.id, $event)"
              >
                {{ child.label }}
              </button>
            </div>
          </template>
          <button
            v-else
            type="button"
            class="sidebar-link"
            :class="{ active: current === item.id }"
            @click="go(item.id, $event)"
          >
            {{ item.label }}
          </button>
        </template>
      </nav>
    </aside>
  `,
};
