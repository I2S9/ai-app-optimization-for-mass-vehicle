import { ref, computed, watch } from 'vue';
import {
  cloneStructure,
  findSection,
  findSubsection,
} from './structureModel.js?v=matrix2';

const DEFAULT_SECTION_COLOR = '#fff2cc';
const DEFAULT_SUBSECTION_COLOR = '#bdd7ee';

function formatSubLabel(name) {
  const t = String(name || '').trim();
  if (!t) return '';
  return t.startsWith('_') ? t.toUpperCase() : `_${t.toUpperCase()}`;
}

export default {
  name: 'MatrixModal',
  props: {
    open: { type: Boolean, default: false },
    state: { type: Object, default: null },
    saving: { type: Boolean, default: false },
  },
  emits: ['close', 'save'],
  setup(props, { emit }) {
    const model = ref(null);
    const selectedIds = ref([]);
    const dragSubId = ref(null);
    const dropTarget = ref(null);
    const editor = ref(null);

    watch(
      () => props.state,
      (s) => {
        if (s?.bd) model.value = cloneStructure(s.bd);
        else model.value = null;
        selectedIds.value = [];
        editor.value = null;
        dropTarget.value = null;
        dragSubId.value = null;
      },
      { immediate: true }
    );

    const selectedSections = computed(() => {
      if (!model.value) return [];
      return selectedIds.value
        .map((id) => findSection(model.value, id))
        .filter(Boolean);
    });

    function toggleSection(id) {
      const idx = selectedIds.value.indexOf(id);
      if (idx >= 0) {
        selectedIds.value = selectedIds.value.filter((x) => x !== id);
        return;
      }
      if (selectedIds.value.length >= 2) {
        selectedIds.value = [selectedIds.value[1], id];
      } else {
        selectedIds.value = [...selectedIds.value, id];
      }
    }

    function isDropTarget(sectionId, index) {
      const t = dropTarget.value;
      return t?.sectionId === sectionId && t?.index === index;
    }

    function onDragStart(subId, e) {
      dragSubId.value = subId;
      if (e?.dataTransfer) {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', subId);
      }
    }

    function onDragEnd() {
      dragSubId.value = null;
      dropTarget.value = null;
    }

    function onSlotDragOver(sectionId, index) {
      if (!dragSubId.value) return;
      dropTarget.value = { sectionId, index };
    }

    function clearDrop(sectionId) {
      if (dropTarget.value?.sectionId === sectionId) dropTarget.value = null;
    }

    function onDropSub(targetSectionId, insertIndex) {
      const from = findSubsection(model.value, dragSubId.value);
      if (!from) return;
      const toSec = findSection(model.value, targetSectionId);
      if (!toSec) return;
      const fromIdx = from.section.subsections.findIndex(
        (s) => s.id === from.subsection.id
      );
      if (fromIdx < 0) return;
      const [item] = from.section.subsections.splice(fromIdx, 1);
      let idx =
        insertIndex == null ? toSec.subsections.length : insertIndex;
      if (from.section.id === toSec.id && fromIdx < idx) idx -= 1;
      toSec.subsections.splice(idx, 0, item);
      dragSubId.value = null;
      dropTarget.value = null;
    }

    function openEditor(payload) {
      editor.value = { ...payload };
    }

    function closeEditor() {
      editor.value = null;
    }

    function openRenameSection(sec) {
      openEditor({
        kind: 'section',
        sectionId: sec.id,
        label: sec.label,
        color: sec.color,
        title: 'Rename section',
        hint: 'Updates the yellow section in Database and Synthesis.',
      });
    }

    function openRenameSubsection(sec, sub) {
      openEditor({
        kind: 'subsection',
        sectionId: sec.id,
        subId: sub.id,
        label: sub.label,
        color: sub.color,
        title: 'Rename sub-section',
        hint: 'Updates the blue band in Database and Synthesis.',
      });
    }

    function openAddSubsection(sec) {
      openEditor({
        kind: 'add-subsection',
        sectionId: sec.id,
        label: '_NEW',
        color: DEFAULT_SUBSECTION_COLOR,
        title: 'Add sub-section',
        hint: 'Creates a new blue band at the end of this section (both sheets on Save).',
      });
    }

    function applyEditor() {
      if (!editor.value || !model.value) return;
      const { kind, sectionId, subId, label, color } = editor.value;
      const val = String(label || '').trim();
      if (!val) return;

      if (kind === 'section') {
        const sec = findSection(model.value, sectionId);
        if (sec) {
          sec.label = val;
          if (color) sec.color = color;
        }
      } else if (kind === 'subsection') {
        const hit = findSubsection(model.value, subId);
        if (hit) {
          hit.subsection.label = formatSubLabel(val) || val;
          if (color) hit.subsection.color = color;
        }
      } else if (kind === 'add-subsection') {
        const sec = findSection(model.value, sectionId);
        if (!sec) return;
        sec.subsections.push({
          id: `sub-${Date.now()}`,
          label: formatSubLabel(val) || '_NEW',
          color: color || DEFAULT_SUBSECTION_COLOR,
          isNew: true,
          startRow: null,
          endRow: null,
        });
      }
      closeEditor();
    }

    function onColorChange(type, id, color) {
      if (!model.value) return;
      if (type === 'section') {
        const sec = findSection(model.value, id);
        if (sec) sec.color = color;
      } else if (type === 'subsection') {
        const hit = findSubsection(model.value, id);
        if (hit) hit.subsection.color = color;
      }
      if (editor.value) {
        if (
          editor.value.kind === 'section' &&
          editor.value.sectionId === id
        ) {
          editor.value.color = color;
        }
        if (
          (editor.value.kind === 'subsection' ||
            editor.value.kind === 'add-subsection') &&
          editor.value.subId === id
        ) {
          editor.value.color = color;
        }
      }
    }

    function close() {
      emit('close');
    }

    function save() {
      if (!model.value) return;
      emit('save', { bd: cloneStructure(model.value) });
    }

    return {
      model,
      selectedIds,
      selectedSections,
      dragSubId,
      dropTarget,
      editor,
      DEFAULT_SECTION_COLOR,
      toggleSection,
      isDropTarget,
      onDragStart,
      onDragEnd,
      onSlotDragOver,
      clearDrop,
      onDropSub,
      openRenameSection,
      openRenameSubsection,
      openAddSubsection,
      applyEditor,
      closeEditor,
      onColorChange,
      close,
      save,
    };
  },
  template: `
    <Teleport to="body">
      <Transition name="matrix-fade">
        <div v-if="open && model" class="matrix-overlay" @click.self="close">
          <div class="matrix-panel" role="dialog" aria-modal="true" aria-label="Structure matrix editor">
            <header class="matrix-header">
              <button type="button" class="matrix-close" aria-label="Close" @click="close">×</button>
              <div class="matrix-title-block">
                <h2 class="matrix-title">Structure matrix</h2>
                <p class="matrix-subtitle">Edits apply to <strong>Database</strong> and <strong>Synthesis</strong> when you save.</p>
              </div>
              <button
                type="button"
                class="matrix-save"
                :disabled="saving"
                @click="save"
              >{{ saving ? 'Saving…' : 'Save' }}</button>
            </header>
            <div class="matrix-body">
              <aside class="matrix-sections">
                <p class="matrix-hint">Yellow sections — select up to 2</p>
                <div
                  v-for="sec in model.sections"
                  :key="sec.id"
                  class="matrix-sec-row"
                  :class="{ selected: selectedIds.includes(sec.id) }"
                >
                  <button
                    type="button"
                    class="matrix-sec-btn"
                    :style="{ background: sec.color }"
                    @click="toggleSection(sec.id)"
                  >
                    <span class="matrix-sec-label">{{ sec.label }}</span>
                  </button>
                  <input
                    type="color"
                    class="matrix-color-input"
                    :value="sec.color"
                    title="Section color"
                    @click.stop
                    @input="onColorChange('section', sec.id, $event.target.value)"
                  />
                  <button
                    type="button"
                    class="matrix-icon-btn"
                    title="Rename section"
                    aria-label="Rename section"
                    @click.stop="openRenameSection(sec)"
                  >✎</button>
                </div>
              </aside>
              <div class="matrix-columns">
                <div
                  v-for="sec in selectedSections"
                  :key="sec.id"
                  class="matrix-col"
                >
                  <div class="matrix-col-head" :style="{ background: sec.color }">
                    <span class="matrix-col-title">{{ sec.label }}</span>
                    <button
                      type="button"
                      class="matrix-icon-btn matrix-icon-btn-light"
                      title="Rename section"
                      @click="openRenameSection(sec)"
                    >✎</button>
                    <button type="button" class="matrix-mini-btn" @click="openAddSubsection(sec)">
                      + Sub-section
                    </button>
                  </div>
                  <ul
                    class="matrix-subs"
                    :class="{ 'is-dragging': !!dragSubId }"
                    @dragleave="clearDrop(sec.id)"
                  >
                    <template v-for="(sub, idx) in sec.subsections" :key="sub.id">
                      <li
                        class="matrix-drop-slot"
                        :class="{ 'is-target': isDropTarget(sec.id, idx) }"
                        @dragenter.prevent="onSlotDragOver(sec.id, idx)"
                        @dragover.prevent="onSlotDragOver(sec.id, idx)"
                        @drop.prevent="onDropSub(sec.id, idx)"
                      >
                        <span class="matrix-drop-line"></span>
                        <span class="matrix-drop-hint">Drop here</span>
                      </li>
                      <li
                        class="matrix-sub"
                        :class="{
                          'is-dragging': dragSubId === sub.id,
                          'is-dimmed': dragSubId && dragSubId !== sub.id
                        }"
                        :style="{ background: sub.color }"
                        draggable="true"
                        @dragstart="onDragStart(sub.id, $event)"
                        @dragend="onDragEnd"
                      >
                        <span class="matrix-drag-handle" title="Drag to reorder" aria-hidden="true">⋮⋮</span>
                        <input
                          type="color"
                          class="matrix-color-input"
                          :value="sub.color"
                          title="Sub-section color"
                          @click.stop
                          @input="onColorChange('subsection', sub.id, $event.target.value)"
                        />
                        <span class="matrix-sub-label">{{ sub.label }}</span>
                        <button
                          type="button"
                          class="matrix-icon-btn matrix-icon-btn-on-color"
                          title="Rename sub-section"
                          @click.stop="openRenameSubsection(sec, sub)"
                        >✎</button>
                      </li>
                    </template>
                    <li
                      class="matrix-drop-slot matrix-drop-slot-end"
                      :class="{ 'is-target': isDropTarget(sec.id, sec.subsections.length) }"
                      @dragenter.prevent="onSlotDragOver(sec.id, sec.subsections.length)"
                      @dragover.prevent="onSlotDragOver(sec.id, sec.subsections.length)"
                      @drop.prevent="onDropSub(sec.id, sec.subsections.length)"
                    >
                      <span class="matrix-drop-line"></span>
                      <span class="matrix-drop-hint">Drop at end</span>
                    </li>
                  </ul>
                  <p v-if="!sec.subsections.length && !dragSubId" class="matrix-empty">
                    No sub-sections yet — use + Sub-section or drag a blue block here.
                  </p>
                </div>
                <div v-if="selectedSections.length === 0" class="matrix-placeholder">
                  <p class="matrix-placeholder-title">Select sections to edit</p>
                  <p>Pick one or two yellow sections on the left. Drag blue sub-sections to reorder or move them between columns. Drop zones appear between items while dragging.</p>
                </div>
                <div v-else-if="selectedSections.length === 1" class="matrix-placeholder matrix-placeholder-col">
                  <p class="matrix-placeholder-title">Add a second section</p>
                  <p>Select another yellow section to compare side by side and drag sub-sections across columns.</p>
                </div>
              </div>
            </div>
          </div>

          <div v-if="editor" class="matrix-editor-backdrop" @click.self="closeEditor">
            <form class="matrix-editor" @submit.prevent="applyEditor">
              <h3 class="matrix-editor-title">{{ editor.title }}</h3>
              <p class="matrix-editor-hint">{{ editor.hint }}</p>
              <label class="matrix-editor-label">
                Name
                <input
                  v-model="editor.label"
                  type="text"
                  class="matrix-editor-input"
                  autocomplete="off"
                  required
                />
              </label>
              <label class="matrix-editor-label matrix-editor-color-row">
                Color
                <input
                  v-model="editor.color"
                  type="color"
                  class="matrix-editor-color"
                />
                <span class="matrix-editor-swatch" :style="{ background: editor.color }"></span>
              </label>
              <div class="matrix-editor-actions">
                <button type="button" class="matrix-editor-cancel" @click="closeEditor">Cancel</button>
                <button type="submit" class="matrix-editor-apply">Apply</button>
              </div>
            </form>
          </div>
        </div>
      </Transition>
    </Teleport>
  `,
};
