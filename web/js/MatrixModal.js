import { ref, computed, watch, nextTick } from 'vue';
import {
  cloneStructure,
  findSection,
  findSubsection,
} from './structureModel.js?v=matrix5';

const DEFAULT_SECTION_COLOR = '#ffff00';
const DEFAULT_SUBSECTION_COLOR = '#00b0f0';

function formatSubLabel(name) {
  const t = String(name || '').trim();
  if (!t) return '';
  return t.startsWith('_') ? t.toUpperCase() : `_${t.toUpperCase()}`;
}

function subSortIndex(model, subId) {
  let i = 0;
  for (const sec of model.sections) {
    for (const sub of sec.subsections) {
      if (sub.id === subId) return i;
      i++;
    }
  }
  return -1;
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
    const selectedSubIds = ref([]);
    const subAnchorId = ref(null);
    const dragSubId = ref(null);
    const dragSubIds = ref([]);
    const dropTarget = ref(null);
    const editing = ref(null);
    const confirmDelete = ref(null);
    const colorInputRef = ref(null);
    const colorPick = ref(null);

    watch(
      () => props.state,
      (s) => {
        if (s?.bd) model.value = cloneStructure(s.bd);
        else model.value = null;
        selectedIds.value = [];
        selectedSubIds.value = [];
        subAnchorId.value = null;
        editing.value = null;
        confirmDelete.value = null;
        dropTarget.value = null;
        dragSubId.value = null;
        dragSubIds.value = [];
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

    function isEditing(kind, sectionId, subId = null) {
      const e = editing.value;
      if (!e || e.kind !== kind) return false;
      if (e.sectionId !== sectionId) return false;
      if (kind === 'subsection') return e.subId === subId;
      return true;
    }

    function startEditSection(sec) {
      editing.value = {
        kind: 'section',
        sectionId: sec.id,
        value: sec.label,
      };
      focusEditInput();
    }

    function startEditSubsection(sec, sub) {
      editing.value = {
        kind: 'subsection',
        sectionId: sec.id,
        subId: sub.id,
        value: sub.label,
      };
      focusEditInput();
    }

    function focusEditInput() {
      nextTick(() => {
        const el = document.querySelector('.matrix-inline-input');
        if (el) {
          el.focus();
          el.select();
        }
      });
    }

    function commitEdit() {
      const e = editing.value;
      if (!e || !model.value) return;
      const val = String(e.value || '').trim();
      editing.value = null;
      if (!val) return;

      if (e.kind === 'section') {
        const sec = findSection(model.value, e.sectionId);
        if (sec) sec.label = val;
      } else if (e.kind === 'subsection') {
        const hit = findSubsection(model.value, e.subId);
        if (hit) hit.subsection.label = formatSubLabel(val) || val;
      }
    }

    function cancelEdit() {
      editing.value = null;
    }

    function onEditKeydown(ev) {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        commitEdit();
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        cancelEdit();
      }
    }

    function isSubSelected(subId) {
      return selectedSubIds.value.includes(subId);
    }

    function clearSubSelection() {
      selectedSubIds.value = [];
      subAnchorId.value = null;
    }

    function toggleSubSelect(subId) {
      const idx = selectedSubIds.value.indexOf(subId);
      if (idx >= 0) {
        selectedSubIds.value = selectedSubIds.value.filter((id) => id !== subId);
      } else {
        selectedSubIds.value = [...selectedSubIds.value, subId];
      }
      subAnchorId.value = subId;
    }

    function rangeSelectSubs(sec, subId) {
      const anchor = subAnchorId.value;
      if (!anchor) {
        toggleSubSelect(subId);
        return;
      }
      const ids = sec.subsections.map((s) => s.id);
      const a = ids.indexOf(anchor);
      const b = ids.indexOf(subId);
      if (a < 0 || b < 0) {
        toggleSubSelect(subId);
        return;
      }
      const lo = Math.min(a, b);
      const hi = Math.max(a, b);
      const range = ids.slice(lo, hi + 1);
      const merged = new Set([...selectedSubIds.value, ...range]);
      selectedSubIds.value = [...merged];
      subAnchorId.value = subId;
    }

    function onSubClick(sec, sub, e) {
      if (e.target.closest('.matrix-icon-btn, .matrix-delete-btn')) return;
      if (e.ctrlKey || e.metaKey) {
        e.preventDefault();
        toggleSubSelect(sub.id);
        return;
      }
      if (e.shiftKey) {
        e.preventDefault();
        rangeSelectSubs(sec, sub.id);
        return;
      }
      if (e.target.closest('.matrix-sub-label, .matrix-inline-input')) {
        startEditSubsection(sec, sub);
      }
    }

    function resolveDragIds(subId) {
      if (
        selectedSubIds.value.length > 0 &&
        selectedSubIds.value.includes(subId)
      ) {
        return [...selectedSubIds.value];
      }
      return [subId];
    }

    function isDropTarget(sectionId, index) {
      const t = dropTarget.value;
      return t?.sectionId === sectionId && t?.index === index;
    }

    function onDragStart(subId, e) {
      const ids = resolveDragIds(subId);
      dragSubIds.value = ids;
      dragSubId.value = subId;
      if (e?.dataTransfer) {
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', ids.join(','));
      }
    }

    function onDragEnd() {
      dragSubId.value = null;
      dragSubIds.value = [];
      dropTarget.value = null;
    }

    function onSlotDragOver(sectionId, index) {
      if (!dragSubIds.value.length && !dragSubId.value) return;
      dropTarget.value = { sectionId, index };
    }

    function clearDrop(sectionId) {
      if (dropTarget.value?.sectionId === sectionId) dropTarget.value = null;
    }

    function onDropSub(targetSectionId, insertIndex) {
      const ids =
        dragSubIds.value.length > 0
          ? dragSubIds.value
          : dragSubId.value
            ? [dragSubId.value]
            : [];
      if (!ids.length || !model.value) return;

      const items = ids
        .map((id) => {
          const hit = findSubsection(model.value, id);
          if (!hit) return null;
          return {
            subsection: hit.subsection,
            order: subSortIndex(model.value, id),
            sourceSectionId: hit.section.id,
            sourceIndex: hit.section.subsections.findIndex(
              (s) => s.id === id
            ),
          };
        })
        .filter(Boolean)
        .sort((a, b) => a.order - b.order);

      const toSec = findSection(model.value, targetSectionId);
      if (!toSec) return;

      let idx =
        insertIndex == null ? toSec.subsections.length : insertIndex;
      let adjust = 0;
      for (const item of items) {
        if (
          item.sourceSectionId === toSec.id &&
          item.sourceIndex >= 0 &&
          item.sourceIndex < idx
        ) {
          adjust++;
        }
      }
      idx -= adjust;

      for (const sec of model.value.sections) {
        for (let i = sec.subsections.length - 1; i >= 0; i--) {
          if (ids.includes(sec.subsections[i].id)) {
            sec.subsections.splice(i, 1);
          }
        }
      }

      for (const item of items) {
        toSec.subsections.splice(idx, 0, item.subsection);
        idx++;
      }

      dragSubId.value = null;
      dragSubIds.value = [];
      dropTarget.value = null;
    }

    function askDeleteSection(sec) {
      confirmDelete.value = {
        kind: 'section',
        sectionId: sec.id,
        name: sec.label,
      };
    }

    function askDeleteSubsection(sec, sub) {
      confirmDelete.value = {
        kind: 'subsection',
        sectionId: sec.id,
        subId: sub.id,
        name: sub.label,
      };
    }

    function cancelDelete() {
      confirmDelete.value = null;
    }

    function applyDelete() {
      const d = confirmDelete.value;
      if (!d || !model.value) return;

      if (d.kind === 'section') {
        model.value.sections = model.value.sections.filter(
          (s) => s.id !== d.sectionId
        );
        selectedIds.value = selectedIds.value.filter((id) => id !== d.sectionId);
        selectedSubIds.value = selectedSubIds.value.filter((subId) => {
          const hit = findSubsection(model.value, subId);
          return hit?.section.id !== d.sectionId;
        });
      } else if (d.kind === 'subsection') {
        const sec = findSection(model.value, d.sectionId);
        if (sec) {
          sec.subsections = sec.subsections.filter((s) => s.id !== d.subId);
        }
        selectedSubIds.value = selectedSubIds.value.filter(
          (id) => id !== d.subId
        );
      }
      confirmDelete.value = null;
    }

    function pickColor(type, sectionId, subId = null) {
      let ids = [];
      let current = DEFAULT_SECTION_COLOR;

      if (type === 'section') {
        const sec = findSection(model.value, sectionId);
        if (!sec) return;
        ids = [sectionId];
        current = sec.color || DEFAULT_SECTION_COLOR;
      } else {
        if (selectedSubIds.value.length > 0) {
          ids = [...selectedSubIds.value];
          const first = findSubsection(model.value, ids[0]);
          current = first?.subsection.color || DEFAULT_SUBSECTION_COLOR;
        } else if (subId) {
          ids = [subId];
          const hit = findSubsection(model.value, subId);
          current = hit?.subsection.color || DEFAULT_SUBSECTION_COLOR;
        }
      }

      colorPick.value = { type, ids, sectionId };
      nextTick(() => {
        const input = colorInputRef.value;
        if (!input) return;
        input.value = current;
        input.click();
      });
    }

    function onColorPicked(e) {
      const color = e.target?.value;
      const pick = colorPick.value;
      if (!color || !pick || !model.value) return;

      if (pick.type === 'section') {
        for (const id of pick.ids) {
          const sec = findSection(model.value, id);
          if (sec) sec.color = color;
        }
      } else {
        for (const id of pick.ids) {
          const hit = findSubsection(model.value, id);
          if (hit) hit.subsection.color = color;
        }
      }
      colorPick.value = null;
    }

    function addSubsection(sec) {
      const id = `sub-${Date.now()}`;
      sec.subsections.push({
        id,
        label: '_NEW',
        color: DEFAULT_SUBSECTION_COLOR,
        isNew: true,
        startRow: null,
        endRow: null,
      });
      selectedSubIds.value = [id];
      subAnchorId.value = id;
      startEditSubsection(sec, sec.subsections[sec.subsections.length - 1]);
      scrollToSub(id);
    }

    function scrollToSub(subId) {
      nextTick(() => {
        const el = document.querySelector(`[data-matrix-sub="${subId}"]`);
        el?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      });
    }

    function sectionStyle(sec) {
      return { background: sec.color || DEFAULT_SECTION_COLOR };
    }

    function subStyle(sub) {
      return { background: sub.color || DEFAULT_SUBSECTION_COLOR };
    }

    function close() {
      emit('close');
    }

    function save() {
      if (!model.value) return;
      if (model.value.sections.length < 2) {
        window.alert(
          'Enregistrement annulé : il doit rester au moins deux sections dans la structure.'
        );
        return;
      }
      emit('save', { bd: cloneStructure(model.value) });
    }

    return {
      model,
      selectedIds,
      selectedSections,
      selectedSubIds,
      dragSubId,
      dropTarget,
      editing,
      confirmDelete,
      colorInputRef,
      DEFAULT_SECTION_COLOR,
      DEFAULT_SUBSECTION_COLOR,
      toggleSection,
      isEditing,
      startEditSection,
      startEditSubsection,
      commitEdit,
      cancelEdit,
      onEditKeydown,
      isSubSelected,
      onSubClick,
      isDropTarget,
      onDragStart,
      onDragEnd,
      onSlotDragOver,
      clearDrop,
      onDropSub,
      askDeleteSection,
      askDeleteSubsection,
      cancelDelete,
      applyDelete,
      pickColor,
      onColorPicked,
      addSubsection,
      sectionStyle,
      subStyle,
      close,
      save,
    };
  },
  template: `
    <Teleport to="body">
      <input
        ref="colorInputRef"
        type="color"
        class="matrix-color-hidden"
        tabindex="-1"
        aria-hidden="true"
        @input="onColorPicked"
      />
      <Transition name="matrix-fade">
        <div v-if="open && model" class="matrix-overlay" @click.self="close">
          <div class="matrix-panel" role="dialog" aria-modal="true" aria-label="Structure matrix editor">
            <header class="matrix-header">
              <button type="button" class="matrix-close" aria-label="Close" @click="close">×</button>
              <div class="matrix-title-block">
                <h2 class="matrix-title">Structure matrix</h2>
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
                <div
                  v-for="sec in model.sections"
                  :key="sec.id"
                  class="matrix-sec-row"
                  :class="{ selected: selectedIds.includes(sec.id) }"
                >
                  <div
                    class="matrix-sec-card"
                    :style="sectionStyle(sec)"
                    @click="toggleSection(sec.id)"
                  >
                    <input
                      v-if="isEditing('section', sec.id)"
                      v-model="editing.value"
                      type="text"
                      class="matrix-inline-input"
                      @click.stop
                      @keydown="onEditKeydown"
                      @blur="commitEdit"
                    />
                    <span
                      v-else
                      class="matrix-sec-label"
                      @click.stop="startEditSection(sec)"
                    >{{ sec.label }}</span>
                    <button
                      type="button"
                      class="matrix-icon-btn"
                      title="Couleur"
                      aria-label="Changer la couleur"
                      @click.stop="pickColor('section', sec.id)"
                    >✎</button>
                    <button
                      type="button"
                      class="matrix-delete-btn"
                      title="Supprimer la section"
                      aria-label="Supprimer la section"
                      @click.stop="askDeleteSection(sec)"
                    >×</button>
                  </div>
                </div>
              </aside>
              <div class="matrix-columns">
                <div
                  v-for="sec in selectedSections"
                  :key="sec.id"
                  class="matrix-col"
                >
                  <div class="matrix-col-head" :style="sectionStyle(sec)">
                    <input
                      v-if="isEditing('section', sec.id)"
                      v-model="editing.value"
                      type="text"
                      class="matrix-inline-input matrix-inline-input-head"
                      @click.stop
                      @keydown="onEditKeydown"
                      @blur="commitEdit"
                    />
                    <span
                      v-else
                      class="matrix-col-title"
                      @click.stop="startEditSection(sec)"
                    >{{ sec.label }}</span>
                    <button
                      type="button"
                      class="matrix-icon-btn matrix-icon-btn-light"
                      title="Couleur"
                      aria-label="Changer la couleur"
                      @click.stop="pickColor('section', sec.id)"
                    >✎</button>
                    <button
                      type="button"
                      class="matrix-icon-btn matrix-icon-btn-light"
                      title="Ajouter une sous-section"
                      aria-label="Ajouter une sous-section"
                      @click.stop="addSubsection(sec)"
                    >+</button>
                  </div>
                  <ul
                    class="matrix-subs"
                    :class="{ 'is-dragging': dragSubId }"
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
                        :data-matrix-sub="sub.id"
                        class="matrix-sub"
                        :class="{
                          'is-dragging': dragSubId === sub.id,
                          'is-dimmed': dragSubId && dragSubId !== sub.id,
                          'is-selected': isSubSelected(sub.id)
                        }"
                        :style="subStyle(sub)"
                        @click="onSubClick(sec, sub, $event)"
                      >
                        <span
                          class="matrix-drag-handle"
                          title="Glisser pour réordonner"
                          draggable="true"
                          @dragstart="onDragStart(sub.id, $event)"
                          @dragend="onDragEnd"
                        >⋮⋮</span>
                        <input
                          v-if="isEditing('subsection', sec.id, sub.id)"
                          v-model="editing.value"
                          type="text"
                          class="matrix-inline-input"
                          @click.stop
                          @keydown="onEditKeydown"
                          @blur="commitEdit"
                        />
                        <span
                          v-else
                          class="matrix-sub-label"
                        >{{ sub.label }}</span>
                        <button
                          type="button"
                          class="matrix-icon-btn matrix-icon-btn-on-color"
                          title="Couleur"
                          aria-label="Changer la couleur"
                          @click.stop="pickColor('subsection', sec.id, sub.id)"
                        >✎</button>
                        <button
                          type="button"
                          class="matrix-delete-btn matrix-delete-btn-on-color"
                          title="Supprimer la sous-section"
                          aria-label="Supprimer la sous-section"
                          @click.stop="askDeleteSubsection(sec, sub)"
                        >×</button>
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
                </div>
              </div>
            </div>
          </div>

          <div v-if="confirmDelete" class="matrix-confirm-backdrop" @click.self="cancelDelete">
            <div class="matrix-confirm" role="alertdialog" aria-modal="true">
              <p class="matrix-confirm-text">
                Supprimer <strong>{{ confirmDelete.name }}</strong> ?
                Cette action sera appliquée à l'enregistrement.
              </p>
              <div class="matrix-confirm-actions">
                <button type="button" class="matrix-editor-cancel" @click="cancelDelete">Annuler</button>
                <button type="button" class="matrix-confirm-delete" @click="applyDelete">Supprimer</button>
              </div>
            </div>
          </div>
        </div>
      </Transition>
    </Teleport>
  `,
};
