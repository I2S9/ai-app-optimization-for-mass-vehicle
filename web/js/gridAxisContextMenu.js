/**
 * Right-click context menu for row/column axis headers (copy, paste, delete).
 */
import { ref, computed } from 'vue';

export function createGridAxisContextMenu({
  emit,
  getClipboard,
  sheetName,
  canOpenRow = () => true,
  canOpenCol = () => true,
}) {
  const ctxMenu = ref({
    visible: false,
    x: 0,
    y: 0,
    axis: null,
    row: null,
    col: null,
    displayRow: null,
    displayCol: null,
  });

  const canPaste = computed(() => {
    const clip = getClipboard();
    const menu = ctxMenu.value;
    if (!clip || !menu.visible || !menu.axis) return false;
    if (clip.sheet !== sheetName) return false;
    return clip.type === menu.axis;
  });

  function selectAxis(axis, row, col) {
    emit('axis-select', {
      sheet: sheetName,
      row: axis === 'row' && row != null ? row : null,
      col: axis === 'col' && col ? col : null,
    });
  }

  function closeCtxMenu() {
    if (ctxMenu.value.visible) {
      ctxMenu.value = {
        visible: false,
        x: 0,
        y: 0,
        axis: null,
        row: null,
        col: null,
        displayRow: null,
        displayCol: null,
      };
    }
  }

  function openMenu(axis, data, event) {
    if (axis === 'row' && !canOpenRow(data)) return;
    if (axis === 'col' && !canOpenCol(data)) return;
    event.preventDefault();
    selectAxis(axis, data.row, data.col);
    ctxMenu.value = {
      visible: true,
      x: event.clientX,
      y: event.clientY,
      axis,
      row: data.row != null ? data.row : null,
      col: data.col != null ? data.col : null,
      displayRow: data.displayRow != null ? data.displayRow : null,
      displayCol: data.displayCol != null ? data.displayCol : null,
    };
  }

  function onRowContextMenu(data, event) {
    openMenu('row', data, event);
  }

  function onColContextMenu(col, displayCol, event) {
    openMenu('col', { col, displayCol: displayCol != null ? displayCol : col }, event);
  }

  function onCtxCopy() {
    closeCtxMenu();
    emit('axis-copy');
  }

  function onCtxPaste() {
    if (!canPaste.value) return;
    closeCtxMenu();
    emit('axis-paste');
  }

  function onCtxDelete() {
    const menu = ctxMenu.value;
    closeCtxMenu();
    if (menu.axis === 'row' && menu.row != null) {
      emit('row-delete', { sheet: sheetName, excelRow: menu.row });
    } else if (menu.axis === 'col' && menu.col) {
      emit('column-delete', { sheet: sheetName, col: menu.col });
    }
  }

  const deleteLabel = computed(() => {
    const menu = ctxMenu.value;
    if (menu.axis === 'col' && menu.displayCol) {
      return `Supprimer la colonne ${menu.displayCol}`;
    }
    if (menu.axis === 'row' && menu.displayRow != null) {
      return `Supprimer la ligne ${menu.displayRow}`;
    }
    return 'Supprimer';
  });

  return {
    ctxMenu,
    canPaste,
    deleteLabel,
    closeCtxMenu,
    onRowContextMenu,
    onColContextMenu,
    onCtxCopy,
    onCtxPaste,
    onCtxDelete,
  };
}
