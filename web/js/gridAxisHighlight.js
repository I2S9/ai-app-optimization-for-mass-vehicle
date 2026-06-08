/**
 * Excel-style row/column axis highlight (grey band + header emphasis).
 */
import { ref } from 'vue';

export function createGridAxisHighlight(getClipboard = () => null) {
  const axisRow = ref(null);
  const axisCol = ref(null);

  function syncFromCell(row, col) {
    axisRow.value = row;
    axisCol.value = col;
  }

  function onRowNumClick(row, event) {
    if (event.button !== 0 || row == null) return;
    event.preventDefault();
    axisRow.value = row;
    axisCol.value = null;
  }

  function onColHeaderClick(col, event) {
    if (event.button !== 0) return;
    event.preventDefault();
    axisCol.value = col;
    axisRow.value = null;
  }

  function isAxisRow(row) {
    return axisRow.value != null && Number(axisRow.value) === Number(row);
  }

  function isAxisCol(col) {
    return axisCol.value != null && axisCol.value === col;
  }

  function clip() {
    const c = getClipboard();
    return c && typeof c === 'object' ? c : null;
  }

  function isCopiedRow(row) {
    const c = clip();
    return (
      c &&
      c.type === 'row' &&
      c.row != null &&
      Number(c.row) === Number(row)
    );
  }

  function isCopiedCol(col) {
    const c = clip();
    return c && c.type === 'col' && c.col != null && c.col === col;
  }

  return {
    axisRow,
    axisCol,
    syncFromCell,
    onRowNumClick,
    onColHeaderClick,
    isAxisRow,
    isAxisCol,
    isCopiedRow,
    isCopiedCol,
  };
}