/**
 * Undo/redo stack for committed cell edits (one step per cell commit).
 */

export function createEditHistory({ maxSteps = 200 } = {}) {
  const undoStack = [];
  const redoStack = [];
  let revision = 0;

  function push(entry) {
    const oldValue = String(entry?.oldValue ?? '');
    const newValue = String(entry?.newValue ?? '');
    if (!entry?.sheet || oldValue === newValue) return;
    undoStack.push({
      sheet: entry.sheet,
      row: entry.row,
      col: entry.col,
      oldValue,
      newValue,
    });
    if (undoStack.length > maxSteps) undoStack.shift();
    redoStack.length = 0;
    revision += 1;
  }

  function undo() {
    const entry = undoStack.pop();
    if (!entry) return null;
    redoStack.push(entry);
    revision += 1;
    return entry;
  }

  function redo() {
    const entry = redoStack.pop();
    if (!entry) return null;
    undoStack.push(entry);
    revision += 1;
    return entry;
  }

  function clear() {
    undoStack.length = 0;
    redoStack.length = 0;
    revision += 1;
  }

  return {
    push,
    undo,
    redo,
    clear,
    get canUndo() {
      return undoStack.length > 0;
    },
    get canRedo() {
      return redoStack.length > 0;
    },
    get revision() {
      return revision;
    },
  };
}
