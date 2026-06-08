/**
 * Undo/redo stack for committed cell edits and bookmark-matrix structure saves.
 */

function cloneSheet(sheet) {
  if (!sheet) return null;
  return JSON.parse(JSON.stringify(sheet));
}

export function createEditHistory({ maxSteps = 200, maxMatrixSteps = 30 } = {}) {
  const undoStack = [];
  const redoStack = [];
  let revision = 0;

  function cloneGapList(list) {
    return Array.isArray(list) ? list.map((g) => ({ ...g })) : [];
  }

  function push(entry) {
    if (entry && entry.type === 'matrix') {
      if (!entry.bdBefore || !entry.bdAfter) return;
      undoStack.push({
        type: 'matrix',
        bdBefore: cloneSheet(entry.bdBefore),
        synBefore: cloneSheet(entry.synBefore),
        bdAfter: cloneSheet(entry.bdAfter),
        synAfter: cloneSheet(entry.synAfter),
      });
      const matrixCount = undoStack.filter((e) => e.type === 'matrix').length;
      if (matrixCount > maxMatrixSteps) {
        const firstMatrix = undoStack.findIndex((e) => e.type === 'matrix');
        if (firstMatrix >= 0) undoStack.splice(firstMatrix, 1);
      }
    } else if (entry && entry.type === 'userGaps') {
      if (!entry.sheet || !entry.before || !entry.after) return;
      undoStack.push({
        type: 'userGaps',
        sheet: entry.sheet,
        before: {
          rowGaps: cloneGapList(entry.before.rowGaps),
          colGaps: cloneGapList(entry.before.colGaps),
        },
        after: {
          rowGaps: cloneGapList(entry.after.rowGaps),
          colGaps: cloneGapList(entry.after.colGaps),
        },
      });
    } else {
      const oldValue = String(entry && entry.oldValue != null ? entry.oldValue : '');
      const newValue = String(entry && entry.newValue != null ? entry.newValue : '');
      if (!entry || !entry.sheet || oldValue === newValue) return;
      undoStack.push({
        type: 'cell',
        sheet: entry.sheet,
        row: entry.row,
        col: entry.col,
        oldValue,
        newValue,
      });
    }
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
