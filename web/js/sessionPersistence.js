/**
 * Local persistence (IndexedDB) until Databricks API is available.
 * Survives F5 on the same machine / browser profile.
 */

const DB_NAME = 'vehicle-mass-platform';
const DB_VERSION = 1;
const STORE = 'snapshots';
const DEFAULT_KEY = 'default';
/** Fallback when IndexedDB is blocked (private mode, policy). Same browser only. */
const LS_SNAPSHOT_KEY = 'vehicle-mass-platform-snapshot-v1';

let dbPromise = null;

function openDb() {
  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  return dbPromise;
}

const PERSIST_VERSION = 2;

/** Cells the user changed — small enough for localStorage + IndexedDB. */
export function extractSheetEdits(raw) {
  if (!raw) return { cells: [], headerRows: {} };
  const cells = (raw.cells || [])
    .filter((c) => c.userEdited)
    .map((c) => ({
      r: c.r,
      c: c.c,
      v: c.v == null ? '' : String(c.v),
      userEdited: true,
    }));
  const headerRows = {};
  for (const [rowKey, cols] of Object.entries(raw.headerRows || {})) {
    for (const [col, cell] of Object.entries(cols)) {
      if (!cell?.userEdited) continue;
      const row = Number(rowKey);
      if (!headerRows[row]) headerRows[row] = {};
      headerRows[row][col] = {
        v: cell.v == null ? '' : String(cell.v),
        userEdited: true,
      };
    }
  }
  return { cells, headerRows };
}

export function applySheetEdits(raw, edits) {
  if (!raw || !edits) return;
  for (const cell of edits.cells || []) {
    upsertRawCell(raw, cell.r, cell.c, cell.v);
  }
  for (const [rowKey, cols] of Object.entries(edits.headerRows || {})) {
    const row = Number(rowKey);
    for (const [col, cell] of Object.entries(cols)) {
      upsertRawCell(raw, row, col, cell.v);
    }
  }
}

export function buildPersistRecord({ bd, syn, revision }) {
  const bdEdits = extractSheetEdits(bd);
  const synEdits = extractSheetEdits(syn);
  const hasBd = bdEdits.cells.length || Object.keys(bdEdits.headerRows).length;
  const hasSyn = synEdits.cells.length || Object.keys(synEdits.headerRows).length;
  if (!hasBd && !hasSyn) return null;
  return {
    version: PERSIST_VERSION,
    revision: revision ?? 0,
    savedAt: new Date().toISOString(),
    bdEdits: hasBd ? bdEdits : null,
    synEdits: hasSyn ? synEdits : null,
  };
}

/** @returns {boolean} true if legacy full snapshot was applied */
export function applyPersistRecord(bdRaw, synRaw, record) {
  if (!record) return false;
  if (record.version === PERSIST_VERSION) {
    if (record.bdEdits) applySheetEdits(bdRaw, record.bdEdits);
    if (record.synEdits && synRaw) applySheetEdits(synRaw, record.synEdits);
    return false;
  }
  if (record.bd) {
    Object.assign(bdRaw, record.bd);
  }
  if (record.syn && synRaw) {
    Object.assign(synRaw, record.syn);
  }
  return Boolean(record.bd || record.syn);
}

/** Update or insert a cell on export-shaped sheet JSON. */
export function upsertRawCell(raw, row, col, value) {
  if (!raw) return;
  const r = Number(row);
  const c = String(col);
  const strVal = value == null ? '' : String(value);

  let cell = (raw.cells || []).find((x) => x.r === r && x.c === c);
  if (!cell) {
    cell = { r, c, v: strVal, userEdited: true };
    if (!raw.cells) raw.cells = [];
    raw.cells.push(cell);
  } else {
    cell.v = strVal;
    cell.userEdited = true;
    delete cell.f;
  }

  const rowHdr = raw.headerRows?.[String(r)] ?? raw.headerRows?.[r];
  if (rowHdr?.[c]) {
    rowHdr[c].v = strVal;
    rowHdr[c].userEdited = true;
    delete rowHdr[c].f;
  }
}

export async function loadLocalSnapshot(
  key = DEFAULT_KEY,
  { timeoutMs = 8000 } = {}
) {
  const load = async () => {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).get(key);
      req.onsuccess = () => resolve(req.result ?? null);
      req.onerror = () => reject(req.error);
    });
  };

  try {
    if (!timeoutMs) return await load();
    let timer;
    const timeout = new Promise((resolve) => {
      timer = setTimeout(() => resolve(null), timeoutMs);
    });
    const result = await Promise.race([load(), timeout]);
    clearTimeout(timer);
    if (result === null) {
      console.warn(
        'loadLocalSnapshot: timeout or empty — trying localStorage fallback'
      );
      return loadLocalSnapshotFromStorage();
    }
    return result;
  } catch (e) {
    console.warn('loadLocalSnapshot failed:', e);
    return loadLocalSnapshotFromStorage();
  }
}

function loadLocalSnapshotFromStorage() {
  try {
    const raw = localStorage.getItem(LS_SNAPSHOT_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    console.warn('localStorage snapshot load failed:', e);
    return null;
  }
}

function mirrorSnapshotToStorage(record) {
  try {
    localStorage.setItem(LS_SNAPSHOT_KEY, JSON.stringify(record));
  } catch (e) {
    console.warn('localStorage snapshot mirror failed:', e);
  }
}

export async function saveLocalSnapshot(payload, key = DEFAULT_KEY) {
  const record = buildPersistRecord(payload);
  if (!record) return null;
  mirrorSnapshotToStorage(record);
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    const req = tx.objectStore(STORE).put(record, key);
    req.onsuccess = () => resolve(record.savedAt);
    req.onerror = () => reject(req.error);
    tx.onerror = () => reject(tx.error);
  });
}

export async function clearLocalSnapshot(key = DEFAULT_KEY) {
  try {
    localStorage.removeItem(LS_SNAPSHOT_KEY);
  } catch {
    /* ignore */
  }
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export function createAutoSave(getPayload, { debounceMs = 1500, onStatus } = {}) {
  let timer = null;
  let saving = false;
  let pending = false;

  async function flush() {
    const data = getPayload();
    if (!data?.bd && !data?.syn) return;
    if (saving) {
      pending = true;
      return;
    }
    saving = true;
    onStatus?.('saving');
    try {
      await saveLocalSnapshot(data);
      onStatus?.('saved');
    } catch (e) {
      console.error('Auto-save failed:', e);
      onStatus?.('error', e);
    } finally {
      saving = false;
      if (pending) {
        pending = false;
        void flush();
      }
    }
  }

  function schedule() {
    onStatus?.('dirty');
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => {
      timer = null;
      void flush();
    }, debounceMs);
  }

  function saveNow() {
    if (timer) {
      clearTimeout(timer);
      timer = null;
    }
    return flush();
  }

  function destroy() {
    if (timer) clearTimeout(timer);
    timer = null;
  }

  return { schedule, saveNow, destroy };
}
