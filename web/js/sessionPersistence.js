/**
 * Local persistence (IndexedDB) + optional Databricks session snapshots.
 * Survives F5 on the same machine / browser profile.
 */
import { buildCellMap } from './bdStore.js';
import { probeSheetApi, saveSession } from './sheetDataApi.js';

const DB_NAME = 'vehicle-mass-platform';
const DB_VERSION = 2;
/**
 * Bump when the grid transform output changes (e.g. outline/bookmark structure)
 * so cached transforms (precomputed pack + IndexedDB L1) are invalidated even
 * though the raw source data is unchanged.
 */
const TRANSFORM_VERSION = 'tv5-identical-structure';
/**
 * Bump when the *structure* logic changes (section/sub-section detection, canonical
 * labels, outline rows). A persisted `bdFull` snapshot freezes the whole BD sheet;
 * if it was produced by an older schema it must NOT override the fresh base data +
 * current code, otherwise old/duplicated structure keeps coming back. On mismatch the
 * loader discards `bdFull` and rebuilds from the project JSON (cell edits still apply).
 */
export const STRUCTURE_SCHEMA_VERSION = 'ss4-identical-structure';
const STORE = 'snapshots';
const STORE_TRANSFORMS = 'transforms';
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
        if (!db.objectStoreNames.contains(STORE_TRANSFORMS)) {
          db.createObjectStore(STORE_TRANSFORMS);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }
  return dbPromise;
}

const PERSIST_VERSION = 3;

/** Cells the user changed — small enough for localStorage + IndexedDB. */
export function extractSheetEdits(raw) {
  if (!raw) {
    return {
      cells: [],
      headerRows: {},
      deletedRows: [],
      deletedCols: [],
      userRowGaps: [],
      userColGaps: [],
    };
  }
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
      if (!cell || !cell.userEdited) continue;
      const row = Number(rowKey);
      if (!headerRows[row]) headerRows[row] = {};
      headerRows[row][col] = {
        v: cell.v == null ? '' : String(cell.v),
        userEdited: true,
      };
    }
  }
  const deletedRows = Array.isArray(raw.deletedRows)
    ? [...new Set(raw.deletedRows.map(Number).filter(Number.isFinite))]
    : [];
  const deletedCols = Array.isArray(raw.deletedCols)
    ? [...new Set(raw.deletedCols.map(String).filter(Boolean))]
    : [];
  const userRowGaps = Array.isArray(raw.userRowGaps)
    ? raw.userRowGaps.map((g) => ({
        id: String(g.id || ''),
        afterExcelRow: Number(g.afterExcelRow),
      })).filter((g) => g.id && Number.isFinite(g.afterExcelRow))
    : [];
  const userColGaps = Array.isArray(raw.userColGaps)
    ? raw.userColGaps.map((g) => ({
        id: String(g.id || ''),
        afterCol: String(g.afterCol || ''),
      })).filter((g) => g.id && g.afterCol)
    : [];
  return { cells, headerRows, deletedRows, deletedCols, userRowGaps, userColGaps };
}

/**
 * Copy in-grid user edits onto export-shaped raw JSON before save/load fingerprint.
 * Needed when Synthesis was shown from a precomputed grid before synRaw finished loading.
 */
/** Detect synthesis JSON damaged by an old matrix apply (missing ADAPTATION band). */
export function synRawLooksHealthy(raw) {
  if (!raw) return true;
  const labelAt = (row) => {
    const idx = raw._cellIndex;
    if (idx instanceof Map) {
      const c = idx.get(`${row}:F`);
      return c && c.v != null ? String(c.v).trim() : '';
    }
    for (const c of raw.cells || []) {
      if (Number(c.r) === row && c.c === 'F') {
        return c.v == null ? '' : String(c.v).trim();
      }
    }
    return '';
  };
  const l25 = labelAt(25).toUpperCase();
  if (!l25 || !l25.includes('ADAPT')) return false;
  const l26 = labelAt(26).toUpperCase();
  if (!l26 || !l26.startsWith('_')) return false;
  const l41 = labelAt(41).toUpperCase();
  if (!l41 || (!l41.includes('ADTH') && !l41.includes('CABIN'))) return false;
  return true;
}

export function syncGridEditsToRaw(gridSheet, raw) {
  if (!gridSheet || !raw) return;
  for (const c of gridSheet.cells || []) {
    if (c && c.userEdited) upsertRawCell(raw, c.r, c.c, c.v);
  }
  for (const [rowKey, cols] of Object.entries(gridSheet.headerRows || {})) {
    const row = Number(rowKey);
    for (const [col, cell] of Object.entries(cols)) {
      if (cell && cell.userEdited) upsertRawCell(raw, row, col, cell.v);
    }
  }
  if (Array.isArray(gridSheet.deletedRows) && gridSheet.deletedRows.length) {
    const set = new Set(
      (raw.deletedRows || []).map(Number).filter(Number.isFinite)
    );
    for (const r of gridSheet.deletedRows) {
      const n = Number(r);
      if (Number.isFinite(n)) set.add(n);
    }
    raw.deletedRows = [...set];
  }
  if (Array.isArray(gridSheet.deletedCols) && gridSheet.deletedCols.length) {
    const set = new Set((raw.deletedCols || []).map(String).filter(Boolean));
    for (const c of gridSheet.deletedCols) {
      const col = String(c || '');
      if (col) set.add(col);
    }
    raw.deletedCols = [...set];
  }
  if (Array.isArray(gridSheet.userRowGaps)) {
    raw.userRowGaps = gridSheet.userRowGaps.map((g) => ({
      id: String(g.id || ''),
      afterExcelRow: Number(g.afterExcelRow),
    })).filter((g) => g.id && Number.isFinite(g.afterExcelRow));
  }
  if (Array.isArray(gridSheet.userColGaps)) {
    raw.userColGaps = gridSheet.userColGaps.map((g) => ({
      id: String(g.id || ''),
      afterCol: String(g.afterCol || ''),
    })).filter((g) => g.id && g.afterCol);
  }
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
  if (Array.isArray(edits.deletedRows) && edits.deletedRows.length) {
    const set = new Set(
      (raw.deletedRows || []).map(Number).filter(Number.isFinite)
    );
    for (const r of edits.deletedRows) {
      const n = Number(r);
      if (Number.isFinite(n)) set.add(n);
    }
    raw.deletedRows = [...set];
  }
  if (Array.isArray(edits.deletedCols) && edits.deletedCols.length) {
    const set = new Set((raw.deletedCols || []).map(String).filter(Boolean));
    for (const c of edits.deletedCols) {
      const col = String(c || '');
      if (col) set.add(col);
    }
    raw.deletedCols = [...set];
  }
  if (Array.isArray(edits.userRowGaps)) {
    raw.userRowGaps = edits.userRowGaps.map((g) => ({
      id: String(g.id || ''),
      afterExcelRow: Number(g.afterExcelRow),
    })).filter((g) => g.id && Number.isFinite(g.afterExcelRow));
  }
  if (Array.isArray(edits.userColGaps)) {
    raw.userColGaps = edits.userColGaps.map((g) => ({
      id: String(g.id || ''),
      afterCol: String(g.afterCol || ''),
    })).filter((g) => g.id && g.afterCol);
  }
}

/**
 * @param {{ bd?: object, syn?: object, revision?: number, structureRevision?: number }} payload
 */
export function buildPersistRecord({
  bd,
  syn,
  revision,
  structureRevision = 0,
}) {
  const hasStructure = structureRevision > 0;
  const bdEdits = extractSheetEdits(bd);
  const synEdits = extractSheetEdits(syn);
  const hasBdEdits =
    bdEdits.cells.length ||
    Object.keys(bdEdits.headerRows).length ||
    bdEdits.deletedRows.length ||
    bdEdits.deletedCols.length ||
    bdEdits.userRowGaps.length ||
    bdEdits.userColGaps.length;
  const hasSynEdits =
    synEdits.cells.length ||
    Object.keys(synEdits.headerRows).length ||
    synEdits.deletedRows.length ||
    synEdits.deletedCols.length ||
    synEdits.userRowGaps.length ||
    synEdits.userColGaps.length;

  if (!hasStructure && !hasBdEdits && !hasSynEdits) return null;

  const record = {
    version: PERSIST_VERSION,
    revision: revision != null ? revision : 0,
    structureRevision: structureRevision != null ? structureRevision : 0,
    structureSchema: STRUCTURE_SCHEMA_VERSION,
    savedAt: new Date().toISOString(),
  };

  if (hasStructure) {
    if (bd) record.bdFull = bd;
    // Never persist synFull: matrix row-remap snapshots reload as a broken grid
    // (ADAPTATION band replaced by FENDERS, etc.). Synthesis rebuilds from the
    // project JSON + synEdits; structure comes from bdFull + matrix re-apply.
    if (hasSynEdits) record.synEdits = synEdits;
  } else {
    if (hasBdEdits) record.bdEdits = bdEdits;
    if (hasSynEdits) record.synEdits = synEdits;
  }

  return record;
}

/**
 * A persisted full-structure snapshot is stale when it was produced by an older
 * structure schema. Such a snapshot must be discarded so the current code rebuilds
 * the structure from the project JSON (otherwise old/duplicated bookmarks return).
 */
export function isStructureSnapshotStale(snapshot) {
  if (!snapshot) return false;
  const hasFrozenStructure =
    Boolean(snapshot.bdFull) ||
    Boolean(snapshot.synFull) ||
    Number(snapshot.structureRevision) > 0;
  if (!hasFrozenStructure) return false;
  return snapshot.structureSchema !== STRUCTURE_SCHEMA_VERSION;
}

/**
 * Strip the frozen-structure parts of a stale snapshot in place, keeping plain
 * cell edits so the user's typed values survive while the structure is rebuilt.
 * @returns {boolean} true when the snapshot was modified.
 */
export function dropStaleStructure(snapshot) {
  if (!isStructureSnapshotStale(snapshot)) return false;
  delete snapshot.bdFull;
  delete snapshot.synFull;
  snapshot.structureRevision = 0;
  return true;
}

/** @returns {boolean} true if legacy full snapshot was applied */
export function applyPersistRecord(bdRaw, synRaw, record) {
  if (!record) return false;

  if (record.version === PERSIST_VERSION) {
    if (record.bdFull) {
      Object.assign(bdRaw, record.bdFull);
    } else if (record.bdEdits) {
      applySheetEdits(bdRaw, record.bdEdits);
    }
    if (synRaw) {
      if (record.synFull) {
        Object.assign(synRaw, record.synFull);
      } else if (record.synEdits) {
        applySheetEdits(synRaw, record.synEdits);
      }
    }
    return Boolean(record.bdFull || record.synFull);
  }

  if (record.version === 2) {
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

/** Update or insert a cell on export-shaped sheet JSON. O(1) via lazy row:col index. */
function ensureRawCellIndex(raw) {
  if (!raw) return null;
  if (raw._cellIndex instanceof Map) return raw._cellIndex;
  const idx = new Map();
  for (const c of raw.cells || []) idx.set(`${c.r}:${c.c}`, c);
  raw._cellIndex = idx;
  return idx;
}

export function upsertRawCell(raw, row, col, value) {
  if (!raw) return;
  const r = Number(row);
  const c = String(col);
  const strVal = value == null ? '' : String(value);
  const idx = ensureRawCellIndex(raw);

  let cell = idx ? idx.get(`${r}:${c}`) : null;
  if (!cell) {
    cell = { r, c, v: strVal, userEdited: true };
    if (!raw.cells) raw.cells = [];
    raw.cells.push(cell);
    if (idx) idx.set(`${r}:${c}`, cell);
  } else {
    cell.v = strVal;
    cell.userEdited = true;
    delete cell.f;
  }

  const rowHdr =
    (raw.headerRows && raw.headerRows[String(r)]) ||
    (raw.headerRows && raw.headerRows[r]);
  if (rowHdr && rowHdr[c]) {
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
      req.onsuccess = () => resolve(req.result != null ? req.result : null);
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
    const json = JSON.stringify(record);
    if (json.length > 512 * 1024) return;
    localStorage.setItem(LS_SNAPSHOT_KEY, json);
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
    if (!data || (!data.bd && !data.syn)) return;
    if (saving) {
      pending = true;
      return;
    }
    saving = true;
    if (onStatus) onStatus('saving');
    try {
      const cfg = await probeSheetApi();
      const cloud =
        cfg &&
        (cfg.cloudPersist ||
          cfg.mode === 'databricks' ||
          cfg.mode === 'supabase' ||
          cfg.mode === 'postgres');
      const remoteOnly = Boolean(cfg && cfg.remoteOnly);
      if (!remoteOnly) {
        await saveLocalSnapshot(data);
      }
      if (cloud && (data.bd || data.syn)) {
        const structRev =
          data.structureRevision != null ? data.structureRevision : 0;
        const res = await saveSession(cfg.projectId || 'default', {
          revision: data.revision != null ? data.revision : 0,
          bd: data.bd,
          syn: data.syn,
          updated_by: 'web',
          structure_revision: structRev,
          structureRevision: structRev,
        });
        if (res && res.conflict) {
          const err = new Error(
            `Conflit de revision (serveur ${res.revision}, local ${data.revision})`
          );
          err.code = 'REVISION_CONFLICT';
          throw err;
        }
      } else if (remoteOnly) {
        throw new Error('API Supabase inaccessible — lancez go-api.bat');
      }
      if (onStatus) onStatus('saved');
    } catch (e) {
      console.error('Auto-save failed:', e);
      if (onStatus) onStatus('error', e);
    } finally {
      saving = false;
      if (pending) {
        pending = false;
        void flush();
      }
    }
  }

  function schedule() {
    if (onStatus) onStatus('dirty');
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

/** Stable id for raw JSON — detects project file updates. */
export function rawFingerprint(raw) {
  if (!raw) return '';
  const cells = raw.cells && raw.cells.length ? raw.cells.length : 0;
  const cols = raw.columns && raw.columns.length ? raw.columns.length : 0;
  const lastRow = raw.lastRow != null ? raw.lastRow : 0;
  // Hash the actual content (r, c, v) of every edited cell — order-independent —
  // so re-editing an already-edited cell still changes the fingerprint and
  // invalidates any stale cached grid transform. A count-only fingerprint would
  // collide when a value changes without the edited-cell count changing, causing
  // a refresh to show the previous value instead of the one just entered.
  let edited = 0;
  let editHash = 0;
  for (const c of raw.cells || []) {
    if (!c.userEdited) continue;
    edited++;
    const s = `${c.r}:${c.c}=${c.v == null ? '' : c.v}`;
    let h = 2166136261;
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    editHash = (editHash + (h >>> 0)) >>> 0;
  }
  const delRows = Array.isArray(raw.deletedRows)
    ? [...raw.deletedRows].map(Number).filter(Number.isFinite).sort((a, b) => a - b)
    : [];
  const delKey = delRows.length ? delRows.join(',') : '';
  const delCols = Array.isArray(raw.deletedCols)
    ? [...raw.deletedCols].map(String).filter(Boolean).sort()
    : [];
  const delColKey = delCols.length ? delCols.join(',') : '';
  const uRowGaps = Array.isArray(raw.userRowGaps)
    ? raw.userRowGaps.map((g) => `${g.afterExcelRow}:${g.id}`).sort().join('|')
    : '';
  const uColGaps = Array.isArray(raw.userColGaps)
    ? raw.userColGaps.map((g) => `${g.afterCol}:${g.id}`).sort().join('|')
    : '';
  return `${TRANSFORM_VERSION}:${cells}:${cols}:${lastRow}:${edited}:${editHash.toString(36)}:del${delKey}:delc${delColKey}:urg${uRowGaps}:ucg${uColGaps}`;
}

export function serializeTransformSheet(sheet) {
  if (!sheet) return null;
  const { cellMap, canonicalSectionMap, ...rest } = sheet;
  const out = { ...rest };
  if (out.sectionHeaderRows instanceof Set) {
    out.sectionHeaderRows = [...out.sectionHeaderRows];
  }
  return out;
}

export function hydrateTransformSheet(data) {
  if (!data) return null;
  return {
    ...data,
    cellMap: buildCellMap(data.cells, data.headerRows),
  };
}

export async function clearSheetTransform(sheetId) {
  try {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_TRANSFORMS, 'readwrite');
      tx.objectStore(STORE_TRANSFORMS).delete(sheetId);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch {
    return undefined;
  }
}

export async function loadSheetTransform(sheetId, fingerprint) {
  try {
    const db = await openDb();
    return new Promise((resolve) => {
      const tx = db.transaction(STORE_TRANSFORMS, 'readonly');
      const req = tx.objectStore(STORE_TRANSFORMS).get(sheetId);
      req.onsuccess = () => {
        const rec = req.result;
        if (rec && rec.fingerprint === fingerprint && rec.sheet) {
          resolve(hydrateTransformSheet(rec.sheet));
        } else resolve(null);
      };
      req.onerror = () => resolve(null);
    });
  } catch {
    return null;
  }
}

export async function saveSheetTransform(sheetId, fingerprint, sheet, { skipIfPrecomputed = false } = {}) {
  if (skipIfPrecomputed) return false;
  try {
    const db = await openDb();
    // IndexedDB uses the structured clone algorithm. In some environments, objects that
    // look like plain arrays/objects can still fail to clone (e.g. proxies / exotic arrays).
    // Normalize through JSON to guarantee cloneability for cached transforms.
    const normalizedSheet = (() => {
      const data = serializeTransformSheet(sheet);
      try {
        return JSON.parse(JSON.stringify(data));
      } catch {
        return data;
      }
    })();
    const payload = {
      fingerprint,
      sheet: normalizedSheet,
      savedAt: new Date().toISOString(),
    };
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_TRANSFORMS, 'readwrite');
      tx.objectStore(STORE_TRANSFORMS).put(payload, sheetId);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => reject(tx.error);
    });
  } catch (e) {
    console.warn('saveSheetTransform failed:', e);
    return false;
  }
}
