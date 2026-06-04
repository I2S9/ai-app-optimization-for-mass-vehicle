/**
 * Fast boot: precomputed grid JSON + IndexedDB transform cache.
 *
 * Storage layers (time/space tradeoff):
 *  L0 — precomputed *-grid.json (build time, O(1) hydrate)
 *  L1 — IndexedDB transform cache (per-browser, skip if fingerprint matches L0)
 *  L2 — raw JSON + in-browser transform (only when user edits differ from project files)
 */
import {
  hydrateTransformSheet,
  loadSheetTransform,
  rawFingerprint,
} from './sessionPersistence.js?v=persist-v6';
import { fetchGridPack } from './sheetDataApi.js?v=p2';

const GRID_URLS = {
  bd: '/public/data/bd-sheet-grid.json',
  syn: '/public/data/synthesis-sheet-grid.json',
};

const packCache = new Map();

export function hasSheetEdits(edits) {
  if (!edits) return false;
  if (edits.cells && edits.cells.length) return true;
  if (edits.deletedRows && edits.deletedRows.length) return true;
  return Object.keys(edits.headerRows || {}).length > 0;
}

export async function fetchJson(url) {
  const res = await fetch(url);
  if (!res.ok) return null;
  return res.json();
}

/** @returns {{ fingerprint: string, sheet: object } | null} */
export async function fetchPrecomputedPack(sheetId) {
  if (packCache.has(sheetId)) return packCache.get(sheetId);
  // Prefer the DB-served grid (JSON-free path); fall back to the static file so the
  // app keeps working while the migration is in progress / in pure static mode.
  let pack = await fetchGridPack(sheetId);
  if (!pack || !pack.sheet) {
    const url = GRID_URLS[sheetId];
    if (!url) return null;
    try {
      pack = await fetchJson(url);
    } catch {
      return null;
    }
  }
  if (!pack || !pack.sheet) return null;
  const hydrated = {
    fingerprint: pack.fingerprint || '',
    sheet: hydrateTransformSheet(pack.sheet),
  };
  packCache.set(sheetId, hydrated);
  return hydrated;
}

/**
 * Resolve a display-ready grid sheet without running transform in-browser.
 * @returns {object | null}
 */
export async function resolveGridSheet(sheetId, raw, { force = false } = {}) {
  if (!force && raw) {
    const fp = rawFingerprint(raw);
    const idb = await loadSheetTransform(sheetId, fp);
    if (idb) return idb;
    const pre = await fetchPrecomputedPack(sheetId);
    if (pre && pre.fingerprint === fp) return pre.sheet;
    return null;
  }
  if (!force && !raw) {
    const pre = await fetchPrecomputedPack(sheetId);
    if (pre && pre.sheet) return pre.sheet;
  }
  return null;
}
