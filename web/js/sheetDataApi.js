/**
 * Progressive sheet loading — API chunks first, full JSON fallback.
 * Works with web/server.mjs (local) or api/ FastAPI + Databricks (prod).
 */

const SHEET_FILES = {
  bd: '/public/data/bd-sheet.json',
  synthesis: '/public/data/synthesis-sheet.json',
};

const CHUNK_ROWS = { bd: 400, synthesis: 200 };

let apiConfigCache = null;
let apiConfigPromise = null;

function apiBase() {
  if (typeof window !== 'undefined' && window.__WGHT_API_BASE__) {
    return String(window.__WGHT_API_BASE__).replace(/\/$/, '');
  }
  return '';
}

function url(path) {
  const base = apiBase();
  return base ? `${base}${path}` : path;
}

async function fetchJson(path, init) {
  const res = await fetch(url(path), init);
  if (!res.ok) {
    throw new Error(`HTTP ${res.status} for ${path}`);
  }
  return res.json();
}

/** @returns {Promise<{ mode: string, chunkedLoad: boolean } | null>} */
export async function probeSheetApi() {
  if (apiConfigCache) return apiConfigCache;
  if (!apiConfigPromise) {
    apiConfigPromise = (async () => {
      try {
        const cfg = await fetchJson('/api/v1/config', {
          signal: AbortSignal.timeout(1200),
        });
        apiConfigCache = cfg;
        return cfg;
      } catch {
        apiConfigCache = { mode: 'static', chunkedLoad: false };
        return apiConfigCache;
      }
    })();
  }
  return apiConfigPromise;
}

export function resetSheetApiProbe() {
  apiConfigCache = null;
  apiConfigPromise = null;
}

function mergeMeta(raw, meta) {
  for (const key of Object.keys(meta)) {
    if (key === 'cells') continue;
    raw[key] = meta[key];
  }
}

/**
 * @param {'bd'|'synthesis'} sheetId
 * @param {{
 *   onFirstChunk?: (raw: object) => void | Promise<void>,
 *   onProgress?: (info: object) => void,
 * }} [opts]
 */
export async function loadSheetRaw(sheetId, opts = {}) {
  const cfg = await probeSheetApi();
  if (!cfg?.chunkedLoad) {
    return fetchJson(SHEET_FILES[sheetId]);
  }

  const meta = await fetchJson(`/api/v1/sheets/${sheetId}/meta`);
  const raw = { ...meta, cells: [] };
  const step = CHUNK_ROWS[sheetId] || 300;
  const rowMin = sheetId === 'bd' ? 2 : 1;
  const rowMax = Number(meta.lastRow) || rowMin;
  let firstDone = false;

  for (let start = rowMin; start <= rowMax; start += step) {
    const end = Math.min(start + step - 1, rowMax);
    const chunk = await fetchJson(
      `/api/v1/sheets/${sheetId}/cells?rowMin=${start}&rowMax=${end}`
    );
    raw.cells.push(...(chunk.cells || []));
    opts.onProgress?.({
      sheetId,
      rowMin: start,
      rowMax: end,
      lastRow: rowMax,
      cellCount: raw.cells.length,
    });
    if (!firstDone && opts.onFirstChunk) {
      firstDone = true;
      await opts.onFirstChunk({ ...raw, cells: [...raw.cells] });
    }
  }

  return raw;
}

/** Load remaining rows after first chunk (when raw already has meta + partial cells). */
export async function loadSheetRemainingCells(sheetId, raw, opts = {}) {
  const cfg = await probeSheetApi();
  if (!cfg?.chunkedLoad) return raw;

  const step = CHUNK_ROWS[sheetId] || 300;
  const rowMin = sheetId === 'bd' ? 2 : 1;
  const rowMax = Number(raw.lastRow) || rowMin;
  let maxLoaded = rowMin - 1;
  for (const c of raw.cells || []) {
    if (c.r > maxLoaded) maxLoaded = c.r;
  }
  let start = maxLoaded + 1;
  if (start <= rowMin) start = rowMin + step;

  for (; start <= rowMax; start += step) {
    const end = Math.min(start + step - 1, rowMax);
    const chunk = await fetchJson(
      `/api/v1/sheets/${sheetId}/cells?rowMin=${start}&rowMax=${end}`
    );
    raw.cells.push(...(chunk.cells || []));
    opts.onProgress?.({
      sheetId,
      rowMin: start,
      rowMax: end,
      lastRow: rowMax,
      cellCount: raw.cells.length,
    });
  }
  return raw;
}

/** Session snapshot — Databricks API when configured. */
export async function fetchSession(projectId = 'default') {
  try {
    return await fetchJson(`/api/v1/sessions/${encodeURIComponent(projectId)}`);
  } catch {
    return null;
  }
}

export async function saveSession(projectId, payload) {
  const res = await fetch(url(`/api/v1/sessions/${encodeURIComponent(projectId)}`), {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(text || `Save failed (${res.status})`);
  }
  return res.json();
}

/**
 * PATCH cell values — BD edits trigger server-side Synthesis recalc (P2).
 * @param {'bd'|'synthesis'} sheetId
 * @param {{ r: number, c: string, v: * }[]} changes
 * @returns {Promise<{ ok: boolean, synPatches?: { r: number, c: string, v: string, mat?: boolean }[] }>}
 */
export async function patchSheetCells(sheetId, changes) {
  const res = await fetch(url(`/api/v1/sheets/${sheetId}/cells`), {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ changes }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(text || `PATCH failed (${res.status})`);
  }
  return res.json();
}

export { mergeMeta, CHUNK_ROWS };
