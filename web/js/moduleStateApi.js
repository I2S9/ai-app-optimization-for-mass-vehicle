/**
 * Remote module state — Supabase via /api/v1/modules/{moduleKey}
 * (Weight Tax, Waterline, CDC Output, …)
 */
import { probeSheetApi } from './sheetDataApi.js';

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

/**
 * @param {string} moduleKey
 * @param {string} [projectId]
 * @returns {Promise<{ revision: number, state: object, updated_at?: string } | null>}
 */
export async function fetchModuleState(moduleKey, projectId = 'default') {
  const path =
    `/api/v1/modules/${encodeURIComponent(moduleKey)}` +
    `?projectId=${encodeURIComponent(projectId)}`;
  const res = await fetch(url(path), { signal: AbortSignal.timeout(12000) });
  if (res.status === 404) return null;
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(text || `Module state load failed (${res.status})`);
  }
  return res.json();
}

/**
 * @param {string} moduleKey
 * @param {{ revision?: number, state: object, updated_by?: string }} payload
 * @param {string} [projectId]
 */
export async function saveModuleState(moduleKey, payload, projectId = 'default') {
  const path =
    `/api/v1/modules/${encodeURIComponent(moduleKey)}` +
    `?projectId=${encodeURIComponent(projectId)}`;
  const res = await fetch(url(path), {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      revision: payload.revision != null ? payload.revision : 0,
      state: payload.state || {},
      updated_by: payload.updated_by || 'web',
    }),
    signal: AbortSignal.timeout(12000),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(text || `Module state save failed (${res.status})`);
  }
  return res.json();
}

/** @returns {Promise<{ cloudPersist: boolean, projectId: string }>} */
export async function moduleCloudConfig() {
  const cfg = await probeSheetApi();
  return {
    cloudPersist: Boolean(cfg && cfg.cloudPersist),
    projectId: (cfg && cfg.projectId) || 'default',
  };
}
