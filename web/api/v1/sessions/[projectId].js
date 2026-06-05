/**
 * Vercel serverless function — Supabase session persistence for the web app.
 *
 * Route: /api/v1/sessions/:projectId  (GET + PUT)
 * Storage: Supabase table `workbook_sessions` (one row per projectId).
 *
 * The browser compresses the BD/SYN snapshots (gz+b64:) BEFORE sending, so this
 * function is a thin proxy: it never inflates the payload. That keeps every request
 * and response under the Vercel 4.5 MB body limit (full BD+SYN JSON is ~12 MB).
 *
 * The Supabase service key is read from the SUPABASE_SERVICE_KEY env var and stays
 * server-side only — it is never exposed to the browser.
 *
 * Required Vercel env vars:
 *   SUPABASE_URL          e.g. https://xxxx.supabase.co
 *   SUPABASE_SERVICE_KEY  service_role key (secret)
 */

const SUPABASE_URL = (process.env.SUPABASE_URL || '').replace(/\/$/, '');
const SUPABASE_KEY =
  process.env.SUPABASE_SERVICE_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY || '';

function supabaseHeaders(extra) {
  return Object.assign(
    {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${SUPABASE_KEY}`,
      'Content-Type': 'application/json',
    },
    extra || {}
  );
}

async function fetchSession(projectId) {
  const u =
    `${SUPABASE_URL}/rest/v1/workbook_sessions?project_id=eq.` +
    `${encodeURIComponent(projectId)}` +
    `&select=project_id,revision,structure_revision,bd_snapshot,syn_snapshot,updated_at,updated_by&limit=1`;
  const res = await fetch(u, { headers: supabaseHeaders() });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  const rows = await res.json();
  if (!rows || !rows.length) return null;
  const row = rows[0];
  return {
    project_id: projectId,
    revision: Number(row.revision || 0),
    structure_revision: Number(row.structure_revision || 0),
    structureRevision: Number(row.structure_revision || 0),
    // Compressed strings, decompressed in the browser.
    bd_snapshot: row.bd_snapshot,
    syn_snapshot: row.syn_snapshot,
    updated_at: row.updated_at,
    updated_by: row.updated_by,
  };
}

async function upsertSession(projectId, payload) {
  const revision = Number(payload.revision || 0);
  const structRev = Number(
    payload.structure_revision != null
      ? payload.structure_revision
      : payload.structureRevision != null
        ? payload.structureRevision
        : 0
  );
  // Revision guard: never let an older client overwrite a newer cloud snapshot.
  const existing = await fetchSession(projectId).catch(() => null);
  if (existing && existing.revision > revision) {
    return { project_id: projectId, revision: existing.revision, ok: false, conflict: true };
  }
  const body = [
    {
      project_id: projectId,
      revision,
      structure_revision: structRev,
      // Already compressed by the browser; store verbatim.
      bd_snapshot: payload.bd_snapshot != null ? payload.bd_snapshot : null,
      syn_snapshot: payload.syn_snapshot != null ? payload.syn_snapshot : null,
      updated_by: payload.updated_by || 'web',
    },
  ];
  const res = await fetch(`${SUPABASE_URL}/rest/v1/workbook_sessions`, {
    method: 'POST',
    headers: supabaseHeaders({ Prefer: 'resolution=merge-duplicates,return=minimal' }),
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  return { project_id: projectId, revision, structureRevision: structRev, ok: true };
}

module.exports = async (req, res) => {
  if (!SUPABASE_URL || !SUPABASE_KEY) {
    res
      .status(503)
      .json({ error: 'Supabase env vars missing (SUPABASE_URL, SUPABASE_SERVICE_KEY)' });
    return;
  }

  const projectId =
    (req.query && (req.query.projectId || req.query.projectid)) || 'default';

  try {
    if (req.method === 'GET') {
      const session = await fetchSession(String(projectId));
      if (!session) {
        res.status(404).json({ error: 'Session not found' });
        return;
      }
      res.status(200).json(session);
      return;
    }

    if (req.method === 'PUT') {
      let body = req.body;
      if (typeof body === 'string') {
        try {
          body = JSON.parse(body);
        } catch {
          body = {};
        }
      }
      const result = await upsertSession(String(projectId), body || {});
      res.status(200).json(result);
      return;
    }

    res.setHeader('Allow', 'GET, PUT');
    res.status(405).json({ error: 'Method not allowed' });
  } catch (e) {
    res.status(503).json({ error: (e && e.message) || String(e) });
  }
};
