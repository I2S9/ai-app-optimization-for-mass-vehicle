/**
 * Vercel serverless — Supabase module state (Weight Tax, Waterline, …)
 * Route: /api/v1/modules/:moduleKey  (GET + PUT)
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

async function fetchModuleState(projectId, moduleKey) {
  const u =
    `${SUPABASE_URL}/rest/v1/module_state?project_id=eq.` +
    `${encodeURIComponent(projectId)}` +
    `&module_key=eq.${encodeURIComponent(moduleKey)}` +
    `&select=project_id,module_key,revision,state_json,updated_at,updated_by&limit=1`;
  const res = await fetch(u, { headers: supabaseHeaders() });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  const rows = await res.json();
  if (!rows || !rows.length) return null;
  const row = rows[0];
  return {
    project_id: projectId,
    module_key: moduleKey,
    revision: Number(row.revision || 0),
    state: row.state_json || {},
    updated_at: row.updated_at,
    updated_by: row.updated_by,
  };
}

async function upsertModuleState(projectId, moduleKey, payload) {
  const revision = Number(payload.revision || 0);
  const existing = await fetchModuleState(projectId, moduleKey).catch(() => null);
  if (existing && existing.revision > revision) {
    return {
      project_id: projectId,
      module_key: moduleKey,
      revision: existing.revision,
      ok: false,
      conflict: true,
    };
  }
  const body = [
    {
      project_id: projectId,
      module_key: moduleKey,
      revision,
      state_json: payload.state || {},
      updated_by: payload.updated_by || 'web',
    },
  ];
  const res = await fetch(`${SUPABASE_URL}/rest/v1/module_state`, {
    method: 'POST',
    headers: supabaseHeaders({ Prefer: 'resolution=merge-duplicates,return=representation' }),
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Supabase ${res.status}: ${await res.text()}`);
  const rows = await res.json();
  const savedRev = rows && rows[0] ? Number(rows[0].revision || revision) : revision;
  return { project_id: projectId, module_key: moduleKey, revision: savedRev, ok: true };
}

module.exports = async (req, res) => {
  if (!SUPABASE_URL || !SUPABASE_KEY) {
    res
      .status(503)
      .json({ error: 'Supabase env vars missing (SUPABASE_URL, SUPABASE_SERVICE_KEY)' });
    return;
  }

  const moduleKey =
    (req.query && (req.query.moduleKey || req.query.modulekey)) || 'unknown';
  const projectId =
    (req.query && (req.query.projectId || req.query.projectid)) || 'default';

  try {
    if (req.method === 'GET') {
      const row = await fetchModuleState(String(projectId), String(moduleKey));
      if (!row) {
        res.status(404).json({ error: 'Module state not found' });
        return;
      }
      res.status(200).json(row);
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
      const result = await upsertModuleState(String(projectId), String(moduleKey), body || {});
      res.status(200).json(result);
      return;
    }

    res.setHeader('Allow', 'GET, PUT');
    res.status(405).json({ error: 'Method not allowed' });
  } catch (e) {
    res.status(503).json({ error: (e && e.message) || String(e) });
  }
};
