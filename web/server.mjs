/**
 * Static server + progressive sheet API + server-side Synthesis calc (P2).
 * Run: node web/server.mjs
 */
import http from 'http';
import fs from 'fs';
import path from 'path';
import zlib from 'zlib';
import { fileURLToPath } from 'url';
import { transformBdSheet, transformSynthesisSheet } from './js/sheetTransform.js';
import { createSynCalcContext } from './js/synMaterialize.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT) || 5173;
const HOST = '127.0.0.1';
/** FastAPI + Supabase — port 8000 (go-api.bat) */
const PROXY_API = (process.env.WGHT_PROXY_API || 'http://127.0.0.1:8000').replace(
  /\/$/,
  ''
);

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.ico': 'image/x-icon',
};

const SHEET_FILES = {
  bd: path.join(__dirname, 'public/data/bd-sheet.json'),
  synthesis: path.join(__dirname, 'public/data/synthesis-sheet.json'),
};

/** @type {Map<string, object>} raw JSON cache */
const sheetCache = new Map();
/** @type {{ bd: object, syn: object, ctx: object } | null} */
let calcRuntime = null;

function loadSheetRaw(sheetId) {
  if (sheetCache.has(sheetId)) return sheetCache.get(sheetId);
  const filePath = SHEET_FILES[sheetId];
  if (!filePath || !fs.existsSync(filePath)) {
    throw new Error(`Missing sheet file for ${sheetId}`);
  }
  const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  sheetCache.set(sheetId, data);
  return data;
}

function getCalcRuntime() {
  if (calcRuntime) return calcRuntime;
  const bdRaw = loadSheetRaw('bd');
  const synRaw = loadSheetRaw('synthesis');
  const bdSheet = transformBdSheet(bdRaw);
  const synSheet = transformSynthesisSheet(synRaw);
  const ctx = createSynCalcContext(bdSheet, synSheet);
  calcRuntime = { bdRaw, synRaw, bdSheet, synSheet, ctx };
  return calcRuntime;
}

function resetCalcRuntime() {
  calcRuntime = null;
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      try {
        const text = Buffer.concat(chunks).toString('utf8') || '{}';
        resolve(JSON.parse(text));
      } catch (e) {
        reject(e);
      }
    });
    req.on('error', reject);
  });
}

function upsertRawCell(raw, row, col, value) {
  const cells = raw.cells || (raw.cells = []);
  let cell = cells.find((c) => c.r === row && c.c === col);
  if (!cell) {
    cell = { r: row, c: col, v: value, userEdited: true };
    cells.push(cell);
  } else {
    cell.v = value;
    cell.userEdited = true;
    delete cell.f;
  }
}

function parseQuery(url) {
  const i = url.indexOf('?');
  if (i < 0) return {};
  const out = {};
  for (const part of url.slice(i + 1).split('&')) {
    const [k, v] = part.split('=');
    if (k) out[decodeURIComponent(k)] = decodeURIComponent(v || '');
  }
  return out;
}

function sendJson(res, obj, status = 200) {
  const body = Buffer.from(JSON.stringify(obj));
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': body.length,
    'Cache-Control': 'no-store',
  });
  res.end(body);
}

async function handleApi(req, res, urlPath, fullUrl) {
  if (urlPath === '/api/v1/config') {
    // Local dev: full JSON is faster than many /cells chunks (re-parse file each request).
    const chunkedLoad = process.env.WGHT_CHUNKED_LOAD === '1';
    const serverCalc = process.env.WGHT_SERVER_CALC === '1';
    sendJson(res, {
      mode: 'local-node',
      chunkedLoad,
      serverCalc,
      version: 2,
      projectId: 'default',
    });
    return true;
  }
  if (urlPath === '/api/v1/health') {
    sendJson(res, { ok: true, backend: 'local-node', serverCalc: true });
    return true;
  }

  const metaMatch = urlPath.match(/^\/api\/v1\/sheets\/(bd|synthesis)\/meta$/);
  if (metaMatch) {
    try {
      const raw = loadSheetRaw(metaMatch[1]);
      const meta = { ...raw };
      delete meta.cells;
      meta.cellCount = (raw.cells || []).length;
      sendJson(res, meta);
    } catch (e) {
      sendJson(res, { error: e.message }, 404);
    }
    return true;
  }

  const cellsMatch = urlPath.match(/^\/api\/v1\/sheets\/(bd|synthesis)\/cells$/);
  if (cellsMatch && req.method === 'GET') {
    const q = parseQuery(fullUrl);
    const rowMin = Number(q.rowMin);
    const rowMax = Number(q.rowMax);
    if (!Number.isFinite(rowMin) || !Number.isFinite(rowMax) || rowMax < rowMin) {
      sendJson(res, { error: 'rowMin and rowMax required' }, 400);
      return true;
    }
    try {
      const raw = loadSheetRaw(cellsMatch[1]);
      const cells = [];
      for (const cell of raw.cells || []) {
        const r = Number(cell.r);
        if (r >= rowMin && r <= rowMax) cells.push(cell);
      }
      sendJson(res, { cells, rowMin, rowMax });
    } catch (e) {
      sendJson(res, { error: e.message }, 404);
    }
    return true;
  }

  const patchMatch = urlPath.match(/^\/api\/v1\/sheets\/(bd|synthesis)\/cells$/);
  if (patchMatch && req.method === 'PATCH') {
    try {
      const body = await readJsonBody(req);
      const changes = Array.isArray(body?.changes) ? body.changes : [];
      if (!changes.length) {
        sendJson(res, { error: 'changes[] required' }, 400);
        return true;
      }
      const sheetId = patchMatch[1];
      const raw = loadSheetRaw(sheetId);
      for (const ch of changes) {
        const r = Number(ch.r);
        const c = String(ch.c || '');
        if (!Number.isFinite(r) || !c) continue;
        upsertRawCell(raw, r, c, ch.v ?? '');
      }
      resetCalcRuntime();
      let synPatches = [];
      if (sheetId === 'bd') {
        const rt = getCalcRuntime();
        for (const ch of changes) {
          rt.ctx.patchBdCell(Number(ch.r), String(ch.c), ch.v ?? '');
        }
        synPatches = rt.ctx.recalcAfterBdEdit();
        rt.ctx.applyPatches(synPatches);
      }
      sendJson(res, { ok: true, synPatches });
    } catch (e) {
      sendJson(res, { error: e.message }, 500);
    }
    return true;
  }

  return false;
}

function serveStatic(req, res, urlPath) {
  if (urlPath === '/') urlPath = '/index.html';
  const filePath = path.join(__dirname, urlPath.replace(/^\//, ''));
  if (!filePath.startsWith(__dirname)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    const ext = path.extname(filePath);
    const noStore = ['.html', '.js', '.css', '.json'].includes(ext);
    const accept = req.headers['accept-encoding'] || '';
    const useGzip =
      ext === '.json' && data.length > 4096 && accept.includes('gzip');

    const headers = {
      'Content-Type': MIME[ext] || 'application/octet-stream',
      'Cache-Control': noStore ? 'no-store, no-cache, must-revalidate' : 'public, max-age=3600',
    };
    if (noStore) headers.Pragma = 'no-cache';

    if (useGzip) {
      zlib.gzip(data, (gzipErr, compressed) => {
        if (gzipErr) {
          res.writeHead(200, headers);
          res.end(data);
          return;
        }
        headers['Content-Encoding'] = 'gzip';
        headers['Content-Length'] = compressed.length;
        res.writeHead(200, headers);
        res.end(compressed);
      });
      return;
    }

    res.writeHead(200, headers);
    res.end(data);
  });
}

function readRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

async function proxyApiToBackend(req, res, fullUrl, urlPath) {
  const target = `${PROXY_API}${fullUrl}`;
  const method = req.method || 'GET';
  const headers = {};
  for (const [k, v] of Object.entries(req.headers)) {
    if (k.toLowerCase() === 'host') continue;
    headers[k] = v;
  }
  let body;
  if (method !== 'GET' && method !== 'HEAD') {
    body = await readRawBody(req);
  }
  try {
    const upstream = await fetch(target, {
      method,
      headers,
      body: body && body.length ? body : undefined,
    });
    const buf = Buffer.from(await upstream.arrayBuffer());
    res.writeHead(upstream.status, {
      'Content-Type': upstream.headers.get('content-type') || 'application/json',
      'Cache-Control': 'no-store',
      'Content-Length': buf.length,
    });
    res.end(buf);
    return true;
  } catch (e) {
    const msg = String(e.message || e);
    sendJson(
      res,
      {
        ok: false,
        error: `API Python injoignable sur ${PROXY_API}`,
        detail: msg,
        hint: 'Ouvrez une fenetre et lancez: api\\go-api.bat  puis rechargez cette page',
        urlPath,
      },
      503
    );
    return true;
  }
}

const server = http.createServer(async (req, res) => {
  const fullUrl = req.url || '/';
  const urlPath = fullUrl.split('?')[0];
  if (urlPath.startsWith('/api/')) {
    // Local-first : on sert config / sheets (meta, cells, PATCH) directement
    // depuis les fichiers JSON de web/public/data. run-bd-server fonctionne donc
    // sans Supabase/Python. Le proxy vers l'API Python (port 8000) ne sert que
    // de repli pour les routes que ce serveur ne gere pas (ex : /sessions).
    if (await handleApi(req, res, urlPath, fullUrl)) return;
    if (await proxyApiToBackend(req, res, fullUrl, urlPath)) return;
    sendJson(
      res,
      {
        error: 'Route API inconnue',
        path: urlPath,
        hint: 'Demarrez api\\go-api.bat puis http://127.0.0.1:5173/api/v1/config',
      },
      404
    );
    return;
  }
  serveStatic(req, res, urlPath);
});

server.listen(PORT, HOST, () => {
  console.log(`WGHT app:     http://${HOST}:${PORT}/`);
  console.log(`API config:   http://${HOST}:${PORT}/api/v1/config  → proxy ${PROXY_API}`);
  console.log('Si config echoue: lancez api\\go-api.bat dans une autre fenetre');
});
server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is already in use. Close the other server or set PORT=8080.`);
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
