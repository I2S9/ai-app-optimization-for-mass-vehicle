/**
 * Static server + progressive sheet API (same origin).
 * Run: node web/server.mjs
 */
import http from 'http';
import fs from 'fs';
import path from 'path';
import zlib from 'zlib';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT) || 5173;
const HOST = '127.0.0.1';

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

/** @type {Map<string, object>} */
const sheetCache = new Map();

function loadSheet(sheetId) {
  if (sheetCache.has(sheetId)) return sheetCache.get(sheetId);
  const filePath = SHEET_FILES[sheetId];
  if (!filePath || !fs.existsSync(filePath)) {
    throw new Error(`Missing sheet file for ${sheetId}`);
  }
  const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  sheetCache.set(sheetId, data);
  return data;
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

function handleApi(req, res, urlPath, fullUrl) {
  if (urlPath === '/api/v1/config') {
    sendJson(res, { mode: 'local-node', chunkedLoad: true, version: 1, projectId: 'default' });
    return true;
  }
  if (urlPath === '/api/v1/health') {
    sendJson(res, { ok: true, backend: 'local-node' });
    return true;
  }

  const metaMatch = urlPath.match(/^\/api\/v1\/sheets\/(bd|synthesis)\/meta$/);
  if (metaMatch) {
    try {
      const raw = loadSheet(metaMatch[1]);
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
  if (cellsMatch) {
    const q = parseQuery(fullUrl);
    const rowMin = Number(q.rowMin);
    const rowMax = Number(q.rowMax);
    if (!Number.isFinite(rowMin) || !Number.isFinite(rowMax) || rowMax < rowMin) {
      sendJson(res, { error: 'rowMin and rowMax required' }, 400);
      return true;
    }
    try {
      const raw = loadSheet(cellsMatch[1]);
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

const server = http.createServer((req, res) => {
  const fullUrl = req.url || '/';
  const urlPath = fullUrl.split('?')[0];
  if (urlPath.startsWith('/api/')) {
    if (handleApi(req, res, urlPath, fullUrl)) return;
    sendJson(res, { error: 'Not found' }, 404);
    return;
  }
  serveStatic(req, res, urlPath);
});

server.listen(PORT, HOST, () => {
  console.log(`WGHT app: http://${HOST}:${PORT}/`);
  console.log('Progressive API: GET /api/v1/config');
});
server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is already in use. Close the other server or set PORT=8080.`);
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
