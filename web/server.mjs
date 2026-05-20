/**
 * Local static server for the BD page (no npm required).
 * Run: node web/server.mjs
 */
import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = 5173;
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.ico': 'image/x-icon',
};
const server = http.createServer((req, res) => {
  let urlPath = req.url?.split('?')[0] || '/';
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
    const headers = {
      'Content-Type': MIME[ext] || 'application/octet-stream',
      'Cache-Control': noStore ? 'no-store, no-cache, must-revalidate' : 'public, max-age=3600',
    };
    if (noStore) headers.Pragma = 'no-cache';
    res.writeHead(200, headers);
    res.end(data);
  });
});
const HOST = '127.0.0.1';
server.listen(PORT, HOST, () => {
  console.log(`BD page: http://${HOST}:${PORT}/`);
});
server.on('error', (err) => {
  if (err.code === 'EADDRINUSE') {
    console.error(`Port ${PORT} is already in use. Close the other server or use run-bd-server.bat.`);
  } else {
    console.error(err.message);
  }
  process.exit(1);
});
