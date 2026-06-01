/**
 * Build precomputed grid JSON from raw exports (run after export-bd/synthesis-sheet).
 * Eliminates ~5 s browser transform on cold start.
 *
 *   node tools/build-grid-sheets.mjs
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { transformBdSheet, transformSynthesisSheet } from '../web/js/sheetTransform.js';
import { rawFingerprint, serializeTransformSheet } from '../web/js/sessionPersistence.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA = path.join(__dirname, '..', 'web', 'public', 'data');

function buildOne(sheetId, rawName, outName, transform) {
  const rawPath = path.join(DATA, rawName);
  const outPath = path.join(DATA, outName);
  if (!fs.existsSync(rawPath)) {
    console.warn(`Skip ${sheetId}: missing ${rawPath}`);
    return;
  }
  console.log(`Transforming ${sheetId}…`);
  const t0 = Date.now();
  const raw = JSON.parse(fs.readFileSync(rawPath, 'utf8'));
  const fp = rawFingerprint(raw);
  const sheet = transform(raw);
  const pack = {
    version: 1,
    fingerprint: fp,
    builtAt: new Date().toISOString(),
    sheet: serializeTransformSheet(sheet),
  };
  fs.writeFileSync(outPath, JSON.stringify(pack));
  const mb = (fs.statSync(outPath).size / (1024 * 1024)).toFixed(1);
  console.log(`  → ${outName} (${mb} MB) in ${((Date.now() - t0) / 1000).toFixed(1)}s`);
}

buildOne('bd', 'bd-sheet.json', 'bd-sheet-grid.json', transformBdSheet);
buildOne('syn', 'synthesis-sheet.json', 'synthesis-sheet-grid.json', transformSynthesisSheet);
console.log('Done.');
