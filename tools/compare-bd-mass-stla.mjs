/**
 * Compare Excel vs JSON Mass (V) on STLA/S rows only.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { transformBdSheet } from '../web/js/sheetTransform.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-bd-export');
const sheetXml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet3.xml'), 'utf8');
const shared = [];
for (const m of fs
  .readFileSync(path.join(TMP, 'xl/sharedStrings.xml'), 'utf8')
  .matchAll(/<si>([\s\S]*?)<\/si>/g)) {
  const parts = [];
  for (const t of m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)) parts.push(t[1]);
  shared.push(parts.join(''));
}

function cellVal(col, row) {
  const re = new RegExp(
    `<c r="${col}${row}"([^/]*?)(\\/>|>([\\s\\S]*?)<\\/c>)`
  );
  const m = sheetXml.match(re);
  if (!m) return '';
  const inner = m[2] === '/>' ? '' : m[3] || '';
  const t = m[1].match(/\bt="([^"]+)"/)?.[1];
  const vM = inner.match(/<v>([^<]*)<\/v>/);
  if (!vM) return '';
  const raw = vM[1];
  if (t === 's') return shared[+raw] ?? raw;
  return raw;
}

const raw = JSON.parse(
  fs.readFileSync(path.join(__dirname, '../web/public/data/bd-sheet.json'), 'utf8')
);
const sheet = transformBdSheet(raw);
const byRow = new Map();
for (const c of sheet.cells) {
  if (c.c === 'V') byRow.set(c.r, c.v != null ? String(c.v) : '');
}

let mism = 0;
let checked = 0;
for (let r = 6; r <= sheet.lastRow; r++) {
  const a = cellVal('A', r);
  if (a !== 'STLA/S' && a !== 'STLA-S') continue;
  checked++;
  const ex = cellVal('V', r);
  const json = byRow.get(r) ?? '';
  const nEx = Number(ex);
  const nJson = Number(json);
  const same =
    ex === json ||
    (!Number.isNaN(nEx) && !Number.isNaN(nJson) && nEx === nJson);
  if (!same) {
    mism++;
    if (mism <= 10) console.log('row', r, { excel: ex, json });
  }
}
console.log('STLA rows checked', checked, 'mismatches', mism);
