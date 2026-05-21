/**
 * Compare Excel column V (Masse/Mass) vs bd-sheet.json cached v vs HyperFormula.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { transformBdSheet } from '../web/js/sheetTransform.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
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

function resolveCellValue(raw, t) {
  if (raw == null || raw === '') return '';
  if (t === 's') return shared[+raw] ?? String(raw);
  const s = String(raw).trim();
  if (/^\d+$/.test(s)) {
    const hit = shared[+s];
    if (hit != null && hit !== '' && !/^-?\d+(\.\d+)?$/.test(String(hit).trim())) return hit;
  }
  return s;
}

const excel = new Map();
const re = /<c r="V(\d+)"([^/]*?)(\/>|>([\s\S]*?)<\/c>)/g;
let m;
while ((m = re.exec(sheetXml)) !== null) {
  const row = +m[1];
  if (row < 6) continue;
  const inner = m[3] === '/>' ? '' : m[4] || '';
  const t = m[2].match(/\bt="([^"]+)"/)?.[1];
  const vM = inner.match(/<v>([^<]*)<\/v>/);
  const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
  const v = vM ? resolveCellValue(vM[1], t) : '';
  excel.set(row, { v: String(v), f: fM?.[1]?.slice(0, 40) || '' });
}

const raw = JSON.parse(
  fs.readFileSync(path.join(ROOT, 'web/public/data/bd-sheet.json'), 'utf8')
);
const transformed = transformBdSheet(raw);
const byRow = new Map();
for (const c of transformed.cells) {
  if (c.c === 'V') byRow.set(c.r, c);
}

function normMass(v) {
  if (v == null || v === '') return '';
  const t = String(v).trim();
  if (t.startsWith('#')) return t;
  const n = Number(t);
  if (!Number.isNaN(n) && t !== '') return String(Math.round(n * 1e6) / 1e6);
  return t;
}

let excelVsJson = 0;
let missingInJson = 0;
let samples = [];
for (const [row, ex] of excel) {
  const cell = byRow.get(row);
  const jsonV = cell?.v != null ? String(cell.v) : '';
  const nEx = normMass(ex.v);
  const nJson = normMass(jsonV);
  if (nEx !== nJson) {
    excelVsJson++;
    if (samples.length < 12)
      samples.push({
        row,
        excel: nEx,
        json: nJson,
        f: cell?.f?.slice(0, 40) || ex.f,
      });
  }
  if (!cell && nEx !== '') missingInJson++;
}

console.log('Excel V cells', excel.size);
console.log('JSON V cells', byRow.size);
console.log('Excel vs JSON mismatches', excelVsJson);
console.log('Excel value missing in JSON', missingInJson);
console.log('samples', samples);
