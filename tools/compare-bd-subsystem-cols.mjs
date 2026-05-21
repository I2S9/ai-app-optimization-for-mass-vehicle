/**
 * Compare Excel AP/AS/AU vs web transform for subsystem columns.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-bd-export');
const WB = path.join(ROOT, 'workbooks', 'base-de-donnees-complete-avec-liens.xlsm');

function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function parseSharedStrings(xml) {
  const strings = [];
  for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    const parts = [];
    for (const t of m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)) parts.push(t[1]);
    strings.push(parts.join(''));
  }
  return strings;
}

function resolveCellValue(raw, t, shared) {
  if (raw == null || raw === '') return raw;
  if (t === 's') return shared[+raw] ?? String(raw);
  const s = String(raw).trim();
  if (/^\d+$/.test(s)) {
    const hit = shared[+s];
    if (hit != null && hit !== '' && !/^-?\d+(\.\d+)?$/.test(String(hit).trim())) {
      return hit;
    }
  }
  return s;
}

function parseFormula(inner, sharedFormulas) {
  const open = inner.match(/<f([^>]*)>([\s\S]*?)<\/f>/);
  if (open) {
    const attrs = open[1];
    const body = open[2].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
    const si = attrs.match(/\bsi="(\d+)"/)?.[1];
    if (attrs.includes('t="shared"') && si != null) {
      if (body) sharedFormulas.set(si, body);
      return body || sharedFormulas.get(si) || null;
    }
    return body || null;
  }
  const self = inner.match(/<f([^/]*)\/>/);
  if (!self) return null;
  const attrs = self[1];
  const si = attrs.match(/\bsi="(\d+)"/)?.[1];
  if (attrs.includes('t="shared"') && si != null) {
    return sharedFormulas.get(si) || null;
  }
  return null;
}

function parseExcelCol(sheetXml, shared, colLetter, rowMin, rowMax) {
  const sharedFormulas = new Map();
  const re = new RegExp(
    `<c r="${colLetter}(\\d+)"([^/]*?)(\\/>|>([\\s\\S]*?)<\\/c>)`,
    'g'
  );
  const rows = new Map();
  let m;
  while ((m = re.exec(sheetXml)) !== null) {
    const row = +m[1];
    if (row < rowMin || row > rowMax) continue;
    const inner = m[3] === '/>' ? '' : m[4] || '';
    const t = m[2].match(/\bt="([^"]+)"/)?.[1];
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    const formula = parseFormula(inner, sharedFormulas);
    let v = vM ? resolveCellValue(vM[1], t, shared) : '';
    rows.set(row, { v: v != null ? String(v) : '', f: formula || '' });
  }
  return rows;
}

async function ensureExtract() {
  const sheetPath = path.join(TMP, 'xl/worksheets/sheet3.xml');
  if (fs.existsSync(sheetPath)) return;
  if (!fs.existsSync(WB)) {
    console.error('Workbook missing:', WB);
    process.exit(1);
  }
  const { execSync } = await import('child_process');
  fs.mkdirSync(TMP, { recursive: true });
  const zip = path.join(TMP, 'base.zip');
  fs.copyFileSync(WB, zip);
  execSync(
    `powershell -NoProfile -Command "Expand-Archive -LiteralPath '${zip.replace(/'/g, "''")}' -DestinationPath '${TMP.replace(/'/g, "''")}' -Force"`,
    { stdio: 'inherit' }
  );
}

function webDisplay(map, row, col, l1Col, l2Col, designCol) {
  const cell = map.get(`${row}:${col}`);
  if (!cell) return '';
  let v = cell.v != null && cell.v !== '' ? String(cell.v) : '';
  if (cell.f && !v) v = '';
  if (v === '#REF!' || v === '#NAME?' || v.startsWith('#')) v = '';
  return v;
}

await ensureExtract();
const sheetXml = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet3.xml'), 'utf8');
const shared = parseSharedStrings(
  fs.readFileSync(path.join(TMP, 'xl/sharedStrings.xml'), 'utf8')
);

const excelAP = parseExcelCol(sheetXml, shared, 'AP', 6, 200);
const excelAS = parseExcelCol(sheetXml, shared, 'AS', 6, 200);
const excelAU = parseExcelCol(sheetXml, shared, 'AU', 6, 200);

const { transformBdSheet } = await import('../web/js/sheetTransform.js');
const { buildCellMap, displayCellValue, bdSubsystemL1Col, bdSubsystemL2Col } =
  await import('../web/js/bdStore.js');
const raw = JSON.parse(
  fs.readFileSync(path.join(ROOT, 'web/public/data/bd-sheet.json'), 'utf8')
);
const t = transformBdSheet(raw);
const map = buildCellMap(t.cells, t.headerRows);
const l1 = bdSubsystemL1Col(t);
const l2 = bdSubsystemL2Col(t);
const design = 'AW';
const sh = t.sectionHeaderRows;
const canon = t.canonicalSectionByLabel;

let apDiff = 0;
let asDiff = 0;
let auDiff = 0;
for (let r = 6; r <= 200; r++) {
  const exAp = excelAP.get(r)?.v || '';
  const exAs = excelAS.get(r)?.v || '';
  const exAu = excelAU.get(r)?.v || '';
  const webL1 = displayCellValue(map, r, l1, sh, canon, l1, l2);
  const webL2 = displayCellValue(map, r, l2, sh, canon, l1, l2);
  const webAu = displayCellValue(map, r, design, sh, canon, l1, l2);
  const rawL1 = map.get(`${r}:${l1}`)?.v || '';
  const rawL2 = map.get(`${r}:${l2}`)?.v || '';
  const rawAu = map.get(`${r}:${design}`)?.v || '';

  if (exAp && exAp !== webL1 && exAp !== rawL1) {
    if (apDiff < 15) console.log('AP', r, 'excel:', exAp, 'web:', webL1, 'raw:', rawL1);
    apDiff++;
  }
  if (exAs && exAs !== webL2 && exAs !== rawL2) {
    if (asDiff < 20) console.log('AS', r, 'excel:', exAs, 'web:', webL2, 'raw:', rawL2, 'f:', map.get(`${r}:${l2}`)?.f?.slice(0, 40));
    asDiff++;
  }
  if (exAu && exAu !== webAu && exAu !== rawAu) {
    if (auDiff < 20) console.log('AU', r, 'excel:', exAu, 'web:', webAu, 'raw:', rawAu);
    auDiff++;
  }
}
console.log('diff counts AP', apDiff, 'AS', asDiff, 'AU', auDiff);
console.log('cols', { l1, l2, design });
console.log('sample excel AS 7-15:', [...excelAS.entries()].filter(([r]) => r >= 7 && r <= 15).map(([r, c]) => ({ r, v: c.v })));
console.log('sample excel AU 26-35:', [...excelAU.entries()].filter(([r]) => r >= 26 && r <= 35).map(([r, c]) => ({ r, v: c.v, f: c.f?.slice(0, 35) })));
console.log('sample excel AP 6-20:', [...excelAP.entries()].filter(([r]) => r >= 6 && r <= 20).map(([r, c]) => ({ r, v: c.v, f: c.f?.slice(0, 50) })));
const nameCount = [...t.cells].filter((c) => c.c === l1 && String(c.v || '').includes('NAME')).length;
console.log('#NAME in raw L1', nameCount, t.cells.filter((c) => c.c === l1 && String(c.v || '').includes('NAME')).slice(0, 5));
