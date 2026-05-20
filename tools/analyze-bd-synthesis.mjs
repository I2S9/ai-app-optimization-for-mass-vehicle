/**
 * Analyze BD and SYNTHESIS sheets from extracted xlsm OOXML.
 * Usage: node tools/analyze-bd-synthesis.mjs
 */
import fs from 'fs';
import path from 'path';

const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-analyze');
const WS = path.join(TMP, 'xl', 'worksheets');
const SHARED = path.join(TMP, 'xl', 'sharedStrings.xml');

function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function numToCol(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function parseCellRef(ref) {
  const m = ref.match(/^([A-Z]+)(\d+)$/);
  return m ? { col: m[1], row: parseInt(m[2], 10), colNum: colToNum(m[1]) } : null;
}

function loadSharedStrings() {
  if (!fs.existsSync(SHARED)) return [];
  const xml = fs.readFileSync(SHARED, 'utf8');
  const strings = [];
  const re = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const t = m[1];
    const parts = [];
    const tre = /<t[^>]*>([^<]*)<\/t>/g;
    let tm;
    while ((tm = tre.exec(t)) !== null) parts.push(tm[1]);
    strings.push(parts.join('') || t.replace(/<[^>]+>/g, ''));
  }
  return strings;
}

function parseSheet(filePath, sharedStrings, opts = {}) {
  const xml = fs.readFileSync(filePath, 'utf8');
  const dimM = xml.match(/<dimension ref="([^"]+)"/);
  const cells = new Map();

  const cellRe = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
  let m;
  while ((m = cellRe.exec(xml)) !== null) {
    const ref = m[1];
    const attrs = m[2];
    const inner = m[3];
    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    const s = attrs.match(/\bs="(\d+)"/)?.[1];
    let value = null;
    let formula = null;
    const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
    if (fM) formula = fM[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) {
      if (t === 's') value = sharedStrings[parseInt(vM[1], 10)] ?? vM[1];
      else value = vM[1];
    }
    const is = inner.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
    if (is) value = is[1];

    cells.set(ref, { ref, formula, value, type: t });
  }

  return { dimension: dimM?.[1], cells };
}

function analyzeColumnStats(cells, maxRow, colLetters) {
  const stats = {};
  for (const col of colLetters) {
    let formula = 0, value = 0, empty = 0, sampleFormula = null, sampleValue = null;
    for (let r = 2; r <= maxRow; r++) {
      const ref = `${col}${r}`;
      const c = cells.get(ref);
      if (!c) { empty++; continue; }
      if (c.formula) {
        formula++;
        if (!sampleFormula) sampleFormula = c.formula.slice(0, 120);
      } else if (c.value != null && c.value !== '') {
        value++;
        if (!sampleValue) sampleValue = String(c.value).slice(0, 80);
      } else empty++;
    }
    stats[col] = { formula, value, empty, sampleFormula, sampleValue };
  }
  return stats;
}

function getRowCells(cells, row, cols) {
  const out = {};
  for (const col of cols) {
    const c = cells.get(`${col}${row}`);
    if (c) out[col] = c.formula ? `=${c.formula}` : c.value;
  }
  return out;
}

const shared = loadSharedStrings();
const bd = parseSheet(path.join(WS, 'sheet3.xml'), shared);
const syn = parseSheet(path.join(WS, 'sheet4.xml'), shared);

// All columns in BD dimension A1:AV3539
const bdCols = [];
for (let i = 1; i <= colToNum('AV'); i++) bdCols.push(numToCol(i));

// Header rows 1-5 for BD
console.log('=== BD HEADER ROWS 1-5 (sample columns A-O, key cols) ===');
const bdSampleCols = bdCols.slice(0, 48);
for (let r = 1; r <= 8; r++) {
  const row = getRowCells(bd.cells, r, bdSampleCols);
  const nonEmpty = Object.entries(row).filter(([, v]) => v != null && v !== '');
  if (nonEmpty.length) console.log(`Row ${r}:`, JSON.stringify(Object.fromEntries(nonEmpty.slice(0, 20))));
}

console.log('\n=== BD COLUMN STATS (rows 2-100, all AV cols) ===');
const bdStats100 = analyzeColumnStats(bd.cells, 100, bdCols);
const bdManual = [];
const bdCalc = [];
for (const [col, s] of Object.entries(bdStats100)) {
  if (s.formula > 0 && s.value === 0) bdCalc.push(col);
  else if (s.value > 0 && s.formula === 0) bdManual.push(col);
  else if (s.formula > 0 && s.value > 0) bdCalc.push(col + '*');
}
console.log('Mostly manual (value only rows 2-100):', bdManual.join(', '));
console.log('Mostly formula rows 2-100:', bdCalc.join(', '));

console.log('\n=== BD COLUMN STATS (rows 2-500) ===');
const bdStats500 = analyzeColumnStats(bd.cells, 500, bdCols);
const summary500 = [];
for (const col of bdCols) {
  const s = bdStats500[col];
  if (!s || (s.formula === 0 && s.value === 0)) continue;
  const pctF = ((s.formula / 499) * 100).toFixed(0);
  summary500.push({ col, formula: s.formula, value: s.value, pctFormula: pctF + '%', sample: (s.sampleFormula || s.sampleValue || '').slice(0, 80) });
}
summary500.sort((a, b) => b.formula - a.formula);
console.log(JSON.stringify(summary500.slice(0, 35), null, 2));

// SYNTHESIS structure - first rows
console.log('\n=== SYNTHESIS DIMENSION ===', syn.dimension);
console.log('=== SYNTHESIS HEADER ROWS 1-15 (cols A-AH sample) ===');
const synCols = [];
for (let i = 1; i <= colToNum('NI'); i++) synCols.push(numToCol(i));
const synHeaderCols = synCols.filter((_, i) => i < 80);
for (let r = 1; r <= 14; r++) {
  const row = getRowCells(syn.cells, r, synHeaderCols);
  const nonEmpty = Object.entries(row).filter(([, v]) => v != null && v !== '');
  if (nonEmpty.length > 0) console.log(`Row ${r} (${nonEmpty.length} cells):`, JSON.stringify(Object.fromEntries(nonEmpty.slice(0, 12))));
}

// SYNTHESIS formula patterns rows 10-20
console.log('\n=== SYNTHESIS SAMPLE FORMULAS row 14 ===');
const r14 = getRowCells(syn.cells, 14, synCols.slice(0, 30));
for (const [col, v] of Object.entries(r14)) {
  if (String(v).startsWith('=')) console.log(col, String(v).slice(0, 150));
}

// Count BD refs in synthesis formulas (sample)
let bdRefCount = 0;
let sumproductCount = 0;
const synFormulas = [];
for (const [, c] of syn.cells) {
  if (c.formula) {
    synFormulas.push(c.formula);
    if (c.formula.includes('BD!')) bdRefCount++;
    if (/SUMPRODUCT/i.test(c.formula)) sumproductCount++;
  }
}
console.log('\n=== SYNTHESIS FORMULA SUMMARY ===');
console.log('Total cells with formula:', synFormulas.length);
console.log('Formulas referencing BD!:', bdRefCount);
console.log('SUMPRODUCT formulas:', sumproductCount);
console.log('Unique formula prefixes (first 200 chars, top 5):');
const prefixes = {};
for (const f of synFormulas.slice(0, 500)) {
  const p = f.slice(0, 60);
  prefixes[p] = (prefixes[p] || 0) + 1;
}
console.log(Object.entries(prefixes).sort((a, b) => b[1] - a[1]).slice(0, 8));

// BD data row sample row 10
console.log('\n=== BD DATA ROW 10 (all non-empty) ===');
const r10 = getRowCells(bd.cells, 10, bdCols);
console.log(JSON.stringify(Object.fromEntries(Object.entries(r10).filter(([, v]) => v != null && v !== '')), null, 2));

// Styles reference - sheet views
const bdXml = fs.readFileSync(path.join(WS, 'sheet3.xml'), 'utf8');
const freeze = bdXml.match(/<pane[^>]*>/g);
console.log('\n=== BD freeze/sheetView ===', freeze?.slice(0, 3));

// merged cells BD
const merges = [...bdXml.matchAll(/<mergeCell ref="([^"]+)"/g)].map((m) => m[1]);
console.log('BD merged ranges count:', merges.length);
console.log('BD merged (first 25):', merges.slice(0, 25));

const synXml = fs.readFileSync(path.join(WS, 'sheet4.xml'), 'utf8');
const synMerges = [...synXml.matchAll(/<mergeCell ref="([^"]+)"/g)].map((m) => m[1]);
console.log('SYNTHESIS merged ranges count:', synMerges.length);
console.log('SYNTHESIS merged (first 20):', synMerges.slice(0, 20));

// Row count with data in BD col A
let bdDataRows = 0;
for (let r = 2; r <= 3539; r++) {
  if (bd.cells.has(`A${r}`) || bd.cells.has(`B${r}`) || bd.cells.has(`O${r}`)) bdDataRows++;
}
console.log('\nBD rows with signal in A/B/O up to 3539:', bdDataRows);
