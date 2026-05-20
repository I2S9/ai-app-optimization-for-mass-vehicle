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
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function loadSharedStrings() {
  const xml = fs.readFileSync(SHARED, 'utf8');
  const strings = [];
  const re = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const parts = [];
    const tre = /<t[^>]*>([^<]*)<\/t>/g;
    let tm;
    while ((tm = tre.exec(m[1])) !== null) parts.push(tm[1]);
    strings.push(parts.join(''));
  }
  return strings;
}
function parseSheet(filePath, sharedStrings) {
  const xml = fs.readFileSync(filePath, 'utf8');
  const cells = new Map();
  const cellRe = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
  let m;
  while ((m = cellRe.exec(xml)) !== null) {
    const ref = m[1];
    const attrs = m[2];
    const inner = m[3];
    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    let value = null, formula = null;
    const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
    if (fM) formula = fM[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = t === 's' ? sharedStrings[parseInt(vM[1], 10)] ?? vM[1] : vM[1];
    cells.set(ref, { formula, value });
  }
  const cols = xml.match(/<cols>([\s\S]*?)<\/cols>/);
  const colWidths = [];
  if (cols) {
    const re = /<col min="(\d+)" max="(\d+)"[^>]*width="([^"]+)"/g;
    let cm;
    while ((cm = re.exec(cols[1])) !== null) {
      for (let i = +cm[1]; i <= +cm[2]; i++)
        colWidths.push({ col: numToCol(i), width: +cm[3] });
    }
  }
  return { cells, colWidths, xml };
}
const shared = loadSharedStrings();
const bd = parseSheet(path.join(WS, 'sheet3.xml'), shared);
const syn = parseSheet(path.join(WS, 'sheet4.xml'), shared);
const bdCols = [];
for (let i = 1; i <= colToNum('AV'); i++) bdCols.push(numToCol(i));
// Full row 1 headers
console.log('=== BD ROW 1 — ALL COLUMN HEADERS (A–AV) ===');
const h1 = {};
for (const c of bdCols) {
  const cell = bd.cells.get(`${c}1`);
  if (cell?.value || cell?.formula) h1[c] = cell.formula ? '=' + cell.formula : cell.value;
}
console.log(JSON.stringify(h1, null, 2));
// Rows 2-4 legend rows in col A
console.log('\n=== BD COLUMN A rows 1-6 (legend) ===');
for (let r = 1; r <= 6; r++) console.log(r, bd.cells.get(`A${r}`)?.value);
// Extended headers row 2 labels in A?
console.log('\n=== BD ROW 2 labels in column A + key cols U-AF ===');
for (let r = 1; r <= 4; r++) {
  const parts = [];
  for (const c of ['A','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT','AU','AV']) {
    const cell = bd.cells.get(`${c}${r}`);
    if (cell) parts.push(`${c}:${cell.formula ? 'F' : cell.value}`);
  }
  console.log('Row', r, parts.join(' | '));
}
// Find first data row (component lines) - row where S has decoupage code
console.log('\n=== BD typical component row 100 ===');
for (const c of bdCols) {
  const cell = bd.cells.get(`A100`) || bd.cells.get(`${c}100`);
}
const r100 = {};
for (const c of bdCols) {
  const cell = bd.cells.get(`${c}100`);
  if (cell) r100[c] = cell.formula ? `=${cell.formula.slice(0,60)}` : cell.value;
}
console.log(JSON.stringify(r100, null, 2));
// Count last row with data in column V
let lastV = 0;
for (let r = 3500; r >= 2; r--) {
  if (bd.cells.has(`V${r}`)) { lastV = r; break; }
}
console.log('\nLast row with V cell:', lastV);
// SYNTHESIS: decode one SUMPRODUCT at F27 or similar
for (const ref of ['G27', 'I27', 'G50', 'G100', 'G200']) {
  const c = syn.cells.get(ref);
  if (c?.formula) {
    console.log(`\n=== SYNTHESIS ${ref} formula (first 500 chars) ===`);
    console.log(c.formula.slice(0, 500));
  }
}
// SYNTHESIS row labels column F rows 15-80
console.log('\n=== SYNTHESIS column F (labels) rows 15-45 ===');
for (let r = 15; r <= 45; r++) {
  const c = syn.cells.get(`F${r}`);
  if (c?.value || c?.formula) console.log(r, c.value || c.formula?.slice(0,40));
}
// SYNTHESIS structure: how many "vehicle columns" - count headers row 10 with SR/HR/XR
console.log('\n=== SYNTHESIS row 10 (range codes) non-empty count ===');
let r10count = 0;
for (let i = 1; i <= colToNum('NI'); i++) {
  const col = numToCol(i);
  if (syn.cells.get(`${col}10`)?.value) r10count++;
}
console.log('Cells in row 10:', r10count);
// BD columns used in SUMPRODUCT - grep first formula
const f27 = syn.cells.get('G27')?.formula || '';
const bdColsInFormula = [...new Set([...f27.matchAll(/BD!\$([A-Z]+)\$/g)].map((m) => m[1]))];
console.log('\nBD columns referenced in G27:', bdColsInFormula);
// SYNTHESIS editable cells? rows 3-8 seem config - count formulas vs values in row 27 area
let synEditable = 0, synFormula = 0;
for (const [ref, cell] of syn.cells) {
  const m = ref.match(/^[A-Z]+(\d+)$/);
  if (!m) continue;
  const row = +m[1];
  if (row >= 3 && row <= 13) {
    if (cell.formula) synFormula++;
    else if (cell.value) synEditable++;
  }
}
console.log('\nSYNTHESIS rows 3-13: values (likely editable config):', synEditable, 'formulas:', synFormula);
// Row types in synthesis - sample row 27 labels
console.log('\n=== SYNTHESIS row 27 headers F, G values ===');
console.log('F27', syn.cells.get('F27')?.value, syn.cells.get('F27')?.formula?.slice(0,80));
console.log('G27 formula exists:', !!syn.cells.get('G27')?.formula);
// BD col widths first 30
console.log('\n=== BD column widths (first 48 cols) ===');
console.log(JSON.stringify(bd.colWidths.slice(0, 48), null, 2));
