/**
 * Compare SYNTHESIS row band colors in Excel vs web label-based classes.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const WB = path.join(ROOT, 'workbooks', 'base-de-donnees-complete-avec-liens.xlsm');
const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-syn-colors');
const MAX_ROW = 530;
const LABEL_COL = 'F';
const SCAN_COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];

const INDEXED_COLORS = {
  0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00', 4: '#0000FF',
  5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF', 13: '#FFFF00', 40: '#FFFFCC',
  41: '#CCFFFF', 44: '#CCFFCC', 46: '#99CCFF', 64: '#333399',
};

const THEME_FILLS = [
  '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31',
  '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72',
];

function colToNum(col) {
  let n = 0;
  for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64);
  return n;
}

function resolveColor(node) {
  if (!node) return null;
  const rgb = node.match(/rgb="([^"]+)"/)?.[1];
  if (rgb) {
    const hex = rgb.length === 8 ? rgb.slice(2) : rgb;
    return `#${hex}`.toLowerCase();
  }
  const indexed = node.match(/indexed="(\d+)"/)?.[1];
  if (indexed != null) return (INDEXED_COLORS[indexed] ?? null)?.toLowerCase();
  const theme = node.match(/theme="(\d+)"/)?.[1];
  const tint = parseFloat(node.match(/tint="([^"]+)"/)?.[1] ?? '0');
  if (theme != null) {
    const base = THEME_FILLS[parseInt(theme, 10)] || null;
    if (!base || tint === 0) return base?.toLowerCase() ?? null;
    return base.toLowerCase();
  }
  return null;
}

function parseStyles(stylesXml) {
  const fills = [{ fg: null }];
  const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
  let fm;
  while ((fm = fillRe.exec(stylesXml)) !== null) {
    const block = fm[1];
    const fg =
      resolveColor(block.match(/<fgColor[^>]*\/>/)?.[0]) ||
      resolveColor(block.match(/<fgColor[^>]*>[\s\S]*?<\/fgColor>/)?.[0]);
    const bg =
      resolveColor(block.match(/<bgColor[^>]*\/>/)?.[0]) ||
      resolveColor(block.match(/<bgColor[^>]*>[\s\S]*?<\/bgColor>/)?.[0]);
    fills.push({ fg: (fg || bg)?.toLowerCase() });
  }
  const xfs = [];
  const xfRe = /<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g;
  let xfm;
  while ((xfm = xfRe.exec(stylesXml)) !== null) {
    const attrs = xfm[1];
    const fillId = parseInt(attrs.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
    const applyFill =
      fillId > 0 || attrs.includes('applyFill="1"') || attrs.includes('applyFill="true"');
    const f = fills[fillId] || {};
    xfs.push(applyFill && f.fg ? f.fg : null);
  }
  return xfs;
}

function parseSharedStrings(xml) {
  const strings = [];
  const re = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const parts = [];
    const tre = /<t[^>]*>([^<]*)<\/t>/g;
    let tm;
    while ((tm = tre.exec(m[1])) !== null) parts.push(tm[1]);
    strings.push(parts.join('') || m[1].replace(/<[^>]+>/g, ''));
  }
  return strings;
}

function normHex(h) {
  if (!h) return null;
  const x = h.toLowerCase();
  if (x === '#ffff00' || x === '#fff2cc' || x === '#ffc000' || x === '#ffffcc') return 'yellow';
  if (
    x === '#00b0f0' || x === '#9bc2e6' || x === '#bdd7ee' || x === '#ddebf7' ||
    x === '#4472c4' || x === '#5b9bd5' || x === '#8db4e2' || x === '#ccffff' ||
    x === '#99ccff' || x === '#0563c1'
  ) return 'blue';
  if (x === '#1f3864' || x === '#203864' || x === '#002060' || x === '#44546a') return 'separator';
  return 'other';
}

function webClass(label, row) {
  if (row === 4) return 'syn-row-separator';
  if (row >= 3 && row <= 14) return 'syn-row-filter';
  if (!label) return 'syn-row-data';
  const t = String(label).trim();
  if (t.startsWith('-')) return 'syn-row-section';
  if (t.startsWith('_')) return 'syn-row-subsection';
  return 'syn-row-data';
}

function excelClass(bg) {
  const k = normHex(bg);
  if (k === 'yellow') return 'syn-row-section';
  if (k === 'blue') return 'syn-row-subsection';
  if (k === 'separator') return 'syn-row-separator';
  return 'syn-row-data';
}

async function main() {
  if (!fs.existsSync(WB)) {
    console.error('Workbook not found:', WB);
    process.exit(1);
  }
  const { execSync } = await import('child_process');
  const sheetPath = path.join(TMP, 'xl', 'worksheets', 'sheet4.xml');
  if (!fs.existsSync(sheetPath)) {
    fs.mkdirSync(TMP, { recursive: true });
    const zip = path.join(TMP, 'base.zip');
    fs.copyFileSync(WB, zip);
    execSync(
      `powershell -NoProfile -Command "Expand-Archive -LiteralPath '${zip.replace(/'/g, "''")}' -DestinationPath '${TMP.replace(/'/g, "''")}' -Force"`,
      { stdio: 'inherit' }
    );
  }

  const shared = parseSharedStrings(fs.readFileSync(path.join(TMP, 'xl', 'sharedStrings.xml'), 'utf8'));
  const styles = parseStyles(fs.readFileSync(path.join(TMP, 'xl', 'styles.xml'), 'utf8'));
  const sheetXml = fs.readFileSync(sheetPath, 'utf8');

  const rowBg = new Map();
  const rowLabel = new Map();
  const cellRe = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
  let m;
  while ((m = cellRe.exec(sheetXml)) !== null) {
    const ref = m[1];
    const col = ref.replace(/\d+$/, '');
    const row = parseInt(ref.replace(/^[A-Z]+/, ''), 10);
    if (row > MAX_ROW || !SCAN_COLS.includes(col)) continue;
    const attrs = m[2];
    const inner = m[3];
    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    const styleIdx = parseInt(attrs.match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    let value = null;
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = t === 's' ? shared[parseInt(vM[1], 10)] ?? vM[1] : vM[1];
    const is = inner.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
    if (is) value = is[1];
    const bg = styleIdx >= 0 ? styles[styleIdx] : null;
    if (col === LABEL_COL && value) rowLabel.set(row, String(value).trim());
    if (col === 'D' && value) {
      const v = String(value).trim();
      if ((v.startsWith('-') || v.startsWith('_')) && !rowLabel.has(row)) rowLabel.set(row, v);
    }
    if (bg && !rowBg.has(row)) rowBg.set(row, bg);
    if (bg && col === LABEL_COL) rowBg.set(row, bg);
  }

  const mismatches = [];
  const labelNoBg = [];
  for (let row = 15; row <= MAX_ROW; row++) {
    const label = rowLabel.get(row) || '';
    const bg = rowBg.get(row);
    const excel = excelClass(bg);
    const web = webClass(label, row);
    if (web === 'syn-row-filter' || web === 'syn-row-separator') continue;
    if (web !== 'syn-row-data' && excel === 'syn-row-data') {
      labelNoBg.push({ row, label, bg, web });
    }
    if (excel !== 'syn-row-data' && excel !== web) {
      mismatches.push({ row, label, bg, excel, web });
    }
  }

  console.log('Label rows without detected Excel band color:', labelNoBg.length);
  for (const x of labelNoBg.slice(0, 25)) console.log(JSON.stringify(x));

  console.log('\nExcel color vs web class mismatches:', mismatches.length);
  for (const x of mismatches.slice(0, 25)) console.log(JSON.stringify(x));

  const samples = [25, 26, 41, 42, 100, 200, 421, 422];
  console.log('\nSample rows (bg from any A-H cell):');
  for (const row of samples) {
    console.log(row, 'label=', rowLabel.get(row), 'bg=', rowBg.get(row), 'web=', webClass(rowLabel.get(row), row));
  }

  const byBg = {};
  for (const [, bg] of rowBg) {
    const k = normHex(bg) || bg || 'none';
    byBg[k] = (byBg[k] || 0) + 1;
  }
  console.log('\nRow fill buckets (first colored cell A-H):', byBg);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
