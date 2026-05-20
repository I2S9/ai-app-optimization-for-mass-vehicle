/**
 * Build rowBands map from Excel SYNTHESIS column F fill colors.
 * Writes rowBands into synthesis-sheet.json export payload.
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

const INDEXED = {
  0: '#000000', 1: '#FFFFFF', 2: '#FF0000', 3: '#00FF00', 4: '#0000FF',
  5: '#FFFF00', 6: '#FF00FF', 7: '#00FFFF', 8: '#000000', 9: '#FFFFFF',
  10: '#FF0000', 11: '#00FF00', 12: '#0000FF', 13: '#FFFF00', 22: '#C0C0C0',
  40: '#FFFFCC', 41: '#CCFFFF', 44: '#CCFFCC', 45: '#FFCCFF', 46: '#99CCFF',
  64: '#333399',
};

const THEME = [
  '#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31',
  '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72',
];

function applyTint(hex, tint) {
  if (!tint) return hex;
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const f = (c) =>
    tint < 0
      ? Math.round(c * (1 + tint))
      : Math.round(c + (255 - c) * tint);
  const h = (n) => n.toString(16).padStart(2, '0');
  return `#${h(f(r))}${h(f(g))}${h(f(b))}`;
}

function resolveColorNode(node) {
  if (!node) return null;
  const rgb = node.match(/rgb="([^"]+)"/)?.[1];
  if (rgb) {
    const h = rgb.length === 8 ? rgb.slice(2) : rgb;
    return `#${h}`.toLowerCase();
  }
  const indexed = node.match(/indexed="(\d+)"/)?.[1];
  if (indexed != null) return INDEXED[indexed] ?? null;
  const theme = node.match(/theme="(\d+)"/)?.[1];
  if (theme != null) {
    const base = THEME[parseInt(theme, 10)];
    const tint = parseFloat(node.match(/tint="([^"]+)"/)?.[1] ?? '0');
    return base ? applyTint(base.toLowerCase(), tint) : null;
  }
  return null;
}

function parseStyles(stylesXml) {
  const fills = [{ fg: null }];
  const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
  let fm;
  while ((fm = fillRe.exec(stylesXml)) !== null) {
    const block = fm[1];
    if (/patternType="none"/.test(block)) {
      fills.push({ fg: null });
      continue;
    }
    const fg =
      resolveColorNode(block.match(/<fgColor[^>]*\/>/)?.[0]) ||
      resolveColorNode(block.match(/<fgColor[^>]*>[\s\S]*?<\/fgColor>/)?.[0]);
    const bg =
      resolveColorNode(block.match(/<bgColor[^>]*\/>/)?.[0]) ||
      resolveColorNode(block.match(/<bgColor[^>]*>[\s\S]*?<\/bgColor>/)?.[0]);
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
    xfs.push(applyFill ? fills[fillId]?.fg ?? null : null);
  }
  return xfs;
}

function bandFromBg(bg) {
  if (!bg) return null;
  const h = bg.toLowerCase();
  if (h === '#ffff00' || h === '#fff2cc' || h === '#ffc000' || h === '#ffffcc' || h === '#ccffcc') {
    return 'section';
  }
  if (
    h === '#00b0f0' || h === '#9bc2e6' || h === '#bdd7ee' || h === '#ddebf7' ||
    h === '#4472c4' || h === '#5b9bd5' || h === '#8db4e2' || h === '#99ccff' ||
    h === '#ccffff' || h === '#0563c1' || h === '#4f81bd'
  ) {
    return 'subsection';
  }
  if (h === '#1f3864' || h === '#203864' || h === '#002060' || h === '#44546a') {
    return 'separator';
  }
  return null;
}

function bandFromLabel(label, row) {
  if (row === 4) return 'separator';
  if (row >= 3 && row <= 14) return 'filter';
  if (!label) return null;
  const t = String(label).trim();
  if (t.startsWith('-')) return 'section';
  if (t.startsWith('_')) return 'subsection';
  return null;
}

async function main() {
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

  const shared = [];
  const ss = fs.readFileSync(path.join(TMP, 'xl', 'sharedStrings.xml'), 'utf8');
  const sre = /<si>([\s\S]*?)<\/si>/g;
  let sm;
  while ((sm = sre.exec(ss)) !== null) {
    const p = [];
    let t;
    const tr = /<t[^>]*>([^<]*)<\/t>/g;
    while ((t = tr.exec(sm[1])) !== null) p.push(t[1]);
    shared.push(p.join(''));
  }

  const xfs = parseStyles(fs.readFileSync(path.join(TMP, 'xl', 'styles.xml'), 'utf8'));
  const sheetXml = fs.readFileSync(sheetPath, 'utf8');

  const fStyle = new Map();
  const fLabel = new Map();
  const cellRe = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
  let m;
  while ((m = cellRe.exec(sheetXml)) !== null) {
    const col = m[1].replace(/\d+$/, '');
    const row = parseInt(m[1].replace(/^[A-Z]+/, ''), 10);
    if (row > MAX_ROW) continue;
    const attrs = m[2];
    const inner = m[3];
    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    const styleIdx = parseInt(attrs.match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    let value = null;
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = t === 's' ? shared[parseInt(vM[1], 10)] ?? vM[1] : vM[1];
    const is = inner.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
    if (is) value = is[1];
    if (col === LABEL_COL) {
      if (value) fLabel.set(row, String(value).trim());
      if (styleIdx >= 0) fStyle.set(row, xfs[styleIdx]);
    }
    if (col === 'D' && value) {
      const v = String(value).trim();
      if ((v.startsWith('-') || v.startsWith('_')) && !fLabel.has(row)) fLabel.set(row, v);
    }
  }

  const rowBands = {};
  const mismatches = [];
  for (let row = 2; row <= MAX_ROW; row++) {
    const label = fLabel.get(row) || '';
    const bg = fStyle.get(row);
    let band = bandFromBg(bg);
    if (!band) band = bandFromLabel(label, row);
    if (band) rowBands[String(row)] = band;

    const webFromLabel = bandFromLabel(label, row);
    if (webFromLabel && band && webFromLabel !== band) {
      mismatches.push({ row, label, bg, band, webFromLabel });
    }
  }

  console.log('rowBands count:', Object.keys(rowBands).length);
  console.log('label/bg mismatches:', mismatches.length);
  for (const x of mismatches.slice(0, 20)) console.log(JSON.stringify(x));

  const out = path.join(ROOT, 'tools', 'synthesis-row-bands.json');
  fs.writeFileSync(out, JSON.stringify({ rowBands, mismatches: mismatches.length }));
  console.log('Written', out);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
