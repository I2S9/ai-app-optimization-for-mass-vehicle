/**
 * ONE-TIME export: SYNTHESIS sheet → web/public/data/synthesis-sheet.json
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const WB = path.join(ROOT, 'workbooks', 'base-de-donnees-complete-avec-liens.xlsm');
const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-bd-export');
const OUT = path.join(ROOT, 'web', 'public', 'data', 'synthesis-sheet.json');
const MAX_ROW = 530;

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

function parseStyles(stylesXml) {
  const fills = [{ fg: null, bg: null }];
  const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
  let fm;
  while ((fm = fillRe.exec(stylesXml)) !== null) {
    const block = fm[1];
    const fg =
      block.match(/<fgColor[^>]*rgb="([^"]+)"/)?.[1] ||
      block.match(/<fgColor[^>]*indexed="(\d+)"/)?.[1];
    fills.push({
      fg: fg ? (fg.length === 6 ? '#' + fg : fg) : null,
    });
  }
  const fonts = [{}];
  const fontRe = /<font>([\s\S]*?)<\/font>/g;
  let fontm;
  while ((fontm = fontRe.exec(stylesXml)) !== null) {
    const block = fontm[1];
    fonts.push({
      bold: /<b\s*\/>/.test(block) || /<b>/.test(block),
      color: block.match(/<color[^>]*rgb="([^"]+)"/)?.[1],
    });
  }
  const xfs = [];
  const xfRe = /<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g;
  let xfm;
  while ((xfm = xfRe.exec(stylesXml)) !== null) {
    const attrs = xfm[1];
    const fillId = parseInt(attrs.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
    const fontId = parseInt(attrs.match(/fontId="(\d+)"/)?.[1] ?? '0', 10);
    const applyFill = attrs.includes('applyFill="1"') || attrs.includes('applyFill="true"');
    const f = fills[fillId] || {};
    const font = fonts[fontId] || {};
    xfs.push({
      backgroundColor: applyFill && f.fg ? (f.fg.startsWith('#') ? f.fg : '#' + f.fg) : null,
      color: font.color ? '#' + font.color : null,
      fontWeight: font.bold ? 'bold' : 'normal',
    });
  }
  return xfs;
}

function decodeXml(s) {
  return s
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"');
}

async function main() {
  if (!fs.existsSync(WB)) {
    console.error('Workbook not found:', WB);
    process.exit(1);
  }
  const { execSync } = await import('child_process');
  if (!fs.existsSync(path.join(TMP, 'xl', 'worksheets', 'sheet4.xml'))) {
    fs.mkdirSync(TMP, { recursive: true });
    const zip = path.join(TMP, 'base.zip');
    fs.copyFileSync(WB, zip);
    execSync(
      `powershell -NoProfile -Command "Expand-Archive -LiteralPath '${zip.replace(/'/g, "''")}' -DestinationPath '${TMP.replace(/'/g, "''")}' -Force"`,
      { stdio: 'inherit' }
    );
  }

  const shared = parseSharedStrings(
    fs.readFileSync(path.join(TMP, 'xl', 'sharedStrings.xml'), 'utf8')
  );
  const styles = parseStyles(fs.readFileSync(path.join(TMP, 'xl', 'styles.xml'), 'utf8'));
  const sheetXml = fs.readFileSync(path.join(TMP, 'xl', 'worksheets', 'sheet4.xml'), 'utf8');

  const colWidths = [];
  const colsM = sheetXml.match(/<cols>([\s\S]*?)<\/cols>/);
  if (colsM) {
    const re = /<col min="(\d+)" max="(\d+)"[^>]*width="([^"]+)"/g;
    let cm;
    while ((cm = re.exec(colsM[1])) !== null) {
      for (let i = +cm[1]; i <= +cm[2]; i++) {
        colWidths.push({ col: numToCol(i), width: Math.round(+cm[3] * 7) });
      }
    }
  }

  const usedCols = new Set(['A', 'B', 'C', 'D', 'E', 'F', 'G']);
  const headers = {};
  const headerRows = {};
  const cells = [];

  const cellRe = /<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g;
  let m;
  while ((m = cellRe.exec(sheetXml)) !== null) {
    const ref = m[1];
    const attrs = m[2];
    const inner = m[3];
    const col = ref.replace(/\d+$/, '');
    const row = parseInt(ref.replace(/^[A-Z]+/, ''), 10);
    if (row > MAX_ROW) continue;

    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    const styleIdx = parseInt(attrs.match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    let formula = null;
    let value = null;
    const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
    if (fM) formula = decodeXml(fM[1]);
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = t === 's' ? shared[parseInt(vM[1], 10)] ?? vM[1] : vM[1];
    const is = inner.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
    if (is) value = is[1];

    const display = value ?? (formula ? null : '');
    const style = styleIdx >= 0 ? styles[styleIdx] : null;
    usedCols.add(col);

    if (row === 1 && display) headers[col] = display;

    if (row <= 14) {
      if (!headerRows[row]) headerRows[row] = {};
      if (display != null || formula)
        headerRows[row][col] = { v: display ?? '', f: formula, s: style };
    }

    if (row >= 3) {
      const entry = { r: row, c: col };
      if (display != null && display !== '') entry.v = String(display);
      if (formula) entry.f = formula;
      if (style?.backgroundColor) entry.bg = style.backgroundColor;
      if (style?.color) entry.fc = style.color;
      if (style?.fontWeight === 'bold') entry.b = 1;
      cells.push(entry);
    }
  }

  const columns = [...usedCols].sort((a, b) => colToNum(a) - colToNum(b));
  const lastRow = Math.min(
    MAX_ROW,
    cells.reduce((max, c) => Math.max(max, c.r), 3)
  );

  const payload = {
    version: 1,
    sheet: 'SYNTHESIS',
    dataStartRow: 3,
    lastRow,
    columns,
    colWidths: colWidths.filter((w) => columns.includes(w.col)),
    headers,
    headerRows,
    cells: cells.filter((c) => c.r <= lastRow),
  };

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(payload));
  console.log('Written:', OUT);
  console.log(
    'Cells:',
    payload.cells.length,
    'Columns:',
    columns.length,
    'Last row:',
    lastRow,
    'Size MB:',
    (fs.statSync(OUT).size / 1e6).toFixed(2)
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
