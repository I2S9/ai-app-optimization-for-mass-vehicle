/**
 * Export SYNTHESIS sheet → web/public/data/synthesis-sheet.json
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
const MAX_COL = 'NI';
const LABEL_COL = 'F';

/** Excel indexed palette (common defaults). */
const INDEXED_COLORS = {
  0: '#000000',
  1: '#FFFFFF',
  2: '#FF0000',
  3: '#00FF00',
  4: '#0000FF',
  5: '#FFFF00',
  6: '#FF00FF',
  7: '#00FFFF',
  8: '#000000',
  9: '#FFFFFF',
  10: '#FF0000',
  11: '#00FF00',
  12: '#0000FF',
  13: '#FFFF00',
  14: '#FF00FF',
  15: '#00FFFF',
  16: '#800000',
  17: '#008000',
  18: '#000080',
  19: '#808000',
  20: '#800080',
  21: '#008080',
  22: '#C0C0C0',
  23: '#808080',
  40: '#FFFFCC',
  41: '#CCFFFF',
  42: '#FFCCCC',
  43: '#CCCCFF',
  44: '#CCFFCC',
  45: '#FFCCFF',
  46: '#99CCFF',
  47: '#FF99CC',
  48: '#CC99FF',
  49: '#FFCC99',
  50: '#3366FF',
  51: '#33CCCC',
  52: '#99CC00',
  53: '#FFCC00',
  54: '#FF9900',
  55: '#FF6600',
  56: '#666699',
  57: '#969696',
  58: '#003366',
  59: '#339966',
  60: '#003300',
  61: '#333300',
  62: '#993300',
  63: '#993366',
  64: '#333399',
  65: '#333333',
};

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

function resolveColor(node) {
  if (!node) return null;
  const rgb = node.match(/rgb="([^"]+)"/)?.[1];
  if (rgb) {
    const hex = rgb.length === 8 ? rgb.slice(2) : rgb;
    return `#${hex}`.toLowerCase();
  }
  const indexed = node.match(/indexed="(\d+)"/)?.[1];
  if (indexed != null) return INDEXED_COLORS[indexed] ?? null;
  const theme = node.match(/theme="(\d+)"/)?.[1];
  if (theme === '0') return '#FFFFFF';
  if (theme === '1') return '#000000';
  return null;
}

function parseStyles(stylesXml) {
  const fills = [{ fg: null }];
  const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
  let fm;
  while ((fm = fillRe.exec(stylesXml)) !== null) {
    const block = fm[1];
    const fg =
      resolveColor(block.match(/<fgColor[^>]*\/>/)?.[0] || block.match(/<fgColor[^>]*>/)?.[0]) ||
      resolveColor(block.match(/<fgColor[^>]*[^/]>[\s\S]*?<\/fgColor>/)?.[0]);
    const bg =
      resolveColor(block.match(/<bgColor[^>]*\/>/)?.[0] || block.match(/<bgColor[^>]*>/)?.[0]);
    fills.push({ fg: fg || bg });
  }
  const fonts = [{}];
  const fontRe = /<font>([\s\S]*?)<\/font>/g;
  let fontm;
  while ((fontm = fontRe.exec(stylesXml)) !== null) {
    const block = fontm[1];
    fonts.push({
      bold: /<b\s*\/>/.test(block) || /<b>/.test(block),
      color: resolveColor(block.match(/<color[^>]*\/>/)?.[0] || block.match(/<color[^>]*>/)?.[0]),
    });
  }
  const xfs = [];
  const xfRe = /<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g;
  let xfm;
  while ((xfm = xfRe.exec(stylesXml)) !== null) {
    const attrs = xfm[1];
    const fillId = parseInt(attrs.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
    const fontId = parseInt(attrs.match(/fontId="(\d+)"/)?.[1] ?? '0', 10);
    const applyFill =
      fillId > 0 ||
      attrs.includes('applyFill="1"') ||
      attrs.includes('applyFill="true"');
    const f = fills[fillId] || {};
    const font = fonts[fontId] || {};
    xfs.push({
      backgroundColor: applyFill && f.fg ? f.fg : null,
      color: font.color,
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

/** Resolve shared-string index even when OOXML omits t="s". */
function resolveCellValue(raw, t, shared) {
  if (raw == null || raw === '') return raw;
  if (t === 's') {
    const idx = parseInt(String(raw), 10);
    return shared[idx] ?? String(raw);
  }
  const s = String(raw).trim();
  if (/^\d+$/.test(s)) {
    const hit = shared[parseInt(s, 10)];
    if (
      hit != null &&
      hit !== '' &&
      !/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(String(hit).trim())
    ) {
      return hit;
    }
  }
  return s;
}

function parseMergeRef(ref) {
  const [a, b] = ref.split(':');
  const parse = (addr) => {
    const col = addr.replace(/\d+$/, '');
    const row = parseInt(addr.replace(/^[A-Z]+/, ''), 10);
    return { col, row, colNum: colToNum(col) };
  };
  const start = parse(a);
  const end = parse(b || a);
  return {
    ref,
    startRow: start.row,
    endRow: end.row,
    startCol: start.col,
    endCol: end.col,
    rowspan: end.row - start.row + 1,
    colspan: end.colNum - start.colNum + 1,
  };
}

function parseRowHeights(sheetXml) {
  const rowHeights = {};
  const rowRe = /<row r="(\d+)"([^>]*)>/g;
  let m;
  while ((m = rowRe.exec(sheetXml)) !== null) {
    const ht = m[2].match(/\bht="([^"]+)"/)?.[1];
    if (ht) rowHeights[m[1]] = Math.round(+ht * 1.33);
  }
  return rowHeights;
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

  const merges = [];
  const mergeRe = /<mergeCell ref="([^"]+)"/g;
  let mm;
  while ((mm = mergeRe.exec(sheetXml)) !== null) {
    merges.push(parseMergeRef(mm[1]));
  }

  const rowHeights = parseRowHeights(sheetXml);
  const usedCols = new Set();
  const headers = {};
  const headerRows = {};
  const cells = [];
  const labelMeta = new Map();

  const cellRe = /<c r="([A-Z]+\d+)"([^/]*?)(\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = cellRe.exec(sheetXml)) !== null) {
    const ref = m[1];
    const attrs = m[2];
    const inner = m[3] === '/>' ? '' : (m[4] || '');
    const col = ref.replace(/\d+$/, '');
    const row = parseInt(ref.replace(/^[A-Z]+/, ''), 10);
    if (row > MAX_ROW || colToNum(col) > colToNum(MAX_COL)) continue;

    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    const styleIdx = parseInt(attrs.match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    let formula = null;
    let value = null;
    const fM = inner.match(/<f[^>]*>([\s\S]*?)<\/f>/);
    if (fM) formula = decodeXml(fM[1]);
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = resolveCellValue(vM[1], t, shared);
    const is = inner.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
    if (is) value = is[1];

    const display = value ?? (formula ? null : '');
    const style = styleIdx >= 0 ? styles[styleIdx] : null;
    usedCols.add(col);

    if (row === 1 && display) headers[col] = display;

    if (row >= 2 && row <= 14) {
      if (!headerRows[row]) headerRows[row] = {};
      if (display != null || formula) {
        headerRows[row][col] = {
          v: display ?? '',
          f: formula,
          s: style,
          bg: style?.backgroundColor,
        };
      }
    }

    if (row >= 2) {
      const entry = { r: row, c: col };
      if (display != null && display !== '') entry.v = String(display);
      // SUMPRODUCT strings are huge — web recalculates from BD when filters change.
      if (formula && !/SUMPRODUCT/i.test(formula)) entry.f = formula;
      else if (formula && (display == null || display === '')) entry.f = 'SUMPRODUCT';
      if (style?.backgroundColor) entry.bg = style.backgroundColor;
      if (style?.color) entry.fc = style.color;
      if (style?.fontWeight === 'bold') entry.b = 1;
      if (entry.v || entry.f || entry.bg || entry.fc || entry.b) cells.push(entry);
      if (col === LABEL_COL) {
        const meta = labelMeta.get(row) || {};
        if (display != null && display !== '') meta.label = String(display).trim();
        if (style?.backgroundColor) meta.bg = style.backgroundColor;
        labelMeta.set(row, meta);
      }
    }
  }

  function classifyLabelBand(row, meta = {}) {
    if (row === 4) return 'separator';
    if (row >= 3 && row <= 14) return 'filter';
    const bg = meta.bg?.toLowerCase();
    if (
      bg === '#ccffcc' ||
      bg === '#ffffcc' ||
      bg === '#ffff00' ||
      bg === '#ff99cc'
    ) {
      return 'section';
    }
    const label = meta.label || '';
    if (label.startsWith('_')) return 'subsection';
    if (label.startsWith('-')) return 'section';
    return null;
  }

  const rowBands = {};
  for (let row = 3; row <= 14; row++) {
    rowBands[String(row)] = row === 4 ? 'separator' : 'filter';
  }
  for (let row = 15; row <= MAX_ROW; row++) {
    const band = classifyLabelBand(row, labelMeta.get(row));
    if (band) rowBands[String(row)] = band;
  }

  const columns = [];
  for (let i = 1; i <= colToNum(MAX_COL); i++) columns.push(numToCol(i));

  const widthByCol = new Map(colWidths.map((w) => [w.col, w.width]));
  const fullColWidths = columns.map((col) => ({
    col,
    width: widthByCol.get(col) || (col === 'F' ? 180 : col <= 'G' ? 72 : 56),
  }));

  const FILTER_LABELS = {
    3: 'Date',
    4: 'Projet',
    5: 'Silhouette',
    6: 'Hybridation',
    7: 'Plaque de conception',
    8: 'Sièges',
    9: 'Spécificité technique',
    10: '',
    11: 'Pôle',
    12: 'Energie',
    13: 'Pack technique',
    14: 'Finition',
  };
  for (const [rowStr, label] of Object.entries(FILTER_LABELS)) {
    const row = parseInt(rowStr, 10);
    const hasF = cells.some((c) => c.r === row && c.c === LABEL_COL && c.v);
    if (!hasF && label) {
      cells.push({ r: row, c: LABEL_COL, v: label });
      if (!headerRows[row]) headerRows[row] = {};
      headerRows[row][LABEL_COL] = { v: label };
    }
  }
  const METRIC_LABELS = {
    16: 'Curb mass :',
    17: 'PM pre-target',
    18: 'Curb mass : last update',
    19: 'Control',
    20: 'Portfolio',
    21: 'Forecast',
    22: '% Forecast/target',
  };
  for (const [rowStr, label] of Object.entries(METRIC_LABELS)) {
    const row = parseInt(rowStr, 10);
    const hasF = cells.some((c) => c.r === row && c.c === LABEL_COL && c.v);
    if (!hasF) {
      cells.push({ r: row, c: LABEL_COL, v: label });
      if (!headerRows[row]) headerRows[row] = {};
      headerRows[row][LABEL_COL] = { v: label };
    }
  }

  let lastUsedRow = 14;
  for (const cell of cells) {
    if (!cell.v && !cell.f) continue;
    if (cell.r > lastUsedRow) lastUsedRow = cell.r;
  }
  for (const [rowStr, cols] of Object.entries(headerRows)) {
    const row = parseInt(rowStr, 10);
    if (
      row > lastUsedRow &&
      Object.values(cols).some((c) => c.v || c.f)
    ) {
      lastUsedRow = row;
    }
  }
  for (const rowStr of Object.keys(rowBands)) {
    lastUsedRow = Math.max(lastUsedRow, parseInt(rowStr, 10));
  }

  const payload = {
    version: 2,
    sheet: 'SYNTHESIS',
    dataStartRow: 15,
    filterRows: [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
    lastRow: lastUsedRow,
    columns,
    colWidths: fullColWidths,
    headers,
    headerRows,
    merges,
    rowHeights,
    rowBands,
    cells: cells.filter((c) => c.r <= MAX_ROW),
  };

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(payload));
  console.log('Written:', OUT);
  console.log(
    'Cells:',
    payload.cells.length,
    'Columns:',
    columns.length,
    'Merges:',
    merges.length,
    'Last row:',
    MAX_ROW,
    'Row bands:',
    Object.keys(rowBands).length,
    'Size MB:',
    (fs.statSync(OUT).size / 1e6).toFixed(2)
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
