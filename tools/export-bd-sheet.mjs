/**
 * ONE-TIME export: BD sheet → web/public/data/bd-sheet.json
 * Not used at runtime by the app.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const WB = path.join(ROOT, 'workbooks', 'base-de-donnees-complete-avec-liens.xlsm');
const TMP = path.join(process.env.TEMP || '/tmp', 'xlsm-bd-export');
const OUT = path.join(ROOT, 'web', 'public', 'data', 'bd-sheet.json');
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
/** Default Excel indexed palette (OOXML spec). */
const EXCEL_INDEXED = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
  '#000000', '#FFFFFF',
];

function resolveColor(attrs, themeColors) {
  if (!attrs) return null;
  const rgb = attrs.match(/rgb="([^"]+)"/)?.[1];
  if (rgb) return '#' + rgb.replace(/^FF/i, '').slice(-6);
  const idx = attrs.match(/indexed="(\d+)"/)?.[1];
  if (idx != null && EXCEL_INDEXED[+idx]) return EXCEL_INDEXED[+idx];
  const theme = attrs.match(/theme="(\d+)"/)?.[1];
  if (theme != null && themeColors[+theme]) return themeColors[+theme];
  return null;
}

function parseThemeColors(stylesXml) {
  const themeColors = [];
  const scheme = stylesXml.match(/<clrScheme[^>]*>([\s\S]*?)<\/clrScheme>/);
  if (!scheme) return themeColors;
  for (const m of scheme[1].matchAll(/<a:srgbClr val="([^"]+)"/g)) {
    themeColors.push('#' + m[1]);
  }
  for (const m of scheme[1].matchAll(/<srgbClr val="([^"]+)"/g)) {
    themeColors.push('#' + m[1]);
  }
  return themeColors;
}

function parseStyles(stylesXml) {
  const themeColors = parseThemeColors(stylesXml);
  const fills = [{ fg: null, bg: null }];
  const fillRe = /<fill>([\s\S]*?)<\/fill>/g;
  let fm;
  while ((fm = fillRe.exec(stylesXml)) !== null) {
    const block = fm[1];
    const fgAttrs = block.match(/<fgColor([^/]*)\/?>/)?.[1];
    const bgAttrs = block.match(/<bgColor([^/]*)\/?>/)?.[1];
    fills.push({
      fg: resolveColor(fgAttrs, themeColors),
      bg: resolveColor(bgAttrs, themeColors),
    });
  }
  const fonts = [{}];
  const fontRe = /<font>([\s\S]*?)<\/font>/g;
  let fontm;
  while ((fontm = fontRe.exec(stylesXml)) !== null) {
    const block = fontm[1];
    fonts.push({
      bold: /<b\s*\/>/.test(block) || /<b>/.test(block),
      italic: /<i\s*\/>/.test(block),
      color: block.match(/<color[^>]*rgb="([^"]+)"/)?.[1],
      size: block.match(/<sz val="([^"]+)"/)?.[1],
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
      fontStyle: font.italic ? 'italic' : 'normal',
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

function parseFormula(inner, sharedFormulas) {
  const open = inner.match(/<f([^>]*)>([\s\S]*?)<\/f>/);
  if (open) {
    const attrs = open[1];
    const body = decodeXml(open[2]);
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
async function main() {
  if (!fs.existsSync(WB)) {
    console.error('Workbook not found:', WB);
    process.exit(1);
  }
  const { execSync } = await import('child_process');
  if (!fs.existsSync(path.join(TMP, 'xl', 'worksheets', 'sheet3.xml'))) {
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
  const sheetXml = fs.readFileSync(path.join(TMP, 'xl', 'worksheets', 'sheet3.xml'), 'utf8');
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
  const maxCol = colToNum('AV');
  const columns = [];
  for (let i = 1; i <= maxCol; i++) columns.push(numToCol(i));
  const headers = {};
  const headerRows = {};
  let cells = [];
  const sharedFormulas = new Map();
  /** [^/]*? attrs so self-closing <c r="AC7" s="63"/> is not merged with the next cell. */
  const cellRe = /<c r="([A-Z]+\d+)"([^/]*?)(\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = cellRe.exec(sheetXml)) !== null) {
    const ref = m[1];
    const attrs = m[2];
    const inner = m[3] === '/>' ? '' : (m[4] || '');
    const col = ref.replace(/\d+$/, '');
    const row = parseInt(ref.replace(/^[A-Z]+/, ''), 10);
    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    const styleIdx = parseInt(attrs.match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    let formula = parseFormula(inner, sharedFormulas);
    let value = null;
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = resolveCellValue(vM[1], t, shared);
    const is = inner.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
    if (is) value = is[1];
    const display = value ?? (formula ? null : '');
    const style = styleIdx >= 0 ? styles[styleIdx] : null;
    if (row === 1 && display) headers[col] = display;
    if (row <= 5) {
      if (!headerRows[row]) headerRows[row] = {};
      if (display != null || formula)
        headerRows[row][col] = { v: display ?? '', f: formula, s: style };
    }
    if (row >= 6) {
      const entry = { r: row, c: col };
      if (display != null && display !== '') entry.v = String(display);
      if (formula) entry.f = formula;
      if (style?.backgroundColor) entry.bg = style.backgroundColor;
      if (style?.color) entry.fc = style.color;
      if (style?.fontWeight === 'bold') entry.b = 1;
      if (entry.v || entry.f || entry.bg || entry.fc || entry.b) cells.push(entry);
    }
  }
  const EN_HEADERS = {
    Date: 'Date',
    Projet: 'Project',
    Silhouette: 'Silhouette',
    'Plaque de conception': 'Design plate',
    Hybridation: 'Hybridization',
    Sièges: 'Seats',
    'Spécificité technique': 'Technical spec',
    'Charge utile': 'Payload',
    'Pack technique': 'Technical pack',
    Roues: 'Wheels',
    Pôle: 'Pole',
    Energie: 'Energy',
    Moteur: 'Engine',
    Boite: 'Gearbox',
    Finition: 'Trim',
    Equipts: 'Equipment',
    Option: 'Option',
    Codification: 'Codification',
    Découpage: 'Breakdown',
    Intitulé: 'Title',
    BVH: 'BVH',
    Masse: 'Mass',
    AV: 'Front (AV)',
    AR: 'Rear (AR)',
    'Désignation technique': 'Technical designation',
    Source: 'Source',
    Pack: 'Pack',
    Reference: 'Reference',
    Metier: 'Trade',
    'Positionnement en X': 'X position',
    'Positionnement en Y': 'Y position',
    'Positionnement en Z': 'Z position',
    'Champ libre': 'Free field',
    'CODE MODULE ': 'Module code',
    'LOT DECPSA': 'Lot DECPSA',
    'Type modulaire': 'Modular type',
    'Attribut technique': 'Technical attribute',
    'Sub-System Level1': 'Sub-system L1',
    'Ligne N°': 'Line #',
    'Sub-System Level2': 'Sub-system L2',
    'Sub-System Design Dpt': 'Design dept',
  };
  const headerEn = {};
  for (const [col, label] of Object.entries(headers)) {
    headerEn[col] = EN_HEADERS[label] || label;
  }
  let lastRow = cells.reduce((max, c) => Math.max(max, c.r), 6);
  const finCell = cells.find((c) => c.c === 'A' && String(c.v || '').trim().toUpperCase() === 'FIN');
  if (finCell) {
    lastRow = finCell.r;
    cells = cells.filter((c) => c.r <= lastRow);
  }
  const payload = {
    version: 1,
    sheet: 'BD',
    dataStartRow: 6,
    lastRow,
    columns,
    colWidths: colWidths,
    headers: headerEn,
    headersRaw: headers,
    headerRows,
    cells,
  };
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(payload));
  console.log('Written:', OUT);
  console.log('Cells:', cells.length, 'Last row:', lastRow, 'Size MB:', (fs.statSync(OUT).size / 1e6).toFixed(2));
}
main().catch((e) => {
  console.error(e);
  process.exit(1);
});
