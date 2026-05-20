/**
 * Analyze BD STLA/S rows, project header row 5, and Excel fills.
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const TMP = path.join(process.env.TEMP || 'xlsm-bd-export');
const WB = path.join(ROOT, 'workbooks', 'base-de-donnees-complete-avec-liens.xlsm');

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
  for (const m of xml.matchAll(/<si>([\s\S]*?)<\/si>/g)) {
    const parts = [];
    for (const tm of m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)) parts.push(tm[1]);
    strings.push(parts.join(''));
  }
  return strings;
}

function parseStyles(stylesXml) {
  const indexed = ['#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF'];
  const ic = stylesXml.match(/<indexedColors>([\s\S]*?)<\/indexedColors>/);
  if (ic) {
    indexed.length = 0;
    for (const m of ic[1].matchAll(/rgbColor rgb="([^"]+)"/g)) {
      indexed.push('#' + m[1]);
    }
  }
  const themeColors = [];
  const theme = stylesXml.match(/<clrScheme[^>]*>([\s\S]*?)<\/clrScheme>/);
  if (theme) {
    for (const m of theme[1].matchAll(/<sysClr[^>]*lastClr="([^"]+)"/g)) {
      themeColors.push('#' + m[1].replace(/^FF/i, ''));
    }
    for (const m of theme[1].matchAll(/<srgbClr val="([^"]+)"/g)) {
      themeColors.push('#' + m[1]);
    }
  }

  const fills = [{ fg: null }];
  for (const m of stylesXml.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) {
    const block = m[1];
    let fg = null;
    const fgM = block.match(/<fgColor([^/]*)\/?>/);
    if (fgM) {
      const attrs = fgM[1];
      const rgb = attrs.match(/rgb="([^"]+)"/)?.[1];
      const idx = attrs.match(/indexed="(\d+)"/)?.[1];
      const th = attrs.match(/theme="(\d+)"/)?.[1];
      const tint = parseFloat(attrs.match(/tint="([^"]+)"/)?.[1] ?? '0');
      if (rgb) fg = '#' + rgb;
      else if (idx != null && indexed[+idx]) fg = indexed[+idx];
      else if (th != null && themeColors[+th]) {
        fg = themeColors[+th];
        if (tint) {
          /* skip tint math for audit */
        }
      }
    }
    fills.push({ fg });
  }

  const xfs = [];
  for (const m of stylesXml.matchAll(/<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g)) {
    const attrs = m[1];
    const fillId = parseInt(attrs.match(/fillId="(\d+)"/)?.[1] ?? '0', 10);
    const applyFill =
      attrs.includes('applyFill="1"') || attrs.includes('applyFill="true"');
    const f = fills[fillId] || {};
    xfs.push({
      backgroundColor: applyFill && f.fg ? f.fg : null,
    });
  }
  return { xfs, fills, indexed, themeColors };
}

function parseSheet(sheetXml, shared, xfs) {
  const cells = new Map();
  for (const m of sheetXml.matchAll(/<c r="([A-Z]+\d+)"([^>]*)>([\s\S]*?)<\/c>/g)) {
    const ref = m[1];
    const attrs = m[2];
    const inner = m[3];
    const t = attrs.match(/\bt="([^"]+)"/)?.[1];
    const styleIdx = parseInt(attrs.match(/\bs="(\d+)"/)?.[1] ?? '-1', 10);
    let value = null;
    const vM = inner.match(/<v>([^<]*)<\/v>/);
    if (vM) value = t === 's' ? shared[parseInt(vM[1], 10)] ?? vM[1] : vM[1];
    const is = inner.match(/<is>[\s\S]*?<t[^>]*>([^<]*)<\/t>/);
    if (is) value = is[1];
    const style = styleIdx >= 0 ? xfs[styleIdx] : null;
    cells.set(ref, { value, bg: style?.backgroundColor });
  }
  return cells;
}

async function main() {
  if (!fs.existsSync(path.join(TMP, 'xl/worksheets/sheet3.xml'))) {
    const { execSync } = await import('child_process');
    fs.mkdirSync(TMP, { recursive: true });
    const zip = path.join(TMP, 'base.zip');
    fs.copyFileSync(WB, zip);
    execSync(
      `powershell -NoProfile -Command "Expand-Archive -LiteralPath '${zip.replace(/'/g, "''")}' -DestinationPath '${TMP.replace(/'/g, "''")}' -Force"`,
      { stdio: 'inherit' }
    );
  }

  const shared = parseSharedStrings(
    fs.readFileSync(path.join(TMP, 'xl/sharedStrings.xml'), 'utf8')
  );
  const { xfs } = parseStyles(
    fs.readFileSync(path.join(TMP, 'xl/styles.xml'), 'utf8')
  );
  const bd = parseSheet(
    fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet3.xml'), 'utf8'),
    shared,
    xfs
  );

  const projectCols = [];
  for (let i = colToNum('B'); i <= colToNum('P'); i++) projectCols.push(numToCol(i));

  console.log('=== Row 5 project columns (Excel fills) ===');
  for (const col of projectCols) {
    const c = bd.get(`${col}5`);
    if (c) console.log(col, c.value, c.bg || '(no rgb fill)');
  }

  console.log('\n=== Row 10 STLA/S project columns ===');
  for (const col of projectCols) {
    const c = bd.get(`${col}10`);
    if (c) console.log(col, c.value, c.bg || '(no rgb fill)');
  }

  let stla = 0;
  let stlaEmptyA = 0;
  for (let r = 6; r <= 3480; r++) {
    const a = bd.get(`A${r}`)?.value;
    if (a === 'STLA/S' || a === 'STLA-S') {
      stla++;
      if (!a) stlaEmptyA++;
    }
  }
  console.log('\nSTLA rows in Excel:', stla);

  let noBg = 0;
  let withBg = 0;
  for (const col of projectCols) {
    const c = bd.get(`${col}5`);
    if (c?.bg) withBg++;
    else noBg++;
  }
  console.log('Row5 project cols with resolved bg:', withBg, 'without:', noBg);
}

main().catch(console.error);
