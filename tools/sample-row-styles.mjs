import fs from 'fs';
import path from 'path';
const TMP = path.join(process.env.TEMP, 'xlsm-bd-export');
const sheet = fs.readFileSync(path.join(TMP, 'xl/worksheets/sheet3.xml'), 'utf8');
const styles = fs.readFileSync(path.join(TMP, 'xl/styles.xml'), 'utf8');
const fills = [];
for (const m of styles.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) {
  const b = m[1];
  fills.push({
    fg: b.match(/<fgColor[^>]*rgb="([^"]+)"/)?.[1],
    fgi: b.match(/<fgColor[^>]*indexed="(\d+)"/)?.[1],
    theme: b.match(/<fgColor[^>]*theme="(\d+)"/)?.[1],
  });
}
const xfs = [];
for (const m of styles.matchAll(/<xf([^/>]*)\/?>/g)) {
  const a = m[1];
  xfs.push({
    fillId: a.match(/fillId="(\d+)"/)?.[1],
    applyFill: /applyFill="1"|applyFill="true"/.test(a),
  });
}
function styleFor(ref) {
  const cm = sheet.match(new RegExp(`<c r="${ref}"([^>]*)>`));
  if (!cm) return null;
  const si = cm[1].match(/s="(\d+)"/)?.[1];
  if (si == null) return null;
  const xf = xfs[+si];
  const fill = fills[+xf?.fillId];
  return { si, fillId: xf?.fillId, fill, applyFill: xf?.applyFill };
}
for (const ref of ['A6', 'AP6', 'AS6', 'A7', 'AS7', 'W7', 'A8', 'B10', 'V10']) {
  console.log(ref, JSON.stringify(styleFor(ref)));
}
