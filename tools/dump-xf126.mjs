import fs from 'fs';
import path from 'path';
const xml = fs.readFileSync(path.join(process.env.TEMP, 'xlsm-syn-colors/xl/styles.xml'), 'utf8');
const fonts = [];
for (const m of xml.matchAll(/<font>([\s\S]*?)<\/font>/g)) fonts.push(m[1]);
let i = 0;
for (const m of xml.matchAll(/<xf([^/]*)(?:\/>|>([\s\S]*?)<\/xf>)/g)) {
  if (i++ !== 126) continue;
  const fontId = parseInt(m[1].match(/fontId="(\d+)"/)?.[1] ?? '0', 10);
  console.log('xf126', m[1].slice(0, 200));
  console.log('font', fonts[fontId]);
  break;
}
