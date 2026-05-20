import fs from 'fs';
import path from 'path';
const xml = fs.readFileSync(path.join(process.env.TEMP, 'xlsm-syn-colors/xl/styles.xml'), 'utf8');
const hits = [];
for (const m of xml.matchAll(/rgb="([^"]+)"/g)) {
  const rgb = m[1].toLowerCase();
  if (rgb.includes('b0f0') || rgb.includes('ffff00') || rgb.includes('ccffcc')) hits.push(rgb);
}
console.log('rgb hits', [...new Set(hits)].slice(0, 30));
const fills = [];
let i = 0;
for (const m of xml.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) {
  const b = m[1];
  if (/b0f0|ffff00|ccffcc|00B0/i.test(b)) fills.push({ i, snippet: b.slice(0, 120) });
  i++;
}
console.log('fill hits', fills.slice(0, 15));
