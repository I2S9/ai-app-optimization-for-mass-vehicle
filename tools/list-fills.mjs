import fs from 'fs';
import path from 'path';
const styles = fs.readFileSync(
  path.join(process.env.TEMP, 'xlsm-bd-export', 'xl', 'styles.xml'),
  'utf8'
);
let i = 0;
for (const m of styles.matchAll(/<fill>([\s\S]*?)<\/fill>/g)) {
  const b = m[1];
  if (b.includes('indexed') || b.includes('rgb')) {
    console.log(i, b.replace(/\s+/g, ' ').slice(0, 120));
  }
  i++;
}
