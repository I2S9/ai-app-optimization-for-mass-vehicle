import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const url = (p) => `file://${join(root, p).replace(/\\/g, '/')}`;
const { synRawLooksHealthy } = await import(url('web/js/sessionPersistence.js'));
const good = JSON.parse(
  readFileSync(join(root, 'web/public/data/synthesis-sheet.json'), 'utf8')
);
const bad = JSON.parse(JSON.stringify(good));
bad.cells = bad.cells.filter((c) => !(c.r === 25 && c.c === 'F'));

if (!synRawLooksHealthy(good)) {
  console.error('FAIL: project file should be healthy');
  process.exit(1);
}
if (synRawLooksHealthy(bad)) {
  console.error('FAIL: corrupted syn should be unhealthy');
  process.exit(1);
}
console.log('ok: syn heal detection works');
