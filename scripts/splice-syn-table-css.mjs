import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const css = fs.readFileSync(path.join(__dirname, '..', 'web', 'css', '_syn-new-tables.css.txt'), 'utf8');
const summaryIdx = css.indexOf('/* DR');
if (summaryIdx < 0) throw new Error('summary block not found');
const panel = css.slice(0, summaryIdx);
const summary = css.slice(summaryIdx);

const gridPath = path.join(__dirname, '..', 'web', 'css', 'bd-grid.css');
let grid = fs.readFileSync(gridPath, 'utf8');

const panelAnchor = '/* AC';
const panelPos = grid.indexOf(panelAnchor);
if (panelPos < 0) throw new Error('panel anchor not found');
grid = grid.slice(0, panelPos) + panel + grid.slice(panelPos);

const spot = '/* Spot blue — display F/G & I/J on listed rows (overrides ADAPTATION band) */';
const spotPos = grid.indexOf(spot);
if (spotPos < 0) throw new Error('spot anchor not found');
grid = grid.slice(0, spotPos) + summary + spot + grid.slice(spotPos + spot.length);

fs.writeFileSync(gridPath, grid, 'utf8');
console.log('spliced panel', panel.length, 'summary', summary.length);
