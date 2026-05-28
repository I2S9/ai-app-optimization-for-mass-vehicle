const tables = [['DR', 'ED'], ['EF', 'EQ'], ['ES', 'FE'], ['FJ', 'FZ']];

function slug(a, b) {
  return `${a.toLowerCase()}-${b.toLowerCase()}`;
}
function edge(a, b) {
  return `${a.toLowerCase()}${b.toLowerCase()}`;
}

function panelDividerBlock(start, end, tag) {
  const states = [
    `tr.syn-proj-table-frame td.data-cell.syn-hdr-edge-${tag}-right:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right:not(.syn-pillar-col)`,
    `tr.syn-header-block.syn-filter-band td.data-cell.syn-hdr-edge-${tag}-right:not(.syn-pillar-col)`,
    `tr.syn-header-block.syn-metric-band td.data-cell.syn-hdr-edge-${tag}-right:not(.syn-pillar-col)`,
    `tr.syn-header-block.syn-header-spacer-white td.data-cell.syn-hdr-edge-${tag}-right:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-filter-col-grey:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-hdr-row-metric-bg:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-val-avenger-like:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-val-p1x:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-hdr-val-p1h:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-hdr-val-hev:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-hdr-val-mhevp2:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-hdr-val-default:not(.syn-pillar-col)`,
    `tr.syn-header-block td.data-cell.syn-hdr-edge-${tag}-right.syn-metric-cj-white:not(.syn-pillar-col)`,
  ];
  const panel = states.map((s) => `.synthesis-grid .bd-table tbody ${s}`).join(',\n');
  const gridStates = states.map((s) =>
    s.replace('tr.syn-proj-table-frame td', 'tr.syn-header-block.syn-hdr-panel-grid td')
  );
  const grid = gridStates.map((s) => `.synthesis-grid .bd-table tbody ${s}`).join(',\n');
  return `/* Bold vertical dividers between display columns ${start}–${end} (rows 3–22) */
${panel} {
  border-right: var(--syn-bold-border) solid #000 !important;
}
/* Bold vertical dividers ${start}–${end} — every panel row 3–22 */
${grid} {
  border-right: var(--syn-bold-border) solid #000 !important;
}
`;
}

function summaryBlock(start, end) {
  const s = slug(start, end);
  const t = edge(start, end);
  const ls = start.toLowerCase();
  const rs = end.toLowerCase();
  return `/* ${start}–${end} summary table — display columns ${start}…${end}, rows 3–22 (display rows 3–23) */
.synthesis-grid .bd-table tbody tr.syn-${s}-table-frame td.data-cell.syn-${s}-cell:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-block.syn-${s}-table-frame td.data-cell.syn-${s}-cell:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-filter-band.syn-${s}-table-frame td.data-cell.syn-${s}-cell:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-metric-band.syn-${s}-table-frame td.data-cell.syn-${s}-cell:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-spacer-white.syn-${s}-table-frame td.data-cell.syn-${s}-cell:not(.syn-pillar-col) {
  border-bottom: var(--syn-bold-border) solid #000 !important;
}
.synthesis-grid .bd-table tbody tr.syn-${s}-edge-top td.data-cell.syn-${s}-cell:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-block.syn-${s}-edge-top td.data-cell.syn-${s}-cell:not(.syn-pillar-col) {
  border-top: var(--syn-bold-border) solid #000 !important;
}
.synthesis-grid .bd-table tbody tr.syn-${s}-edge-bottom td.data-cell.syn-${s}-cell:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-block.syn-${s}-edge-bottom td.data-cell.syn-${s}-cell:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-panel-end.syn-${s}-edge-bottom td.data-cell.syn-${s}-cell:not(.syn-pillar-col) {
  border-bottom: var(--syn-bold-border) solid #000 !important;
}
.synthesis-grid .bd-table tbody tr.syn-${s}-table-frame td.data-cell.syn-hdr-edge-${t}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-block td.data-cell.syn-hdr-edge-${t}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-filter-band td.data-cell.syn-hdr-edge-${t}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-metric-band td.data-cell.syn-hdr-edge-${t}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-spacer-white td.data-cell.syn-hdr-edge-${t}-right:not(.syn-pillar-col) {
  border-right: var(--syn-bold-border) solid #000 !important;
}
.synthesis-grid .bd-table tbody tr.syn-${s}-table-frame td.data-cell.syn-hdr-edge-${ls}-left:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-block td.data-cell.syn-hdr-edge-${ls}-left:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-filter-band td.data-cell.syn-hdr-edge-${ls}-left:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-metric-band td.data-cell.syn-hdr-edge-${ls}-left:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-spacer-white td.data-cell.syn-hdr-edge-${ls}-left:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody td.data-cell.syn-hdr-edge-${ls}-left:not(.syn-pillar-col) {
  border-left: var(--syn-bold-border) solid #000 !important;
  box-shadow: inset 2px 0 0 #000 !important;
}
.synthesis-grid .bd-table tbody tr.syn-${s}-table-frame td.data-cell.syn-hdr-edge-${rs}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-block td.data-cell.syn-hdr-edge-${rs}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-filter-band td.data-cell.syn-hdr-edge-${rs}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-metric-band td.data-cell.syn-hdr-edge-${rs}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody tr.syn-header-spacer-white td.data-cell.syn-hdr-edge-${rs}-right:not(.syn-pillar-col),
.synthesis-grid .bd-table tbody td.data-cell.syn-hdr-edge-${rs}-right:not(.syn-pillar-col) {
  border-right: var(--syn-bold-border) solid #000 !important;
  box-shadow: inset -2px 0 0 #000 !important;
}
`;
}

let out = '';
for (const [a, b] of tables) out += panelDividerBlock(a, b, edge(a, b)) + '\n';
for (const [a, b] of tables) out += summaryBlock(a, b);

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const __dirname = path.dirname(fileURLToPath(import.meta.url));
fs.writeFileSync(path.join(__dirname, '..', 'web', 'css', '_syn-new-tables.css.txt'), out, 'utf8');
