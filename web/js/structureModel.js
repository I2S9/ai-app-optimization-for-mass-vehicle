/**
 * Extract / apply workbook section structure for the matrix editor.
 */
import {
  buildCellMap,
  displayValue,
  getCell,
  getSubSectionLabel,
  isSubSectionRow,
  isSectionRow,
  isCaBandRow,
  formatBlueBandLabel,
} from './bdStore.js';
import {
  synLabel,
  isSynL1SectionLabel,
  isSynL2SubsectionLabel,
  findSynAdaptationRow,
  SYN_HEADER_PANEL_LAST_ROW,
  SYN_SKIPPED_ROWS,
} from './synStore.js';

const DEFAULT_SECTION_COLOR = '#fff2cc';
const DEFAULT_SUBSECTION_COLOR = '#bdd7ee';

function uid(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

function rowRange(start, end) {
  const rows = [];
  for (let r = start; r <= end; r++) rows.push(r);
  return rows;
}

function getBdSectionLabel(map, row, sheet) {
  if (isCaBandRow(map, row)) {
    if (row === 5) return '-ADAPTATION';
    if (row === 139) return '-ADTH';
  }
  const ap = displayValue(getCell(map, row, 'AP'));
  if (ap) return String(ap).trim();
  const a = displayValue(getCell(map, row, 'A'));
  if (a && String(a).trim().startsWith('-')) return String(a).trim();
  return `Section ${row}`;
}

/** Ordered L1 section header rows in the BD sheet. */
function bdSectionHeaderRows(sheet, map) {
  const fromSheet = sheet.sectionHeaderRows;
  const set = fromSheet instanceof Set ? fromSheet : new Set(fromSheet || []);
  const rows = [...set].sort((a, b) => a - b);
  if (!rows.length) {
    for (let r = sheet.dataStartRow || 6; r <= sheet.lastRow; r++) {
      const ap = displayValue(getCell(map, r, 'AP'));
      if (ap && ap.trim()) rows.push(r);
    }
    rows.sort((a, b) => a - b);
  }
  return [...new Set(rows)];
}

export function extractBdStructure(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const sh = sheet.sectionHeaderRows;
  const headers = bdSectionHeaderRows(sheet, map);
  const colors = sheet.matrixColors || {};
  const sections = [];

  for (let i = 0; i < headers.length; i++) {
    const headerRow = headers[i];
    const endRow =
      i + 1 < headers.length ? headers[i + 1] - 1 : sheet.lastRow;
    const subStarts = [];
    for (let r = headerRow + 1; r <= endRow; r++) {
      if (isSubSectionRow(map, r, sh)) subStarts.push(r);
    }
    const subsections = [];
    for (let j = 0; j < subStarts.length; j++) {
      const startRow = subStarts[j];
      const subEnd =
        j + 1 < subStarts.length ? subStarts[j + 1] - 1 : endRow;
      const label = getSubSectionLabel(map, startRow) || `_Row ${startRow}`;
      subsections.push({
        id: uid('sub'),
        label,
        startRow,
        endRow: subEnd,
        color: colors[startRow] || DEFAULT_SUBSECTION_COLOR,
      });
    }
    sections.push({
      id: uid('sec'),
      label: getBdSectionLabel(map, headerRow, sheet),
      headerRow,
      endRow,
      color: colors[headerRow] || DEFAULT_SECTION_COLOR,
      subsections,
      customLines: [],
    });
  }
  return {
    prefixEndRow: (sheet.dataStartRow || 6) - 1,
    finRow: sheet.finRow ?? sheet.lastRow,
    sections,
  };
}

export function extractSynStructure(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const adapt = findSynAdaptationRow(map, sheet);
  const prefixEndRow = SYN_HEADER_PANEL_LAST_ROW;
  const colors = sheet.matrixColors || {};
  const last = sheet.effectiveLastRow ?? sheet.lastRow ?? 530;
  const l1Rows = [];
  for (let r = adapt; r <= last; r++) {
    const label = synLabel(map, r);
    if (isSynL1SectionLabel(label)) l1Rows.push(r);
  }
  const sections = [];
  for (let i = 0; i < l1Rows.length; i++) {
    const headerRow = l1Rows[i];
    const endRow = i + 1 < l1Rows.length ? l1Rows[i + 1] - 1 : last;
    const subStarts = [];
    for (let r = headerRow + 1; r <= endRow; r++) {
      const label = synLabel(map, r);
      if (isSynL2SubsectionLabel(label)) subStarts.push(r);
    }
    const subsections = [];
    for (let j = 0; j < subStarts.length; j++) {
      const startRow = subStarts[j];
      const subEnd =
        j + 1 < subStarts.length ? subStarts[j + 1] - 1 : endRow;
      subsections.push({
        id: uid('sub'),
        label: synLabel(map, startRow) || `_Row ${startRow}`,
        startRow,
        endRow: subEnd,
        color: colors[startRow] || DEFAULT_SUBSECTION_COLOR,
        bdLabel: null,
      });
    }
    sections.push({
      id: uid('sec'),
      label: synLabel(map, headerRow) || `Section ${headerRow}`,
      headerRow,
      endRow,
      color: colors[headerRow] || DEFAULT_SECTION_COLOR,
      subsections,
      customLines: [],
    });
  }
  return { prefixEndRow, finRow: last, sections, skippedRows: [...SYN_SKIPPED_ROWS] };
}

/** Deep clone for the matrix editor UI. */
export function cloneStructure(model) {
  return JSON.parse(JSON.stringify(model));
}

export function findSection(model, id) {
  return model.sections.find((s) => s.id === id);
}

export function findSubsection(model, subId) {
  for (const sec of model.sections) {
    const sub = sec.subsections.find((s) => s.id === subId);
    if (sub) return { section: sec, subsection: sub };
  }
  return null;
}

/** Preserve data rows (STLA/S, project lines…) in place when reordering structure. */
function bdBlocksFromStructure(model) {
  const blocks = [];
  for (const sec of model.sections) {
    blocks.push({
      type: 'section',
      rows: [sec.headerRow],
      meta: sec,
    });
    const subs = [...sec.subsections].sort((a, b) => a.startRow - b.startRow);
    let cursor = sec.headerRow + 1;
    const end = sec.endRow ?? sec.headerRow;
    for (const sub of subs) {
      if (sub.isNew) continue;
      while (cursor < sub.startRow) {
        blocks.push({ type: 'passthrough', rows: [cursor++] });
      }
      const rows = rowRange(sub.startRow, sub.endRow);
      blocks.push({
        type: 'subsection',
        rows,
        meta: { section: sec, subsection: sub },
      });
      cursor = sub.endRow + 1;
    }
    while (cursor <= end) {
      blocks.push({ type: 'passthrough', rows: [cursor++] });
    }
  }
  return blocks;
}

function remapSheetRows(sheet, oldToNew, newLastRow) {
  const cells = [];
  for (const c of sheet.cells || []) {
    const nr = oldToNew.get(c.r);
    if (nr == null) continue;
    cells.push({ ...c, r: nr });
  }
  const headerRows = {};
  for (const [rowStr, cols] of Object.entries(sheet.headerRows || {})) {
    const nr = oldToNew.get(parseInt(rowStr, 10));
    if (nr == null) continue;
    headerRows[String(nr)] = cols;
  }
  const merges = (sheet.merges || [])
    .map((m) => {
      const sr = oldToNew.get(m.startRow);
      const er = oldToNew.get(m.endRow);
      if (sr == null || er == null) return null;
      return { ...m, startRow: sr, endRow: er };
    })
    .filter(Boolean);
  const matrixColors = {};
  for (const [oldR, color] of Object.entries(sheet.matrixColors || {})) {
    const nr = oldToNew.get(parseInt(oldR, 10));
    if (nr != null) matrixColors[nr] = color;
  }
  return {
    ...sheet,
    cells,
    headerRows,
    merges,
    matrixColors,
    lastRow: newLastRow,
    finRow: newLastRow,
  };
}

function isCaChapterRow(row) {
  return row === 5 || row === 139;
}

function insertNewBdRows(sheet, model) {
  for (const sec of model.sections) {
    for (const sub of sec.subsections) {
      if (!sub.isNew || sub.startRow == null) continue;
      const lbl = formatBlueBandLabel(sub.label);
      setCell(sheet, sub.startRow, 'A', lbl);
      setCell(sheet, sub.startRow, 'AS', lbl);
      setCell(sheet, sub.startRow, 'AP', sec.label);
    }
    for (const line of sec.customLines || []) {
      if (!line.isNew || line.startRow == null) continue;
      setCell(sheet, line.startRow, 'A', line.label);
      setCell(sheet, line.startRow, 'AP', sec.label);
    }
  }
}

function applyBdLabels(sheet, model) {
  for (const sec of model.sections) {
    const hr = sec.headerRow;
    if (isCaChapterRow(hr)) {
      setCell(sheet, hr, 'A', sec.label);
    } else {
      setCell(sheet, hr, 'AP', sec.label);
    }
    for (const sub of sec.subsections) {
      const lbl = formatBlueBandLabel(sub.label);
      setCell(sheet, sub.startRow, 'A', lbl);
      setCell(sheet, sub.startRow, 'AS', lbl);
    }
    for (const line of sec.customLines || []) {
      if (line.isNew) continue;
      setCell(sheet, line.startRow, 'A', line.label);
    }
  }
  return sheet;
}

function setCell(sheet, row, col, value) {
  const idx = sheet.cells.findIndex((c) => c.r === row && c.c === col);
  if (idx >= 0) {
    sheet.cells[idx] = { ...sheet.cells[idx], v: value };
    delete sheet.cells[idx].f;
  } else {
    sheet.cells.push({ r: row, c: col, v: value });
  }
}

function buildBdRowMap(raw, model) {
  const prefixEnd = model.prefixEndRow ?? 5;
  const oldToNew = new Map();
  for (let r = 1; r <= prefixEnd; r++) oldToNew.set(r, r);
  let next = prefixEnd + 1;
  const blocks = bdBlocksFromStructure(model);
  for (const block of blocks) {
    for (const oldR of block.rows) {
      oldToNew.set(oldR, next++);
    }
    if (block.type === 'section') block.meta.headerRow = oldToNew.get(block.rows[0]);
    if (block.type === 'subsection') {
      block.meta.subsection.startRow = oldToNew.get(block.rows[0]);
      block.meta.subsection.endRow = oldToNew.get(block.rows[block.rows.length - 1]);
    }
  }
  for (const sec of model.sections) {
    const lastSub = sec.subsections[sec.subsections.length - 1];
    sec.endRow = lastSub?.endRow ?? sec.headerRow;
  }
  let insertAt = next;
  for (const sec of model.sections) {
    for (const sub of sec.subsections) {
      if (!sub.isNew) continue;
      sub.startRow = insertAt;
      sub.endRow = insertAt;
      oldToNew.set(`new:${sub.id}`, insertAt);
      insertAt++;
    }
  }
  return { oldToNew, newLastRow: insertAt - 1 };
}

function syncSynToBdLabels(synModel, bdModel) {
  const byLabel = new Map();
  for (const sec of bdModel.sections) {
    byLabel.set(normalizeKey(sec.label), sec);
  }
  for (const sec of synModel.sections) {
    const bd = byLabel.get(normalizeKey(sec.label));
    if (!bd) continue;
    const bdSubs = bd.subsections.map((s) => normalizeKey(s.label));
    sec.subsections.forEach((sub, i) => {
      if (bdSubs[i]) sub.bdLabel = bd.subsections[i]?.label;
    });
  }
}

function normalizeKey(label) {
  return String(label || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

function buildSynBlocks(synModel, bdModel) {
  syncSynToBdLabels(synModel, bdModel);
  const blocks = [];
  const prefixEnd = synModel.prefixEndRow ?? 22;
  blocks.push({ type: 'prefix', rows: rowRange(1, prefixEnd) });
  const bdOrder = [];
  for (const sec of bdModel.sections) {
    for (const sub of sec.subsections) {
      bdOrder.push({ secLabel: sec.label, subLabel: sub.label });
    }
  }
  const synBySub = new Map();
  for (const sec of synModel.sections) {
    for (const sub of sec.subsections) {
      synBySub.set(normalizeKey(sub.label), { sec, sub });
    }
  }
  for (const sec of synModel.sections) {
    blocks.push({ type: 'section', rows: [sec.headerRow], meta: sec });
  }
  for (const { secLabel, subLabel } of bdOrder) {
    const hit = synBySub.get(normalizeKey(subLabel));
    if (!hit) continue;
    const { sub } = hit;
    if (sub._placed) continue;
    sub._placed = true;
    blocks.push({
      type: 'subsection',
      rows: rowRange(sub.startRow, sub.endRow),
      meta: sub,
    });
  }
  for (const sec of synModel.sections) {
    for (const sub of sec.subsections) {
      if (sub._placed) continue;
      blocks.push({
        type: 'subsection',
        rows: rowRange(sub.startRow, sub.endRow),
        meta: sub,
      });
    }
  }
  return blocks;
}

function applySynRowOrder(synRaw, synModel, bdModel) {
  const blocks = buildSynBlocks(synModel, bdModel);
  const oldToNew = new Map();
  let next = 1;
  for (const block of blocks) {
    for (const oldR of block.rows) {
      oldToNew.set(oldR, next++);
    }
    if (block.type === 'section') block.meta.headerRow = oldToNew.get(block.rows[0]);
    if (block.type === 'subsection') {
      block.meta.startRow = oldToNew.get(block.rows[0]);
      block.meta.endRow = oldToNew.get(block.rows[block.rows.length - 1]);
    }
  }
  const skipped = new Set(synModel.skippedRows || []);
  let gap = 0;
  for (const skip of [...skipped].sort((a, b) => a - b)) {
    for (const [oldR, newR] of [...oldToNew.entries()]) {
      if (oldR > skip) oldToNew.set(oldR, newR + 1);
    }
    gap++;
  }
  const sheet = remapSheetRows(synRaw, oldToNew, next - 1 + gap);
  applySynLabels(sheet, synModel);
  return sheet;
}

function applySynLabels(sheet, model) {
  for (const sec of model.sections) {
    setCell(sheet, sec.headerRow, 'F', sec.label);
    for (const sub of sec.subsections) {
      const lbl = formatBlueBandLabel(sub.label);
      setCell(sheet, sub.startRow, 'F', lbl);
    }
  }
  return sheet;
}

function collectMatrixColors(model) {
  const colors = {};
  for (const sec of model.sections) {
    colors[sec.headerRow] = sec.color;
    for (const sub of sec.subsections) {
      colors[sub.startRow] = sub.color;
    }
    for (const line of sec.customLines || []) {
      if (!line.isNew) colors[line.startRow] = line.color;
    }
  }
  return colors;
}

function linkBdStructureToSyn(bd, syn) {
  const synByLabel = new Map();
  for (const s of syn.sections) synByLabel.set(normalizeKey(s.label), s);
  for (const sec of bd.sections) {
    const synSec = synByLabel.get(normalizeKey(sec.label));
    if (!synSec) continue;
    sec.synLinkId = synSec.id;
    const synSubByLabel = new Map();
    for (const sub of synSec.subsections) {
      synSubByLabel.set(normalizeKey(sub.label), sub);
    }
    for (const sub of sec.subsections) {
      const synSub = synSubByLabel.get(normalizeKey(sub.label));
      if (synSub) sub.synLinkId = synSub.id;
    }
  }
}

export function alignSynModelToBd(synModel, bdModel) {
  const synBySec = new Map();
  const synSecById = new Map();
  for (const sec of synModel.sections) {
    synBySec.set(normalizeKey(sec.label), sec);
    synSecById.set(sec.id, sec);
  }
  const sections = [];
  for (const bdSec of bdModel.sections) {
    let synSec =
      (bdSec.synLinkId && synSecById.get(bdSec.synLinkId)) ||
      synBySec.get(normalizeKey(bdSec.label));
    const sec = synSec
      ? { ...synSec, label: bdSec.label, color: bdSec.color, subsections: [] }
      : {
          id: uid('sec'),
          label: bdSec.label,
          headerRow: null,
          endRow: null,
          color: bdSec.color,
          subsections: [],
          customLines: [],
        };
    const synSubs = new Map();
    const synSubById = new Map();
    for (const sub of synSec?.subsections || []) {
      synSubs.set(normalizeKey(sub.label), sub);
      synSubById.set(sub.id, sub);
    }
    for (const bdSub of bdSec.subsections) {
      const existing =
        (bdSub.synLinkId && synSubById.get(bdSub.synLinkId)) ||
        synSubs.get(normalizeKey(bdSub.label));
      sec.subsections.push(
        existing
          ? {
              ...existing,
              label: bdSub.label,
              color: bdSub.color,
              synLinkId: existing.id,
            }
          : {
              id: uid('sub'),
              label: bdSub.label,
              color: bdSub.color,
              isNew: true,
              startRow: null,
              endRow: null,
            }
      );
    }
    sections.push(sec);
  }
  return { ...synModel, sections };
}

export function applyStructureToBdRaw(bdRaw, model) {
  const m = cloneStructure(model);
  const { oldToNew, newLastRow } = buildBdRowMap(bdRaw, m);
  let sheet = remapSheetRows(bdRaw, oldToNew, newLastRow);
  insertNewBdRows(sheet, m);
  sheet.matrixColors = collectMatrixColors(m);
  sheet = applyBdLabels(sheet, m);
  return { sheet, model: m };
}

export function applyStructureToSynRaw(synRaw, synModel, bdModel) {
  const m = alignSynModelToBd(synModel, bdModel);
  const sheet = applySynRowOrder(synRaw, m, bdModel);
  sheet.matrixColors = collectMatrixColors(m);
  for (const sec of m.sections) {
    for (const sub of sec.subsections) {
      if (sub.isNew && sub.startRow != null) {
        setCell(sheet, sub.startRow, 'F', formatBlueBandLabel(sub.label));
      }
    }
  }
  applySynLabels(sheet, m);
  return { sheet, model: m };
}

export function applyMatrixSave(bdRaw, synRaw, bdModel, synModel) {
  const bd = applyStructureToBdRaw(bdRaw, bdModel);
  const syn = synModel
    ? applyStructureToSynRaw(synRaw, synModel, bd.model)
    : { sheet: synRaw, model: null };
  return {
    bdRaw: bd.sheet,
    synRaw: syn.sheet,
    bdModel: bd.model,
    synModel: syn.model,
  };
}

export function buildMatrixState(bdSheet, synSheet) {
  const bd = extractBdStructure(bdSheet);
  const syn = synSheet ? extractSynStructure(synSheet) : null;
  if (syn) linkBdStructureToSyn(bd, syn);
  return { bd: cloneStructure(bd), syn: syn ? cloneStructure(syn) : null };
}
