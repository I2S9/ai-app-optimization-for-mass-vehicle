/**
 * Extract / apply workbook section structure for the matrix editor.
 */
import {
  buildCellMap,
  displayValue,
  getCell,
  isSectionRow,
  isCaBandRow,
  formatBlueBandLabel,
  computeSectionHeaderRows,
  computeOutlineRows,
  bdSubsystemL1Col,
  bdSubsystemL2Col,
  isBdL2BookmarkStart,
  getBdL2BookmarkLabel,
} from './bdStore.js';
import {
  synLabel,
  isSynL1SectionLabel,
  isSynL2SubsectionLabel,
  findSynAdaptationRow,
  SYN_HEADER_PANEL_LAST_ROW,
  SYN_SKIPPED_ROWS,
  SYN_MAX_EXCEL_ROW,
} from './synStore.js';
import { translateValue } from './bdTranslate.js';

/** Same as bd-grid.css row-section / row-subsection */
const DEFAULT_SECTION_COLOR = '#ffff00';
const DEFAULT_SUBSECTION_COLOR = '#00b0f0';

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
  const l1Col = bdSubsystemL1Col(sheet);
  const ap = displayValue(getCell(map, row, l1Col));
  if (ap) return translateValue(String(ap).trim());
  const a = displayValue(getCell(map, row, 'A'));
  if (a && String(a).trim().startsWith('-')) {
    return translateValue(String(a).trim());
  }
  return `Section ${row}`;
}

/** Ordered L1 section header rows in the BD sheet. */
function bdSectionHeaderRows(sheet) {
  const fromSheet = sheet.sectionHeaderRows;
  const set = fromSheet instanceof Set ? fromSheet : new Set(fromSheet || []);
  if (set.size > 2) return [...set].sort((a, b) => a - b);
  const { rows } = computeSectionHeaderRows(sheet);
  return [...rows].sort((a, b) => a - b);
}

export function extractBdStructure(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const sh =
    sheet.sectionHeaderRows ||
    computeSectionHeaderRows(sheet).rows;
  const l2Col = bdSubsystemL2Col(sheet);
  const headers = bdSectionHeaderRows(sheet);
  const colors = sheet.matrixColors || {};
  const sections = [];

  for (let i = 0; i < headers.length; i++) {
    const headerRow = headers[i];
    const endRow =
      i + 1 < headers.length ? headers[i + 1] - 1 : sheet.lastRow;
    const subStarts = [];
    for (let r = headerRow + 1; r <= endRow; r++) {
      if (isBdL2BookmarkStart(map, r, sh, l2Col)) subStarts.push(r);
    }
    const subsections = [];
    for (let j = 0; j < subStarts.length; j++) {
      const startRow = subStarts[j];
      const subEnd =
        j + 1 < subStarts.length ? subStarts[j + 1] - 1 : endRow;
      const label =
        getBdL2BookmarkLabel(map, startRow, sheet, sh, l2Col) ||
        `_Row ${startRow}`;
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
  const last = sheet.effectiveLastRow ?? sheet.lastRow ?? SYN_MAX_EXCEL_ROW;
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

/**
 * Move each L1 section as one contiguous band (header → endRow), in matrix column order.
 * Avoids dropping rows between sub-sections or between sections when reordering.
 */
function bdBlocksFromStructure(model) {
  const blocks = [];
  for (const sec of model.sections) {
    const start = sec.headerRow;
    if (start == null) continue;
    const end = sec.endRow ?? start;
    if (end < start) continue;
    blocks.push({
      type: 'section-band',
      rows: rowRange(start, end),
      meta: sec,
    });
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
  const l1Col = bdSubsystemL1Col(sheet);
  const l2Col = bdSubsystemL2Col(sheet);
  for (const sec of model.sections) {
    for (const sub of sec.subsections) {
      if (!sub.isNew || sub.startRow == null) continue;
      const lbl = formatBlueBandLabel(sub.label);
      setCell(sheet, sub.startRow, 'A', lbl);
      setCell(sheet, sub.startRow, l2Col, lbl);
      setCell(sheet, sub.startRow, l1Col, sec.label);
    }
    for (const line of sec.customLines || []) {
      if (!line.isNew || line.startRow == null) continue;
      setCell(sheet, line.startRow, 'A', line.label);
      setCell(sheet, line.startRow, l1Col, sec.label);
    }
  }
}

function applyBdLabels(sheet, model) {
  const l1Col = bdSubsystemL1Col(sheet);
  const l2Col = bdSubsystemL2Col(sheet);
  for (const sec of model.sections) {
    const hr = sec.headerRow;
    if (isCaChapterRow(hr)) {
      setCell(sheet, hr, 'A', sec.label);
    } else {
      setCell(sheet, hr, l1Col, sec.label);
    }
    for (const sub of sec.subsections) {
      const lbl = formatBlueBandLabel(sub.label);
      setCell(sheet, sub.startRow, 'A', lbl);
      setCell(sheet, sub.startRow, l2Col, lbl);
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
    sheet.cells[idx] = {
      ...sheet.cells[idx],
      v: value,
      userEdited: true,
    };
    delete sheet.cells[idx].f;
  } else {
    sheet.cells.push({ r: row, c: col, v: value, userEdited: true });
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
    if (block.type === 'section-band') {
      const sec = block.meta;
      sec.headerRow = oldToNew.get(block.rows[0]);
      sec.endRow = oldToNew.get(block.rows[block.rows.length - 1]);
      for (const sub of sec.subsections) {
        if (sub.isNew) continue;
        const nrStart = oldToNew.get(sub.startRow);
        const nrEnd = oldToNew.get(sub.endRow);
        if (nrStart != null) sub.startRow = nrStart;
        if (nrEnd != null) sub.endRow = nrEnd;
      }
    }
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
    colors[sec.headerRow] = sec.color || DEFAULT_SECTION_COLOR;
    for (const sub of sec.subsections) {
      colors[sub.startRow] = sub.color || DEFAULT_SUBSECTION_COLOR;
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

function attachBdStructureMeta(sheet, model) {
  const rows = new Set([5, 139]);
  for (const sec of model?.sections || []) {
    if (sec.headerRow != null) rows.add(sec.headerRow);
  }
  const { rows: computedRows, canonicalByLabel } = computeSectionHeaderRows(sheet);
  for (const r of computedRows) rows.add(r);
  sheet.sectionHeaderRows = [...rows].sort((a, b) => a - b);
  sheet.canonicalSectionByLabel = Object.fromEntries(canonicalByLabel);
  sheet.outlineRows = computeOutlineRows(sheet).filter(
    (r) => r <= sheet.lastRow
  );
  return sheet;
}

export function applyStructureToBdRaw(bdRaw, model) {
  const m = cloneStructure(model);
  const { oldToNew, newLastRow } = buildBdRowMap(bdRaw, m);
  let sheet = remapSheetRows(bdRaw, oldToNew, newLastRow);
  insertNewBdRows(sheet, m);
  sheet.matrixColors = collectMatrixColors(m);
  sheet = applyBdLabels(sheet, m);
  sheet = attachBdStructureMeta(sheet, m);
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
