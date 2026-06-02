/**
 * Extract / apply workbook section structure for the matrix editor.
 */
import {
  buildCellMap,
  displayValue,
  getCell,
  getBdSectionTitleFromRow,
  sanitizeStructureLabel,
  isFormulaLike,
  isSectionLabel,
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
} from './synStore.js?v=syn-cicy1';
import { translateSubsystemLabel, translateValue } from './bdTranslate.js';

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
  const title = getBdSectionTitleFromRow(map, row, sheet);
  return title || `Section ${row}`;
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
      // CA chapter bands (-ADAPTATION / -ADTH) store their label in column W,
      // not the L1 column. Flag them so the label is written/read consistently.
      caBand: isCaBandRow(map, headerRow),
    });
  }
  // The first section header (row 5 = -ADAPTATION) must NOT be swallowed by the
  // prefix, otherwise it collides with prefixEnd and shifts on apply.
  const firstHeader = headers.length ? headers[0] : sheet.dataStartRow || 6;
  const model = {
    prefixEndRow: Math.max(
      0,
      Math.min((sheet.dataStartRow || 6) - 1, firstHeader - 1)
    ),
    finRow: sheet.finRow != null ? sheet.finRow : sheet.lastRow,
    sections,
  };
  applyStructureArchiveToModel(model, sheet.structureArchive);
  return model;
}

export function extractSynStructure(sheet) {
  const map = buildCellMap(sheet.cells, sheet.headerRows);
  const adapt = findSynAdaptationRow(map, sheet);
  const prefixEndRow = SYN_HEADER_PANEL_LAST_ROW;
  const colors = sheet.matrixColors || {};
  const last =
    sheet.effectiveLastRow != null
      ? sheet.effectiveLastRow
      : sheet.lastRow != null
        ? sheet.lastRow
        : SYN_MAX_EXCEL_ROW;
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
        label: translateSubsystemLabel(synLabel(map, startRow) || `_Row ${startRow}`),
        startRow,
        endRow: subEnd,
        color: colors[startRow] || DEFAULT_SUBSECTION_COLOR,
        bdLabel: null,
      });
    }
    sections.push({
      id: uid('sec'),
      label: translateSubsystemLabel(synLabel(map, headerRow) || `Section ${headerRow}`),
      headerRow,
      endRow,
      color: colors[headerRow] || DEFAULT_SECTION_COLOR,
      subsections,
      customLines: [],
    });
  }
  const model = {
    prefixEndRow,
    finRow: last,
    sections,
    skippedRows: [...SYN_SKIPPED_ROWS],
  };
  applyStructureArchiveToModel(model, sheet.structureArchive);
  return model;
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
  const active = [];
  const archived = [];
  for (const sec of model.sections) {
    const start = sec.headerRow;
    if (start == null) continue;
    const end = sec.endRow != null ? sec.endRow : start;
    if (end < start) continue;
    const block = {
      type: 'section-band',
      rows: rowRange(start, end),
      meta: sec,
    };
    (sec.archived ? archived : active).push(block);
  }
  return [...active, ...archived];
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

/**
 * BD CA chapter rows (W column) only mirror their title onto synthesis L1 bands.
 * They must never drive synthesis row bands or sub-section lists (different layout).
 */
const CA_BAND_SYN_LABEL_MIRROR = [
  { bdHeaderRow: 5, synHeaderRow: 25 },
  { bdHeaderRow: 139, synHeaderRow: 41 },
];

function syncCaBandLabelsOntoSyn(synModel, bdModel) {
  if (!synModel || !bdModel) return;
  for (const { bdHeaderRow, synHeaderRow } of CA_BAND_SYN_LABEL_MIRROR) {
    const bdSec = bdModel.sections.find(
      (s) => s.caBand && s.headerRow === bdHeaderRow
    );
    const synSec = synModel.sections.find((s) => s.headerRow === synHeaderRow);
    if (bdSec && synSec) synSec.label = bdSec.label;
  }
}

function insertNewBdRows(sheet, model) {
  const l1Col = bdSubsystemL1Col(sheet);
  const l2Col = bdSubsystemL2Col(sheet);
  for (const sec of model.sections) {
    if (sec.isNew && sec.headerRow != null) {
      setCell(sheet, sec.headerRow, l1Col, sec.label);
      setCell(sheet, sec.headerRow, 'A', sec.label);
    }
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
  if (!sheet.canonicalSectionByLabel) sheet.canonicalSectionByLabel = {};
  for (const sec of model.sections) {
    const hr = sec.headerRow;
    const label = sanitizeStructureLabel(sec.label);
    if (!label) continue;
    sec.label = label;
    if (sec.caBand || isCaChapterRow(hr)) {
      // CA chapter label lives in column W (its canonical display cell).
      setCell(sheet, hr, 'W', label);
      sheet.canonicalSectionByLabel[label] = hr;
    } else {
      setCell(sheet, hr, l1Col, label);
      if (isSectionLabel(label)) {
        sheet.canonicalSectionByLabel[label] = hr;
      }
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
  const prefixEnd = model.prefixEndRow != null ? model.prefixEndRow : 5;
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
      // Re-map any sub-section that already owns a row. Driven by startRow (not
      // the isNew flag) so a stale isNew flag can never re-allocate a duplicate.
      for (const sub of sec.subsections) {
        if (sub.startRow == null) continue;
        const nrStart = oldToNew.get(sub.startRow);
        const nrEnd = oldToNew.get(sub.endRow);
        if (nrStart != null) sub.startRow = nrStart;
        if (nrEnd != null) sub.endRow = nrEnd;
      }
    }
  }
  // Allocate fresh rows ONLY for items that don't already own one. A section
  // with a headerRow is handled by its band above; a sub with a startRow was
  // re-mapped above — so nothing here can duplicate an existing row.
  let insertAt = next;
  for (const sec of model.sections) {
    const isNewSec = sec.headerRow == null;
    if (isNewSec) {
      sec.headerRow = insertAt;
      sec.endRow = insertAt;
      oldToNew.set(`newsec:${sec.id}`, insertAt);
      insertAt++;
    }
    for (const sub of sec.subsections) {
      if (sub.startRow != null) continue;
      sub.startRow = insertAt;
      sub.endRow = insertAt;
      oldToNew.set(`new:${sub.id}`, insertAt);
      insertAt++;
    }
    if (isNewSec) sec.endRow = insertAt - 1;
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
      if (bdSubs[i]) {
        sub.bdLabel = bd.subsections[i] ? bd.subsections[i].label : undefined;
      }
    });
  }
}

function normalizeKey(label) {
  return String(label || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

function sectionArchiveKey(label) {
  return normalizeKey(label);
}

function subsectionArchiveKey(sectionLabel, subLabel) {
  return `${normalizeKey(sectionLabel)}|${normalizeKey(subLabel)}`;
}

/** Persisted on raw sheets after matrix apply (label keys survive row remaps). */
export function structureArchiveFromModel(model) {
  const archive = { sections: {}, subsections: {} };
  if (!model || !model.sections) return archive;
  for (const sec of model.sections) {
    if (sec.archived) {
      archive.sections[sectionArchiveKey(sec.label)] = {
        restoreIndex: sec.archiveRestoreIndex != null ? sec.archiveRestoreIndex : 0,
      };
    }
    let subActive = 0;
    for (const sub of sec.subsections) {
      if (sub.archived) {
        archive.subsections[subsectionArchiveKey(sec.label, sub.label)] = {
          restoreIndex:
            sub.archiveRestoreIndex != null ? sub.archiveRestoreIndex : subActive,
        };
      } else {
        subActive++;
      }
    }
  }
  return archive;
}

export function applyStructureArchiveToModel(model, archive) {
  if (!model || !model.sections || !archive) return;
  for (const sec of model.sections) {
    const hit = archive.sections
      ? archive.sections[sectionArchiveKey(sec.label)]
      : undefined;
    if (hit) {
      sec.archived = true;
      sec.archiveRestoreIndex = hit.restoreIndex != null ? hit.restoreIndex : 0;
    } else {
      delete sec.archived;
      delete sec.archiveRestoreIndex;
    }
    let subActive = 0;
    for (const sub of sec.subsections) {
      const subHit =
        archive.subsections
          ? archive.subsections[subsectionArchiveKey(sec.label, sub.label)]
          : undefined;
      if (subHit) {
        sub.archived = true;
        sub.archiveRestoreIndex =
          subHit.restoreIndex != null ? subHit.restoreIndex : subActive;
      } else {
        delete sub.archived;
        delete sub.archiveRestoreIndex;
      }
      if (!sub.archived) subActive++;
    }
  }
  sortModelArchiveToEnd(model);
}

/** Active items first, archived at the end (sections and subsections). */
export function sortModelArchiveToEnd(model) {
  if (!model || !model.sections) return;
  const activeSecs = [];
  const archivedSecs = [];
  for (const sec of model.sections) {
    (sec.archived ? archivedSecs : activeSecs).push(sec);
  }
  for (const sec of [...activeSecs, ...archivedSecs]) {
    const activeSubs = [];
    const archivedSubs = [];
    for (const sub of sec.subsections) {
      (sub.archived ? archivedSubs : activeSubs).push(sub);
    }
    sec.subsections = [...activeSubs, ...archivedSubs];
  }
  model.sections = [...activeSecs, ...archivedSecs];
}

export function collectArchivedRowBands(model) {
  const bands = [];
  if (!model || !model.sections) return bands;
  for (const sec of model.sections) {
    if (sec.archived && sec.headerRow != null) {
      bands.push({
        start: sec.headerRow,
        end: sec.endRow != null ? sec.endRow : sec.headerRow,
      });
      continue;
    }
    for (const sub of sec.subsections) {
      if (sub.archived && sub.startRow != null) {
        bands.push({
          start: sub.startRow,
          end: sub.endRow != null ? sub.endRow : sub.startRow,
        });
      }
    }
  }
  return bands;
}

export function isExcelRowArchived(sheet, row) {
  const bands = sheet ? sheet.archivedRowBands : null;
  if (!bands || !bands.length || row == null) return false;
  for (const b of bands) {
    if (row >= b.start && row <= b.end) return true;
  }
  return false;
}

function activeSections(model) {
  return ((model && model.sections) || []).filter((s) => !s.archived);
}

export function archiveSection(model, secId) {
  const sections = model ? model.sections : null;
  if (!sections) return;
  const idx = sections.findIndex((s) => s.id === secId);
  if (idx < 0) return;
  const sec = sections[idx];
  if (sec.archived) return;
  sec.archiveRestoreIndex = activeSections(model).findIndex((s) => s.id === secId);
  sec.archived = true;
  sortModelArchiveToEnd(model);
}

export function restoreSection(model, secId) {
  const sections = model ? model.sections : null;
  if (!sections) return;
  const sec = sections.find((s) => s.id === secId);
  if (!sec || !sec.archived) return;
  const active = sections.filter((s) => !s.archived && s.id !== secId);
  let at = sec.archiveRestoreIndex != null ? sec.archiveRestoreIndex : active.length;
  at = Math.max(0, Math.min(at, active.length));
  sec.archived = false;
  delete sec.archiveRestoreIndex;
  const archived = sections.filter((s) => s.archived);
  model.sections = [
    ...active.slice(0, at),
    sec,
    ...active.slice(at),
    ...archived,
  ];
}

export function archiveSubsection(model, subId) {
  const hit = findSubsection(model, subId);
  if (!hit || hit.subsection.archived) return;
  const { section, subsection } = hit;
  const active = section.subsections.filter((s) => !s.archived);
  subsection.archiveRestoreIndex = active.findIndex((s) => s.id === subId);
  subsection.archived = true;
  sortModelArchiveToEnd(model);
}

export function restoreSubsection(model, subId) {
  const hit = findSubsection(model, subId);
  if (!hit || !hit.subsection || !hit.subsection.archived) return;
  const { section, subsection } = hit;
  const active = section.subsections.filter(
    (s) => !s.archived && s.id !== subId
  );
  let at =
    subsection.archiveRestoreIndex != null
      ? subsection.archiveRestoreIndex
      : active.length;
  at = Math.max(0, Math.min(at, active.length));
  subsection.archived = false;
  delete subsection.archiveRestoreIndex;
  const archived = section.subsections.filter((s) => s.archived);
  section.subsections = [
    ...active.slice(0, at),
    subsection,
    ...active.slice(at),
    ...archived,
  ];
}

/**
 * Synthesis row ordering mirrors the BD approach: each section is moved as one
 * contiguous band (header → endRow). This keeps EVERY body row of a section,
 * including the ~95% of sections that have no L2 sub-sections, so applying a
 * matrix change never drops synthesis data. Section order follows synModel
 * (already aligned to the BD/matrix order); archived sections move to the end.
 */
function buildSynBlocks(synModel, bdModel) {
  syncSynToBdLabels(synModel, bdModel);
  const blocks = [];
  const prefixEnd = synModel.prefixEndRow != null ? synModel.prefixEndRow : 22;
  blocks.push({ type: 'prefix', rows: rowRange(1, prefixEnd) });
  const active = [];
  const archived = [];
  const withRows = synModel.sections.filter((sec) => sec.headerRow != null);
  withRows.sort((a, b) => a.headerRow - b.headerRow);
  for (const sec of withRows) {
    const start = sec.headerRow;
    const end = sec.endRow != null && sec.endRow >= start ? sec.endRow : start;
    (sec.archived ? archived : active).push({
      type: 'section-band',
      rows: rowRange(start, end),
      meta: sec,
    });
  }
  return [...blocks, ...active, ...archived];
}

function applySynRowOrder(synRaw, synModel, bdModel) {
  const blocks = buildSynBlocks(synModel, bdModel);
  const oldToNew = new Map();
  let next = 1;
  for (const block of blocks) {
    for (const oldR of block.rows) oldToNew.set(oldR, next++);
  }
  const skipped = new Set(synModel.skippedRows || []);
  let gap = 0;
  for (const skip of [...skipped].sort((a, b) => a - b)) {
    for (const [oldR, newR] of [...oldToNew.entries()]) {
      if (oldR > skip) oldToNew.set(oldR, newR + 1);
    }
    gap++;
  }
  // Re-map each section / sub-section to its FINAL new row (after skip gaps).
  for (const sec of synModel.sections) {
    if (sec.headerRow != null && oldToNew.has(sec.headerRow)) {
      const newEnd = oldToNew.has(sec.endRow)
        ? oldToNew.get(sec.endRow)
        : oldToNew.get(sec.headerRow);
      sec.headerRow = oldToNew.get(sec.headerRow);
      sec.endRow = newEnd;
      for (const sub of sec.subsections) {
        if (sub.startRow != null && oldToNew.has(sub.startRow)) {
          sub.endRow = oldToNew.has(sub.endRow)
            ? oldToNew.get(sub.endRow)
            : oldToNew.get(sub.startRow);
          sub.startRow = oldToNew.get(sub.startRow);
        }
      }
    }
  }
  // Append rows ONLY for items the user just added in the matrix. Sections /
  // sub-sections that exist in BD but not in synthesis (headerRow/startRow null
  // without isNew) stay absent — synthesis is a summary, not a 1:1 mirror.
  let tail = next - 1 + gap;
  for (const sec of synModel.sections) {
    if (sec.headerRow == null) {
      if (!sec.isNew) continue;
      sec.headerRow = ++tail;
      sec.endRow = sec.headerRow;
    }
    for (const sub of sec.subsections) {
      if (sub.startRow != null || !sub.isNew) continue;
      sub.startRow = ++tail;
      sub.endRow = sub.startRow;
      if (sub.startRow > (sec.endRow || 0)) sec.endRow = sub.startRow;
    }
  }
  const sheet = remapSheetRows(synRaw, oldToNew, tail);
  applySynLabels(sheet, synModel);
  return sheet;
}

function applySynLabels(sheet, model) {
  for (const sec of model.sections) {
    const label = sanitizeStructureLabel(sec.label);
    if (!label || sec.headerRow == null) continue;
    sec.label = label;
    setCell(sheet, sec.headerRow, 'F', label);
    for (const sub of sec.subsections) {
      if (sub.startRow == null) continue;
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
    if (sec.caBand) continue;
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
  const consumed = new Set();
  for (const bdSec of bdModel.sections) {
    // CA bands (rows 5 / 139) are BD-only structure — never merge into synthesis.
    if (bdSec.caBand) continue;
    // Each synthesis section may be claimed by AT MOST ONE BD section, so a
    // rename/reorder updates the existing row band instead of duplicating it.
    let synSec = null;
    const byId = bdSec.synLinkId ? synSecById.get(bdSec.synLinkId) : null;
    if (byId && !consumed.has(byId.id)) {
      synSec = byId;
    } else {
      const byLabel = synBySec.get(normalizeKey(bdSec.label));
      if (byLabel && !consumed.has(byLabel.id)) synSec = byLabel;
    }
    if (synSec) consumed.add(synSec.id);
    const sec = synSec
      ? {
          ...synSec,
          label: bdSec.label,
          color: bdSec.color,
          archived: bdSec.archived,
          archiveRestoreIndex: bdSec.archiveRestoreIndex,
          subsections: [],
        }
      : {
          id: uid('sec'),
          label: bdSec.label,
          headerRow: null,
          endRow: null,
          color: bdSec.color,
          archived: bdSec.archived,
          archiveRestoreIndex: bdSec.archiveRestoreIndex,
          subsections: [],
          customLines: [],
          // Only sections the user just added in the matrix are materialised as
          // new synthesis rows; BD-only sections stay absent (synthesis is a
          // summary and must not gain hundreds of blank rows).
          isNew: Boolean(bdSec.isNew),
        };
    const synSubs = new Map();
    const synSubById = new Map();
    const synSecSubs = synSec && synSec.subsections ? synSec.subsections : [];
    for (const sub of synSecSubs) {
      synSubs.set(normalizeKey(sub.label), sub);
      synSubById.set(sub.id, sub);
    }
    const consumedSubs = new Set();
    for (const bdSub of bdSec.subsections) {
      // Each synthesis sub-section is claimed by AT MOST ONE BD sub-section so a
      // rename never spawns a duplicate row pointing at the same band.
      let existing = null;
      const subById = bdSub.synLinkId ? synSubById.get(bdSub.synLinkId) : null;
      if (subById && !consumedSubs.has(subById.id)) {
        existing = subById;
      } else {
        const subByLabel = synSubs.get(normalizeKey(bdSub.label));
        if (subByLabel && !consumedSubs.has(subByLabel.id)) existing = subByLabel;
      }
      if (existing) consumedSubs.add(existing.id);
      sec.subsections.push(
        existing
          ? {
              ...existing,
              label: bdSub.label,
              color: bdSub.color,
              archived: bdSub.archived,
              archiveRestoreIndex: bdSub.archiveRestoreIndex,
              synLinkId: existing.id,
            }
          : {
              id: uid('sub'),
              label: bdSub.label,
              color: bdSub.color,
              archived: bdSub.archived,
              archiveRestoreIndex: bdSub.archiveRestoreIndex,
              // Materialise only sub-sections the user added in this session;
              // pre-existing BD sub-sections absent from synthesis stay absent.
              isNew: Boolean(bdSub.isNew),
              startRow: null,
              endRow: null,
            }
      );
    }
    sections.push(sec);
  }
  // Preserve synthesis-only sections (no BD counterpart) so their rows are
  // never dropped on apply. They keep their original rows / order and are
  // appended after the BD-aligned sections.
  for (const synSec of synModel.sections) {
    if (consumed.has(synSec.id)) continue;
    sections.push({ ...synSec, subsections: [...(synSec.subsections || [])] });
  }
  const out = { ...synModel, sections };
  syncCaBandLabelsOntoSyn(out, bdModel);
  return out;
}

function attachBdStructureMeta(sheet, model) {
  const rows = new Set([5, 139]);
  for (const sec of (model && model.sections) || []) {
    if (sec.headerRow != null) rows.add(sec.headerRow);
  }
  const { rows: computedRows, canonicalByLabel } = computeSectionHeaderRows(sheet);
  for (const r of computedRows) rows.add(r);
  for (const sec of (model && model.sections) || []) {
    if (
      sec.caBand &&
      sec.headerRow != null &&
      sec.label &&
      !isFormulaLike(sec.label)
    ) {
      canonicalByLabel.set(String(sec.label).trim(), sec.headerRow);
    }
  }
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
  sheet.structureArchive = structureArchiveFromModel(m);
  sheet.archivedRowBands = collectArchivedRowBands(m);
  // New rows are now materialised — drop the "new" flag so re-applying the
  // same session model does not allocate duplicate rows.
  for (const sec of m.sections) {
    delete sec.isNew;
    for (const sub of sec.subsections) delete sub.isNew;
  }
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
  sheet.structureArchive = structureArchiveFromModel(m);
  sheet.archivedRowBands = collectArchivedRowBands(m);
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
