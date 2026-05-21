import fs from 'fs';
import { transformBdSheet } from '../web/js/sheetTransform.js';
import {
  buildCellMap,
  displayCellValue,
  bdSubsystemL1Col,
  bdSubsystemL2Col,
  bdMassCol,
  isSubSectionRow,
  shouldDateColBlue,
  isStructureRow,
} from '../web/js/bdStore.js';

const t = transformBdSheet(
  JSON.parse(fs.readFileSync('web/public/data/bd-sheet.json', 'utf8'))
);
const map = buildCellMap(t.cells, t.headerRows);
const mass = bdMassCol(t);
const l1 = bdSubsystemL1Col(t);
const l2 = bdSubsystemL2Col(t);
const sh = t.sectionHeaderRows;
for (const r of [7, 8, 9, 246]) {
  console.log(r, {
    web: displayCellValue(map, r, mass, sh, t.canonicalSectionByLabel, l1, l2),
    cell: map.get(`${r}:${mass}`),
    sub: isSubSectionRow(map, r, sh, l2),
    blue: shouldDateColBlue(map, r, sh),
    struct: isStructureRow(map, r, sh),
  });
}
