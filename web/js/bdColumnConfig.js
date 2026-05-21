/** BD column letters after Y/Z insert following AD (X position). */
export const BD_POSITION_COLS = new Set(['AD', 'AE', 'AF']);
/** Sub-system L1 in raw export JSON (before Y/Z insert). */
export const BD_SUBSYSTEM_L1_COL_RAW = 'AP';
/** Sub-system L1 after {@link BD_POSITION_COLS} insert in {@link transformBdSheet}. */
export const BD_SUBSYSTEM_L1_COL = 'AR';
/** Internal Excel lookups (hidden in grid): AQ → AS, AT → AV, AR → AT. */
export const BD_SUBSYSTEM_L2_COL_RAW = 'AS';
export const BD_SUBSYSTEM_L2_COL = 'AU';
/** Sub-System Design Dpt (raw AU → AW). */
export const BD_DESIGN_DEPT_COL_RAW = 'AU';
export const BD_DESIGN_DEPT_COL = 'AW';
/** "Free field" (was AE) shifts to AG. */
export const BD_FREE_FIELD_COL = 'AG';
/** Métier / Trade (column AC in the BD sheet). */
export const BD_TRADE_COL = 'AC';
/** Filter / legend column Project (B). */
export const BD_PROJECT_COL = 'B';
/** First vehicle-data column for row striping (Silhouette = C). */
export const BD_SILHOUETTE_COL = 'C';
export const BD_CODIFICATION_COL = 'R';
export const BD_TITLE_COL = 'T';
/** Mass (Masse) — must match Excel cached values in export. */
export const BD_MASS_COL = 'V';
/** Default grid width for Codification (narrower than Excel, not minimal). */
export const BD_CODIFICATION_WIDTH = 36;
/** Modular type (Type modulaire) — grey like Codification / Title. */
export const BD_MODULAR_TYPE_COL = 'AJ';
/** Front (AV) and Rear (AR) mass columns — always empty in the grid. */
export const BD_MASS_AV_AR_COLS = new Set(['W', 'X']);
