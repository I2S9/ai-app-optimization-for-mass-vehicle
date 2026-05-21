/** BD column letters after Y/Z insert following AD (X position). */
export const BD_POSITION_COLS = new Set(['AD', 'AE', 'AF']);
/** Sub-system L1 in raw export JSON (before Y/Z insert). */
export const BD_SUBSYSTEM_L1_COL_RAW = 'AP';
/** Sub-system L1 after {@link BD_POSITION_COLS} insert in {@link transformBdSheet}. */
export const BD_SUBSYSTEM_L1_COL = 'AR';
/** "Free field" (was AE) shifts to AG. */
export const BD_FREE_FIELD_COL = 'AG';
/** Métier / Trade (column AC in the BD sheet). */
export const BD_TRADE_COL = 'AC';
/** Front (AV) and Rear (AR) mass columns — always empty in the grid. */
export const BD_MASS_AV_AR_COLS = new Set(['W', 'X']);
