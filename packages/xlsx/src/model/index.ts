/**
 * L1 - the document model. Rows, cells, references, merges and names as
 * structure, with edits that keep every row-dependent invariant.
 *
 * The layer's seam lives in `./types.js`; the classes here implement it.
 */
export {
  decodeCell,
  decodeColumn,
  encodeCell,
  encodeColumn,
  relocateFormulaRows,
  relocateReference,
  relocateSqref,
  RowRelocation,
  DELETED_REFERENCE,
  type CellReference,
} from "./references.js";
export { excelSerialToDate, isDateFormatCode, StyleTable } from "./styles.js";
export {
  WorksheetCell,
  WorksheetModel,
  WorksheetRow,
  type RelocationCounters,
  type WorksheetCellValue,
  type WorksheetHost,
} from "./worksheet-model.js";
export { WorkbookModel, type DeleteRowsOptions } from "./workbook-model.js";
