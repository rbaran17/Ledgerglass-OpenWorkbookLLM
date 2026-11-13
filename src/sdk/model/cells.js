/** Addresses, ranges and selection areas. */
export {
  parseRangeAddress,
  parseSelectionAreas,
  columnIndexToLetters,
  isCellAddressInRanges,
  countSelectionCells,
  enumerateSelectionCells,
  rangesIntersect,
  findOccupiedRangeBounds,
  formatOccupiedRangeAddress,
} from "../../app/03-workbook-model.js";

/** Sheet names, formula text and cell records. */
export {
  quoteSheetName,
  unquoteSheetName,
  isFormulaCellValue,
  normalizeCellRecordInput,
} from "../../app/03-workbook-model.js";

/** Cell styles, borders, CSS conversion and hyperlinks. */
export {
  validateCellStyle,
  cellStyleToCss,
  applyBorderEditTransaction,
  borderLineStyles,
  setCellHyperlink,
} from "../../app/03-workbook-model.js";

/** Number formats and their display text. */
export {
  validateNumberFormatSettings,
  buildNumberFormatCode,
  formatCellValueWithNumberFormat,
  applyNumberFormatTransaction,
  supportedCurrencyCodes,
  supportedNumberLocales,
} from "../../app/03-workbook-model.js";

/** AutoSum for an empty total cell. */
export { applyAutoSum } from "../../app/03-workbook-model.js";
