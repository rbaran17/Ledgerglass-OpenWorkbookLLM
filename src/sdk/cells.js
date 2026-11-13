/**
 * The cells area of the SDK: the model functions in ./model/cells.js, with edits recorded
 * (see ./recording.js). Reference page: docs/api/cells.md.
 */
import * as model from "./model/cells.js";
import { recordEdits } from "./recording.js";

const area = recordEdits(model);

export const parseRangeAddress = area.parseRangeAddress;
export const parseSelectionAreas = area.parseSelectionAreas;
export const columnIndexToLetters = area.columnIndexToLetters;
export const isCellAddressInRanges = area.isCellAddressInRanges;
export const countSelectionCells = area.countSelectionCells;
export const enumerateSelectionCells = area.enumerateSelectionCells;
export const rangesIntersect = area.rangesIntersect;
export const findOccupiedRangeBounds = area.findOccupiedRangeBounds;
export const formatOccupiedRangeAddress = area.formatOccupiedRangeAddress;
export const quoteSheetName = area.quoteSheetName;
export const unquoteSheetName = area.unquoteSheetName;
export const isFormulaCellValue = area.isFormulaCellValue;
export const normalizeCellRecordInput = area.normalizeCellRecordInput;
export const validateCellStyle = area.validateCellStyle;
export const cellStyleToCss = area.cellStyleToCss;
export const applyBorderEditTransaction = area.applyBorderEditTransaction;
export const borderLineStyles = area.borderLineStyles;
export const setCellHyperlink = area.setCellHyperlink;
export const validateNumberFormatSettings = area.validateNumberFormatSettings;
export const buildNumberFormatCode = area.buildNumberFormatCode;
export const formatCellValueWithNumberFormat =
  area.formatCellValueWithNumberFormat;
export const applyNumberFormatTransaction = area.applyNumberFormatTransaction;
export const supportedCurrencyCodes = area.supportedCurrencyCodes;
export const supportedNumberLocales = area.supportedNumberLocales;
export const applyAutoSum = area.applyAutoSum;
