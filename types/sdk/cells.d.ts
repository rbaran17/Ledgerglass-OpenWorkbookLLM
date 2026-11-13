/**
 * Declarations for the cells area of the SDK: addresses and ranges, cell records, styles, borders,
 * hyperlinks, number formats and AutoSum. Addresses are A1 text such as "B7". Ranges are zero-based
 * row and column indexes, so "A1" is row 0 and column 0.
 */
import type {
  Workbook,
  Sheet,
  Cell,
  CellPatch,
  RawCellContent,
} from "../index";

/** Zero-based bounds of a rectangle of cells. Both ends are inclusive. */
export interface RangeBounds {
  firstRow: number;
  lastRow: number;
  firstCol: number;
  lastCol: number;
}

/** One edge of a cell. */
export type CellBorderEdge = "top" | "right" | "bottom" | "left";

/** Line styles a border edge can use. "none" removes the edge. */
export type BorderLineStyle =
  "none" | "thin" | "medium" | "thick" | "dotted" | "dashed" | "double";

/** A border edge. The color is six or three digit hexadecimal and defaults to "#000000". */
export interface CellBorder {
  style: BorderLineStyle;
  color?: string;
  [extra: string]: unknown;
}

/** Border edges of a cell. Edges that are not listed keep their previous setting. */
export type CellBorders = Partial<Record<CellBorderEdge, CellBorder>>;

/** Direct formatting of a cell. Only the listed settings are accepted. */
export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  wrap?: boolean;
  fontFamily?: string;
  /** Whole number from 8 to 72. */
  fontSize?: number;
  /** Background color, as hexadecimal. */
  fill?: string;
  /** Text color, as hexadecimal. */
  color?: string;
  align?: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  /** Whole number from 0 to 15. */
  indent?: number;
  border?: "all" | "bottom" | "none";
  borders?: CellBorders;
  [extra: string]: unknown;
}

/** The CSS properties a cell style produces. Values are CSS text, except the numeric padding. */
export interface CellStyleCss {
  background?: string;
  color?: string;
  fontWeight: 400 | 700;
  fontStyle: "normal" | "italic";
  fontFamily?: string;
  /** Font size in points, such as "11pt". */
  fontSize?: string;
  textDecoration: string;
  textAlign: "left" | "center" | "right";
  verticalAlign?: "top" | "middle" | "bottom";
  paddingLeft: number;
  paddingRight: number;
  borderTop?: string;
  borderRight?: string;
  borderBottom?: string;
  borderLeft?: string;
  [extra: string]: unknown;
}

/** Options for CSS conversion. */
export interface CellStyleCssOptions {
  /** CSS border used on any edge the style does not set. */
  defaultBorder?: string;
  [extra: string]: unknown;
}

/** A link to a web address or email address, or to a cell on a sheet in the same workbook. */
export type Hyperlink = ExternalHyperlink | InternalHyperlink;

/** A link to an https://, http:// or mailto: address. The stored address is normalized, so a trailing slash can be added. */
export interface ExternalHyperlink {
  kind: "external";
  url: string;
  label?: string;
  [extra: string]: unknown;
}

/** A link to a cell on a sheet in the same workbook. */
export interface InternalHyperlink {
  kind: "internal";
  sheetId: string;
  address: string;
  label?: string;
  [extra: string]: unknown;
}

/** Options for edits that check protection and read-only state. */
export interface EditOptions {
  /** When true the edit throws, because the workbook is read-only. */
  readOnly?: boolean;
  [extra: string]: unknown;
}

/** Options for applyBorderEditTransaction. */
export interface BorderEditOptions extends EditOptions {
  /** Applies a preset to every cell in the selection, instead of only the edges in the style. */
  borders?: {
    preset: "all" | "outline" | "inside" | "clear";
    /** The line to use for the preset. Defaults to a thin black line. */
    line?: CellBorder;
  };
}

/** Number format categories. */
export type NumberFormatCategory =
  | "general"
  | "number"
  | "currency"
  | "accounting"
  | "percent"
  | "date"
  | "time"
  | "datetime"
  | "duration"
  | "custom";

/** Currency codes the number format settings accept. */
export type CurrencyCode =
  "USD" | "EUR" | "GBP" | "JPY" | "CAD" | "AUD" | "INR" | "CHF";

/** Display locales the number format settings accept. */
export type NumberLocale =
  "en-US" | "en-GB" | "de-DE" | "fr-FR" | "ja-JP" | "en-IN";

/** How a cell's number is shown. Which fields apply depends on the category. */
export interface NumberFormatSettings {
  category: NumberFormatCategory;
  /** Defaults to "en-US". */
  locale?: NumberLocale;
  /** Only for the currency and accounting categories. Defaults to "USD". */
  currency?: CurrencyCode;
  /** Whole number from 0 to 12. Defaults to 2, or 0 for JPY. */
  decimals?: number;
  /** Defaults to true, except for percent. */
  useGrouping?: boolean;
  /** Units, thousands or millions. Percent and custom formats use 1. */
  scale?: 1 | 1000 | 1000000;
  /** Defaults to "minus". Accounting always uses "parentheses". */
  negative?: "minus" | "parentheses";
  /** Excel-style pattern. Required for date, time, datetime, duration and custom formats. */
  pattern?: string;
  [extra: string]: unknown;
}

/** Number format settings after validation. Fields that do not apply to the category are absent. */
export interface NormalizedNumberFormat extends NumberFormatSettings {
  locale: NumberLocale;
  [extra: string]: unknown;
}

/** Display text for a value, with accounting parts when the category is accounting. */
export interface FormattedCellText {
  text: string;
  accounting?: {
    symbol: string;
    amount: string;
    /** "start" puts the currency symbol before the amount. */
    symbolPosition: string;
    [extra: string]: unknown;
  };
  [extra: string]: unknown;
}

/** Border line styles, in the order the settings accept them. */
export const borderLineStyles: readonly BorderLineStyle[];

/** Currency codes accepted by NumberFormatSettings.currency. */
export const supportedCurrencyCodes: readonly CurrencyCode[];

/** Locales accepted by NumberFormatSettings.locale. */
export const supportedNumberLocales: readonly NumberLocale[];

/** Parses an A1 address or a range such as "A1:C4" into zero-based bounds. Reversed corners are normalized, and "$" signs are ignored. Throws if the text is not a valid address or range, or if it lies outside the workbook limits. */
export function parseRangeAddress(rangeText: string): RangeBounds;

/** Parses comma-separated areas, or an array of areas, into their normalized A1 form. Duplicates are removed and the order of the first occurrence is kept. Throws if the list is empty, has more than 256 areas, or contains an invalid area. */
export function parseSelectionAreas(
  selectionInput: string | string[],
): string[];

/** Converts a zero-based column index to its letters, so 0 gives "A" and 26 gives "AA". Throws if the index is not an integer from 0 to 16383. */
export function columnIndexToLetters(columnIndex: number): string;

/** Returns true when the active cell address lies inside one of the given areas. Throws if the address or an area is invalid. */
export function isCellAddressInRanges(
  selectionAreas: string[],
  active: string,
): boolean;

/** Counts the distinct cells covered by one or more areas. Overlapping cells are counted once. Throws if an area is invalid. */
export function countSelectionCells(selectionAreas: string | string[]): number;

/** Lists the A1 addresses of every cell in one or more areas, in the order the areas are given and row by row inside each area. Throws if the selection has more cells than maxCellCount, which defaults to 500. */
export function enumerateSelectionCells(
  selection: string | string[],
  maxCellCount?: number,
): string[];

/** Returns true when two rectangles share at least one cell. Both arguments must come from parseRangeAddress or have the same bounds fields. */
export function rangesIntersect(
  first: RangeBounds,
  second: RangeBounds,
): boolean;

/** Returns the smallest rectangle that holds every populated cell on the sheet, or null when the sheet is empty. Formatting counts as occupied unless shouldIncludeFormatting is false. Does not change the sheet. */
export function findOccupiedRangeBounds(
  sheet: Sheet,
  shouldIncludeFormatting?: boolean,
): RangeBounds | null;

/** Returns the A1 address of the occupied rectangle of a sheet, such as "A1:C5", or "A1" when nothing is occupied. Formatting counts unless shouldIncludeFormatting is false. Does not change the sheet. */
export function formatOccupiedRangeAddress(
  sheet: Sheet,
  shouldIncludeFormatting?: boolean,
): string;

/** Wraps a sheet name in single quotes for use in a formula, doubling any single quote inside it. The name is always quoted, even when quotes are not required. */
export function quoteSheetName(sheetName: string): string;

/** Removes one pair of surrounding single quotes and turns doubled quotes back into one. Text without a leading quote is returned unchanged. */
export function unquoteSheetName(quotedSheetName: string): string;

/** Returns true when the cell holds a formula, meaning its raw text starts with "=" and it was not stored as literal text. Returns false for a missing cell. */
export function isFormulaCellValue(cell: Cell | null | undefined): boolean;

/** Checks a cell record and returns a copy with only the known fields. Throws if raw is missing or is not text, a finite number, a boolean or null. */
export function normalizeCellRecordInput(cell: Cell): Cell;

/** Checks a style object and returns a copy that holds only the validated settings. Throws if it has an unknown setting or a value that does not fit the setting. The input is not changed. */
export function validateCellStyle(cellStyle: CellStyle): CellStyle;

/** Converts a cell style and its value to CSS properties. Numbers are right aligned and other values left aligned unless the style sets an alignment. Does not change its arguments. */
export function cellStyleToCss(
  cellStyle: CellStyle | undefined,
  cellValue: RawCellContent,
  options?: CellStyleCssOptions,
): CellStyleCss;

/**
 * Merges a style into the cells of a selection and returns a new workbook, leaving the argument unchanged.
 * The returned workbook's revision is one higher, and its events has one new event, "Apply border edit transaction", with an empty change list.
 * A string is one area, so pass an array for several. Throws if the workbook is read-only, the sheet is protected, the selection has more than 500 cells, or it touches a merged area.
 */
export function applyBorderEditTransaction(
  workbook: Workbook,
  sheetId: string,
  selectionAreas: string | string[],
  stylePatch: CellStyle,
  options?: BorderEditOptions,
): Workbook;

/**
 * Sets or removes the hyperlink of one cell and returns a new workbook. Pass null to remove the link.
 * The returned workbook's revision is one higher, and its events has one new event, "Set cell hyperlink", with an empty change list.
 * Throws if the link is invalid, the target sheet is missing, the workbook is read-only, or the cell is inside a merged area and is not its first cell.
 */
export function setCellHyperlink(
  workbook: Workbook,
  sheetId: string,
  address: string,
  hyperlink: Hyperlink | null,
  options?: EditOptions,
): Workbook;

/** Checks number format settings and returns them with defaults filled in. Throws if a field is unknown, out of range, or does not suit the category. The input is not changed. */
export function validateNumberFormatSettings(
  settings: NumberFormatSettings,
): NormalizedNumberFormat;

/** Returns the Excel number format code for the settings, such as "#,##0.00". Throws if the settings are invalid. */
export function buildNumberFormatCode(settings: NumberFormatSettings): string;

/**
 * Formats a value for display. Booleans show TRUE or FALSE, null shows an empty string, and an invalid number shows "#NUM!".
 * dateSystem is "1900" by default, and any value other than "1904" is treated as "1900". Throws if the settings are invalid.
 */
export function formatCellValueWithNumberFormat(
  value: RawCellContent,
  settings: NumberFormatSettings,
  dateSystem?: "1900" | "1904",
): FormattedCellText;

/**
 * Applies a number format to every cell of a selection and returns a new workbook, leaving the argument unchanged.
 * The returned workbook's revision is one higher, and its events has one new event, "Apply number format transaction", with an empty change list.
 * A string is one area, so pass an array for several. Throws if the settings are invalid, the workbook is read-only, the sheet is protected, or the selection has more than 500 cells.
 */
export function applyNumberFormatTransaction(
  workbook: Workbook,
  sheetId: string,
  selectionAreas: string | string[],
  numberFormat: NumberFormatSettings,
  options?: EditOptions,
): Workbook;

/**
 * Returns a patch that sums the numbers or formulas directly above or to the left of an empty cell.
 * The patch is not applied; pass it to a cell patch transaction to write it. Throws if the cell is not empty, is not a single address, or has no numbers to sum.
 */
export function applyAutoSum(
  workbook: Workbook,
  sheetId: string,
  activeCellAddress: string,
): CellPatch;
