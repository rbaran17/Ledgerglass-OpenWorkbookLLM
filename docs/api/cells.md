# Cells

The cells area holds the functions that read and write cell addresses, ranges and selections, and the functions that set cell styles, borders, hyperlinks, number formats and AutoSum formulas. Addresses are A1 text such as `B7`. Range bounds are zero-based, so `A1` is row 0 and column 0. The edit functions take a workbook and return a new one; the argument is never changed.

## Addresses and ranges

### parseRangeAddress

```ts
export function parseRangeAddress(rangeText: string): RangeBounds;
```

Parses `A1` or a range such as `A1:C4` into zero-based bounds. Reversed corners are normalized and `$` signs are ignored. Throws if the text is not a valid address or range, or if a cell lies outside the workbook limits.

```js
import { parseRangeAddress } from "@ledgerglass/workbook";

parseRangeAddress("b5:a1");
// { firstRow: 0, lastRow: 4, firstCol: 0, lastCol: 1 }
```

### parseSelectionAreas

```ts
export function parseSelectionAreas(
  selectionInput: string | string[],
): string[];
```

Parses comma-separated areas, or an array of areas, into their normalized A1 form. Duplicates are removed and the first occurrence keeps its place. Throws if the list is empty, has more than 256 areas, or contains an invalid area.

```js
import { parseSelectionAreas } from "@ledgerglass/workbook";

parseSelectionAreas("b2:a1, c3 ,a1,$D$4");
// ["A1:B2", "C3", "A1", "D4"]
```

### columnIndexToLetters

```ts
export function columnIndexToLetters(columnIndex: number): string;
```

Converts a zero-based column index to its letters, so `0` gives `A` and `26` gives `AA`. Throws if the index is not an integer from 0 to 16383.

### isCellAddressInRanges

```ts
export function isCellAddressInRanges(
  selectionAreas: string[],
  active: string,
): boolean;
```

Returns `true` when the active cell lies inside one of the areas. Throws if the active address is not valid. The active address is case-insensitive.

### countSelectionCells

```ts
export function countSelectionCells(selectionAreas: string | string[]): number;
```

Counts the distinct cells covered by one or more areas. A cell that two areas share is counted once. Throws if an area is invalid.

### enumerateSelectionCells

```ts
export function enumerateSelectionCells(
  selection: string | string[],
  maxCellCount?: number,
): string[];
```

Lists the A1 address of every cell in one or more areas, area by area and row by row inside each area. Throws if the selection has more cells than `maxCellCount`, which defaults to 500.

```js
import { enumerateSelectionCells } from "@ledgerglass/workbook";

enumerateSelectionCells("B1:C2,A1");
// ["B1", "C1", "B2", "C2", "A1"]
```

### rangesIntersect

```ts
export function rangesIntersect(
  first: RangeBounds,
  second: RangeBounds,
): boolean;
```

Returns `true` when two rectangles share at least one cell. Pass bounds from `parseRangeAddress`. Does not change either argument.

### findOccupiedRangeBounds

```ts
export function findOccupiedRangeBounds(
  sheet: Sheet,
  shouldIncludeFormatting?: boolean,
): RangeBounds | null;
```

Returns the smallest rectangle that holds every populated cell on the sheet, or `null` when the sheet is empty. A cell with formatting counts as occupied unless `shouldIncludeFormatting` is `false`. Does not change the sheet.

### formatOccupiedRangeAddress

```ts
export function formatOccupiedRangeAddress(
  sheet: Sheet,
  shouldIncludeFormatting?: boolean,
): string;
```

Returns the occupied rectangle as A1 text, such as `A1:C5`, or `A1` when nothing is occupied. Formatting counts unless `shouldIncludeFormatting` is `false`. Does not change the sheet.

## Sheet names, formulas and cell records

### quoteSheetName

```ts
export function quoteSheetName(sheetName: string): string;
```

Wraps a sheet name in single quotes for use in a formula, and doubles any single quote inside it. The name is always quoted, even when quotes are not needed.

```js
import { quoteSheetName } from "@ledgerglass/workbook";

quoteSheetName("O'Brien"); // "'O''Brien'"
```

### unquoteSheetName

```ts
export function unquoteSheetName(quotedSheetName: string): string;
```

Removes one pair of surrounding single quotes and turns doubled quotes back into one. Text without a leading quote is returned unchanged. Use it only on text that was quoted by `quoteSheetName`.

```js
import { unquoteSheetName } from "@ledgerglass/workbook";

unquoteSheetName("'O''Brien'"); // "O'Brien"
```

### isFormulaCellValue

```ts
export function isFormulaCellValue(cell: Cell | null | undefined): boolean;
```

Returns `true` when the cell holds a formula: its raw text starts with `=` and it was not stored as literal text. Returns `false` for a value, for a formula stored as literal text, and for a missing cell.

### normalizeCellRecordInput

```ts
export function normalizeCellRecordInput(cell: Cell): Cell;
```

Checks a cell record and returns a copy that holds only the known fields. Unknown fields are dropped and a number format is filled in with its defaults. Throws if `raw` is missing or is not text, a finite number, a boolean or `null`.

## Styles and hyperlinks

### validateCellStyle

```ts
export function validateCellStyle(cellStyle: CellStyle): CellStyle;
```

Checks a style object and returns a copy that holds only the validated settings. Throws if the object has an unknown setting or a value that does not fit it, such as a color that is not hexadecimal or a font size outside 8 to 72. The input is not changed.

```js
import { validateCellStyle } from "@ledgerglass/workbook";

validateCellStyle({ borders: { left: { style: "thin" } } });
// { borders: { left: { style: "thin", color: "#000000" } } }
```

### cellStyleToCss

```ts
export function cellStyleToCss(
  cellStyle: CellStyle | undefined,
  cellValue: RawCellContent,
  options?: CellStyleCssOptions,
): CellStyleCss;
```

Converts a style and the value in the cell to CSS properties. A number is right aligned, and any other value is left aligned, unless the style sets an alignment. `options.defaultBorder` is used on every edge the style does not set. Does not change its arguments.

### applyBorderEditTransaction

```ts
export function applyBorderEditTransaction(
  workbook: Workbook,
  sheetId: string,
  selectionAreas: string | string[],
  stylePatch: CellStyle,
  options?: BorderEditOptions,
): Workbook;
```

Merges a style into every cell of a selection and returns a new workbook. A string is one area, so `"A1:B2,C3"` throws; pass an array for several areas. The selection may hold at most 500 cells, and it may not touch a merged area. The returned workbook's `revision` is one higher, and its `events` has one new event, `Apply border edit transaction`, with an empty change list.

```js
import {
  applyBorderEditTransaction,
  createBlankWorkbook,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Budget");
const sheetId = workbook.sheets[0].id;

const next = applyBorderEditTransaction(
  workbook,
  sheetId,
  "A1:B2",
  {},
  { borders: { preset: "outline", line: { style: "medium" } } },
);
next.sheets[0].cells.A1.style.borders;
// { top: { style: "medium", color: "#000000" }, left: { style: "medium", color: "#000000" } }
```

### borderLineStyles

```ts
export const borderLineStyles: readonly BorderLineStyle[];
```

Lists the border line styles in order: `none`, `thin`, `medium`, `thick`, `dotted`, `dashed` and `double`. The array is not frozen at run time, so do not change it.

### setCellHyperlink

```ts
export function setCellHyperlink(
  workbook: Workbook,
  sheetId: string,
  address: string,
  hyperlink: Hyperlink | null,
  options?: EditOptions,
): Workbook;
```

Sets the hyperlink of one cell and returns a new workbook. Pass `null` to remove the link. Throws if the link is invalid, the sheet is missing, the workbook is read-only, or the cell is inside a merged area and is not its first cell. The returned workbook's `revision` is one higher, and its `events` has one new event, `Set cell hyperlink`, with an empty change list.

```js
import { createBlankWorkbook, setCellHyperlink } from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Links");
const sheetId = workbook.sheets[0].id;

const next = setCellHyperlink(workbook, sheetId, "A1", {
  kind: "external",
  url: "https://example.com/docs",
  label: "Docs",
});
next.sheets[0].cells.A1.hyperlink;
// { kind: "external", url: "https://example.com/docs", label: "Docs" }
```

## Number formats

### validateNumberFormatSettings

```ts
export function validateNumberFormatSettings(
  settings: NumberFormatSettings,
): NormalizedNumberFormat;
```

Checks number format settings and returns them with the defaults filled in. Throws if a field is unknown, out of range, or does not suit the category, such as a currency on a number format. The input is not changed.

```js
import { validateNumberFormatSettings } from "@ledgerglass/workbook";

validateNumberFormatSettings({ category: "currency", currency: "JPY" });
// { category: "currency", locale: "en-US", currency: "JPY", decimals: 0, useGrouping: true, scale: 1, negative: "minus" }
```

### buildNumberFormatCode

```ts
export function buildNumberFormatCode(settings: NumberFormatSettings): string;
```

Returns the Excel number format code for the settings, such as `#,##0.00`. Non-US locales add an LCID prefix such as `[$-407]`. Throws if the settings are invalid.

```js
import { buildNumberFormatCode } from "@ledgerglass/workbook";

buildNumberFormatCode({ category: "percent", decimals: 1 }); // "0.0%"
```

### formatCellValueWithNumberFormat

```ts
export function formatCellValueWithNumberFormat(
  value: RawCellContent,
  settings: NumberFormatSettings,
  dateSystem?: "1900" | "1904",
): FormattedCellText;
```

Formats a value for display and returns `{ text }`. Booleans show `TRUE` or `FALSE`, `null` shows an empty string, and `NaN` or an infinite number shows `#NUM!`. The accounting category also returns its symbol and amount. Any `dateSystem` other than `"1904"` uses the 1900 system. Throws if the settings are invalid.

```js
import { formatCellValueWithNumberFormat } from "@ledgerglass/workbook";

formatCellValueWithNumberFormat(1234.5, { category: "number", decimals: 2 })
  .text;
// "1,234.50"
formatCellValueWithNumberFormat(0.256, { category: "percent", decimals: 1 })
  .text;
// "25.6%"
```

### applyNumberFormatTransaction

```ts
export function applyNumberFormatTransaction(
  workbook: Workbook,
  sheetId: string,
  selectionAreas: string | string[],
  numberFormat: NumberFormatSettings,
  options?: EditOptions,
): Workbook;
```

Applies a number format to every cell of a selection and returns a new workbook. A string is one area, so `"A1:B2,C3"` throws; pass an array for several areas. The selection may hold at most 500 cells. The returned workbook's `revision` is one higher, and its `events` has one new event, `Apply number format transaction`, with an empty change list.

```js
import {
  applyNumberFormatTransaction,
  createBlankWorkbook,
} from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Prices");
const sheetId = workbook.sheets[0].id;

const next = applyNumberFormatTransaction(workbook, sheetId, ["A1:A10", "C1"], {
  category: "currency",
  currency: "EUR",
  decimals: 2,
});
next.sheets[0].cells.A1.numFmt;
// "\"€\"#,##0.00"
```

### supportedCurrencyCodes

```ts
export const supportedCurrencyCodes: readonly CurrencyCode[];
```

Lists the currency codes that `currency` accepts: `USD`, `EUR`, `GBP`, `JPY`, `CAD`, `AUD`, `INR` and `CHF`. The array is not frozen at run time, so do not change it.

### supportedNumberLocales

```ts
export const supportedNumberLocales: readonly NumberLocale[];
```

Lists the locales that `locale` accepts: `en-US`, `en-GB`, `de-DE`, `fr-FR`, `ja-JP` and `en-IN`. The array is not frozen at run time, so do not change it.

## AutoSum

### applyAutoSum

```ts
export function applyAutoSum(
  workbook: Workbook,
  sheetId: string,
  activeCellAddress: string,
): CellPatch;
```

Returns a patch that puts a `SUM` formula in an empty cell. The formula covers the numbers or formulas directly above the cell, or else to its left. The patch is not applied; pass it to a cell patch transaction to write it. Throws if the cell is not empty, if more than one cell is given, or if there is nothing to sum next to it.

```js
import { applyAutoSum, createBlankWorkbook } from "@ledgerglass/workbook";

const workbook = createBlankWorkbook("Totals");
const sheetId = workbook.sheets[0].id;
const sheet = workbook.sheets[0];
const filled = {
  ...workbook,
  sheets: [
    { ...sheet, cells: { A1: { raw: 10 }, A2: { raw: 20 }, A3: { raw: 30 } } },
  ],
};

applyAutoSum(filled, sheetId, "A4").cell;
// { raw: "=SUM(A1:A3)" }
```
