# Worksheets

Worksheet functions add, remove, rename, reorder and hide the sheets of a workbook, and set the workbook's name, engagement and period. They also set frozen and split panes, merged cells, hidden rows and columns, default sizes, outline groups and protection for a single sheet. Functions that change a workbook return a new one and leave the argument as it was. The functions here do not change `revision` or add an entry to `events`, unlike cell edits.

### addWorksheet

```ts
export function addWorksheet(
  workbook: Workbook,
  worksheetName: string,
): Workbook;
```

Adds a sheet named `worksheetName` at the end of the tab order. Names are 1 to 31 characters, cannot contain `\ / [ ] * ? :` or control characters, and must be unique without regard to case. Throws if the name is invalid or already used, or if the workbook structure is protected.

```js
import { addWorksheet, createBlankWorkbook } from "@ledgerglass/workbook";

const workbook = addWorksheet(createBlankWorkbook("Budget"), "Data");
console.log(workbook.sheets.map((sheet) => sheet.name)); // ["Sheet1", "Data"]
```

### renameWorksheet

```ts
export function renameWorksheet(
  workbook: Workbook,
  sheetId: string,
  newName: string,
): Workbook;
```

Renames the sheet with id `sheetId` and rewrites formulas that used the old name, so `=Data!B2` becomes `='Summary'!B2` after renaming `Data` to `Summary`. Returns the same workbook when the name does not change. Throws if the name is invalid or used by another sheet, or if the sheet id is not found.

```js
import { renameWorksheet } from "@ledgerglass/workbook";

const renamed = renameWorksheet(workbook, dataSheetId, "Summary");
```

### moveWorksheet

```ts
export function moveWorksheet(
  workbook: Workbook,
  sheetId: string,
  targetIndex: number,
): Workbook;
```

Moves the sheet to `targetIndex` in the tab order, counted from zero. Returns the same workbook when the sheet is already at that index. Throws if the index is not an integer inside the workbook, or if a formula uses a three-dimensional reference such as `Data:Other!A1`.

### duplicateWorksheet

```ts
export function duplicateWorksheet(
  workbook: Workbook,
  sheetId: string,
): Workbook;
```

Adds a copy of the sheet directly after the original, named with the next free `(n)` suffix such as `Data (2)`. Formulas on the copy that refer to the original sheet are rewritten to refer to the copy. Throws if a formula uses a three-dimensional reference.

```js
import { duplicateWorksheet } from "@ledgerglass/workbook";

const withCopy = duplicateWorksheet(workbook, dataSheetId);
console.log(withCopy.sheets.at(-1).name); // "Data (2)"
```

### deleteWorksheet

```ts
export function deleteWorksheet(
  workbook: Workbook,
  sheetId: string,
  isDeletionConfirmed?: boolean,
): Workbook;
```

Removes the sheet. Formulas that referred to it become `#REF!`. Throws if anything in the workbook depends on the sheet and `isDeletionConfirmed` is not `true`, if it is the only worksheet, or if no visible worksheet would remain. Use `describeSheetDeletionImpacts` first to list what would change.

```js
import {
  deleteWorksheet,
  describeSheetDeletionImpacts,
} from "@ledgerglass/workbook";

const impacts = describeSheetDeletionImpacts(workbook, dataSheetId);
const trimmed = deleteWorksheet(workbook, dataSheetId, impacts.length > 0);
```

### setWorksheetVisibility

```ts
export function setWorksheetVisibility(
  workbook: Workbook,
  sheetId: string,
  visibility: WorksheetVisibility,
): Workbook;
```

Sets the sheet to `"visible"`, `"hidden"` or `"veryHidden"`. A very hidden sheet is also marked hidden. Returns the same workbook when the visibility does not change. Throws for any other value, or when the change would leave no visible worksheet.

### setWorksheetTabColor

```ts
export function setWorksheetTabColor(
  workbook: Workbook,
  sheetId: string,
  tabColor: string | undefined,
): Workbook;
```

Sets the tab colour to a six-digit colour such as `"#FF0000"`, which is stored in lower case, or removes it when `tabColor` is `undefined`. Throws if the colour is not a `#` followed by six hexadecimal digits.

### renameWorkbook

```ts
export function renameWorkbook(workbook: Workbook, newName: string): Workbook;
```

Sets the workbook name after trimming spaces. Throws if the trimmed name is empty or longer than 200 characters.

### updateWorkbookProperties

```ts
export function updateWorkbookProperties(
  workbook: Workbook,
  propertyUpdates: WorkbookPropertyUpdates,
): Workbook;
```

Changes any of `name`, `engagement` and `period`, and keeps the fields that are not given. An empty engagement or period removes it from the metadata. Throws for any other field name, an engagement longer than 200 characters, or a period longer than 100 characters.

```js
import { updateWorkbookProperties } from "@ledgerglass/workbook";

const audit = updateWorkbookProperties(workbook, {
  engagement: "FY26 audit",
  period: "Q1",
});
console.log(audit.metadata.engagement); // "FY26 audit"
```

### createWorkbookFromTemplate

```ts
export function createWorkbookFromTemplate(
  workbook: Workbook,
  details: WorkbookTemplateDetails,
): Workbook;
```

Returns a new workbook with a copy of the template's sheets and settings, a new id, revision 0, and the given name, engagement and period. The metadata records the template's id, name and revision. The template is not changed. Throws if any of the three details is missing or empty.

### protectWorkbookStructure

```ts
export function protectWorkbookStructure(
  workbook: Workbook,
  password: string,
): Promise<Workbook>;
```

Protects the workbook structure, so that worksheets cannot be added or renamed while it is protected. The password is stored as a salted SHA-256 hash; an empty string protects the structure without a password. Throws if `password` is not text or is longer than 200 characters.

### unprotectWorkbookStructure

```ts
export function unprotectWorkbookStructure(
  workbook: Workbook,
  password: string,
): Promise<Workbook>;
```

Removes the structure protection. Throws if a password was set and `password` does not match it. Resolves to the same workbook when the structure is not protected.

```js
import {
  protectWorkbookStructure,
  unprotectWorkbookStructure,
} from "@ledgerglass/workbook";

const locked = await protectWorkbookStructure(workbook, "review-2026");
const open = await unprotectWorkbookStructure(locked, "review-2026");
```

### updateSheetProtection

```ts
export function updateSheetProtection(
  workbook: Workbook,
  sheetId: string,
  password: string,
  isProtecting: boolean,
): Promise<Workbook>;
```

Protects the sheet when `isProtecting` is `true`, or removes its protection when it is `false`. An empty password protects or unprotects without one. Throws when protecting a sheet that is already protected, when unprotecting with a password that does not match, or when `password` is not text.

While a sheet is protected, freeze and split panes, sizes, hidden rows and columns, and outline groups cannot be changed. Renaming the sheet and changing its visibility are still allowed.

### summarizeWorksheets

```ts
export function summarizeWorksheets(
  workbook: Workbook,
  options?: WorksheetSummaryOptions,
): WorksheetSummary[];
```

Returns one summary per sheet in tab order, with its id, name, visibility, tab colour, stored cell count, formula count and whether the structure is locked. `options.allowedSheetIds` limits the list to the given ids. The workbook is not changed.

```js
import { summarizeWorksheets } from "@ledgerglass/workbook";

for (const sheet of summarizeWorksheets(workbook)) {
  console.log(sheet.name, sheet.visibility, sheet.formulaCellCount);
}
```

### describeSheetDeletionImpacts

```ts
export function describeSheetDeletionImpacts(
  workbook: Workbook,
  sheetId: string,
): DeletionImpact[];
```

Lists each formula elsewhere in the workbook that deleting the sheet would change. Each entry has a `kind`, a `label` for people to read, and the `sheetId` and `address` where it applies. The list is empty when nothing depends on the sheet. Throws if the sheet id is not found.

### findSheetOrThrow

```ts
export function findSheetOrThrow(workbook: Workbook, sheetId: string): Sheet;
```

Returns the sheet with the given id, as it is stored in the workbook. Throws `Sheet not found: <id>` when no sheet has that id.

### setFreezePanes

```ts
export function setFreezePanes(
  workbook: Workbook,
  sheetId: string,
  frozenRows: number,
  frozenCols?: number,
): Workbook;
```

Freezes `frozenRows` rows at the top and `frozenCols` columns at the left of the sheet. Omitting `frozenCols` freezes no columns. Any split panes are removed. Throws if a count is negative or larger than the sheet.

### setFreezePanesAtCell

```ts
export function setFreezePanesAtCell(
  workbook: Workbook,
  sheetId: string,
  cellAddress: string,
): Workbook;
```

Freezes the rows above and the columns to the left of the cell, so `"B3"` freezes two rows and one column. `"A1"` freezes nothing. Throws for an address that is not a valid cell.

```js
import { setFreezePanesAtCell } from "@ledgerglass/workbook";

const headers = setFreezePanesAtCell(workbook, dataSheetId, "B3");
```

### setSplitPanes

```ts
export function setSplitPanes(
  workbook: Workbook,
  sheetId: string,
  splitPanes: SplitPaneState | undefined,
): Workbook;
```

Splits the sheet at the zero-based cell `row` and `col`, with `horizontal` and `vertical` divider positions as percentages. Setting a split clears any frozen panes. Passing `undefined` removes the split. Throws if neither divider is given, a divider is outside 20 to 80, or the anchor cell is outside the sheet.

```js
import { setSplitPanes } from "@ledgerglass/workbook";

const split = setSplitPanes(workbook, dataSheetId, {
  horizontal: 50,
  row: 1,
  col: 0,
});
```

### mergeWorksheetCells

```ts
export function mergeWorksheetCells(
  workbook: Workbook,
  sheetId: string,
  rangeText: string,
  unmerge?: boolean,
): Workbook;
```

Merges the range into one cell and records it in the sheet's `merges`. The top-left cell keeps its value and is centred. With `unmerge` set to `true`, removes the merges that overlap the range instead. Throws for a range over 500 cells, a merge that overlaps an existing merge, or a protected sheet.

```js
import { mergeWorksheetCells } from "@ledgerglass/workbook";

const titled = mergeWorksheetCells(workbook, dataSheetId, "B2:D2");
```

### setHiddenAxisRange

```ts
export function setHiddenAxisRange(
  workbook: Workbook,
  sheetId: string,
  axis: WorksheetAxis,
  start: number,
  end: number,
  isHidden: boolean,
): Workbook;
```

Hides or shows the rows or columns from `start` to `end`, both zero-based and inclusive. Hiding a range that touches an existing hidden run merges them. Throws if `axis` is not `"row"` or `"column"`, if `isHidden` is not a boolean, if the range is outside the sheet, or if the sheet is protected.

### readHiddenAxisIntervals

```ts
export function readHiddenAxisIntervals(
  sheet: Sheet,
  axis: WorksheetAxis,
): HiddenAxisInterval[];
```

Returns the hidden rows or columns of a sheet as sorted, merged intervals with inclusive `start` and `end` values. Returns an empty list when nothing is hidden. The sheet is not changed. Throws if `axis` is not `"row"` or `"column"`.

### applySheetSizeEdits

```ts
export function applySheetSizeEdits(
  workbook: Workbook,
  sheetId: string,
  sizeEdits: SheetSizeEdits,
  options?: SheetSizeEditOptions,
): Workbook;
```

Sets the default row height (21 to 546 pixels) and default column width in pixels, and clears custom row heights or column widths when `resetRowOverrides` or `resetColumnOverrides` is `true`. Returns the same workbook when nothing changes. Throws for an unknown key or a size out of range, when `options.readOnly` is set, when `options.expectedSnapshot` no longer matches the workbook, or when the sheet is protected.

```js
import { applySheetSizeEdits } from "@ledgerglass/workbook";

const taller = applySheetSizeEdits(workbook, dataSheetId, { rowHeight: 30 });
```

### removeOutlineGroup

```ts
export function removeOutlineGroup(
  workbook: Workbook,
  sheetId: string,
  outlineGroupId: string,
  options?: OutlineEditOptions,
): Workbook;
```

Removes the outline group with the given id from the sheet. Throws if no group on the sheet has that id, or if the sheet is protected.

### setOutlineGroupCollapsed

```ts
export function setOutlineGroupCollapsed(
  workbook: Workbook,
  sheetId: string,
  outlineGroupId: string,
  isCollapsed: boolean,
  options?: OutlineEditOptions,
): Workbook;
```

Collapses or expands the outline group with the given id. Throws if `isCollapsed` is not a boolean, if no group has that id, or if the sheet is protected.
