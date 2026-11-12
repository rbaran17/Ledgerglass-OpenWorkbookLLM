# Tables, filters and pivots

Worksheet tables give a range a header row, named columns and optional calculated columns. Filters choose which rows of a header-and-data range match. Pivot tables summarise a source range by row and column fields. Every function here either reads a workbook or returns a changed copy of one, and results are plain workbook objects.

## How changes come back

Most functions here return a new workbook and leave the workbook they were given unchanged. Preview functions return a plan first, and a matching `apply…` function turns the plan into a workbook. A plan keeps a JSON snapshot of the workbook it was made from, so applying it to a different workbook throws.

These functions do not change `revision` or add to `events`. A host that keeps an event history must record these changes itself.

## Tables

### createWorksheetTable

```ts
function createWorksheetTable(
  workbook: Workbook,
  sheetId: string,
  options: CreateWorksheetTableOptions,
): Workbook;
```

Adds a table over `options.range` and returns a new workbook. The first row of the range gives the column names unless `headers` is given. It throws if the sheet is protected, the range overlaps merged cells or another table, the name is taken or invalid, or the headers need review (blank, repeated, non-text or formula headers).

```js
import {
  createBlankWorkbook,
  createWorksheetTable,
} from "@ledgerglass/workbook";

const book = createBlankWorkbook("Sales");
const sheet = book.sheets[0];
Object.assign(sheet.cells, {
  A1: { raw: "Region" },
  B1: { raw: "Units" },
  C1: { raw: "Price" },
  A2: { raw: "North" },
  B2: { raw: 10 },
  C2: { raw: 2.5 },
});
const withTable = createWorksheetTable(book, sheet.id, {
  name: "Sales",
  range: "A1:C2",
});
withTable.sheets[0].tables[0].columns.map((column) => column.name); // ["Region", "Units", "Price"]
```

### applyTableOperation

```ts
function applyTableOperation(
  workbook: Workbook,
  sheetId: string,
  tableId: string,
  request: TableOperation,
  options?: { readOnly?: boolean },
): TableOperationPlan;
```

Previews one change to a table and returns a plan. The kinds are `appendRows` (with `rows` or `count`, not both), `resize` (with `range`), `updateTable` (with `options`), `setColumnFormula` (with `columnId` and `formula`) and `clearColumnFormula`. The plan's `next` holds the workbook the change would produce; nothing is applied until the plan goes to `applyTableOperationPreview`. It throws for an unknown table, column or operation, or for values the operation cannot accept.

```js
import {
  applyTableOperation,
  applyTableOperationPreview,
} from "@ledgerglass/workbook";

const tableId = withTable.sheets[0].tables[0].id;
const plan = applyTableOperation(withTable, sheet.id, tableId, {
  kind: "appendRows",
  rows: [["East", 5, 2]],
});
const appended = applyTableOperationPreview(withTable, plan);
appended.sheets[0].tables[0].range; // "A1:C3"
```

### applyTableOperationPreview

```ts
function applyTableOperationPreview(
  workbook: Workbook,
  plan: TableOperationPlan,
  options?: { readOnly?: boolean },
): Workbook;
```

Returns the workbook that a previewed table operation produces. If the operation would change nothing, it returns the `workbook` object it was given. It throws if the workbook is not the one the plan was made from, if the plan no longer matches a fresh preview, or if the workbook is read-only.

### findTableForRange

```ts
function findTableForRange(
  workbook: Workbook,
  sheetId: string,
  range: string,
): string;
```

Returns the A1 range of the table that holds a single cell. If no table holds it, it returns the used range of the sheet. A range of more than one cell is returned as A1 text, and it throws if the sheet does not exist.

```js
import { findTableForRange } from "@ledgerglass/workbook";

findTableForRange(withTable, sheet.id, "B2"); // "A1:C2"
```

### suggestTableHeaderName

```ts
function suggestTableHeaderName(
  workbook: Workbook,
  sheetId: string,
  range: string,
): TableHeaderSuggestion[];
```

Suggests one column name for each cell in the first row of `range`, and does not change the workbook. A blank header becomes `ColumnN`, and a repeated name gets a number. A suggestion that differs from the cell's text, or a header that is a formula, is marked `requiresResolution`.

```js
import { suggestTableHeaderName } from "@ledgerglass/workbook";

suggestTableHeaderName(book, sheet.id, "A1:C1").map((s) => s.suggested); // ["Region", "Units", "Price"]
```

### setTableColumnDataType

```ts
function setTableColumnDataType(
  workbook: Workbook,
  sheetId: string,
  tableId: string,
  columnId: string,
): TableColumnTypeCheck;
```

Checks the values in a column against the column's current type and lists the ones that do not match. Despite the name, it sets nothing and does not change the workbook. A column of type `any` has no exceptions, and at most 500 data rows are checked. It throws if the table or column is not found.

```js
import { setTableColumnDataType } from "@ledgerglass/workbook";

const check = setTableColumnDataType(
  withTable,
  sheet.id,
  tableId,
  withTable.sheets[0].tables[0].columns[1].id,
);
check.dataType; // "any"
check.exceptions; // []
```

### previewColumnTypeChange

```ts
function previewColumnTypeChange(
  workbook: Workbook,
  sheetId: string,
  tableId: string,
  request: ColumnTypeRequest,
  options?: { readOnly?: boolean },
): ColumnTypePlan;
```

Plans a change of a column's type and returns the plan without changing the workbook. With `mode: "convert"`, values that parse as the new type are rewritten; for numbers, `decimal` and `group` say how the source text is read. With `mode: "tag"`, the type is set and values are only checked. A value that does not convert is an exception: with `invalid: "reject"` it blocks the change (`canApply` is false), and with `"keep"` it stays as it is. It throws for an invalid request, an unknown table or column, a protected sheet, or a table with more than 500 data rows.

```js
import { previewColumnTypeChange } from "@ledgerglass/workbook";

const tableColumns = withTable.sheets[0].tables[0].columns;
const plan = previewColumnTypeChange(withTable, sheet.id, tableId, {
  columnId: tableColumns[1].id,
  dataType: "number",
  mode: "convert",
  invalid: "reject",
});
plan.canApply; // true when no value was an exception
```

### applyColumnTypePreview

```ts
function applyColumnTypePreview(
  workbook: Workbook,
  plan: ColumnTypePlan,
  options?: { readOnly?: boolean },
): Workbook;
```

Returns the workbook that a column type plan produces. It throws if the workbook has changed since the preview, if the plan is stale, or if `canApply` is false. Resolve the listed exceptions, or use `invalid: "keep"`, and preview again.

```js
import { applyColumnTypePreview } from "@ledgerglass/workbook";

const typed = applyColumnTypePreview(withTable, plan);
typed.sheets[0].tables[0].columns[1].dataType; // "number"
```

### tableColumnDataTypes

```ts
const tableColumnDataTypes: readonly TableColumnDataType[];
```

The five column types: `"any"`, `"text"`, `"number"`, `"boolean"` and `"date"`. `"any"` means no type is set, and applying it removes the column's type. `previewColumnTypeChange` refuses any other value.

```js
import { tableColumnDataTypes } from "@ledgerglass/workbook";

tableColumnDataTypes.includes("number"); // true
```

## Filters

### applyFilterConditions

```ts
function applyFilterConditions(
  workbook: Workbook,
  sheetId: string,
  filterRequest: FilterRequest,
): FilterResult;
```

Applies a filter to the rows below a header row and returns which rows match. The workbook is not changed; hiding `hiddenRows` is up to the caller. Rows are reported as zero-based intervals. It throws if the range has no data rows, a column is outside the range or is repeated, or a condition is invalid.

```js
import { applyFilterConditions } from "@ledgerglass/workbook";

const result = applyFilterConditions(withTable, sheet.id, {
  range: "A1:C2",
  columns: [{ column: 0, values: ["North"] }],
});
result.context.matchedRows; // 1
```

### buildFilterFacetValues

```ts
function buildFilterFacetValues(
  workbook: Workbook,
  sheetId: string,
  filterRequest: FilterRequest,
  facetColumn: number,
  searchText?: string,
): FilterFacetValues;
```

Lists the distinct values of one column among the rows that the other filters keep, with a count of rows for each. This is the list a filter menu shows. `searchText` narrows the list and matches without regard to case. It throws if `facetColumn` is outside the filter range.

```js
import { buildFilterFacetValues } from "@ledgerglass/workbook";

buildFilterFacetValues(withTable, sheet.id, { range: "A1:C2", columns: [] }, 0)
  .values;
// [{ value: "North", count: 1, key: "string:north" }]
```

### createSavedFilterView

```ts
function createSavedFilterView(
  workbook: Workbook,
  sheetId: string,
  definition: FilterRequest,
  viewName: string,
  viewId?: string,
): Workbook;
```

Saves a named filter in `workbook.savedFilterViews` and returns the new workbook. The view records the row counts at the time it was saved. Pass `viewId` to replace an existing view, which keeps its id. It throws if the name is not 1 to 80 characters or the filter is invalid.

```js
import { createSavedFilterView } from "@ledgerglass/workbook";

const saved = createSavedFilterView(
  withTable,
  sheet.id,
  { range: "A1:C2", columns: [{ column: 0, values: ["North"] }] },
  "North only",
);
saved.savedFilterViews[0].population.matchedRows; // 1
```

### addSavedFilterView

```ts
function addSavedFilterView(
  workbook: Workbook,
  viewId: string,
  newName: string,
): Workbook;
```

Renames a saved view and returns the new workbook. It throws if no view has that id, or if the name is blank or longer than 80 characters.

### removeSavedFilterView

```ts
function removeSavedFilterView(workbook: Workbook, viewId: string): Workbook;
```

Deletes a saved view and returns the new workbook. It throws if no view has that id.

## Structured references

### rewriteStructuredReferences

```ts
function rewriteStructuredReferences(
  workbook: Workbook,
  sheetId: string,
  referenceText: string,
  address?: string,
): { sheetId: string; range: string } | null;
```

Resolves a table reference to the sheet and range it covers, without changing the workbook. It accepts `Table[Column]`, `Table[[First]:[Last]]`, `Table[#All]`, `Table[#Headers]`, `Table[#Totals]` and `Table[#This Row]`. `[@Column]` needs `address`, and resolves only inside a table's data rows. It returns null when the reference does not resolve.

```js
import { rewriteStructuredReferences } from "@ledgerglass/workbook";

rewriteStructuredReferences(withTable, sheet.id, "Sales[Units]"); // { sheetId, range: "B2:B2" }
```

## Pivot tables

### getPivotSourceFields

```ts
function getPivotSourceFields(
  workbook: Workbook,
  sheetId: string,
  range: string,
): string[];
```

Returns the header texts of a source range in column order, so a host can offer them as row, column and value fields. It does not change the workbook, and it throws if the sheet does not exist.

```js
import { getPivotSourceFields } from "@ledgerglass/workbook";

getPivotSourceFields(withTable, sheet.id, "A1:C2"); // ["Region", "Units", "Price"]
```

### buildPivotTableOutput

```ts
function buildPivotTableOutput(
  workbook: Workbook,
  request: PivotRequest,
): Workbook;
```

Builds a pivot from a source range and returns a new workbook. Without `destination` the output goes to a new sheet named after the pivot; with one, it is written at that address. The output sheet keeps the pivot's settings and a refresh record, which `getPivotRefreshState` checks later. It throws for an invalid request or for a field that is not in the source.

```js
import { buildPivotTableOutput } from "@ledgerglass/workbook";

const withPivot = buildPivotTableOutput(withTable, {
  name: "By region",
  source: { sheetId: sheet.id, range: "A1:C2" },
  rowFields: ["Region"],
  columnFields: [],
  values: [{ field: "Units", operation: "sum" }],
});
withPivot.sheets[1].pivots[0].refresh.measures[0].result; // 10
```

### refreshPivotTable

```ts
function refreshPivotTable(
  workbook: Workbook,
  pivotId: string,
  destination?: { sheetId?: string; sheetName?: string; address?: string },
): Workbook;
```

Rebuilds a pivot from its current source. Without `destination` the result goes to a new sheet named "<name> refreshed", and the earlier output is left as it was. With `destination`, the output is written at that address. It throws if no pivot has that id.

### getPivotRefreshState

```ts
function getPivotRefreshState(
  workbook: Workbook,
  pivotId: string,
): PivotRefreshState;
```

Reports whether a pivot still matches its source. `state` is `"current"` when the source and output match the refresh record, `"stale"` when either has changed, `"unavailable"` when a field can no longer be found, and `"unverified"` when there is no refresh record. `outputModified` says whether the output cells were edited. It throws if no pivot has that id.

```js
import { getPivotRefreshState } from "@ledgerglass/workbook";

const pivotId = withPivot.sheets[1].pivots[0].id;
getPivotRefreshState(withPivot, pivotId).state; // "current"
```

### buildPivotDrillDown

```ts
function buildPivotDrillDown(
  workbook: Workbook,
  pivotId: string,
  cellAddress: string,
): PivotDrillDown;
```

Lists the source rows behind one value cell, the grand total, or a subtotal of the pivot. Each record gives the source row's locator and values. It throws if the cell is a header or group label, or if no pivot has that id.

```js
import { buildPivotDrillDown } from "@ledgerglass/workbook";

const drill = buildPivotDrillDown(withPivot, pivotId, "B2");
drill.records.length; // 1
```

### extractPivotDrillThrough

```ts
function extractPivotDrillThrough(
  workbook: Workbook,
  pivotId: string,
  cellAddress: string,
  previewContext: { snapshot: string; readOnly?: boolean },
): Workbook;
```

Writes the rows behind a pivot value cell to a new sheet named "<name> details", one row per source row with a locator. Pass `snapshot: JSON.stringify(workbook)` taken when the drill-through was previewed. It throws if the workbook has changed, is read-only, or the cell cannot be drilled.

```js
import { extractPivotDrillThrough } from "@ledgerglass/workbook";

const extracted = extractPivotDrillThrough(withPivot, pivotId, "B2", {
  snapshot: JSON.stringify(withPivot),
});
extracted.sheets.at(-1).name; // "By region details"
```

### getPivotSlicerItems

```ts
function getPivotSlicerItems(
  workbook: Workbook,
  pivotId: string,
  field: string,
): RawCellContent[];
```

Returns the distinct items of a source field, in the order they first appear, for building a slicer. It throws if the field is not in the pivot's source.

```js
import { getPivotSlicerItems } from "@ledgerglass/workbook";

getPivotSlicerItems(withPivot, pivotId, "Region"); // ["North"]
```

### applyPivotViewChange

```ts
function applyPivotViewChange(
  workbook: Workbook,
  pivotId: string,
  viewChanges: PivotViewChanges,
  previewContext: { snapshot: string; readOnly?: boolean },
): Workbook;
```

Builds a changed copy of a pivot on a new sheet named "<name> filtered". The changes are `slicers` (keep only the listed items), `sorts`, `values` (replace the aggregations) and `chart`. The original pivot is not changed. It throws if the workbook has changed or is read-only, or if the changes include a key that is not listed.

```js
import { applyPivotViewChange } from "@ledgerglass/workbook";

const sorted = applyPivotViewChange(
  withPivot,
  pivotId,
  {
    sorts: [{ field: "Region", direction: "descending" }],
  },
  { snapshot: JSON.stringify(withPivot) },
);
```

### buildPivotChartModel

```ts
function buildPivotChartModel(book: Workbook, pivotId: string): PivotChartModel;
```

Returns the data behind a pivot chart: a chart kind (`"column"`, `"bar"` or `"line"`), the measure name, and one point for each row group. It charts the first value field unless the pivot's chart setting names another. It draws nothing; pass the points to a chart library. It throws if no pivot has that id.

```js
import { buildPivotChartModel } from "@ledgerglass/workbook";

buildPivotChartModel(withPivot, pivotId).points; // [{ label: "North", value: 10 }]
```

### pivotAggregateOperations

```ts
const pivotAggregateOperations: readonly PivotAggregateOperation[];
```

The aggregations a pivot value field may use: `"sum"`, `"count"`, `"average"`, `"min"`, `"max"` and `"distinctCount"`. `buildPivotTableOutput` refuses any other operation.

```js
import { pivotAggregateOperations } from "@ledgerglass/workbook";

pivotAggregateOperations.includes("median"); // false
```
