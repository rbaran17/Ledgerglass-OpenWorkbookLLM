import type { Workbook, RawCellContent } from "../index";

/** The column types a table column can carry. "any" means no type is set. */
export type TableColumnDataType =
  "any" | "text" | "number" | "boolean" | "date";

/** One column of a table, as stored in `sheet.tables[i].columns`. */
export interface TableColumn {
  id: string;
  name: string;
  /** The formula of a calculated column, beginning with "=". Absent on other columns. */
  calculatedFormula?: string;
  /** The column's type. Absent when the type is "any". */
  dataType?: Exclude<TableColumnDataType, "any">;
  [extra: string]: unknown;
}

/** A table, as stored in `sheet.tables`. */
export interface WorksheetTable {
  id: string;
  name: string;
  /** The table's full range in A1 notation, including the header row and any totals row. */
  range: string;
  columns: TableColumn[];
  style: string;
  showTotals: boolean;
  totals: Record<string, unknown>;
  [extra: string]: unknown;
}

export interface CreateWorksheetTableOptions {
  /** 1 to 255 characters: a letter or underscore first, then letters, digits, underscores or dots. */
  name: string;
  /** A1 range that includes the header row, such as "A1:C10". */
  range: string;
  /** One reviewed name for each column. Required when the first row has blank, duplicate or formula headers. */
  headers?: string[];
  /** A table style name. Defaults to "blue". */
  style?: string;
  showTotals?: boolean;
  totals?: Record<string, unknown>;
  [extra: string]: unknown;
}

/** An operation on an existing table. Each kind accepts only the fields listed for it. */
export type TableOperation =
  | { kind: "setColumnFormula"; columnId: string; formula: string }
  | { kind: "clearColumnFormula"; columnId: string }
  | { kind: "appendRows"; rows: RawCellContent[][] }
  | { kind: "appendRows"; count: number }
  | { kind: "resize"; range: string }
  | {
      kind: "updateTable";
      options: {
        name: string;
        range: string;
        style?: string;
        showTotals?: boolean;
        totals?: Record<string, unknown>;
        [extra: string]: unknown;
      };
    };

/** A previewed table operation. Pass it to `applyTableOperationPreview` to get the changed workbook. */
export interface TableOperationPlan {
  /** The JSON of the workbook the preview was made from. Checked again on apply. */
  snapshot: string;
  sheetId: string;
  tableId: string;
  request: TableOperation;
  beforeRange: string;
  afterRange: string;
  changes: unknown[];
  anchorChanges: unknown[];
  /** The workbook with the operation applied. */
  next: Workbook;
  [extra: string]: unknown;
}

/** A header suggestion for one column, from `suggestTableHeaderName`. */
export interface TableHeaderSuggestion {
  /** The header cell's address, such as "B1". */
  address: string;
  /** The cell's raw value, or null when the cell is empty. */
  original: RawCellContent;
  suggested: string;
  /** True when the suggestion differs from the cell's text, or the cell holds a formula that is not a literal. */
  requiresResolution: boolean;
}

/** The result of checking a column's values, from `setTableColumnDataType`. */
export interface TableColumnTypeCheck {
  /** The type the values were checked against: the column's current type. */
  dataType: TableColumnDataType;
  /** How many data rows were checked, at most 500. */
  checked: number;
  /** How many data rows the table has. */
  total: number;
  truncated: boolean;
  exceptions: Array<{ address: string; value: unknown }>;
  [extra: string]: unknown;
}

/** A request to convert a column's values to a type. */
export interface ColumnTypeRequest {
  columnId: string;
  dataType: TableColumnDataType;
  /** "convert" rewrites values; "tag" only sets the column type and checks values. */
  mode: "convert" | "tag";
  /** "reject" blocks the change while exceptions remain; "keep" leaves those values unchanged. */
  invalid: "reject" | "keep";
  /** The decimal separator in source text. Defaults to ".". */
  decimal?: "." | ",";
  /** The group (thousands) separator in source text. Defaults to ",". Use "" for none. */
  group?: "" | "." | "," | " ";
}

/** A previewed column type change. Pass it to `applyColumnTypePreview` to get the changed workbook. */
export interface ColumnTypePlan {
  snapshot: string;
  sheetId: string;
  tableId: string;
  request: ColumnTypeRequest;
  beforeType: TableColumnDataType;
  /** False while exceptions remain under "reject". */
  canApply: boolean;
  changes: Array<{ address: string; before?: unknown; after: unknown }>;
  exceptions: Array<{ address: string; value: RawCellContent; reason: string }>;
  /** The workbook with the type change applied. */
  next: Workbook;
  [extra: string]: unknown;
}

/** The ways a condition can compare a cell with its criterion. */
export type FilterOperator =
  | "equals"
  | "notEquals"
  | "contains"
  | "notContains"
  | "startsWith"
  | "endsWith"
  | "wildcard"
  | "gt"
  | "gte"
  | "lt"
  | "lte"
  | "blank"
  | "notBlank";

/** One comparison on one column. The column index is absolute: 0 is column A. */
export interface FilterCondition {
  column: number;
  type: "value" | "text" | "number" | "date";
  operator: FilterOperator;
  /** Not used by "blank" and "notBlank". Text criteria need a string. */
  value?: RawCellContent;
  [extra: string]: unknown;
}

/** The filter on one column. It may list values, give up to two conditions, or both. */
export interface FilterColumn {
  column: number;
  /** Keep only rows whose value is in this list. */
  values?: RawCellContent[];
  /** Up to two conditions on this column. */
  conditions?: FilterCondition[];
  /** How this column's values and conditions combine: "and" (the default) or "or". */
  join?: "and" | "or";
  [extra: string]: unknown;
}

/** A filter over a header row and the rows below it. Columns that are not listed are not filtered. */
export interface FilterRequest {
  /** A1 range whose first row is the header, such as "A1:C10". */
  range: string;
  columns: FilterColumn[];
  /** Groups joined by OR. Each group's conditions are joined by AND. */
  advanced?: FilterCondition[][];
  /** Compare text case-sensitively. Defaults to false. */
  caseSensitive?: boolean;
  /** Keep only the first row of each set of identical rows. */
  unique?: boolean;
  [extra: string]: unknown;
}

/** A run of consecutive rows, with zero-based indexes: `{ start: 2, end: 2 }` is the third row. */
export interface TableRowInterval {
  start: number;
  end: number;
}

/** The outcome of applying a filter. */
export interface FilterResult {
  /** Rows to hide, as zero-based intervals. */
  hiddenRows: TableRowInterval[];
  /** Rows that did not match, as zero-based intervals. */
  excludedRows: TableRowInterval[];
  fields: Array<{ column: number; label: string }>;
  context: {
    workbookId: string;
    revision: number;
    sheetId: string;
    sheetName: string;
    range: string;
    definition: FilterRequest;
    totalRows: number;
    matchedRows: number;
    visibleRows: number;
    filteredRows: number;
    manuallyHiddenRows: number;
    hiddenMatchingRows: number;
    [extra: string]: unknown;
  };
}

/** The distinct values of one column, with how many matching rows hold each. */
export interface FilterFacetValues {
  values: Array<{ value: RawCellContent; count: number; key: string }>;
  distinct: number;
  truncated: boolean;
}

/** A named, saved filter. It is stored in `workbook.savedFilterViews`. */
export interface SavedFilterView {
  id: string;
  name: string;
  sheetId: string;
  sheetName: string;
  sourceWorkbookId: string;
  sourceRevision: number;
  savedAt: string;
  definition: FilterRequest;
  /** Row counts recorded when the view was saved. */
  population: {
    totalRows: number;
    matchedRows: number;
    visibleRows: number;
    hiddenMatchingRows: number;
  };
  [extra: string]: unknown;
}

/** A column of the source, or a pivot field, referenced by its header text. */
export type PivotAggregateOperation =
  "sum" | "count" | "average" | "min" | "max" | "distinctCount";

export interface PivotValueField {
  /** The header text of a source column. */
  field: string;
  operation: PivotAggregateOperation;
  /** The display name in the output. Defaults to the operation and field, such as "Sum of Units". */
  name?: string;
  /** "value" shows the amount. The percent modes show each value as a share of a total. */
  showAs?: "value" | "percentGrand" | "percentRow" | "percentColumn";
  [extra: string]: unknown;
}

/** A source filter applied before aggregation. */
export interface PivotSourceFilter {
  field: string;
  operator: FilterOperator;
  value?: RawCellContent;
  [extra: string]: unknown;
}

export interface PivotSource {
  sheetId: string;
  /** A1 range whose first row is the header. */
  range: string;
  [extra: string]: unknown;
}

/** The settings of a pivot table. The output is written to a new sheet unless `destination` is given. */
export interface PivotRequest {
  /** 1 to 100 characters. */
  name: string;
  source: PivotSource;
  rowFields: string[];
  columnFields: string[];
  values: PivotValueField[];
  filters?: PivotSourceFilter[];
  /** Show a grand total row and column. Defaults to true. */
  grandTotals?: boolean;
  subtotals?: boolean;
  /** Where to write the output. Omit it to create a new sheet. */
  destination?: { sheetId?: string; sheetName?: string; address?: string };
  [extra: string]: unknown;
}

/** A pivot as stored on its output sheet, in `sheet.pivots`. */
export interface PivotDefinition extends PivotRequest {
  id: string;
  outputRange: string;
  sourceRevision: number;
  /** The fingerprints and figures taken when the pivot was last built or refreshed. */
  refresh?: {
    at: string;
    sourceSheetName: string;
    sourceRange: string;
    sourceFingerprint: string;
    outputFingerprint: string;
    sourceRows: number;
    includedRows: number;
    excludedRows: number;
    measures: Array<{
      name: string;
      field: string;
      operation: PivotAggregateOperation;
      result: unknown;
      numericCount: number;
      nonblankCount: number;
    }>;
    [extra: string]: unknown;
  };
  [extra: string]: unknown;
}

/** Whether a pivot's output still matches its source. */
export interface PivotRefreshState {
  /** "current" when source and output match; "stale" when either changed; "unverified" when there is no refresh record; "unavailable" when a field can no longer be found. */
  state: "current" | "stale" | "unverified" | "unavailable";
  reason: string;
  outputModified: boolean;
  definition: PivotDefinition;
}

/** A pivot value cell broken down to its source rows. */
export interface PivotDrillDown {
  context: Record<string, unknown>;
  address: string;
  valueIndex: number;
  measure: { name: string; [extra: string]: unknown };
  underlyingValue: unknown;
  displayedValue: unknown;
  rowGroup: unknown;
  columnGroup: unknown;
  /** The source column headers. */
  fields: string[];
  records: Array<{
    locator: {
      sheetName: string;
      row: number;
      range: string;
      [extra: string]: unknown;
    };
    values: RawCellContent[];
    cells: Array<{ address: string; cell: unknown }>;
  }>;
  [extra: string]: unknown;
}

/** Changes to a pivot's view. Each change builds a new output sheet; the original pivot is not changed. */
export interface PivotViewChanges {
  /** Keep only the listed items of a field. */
  slicers?: Array<{ field: string; selected: RawCellContent[] }>;
  /** Sort the row labels of a field. */
  sorts?: Array<{ field: string; direction: "ascending" | "descending" }>;
  /** Replace the value fields and their Show values as settings. */
  values?: PivotValueField[];
  /** Chart settings. `valueIndex` is the position in the value fields. */
  chart?: { kind: "column" | "bar" | "line"; valueIndex: number };
  /** Date timelines. Accepted, but not described here: they need an explicit period and date system. */
  timelines?: Array<Record<string, unknown>>;
}

/** The data behind a pivot chart. */
export interface PivotChartModel {
  kind: "column" | "bar" | "line";
  measure: string;
  showAs: "value" | "percentGrand" | "percentRow" | "percentColumn";
  points: Array<{ label: string; value: number | null }>;
  sourceRevision: number;
}

/** Guard for the preview functions that refuse a change when the workbook differs from when they were previewed. */
export interface PreviewContext {
  /** The JSON of the workbook the preview was made from. */
  snapshot: string;
  readOnly?: boolean;
}

/** Column types a table column may carry. Use these values for `dataType`. */
export const tableColumnDataTypes: readonly TableColumnDataType[];

/** The aggregations a pivot value field may use. */
export const pivotAggregateOperations: readonly PivotAggregateOperation[];

/**
 * Adds a table to a sheet and returns a new workbook. The argument is not changed. Throws if the sheet is
 * protected or merged cells overlap the range, the name is taken or invalid, or the headers need review.
 */
export function createWorksheetTable(
  workbook: Workbook,
  sheetId: string,
  options: CreateWorksheetTableOptions,
): Workbook;

/**
 * Previews a change to a table: appending rows, resizing, or setting or clearing a calculated column.
 * Returns a plan, not a workbook; read `plan.next`, or pass the plan to `applyTableOperationPreview`.
 * Throws for an unknown table, column or operation, or for values the operation cannot accept.
 */
export function applyTableOperation(
  workbook: Workbook,
  sheetId: string,
  tableId: string,
  request: TableOperation,
  options?: { readOnly?: boolean },
): TableOperationPlan;

/**
 * Returns the workbook a previewed table operation produces. Returns the same object when nothing changes.
 * Throws if the workbook differs from the one previewed, or if the plan no longer matches a fresh preview.
 */
export function applyTableOperationPreview(
  workbook: Workbook,
  plan: TableOperationPlan,
  options?: { readOnly?: boolean },
): Workbook;

/**
 * Returns the A1 range of the table that contains a single cell, or of the sheet's occupied cells when no
 * table does. A multi-cell range is returned in A1 form. Throws if the sheet does not exist.
 */
export function findTableForRange(
  workbook: Workbook,
  sheetId: string,
  range: string,
): string;

/**
 * Suggests a name for each column from the header row of a range. Does not change the workbook.
 * A suggestion needs review when the header is blank, repeated, not text, or a formula.
 */
export function suggestTableHeaderName(
  workbook: Workbook,
  sheetId: string,
  range: string,
): TableHeaderSuggestion[];

/**
 * Counts the values in a table column that do not match the column's current type. It does not change the
 * workbook. A column with type "any" has no exceptions. Throws if the table or column is not found.
 */
export function setTableColumnDataType(
  workbook: Workbook,
  sheetId: string,
  tableId: string,
  columnId: string,
): TableColumnTypeCheck;

/**
 * Previews a change of column type. Returns a plan with the changed values, the exceptions, and `canApply`.
 * Does not change the workbook. Throws for an invalid request, an unknown table or column, a protected sheet, or a
 * table with more than 500 data rows.
 */
export function previewColumnTypeChange(
  workbook: Workbook,
  sheetId: string,
  tableId: string,
  request: ColumnTypeRequest,
  options?: { readOnly?: boolean },
): ColumnTypePlan;

/**
 * Returns the workbook a previewed column type change produces. Throws if the workbook has changed, if the
 * preview is stale, or if `canApply` is false.
 */
export function applyColumnTypePreview(
  workbook: Workbook,
  plan: ColumnTypePlan,
  options?: { readOnly?: boolean },
): Workbook;

/**
 * Applies a filter to the rows below a header and returns the outcome. It does not change the workbook; apply
 * `hiddenRows` yourself if you want the rows hidden. Throws for an invalid range, column or condition.
 */
export function applyFilterConditions(
  workbook: Workbook,
  sheetId: string,
  filterRequest: FilterRequest,
): FilterResult;

/**
 * Lists the distinct values of one column among the rows that the other filters keep. Values are matched
 * case-insensitively against `searchText`. Throws if the column is outside the filter range.
 */
export function buildFilterFacetValues(
  workbook: Workbook,
  sheetId: string,
  filterRequest: FilterRequest,
  facetColumn: number,
  searchText?: string,
): FilterFacetValues;

/**
 * Saves a named filter view and returns a new workbook. Pass `viewId` to replace an existing view. Throws if the
 * name is not 1 to 80 characters or the filter is invalid.
 */
export function createSavedFilterView(
  workbook: Workbook,
  sheetId: string,
  definition: FilterRequest,
  viewName: string,
  viewId?: string,
): Workbook;

/**
 * Renames a saved filter view and returns a new workbook. Throws if the view does not exist or the name is not 1
 * to 80 characters.
 */
export function addSavedFilterView(
  workbook: Workbook,
  viewId: string,
  newName: string,
): Workbook;

/**
 * Deletes a saved filter view and returns a new workbook. Throws if the view does not exist.
 */
export function removeSavedFilterView(
  workbook: Workbook,
  viewId: string,
): Workbook;

/**
 * Resolves a structured reference such as `Sales[Units]`, `Sales[[Units]:[Price]]` or `[@Units]` to a range.
 * Returns null when the reference does not resolve. Pass `address` for `[@Column]` and `#This Row` references.
 */
export function rewriteStructuredReferences(
  workbook: Workbook,
  sheetId: string,
  referenceText: string,
  address?: string,
): { sheetId: string; range: string } | null;

/**
 * Returns the header texts of a source range, in column order. Throws if the sheet does not exist.
 */
export function getPivotSourceFields(
  workbook: Workbook,
  sheetId: string,
  range: string,
): string[];

/**
 * Builds a pivot table and returns a new workbook. The output goes to a new sheet unless `destination` is given.
 * Throws for an invalid request or for a field that is not in the source.
 */
export function buildPivotTableOutput(
  workbook: Workbook,
  request: PivotRequest,
): Workbook;

/**
 * Rebuilds an existing pivot from its source. Without `destination` the result goes to a new sheet; with it,
 * the output is written at that address. Throws if the pivot is not found.
 */
export function refreshPivotTable(
  workbook: Workbook,
  pivotId: string,
  destination?: { sheetId?: string; sheetName?: string; address?: string },
): Workbook;

/**
 * Reports whether a pivot's output still matches its source and whether the output was edited. Does not change
 * the workbook. Throws if the pivot is not found.
 */
export function getPivotRefreshState(
  workbook: Workbook,
  pivotId: string,
): PivotRefreshState;

/**
 * Lists the source rows behind one pivot value cell, the grand total or a subtotal. Throws if the cell is a
 * header or group label, or if the pivot is not found.
 */
export function buildPivotDrillDown(
  workbook: Workbook,
  pivotId: string,
  cellAddress: string,
): PivotDrillDown;

/**
 * Writes the source rows behind a pivot value cell to a new sheet and returns the new workbook. Throws if the
 * workbook has changed since `previewContext.snapshot`, if it is read-only, or if the cell cannot be drilled.
 */
export function extractPivotDrillThrough(
  workbook: Workbook,
  pivotId: string,
  cellAddress: string,
  previewContext: PreviewContext,
): Workbook;

/**
 * Returns the distinct items of a pivot field, for building a slicer. Throws if the field is not in the pivot's source.
 */
export function getPivotSlicerItems(
  workbook: Workbook,
  pivotId: string,
  field: string,
): RawCellContent[];

/**
 * Builds a filtered, sorted, re-valued or charted copy of a pivot on a new sheet and returns the new workbook.
 * The original pivot is not changed. Throws if the workbook has changed since `previewContext.snapshot`, or for a
 * change key other than slicers, timelines, sorts, values or chart.
 */
export function applyPivotViewChange(
  workbook: Workbook,
  pivotId: string,
  viewChanges: PivotViewChanges,
  previewContext: PreviewContext,
): Workbook;

/**
 * Returns the points of a pivot chart: one label and value for each row group of the value field that the pivot
 * chart setting names (the first value field by default). Throws if the pivot is not found.
 */
export function buildPivotChartModel(
  book: Workbook,
  pivotId: string,
): PivotChartModel;
