/** Worksheet tables: create them, change their rows and formulas, and check or convert column types. */
export {
  createWorksheetTable,
  applyTableOperation,
  applyTableOperationPreview,
  findTableForRange,
  suggestTableHeaderName,
  setTableColumnDataType,
  previewColumnTypeChange,
  applyColumnTypePreview,
  tableColumnDataTypes,
} from "../app/03-workbook-model.js";

/** Filters: test rows against conditions, list the values of a column, and save named filter views. */
export {
  applyFilterConditions,
  buildFilterFacetValues,
  createSavedFilterView,
  addSavedFilterView,
  removeSavedFilterView,
} from "../app/03-workbook-model.js";

/** Structured references: resolve a reference such as `Sales[Units]` to a range. */
export { rewriteStructuredReferences } from "../app/03-workbook-model.js";

/** Pivot tables: build, refresh, check, drill into, filter and chart pivot results. */
export {
  getPivotSourceFields,
  buildPivotTableOutput,
  refreshPivotTable,
  getPivotRefreshState,
  buildPivotDrillDown,
  extractPivotDrillThrough,
  getPivotSlicerItems,
  applyPivotViewChange,
  buildPivotChartModel,
  pivotAggregateOperations,
} from "../app/03-workbook-model.js";
