/**
 * The tables area of the SDK: the model functions in ./model/tables.js, with edits recorded
 * (see ./recording.js). Reference page: docs/api/tables.md.
 */
import * as model from "./model/tables.js";
import { recordEdits } from "./recording.js";

const area = recordEdits(model);

export const createWorksheetTable = area.createWorksheetTable;
export const applyTableOperation = area.applyTableOperation;
export const applyTableOperationPreview = area.applyTableOperationPreview;
export const findTableForRange = area.findTableForRange;
export const suggestTableHeaderName = area.suggestTableHeaderName;
export const checkTableColumnValues = area.checkTableColumnValues;
export const previewColumnTypeChange = area.previewColumnTypeChange;
export const applyColumnTypePreview = area.applyColumnTypePreview;
export const tableColumnDataTypes = area.tableColumnDataTypes;
export const applyFilterConditions = area.applyFilterConditions;
export const buildFilterFacetValues = area.buildFilterFacetValues;
export const createSavedFilterView = area.createSavedFilterView;
export const renameSavedFilterView = area.renameSavedFilterView;
export const removeSavedFilterView = area.removeSavedFilterView;
export const rewriteStructuredReferences = area.rewriteStructuredReferences;
export const getPivotSourceFields = area.getPivotSourceFields;
export const buildPivotTableOutput = area.buildPivotTableOutput;
export const refreshPivotTable = area.refreshPivotTable;
export const getPivotRefreshState = area.getPivotRefreshState;
export const buildPivotDrillDown = area.buildPivotDrillDown;
export const extractPivotDrillThrough = area.extractPivotDrillThrough;
export const getPivotSlicerItems = area.getPivotSlicerItems;
export const applyPivotViewChange = area.applyPivotViewChange;
export const buildPivotChartModel = area.buildPivotChartModel;
export const pivotAggregateOperations = area.pivotAggregateOperations;
