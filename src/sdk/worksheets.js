/** Worksheet creation, removal, ordering, visibility and tab colour. */
export {
  addWorksheet,
  renameWorksheet,
  moveWorksheet,
  duplicateWorksheet,
  deleteWorksheet,
  setWorksheetVisibility,
  setWorksheetTabColor,
} from "../app/03-workbook-model.js";

/** Workbook-level name, engagement and period, and the workbook template. */
export {
  renameWorkbook,
  updateWorkbookProperties,
  createWorkbookFromTemplate,
} from "../app/03-workbook-model.js";

/** Protection of the workbook structure and of individual worksheets. */
export {
  protectWorkbookStructure,
  unprotectWorkbookStructure,
  updateSheetProtection,
} from "../app/03-workbook-model.js";

/** Worksheet lookup, summaries and the impact of deleting a worksheet. */
export {
  summarizeWorksheets,
  describeSheetDeletionImpacts,
  findSheetOrThrow,
} from "../app/03-workbook-model.js";

/** Frozen and split panes, merged cells, hidden rows and columns, sizes and outline groups. */
export {
  setFreezePanes,
  setFreezePanesAtCell,
  setSplitPanes,
  mergeWorksheetCells,
  setHiddenAxisRange,
  readHiddenAxisIntervals,
  applySheetSizeEdits,
  removeOutlineGroup,
  setOutlineGroupCollapsed,
} from "../app/03-workbook-model.js";
