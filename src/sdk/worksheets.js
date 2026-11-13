/**
 * The worksheets area of the SDK: the model functions in ./model/worksheets.js, with edits recorded
 * (see ./recording.js). Reference page: docs/api/worksheets.md.
 */
import * as model from "./model/worksheets.js";
import { recordEdits } from "./recording.js";

const area = recordEdits(model);

export const addWorksheet = area.addWorksheet;
export const renameWorksheet = area.renameWorksheet;
export const moveWorksheet = area.moveWorksheet;
export const duplicateWorksheet = area.duplicateWorksheet;
export const deleteWorksheet = area.deleteWorksheet;
export const setWorksheetVisibility = area.setWorksheetVisibility;
export const setWorksheetTabColor = area.setWorksheetTabColor;
export const renameWorkbook = area.renameWorkbook;
export const updateWorkbookProperties = area.updateWorkbookProperties;
export const createWorkbookFromTemplate = area.createWorkbookFromTemplate;
export const protectWorkbookStructure = area.protectWorkbookStructure;
export const unprotectWorkbookStructure = area.unprotectWorkbookStructure;
export const updateSheetProtection = area.updateSheetProtection;
export const summarizeWorksheets = area.summarizeWorksheets;
export const describeSheetDeletionImpacts = area.describeSheetDeletionImpacts;
export const findSheetOrThrow = area.findSheetOrThrow;
export const setFreezePanes = area.setFreezePanes;
export const setFreezePanesAtCell = area.setFreezePanesAtCell;
export const setSplitPanes = area.setSplitPanes;
export const mergeWorksheetCells = area.mergeWorksheetCells;
export const setHiddenAxisRange = area.setHiddenAxisRange;
export const readHiddenAxisIntervals = area.readHiddenAxisIntervals;
export const applySheetSizeEdits = area.applySheetSizeEdits;
export const removeOutlineGroup = area.removeOutlineGroup;
export const setOutlineGroupCollapsed = area.setOutlineGroupCollapsed;
