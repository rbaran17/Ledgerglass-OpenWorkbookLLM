import i from "react/jsx-runtime";
import { WorkbookRibbon } from "../26-workbook-ribbon.js";
import { _t, applyAutoSum, createBlankWorkbook } from "../03-workbook-model.js";
import { buildSampleLedgerWorkbook } from "../24-merge-audit-and-ribbon-helpers.js";
import {
  applyCellPreview,
  buildFillPreview,
} from "../04-cell-input-paste-and-fill.js";

function EditorRibbonBar({
  selectionState,
  dialogState,
  filterSortState,
  viewDisplayState,
  operationStatusState,
  findReplaceState,
  importExportState,
  readOnly,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  requireSingleSelectionRange,
  scrollActivePaneVertically,
  undoStack,
  redoStack,
  hasFrozenRows,
  activeCell,
  clearSelectedCellContents,
  runEditorCommand,
  commitWorkbookEdit,
  runGuardedAction,
  applyCellPatches,
  applyZoomPercent,
  applyCellFormatPatch,
  undoOrRedo,
  exportWorkbookAs,
  replaceActiveWorkbook,
  toggleFreezeTopRow,
  copyOrCutSelection,
  pasteFromClipboard,
  autoFitSelectedAxis,
}) {
  return i.jsx(WorkbookRibbon, {
    onTool: runEditorCommand,
    sheetProtected: !!activeSheet.protection?.enabled,
    allowProtectionChange: !readOnly,
    workbookStructureProtected: !!activeWorkbook.structureProtection?.enabled,
    readOnly: isEditingLocked,
    busy: operationStatusState.isOperationBusy,
    canUndo: undoStack.length > 0,
    canRedo: redoStack.length > 0,
    cell: activeCell,
    showFormulas: viewDisplayState.isShowingFormulas,
    gridlines: viewDisplayState.isGridlinesVisible,
    headings: viewDisplayState.isHeadingsVisible,
    freezeRows: hasFrozenRows,
    zoom: viewDisplayState.zoomPercent,
    onFormat: applyCellFormatPatch,
    onUndo: () => undoOrRedo(),
    onRedo: () => undoOrRedo(true),
    onImport: () => importExportState.importFileInputRef.current?.click(),
    onExport: (fe) => void exportWorkbookAs(fe),
    onNew: () => replaceActiveWorkbook(createBlankWorkbook()),
    onDemo: () => replaceActiveWorkbook(buildSampleLedgerWorkbook()),
    onRename: () =>
      dialogState.setPromptDialogState({
        kind: "rename",
        value: activeWorkbook.name,
      }),
    onAddSheet: () =>
      dialogState.setPromptDialogState({
        kind: "sheet",
        value: `Sheet${activeWorkbook.sheets.length + 1}`,
      }),
    onAddRow: () => runEditorCommand("insertRow"),
    onAddColumn: () => runEditorCommand("insertColumn"),
    onAutoFitColumns: () => autoFitSelectedAxis("column"),
    onAutoFitRows: () => autoFitSelectedAxis("row"),
    onAutosum: () =>
      runGuardedAction(() =>
        applyCellPatches(
          [
            applyAutoSum(
              activeWorkbook,
              activeSheet.id,
              selectionState.activeCellAddress,
            ),
          ],
          "AutoSum",
        ),
      ),
    onFillDown: () =>
      runGuardedAction(() => {
        (requireSingleSelectionRange(),
          commitWorkbookEdit(
            applyCellPreview(
              activeWorkbook,
              buildFillPreview(
                activeWorkbook,
                activeSheet.id,
                selectionState.activeSelectionRange,
                "down",
              ),
            ),
            "Fill down",
          ));
      }),
    onClear: () => runGuardedAction(clearSelectedCellContents),
    onCopy: () => void copyOrCutSelection(false),
    onCut: () => void copyOrCutSelection(true),
    onPaste: () => void pasteFromClipboard(),
    onFind: () => findReplaceState.setIsFindBarVisible((fe) => !fe),
    onSort: (fe) => {
      (filterSortState.setActiveColumnSort(
        fe
          ? {
              col: _t(selectionState.activeCellAddress).col,
              direction: fe === "asc" ? 1 : -1,
            }
          : null,
      ),
        scrollActivePaneVertically(0),
        selectionState.gridContainerRef.current &&
          (selectionState.gridContainerRef.current.scrollTop = 0));
    },
    onToggleFilter: () => {
      (filterSortState.setIsFilterBarVisible((fe) => !fe),
        filterSortState.setFilterValueSearch(""));
    },
    onShowFormulas: () => viewDisplayState.setIsShowingFormulas((fe) => !fe),
    onGridlines: () => viewDisplayState.setIsGridlinesVisible((fe) => !fe),
    onHeadings: () => viewDisplayState.setIsHeadingsVisible((fe) => !fe),
    onFreeze: toggleFreezeTopRow,
    onZoom: applyZoomPercent,
    onPanel: dialogState.setSidePanelKind,
    onNote: () =>
      dialogState.setPromptDialogState({
        kind: "note",
        value: activeCell?.note ?? "",
      }),
    onReviewed: () =>
      applyCellFormatPatch({
        tickmark: activeCell?.tickmark ? "" : "✓",
      }),
  });
}

export { EditorRibbonBar };
