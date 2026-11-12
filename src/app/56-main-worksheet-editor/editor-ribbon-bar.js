import i from "react/jsx-runtime";
import { WorkbookRibbon } from "../26-workbook-ribbon.js";
import { _t, applyAutoSum, createBlankWorkbook } from "../03-workbook-model.js";
import { buildSampleLedgerWorkbook } from "../24-merge-audit-and-ribbon-helpers.js";
import {
  applyCellPreview,
  buildFillPreview,
} from "../04-cell-input-paste-and-fill.js";

function EditorRibbonBar({
  viewAndSelectionActions,
  cellEditActions,
  gridInteractionActions,
  navigationActions,
  fileAndAnalysisActions,
  clipboardAndFindActions,
  rowColumnAndFillActions,
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
  undoStack,
  redoStack,
  hasFrozenRows,
  activeCell,
}) {
  return i.jsx(WorkbookRibbon, {
    onTool: gridInteractionActions.runEditorCommand,
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
    onFormat: cellEditActions.applyCellFormatPatch,
    onUndo: () => navigationActions.undoOrRedo(),
    onRedo: () => navigationActions.undoOrRedo(true),
    onImport: () => importExportState.importFileInputRef.current?.click(),
    onExport: (exportFormat) =>
      void fileAndAnalysisActions.exportWorkbookAs(exportFormat),
    onNew: () => navigationActions.replaceActiveWorkbook(createBlankWorkbook()),
    onDemo: () =>
      navigationActions.replaceActiveWorkbook(buildSampleLedgerWorkbook()),
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
    onAddRow: () => gridInteractionActions.runEditorCommand("insertRow"),
    onAddColumn: () => gridInteractionActions.runEditorCommand("insertColumn"),
    onAutoFitColumns: () =>
      rowColumnAndFillActions.autoFitSelectedAxis("column"),
    onAutoFitRows: () => rowColumnAndFillActions.autoFitSelectedAxis("row"),
    onAutosum: () =>
      viewAndSelectionActions.runGuardedAction(() =>
        cellEditActions.applyCellPatches(
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
      viewAndSelectionActions.runGuardedAction(() => {
        (viewAndSelectionActions.requireSingleSelectionRange(),
          navigationActions.commitWorkbookEdit(
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
    onClear: () =>
      viewAndSelectionActions.runGuardedAction(
        cellEditActions.clearSelectedCellContents,
      ),
    onCopy: () => void clipboardAndFindActions.copyOrCutSelection(false),
    onCut: () => void clipboardAndFindActions.copyOrCutSelection(true),
    onPaste: () => void clipboardAndFindActions.pasteFromClipboard(),
    onFind: () =>
      findReplaceState.setIsFindBarVisible(
        (isFindBarVisible) => !isFindBarVisible,
      ),
    onSort: (sortDirection) => {
      (filterSortState.setActiveColumnSort(
        sortDirection
          ? {
              col: _t(selectionState.activeCellAddress).col,
              direction: sortDirection === "asc" ? 1 : -1,
            }
          : null,
      ),
        viewAndSelectionActions.scrollActivePaneVertically(0),
        selectionState.gridContainerRef.current &&
          (selectionState.gridContainerRef.current.scrollTop = 0));
    },
    onToggleFilter: () => {
      (filterSortState.setIsFilterBarVisible(
        (isFilterBarVisible) => !isFilterBarVisible,
      ),
        filterSortState.setFilterValueSearch(""));
    },
    onShowFormulas: () =>
      viewDisplayState.setIsShowingFormulas(
        (isShowingFormulas) => !isShowingFormulas,
      ),
    onGridlines: () =>
      viewDisplayState.setIsGridlinesVisible(
        (isGridlinesVisible) => !isGridlinesVisible,
      ),
    onHeadings: () =>
      viewDisplayState.setIsHeadingsVisible(
        (isHeadingsVisible) => !isHeadingsVisible,
      ),
    onFreeze: navigationActions.toggleFreezeTopRow,
    onZoom: navigationActions.applyZoomPercent,
    onPanel: dialogState.setSidePanelKind,
    onNote: () =>
      dialogState.setPromptDialogState({
        kind: "note",
        value: activeCell?.note ?? "",
      }),
    onReviewed: () =>
      cellEditActions.applyCellFormatPatch({
        tickmark: activeCell?.tickmark ? "" : "✓",
      }),
  });
}

export { EditorRibbonBar };
