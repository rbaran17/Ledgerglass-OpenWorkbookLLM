import i from "react/jsx-runtime";
import { WorkbookRibbon } from "../26-workbook-ribbon.js";
import { _t, applyAutoSum, createBlankWorkbook } from "../03-workbook-model.js";
import { buildSampleLedgerWorkbook } from "../24-merge-audit-and-ribbon-helpers.js";
import {
  applyCellPreview,
  buildFillPreview,
} from "../04-cell-input-paste-and-fill.js";

function renderEditorRibbon({
  readOnly,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  activeCellAddress,
  activeSelectionRange,
  requireSingleSelectionRange,
  setSidePanelKind,
  setFilterValueSearch,
  setActiveColumnSort,
  isShowingFormulas,
  setIsShowingFormulas,
  scrollActivePaneVertically,
  undoStack,
  redoStack,
  isOperationBusy,
  setPromptDialogState,
  isGridlinesVisible,
  setIsGridlinesVisible,
  isHeadingsVisible,
  setIsHeadingsVisible,
  zoomPercent,
  setIsFilterBarVisible,
  setIsFindBarVisible,
  hasFrozenRows,
  importFileInputRef,
  gridContainerRef,
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
    busy: isOperationBusy,
    canUndo: undoStack.length > 0,
    canRedo: redoStack.length > 0,
    cell: activeCell,
    showFormulas: isShowingFormulas,
    gridlines: isGridlinesVisible,
    headings: isHeadingsVisible,
    freezeRows: hasFrozenRows,
    zoom: zoomPercent,
    onFormat: applyCellFormatPatch,
    onUndo: () => undoOrRedo(),
    onRedo: () => undoOrRedo(true),
    onImport: () => importFileInputRef.current?.click(),
    onExport: (fe) => void exportWorkbookAs(fe),
    onNew: () => replaceActiveWorkbook(createBlankWorkbook()),
    onDemo: () => replaceActiveWorkbook(buildSampleLedgerWorkbook()),
    onRename: () =>
      setPromptDialogState({
        kind: "rename",
        value: activeWorkbook.name,
      }),
    onAddSheet: () =>
      setPromptDialogState({
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
          [applyAutoSum(activeWorkbook, activeSheet.id, activeCellAddress)],
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
                activeSelectionRange,
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
    onFind: () => setIsFindBarVisible((fe) => !fe),
    onSort: (fe) => {
      (setActiveColumnSort(
        fe
          ? {
              col: _t(activeCellAddress).col,
              direction: fe === "asc" ? 1 : -1,
            }
          : null,
      ),
        scrollActivePaneVertically(0),
        gridContainerRef.current && (gridContainerRef.current.scrollTop = 0));
    },
    onToggleFilter: () => {
      (setIsFilterBarVisible((fe) => !fe), setFilterValueSearch(""));
    },
    onShowFormulas: () => setIsShowingFormulas((fe) => !fe),
    onGridlines: () => setIsGridlinesVisible((fe) => !fe),
    onHeadings: () => setIsHeadingsVisible((fe) => !fe),
    onFreeze: toggleFreezeTopRow,
    onZoom: applyZoomPercent,
    onPanel: setSidePanelKind,
    onNote: () =>
      setPromptDialogState({
        kind: "note",
        value: activeCell?.note ?? "",
      }),
    onReviewed: () =>
      applyCellFormatPatch({
        tickmark: activeCell?.tickmark ? "" : "✓",
      }),
  });
}

export { renderEditorRibbon };
