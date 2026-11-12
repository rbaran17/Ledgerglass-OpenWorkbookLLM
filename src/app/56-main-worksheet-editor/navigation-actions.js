import {
  applyWorkbookChange,
  recordCrossWorkbookImport,
} from "../24-merge-audit-and-ribbon-helpers.js";
import {
  Pv,
  _t,
  cn,
  columnIndexToLetters,
  createAxisVisibilityIndex,
  lD,
  parseRangeAddress,
  parseSelectionAreas,
  setFreezePanes,
} from "../03-workbook-model.js";
import { logicalToPhysicalPixels } from "../53-find-replace-go-to.js";
import { wy } from "../55-external-source-refresh.js";
import { parseGoToReference } from "../22-themes-and-cell-styles.js";
import { isDeepFrozenValue } from "../01-schemas-date-functions-and-validation.js";

function createNavigationActions({
  viewAndSelectionActions,
  persistenceActions,
  cellEditActions,
  dialogState,
  selectionState,
  navigationState,
  formulaEntryState,
  filterSortState,
  operationStatusState,
  auditState,
  viewDisplayState,
  findReplaceState,
  actor,
  readOnly,
  activeWorkbook,
  activeSheetId,
  setActiveSheetId,
  activeSheet,
  isEditingLocked,
  selectionAreas,
  sheetSplitPanes,
  splitPaneSizes,
  undoStack,
  setUndoStack,
  redoStack,
  setRedoStack,
  hasFrozenRows,
  rowVisibilityIndex,
  columnVisibilityIndex,
  rowHeaderWidth,
  latestEditStateRef,
  scrollableRowLayout,
  scrollableColumnLayout,
  frozenRowsHeight,
  frozenColumnsWidth,
}) {
  function commitWorkbookEdit(workbook, actionLabel) {
    const book = latestEditStateRef.current.book,
      nextBook = applyWorkbookChange(book, workbook, actor, actionLabel),
      sheet = workbook.sheets.find(
        (sheet) =>
          !book.sheets.some((existingSheet) => existingSheet.id === sheet.id),
      );
    (persistenceActions.commitWorkbookChange(
      nextBook,
      true,
      actionLabel === "Unprotect sheet" &&
        !workbook.sheets.find((sheet) => sheet.id === activeSheet.id)
          ?.protection?.enabled,
    ),
      operationStatusState.setStatusNotice(actionLabel),
      sheet &&
        (dialogState.setActiveDialogMode(null),
        switchActiveWorksheet(sheet.id)));
  }
  function replaceWorkbookSnapshot(workbook, actionLabel) {
    if (readOnly) throw new Error("This workbook is read-only.");
    const book = latestEditStateRef.current.book;
    if (JSON.stringify(book) === JSON.stringify(workbook)) return;
    const isStructureProtectionChanged =
      JSON.stringify(book.structureProtection) !==
      JSON.stringify(workbook.structureProtection);
    (persistenceActions.commitWorkbookChange(
      applyWorkbookChange(book, workbook, actor, actionLabel),
      true,
      true,
    ),
      isStructureProtectionChanged && (setUndoStack([]), setRedoStack([])),
      workbook.sheets.some(
        (sheet) => sheet.id === activeSheet.id && !sheet.hidden,
      ) ||
        viewAndSelectionActions.applyWorkbookView(
          viewAndSelectionActions.buildWorkbookViewState(workbook),
        ),
      operationStatusState.setStatusNotice(actionLabel));
  }
  function commitWorkbookWithNewSheet(workbook, actionLabel) {
    const sheet = workbook.sheets.find(
      (sheet) =>
        !latestEditStateRef.current.book.sheets.some(
          (existingSheet) => existingSheet.id === sheet.id,
        ),
    );
    (replaceWorkbookSnapshot(workbook, actionLabel),
      sheet &&
        (dialogState.setActiveDialogMode(null),
        switchActiveWorksheet(sheet.id)));
  }
  function applyZoomPercent(zoomPercent) {
    if (!Number.isFinite(zoomPercent) || zoomPercent < 50 || zoomPercent > 200)
      return;
    const zoomScale = zoomPercent / 100,
      paneScrollPatch = {};
    for (const paneId of Pv(sheetSplitPanes)) {
      const paneSize = splitPaneSizes.current.current[paneId],
        viewport = viewAndSelectionActions.computeGridViewport(paneSize),
        availableHeight = Math.max(
          1,
          paneSize.height -
            ((viewDisplayState.isHeadingsVisible ? 22 : 0) + frozenRowsHeight) *
              zoomScale,
        ),
        availableWidth = Math.max(
          1,
          paneSize.width - (rowHeaderWidth + frozenColumnsWidth) * zoomScale,
        ),
        scrollTop = logicalToPhysicalPixels(
          viewport.rowWindow.logicalOffset,
          scrollableRowLayout.total,
          availableHeight,
          zoomScale,
        ),
        left = logicalToPhysicalPixels(
          viewport.columnWindow.logicalOffset,
          scrollableColumnLayout.total,
          availableWidth,
          zoomScale,
        );
      paneScrollPatch[paneId] = { top: scrollTop, left: left };
      const paneElement = splitPaneSizes.refs.current[paneId];
      paneElement &&
        ((paneElement.scrollTop = scrollTop), (paneElement.scrollLeft = left));
    }
    (splitPaneSizes.patch(paneScrollPatch),
      viewDisplayState.setZoomPercent(zoomPercent),
      navigationState.setPendingScrollAddress(
        selectionState.activeCellAddress,
      ));
  }
  function selectCellAddress(
    cellAddress,
    shouldExtendSelection = false,
    shouldAddArea = false,
  ) {
    if (!cellEditActions.commitCellEdit()) return;
    if (
      !shouldExtendSelection &&
      !filterSortState.filterValueSearch &&
      !filterSortState.activeColumnSort
    ) {
      const containingMerge = (activeSheet.merges ?? []).find((merge) =>
        lD(merge, cellAddress),
      );
      containingMerge && (cellAddress = containingMerge.split(":")[0]);
    }
    const selectionArea = shouldExtendSelection
      ? wy(selectionState.selectionAnchorAddress, cellAddress)
      : cellAddress;
    let parsedSelectionAreas;
    try {
      parsedSelectionAreas = parseSelectionAreas(
        shouldAddArea ? [...selectionAreas, selectionArea] : [selectionArea],
      );
    } catch (error) {
      operationStatusState.setErrorNotice(error.message);
      return;
    }
    (viewAndSelectionActions.requestCellScroll(cellAddress),
      (selectionState.activeCellAddressRef.current = cellAddress),
      selectionState.setActiveCellAddress(
        shouldExtendSelection
          ? selectionState.selectionAnchorAddress
          : cellAddress,
      ),
      viewAndSelectionActions.applySelectionAreas(parsedSelectionAreas),
      shouldExtendSelection ||
        selectionState.setSelectionAnchorAddress(cellAddress),
      formulaEntryState.setCellEditState(null));
  }
  function switchActiveWorksheet(sheetId) {
    if (
      formulaEntryState.activeCellEditRef.current?.value.startsWith("=") &&
      sheetId !== activeSheet.id
    ) {
      (setActiveSheetId(sheetId),
        filterSortState.setActiveColumnSort(null),
        filterSortState.setFilterValueSearch(""),
        viewAndSelectionActions.scrollActivePaneVertically(0),
        viewAndSelectionActions.scrollActivePaneHorizontally(0),
        formulaEntryState.setCellEditState(null),
        (formulaEntryState.formulaPointerStateRef.current.anchor = undefined),
        (formulaEntryState.formulaPointerStateRef.current.head = undefined),
        formulaEntryState.setActiveFormulaReference(null),
        cellEditActions.restoreFormulaEditorFocus());
      return;
    }
    cellEditActions.commitCellEdit() &&
      ((selectionState.activeCellAddressRef.current = "A1"),
      setActiveSheetId(sheetId),
      selectionState.setActiveCellAddress("A1"),
      selectionState.setSelectionAnchorAddress("A1"),
      viewAndSelectionActions.selectSingleRange("A1"),
      filterSortState.setActiveColumnSort(null),
      filterSortState.setFilterValueSearch(""),
      viewAndSelectionActions.scrollActivePaneVertically(0),
      viewAndSelectionActions.scrollActivePaneHorizontally(0),
      auditState.setAuditAnalysisResult(null),
      auditState.setAuditSourceRange(""),
      selectionState.gridContainerRef.current &&
        ((selectionState.gridContainerRef.current.scrollTop = 0),
        (selectionState.gridContainerRef.current.scrollLeft = 0)),
      viewAndSelectionActions.focusGridSurface());
  }
  function navigateToSheetCell(sheetId, cellAddress) {
    const sheet = latestEditStateRef.current.book.sheets.find(
      (sheet) => sheet.id === sheetId,
    );
    if (!sheet || sheet.hidden)
      throw Error(
        "Unhide the target worksheet with Manage worksheets before navigating to it.",
      );
    const parsedAddress = _t(cellAddress);
    if (
      createAxisVisibilityIndex(sheet.rowCount, cn(sheet, "row")).isHidden(
        parsedAddress.row,
      ) ||
      createAxisVisibilityIndex(sheet.colCount, cn(sheet, "column")).isHidden(
        parsedAddress.col,
      )
    )
      throw Error(
        "The target cell is hidden. Unhide its row/column or expand its outline before navigating.",
      );
    cellEditActions.commitCellEdit() &&
      (filterSortState.setFilterStateByScope((previousFilterState) => {
        const nextFilterState = { ...previousFilterState };
        return (
          delete nextFilterState[JSON.stringify([activeWorkbook.id, sheetId])],
          nextFilterState
        );
      }),
      setActiveSheetId(sheetId),
      selectionState.setActiveCellAddress(cellAddress),
      selectionState.setSelectionAnchorAddress(cellAddress),
      (selectionState.activeCellAddressRef.current = cellAddress),
      viewAndSelectionActions.selectSingleRange(cellAddress),
      filterSortState.setFilterValueSearch(""),
      filterSortState.setActiveColumnSort(null),
      navigationState.setPendingScrollAddress(cellAddress));
  }
  function selectSheetRanges(sheetId, ranges) {
    const selectionAreas = parseSelectionAreas(ranges);
    (navigateToSheetCell(sheetId, selectionAreas.at(-1).split(":")[0]),
      viewAndSelectionActions.applySelectionAreas(selectionAreas),
      dialogState.setActiveDialogMode(null),
      viewAndSelectionActions.focusGridSurface());
  }
  function commitNameBoxReference() {
    viewAndSelectionActions.runGuardedAction(() => {
      if (!cellEditActions.commitCellEdit()) return;
      const goToReference = parseGoToReference(
          activeWorkbook,
          activeSheet.id,
          navigationState.goToReferenceText,
        ),
        ranges = goToReference.ranges,
        rangeAddress = parseRangeAddress(ranges.at(-1)),
        firstCellAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
        lastCellAddress = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
      if (
        goToReference.sheetId === activeSheet.id &&
        (rowVisibilityIndex.isHidden(rangeAddress.firstRow) ||
          columnVisibilityIndex.isHidden(rangeAddress.firstCol))
      )
        throw Error(
          `${firstCellAddress} is hidden. Use View → Hide & Unhide or Data → Group & Outline to reveal it.`,
        );
      (goToReference.sheetId !== activeSheet.id &&
        navigateToSheetCell(goToReference.sheetId, firstCellAddress),
        selectionState.setActiveCellAddress(firstCellAddress),
        selectionState.setSelectionAnchorAddress(firstCellAddress),
        (selectionState.activeCellAddressRef.current = lastCellAddress),
        viewAndSelectionActions.applySelectionAreas(ranges),
        goToReference.sheetId === activeSheet.id &&
          viewAndSelectionActions.requestCellScroll(firstCellAddress));
    });
  }
  function undoOrRedo(isRedo = false) {
    viewAndSelectionActions.runGuardedAction(() => {
      if (isEditingLocked || latestEditStateRef.current.readOnly)
        throw new Error("This workbook is read-only.");
      const historyEntry = isRedo
        ? redoStack[redoStack.length - 1]
        : undoStack[undoStack.length - 1];
      if (!historyEntry) return;
      if (
        historyEntry.id === activeWorkbook.id &&
        JSON.stringify(historyEntry.structureProtection) !==
          JSON.stringify(activeWorkbook.structureProtection)
      )
        throw new Error(
          "Undo cannot change workbook structure protection. Use Manage worksheets to unlock it.",
        );
      const workbook =
        historyEntry.id !== activeWorkbook.id
          ? recordCrossWorkbookImport(
              activeWorkbook,
              historyEntry,
              actor,
              isRedo ? "redo" : "undo",
            )
          : applyWorkbookChange(
              activeWorkbook,
              {
                ...(isDeepFrozenValue(historyEntry)
                  ? historyEntry
                  : structuredClone(historyEntry)),
                ...(activeWorkbook.rowOrigins
                  ? { rowOrigins: structuredClone(activeWorkbook.rowOrigins) }
                  : {}),
                revision: activeWorkbook.revision,
                events: activeWorkbook.events,
              },
              actor,
              isRedo ? "Redo workbook change" : "Undo workbook change",
              { restoringHistory: true },
            );
      (persistenceActions.commitWorkbookChange(workbook, false),
        isRedo
          ? (setRedoStack(redoStack.slice(0, -1)),
            setUndoStack([...undoStack, activeWorkbook]))
          : (setUndoStack(undoStack.slice(0, -1)),
            setRedoStack([...redoStack, activeWorkbook])),
        workbook.sheets.some(
          (sheet) => sheet.id === activeSheetId && !sheet.hidden,
        ) ||
          switchActiveWorksheet(
            workbook.sheets.find((sheet) => !sheet.hidden).id,
          ));
    });
  }
  function replaceActiveWorkbook(workbook) {
    viewAndSelectionActions.runGuardedAction(() => {
      (persistenceActions.commitWorkbookChange(workbook),
        setUndoStack([]),
        setRedoStack([]),
        switchActiveWorksheet(
          workbook.sheets.find((sheet) => !sheet.hidden).id,
        ),
        filterSortState.setIsFilterBarVisible(false),
        findReplaceState.setIsFindBarVisible(false),
        dialogState.setSidePanelKind(null));
    });
  }
  function toggleFreezeTopRow() {
    viewAndSelectionActions.runGuardedAction(() =>
      commitWorkbookEdit(
        setFreezePanes(
          activeWorkbook,
          activeSheet.id,
          hasFrozenRows ? 0 : 1,
          activeSheet.frozenCols ?? 0,
        ),
        hasFrozenRows ? "Unfreeze rows" : "Freeze top row",
      ),
    );
  }
  function moveSelectionAfterCommit(commitKey, isReverse) {
    const pendingCrossSheetTarget =
      navigationState.pendingCrossSheetTargetRef.current;
    if (
      ((navigationState.pendingCrossSheetTargetRef.current = null),
      pendingCrossSheetTarget)
    ) {
      const sheet = activeWorkbook.sheets.find(
          (sheet) => sheet.id === pendingCrossSheetTarget.sheetId,
        ),
        targetCell = _t(pendingCrossSheetTarget.address),
        axis = commitKey === "Enter" ? "row" : "column",
        nextIndex =
          createAxisVisibilityIndex(
            axis === "row" ? sheet.rowCount : sheet.colCount,
            cn(sheet, axis),
          ).next(
            axis === "row" ? targetCell.row : targetCell.col,
            isReverse ? -1 : 1,
          ) ?? (axis === "row" ? targetCell.row : targetCell.col),
        targetAddress = `${columnIndexToLetters(axis === "column" ? nextIndex : targetCell.col)}${(axis === "row" ? nextIndex : targetCell.row) + 1}`;
      (setActiveSheetId(pendingCrossSheetTarget.sheetId),
        selectionState.setActiveCellAddress(targetAddress),
        selectionState.setSelectionAnchorAddress(targetAddress),
        viewAndSelectionActions.selectSingleRange(targetAddress),
        navigationState.setPendingScrollAddress(targetAddress),
        viewAndSelectionActions.focusGridSurface());
      return;
    }
    const activeCell = _t(selectionState.activeCellAddress),
      nextRowIndex =
        commitKey === "Enter"
          ? viewAndSelectionActions.stepVisibleRowIndex(
              activeCell.row,
              isReverse ? -1 : 1,
            )
          : activeCell.row,
      nextColIndex =
        commitKey === "Tab"
          ? viewAndSelectionActions.stepVisibleColumnIndex(
              activeCell.col,
              isReverse ? -1 : 1,
            )
          : activeCell.col;
    (selectCellAddress(
      `${columnIndexToLetters(nextColIndex)}${nextRowIndex + 1}`,
    ),
      viewAndSelectionActions.focusGridSurface());
  }
  function selectAllCells() {
    if (!cellEditActions.commitCellEdit()) return;
    viewAndSelectionActions.selectSingleRange(
      `A1:${columnIndexToLetters(activeSheet.colCount - 1)}${activeSheet.rowCount}`,
    );
    const firstVisibleCellAddress = `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`;
    (selectionState.setActiveCellAddress(firstVisibleCellAddress),
      selectionState.setSelectionAnchorAddress(firstVisibleCellAddress),
      (selectionState.activeCellAddressRef.current = firstVisibleCellAddress),
      viewAndSelectionActions.focusGridSurface());
  }
  function selectRowOrColumn(selectionKind, lineIndex, shouldExtendSelection) {
    if (!cellEditActions.commitCellEdit()) return;
    const anchorIndex = _t(selectionState.selectionAnchorAddress)[
        selectionKind === "column" ? "col" : "row"
      ],
      startIndex = shouldExtendSelection
        ? Math.min(anchorIndex, lineIndex)
        : lineIndex,
      endIndex = shouldExtendSelection
        ? Math.max(anchorIndex, lineIndex)
        : lineIndex,
      activeCellAddress =
        selectionKind === "column"
          ? `${columnIndexToLetters(lineIndex)}${rowVisibilityIndex.logicalIndex(0) + 1}`
          : `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${lineIndex + 1}`;
    (viewAndSelectionActions.selectSingleRange(
      selectionKind === "column"
        ? `${columnIndexToLetters(startIndex)}1:${columnIndexToLetters(endIndex)}${activeSheet.rowCount}`
        : `A${startIndex + 1}:${columnIndexToLetters(activeSheet.colCount - 1)}${endIndex + 1}`,
    ),
      (selectionState.activeCellAddressRef.current = activeCellAddress),
      selectionState.setActiveCellAddress(
        shouldExtendSelection
          ? selectionState.selectionAnchorAddress
          : activeCellAddress,
      ),
      shouldExtendSelection ||
        selectionState.setSelectionAnchorAddress(activeCellAddress),
      viewAndSelectionActions.focusGridSurface());
  }
  return {
    commitWorkbookEdit,
    replaceWorkbookSnapshot,
    commitWorkbookWithNewSheet,
    applyZoomPercent,
    selectCellAddress,
    switchActiveWorksheet,
    navigateToSheetCell,
    selectSheetRanges,
    commitNameBoxReference,
    undoOrRedo,
    replaceActiveWorkbook,
    toggleFreezeTopRow,
    moveSelectionAfterCommit,
    selectAllCells,
    selectRowOrColumn,
  };
}

export { createNavigationActions };
