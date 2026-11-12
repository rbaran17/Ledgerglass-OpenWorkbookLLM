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
  function commitWorkbookEdit(workbook, Oe) {
    const book = latestEditStateRef.current.book,
      Ue = applyWorkbookChange(book, workbook, actor, Oe),
      Xe = workbook.sheets.find(
        (sheet) => !book.sheets.some((jt) => jt.id === sheet.id),
      );
    (persistenceActions.commitWorkbookChange(
      Ue,
      true,
      Oe === "Unprotect sheet" &&
        !workbook.sheets.find((sheet) => sheet.id === activeSheet.id)
          ?.protection?.enabled,
    ),
      operationStatusState.setStatusNotice(Oe),
      Xe &&
        (dialogState.setActiveDialogMode(null), switchActiveWorksheet(Xe.id)));
  }
  function replaceWorkbookSnapshot(workbook, Oe) {
    if (readOnly) throw new Error("This workbook is read-only.");
    const book = latestEditStateRef.current.book;
    if (JSON.stringify(book) === JSON.stringify(workbook)) return;
    const Ue =
      JSON.stringify(book.structureProtection) !==
      JSON.stringify(workbook.structureProtection);
    (persistenceActions.commitWorkbookChange(
      applyWorkbookChange(book, workbook, actor, Oe),
      true,
      true,
    ),
      Ue && (setUndoStack([]), setRedoStack([])),
      workbook.sheets.some(
        (sheet) => sheet.id === activeSheet.id && !sheet.hidden,
      ) ||
        viewAndSelectionActions.applyWorkbookView(
          viewAndSelectionActions.buildWorkbookViewState(workbook),
        ),
      operationStatusState.setStatusNotice(Oe));
  }
  function commitWorkbookWithNewSheet(workbook, Oe) {
    const Be = workbook.sheets.find(
      (sheet) =>
        !latestEditStateRef.current.book.sheets.some(
          (Xe) => Xe.id === sheet.id,
        ),
    );
    (replaceWorkbookSnapshot(workbook, Oe),
      Be &&
        (dialogState.setActiveDialogMode(null), switchActiveWorksheet(Be.id)));
  }
  function applyZoomPercent(fe) {
    if (!Number.isFinite(fe) || fe < 50 || fe > 200) return;
    const zoomScale = fe / 100,
      Be = {};
    for (const Ue of Pv(sheetSplitPanes)) {
      const Xe = splitPaneSizes.current.current[Ue],
        gt = viewAndSelectionActions.computeGridViewport(Xe),
        Nt = Math.max(
          1,
          Xe.height -
            ((viewDisplayState.isHeadingsVisible ? 22 : 0) + frozenRowsHeight) *
              zoomScale,
        ),
        wt = Math.max(
          1,
          Xe.width - (rowHeaderWidth + frozenColumnsWidth) * zoomScale,
        ),
        jt = logicalToPhysicalPixels(
          gt.rowWindow.logicalOffset,
          scrollableRowLayout.total,
          Nt,
          zoomScale,
        ),
        left = logicalToPhysicalPixels(
          gt.columnWindow.logicalOffset,
          scrollableColumnLayout.total,
          wt,
          zoomScale,
        );
      Be[Ue] = { top: jt, left: left };
      const Mr = splitPaneSizes.refs.current[Ue];
      Mr && ((Mr.scrollTop = jt), (Mr.scrollLeft = left));
    }
    (splitPaneSizes.patch(Be),
      viewDisplayState.setZoomPercent(fe),
      navigationState.setPendingScrollAddress(
        selectionState.activeCellAddress,
      ));
  }
  function selectCellAddress(fe, Oe = false, Be = false) {
    if (!cellEditActions.commitCellEdit()) return;
    if (
      !Oe &&
      !filterSortState.filterValueSearch &&
      !filterSortState.activeColumnSort
    ) {
      const gt = (activeSheet.merges ?? []).find((Nt) => lD(Nt, fe));
      gt && (fe = gt.split(":")[0]);
    }
    const Ue = Oe ? wy(selectionState.selectionAnchorAddress, fe) : fe;
    let Xe;
    try {
      Xe = parseSelectionAreas(Be ? [...selectionAreas, Ue] : [Ue]);
    } catch (error) {
      operationStatusState.setErrorNotice(error.message);
      return;
    }
    (viewAndSelectionActions.requestCellScroll(fe),
      (selectionState.activeCellAddressRef.current = fe),
      selectionState.setActiveCellAddress(
        Oe ? selectionState.selectionAnchorAddress : fe,
      ),
      viewAndSelectionActions.applySelectionAreas(Xe),
      Oe || selectionState.setSelectionAnchorAddress(fe),
      formulaEntryState.setCellEditState(null));
  }
  function switchActiveWorksheet(fe) {
    if (
      formulaEntryState.activeCellEditRef.current?.value.startsWith("=") &&
      fe !== activeSheet.id
    ) {
      (setActiveSheetId(fe),
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
      setActiveSheetId(fe),
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
  function navigateToSheetCell(fe, Oe) {
    const Be = latestEditStateRef.current.book.sheets.find(
      (sheet) => sheet.id === fe,
    );
    if (!Be || Be.hidden)
      throw Error(
        "Unhide the target worksheet with Manage worksheets before navigating to it.",
      );
    const Ue = _t(Oe);
    if (
      createAxisVisibilityIndex(Be.rowCount, cn(Be, "row")).isHidden(Ue.row) ||
      createAxisVisibilityIndex(Be.colCount, cn(Be, "column")).isHidden(Ue.col)
    )
      throw Error(
        "The target cell is hidden. Unhide its row/column or expand its outline before navigating.",
      );
    cellEditActions.commitCellEdit() &&
      (filterSortState.setFilterStateByScope((Xe) => {
        const gt = { ...Xe };
        return (delete gt[JSON.stringify([activeWorkbook.id, fe])], gt);
      }),
      setActiveSheetId(fe),
      selectionState.setActiveCellAddress(Oe),
      selectionState.setSelectionAnchorAddress(Oe),
      (selectionState.activeCellAddressRef.current = Oe),
      viewAndSelectionActions.selectSingleRange(Oe),
      filterSortState.setFilterValueSearch(""),
      filterSortState.setActiveColumnSort(null),
      navigationState.setPendingScrollAddress(Oe));
  }
  function selectSheetRanges(fe, Oe) {
    const Be = parseSelectionAreas(Oe);
    (navigateToSheetCell(fe, Be.at(-1).split(":")[0]),
      viewAndSelectionActions.applySelectionAreas(Be),
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
        Ue = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
        Xe = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
      if (
        goToReference.sheetId === activeSheet.id &&
        (rowVisibilityIndex.isHidden(rangeAddress.firstRow) ||
          columnVisibilityIndex.isHidden(rangeAddress.firstCol))
      )
        throw Error(
          `${Ue} is hidden. Use View → Hide & Unhide or Data → Group & Outline to reveal it.`,
        );
      (goToReference.sheetId !== activeSheet.id &&
        navigateToSheetCell(goToReference.sheetId, Ue),
        selectionState.setActiveCellAddress(Ue),
        selectionState.setSelectionAnchorAddress(Ue),
        (selectionState.activeCellAddressRef.current = Xe),
        viewAndSelectionActions.applySelectionAreas(ranges),
        goToReference.sheetId === activeSheet.id &&
          viewAndSelectionActions.requestCellScroll(Ue));
    });
  }
  function undoOrRedo(fe = false) {
    viewAndSelectionActions.runGuardedAction(() => {
      if (isEditingLocked || latestEditStateRef.current.readOnly)
        throw new Error("This workbook is read-only.");
      const Oe = fe
        ? redoStack[redoStack.length - 1]
        : undoStack[undoStack.length - 1];
      if (!Oe) return;
      if (
        Oe.id === activeWorkbook.id &&
        JSON.stringify(Oe.structureProtection) !==
          JSON.stringify(activeWorkbook.structureProtection)
      )
        throw new Error(
          "Undo cannot change workbook structure protection. Use Manage worksheets to unlock it.",
        );
      const workbook =
        Oe.id !== activeWorkbook.id
          ? recordCrossWorkbookImport(
              activeWorkbook,
              Oe,
              actor,
              fe ? "redo" : "undo",
            )
          : applyWorkbookChange(
              activeWorkbook,
              {
                ...(isDeepFrozenValue(Oe) ? Oe : structuredClone(Oe)),
                ...(activeWorkbook.rowOrigins
                  ? { rowOrigins: structuredClone(activeWorkbook.rowOrigins) }
                  : {}),
                revision: activeWorkbook.revision,
                events: activeWorkbook.events,
              },
              actor,
              fe ? "Redo workbook change" : "Undo workbook change",
              { restoringHistory: true },
            );
      (persistenceActions.commitWorkbookChange(workbook, false),
        fe
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
  function moveSelectionAfterCommit(fe, Oe) {
    const Be = navigationState.pendingCrossSheetTargetRef.current;
    if (((navigationState.pendingCrossSheetTargetRef.current = null), Be)) {
      const Nt = activeWorkbook.sheets.find((sheet) => sheet.id === Be.sheetId),
        wt = _t(Be.address),
        jt = fe === "Enter" ? "row" : "column",
        Mr =
          createAxisVisibilityIndex(
            jt === "row" ? Nt.rowCount : Nt.colCount,
            cn(Nt, jt),
          ).next(jt === "row" ? wt.row : wt.col, Oe ? -1 : 1) ??
          (jt === "row" ? wt.row : wt.col),
        Br = `${columnIndexToLetters(jt === "column" ? Mr : wt.col)}${(jt === "row" ? Mr : wt.row) + 1}`;
      (setActiveSheetId(Be.sheetId),
        selectionState.setActiveCellAddress(Br),
        selectionState.setSelectionAnchorAddress(Br),
        viewAndSelectionActions.selectSingleRange(Br),
        navigationState.setPendingScrollAddress(Br),
        viewAndSelectionActions.focusGridSurface());
      return;
    }
    const Ue = _t(selectionState.activeCellAddress),
      Xe =
        fe === "Enter"
          ? viewAndSelectionActions.stepVisibleRowIndex(Ue.row, Oe ? -1 : 1)
          : Ue.row,
      gt =
        fe === "Tab"
          ? viewAndSelectionActions.stepVisibleColumnIndex(Ue.col, Oe ? -1 : 1)
          : Ue.col;
    (selectCellAddress(`${columnIndexToLetters(gt)}${Xe + 1}`),
      viewAndSelectionActions.focusGridSurface());
  }
  function selectAllCells() {
    if (!cellEditActions.commitCellEdit()) return;
    viewAndSelectionActions.selectSingleRange(
      `A1:${columnIndexToLetters(activeSheet.colCount - 1)}${activeSheet.rowCount}`,
    );
    const fe = `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`;
    (selectionState.setActiveCellAddress(fe),
      selectionState.setSelectionAnchorAddress(fe),
      (selectionState.activeCellAddressRef.current = fe),
      viewAndSelectionActions.focusGridSurface());
  }
  function selectRowOrColumn(fe, Oe, Be) {
    if (!cellEditActions.commitCellEdit()) return;
    const Ue = _t(selectionState.selectionAnchorAddress)[
        fe === "column" ? "col" : "row"
      ],
      Xe = Be ? Math.min(Ue, Oe) : Oe,
      gt = Be ? Math.max(Ue, Oe) : Oe,
      Nt =
        fe === "column"
          ? `${columnIndexToLetters(Oe)}${rowVisibilityIndex.logicalIndex(0) + 1}`
          : `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${Oe + 1}`;
    (viewAndSelectionActions.selectSingleRange(
      fe === "column"
        ? `${columnIndexToLetters(Xe)}1:${columnIndexToLetters(gt)}${activeSheet.rowCount}`
        : `A${Xe + 1}:${columnIndexToLetters(activeSheet.colCount - 1)}${gt + 1}`,
    ),
      (selectionState.activeCellAddressRef.current = Nt),
      selectionState.setActiveCellAddress(
        Be ? selectionState.selectionAnchorAddress : Nt,
      ),
      Be || selectionState.setSelectionAnchorAddress(Nt),
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
