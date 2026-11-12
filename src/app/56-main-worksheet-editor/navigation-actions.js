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
  selectSingleRange,
  applySelectionAreas,
  selectionAreas,
  sheetSplitPanes,
  splitPaneSizes,
  scrollActivePaneVertically,
  scrollActivePaneHorizontally,
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
  computeGridViewport,
  buildWorkbookViewState,
  applyWorkbookView,
  commitWorkbookChange,
  runGuardedAction,
  restoreFormulaEditorFocus,
  focusGridSurface,
  commitCellEdit,
  requestCellScroll,
  stepVisibleRowIndex,
  stepVisibleColumnIndex,
}) {
  function commitWorkbookEdit(fe, Oe) {
    const Be = latestEditStateRef.current.book,
      Ue = applyWorkbookChange(Be, fe, actor, Oe),
      Xe = fe.sheets.find((wt) => !Be.sheets.some((jt) => jt.id === wt.id));
    (commitWorkbookChange(
      Ue,
      true,
      Oe === "Unprotect sheet" &&
        !fe.sheets.find((wt) => wt.id === activeSheet.id)?.protection?.enabled,
    ),
      operationStatusState.setStatusNotice(Oe),
      Xe &&
        (dialogState.setActiveDialogMode(null), switchActiveWorksheet(Xe.id)));
  }
  function replaceWorkbookSnapshot(fe, Oe) {
    if (readOnly) throw new Error("This workbook is read-only.");
    const Be = latestEditStateRef.current.book;
    if (JSON.stringify(Be) === JSON.stringify(fe)) return;
    const Ue =
      JSON.stringify(Be.structureProtection) !==
      JSON.stringify(fe.structureProtection);
    (commitWorkbookChange(applyWorkbookChange(Be, fe, actor, Oe), true, true),
      Ue && (setUndoStack([]), setRedoStack([])),
      fe.sheets.some((Xe) => Xe.id === activeSheet.id && !Xe.hidden) ||
        applyWorkbookView(buildWorkbookViewState(fe)),
      operationStatusState.setStatusNotice(Oe));
  }
  function commitWorkbookWithNewSheet(fe, Oe) {
    const Be = fe.sheets.find(
      (Ue) =>
        !latestEditStateRef.current.book.sheets.some((Xe) => Xe.id === Ue.id),
    );
    (replaceWorkbookSnapshot(fe, Oe),
      Be &&
        (dialogState.setActiveDialogMode(null), switchActiveWorksheet(Be.id)));
  }
  function applyZoomPercent(fe) {
    if (!Number.isFinite(fe) || fe < 50 || fe > 200) return;
    const Oe = fe / 100,
      Be = {};
    for (const Ue of Pv(sheetSplitPanes)) {
      const Xe = splitPaneSizes.current.current[Ue],
        gt = computeGridViewport(Xe),
        Nt = Math.max(
          1,
          Xe.height -
            ((viewDisplayState.isHeadingsVisible ? 22 : 0) + frozenRowsHeight) *
              Oe,
        ),
        wt = Math.max(1, Xe.width - (rowHeaderWidth + frozenColumnsWidth) * Oe),
        jt = logicalToPhysicalPixels(
          gt.rowWindow.logicalOffset,
          scrollableRowLayout.total,
          Nt,
          Oe,
        ),
        rr = logicalToPhysicalPixels(
          gt.columnWindow.logicalOffset,
          scrollableColumnLayout.total,
          wt,
          Oe,
        );
      Be[Ue] = { top: jt, left: rr };
      const Mr = splitPaneSizes.refs.current[Ue];
      Mr && ((Mr.scrollTop = jt), (Mr.scrollLeft = rr));
    }
    (splitPaneSizes.patch(Be),
      viewDisplayState.setZoomPercent(fe),
      navigationState.setPendingScrollAddress(
        selectionState.activeCellAddress,
      ));
  }
  function selectCellAddress(fe, Oe = false, Be = false) {
    if (!commitCellEdit()) return;
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
    } catch (gt) {
      operationStatusState.setErrorNotice(gt.message);
      return;
    }
    (requestCellScroll(fe),
      (selectionState.activeCellAddressRef.current = fe),
      selectionState.setActiveCellAddress(
        Oe ? selectionState.selectionAnchorAddress : fe,
      ),
      applySelectionAreas(Xe),
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
        scrollActivePaneVertically(0),
        scrollActivePaneHorizontally(0),
        formulaEntryState.setCellEditState(null),
        (formulaEntryState.formulaPointerStateRef.current.anchor = undefined),
        (formulaEntryState.formulaPointerStateRef.current.head = undefined),
        formulaEntryState.setActiveFormulaReference(null),
        restoreFormulaEditorFocus());
      return;
    }
    commitCellEdit() &&
      ((selectionState.activeCellAddressRef.current = "A1"),
      setActiveSheetId(fe),
      selectionState.setActiveCellAddress("A1"),
      selectionState.setSelectionAnchorAddress("A1"),
      selectSingleRange("A1"),
      filterSortState.setActiveColumnSort(null),
      filterSortState.setFilterValueSearch(""),
      scrollActivePaneVertically(0),
      scrollActivePaneHorizontally(0),
      auditState.setAuditAnalysisResult(null),
      auditState.setAuditSourceRange(""),
      selectionState.gridContainerRef.current &&
        ((selectionState.gridContainerRef.current.scrollTop = 0),
        (selectionState.gridContainerRef.current.scrollLeft = 0)),
      focusGridSurface());
  }
  function navigateToSheetCell(fe, Oe) {
    const Be = latestEditStateRef.current.book.sheets.find(
      (Xe) => Xe.id === fe,
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
    commitCellEdit() &&
      (filterSortState.setFilterStateByScope((Xe) => {
        const gt = { ...Xe };
        return (delete gt[JSON.stringify([activeWorkbook.id, fe])], gt);
      }),
      setActiveSheetId(fe),
      selectionState.setActiveCellAddress(Oe),
      selectionState.setSelectionAnchorAddress(Oe),
      (selectionState.activeCellAddressRef.current = Oe),
      selectSingleRange(Oe),
      filterSortState.setFilterValueSearch(""),
      filterSortState.setActiveColumnSort(null),
      navigationState.setPendingScrollAddress(Oe));
  }
  function selectSheetRanges(fe, Oe) {
    const Be = parseSelectionAreas(Oe);
    (navigateToSheetCell(fe, Be.at(-1).split(":")[0]),
      applySelectionAreas(Be),
      dialogState.setActiveDialogMode(null),
      focusGridSurface());
  }
  function commitNameBoxReference() {
    runGuardedAction(() => {
      if (!commitCellEdit()) return;
      const fe = parseGoToReference(
          activeWorkbook,
          activeSheet.id,
          navigationState.goToReferenceText,
        ),
        Oe = fe.ranges,
        Be = parseRangeAddress(Oe.at(-1)),
        Ue = `${columnIndexToLetters(Be.firstCol)}${Be.firstRow + 1}`,
        Xe = `${columnIndexToLetters(Be.lastCol)}${Be.lastRow + 1}`;
      if (
        fe.sheetId === activeSheet.id &&
        (rowVisibilityIndex.isHidden(Be.firstRow) ||
          columnVisibilityIndex.isHidden(Be.firstCol))
      )
        throw Error(
          `${Ue} is hidden. Use View → Hide & Unhide or Data → Group & Outline to reveal it.`,
        );
      (fe.sheetId !== activeSheet.id && navigateToSheetCell(fe.sheetId, Ue),
        selectionState.setActiveCellAddress(Ue),
        selectionState.setSelectionAnchorAddress(Ue),
        (selectionState.activeCellAddressRef.current = Xe),
        applySelectionAreas(Oe),
        fe.sheetId === activeSheet.id && requestCellScroll(Ue));
    });
  }
  function undoOrRedo(fe = false) {
    runGuardedAction(() => {
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
      const Be =
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
      (commitWorkbookChange(Be, false),
        fe
          ? (setRedoStack(redoStack.slice(0, -1)),
            setUndoStack([...undoStack, activeWorkbook]))
          : (setUndoStack(undoStack.slice(0, -1)),
            setRedoStack([...redoStack, activeWorkbook])),
        Be.sheets.some((Ue) => Ue.id === activeSheetId && !Ue.hidden) ||
          switchActiveWorksheet(Be.sheets.find((Ue) => !Ue.hidden).id));
    });
  }
  function replaceActiveWorkbook(fe) {
    runGuardedAction(() => {
      (commitWorkbookChange(fe),
        setUndoStack([]),
        setRedoStack([]),
        switchActiveWorksheet(fe.sheets.find((Oe) => !Oe.hidden).id),
        filterSortState.setIsFilterBarVisible(false),
        findReplaceState.setIsFindBarVisible(false),
        dialogState.setSidePanelKind(null));
    });
  }
  function toggleFreezeTopRow() {
    runGuardedAction(() =>
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
      const Nt = activeWorkbook.sheets.find((xr) => xr.id === Be.sheetId),
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
        selectSingleRange(Br),
        navigationState.setPendingScrollAddress(Br),
        focusGridSurface());
      return;
    }
    const Ue = _t(selectionState.activeCellAddress),
      Xe = fe === "Enter" ? stepVisibleRowIndex(Ue.row, Oe ? -1 : 1) : Ue.row,
      gt = fe === "Tab" ? stepVisibleColumnIndex(Ue.col, Oe ? -1 : 1) : Ue.col;
    (selectCellAddress(`${columnIndexToLetters(gt)}${Xe + 1}`),
      focusGridSurface());
  }
  function selectAllCells() {
    if (!commitCellEdit()) return;
    selectSingleRange(
      `A1:${columnIndexToLetters(activeSheet.colCount - 1)}${activeSheet.rowCount}`,
    );
    const fe = `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`;
    (selectionState.setActiveCellAddress(fe),
      selectionState.setSelectionAnchorAddress(fe),
      (selectionState.activeCellAddressRef.current = fe),
      focusGridSurface());
  }
  function selectRowOrColumn(fe, Oe, Be) {
    if (!commitCellEdit()) return;
    const Ue = _t(selectionState.selectionAnchorAddress)[
        fe === "column" ? "col" : "row"
      ],
      Xe = Be ? Math.min(Ue, Oe) : Oe,
      gt = Be ? Math.max(Ue, Oe) : Oe,
      Nt =
        fe === "column"
          ? `${columnIndexToLetters(Oe)}${rowVisibilityIndex.logicalIndex(0) + 1}`
          : `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${Oe + 1}`;
    (selectSingleRange(
      fe === "column"
        ? `${columnIndexToLetters(Xe)}1:${columnIndexToLetters(gt)}${activeSheet.rowCount}`
        : `A${Xe + 1}:${columnIndexToLetters(activeSheet.colCount - 1)}${gt + 1}`,
    ),
      (selectionState.activeCellAddressRef.current = Nt),
      selectionState.setActiveCellAddress(
        Be ? selectionState.selectionAnchorAddress : Nt,
      ),
      Be || selectionState.setSelectionAnchorAddress(Nt),
      focusGridSurface());
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
