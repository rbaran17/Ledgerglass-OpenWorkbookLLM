import {
  _t,
  calculateWorkbookWithSettings,
  columnIndexToLetters,
  createUuid,
  defaultColumnWidthOf,
  defaultRowHeightOf,
  parseRangeAddress,
} from "../03-workbook-model.js";
import {
  autoFitColumnWidths,
  autoFitRowHeights,
} from "../53-find-replace-go-to.js";
import { wy } from "../55-external-source-refresh.js";

function createRowColumnAndFillActions({
  viewAndSelectionActions,
  persistenceActions,
  cellEditActions,
  selectionState,
  viewDisplayState,
  dragFillState,
  actor,
  activeSheet,
  isEditingLocked,
  fillDragStateRef,
  columnVisibilityIndex,
  defaultRowHeight,
  textMeasurer,
  latestEditStateRef,
  selectionRangeBounds,
  getAllCellEntries,
  effectiveRowHeights,
  columnAxisLayout,
  frozenRowIndexes,
  frozenColumnIndexes,
  shownRows,
  shownColumns,
}) {
  function resolveSelectedAxisIndexes(axis, axisIndex) {
    viewAndSelectionActions.requireSingleSelectionRange();
    const { firstRow, lastRow, firstCol, lastCol } = selectionRangeBounds,
      isFullAxisSelected =
        axis === "column"
          ? firstRow === 0 && lastRow === activeSheet.rowCount - 1
          : firstCol === 0 && lastCol === activeSheet.colCount - 1,
      firstIndex = axis === "column" ? firstCol : firstRow,
      lastIndex = axis === "column" ? lastCol : lastRow;
    if (
      axisIndex !== undefined &&
      (!isFullAxisSelected || axisIndex < firstIndex || axisIndex > lastIndex)
    )
      return [axisIndex];
    const selectedIndexes = new Set();
    for (const cellEntry of getAllCellEntries()) {
      const entryIndex = axis === "column" ? cellEntry.col : cellEntry.row;
      entryIndex >= firstIndex &&
        entryIndex <= lastIndex &&
        selectedIndexes.add(entryIndex);
    }
    const axisSizes =
      axis === "column" ? activeSheet.columnWidths : activeSheet.rowHeights;
    for (const sizeKey of Object.keys(axisSizes ?? {})) {
      const sizeIndex = Number(sizeKey);
      sizeIndex >= firstIndex &&
        sizeIndex <= lastIndex &&
        selectedIndexes.add(sizeIndex);
    }
    if (lastIndex - firstIndex < 500)
      for (let rangeIndex = firstIndex; rangeIndex <= lastIndex; rangeIndex++)
        selectedIndexes.add(rangeIndex);
    return [...selectedIndexes].sort(
      (leftIndex, rightIndex) => leftIndex - rightIndex,
    );
  }
  function applyAxisSizes(axis, sizesByIndex, action) {
    const book = latestEditStateRef.current.book,
      sheet = book.sheets.find((sheet) => sheet.id === activeSheet.id),
      sizeProperty = axis === "column" ? "columnWidths" : "rowHeights";
    if (
      !Object.entries(sizesByIndex).some(
        ([axisIndex, newSize]) =>
          (sheet[sizeProperty]?.[axisIndex] ??
            (axis === "column"
              ? defaultColumnWidthOf(sheet)
              : defaultRowHeightOf(sheet))) !== newSize,
      )
    )
      return;
    const activeCell = _t(selectionState.activeCellAddress),
      isActiveCellShown =
        (shownColumns.includes(activeCell.col) ||
          frozenColumnIndexes.includes(activeCell.col)) &&
        (shownRows.includes(activeCell.row) ||
          frozenRowIndexes.includes(activeCell.row));
    (persistenceActions.commitWorkbookChange({
      ...book,
      revision: book.revision + 1,
      sheets: book.sheets.map((sheet) =>
        sheet.id === activeSheet.id
          ? {
              ...sheet,
              [sizeProperty]: { ...sheet[sizeProperty], ...sizesByIndex },
            }
          : sheet,
      ),
      events: [
        ...book.events,
        {
          id: createUuid(),
          timestamp: new Date().toISOString(),
          actor,
          action: action,
          changes: [],
        },
      ],
    }),
      isActiveCellShown &&
        viewAndSelectionActions.requestCellScroll(
          selectionState.activeCellAddress,
        ));
  }
  function autoFitSelectedAxis(axis, axisIndex) {
    viewAndSelectionActions.runGuardedAction(() => {
      if (
        isEditingLocked ||
        latestEditStateRef.current.readOnly ||
        !cellEditActions.commitCellEdit()
      )
        return;
      const selectedAxisIndexes = resolveSelectedAxisIndexes(axis, axisIndex),
        sheet = latestEditStateRef.current.book.sheets.find(
          (sheet) => sheet.id === activeSheet.id,
        ),
        workbookWithSettings = calculateWorkbookWithSettings(
          latestEditStateRef.current.book,
        ),
        autoFitSizes =
          axis === "column"
            ? autoFitColumnWidths(
                sheet,
                selectedAxisIndexes,
                workbookWithSettings,
                textMeasurer,
                viewDisplayState.isShowingFormulas,
                latestEditStateRef.current.book.datePolicy?.dateSystem,
              )
            : autoFitRowHeights(
                sheet,
                selectedAxisIndexes,
                workbookWithSettings,
                textMeasurer,
                viewDisplayState.isShowingFormulas,
                latestEditStateRef.current.book.datePolicy?.dateSystem,
              );
      (applyAxisSizes(
        axis,
        autoFitSizes,
        `AutoFit ${axis === "column" ? "column widths" : "row heights"}`,
      ),
        viewDisplayState.setAxisResizeState(null),
        viewAndSelectionActions.focusGridSurface());
    });
  }
  function beginAxisResize(axis, axisIndex, event) {
    if (
      (event.preventDefault(),
      event.stopPropagation(),
      event.button !== 0 ||
        isEditingLocked ||
        !cellEditActions.commitCellEdit())
    )
      return;
    const initialSize =
      axis === "column"
        ? columnAxisLayout.size(columnVisibilityIndex.viewIndex(axisIndex) ?? 0)
        : (effectiveRowHeights[axisIndex] ?? defaultRowHeight);
    viewDisplayState.setAxisResizeState({
      axis: axis,
      index: axisIndex,
      start: axis === "column" ? event.clientX : event.clientY,
      initial: initialSize,
      size: initialSize,
    });
  }
  function beginFillDrag(event) {
    (event.preventDefault(),
      event.stopPropagation(),
      !(
        event.button !== 0 ||
        isEditingLocked ||
        !cellEditActions.commitCellEdit()
      ) &&
        viewAndSelectionActions.runGuardedAction(() => {
          viewAndSelectionActions.getSelectedCellAddresses();
          const fillDragState = {
            book: latestEditStateRef.current.book,
            sheetId: activeSheet.id,
            sourceRange: selectionState.activeSelectionRange,
            targetRange: selectionState.activeSelectionRange,
          };
          ((dragFillState.dragAnchorRef.current = null),
            (fillDragStateRef.current = fillDragState),
            dragFillState.setFillDragState(fillDragState));
        }));
  }
  function extendFillDragTo(cellAddress) {
    const fillDragState = fillDragStateRef.current;
    if (!fillDragState) return;
    const rangeAddress = parseRangeAddress(fillDragState.sourceRange),
      cellPosition = _t(cellAddress),
      rowDistance = Math.max(
        rangeAddress.firstRow - cellPosition.row,
        cellPosition.row - rangeAddress.lastRow,
        0,
      ),
      colDistance = Math.max(
        rangeAddress.firstCol - cellPosition.col,
        cellPosition.col - rangeAddress.lastCol,
        0,
      ),
      targetFirstRow =
        rowDistance >= colDistance
          ? Math.min(rangeAddress.firstRow, cellPosition.row)
          : rangeAddress.firstRow,
      targetLastRow =
        rowDistance >= colDistance
          ? Math.max(rangeAddress.lastRow, cellPosition.row)
          : rangeAddress.lastRow,
      targetFirstCol =
        colDistance > rowDistance
          ? Math.min(rangeAddress.firstCol, cellPosition.col)
          : rangeAddress.firstCol,
      targetLastCol =
        colDistance > rowDistance
          ? Math.max(rangeAddress.lastCol, cellPosition.col)
          : rangeAddress.lastCol,
      targetRange = wy(
        `${columnIndexToLetters(targetFirstCol)}${targetFirstRow + 1}`,
        `${columnIndexToLetters(targetLastCol)}${targetLastRow + 1}`,
      ),
      nextFillDragState = { ...fillDragState, targetRange: targetRange };
    ((fillDragStateRef.current = nextFillDragState),
      dragFillState.setFillDragState(nextFillDragState));
  }
  return {
    applyAxisSizes,
    autoFitSelectedAxis,
    beginAxisResize,
    beginFillDrag,
    extendFillDragTo,
  };
}

export { createRowColumnAndFillActions };
