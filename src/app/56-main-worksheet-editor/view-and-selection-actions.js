import {
  Pv,
  anchorsMatchOnAxis,
  enumerateSelectionCells,
  parseSelectionAreas,
} from "../03-workbook-model.js";
import { computeVirtualizedRange } from "../54-workbook-save-conflicts.js";
import {
  logicalToPhysicalPixels,
  resolveCellNavigation,
} from "../53-find-replace-go-to.js";

function createViewAndSelectionActions({
  selectionState,
  filterSortState,
  navigationState,
  operationStatusState,
  viewDisplayState,
  clipboardState,
  activeReliabilityMonitor,
  setActiveSheetId,
  activeSheet,
  selectionAreas,
  sheetSplitPanes,
  splitPaneSizes,
  rowVisibilityIndex,
  columnVisibilityIndex,
  zoomScale,
  rowHeaderWidth,
  getAllCellEntries,
  sortedRowIndexes,
  yn,
  frozenRowIndexes,
  frozenColumnIndexes,
  scrollableRowLayout,
  scrollableColumnLayout,
  logicalRowAt,
  frozenRowsHeight,
  frozenColumnsWidth,
}) {
  function selectSingleRange(range) {
    (selectionState.setActiveSelectionRange(range),
      selectionState.setCompletedSelectionAreas([]));
  }
  function applySelectionAreas(selectionText) {
    const selectionAreas = parseSelectionAreas(selectionText);
    (selectionState.setCompletedSelectionAreas(selectionAreas.slice(0, -1)),
      selectionState.setActiveSelectionRange(selectionAreas.at(-1)));
  }
  function requireSingleSelectionRange() {
    if (selectionAreas.length > 1)
      throw Error(
        "This command requires one rectangular range. Click a cell or enter a single range to continue.",
      );
  }
  function scrollActivePaneVertically(scrollTop) {
    applyPaneScrollOffsets(
      splitPaneSizes.active.current,
      scrollTop,
      splitPaneSizes.current.current[splitPaneSizes.active.current].left,
    );
  }
  function scrollActivePaneHorizontally(scrollLeft) {
    applyPaneScrollOffsets(
      splitPaneSizes.active.current,
      splitPaneSizes.current.current[splitPaneSizes.active.current].top,
      scrollLeft,
    );
  }
  function clearCutSession() {
    (clipboardState.setClipboardSnapshotState(null),
      clipboardState.setMoveCellsSource(null),
      (clipboardState.internalClipboardRef.current = null));
  }
  function computeGridViewport(viewportRect) {
    const rowViewport = Math.max(
        1,
        viewportRect.height -
          ((viewDisplayState.isHeadingsVisible ? 22 : 0) + frozenRowsHeight) *
            zoomScale,
      ),
      columnViewport = Math.max(
        1,
        viewportRect.width - (rowHeaderWidth + frozenColumnsWidth) * zoomScale,
      ),
      virtualizedRange = computeVirtualizedRange(
        scrollableRowLayout,
        viewportRect.top,
        rowViewport,
        zoomScale,
        6,
      ),
      columnWindow = computeVirtualizedRange(
        scrollableColumnLayout,
        viewportRect.left,
        columnViewport,
        zoomScale,
        3,
      ),
      shownRows = Array.from(
        { length: virtualizedRange.last - virtualizedRange.first },
        (Mr, rowOffset) => logicalRowAt(virtualizedRange.first + rowOffset),
      ),
      shownColumns = Array.from(
        { length: columnWindow.last - columnWindow.first },
        (Mr, columnOffset) =>
          columnVisibilityIndex.logicalIndex(
            columnWindow.first + columnOffset + yn,
          ),
      ),
      leading = columnWindow.leading,
      trailing = columnWindow.trailing,
      visibleColumnCount =
        shownColumns.length +
        frozenColumnIndexes.length +
        Number(viewDisplayState.isHeadingsVisible) +
        +(leading > 0) +
        +(trailing > 0);
    return {
      rowViewport: rowViewport,
      columnViewport: columnViewport,
      rowWindow: virtualizedRange,
      columnWindow: columnWindow,
      shownRows: shownRows,
      shownColumns: shownColumns,
      leftGap: leading,
      rightGap: trailing,
      visibleColumnCount: visibleColumnCount,
    };
  }
  function applyPaneScrollOffsets(sourcePaneId, scrollTop, scrollLeft) {
    const gridViewport = computeGridViewport(
        splitPaneSizes.current.current[sourcePaneId],
      ),
      logicalOffset = computeVirtualizedRange(
        scrollableRowLayout,
        scrollTop,
        gridViewport.rowViewport,
        zoomScale,
        0,
      ).logicalOffset,
      logicalColumnOffset = computeVirtualizedRange(
        scrollableColumnLayout,
        scrollLeft,
        gridViewport.columnViewport,
        zoomScale,
        0,
      ).logicalOffset,
      paneOffsets = {};
    for (const targetPaneId of Pv(sheetSplitPanes)) {
      const gridViewport = computeGridViewport(
          splitPaneSizes.current.current[targetPaneId],
        ),
        targetPaneScroll = splitPaneSizes.current.current[targetPaneId],
        targetScrollTop = anchorsMatchOnAxis(sourcePaneId, targetPaneId, "row")
          ? logicalToPhysicalPixels(
              logicalOffset,
              scrollableRowLayout.total,
              gridViewport.rowViewport,
              zoomScale,
            )
          : targetPaneScroll.top,
        left = anchorsMatchOnAxis(sourcePaneId, targetPaneId, "column")
          ? logicalToPhysicalPixels(
              logicalColumnOffset,
              scrollableColumnLayout.total,
              gridViewport.columnViewport,
              zoomScale,
            )
          : targetPaneScroll.left;
      paneOffsets[targetPaneId] = { top: targetScrollTop, left: left };
      const paneElement = splitPaneSizes.refs.current[targetPaneId];
      paneElement &&
        (Math.abs(paneElement.scrollTop - targetScrollTop) > 0.5 &&
          (paneElement.scrollTop = targetScrollTop),
        Math.abs(paneElement.scrollLeft - left) > 0.5 &&
          (paneElement.scrollLeft = left));
    }
    splitPaneSizes.patch(paneOffsets);
  }
  function getSelectedCellAddresses() {
    return enumerateSelectionCells(selectionAreas);
  }
  function buildWorkbookViewState(
    workbook,
    sheetId = activeSheet.id,
    selectionRange = selectionState.activeSelectionRange,
  ) {
    const resolvedSheet =
      workbook.sheets.find((sheet) => sheet.id === sheetId && !sheet.hidden) ??
      workbook.sheets.find((sheet) => !sheet.hidden);
    return {
      activeSheetId: resolvedSheet.id,
      selection: {
        sheetId: resolvedSheet.id,
        range: resolvedSheet.id === sheetId ? selectionRange : "A1",
      },
    };
  }
  function applyWorkbookView(workbookView) {
    const [startAddress, endAddress = startAddress] =
      workbookView.selection.range.split(":");
    (setActiveSheetId(workbookView.activeSheetId),
      selectionState.setActiveCellAddress(startAddress),
      selectionState.setSelectionAnchorAddress(startAddress),
      selectSingleRange(workbookView.selection.range),
      (selectionState.activeCellAddressRef.current = endAddress),
      filterSortState.setActiveColumnSort(null),
      filterSortState.setFilterValueSearch(""),
      scrollActivePaneVertically(0),
      scrollActivePaneHorizontally(0),
      navigationState.setPendingScrollAddress(startAddress));
  }
  function notifyHostCallback(fe, Oe) {
    const showAcknowledgeError = () =>
      operationStatusState.setErrorNotice(
        "The application could not acknowledge this operation. Your local workbook is retained.",
      );
    try {
      const runResult = activeReliabilityMonitor.run("host", fe, Oe);
      runResult &&
        typeof runResult.then == "function" &&
        Promise.resolve(runResult).catch(showAcknowledgeError);
    } catch {
      showAcknowledgeError();
    }
  }
  function runGuardedAction(action) {
    try {
      action();
    } catch (error) {
      operationStatusState.setErrorNotice(
        error instanceof Error
          ? error.message
          : "The operation could not be completed.",
      );
    }
  }
  function focusGridSurface() {
    requestAnimationFrame(() =>
      selectionState.gridContainerRef.current?.focus({ preventScroll: true }),
    );
  }
  function requestCellScroll(cellAddress) {
    navigationState.setPendingScrollAddress(cellAddress);
  }
  function stepVisibleRowIndex(rowIndex, rowDelta) {
    if (!sortedRowIndexes) {
      const viewRowIndex =
        rowVisibilityIndex.viewIndex(rowIndex) ??
        rowVisibilityIndex.countBefore(rowIndex);
      return rowVisibilityIndex.logicalIndex(
        Math.max(
          0,
          Math.min(rowVisibilityIndex.count - 1, viewRowIndex + rowDelta),
        ),
      );
    }
    const orderedRowIndexes = [...frozenRowIndexes, ...sortedRowIndexes],
      rowPosition = orderedRowIndexes.indexOf(rowIndex);
    return (
      orderedRowIndexes[
        Math.max(
          0,
          Math.min(orderedRowIndexes.length - 1, rowPosition + rowDelta),
        )
      ] ?? rowIndex
    );
  }
  function stepVisibleColumnIndex(columnIndex, columnDelta) {
    const columnViewIndex =
      columnVisibilityIndex.viewIndex(columnIndex) ??
      columnVisibilityIndex.countBefore(columnIndex);
    return columnVisibilityIndex.logicalIndex(
      Math.max(
        0,
        Math.min(
          columnVisibilityIndex.count - 1,
          columnViewIndex + columnDelta,
        ),
      ),
    );
  }
  function jumpToDataEdge(row, col, isVertical, direction) {
    const orderedRowIndexes =
        isVertical && sortedRowIndexes
          ? [...frozenRowIndexes, ...sortedRowIndexes]
          : null,
      visibilityIndex = isVertical ? rowVisibilityIndex : columnVisibilityIndex,
      positionByIndex = orderedRowIndexes
        ? new Map(
            orderedRowIndexes.map((orderedIndex, orderedPosition) => [
              orderedIndex,
              orderedPosition,
            ]),
          )
        : null,
      viewPositionOf = (logicalIndex) =>
        positionByIndex
          ? (positionByIndex.get(logicalIndex) ?? null)
          : visibilityIndex.viewIndex(logicalIndex),
      currentPosition = viewPositionOf(isVertical ? row : col);
    if (currentPosition === null) return { row: row, col: col };
    const dataPositions = getAllCellEntries().flatMap((entry) => {
        if (
          entry.cell.raw === null ||
          entry.cell.raw === "" ||
          (isVertical ? entry.col !== col : entry.row !== row)
        )
          return [];
        const cellPosition = viewPositionOf(isVertical ? entry.row : entry.col);
        return cellPosition === null ? [] : [cellPosition];
      }),
      cellNavigation = resolveCellNavigation(
        currentPosition,
        orderedRowIndexes?.length ?? visibilityIndex.count,
        dataPositions,
        direction,
      ),
      targetIndex = orderedRowIndexes
        ? orderedRowIndexes[cellNavigation]
        : visibilityIndex.logicalIndex(cellNavigation);
    return isVertical
      ? { row: targetIndex, col: col }
      : { row: row, col: targetIndex };
  }
  return {
    selectSingleRange,
    applySelectionAreas,
    requireSingleSelectionRange,
    scrollActivePaneVertically,
    scrollActivePaneHorizontally,
    clearCutSession,
    computeGridViewport,
    applyPaneScrollOffsets,
    getSelectedCellAddresses,
    buildWorkbookViewState,
    applyWorkbookView,
    notifyHostCallback,
    runGuardedAction,
    focusGridSurface,
    requestCellScroll,
    stepVisibleRowIndex,
    stepVisibleColumnIndex,
    jumpToDataEdge,
  };
}

export { createViewAndSelectionActions };
