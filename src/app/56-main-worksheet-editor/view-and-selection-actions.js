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
  function selectSingleRange(fe) {
    (selectionState.setActiveSelectionRange(fe),
      selectionState.setCompletedSelectionAreas([]));
  }
  function applySelectionAreas(fe) {
    const Oe = parseSelectionAreas(fe);
    (selectionState.setCompletedSelectionAreas(Oe.slice(0, -1)),
      selectionState.setActiveSelectionRange(Oe.at(-1)));
  }
  function requireSingleSelectionRange() {
    if (selectionAreas.length > 1)
      throw Error(
        "This command requires one rectangular range. Click a cell or enter a single range to continue.",
      );
  }
  function scrollActivePaneVertically(fe) {
    applyPaneScrollOffsets(
      splitPaneSizes.active.current,
      fe,
      splitPaneSizes.current.current[splitPaneSizes.active.current].left,
    );
  }
  function scrollActivePaneHorizontally(fe) {
    applyPaneScrollOffsets(
      splitPaneSizes.active.current,
      splitPaneSizes.current.current[splitPaneSizes.active.current].top,
      fe,
    );
  }
  function clearCutSession() {
    (clipboardState.setClipboardSnapshotState(null),
      clipboardState.setMoveCellsSource(null),
      (clipboardState.internalClipboardRef.current = null));
  }
  function computeGridViewport(fe) {
    const rowViewport = Math.max(
        1,
        fe.height -
          ((viewDisplayState.isHeadingsVisible ? 22 : 0) + frozenRowsHeight) *
            zoomScale,
      ),
      columnViewport = Math.max(
        1,
        fe.width - (rowHeaderWidth + frozenColumnsWidth) * zoomScale,
      ),
      virtualizedRange = computeVirtualizedRange(
        scrollableRowLayout,
        fe.top,
        rowViewport,
        zoomScale,
        6,
      ),
      Xe = computeVirtualizedRange(
        scrollableColumnLayout,
        fe.left,
        columnViewport,
        zoomScale,
        3,
      ),
      shownRows = Array.from(
        { length: virtualizedRange.last - virtualizedRange.first },
        (Mr, Br) => logicalRowAt(virtualizedRange.first + Br),
      ),
      shownColumns = Array.from({ length: Xe.last - Xe.first }, (Mr, Br) =>
        columnVisibilityIndex.logicalIndex(Xe.first + Br + yn),
      ),
      leading = Xe.leading,
      trailing = Xe.trailing,
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
      columnWindow: Xe,
      shownRows: shownRows,
      shownColumns: shownColumns,
      leftGap: leading,
      rightGap: trailing,
      visibleColumnCount: visibleColumnCount,
    };
  }
  function applyPaneScrollOffsets(fe, Oe, Be) {
    const gridViewport = computeGridViewport(
        splitPaneSizes.current.current[fe],
      ),
      logicalOffset = computeVirtualizedRange(
        scrollableRowLayout,
        Oe,
        gridViewport.rowViewport,
        zoomScale,
        0,
      ).logicalOffset,
      gt = computeVirtualizedRange(
        scrollableColumnLayout,
        Be,
        gridViewport.columnViewport,
        zoomScale,
        0,
      ).logicalOffset,
      Nt = {};
    for (const wt of Pv(sheetSplitPanes)) {
      const jt = computeGridViewport(splitPaneSizes.current.current[wt]),
        rr = splitPaneSizes.current.current[wt],
        Mr = anchorsMatchOnAxis(fe, wt, "row")
          ? logicalToPhysicalPixels(
              logicalOffset,
              scrollableRowLayout.total,
              jt.rowViewport,
              zoomScale,
            )
          : rr.top,
        left = anchorsMatchOnAxis(fe, wt, "column")
          ? logicalToPhysicalPixels(
              gt,
              scrollableColumnLayout.total,
              jt.columnViewport,
              zoomScale,
            )
          : rr.left;
      Nt[wt] = { top: Mr, left: left };
      const xr = splitPaneSizes.refs.current[wt];
      xr &&
        (Math.abs(xr.scrollTop - Mr) > 0.5 && (xr.scrollTop = Mr),
        Math.abs(xr.scrollLeft - left) > 0.5 && (xr.scrollLeft = left));
    }
    splitPaneSizes.patch(Nt);
  }
  function getSelectedCellAddresses() {
    return enumerateSelectionCells(selectionAreas);
  }
  function buildWorkbookViewState(
    workbook,
    Oe = activeSheet.id,
    Be = selectionState.activeSelectionRange,
  ) {
    const Ue =
      workbook.sheets.find((sheet) => sheet.id === Oe && !sheet.hidden) ??
      workbook.sheets.find((sheet) => !sheet.hidden);
    return {
      activeSheetId: Ue.id,
      selection: { sheetId: Ue.id, range: Ue.id === Oe ? Be : "A1" },
    };
  }
  function applyWorkbookView(fe) {
    const [Oe, Be = Oe] = fe.selection.range.split(":");
    (setActiveSheetId(fe.activeSheetId),
      selectionState.setActiveCellAddress(Oe),
      selectionState.setSelectionAnchorAddress(Oe),
      selectSingleRange(fe.selection.range),
      (selectionState.activeCellAddressRef.current = Be),
      filterSortState.setActiveColumnSort(null),
      filterSortState.setFilterValueSearch(""),
      scrollActivePaneVertically(0),
      scrollActivePaneHorizontally(0),
      navigationState.setPendingScrollAddress(Oe));
  }
  function notifyHostCallback(fe, Oe) {
    const Be = () =>
      operationStatusState.setErrorNotice(
        "The application could not acknowledge this operation. Your local workbook is retained.",
      );
    try {
      const Ue = activeReliabilityMonitor.run("host", fe, Oe);
      Ue && typeof Ue.then == "function" && Promise.resolve(Ue).catch(Be);
    } catch {
      Be();
    }
  }
  function runGuardedAction(fe) {
    try {
      fe();
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
  function requestCellScroll(fe) {
    navigationState.setPendingScrollAddress(fe);
  }
  function stepVisibleRowIndex(fe, Oe) {
    if (!sortedRowIndexes) {
      const Xe =
        rowVisibilityIndex.viewIndex(fe) ?? rowVisibilityIndex.countBefore(fe);
      return rowVisibilityIndex.logicalIndex(
        Math.max(0, Math.min(rowVisibilityIndex.count - 1, Xe + Oe)),
      );
    }
    const Be = [...frozenRowIndexes, ...sortedRowIndexes],
      Ue = Be.indexOf(fe);
    return Be[Math.max(0, Math.min(Be.length - 1, Ue + Oe))] ?? fe;
  }
  function stepVisibleColumnIndex(fe, Oe) {
    const Be =
      columnVisibilityIndex.viewIndex(fe) ??
      columnVisibilityIndex.countBefore(fe);
    return columnVisibilityIndex.logicalIndex(
      Math.max(0, Math.min(columnVisibilityIndex.count - 1, Be + Oe)),
    );
  }
  function jumpToDataEdge(row, col, Be, Ue) {
    const Xe =
        Be && sortedRowIndexes
          ? [...frozenRowIndexes, ...sortedRowIndexes]
          : null,
      gt = Be ? rowVisibilityIndex : columnVisibilityIndex,
      Nt = Xe ? new Map(Xe.map((xr, Sr) => [xr, Sr])) : null,
      wt = (xr) => (Nt ? (Nt.get(xr) ?? null) : gt.viewIndex(xr)),
      jt = wt(Be ? row : col);
    if (jt === null) return { row: row, col: col };
    const rr = getAllCellEntries().flatMap((xr) => {
        if (
          xr.cell.raw === null ||
          xr.cell.raw === "" ||
          (Be ? xr.col !== col : xr.row !== row)
        )
          return [];
        const Sr = wt(Be ? xr.row : xr.col);
        return Sr === null ? [] : [Sr];
      }),
      cellNavigation = resolveCellNavigation(
        jt,
        Xe?.length ?? gt.count,
        rr,
        Ue,
      ),
      Br = Xe ? Xe[cellNavigation] : gt.logicalIndex(cellNavigation);
    return Be ? { row: Br, col: col } : { row: row, col: Br };
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
