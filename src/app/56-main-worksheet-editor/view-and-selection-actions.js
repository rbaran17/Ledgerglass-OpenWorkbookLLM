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
  activeReliabilityMonitor,
  setActiveSheetId,
  activeSheet,
  setActiveCellAddress,
  setSelectionAnchorAddress,
  activeCellAddressRef,
  activeSelectionRange,
  setActiveSelectionRange,
  setCompletedSelectionAreas,
  selectionAreas,
  setFilterValueSearch,
  setActiveColumnSort,
  sheetSplitPanes,
  splitPaneSizes,
  setPendingScrollAddress,
  setErrorNotice,
  isHeadingsVisible,
  internalClipboardRef,
  setClipboardSnapshotState,
  setMoveCellsSource,
  rowVisibilityIndex,
  columnVisibilityIndex,
  zoomScale,
  rowHeaderWidth,
  gridContainerRef,
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
    (setActiveSelectionRange(fe), setCompletedSelectionAreas([]));
  }
  function applySelectionAreas(fe) {
    const Oe = parseSelectionAreas(fe);
    (setCompletedSelectionAreas(Oe.slice(0, -1)),
      setActiveSelectionRange(Oe.at(-1)));
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
    (setClipboardSnapshotState(null),
      setMoveCellsSource(null),
      (internalClipboardRef.current = null));
  }
  function computeGridViewport(fe) {
    const Oe = Math.max(
        1,
        fe.height -
          ((isHeadingsVisible ? 22 : 0) + frozenRowsHeight) * zoomScale,
      ),
      Be = Math.max(
        1,
        fe.width - (rowHeaderWidth + frozenColumnsWidth) * zoomScale,
      ),
      Ue = computeVirtualizedRange(
        scrollableRowLayout,
        fe.top,
        Oe,
        zoomScale,
        6,
      ),
      Xe = computeVirtualizedRange(
        scrollableColumnLayout,
        fe.left,
        Be,
        zoomScale,
        3,
      ),
      gt = Array.from({ length: Ue.last - Ue.first }, (Mr, Br) =>
        logicalRowAt(Ue.first + Br),
      ),
      Nt = Array.from({ length: Xe.last - Xe.first }, (Mr, Br) =>
        columnVisibilityIndex.logicalIndex(Xe.first + Br + yn),
      ),
      wt = Xe.leading,
      jt = Xe.trailing,
      rr =
        Nt.length +
        frozenColumnIndexes.length +
        Number(isHeadingsVisible) +
        +(wt > 0) +
        +(jt > 0);
    return {
      rowViewport: Oe,
      columnViewport: Be,
      rowWindow: Ue,
      columnWindow: Xe,
      shownRows: gt,
      shownColumns: Nt,
      leftGap: wt,
      rightGap: jt,
      visibleColumnCount: rr,
    };
  }
  function applyPaneScrollOffsets(fe, Oe, Be) {
    const Ue = computeGridViewport(splitPaneSizes.current.current[fe]),
      Xe = computeVirtualizedRange(
        scrollableRowLayout,
        Oe,
        Ue.rowViewport,
        zoomScale,
        0,
      ).logicalOffset,
      gt = computeVirtualizedRange(
        scrollableColumnLayout,
        Be,
        Ue.columnViewport,
        zoomScale,
        0,
      ).logicalOffset,
      Nt = {};
    for (const wt of Pv(sheetSplitPanes)) {
      const jt = computeGridViewport(splitPaneSizes.current.current[wt]),
        rr = splitPaneSizes.current.current[wt],
        Mr = anchorsMatchOnAxis(fe, wt, "row")
          ? logicalToPhysicalPixels(
              Xe,
              scrollableRowLayout.total,
              jt.rowViewport,
              zoomScale,
            )
          : rr.top,
        Br = anchorsMatchOnAxis(fe, wt, "column")
          ? logicalToPhysicalPixels(
              gt,
              scrollableColumnLayout.total,
              jt.columnViewport,
              zoomScale,
            )
          : rr.left;
      Nt[wt] = { top: Mr, left: Br };
      const xr = splitPaneSizes.refs.current[wt];
      xr &&
        (Math.abs(xr.scrollTop - Mr) > 0.5 && (xr.scrollTop = Mr),
        Math.abs(xr.scrollLeft - Br) > 0.5 && (xr.scrollLeft = Br));
    }
    splitPaneSizes.patch(Nt);
  }
  function getSelectedCellAddresses() {
    return enumerateSelectionCells(selectionAreas);
  }
  function buildWorkbookViewState(
    fe,
    Oe = activeSheet.id,
    Be = activeSelectionRange,
  ) {
    const Ue =
      fe.sheets.find((Xe) => Xe.id === Oe && !Xe.hidden) ??
      fe.sheets.find((Xe) => !Xe.hidden);
    return {
      activeSheetId: Ue.id,
      selection: { sheetId: Ue.id, range: Ue.id === Oe ? Be : "A1" },
    };
  }
  function applyWorkbookView(fe) {
    const [Oe, Be = Oe] = fe.selection.range.split(":");
    (setActiveSheetId(fe.activeSheetId),
      setActiveCellAddress(Oe),
      setSelectionAnchorAddress(Oe),
      selectSingleRange(fe.selection.range),
      (activeCellAddressRef.current = Be),
      setActiveColumnSort(null),
      setFilterValueSearch(""),
      scrollActivePaneVertically(0),
      scrollActivePaneHorizontally(0),
      setPendingScrollAddress(Oe));
  }
  function notifyHostCallback(fe, Oe) {
    const Be = () =>
      setErrorNotice(
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
    } catch (Oe) {
      setErrorNotice(
        Oe instanceof Error
          ? Oe.message
          : "The operation could not be completed.",
      );
    }
  }
  function focusGridSurface() {
    requestAnimationFrame(() => {
      return gridContainerRef.current?.focus({ preventScroll: true });
    });
  }
  function requestCellScroll(fe) {
    setPendingScrollAddress(fe);
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
  function jumpToDataEdge(fe, Oe, Be, Ue) {
    const Xe =
        Be && sortedRowIndexes
          ? [...frozenRowIndexes, ...sortedRowIndexes]
          : null,
      gt = Be ? rowVisibilityIndex : columnVisibilityIndex,
      Nt = Xe ? new Map(Xe.map((xr, Sr) => [xr, Sr])) : null,
      wt = (xr) => (Nt ? (Nt.get(xr) ?? null) : gt.viewIndex(xr)),
      jt = wt(Be ? fe : Oe);
    if (jt === null) return { row: fe, col: Oe };
    const rr = getAllCellEntries().flatMap((xr) => {
        if (
          xr.cell.raw === null ||
          xr.cell.raw === "" ||
          (Be ? xr.col !== Oe : xr.row !== fe)
        )
          return [];
        const Sr = wt(Be ? xr.row : xr.col);
        return Sr === null ? [] : [Sr];
      }),
      Mr = resolveCellNavigation(jt, Xe?.length ?? gt.count, rr, Ue),
      Br = Xe ? Xe[Mr] : gt.logicalIndex(Mr);
    return Be ? { row: Br, col: Oe } : { row: fe, col: Br };
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
