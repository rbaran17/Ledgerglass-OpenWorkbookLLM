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
  actor,
  activeSheet,
  isEditingLocked,
  activeCellAddress,
  activeSelectionRange,
  requireSingleSelectionRange,
  isShowingFormulas,
  dragAnchorRef,
  setFillDragState,
  fillDragStateRef,
  setAxisResizeState,
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
  getSelectedCellAddresses,
  commitWorkbookChange,
  runGuardedAction,
  focusGridSurface,
  commitCellEdit,
  requestCellScroll,
}) {
  function resolveSelectedAxisIndexes(fe, Oe) {
    requireSingleSelectionRange();
    const { firstRow, lastRow, firstCol, lastCol } = selectionRangeBounds,
      Nt =
        fe === "column"
          ? firstRow === 0 && lastRow === activeSheet.rowCount - 1
          : firstCol === 0 && lastCol === activeSheet.colCount - 1,
      wt = fe === "column" ? firstCol : firstRow,
      jt = fe === "column" ? lastCol : lastRow;
    if (Oe !== undefined && (!Nt || Oe < wt || Oe > jt)) return [Oe];
    const rr = new Set();
    for (const Br of getAllCellEntries()) {
      const xr = fe === "column" ? Br.col : Br.row;
      xr >= wt && xr <= jt && rr.add(xr);
    }
    const Mr =
      fe === "column" ? activeSheet.columnWidths : activeSheet.rowHeights;
    for (const Br of Object.keys(Mr ?? {})) {
      const xr = Number(Br);
      xr >= wt && xr <= jt && rr.add(xr);
    }
    if (jt - wt < 500) for (let Br = wt; Br <= jt; Br++) rr.add(Br);
    return [...rr].sort((Br, xr) => Br - xr);
  }
  function applyAxisSizes(fe, Oe, Be) {
    const Ue = latestEditStateRef.current.book,
      Xe = Ue.sheets.find((rr) => rr.id === activeSheet.id),
      gt = fe === "column" ? "columnWidths" : "rowHeights";
    if (
      !Object.entries(Oe).some(([rr, Mr]) => {
        return (
          (Xe[gt]?.[rr] ??
            (fe === "column"
              ? defaultColumnWidthOf(Xe)
              : defaultRowHeightOf(Xe))) !== Mr
        );
      })
    )
      return;
    const wt = _t(activeCellAddress),
      jt =
        (shownColumns.includes(wt.col) ||
          frozenColumnIndexes.includes(wt.col)) &&
        (shownRows.includes(wt.row) || frozenRowIndexes.includes(wt.row));
    (commitWorkbookChange({
      ...Ue,
      revision: Ue.revision + 1,
      sheets: Ue.sheets.map((rr) =>
        rr.id === activeSheet.id ? { ...rr, [gt]: { ...rr[gt], ...Oe } } : rr,
      ),
      events: [
        ...Ue.events,
        {
          id: createUuid(),
          timestamp: new Date().toISOString(),
          actor,
          action: Be,
          changes: [],
        },
      ],
    }),
      jt && requestCellScroll(activeCellAddress));
  }
  function autoFitSelectedAxis(fe, Oe) {
    runGuardedAction(() => {
      if (
        isEditingLocked ||
        latestEditStateRef.current.readOnly ||
        !commitCellEdit()
      )
        return;
      const Be = resolveSelectedAxisIndexes(fe, Oe),
        Ue = latestEditStateRef.current.book.sheets.find(
          (jt) => jt.id === activeSheet.id,
        ),
        Xe = calculateWorkbookWithSettings(latestEditStateRef.current.book),
        gt =
          fe === "column"
            ? autoFitColumnWidths(
                Ue,
                Be,
                Xe,
                textMeasurer,
                isShowingFormulas,
                latestEditStateRef.current.book.datePolicy?.dateSystem,
              )
            : autoFitRowHeights(
                Ue,
                Be,
                Xe,
                textMeasurer,
                isShowingFormulas,
                latestEditStateRef.current.book.datePolicy?.dateSystem,
              );
      (applyAxisSizes(
        fe,
        gt,
        `AutoFit ${fe === "column" ? "column widths" : "row heights"}`,
      ),
        setAxisResizeState(null),
        focusGridSurface());
    });
  }
  function beginAxisResize(fe, Oe, Be) {
    if (
      (Be.preventDefault(),
      Be.stopPropagation(),
      Be.button !== 0 || isEditingLocked || !commitCellEdit())
    )
      return;
    const Ue =
      fe === "column"
        ? columnAxisLayout.size(columnVisibilityIndex.viewIndex(Oe) ?? 0)
        : (effectiveRowHeights[Oe] ?? defaultRowHeight);
    setAxisResizeState({
      axis: fe,
      index: Oe,
      start: fe === "column" ? Be.clientX : Be.clientY,
      initial: Ue,
      size: Ue,
    });
  }
  function beginFillDrag(fe) {
    (fe.preventDefault(),
      fe.stopPropagation(),
      !(fe.button !== 0 || isEditingLocked || !commitCellEdit()) &&
        runGuardedAction(() => {
          getSelectedCellAddresses();
          const Oe = {
            book: latestEditStateRef.current.book,
            sheetId: activeSheet.id,
            sourceRange: activeSelectionRange,
            targetRange: activeSelectionRange,
          };
          ((dragAnchorRef.current = null),
            (fillDragStateRef.current = Oe),
            setFillDragState(Oe));
        }));
  }
  function extendFillDragTo(fe) {
    const Oe = fillDragStateRef.current;
    if (!Oe) return;
    const Be = parseRangeAddress(Oe.sourceRange),
      Ue = _t(fe),
      Xe = Math.max(Be.firstRow - Ue.row, Ue.row - Be.lastRow, 0),
      gt = Math.max(Be.firstCol - Ue.col, Ue.col - Be.lastCol, 0),
      Nt = Xe >= gt ? Math.min(Be.firstRow, Ue.row) : Be.firstRow,
      wt = Xe >= gt ? Math.max(Be.lastRow, Ue.row) : Be.lastRow,
      jt = gt > Xe ? Math.min(Be.firstCol, Ue.col) : Be.firstCol,
      rr = gt > Xe ? Math.max(Be.lastCol, Ue.col) : Be.lastCol,
      Mr = wy(
        `${columnIndexToLetters(jt)}${Nt + 1}`,
        `${columnIndexToLetters(rr)}${wt + 1}`,
      ),
      Br = { ...Oe, targetRange: Mr };
    ((fillDragStateRef.current = Br), setFillDragState(Br));
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
