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
  function resolveSelectedAxisIndexes(fe, Oe) {
    viewAndSelectionActions.requireSingleSelectionRange();
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
  function applyAxisSizes(fe, Oe, action) {
    const book = latestEditStateRef.current.book,
      sheet = book.sheets.find((sheet) => sheet.id === activeSheet.id),
      gt = fe === "column" ? "columnWidths" : "rowHeights";
    if (
      !Object.entries(Oe).some(
        ([rr, Mr]) =>
          (sheet[gt]?.[rr] ??
            (fe === "column"
              ? defaultColumnWidthOf(sheet)
              : defaultRowHeightOf(sheet))) !== Mr,
      )
    )
      return;
    const wt = _t(selectionState.activeCellAddress),
      jt =
        (shownColumns.includes(wt.col) ||
          frozenColumnIndexes.includes(wt.col)) &&
        (shownRows.includes(wt.row) || frozenRowIndexes.includes(wt.row));
    (persistenceActions.commitWorkbookChange({
      ...book,
      revision: book.revision + 1,
      sheets: book.sheets.map((sheet) =>
        sheet.id === activeSheet.id
          ? { ...sheet, [gt]: { ...sheet[gt], ...Oe } }
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
      jt &&
        viewAndSelectionActions.requestCellScroll(
          selectionState.activeCellAddress,
        ));
  }
  function autoFitSelectedAxis(fe, Oe) {
    viewAndSelectionActions.runGuardedAction(() => {
      if (
        isEditingLocked ||
        latestEditStateRef.current.readOnly ||
        !cellEditActions.commitCellEdit()
      )
        return;
      const selectedAxisIndexes = resolveSelectedAxisIndexes(fe, Oe),
        sheet = latestEditStateRef.current.book.sheets.find(
          (sheet) => sheet.id === activeSheet.id,
        ),
        workbookWithSettings = calculateWorkbookWithSettings(
          latestEditStateRef.current.book,
        ),
        gt =
          fe === "column"
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
        fe,
        gt,
        `AutoFit ${fe === "column" ? "column widths" : "row heights"}`,
      ),
        viewDisplayState.setAxisResizeState(null),
        viewAndSelectionActions.focusGridSurface());
    });
  }
  function beginAxisResize(axis, Oe, Be) {
    if (
      (Be.preventDefault(),
      Be.stopPropagation(),
      Be.button !== 0 || isEditingLocked || !cellEditActions.commitCellEdit())
    )
      return;
    const Ue =
      axis === "column"
        ? columnAxisLayout.size(columnVisibilityIndex.viewIndex(Oe) ?? 0)
        : (effectiveRowHeights[Oe] ?? defaultRowHeight);
    viewDisplayState.setAxisResizeState({
      axis: axis,
      index: Oe,
      start: axis === "column" ? Be.clientX : Be.clientY,
      initial: Ue,
      size: Ue,
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
          const Oe = {
            book: latestEditStateRef.current.book,
            sheetId: activeSheet.id,
            sourceRange: selectionState.activeSelectionRange,
            targetRange: selectionState.activeSelectionRange,
          };
          ((dragFillState.dragAnchorRef.current = null),
            (fillDragStateRef.current = Oe),
            dragFillState.setFillDragState(Oe));
        }));
  }
  function extendFillDragTo(fe) {
    const Oe = fillDragStateRef.current;
    if (!Oe) return;
    const rangeAddress = parseRangeAddress(Oe.sourceRange),
      Ue = _t(fe),
      Xe = Math.max(
        rangeAddress.firstRow - Ue.row,
        Ue.row - rangeAddress.lastRow,
        0,
      ),
      gt = Math.max(
        rangeAddress.firstCol - Ue.col,
        Ue.col - rangeAddress.lastCol,
        0,
      ),
      Nt =
        Xe >= gt
          ? Math.min(rangeAddress.firstRow, Ue.row)
          : rangeAddress.firstRow,
      wt =
        Xe >= gt
          ? Math.max(rangeAddress.lastRow, Ue.row)
          : rangeAddress.lastRow,
      jt =
        gt > Xe
          ? Math.min(rangeAddress.firstCol, Ue.col)
          : rangeAddress.firstCol,
      rr =
        gt > Xe ? Math.max(rangeAddress.lastCol, Ue.col) : rangeAddress.lastCol,
      targetRange = wy(
        `${columnIndexToLetters(jt)}${Nt + 1}`,
        `${columnIndexToLetters(rr)}${wt + 1}`,
      ),
      Br = { ...Oe, targetRange: targetRange };
    ((fillDragStateRef.current = Br), dragFillState.setFillDragState(Br));
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
