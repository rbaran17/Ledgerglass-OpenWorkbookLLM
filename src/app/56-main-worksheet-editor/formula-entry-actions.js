import {
  $ae,
  findFormulaReferenceSpan,
  replaceFormulaReferenceSpan,
} from "../04-cell-input-paste-and-fill.js";
import { wy } from "../55-external-source-refresh.js";
import {
  _t,
  columnIndexToLetters,
  parseRangeAddress,
} from "../03-workbook-model.js";

function createFormulaEntryActions({
  cellEditActions,
  viewAndSelectionActions,
  navigationActions,
  formulaEntryState,
  selectionState,
  operationStatusState,
  dragFillState,
  activeWorkbook,
  activeSheet,
}) {
  function insertFormulaReference(fe, Oe = false, Be) {
    const Ue = formulaEntryState.activeCellEditRef.current,
      Xe = formulaEntryState.formulaPointerStateRef.current;
    if (
      !Ue?.value.startsWith("=") ||
      cellEditActions.isFormulaEditingLocked() ||
      operationStatusState.isOperationBusy
    )
      return false;
    const gt = Xe.span ?? findFormulaReferenceSpan(Ue.value, Xe.start, Xe.end);
    if (!gt) return false;
    const Nt =
        Oe && Xe.sheetId === activeSheet.id && Xe.anchor ? Xe.anchor : fe,
      range = Be?.range ?? (Nt === fe ? fe : wy(Nt, fe)),
      jt = replaceFormulaReferenceSpan(
        Ue.value,
        gt,
        range,
        Ue.sheetId === activeSheet.id ? undefined : activeSheet.name,
      );
    return (
      cellEditActions.beginCellEdit(
        jt.value,
        !!Xe.input?.classList.contains("lg-cell-editor"),
      ),
      (Xe.span = jt.span),
      (Xe.start = Xe.end = jt.span.end),
      (Xe.anchor = Nt),
      (Xe.head = fe),
      (Xe.sheetId = activeSheet.id),
      formulaEntryState.setFormulaCaretPosition(jt.span.end),
      formulaEntryState.setActiveFormulaReference({
        sheetId: activeSheet.id,
        range: range,
        bounds: Be?.bounds ?? parseRangeAddress(range),
      }),
      cellEditActions.restoreFormulaEditorFocus(),
      true
    );
  }
  function handleFormulaPointerDown(event) {
    if (
      event.button !== 0 ||
      !formulaEntryState.activeCellEditRef.current?.value.startsWith("=")
    )
      return false;
    const target = event.target;
    if (target.closest('input,button,a,[role="separator"]')) return false;
    const Be = target.closest("[data-address]"),
      Ue = target.closest("[data-column]"),
      Xe = target.closest("[data-row-header]");
    let gt = false;
    if (Be) gt = insertFormulaReference(Be.dataset.address, event.shiftKey);
    else if (Ue) {
      const wt = Number(Ue.dataset.column);
      gt = insertFormulaReference(`${columnIndexToLetters(wt)}1`, false, {
        range: `${columnIndexToLetters(wt)}:${columnIndexToLetters(wt)}`,
        bounds: {
          firstRow: 0,
          lastRow: activeSheet.rowCount - 1,
          firstCol: wt,
          lastCol: wt,
        },
      });
    } else if (Xe) {
      const wt = Number(Xe.dataset.rowHeader);
      gt = insertFormulaReference(`A${wt + 1}`, false, {
        range: `${wt + 1}:${wt + 1}`,
        bounds: {
          firstRow: wt,
          lastRow: wt,
          firstCol: 0,
          lastCol: activeSheet.colCount - 1,
        },
      });
    }
    return gt
      ? (event.preventDefault(),
        event.stopPropagation(),
        (formulaEntryState.formulaPointerStateRef.current.dragging = !!Be),
        (dragFillState.dragAnchorRef.current = null),
        true)
      : false;
  }
  function handleFormulaKeyDown(event) {
    if (
      event.nativeEvent.isComposing ||
      event.altKey ||
      cellEditActions.isFormulaEditingLocked() ||
      operationStatusState.isOperationBusy
    )
      return false;
    const currentTarget = event.currentTarget;
    if (!currentTarget.value.startsWith("=")) return false;
    cellEditActions.syncFormulaEditorSelection(currentTarget);
    const Be = formulaEntryState.formulaPointerStateRef.current;
    if (event.key === "F2" && !event.ctrlKey && !event.metaKey)
      return (
        event.preventDefault(),
        formulaEntryState.activeCellEditRef.current ||
          cellEditActions.beginCellEdit(currentTarget.value),
        (Be.mode = !Be.mode),
        formulaEntryState.setIsFormulaPointMode(Be.mode),
        true
      );
    if (event.key === "F4" && !event.ctrlKey && !event.metaKey) {
      const jt = $ae(currentTarget.value, Be.start, Be.end);
      return jt
        ? (event.preventDefault(),
          cellEditActions.beginCellEdit(
            jt.value,
            currentTarget.classList.contains("lg-cell-editor"),
          ),
          (Be.span = jt.span),
          (Be.start = Be.end = jt.span.end),
          formulaEntryState.setFormulaCaretPosition(jt.span.end),
          cellEditActions.restoreFormulaEditorFocus(),
          true)
        : false;
    }
    if (
      (event.ctrlKey || event.metaKey) &&
      ["PageUp", "PageDown"].includes(event.key)
    ) {
      event.preventDefault();
      const jt = activeWorkbook.sheets.filter((sheet) => !sheet.hidden),
        rr = jt.findIndex((Br) => Br.id === activeSheet.id),
        Mr = jt[rr + (event.key === "PageUp" ? -1 : 1)];
      return (Mr && navigationActions.switchActiveWorksheet(Mr.id), true);
    }
    if (
      !Be.mode ||
      !event.key.startsWith("Arrow") ||
      !formulaEntryState.activeCellEditRef.current ||
      !(
        Be.span ??
        findFormulaReferenceSpan(currentTarget.value, Be.start, Be.end)
      )
    )
      return false;
    event.preventDefault();
    const Ue = _t(
        Be.sheetId === activeSheet.id && Be.head
          ? Be.head
          : selectionState.activeCellAddress,
      ),
      Xe = ["ArrowUp", "ArrowDown"].includes(event.key),
      gt = ["ArrowUp", "ArrowLeft"].includes(event.key) ? -1 : 1,
      Nt =
        event.ctrlKey || event.metaKey
          ? viewAndSelectionActions.jumpToDataEdge(Ue.row, Ue.col, Xe, gt)
          : {
              row: Xe
                ? viewAndSelectionActions.stepVisibleRowIndex(Ue.row, gt)
                : Ue.row,
              col: Xe
                ? Ue.col
                : viewAndSelectionActions.stepVisibleColumnIndex(Ue.col, gt),
            },
      wt = `${columnIndexToLetters(Nt.col)}${Nt.row + 1}`;
    return (
      insertFormulaReference(wt, event.shiftKey),
      viewAndSelectionActions.requestCellScroll(wt),
      true
    );
  }
  return {
    insertFormulaReference,
    handleFormulaPointerDown,
    handleFormulaKeyDown,
  };
}

export { createFormulaEntryActions };
