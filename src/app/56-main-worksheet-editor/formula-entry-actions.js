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
  formulaEntryState,
  selectionState,
  operationStatusState,
  dragFillState,
  activeWorkbook,
  activeSheet,
  syncFormulaEditorSelection,
  restoreFormulaEditorFocus,
  isFormulaEditingLocked,
  beginCellEdit,
  requestCellScroll,
  switchActiveWorksheet,
  stepVisibleRowIndex,
  stepVisibleColumnIndex,
  jumpToDataEdge,
}) {
  function insertFormulaReference(fe, Oe = false, Be) {
    const Ue = formulaEntryState.activeCellEditRef.current,
      Xe = formulaEntryState.formulaPointerStateRef.current;
    if (
      !Ue?.value.startsWith("=") ||
      isFormulaEditingLocked() ||
      operationStatusState.isOperationBusy
    )
      return false;
    const gt = Xe.span ?? findFormulaReferenceSpan(Ue.value, Xe.start, Xe.end);
    if (!gt) return false;
    const Nt =
        Oe && Xe.sheetId === activeSheet.id && Xe.anchor ? Xe.anchor : fe,
      wt = Be?.range ?? (Nt === fe ? fe : wy(Nt, fe)),
      jt = replaceFormulaReferenceSpan(
        Ue.value,
        gt,
        wt,
        Ue.sheetId === activeSheet.id ? undefined : activeSheet.name,
      );
    return (
      beginCellEdit(jt.value, !!Xe.input?.classList.contains("lg-cell-editor")),
      (Xe.span = jt.span),
      (Xe.start = Xe.end = jt.span.end),
      (Xe.anchor = Nt),
      (Xe.head = fe),
      (Xe.sheetId = activeSheet.id),
      formulaEntryState.setFormulaCaretPosition(jt.span.end),
      formulaEntryState.setActiveFormulaReference({
        sheetId: activeSheet.id,
        range: wt,
        bounds: Be?.bounds ?? parseRangeAddress(wt),
      }),
      restoreFormulaEditorFocus(),
      true
    );
  }
  function handleFormulaPointerDown(fe) {
    if (
      fe.button !== 0 ||
      !formulaEntryState.activeCellEditRef.current?.value.startsWith("=")
    )
      return false;
    const Oe = fe.target;
    if (Oe.closest('input,button,a,[role="separator"]')) return false;
    const Be = Oe.closest("[data-address]"),
      Ue = Oe.closest("[data-column]"),
      Xe = Oe.closest("[data-row-header]");
    let gt = false;
    if (Be) gt = insertFormulaReference(Be.dataset.address, fe.shiftKey);
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
      ? (fe.preventDefault(),
        fe.stopPropagation(),
        (formulaEntryState.formulaPointerStateRef.current.dragging = !!Be),
        (dragFillState.dragAnchorRef.current = null),
        true)
      : false;
  }
  function handleFormulaKeyDown(fe) {
    if (
      fe.nativeEvent.isComposing ||
      fe.altKey ||
      isFormulaEditingLocked() ||
      operationStatusState.isOperationBusy
    )
      return false;
    const Oe = fe.currentTarget;
    if (!Oe.value.startsWith("=")) return false;
    syncFormulaEditorSelection(Oe);
    const Be = formulaEntryState.formulaPointerStateRef.current;
    if (fe.key === "F2" && !fe.ctrlKey && !fe.metaKey)
      return (
        fe.preventDefault(),
        formulaEntryState.activeCellEditRef.current || beginCellEdit(Oe.value),
        (Be.mode = !Be.mode),
        formulaEntryState.setIsFormulaPointMode(Be.mode),
        true
      );
    if (fe.key === "F4" && !fe.ctrlKey && !fe.metaKey) {
      const jt = $ae(Oe.value, Be.start, Be.end);
      return jt
        ? (fe.preventDefault(),
          beginCellEdit(jt.value, Oe.classList.contains("lg-cell-editor")),
          (Be.span = jt.span),
          (Be.start = Be.end = jt.span.end),
          formulaEntryState.setFormulaCaretPosition(jt.span.end),
          restoreFormulaEditorFocus(),
          true)
        : false;
    }
    if ((fe.ctrlKey || fe.metaKey) && ["PageUp", "PageDown"].includes(fe.key)) {
      fe.preventDefault();
      const jt = activeWorkbook.sheets.filter((Br) => !Br.hidden),
        rr = jt.findIndex((Br) => Br.id === activeSheet.id),
        Mr = jt[rr + (fe.key === "PageUp" ? -1 : 1)];
      return (Mr && switchActiveWorksheet(Mr.id), true);
    }
    if (
      !Be.mode ||
      !fe.key.startsWith("Arrow") ||
      !formulaEntryState.activeCellEditRef.current ||
      !(Be.span ?? findFormulaReferenceSpan(Oe.value, Be.start, Be.end))
    )
      return false;
    fe.preventDefault();
    const Ue = _t(
        Be.sheetId === activeSheet.id && Be.head
          ? Be.head
          : selectionState.activeCellAddress,
      ),
      Xe = ["ArrowUp", "ArrowDown"].includes(fe.key),
      gt = ["ArrowUp", "ArrowLeft"].includes(fe.key) ? -1 : 1,
      Nt =
        fe.ctrlKey || fe.metaKey
          ? jumpToDataEdge(Ue.row, Ue.col, Xe, gt)
          : {
              row: Xe ? stepVisibleRowIndex(Ue.row, gt) : Ue.row,
              col: Xe ? Ue.col : stepVisibleColumnIndex(Ue.col, gt),
            },
      wt = `${columnIndexToLetters(Nt.col)}${Nt.row + 1}`;
    return (
      insertFormulaReference(wt, fe.shiftKey),
      requestCellScroll(wt),
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
