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
  function insertFormulaReference(
    targetAddress,
    isExtendingRange = false,
    rangeOverride,
  ) {
    const activeCellEdit = formulaEntryState.activeCellEditRef.current,
      formulaPointerState = formulaEntryState.formulaPointerStateRef.current;
    if (
      !activeCellEdit?.value.startsWith("=") ||
      cellEditActions.isFormulaEditingLocked() ||
      operationStatusState.isOperationBusy
    )
      return false;
    const referenceSpan =
      formulaPointerState.span ??
      findFormulaReferenceSpan(
        activeCellEdit.value,
        formulaPointerState.start,
        formulaPointerState.end,
      );
    if (!referenceSpan) return false;
    const rangeStartAddress =
        isExtendingRange &&
        formulaPointerState.sheetId === activeSheet.id &&
        formulaPointerState.anchor
          ? formulaPointerState.anchor
          : targetAddress,
      range =
        rangeOverride?.range ??
        (rangeStartAddress === targetAddress
          ? targetAddress
          : wy(rangeStartAddress, targetAddress)),
      replacement = replaceFormulaReferenceSpan(
        activeCellEdit.value,
        referenceSpan,
        range,
        activeCellEdit.sheetId === activeSheet.id
          ? undefined
          : activeSheet.name,
      );
    return (
      cellEditActions.beginCellEdit(
        replacement.value,
        !!formulaPointerState.input?.classList.contains("lg-cell-editor"),
      ),
      (formulaPointerState.span = replacement.span),
      (formulaPointerState.start = formulaPointerState.end =
        replacement.span.end),
      (formulaPointerState.anchor = rangeStartAddress),
      (formulaPointerState.head = targetAddress),
      (formulaPointerState.sheetId = activeSheet.id),
      formulaEntryState.setFormulaCaretPosition(replacement.span.end),
      formulaEntryState.setActiveFormulaReference({
        sheetId: activeSheet.id,
        range: range,
        bounds: rangeOverride?.bounds ?? parseRangeAddress(range),
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
    const addressElement = target.closest("[data-address]"),
      columnElement = target.closest("[data-column]"),
      rowHeaderElement = target.closest("[data-row-header]");
    let isHandled = false;
    if (addressElement)
      isHandled = insertFormulaReference(
        addressElement.dataset.address,
        event.shiftKey,
      );
    else if (columnElement) {
      const columnIndex = Number(columnElement.dataset.column);
      isHandled = insertFormulaReference(
        `${columnIndexToLetters(columnIndex)}1`,
        false,
        {
          range: `${columnIndexToLetters(columnIndex)}:${columnIndexToLetters(columnIndex)}`,
          bounds: {
            firstRow: 0,
            lastRow: activeSheet.rowCount - 1,
            firstCol: columnIndex,
            lastCol: columnIndex,
          },
        },
      );
    } else if (rowHeaderElement) {
      const rowIndex = Number(rowHeaderElement.dataset.rowHeader);
      isHandled = insertFormulaReference(`A${rowIndex + 1}`, false, {
        range: `${rowIndex + 1}:${rowIndex + 1}`,
        bounds: {
          firstRow: rowIndex,
          lastRow: rowIndex,
          firstCol: 0,
          lastCol: activeSheet.colCount - 1,
        },
      });
    }
    return isHandled
      ? (event.preventDefault(),
        event.stopPropagation(),
        (formulaEntryState.formulaPointerStateRef.current.dragging =
          !!addressElement),
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
    const formulaPointerState =
      formulaEntryState.formulaPointerStateRef.current;
    if (event.key === "F2" && !event.ctrlKey && !event.metaKey)
      return (
        event.preventDefault(),
        formulaEntryState.activeCellEditRef.current ||
          cellEditActions.beginCellEdit(currentTarget.value),
        (formulaPointerState.mode = !formulaPointerState.mode),
        formulaEntryState.setIsFormulaPointMode(formulaPointerState.mode),
        true
      );
    if (event.key === "F4" && !event.ctrlKey && !event.metaKey) {
      const replacement = $ae(
        currentTarget.value,
        formulaPointerState.start,
        formulaPointerState.end,
      );
      return replacement
        ? (event.preventDefault(),
          cellEditActions.beginCellEdit(
            replacement.value,
            currentTarget.classList.contains("lg-cell-editor"),
          ),
          (formulaPointerState.span = replacement.span),
          (formulaPointerState.start = formulaPointerState.end =
            replacement.span.end),
          formulaEntryState.setFormulaCaretPosition(replacement.span.end),
          cellEditActions.restoreFormulaEditorFocus(),
          true)
        : false;
    }
    if (
      (event.ctrlKey || event.metaKey) &&
      ["PageUp", "PageDown"].includes(event.key)
    ) {
      event.preventDefault();
      const visibleSheets = activeWorkbook.sheets.filter(
          (sheet) => !sheet.hidden,
        ),
        activeSheetIndex = visibleSheets.findIndex(
          (sheet) => sheet.id === activeSheet.id,
        ),
        targetSheet =
          visibleSheets[activeSheetIndex + (event.key === "PageUp" ? -1 : 1)];
      return (
        targetSheet && navigationActions.switchActiveWorksheet(targetSheet.id),
        true
      );
    }
    if (
      !formulaPointerState.mode ||
      !event.key.startsWith("Arrow") ||
      !formulaEntryState.activeCellEditRef.current ||
      !(
        formulaPointerState.span ??
        findFormulaReferenceSpan(
          currentTarget.value,
          formulaPointerState.start,
          formulaPointerState.end,
        )
      )
    )
      return false;
    event.preventDefault();
    const activeCellPosition = _t(
        formulaPointerState.sheetId === activeSheet.id &&
          formulaPointerState.head
          ? formulaPointerState.head
          : selectionState.activeCellAddress,
      ),
      isVerticalArrow = ["ArrowUp", "ArrowDown"].includes(event.key),
      stepDirection = ["ArrowUp", "ArrowLeft"].includes(event.key) ? -1 : 1,
      nextCellPosition =
        event.ctrlKey || event.metaKey
          ? viewAndSelectionActions.jumpToDataEdge(
              activeCellPosition.row,
              activeCellPosition.col,
              isVerticalArrow,
              stepDirection,
            )
          : {
              row: isVerticalArrow
                ? viewAndSelectionActions.stepVisibleRowIndex(
                    activeCellPosition.row,
                    stepDirection,
                  )
                : activeCellPosition.row,
              col: isVerticalArrow
                ? activeCellPosition.col
                : viewAndSelectionActions.stepVisibleColumnIndex(
                    activeCellPosition.col,
                    stepDirection,
                  ),
            },
      nextAddress = `${columnIndexToLetters(nextCellPosition.col)}${nextCellPosition.row + 1}`;
    return (
      insertFormulaReference(nextAddress, event.shiftKey),
      viewAndSelectionActions.requestCellScroll(nextAddress),
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
