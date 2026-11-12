import {
  DataValidationError,
  applyCellPatchTransaction,
  lD,
} from "../03-workbook-model.js";
import {
  autoCloseFormulaParens,
  parseLocalizedCellEntry,
} from "../04-cell-input-paste-and-fill.js";

function createCellEditActions({
  viewAndSelectionActions,
  persistenceActions,
  persistenceState,
  formulaEntryState,
  selectionState,
  navigationState,
  filterSortState,
  operationStatusState,
  actor,
  readOnly,
  workbookMessages,
  getEditableCellText,
  setActiveSheetId,
  activeSheet,
  workbookCalculator,
  latestEditStateRef,
}) {
  function clearSelectedCellContents() {
    applyCellPatches(
      viewAndSelectionActions.getSelectedCellAddresses().map((address) => ({
        sheetId: activeSheet.id,
        address: address,
        cell: { ...activeSheet.cells[address], raw: null, literal: false },
      })),
      "Clear contents",
    );
  }
  function applyCellPatches(
    patches,
    action = "Edit cells",
    selectionGuard,
    expectedSheetId,
  ) {
    try {
      if (
        expectedSheetId &&
        patches.some((patch) => patch.sheetId !== expectedSheetId)
      )
        throw Error("The formula destination changed.");
      persistenceActions.commitWorkbookChange(
        applyCellPatchTransaction(
          latestEditStateRef.current.book,
          patches,
          actor,
          action,
        ),
        true,
        false,
        undefined,
        true,
        expectedSheetId,
      );
    } catch (error) {
      throw (
        error instanceof DataValidationError &&
          formulaEntryState.setValidationReviewState({
            error: error,
            patches: structuredClone(patches),
            action: action,
            snapshot: JSON.stringify(latestEditStateRef.current.book),
            selectionGuard: selectionGuard,
          }),
        error
      );
    }
  }
  function resetFormulaEditSession() {
    ((formulaEntryState.formulaPointerStateRef.current = {
      input: null,
      start: 0,
      end: 0,
      dragging: false,
      mode: true,
    }),
      formulaEntryState.setActiveFormulaReference(null),
      formulaEntryState.setIsFormulaPointMode(true));
  }
  function syncFormulaEditorSelection(inputElement, shouldForceReset = false) {
    const pointerState = formulaEntryState.formulaPointerStateRef.current,
      selectionStart = inputElement.selectionStart ?? inputElement.value.length,
      selectionEnd = inputElement.selectionEnd ?? selectionStart;
    ((pointerState.input = inputElement),
      !(pointerState.restoring && !shouldForceReset) &&
        ((pointerState.restoring = false),
        (pointerState.start = selectionStart),
        (pointerState.end = selectionEnd),
        (shouldForceReset ||
          (pointerState.span &&
            (selectionStart !== pointerState.span.end ||
              selectionEnd !== selectionStart))) &&
          ((pointerState.span = undefined),
          (pointerState.anchor = undefined),
          (pointerState.head = undefined),
          formulaEntryState.setActiveFormulaReference(null))));
  }
  function restoreFormulaEditorFocus() {
    ((formulaEntryState.formulaPointerStateRef.current.restoring = true),
      requestAnimationFrame(() => {
        if (!formulaEntryState.activeCellEditRef.current) return;
        const pointerState = formulaEntryState.formulaPointerStateRef.current,
          inputElement = pointerState.input?.isConnected
            ? pointerState.input
            : formulaEntryState.formulaInputRef.current;
        (inputElement &&
          ((pointerState.input = inputElement),
          inputElement.focus({ preventScroll: true }),
          inputElement.setSelectionRange(pointerState.start, pointerState.end)),
          (pointerState.restoring = false));
      }));
  }
  function isFormulaEditingLocked() {
    const sheet = latestEditStateRef.current.book.sheets.find(
      (sheet) =>
        sheet.id ===
        (formulaEntryState.activeCellEditRef.current?.sheetId ??
          activeSheet.id),
    );
    return readOnly || !sheet || !!sheet.protection?.enabled;
  }
  function switchToEditSheet(cellEdit) {
    cellEdit.sheetId !== activeSheet.id &&
      (setActiveSheetId(cellEdit.sheetId),
      selectionState.setActiveCellAddress(cellEdit.address),
      selectionState.setSelectionAnchorAddress(cellEdit.address),
      viewAndSelectionActions.selectSingleRange(cellEdit.address),
      filterSortState.setActiveColumnSort(null),
      filterSortState.setFilterValueSearch(""),
      navigationState.setPendingScrollAddress(cellEdit.address));
  }
  function beginCellEdit(editText, shouldSetCellEditState = false) {
    const spillAnchorAddress = formulaEntryState.activeCellEditRef.current
      ? undefined
      : workbookCalculator.getSpillAnchor(
          activeSheet.id,
          selectionState.activeCellAddress,
        );
    if (
      spillAnchorAddress &&
      spillAnchorAddress !== selectionState.activeCellAddress
    ) {
      operationStatusState.setErrorNotice(
        `This cell belongs to the array formula in ${spillAnchorAddress}. Edit that cell instead.`,
      );
      return;
    }
    (persistenceState.saveAbortControllerRef.current?.abort(),
      (formulaEntryState.activeCellEditRef.current = {
        sheetId:
          formulaEntryState.activeCellEditRef.current?.sheetId ??
          activeSheet.id,
        address:
          formulaEntryState.activeCellEditRef.current?.address ??
          selectionState.activeCellAddress,
        value: editText,
        locale: formulaEntryState.activeCellEditRef.current
          ? formulaEntryState.activeCellEditRef.current.locale
          : workbookMessages.entryLocale,
        originalText:
          formulaEntryState.activeCellEditRef.current?.originalText ??
          getEditableCellText(
            activeSheet.cells[selectionState.activeCellAddress],
          ),
      }),
      formulaEntryState.setCellEditText(editText),
      shouldSetCellEditState && formulaEntryState.setCellEditState(editText));
  }
  function cancelCellEdit() {
    const activeCellEdit = formulaEntryState.activeCellEditRef.current;
    (activeCellEdit && switchToEditSheet(activeCellEdit),
      resetFormulaEditSession(),
      (formulaEntryState.activeCellEditRef.current = null),
      formulaEntryState.setCellEditState(null),
      formulaEntryState.setCellEditText(
        activeCellEdit?.originalText ??
          getEditableCellText(
            activeSheet.cells[selectionState.activeCellAddress],
          ),
      ),
      viewAndSelectionActions.focusGridSurface());
  }
  function commitCellEdit(fe) {
    const activeCellEdit = formulaEntryState.activeCellEditRef.current;
    if (!activeCellEdit) return true;
    try {
      const entryText = autoCloseFormulaParens(activeCellEdit.value),
        cell = latestEditStateRef.current.book.sheets.find(
          (sheet) => sheet.id === activeCellEdit.sheetId,
        )?.cells[activeCellEdit.address];
      return (
        entryText !== activeCellEdit.originalText &&
          applyCellPatches(
            [
              {
                sheetId: activeCellEdit.sheetId,
                address: activeCellEdit.address,
                cell: parseLocalizedCellEntry(
                  entryText,
                  cell,
                  latestEditStateRef.current.book.datePolicy?.dateSystem,
                  activeCellEdit.locale,
                ),
              },
            ],
            "Edit cells",
            undefined,
            activeCellEdit.sheetId,
          ),
        (navigationState.pendingCrossSheetTargetRef.current =
          activeCellEdit.sheetId !== activeSheet.id
            ? {
                sheetId: activeCellEdit.sheetId,
                address: activeCellEdit.address,
              }
            : null),
        switchToEditSheet(activeCellEdit),
        resetFormulaEditSession(),
        (formulaEntryState.activeCellEditRef.current = null),
        formulaEntryState.setCellEditState(null),
        formulaEntryState.setCellEditText(entryText),
        true
      );
    } catch (error) {
      return (
        operationStatusState.setErrorNotice(
          error instanceof Error
            ? error.message
            : "The cell could not be saved.",
        ),
        false
      );
    }
  }
  function applyCellFormatPatch(formatPatch) {
    viewAndSelectionActions.runGuardedAction(() => {
      const selectedAddresses =
        viewAndSelectionActions.getSelectedCellAddresses();
      if (
        formatPatch.style &&
        (Object.hasOwn(formatPatch.style, "border") ||
          Object.hasOwn(formatPatch.style, "borders")) &&
        selectedAddresses.some((address) =>
          (activeSheet.merges ?? []).some((mergeRange) =>
            lD(mergeRange, address),
          ),
        )
      )
        throw Error("Unmerge selected cells before changing border edges.");
      applyCellPatches(
        selectedAddresses.map((address) => ({
          sheetId: activeSheet.id,
          address: address,
          cell: {
            ...activeSheet.cells[address],
            raw: activeSheet.cells[address]?.raw ?? null,
            ...formatPatch,
            ...(formatPatch.style
              ? {
                  style: {
                    ...activeSheet.cells[address]?.style,
                    ...formatPatch.style,
                  },
                }
              : {}),
          },
        })),
        "Format selection",
      );
    });
  }
  return {
    clearSelectedCellContents,
    applyCellPatches,
    resetFormulaEditSession,
    syncFormulaEditorSelection,
    restoreFormulaEditorFocus,
    isFormulaEditingLocked,
    beginCellEdit,
    cancelCellEdit,
    commitCellEdit,
    applyCellFormatPatch,
  };
}

export { createCellEditActions };
