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
  selectSingleRange,
  workbookCalculator,
  latestEditStateRef,
  getSelectedCellAddresses,
  commitWorkbookChange,
  runGuardedAction,
  focusGridSurface,
}) {
  function clearSelectedCellContents() {
    applyCellPatches(
      getSelectedCellAddresses().map((fe) => ({
        sheetId: activeSheet.id,
        address: fe,
        cell: { ...activeSheet.cells[fe], raw: null, literal: false },
      })),
      "Clear contents",
    );
  }
  function applyCellPatches(fe, Oe = "Edit cells", Be, Ue) {
    try {
      if (Ue && fe.some((Xe) => Xe.sheetId !== Ue))
        throw Error("The formula destination changed.");
      commitWorkbookChange(
        applyCellPatchTransaction(
          latestEditStateRef.current.book,
          fe,
          actor,
          Oe,
        ),
        true,
        false,
        undefined,
        true,
        Ue,
      );
    } catch (Xe) {
      throw (
        Xe instanceof DataValidationError &&
          formulaEntryState.setValidationReviewState({
            error: Xe,
            patches: structuredClone(fe),
            action: Oe,
            snapshot: JSON.stringify(latestEditStateRef.current.book),
            selectionGuard: Be,
          }),
        Xe
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
  function syncFormulaEditorSelection(fe, Oe = false) {
    const Be = formulaEntryState.formulaPointerStateRef.current,
      Ue = fe.selectionStart ?? fe.value.length,
      Xe = fe.selectionEnd ?? Ue;
    ((Be.input = fe),
      !(Be.restoring && !Oe) &&
        ((Be.restoring = false),
        (Be.start = Ue),
        (Be.end = Xe),
        (Oe || (Be.span && (Ue !== Be.span.end || Xe !== Ue))) &&
          ((Be.span = undefined),
          (Be.anchor = undefined),
          (Be.head = undefined),
          formulaEntryState.setActiveFormulaReference(null))));
  }
  function restoreFormulaEditorFocus() {
    ((formulaEntryState.formulaPointerStateRef.current.restoring = true),
      requestAnimationFrame(() => {
        if (!formulaEntryState.activeCellEditRef.current) return;
        const fe = formulaEntryState.formulaPointerStateRef.current,
          Oe = fe.input?.isConnected
            ? fe.input
            : formulaEntryState.formulaInputRef.current;
        (Oe &&
          ((fe.input = Oe),
          Oe.focus({ preventScroll: true }),
          Oe.setSelectionRange(fe.start, fe.end)),
          (fe.restoring = false));
      }));
  }
  function isFormulaEditingLocked() {
    const fe = latestEditStateRef.current.book.sheets.find(
      (Be) =>
        Be.id ===
        (formulaEntryState.activeCellEditRef.current?.sheetId ??
          activeSheet.id),
    );
    return readOnly || !fe || !!fe.protection?.enabled;
  }
  function switchToEditSheet(fe) {
    fe.sheetId !== activeSheet.id &&
      (setActiveSheetId(fe.sheetId),
      selectionState.setActiveCellAddress(fe.address),
      selectionState.setSelectionAnchorAddress(fe.address),
      selectSingleRange(fe.address),
      filterSortState.setActiveColumnSort(null),
      filterSortState.setFilterValueSearch(""),
      navigationState.setPendingScrollAddress(fe.address));
  }
  function beginCellEdit(fe, Oe = false) {
    const Be = formulaEntryState.activeCellEditRef.current
      ? undefined
      : workbookCalculator.getSpillAnchor(
          activeSheet.id,
          selectionState.activeCellAddress,
        );
    if (Be && Be !== selectionState.activeCellAddress) {
      operationStatusState.setErrorNotice(
        `This cell belongs to the array formula in ${Be}. Edit that cell instead.`,
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
        value: fe,
        locale: formulaEntryState.activeCellEditRef.current
          ? formulaEntryState.activeCellEditRef.current.locale
          : workbookMessages.entryLocale,
        originalText:
          formulaEntryState.activeCellEditRef.current?.originalText ??
          getEditableCellText(
            activeSheet.cells[selectionState.activeCellAddress],
          ),
      }),
      formulaEntryState.setCellEditText(fe),
      Oe && formulaEntryState.setCellEditState(fe));
  }
  function cancelCellEdit() {
    const fe = formulaEntryState.activeCellEditRef.current;
    (fe && switchToEditSheet(fe),
      resetFormulaEditSession(),
      (formulaEntryState.activeCellEditRef.current = null),
      formulaEntryState.setCellEditState(null),
      formulaEntryState.setCellEditText(
        fe?.originalText ??
          getEditableCellText(
            activeSheet.cells[selectionState.activeCellAddress],
          ),
      ),
      focusGridSurface());
  }
  function commitCellEdit(fe) {
    const Oe = formulaEntryState.activeCellEditRef.current;
    if (!Oe) return true;
    try {
      const Xe = autoCloseFormulaParens(Oe.value),
        gt = latestEditStateRef.current.book.sheets.find(
          (Nt) => Nt.id === Oe.sheetId,
        )?.cells[Oe.address];
      return (
        Xe !== Oe.originalText &&
          applyCellPatches(
            [
              {
                sheetId: Oe.sheetId,
                address: Oe.address,
                cell: parseLocalizedCellEntry(
                  Xe,
                  gt,
                  latestEditStateRef.current.book.datePolicy?.dateSystem,
                  Oe.locale,
                ),
              },
            ],
            "Edit cells",
            undefined,
            Oe.sheetId,
          ),
        (navigationState.pendingCrossSheetTargetRef.current =
          Oe.sheetId !== activeSheet.id
            ? { sheetId: Oe.sheetId, address: Oe.address }
            : null),
        switchToEditSheet(Oe),
        resetFormulaEditSession(),
        (formulaEntryState.activeCellEditRef.current = null),
        formulaEntryState.setCellEditState(null),
        formulaEntryState.setCellEditText(Xe),
        true
      );
    } catch (Xe) {
      return (
        operationStatusState.setErrorNotice(
          Xe instanceof Error ? Xe.message : "The cell could not be saved.",
        ),
        false
      );
    }
  }
  function applyCellFormatPatch(fe) {
    runGuardedAction(() => {
      const Oe = getSelectedCellAddresses();
      if (
        fe.style &&
        (Object.hasOwn(fe.style, "border") ||
          Object.hasOwn(fe.style, "borders")) &&
        Oe.some((Be) => (activeSheet.merges ?? []).some((Ue) => lD(Ue, Be)))
      )
        throw Error("Unmerge selected cells before changing border edges.");
      applyCellPatches(
        Oe.map((Be) => ({
          sheetId: activeSheet.id,
          address: Be,
          cell: {
            ...activeSheet.cells[Be],
            raw: activeSheet.cells[Be]?.raw ?? null,
            ...fe,
            ...(fe.style
              ? {
                  style: {
                    ...activeSheet.cells[Be]?.style,
                    ...fe.style,
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
