import { buildSourceImportRecord } from "../09-workbook-links.js";
import {
  _t,
  addWorksheet,
  columnIndexToLetters,
  enumerateSelectionCells,
  findOccupiedRangeBounds,
  parseRangeAddress,
  renameWorkbook,
  resolveEffectiveCellStyle,
} from "../03-workbook-model.js";
import {
  applyCellPreview,
  buildFillPreview,
} from "../04-cell-input-paste-and-fill.js";
import { Tp } from "../53-find-replace-go-to.js";

function createGridInteractionActions({
  getEditableCellText,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  setActiveDialogMode,
  setAutomationInspection,
  setContextMenuState,
  appRootRef,
  setTabBarSheetId,
  activeCellAddress,
  setActiveCellAddress,
  setSelectionAnchorAddress,
  activeCellAddressRef,
  activeSelectionRange,
  completedSelectionAreas,
  selectSingleRange,
  selectionAreas,
  requireSingleSelectionRange,
  activeCellEditRef,
  formulaPointerStateRef,
  setIsFormulaPointMode,
  filterValueSearch,
  setActiveFilterColumn,
  activeColumnSort,
  splitPaneSizes,
  setStatusNotice,
  setErrorNotice,
  isOperationBusy,
  setXlsxConversionReview,
  promptDialogState,
  setPromptDialogState,
  setIsFindBarVisible,
  clipboardSnapshotState,
  setMoveCellsSource,
  clearCutSession,
  rowVisibilityIndex,
  columnVisibilityIndex,
  defaultRowHeight,
  zoomScale,
  gridContainerRef,
  findInputRef,
  runCalculationNow,
  mergeRanges,
  activeCell,
  rowViewport,
  getSelectedCellAddresses,
  commitWorkbookEdit,
  replaceWorkbookSnapshot,
  runGuardedAction,
  applyCellPatches,
  beginCellEdit,
  commitCellEdit,
  requestCellScroll,
  selectCellAddress,
  switchActiveWorksheet,
  applyCellFormatPatch,
  undoOrRedo,
  stepVisibleRowIndex,
  stepVisibleColumnIndex,
  jumpToDataEdge,
  selectAllCells,
  selectRowOrColumn,
}) {
  function runEditorCommand(fe) {
    if (fe === "xlsxSource") {
      runGuardedAction(() =>
        setXlsxConversionReview({
          snapshot: activeWorkbook,
          report: buildSourceImportRecord(activeWorkbook),
          fileName: activeWorkbook.xlsxSource.fileName,
          inspectOnly: true,
        }),
      );
      return;
    }
    if (commitCellEdit()) {
      if (
        (fe === "populationFilter" && setActiveFilterColumn(undefined),
        selectionAreas.length > 1 &&
          ![
            "search",
            "numberFormat",
            "cellStyle",
            "themes",
            "goToSpecial",
            "comments",
            "hyperlink",
            "sheets",
            "workbookProperties",
            "versions",
            "tracePrecedents",
            "traceDependents",
            "formulaWatches",
            "evaluateFormula",
            "functionCompatibility",
            "calculation",
            "calculateNow",
            "automation",
            "externalSources",
            "workbookFormats",
            "importSources",
            "workbookLinks",
            "extensions",
          ].includes(fe))
      ) {
        setErrorNotice(
          "This command requires one rectangular range. Click a cell or enter a single range to continue.",
        );
        return;
      }
      if (fe === "calculateNow") {
        runCalculationNow();
        return;
      }
      if (fe === "pasteSpecial" && clipboardSnapshotState) {
        setMoveCellsSource(clipboardSnapshotState.snapshot);
        return;
      }
      (fe === "sheets" && setTabBarSheetId(activeSheet.id),
        fe === "automation" && setAutomationInspection(undefined),
        setActiveDialogMode(fe));
    }
  }
  function handleGridKeyDown(fe) {
    var wt;
    if (
      fe.target !== fe.currentTarget ||
      fe.defaultPrevented ||
      fe.nativeEvent.isComposing ||
      (fe.altKey &&
        !((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "v")) ||
      fe.getModifierState("AltGraph")
    )
      return;
    const Oe = fe.ctrlKey || fe.metaKey,
      Be = fe.key.toLowerCase();
    if (
      (Oe &&
        ["tab", "enter", "f2", "f5", "f6", "delete", "backspace"].includes(
          Be,
        )) ||
      (Oe &&
        fe.shiftKey &&
        ![
          "z",
          "f",
          "home",
          "end",
          " ",
          "arrowleft",
          "arrowright",
          "arrowup",
          "arrowdown",
        ].includes(Be))
    )
      return;
    if (
      isOperationBusy &&
      ((Oe && ["z", "y", "b", "i", "u", "d", "r", "v"].includes(Be)) ||
        (!Oe &&
          (fe.key.length === 1 ||
            ["F2", "Delete", "Backspace"].includes(fe.key))))
    ) {
      (fe.preventDefault(),
        setStatusNotice("Wait for the current workbook operation to finish."));
      return;
    }
    if (
      (fe.key === "ContextMenu" && !fe.shiftKey && !Oe) ||
      (fe.key === "F10" && fe.shiftKey && !fe.ctrlKey && !fe.metaKey)
    ) {
      fe.preventDefault();
      const rr = (
        fe.currentTarget.querySelector(
          `[data-address="${activeCellAddress}"]`,
        ) ?? fe.currentTarget
      ).getBoundingClientRect();
      setContextMenuState({
        x: rr.left + Math.min(40, rr.width),
        y: rr.top + Math.min(24, rr.height),
        kind:
          parseRangeAddress(activeSelectionRange).lastRow -
            parseRangeAddress(activeSelectionRange).firstRow +
            1 ===
          activeSheet.rowCount
            ? "column"
            : parseRangeAddress(activeSelectionRange).lastCol -
                  parseRangeAddress(activeSelectionRange).firstCol +
                  1 ===
                activeSheet.colCount
              ? "row"
              : "cell",
      });
      return;
    }
    if (fe.key === "Escape" && clipboardSnapshotState) {
      (fe.preventDefault(),
        clearCutSession(),
        setStatusNotice("Cut canceled; source cells remain unchanged."));
      return;
    }
    if (fe.key === "Escape" && completedSelectionAreas.length) {
      (fe.preventDefault(),
        selectSingleRange(activeCellAddress),
        setSelectionAnchorAddress(activeCellAddress),
        (activeCellAddressRef.current = activeCellAddress));
      return;
    }
    if (fe.key === "Escape" && !fe.ctrlKey && !fe.metaKey && !fe.shiftKey) {
      (fe.preventDefault(),
        (wt = appRootRef.current?.querySelector(
          '[role="tab"][aria-selected="true"]',
        )) == null || wt.focus());
      return;
    }
    if (
      (fe.key === "F5" && !fe.shiftKey && !Oe) ||
      ((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "g")
    ) {
      (fe.preventDefault(), runEditorCommand("goToSpecial"));
      return;
    }
    if (
      (fe.ctrlKey || fe.metaKey) &&
      fe.shiftKey &&
      fe.key.toLowerCase() === "f"
    ) {
      (fe.preventDefault(), runEditorCommand("search"));
      return;
    }
    if (completedSelectionAreas.length && ["Tab", "Enter"].includes(fe.key)) {
      (fe.preventDefault(),
        runGuardedAction(() => {
          const jt = enumerateSelectionCells(selectionAreas).filter((xr) => {
            const Sr = _t(xr);
            return (
              !rowVisibilityIndex.isHidden(Sr.row) &&
              !columnVisibilityIndex.isHidden(Sr.col)
            );
          });
          if (!jt.length) return;
          const rr = jt.indexOf(activeCellAddress),
            Mr = fe.shiftKey ? -1 : 1,
            Br = jt[(rr + Mr + jt.length) % jt.length];
          commitCellEdit() &&
            (setActiveCellAddress(Br),
            setSelectionAnchorAddress(Br),
            (activeCellAddressRef.current = Br),
            requestCellScroll(Br));
        }));
      return;
    }
    const Ue = _t(
      fe.shiftKey && fe.key.startsWith("Arrow")
        ? activeCellAddressRef.current
        : activeCellAddress,
    );
    let Xe = Ue.row,
      gt = Ue.col;
    if ((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "z") {
      (fe.preventDefault(), undoOrRedo(fe.shiftKey));
      return;
    }
    if (fe.ctrlKey || fe.metaKey) {
      const jt = fe.key.toLowerCase();
      if (jt === "1") {
        (fe.preventDefault(), runEditorCommand("numberFormat"));
        return;
      }
      if (jt === "v" && fe.altKey) {
        (fe.preventDefault(), runEditorCommand("pasteSpecial"));
        return;
      }
      if (jt === " ") {
        (fe.preventDefault(), selectRowOrColumn("column", Ue.col, fe.shiftKey));
        return;
      }
      if (jt === "pageup" || jt === "pagedown") {
        fe.preventDefault();
        const rr = activeWorkbook.sheets.filter((xr) => !xr.hidden),
          Mr = rr.findIndex((xr) => xr.id === activeSheet.id),
          Br = rr[Mr + (jt === "pageup" ? -1 : 1)];
        Br && commitCellEdit() && switchActiveWorksheet(Br.id);
        return;
      }
      if (jt.startsWith("arrow")) {
        fe.preventDefault();
        const rr = jumpToDataEdge(
          Ue.row,
          Ue.col,
          jt === "arrowup" || jt === "arrowdown",
          jt === "arrowup" || jt === "arrowleft" ? -1 : 1,
        );
        selectCellAddress(
          `${columnIndexToLetters(rr.col)}${rr.row + 1}`,
          fe.shiftKey,
        );
        return;
      }
      if (jt === "a") {
        (fe.preventDefault(), selectAllCells());
        return;
      }
      if (["b", "i", "u"].includes(jt)) {
        if ((fe.preventDefault(), !isEditingLocked)) {
          const rr = jt === "b" ? "bold" : jt === "i" ? "italic" : "underline";
          applyCellFormatPatch({
            style: { [rr]: !resolveEffectiveCellStyle(activeCell?.style)[rr] },
          });
        }
        return;
      }
      if (jt === "d") {
        (fe.preventDefault(),
          isEditingLocked ||
            runGuardedAction(() => {
              (requireSingleSelectionRange(),
                commitWorkbookEdit(
                  applyCellPreview(
                    activeWorkbook,
                    buildFillPreview(
                      activeWorkbook,
                      activeSheet.id,
                      activeSelectionRange,
                      "down",
                    ),
                  ),
                  "Fill down",
                ));
            }));
        return;
      }
      if (jt === "r") {
        (fe.preventDefault(),
          isEditingLocked ||
            runGuardedAction(() => {
              (requireSingleSelectionRange(),
                commitWorkbookEdit(
                  applyCellPreview(
                    activeWorkbook,
                    buildFillPreview(
                      activeWorkbook,
                      activeSheet.id,
                      activeSelectionRange,
                      "right",
                    ),
                  ),
                  "Fill right",
                ));
            }));
        return;
      }
      if (jt === "y") {
        (fe.preventDefault(), undoOrRedo(true));
        return;
      }
      if (jt === "f" || jt === "h") {
        (fe.preventDefault(),
          setIsFindBarVisible(true),
          requestAnimationFrame(() => {
            return findInputRef.current?.focus();
          }));
        return;
      }
      if (jt === "home") {
        (fe.preventDefault(),
          selectCellAddress(
            `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`,
            fe.shiftKey,
          ));
        return;
      }
      if (jt === "end") {
        fe.preventDefault();
        const rr = findOccupiedRangeBounds(activeSheet);
        selectCellAddress(
          `${columnIndexToLetters(columnVisibilityIndex.next(rr?.lastCol ?? 0, -1, true) ?? columnVisibilityIndex.logicalIndex(0))}${(rowVisibilityIndex.next(rr?.lastRow ?? 0, -1, true) ?? rowVisibilityIndex.logicalIndex(0)) + 1}`,
          fe.shiftKey,
        );
        return;
      }
      return;
    }
    if (fe.shiftKey && fe.key === " ") {
      (fe.preventDefault(), selectRowOrColumn("row", Ue.row, false));
      return;
    }
    if (["Home", "PageUp", "PageDown"].includes(fe.key)) {
      (fe.preventDefault(),
        fe.key === "Home"
          ? (gt = columnVisibilityIndex.logicalIndex(0))
          : (Xe = stepVisibleRowIndex(
              Xe,
              Math.max(
                1,
                Math.floor(rowViewport / zoomScale / defaultRowHeight) - 1,
              ) * (fe.key === "PageUp" ? -1 : 1),
            )),
        selectCellAddress(`${columnIndexToLetters(gt)}${Xe + 1}`, fe.shiftKey));
      return;
    }
    if (
      [
        "ArrowDown",
        "ArrowUp",
        "ArrowLeft",
        "ArrowRight",
        "Tab",
        "Enter",
      ].includes(fe.key)
    ) {
      if (
        (fe.preventDefault(),
        fe.key === "ArrowDown" && (Xe = stepVisibleRowIndex(Xe, 1)),
        fe.key === "Enter" &&
          (Xe = stepVisibleRowIndex(Xe, fe.shiftKey ? -1 : 1)),
        fe.key === "ArrowUp" && (Xe = stepVisibleRowIndex(Xe, -1)),
        fe.key === "ArrowLeft" && (gt = stepVisibleColumnIndex(gt, -1)),
        (fe.key === "ArrowRight" || fe.key === "Tab") &&
          (gt = stepVisibleColumnIndex(
            gt,
            fe.shiftKey && fe.key === "Tab" ? -1 : 1,
          )),
        !fe.shiftKey && !filterValueSearch && !activeColumnSort)
      ) {
        const rr = mergeRanges.find((Mr) => Tp(Mr, Ue.row, Ue.col));
        rr &&
          (gt > Ue.col &&
            (gt = columnVisibilityIndex.next(rr.lastCol, 1) ?? Ue.col),
          gt < Ue.col &&
            (gt = columnVisibilityIndex.next(rr.firstCol, -1) ?? Ue.col),
          Xe > Ue.row &&
            (Xe = rowVisibilityIndex.next(rr.lastRow, 1) ?? Ue.row),
          Xe < Ue.row &&
            (Xe = rowVisibilityIndex.next(rr.firstRow, -1) ?? Ue.row));
      }
      const jt = `${columnIndexToLetters(Math.max(0, Math.min(activeSheet.colCount - 1, gt)))}${Math.max(1, Math.min(activeSheet.rowCount, Xe + 1))}`;
      selectCellAddress(jt, fe.shiftKey && !["Tab", "Enter"].includes(fe.key));
      return;
    }
    (!isEditingLocked &&
      (fe.key === "Delete" || fe.key === "Backspace") &&
      (fe.preventDefault(),
      runGuardedAction(() =>
        applyCellPatches(
          getSelectedCellAddresses().map((jt) => ({
            sheetId: activeSheet.id,
            address: jt,
            cell: { ...activeSheet.cells[jt], raw: null },
          })),
          "Clear selection",
        ),
      )),
      !isEditingLocked &&
        (fe.key === "F2" ||
          (fe.key.length === 1 && !fe.ctrlKey && !fe.metaKey && !fe.altKey)) &&
        (fe.preventDefault(),
        (formulaPointerStateRef.current.mode = fe.key !== "F2"),
        setIsFormulaPointMode(formulaPointerStateRef.current.mode),
        beginCellEdit(
          fe.key === "F2" ? getEditableCellText(activeCell) : fe.key,
          true,
        )));
  }
  function commitDialogInput() {
    promptDialogState &&
      runGuardedAction(() => {
        if (promptDialogState.kind === "note")
          applyCellFormatPatch({ note: promptDialogState.value });
        else if (promptDialogState.kind === "rename")
          replaceWorkbookSnapshot(
            renameWorkbook(activeWorkbook, promptDialogState.value),
            "Rename workbook",
          );
        else {
          const fe = addWorksheet(activeWorkbook, promptDialogState.value),
            Oe = fe.sheets.at(-1);
          (replaceWorkbookSnapshot(fe, `Add sheet: ${Oe.name}`),
            switchActiveWorksheet(Oe.id));
        }
        setPromptDialogState(null);
      });
  }
  function activateGridPane(fe) {
    var Oe;
    return splitPaneSizes.active.current === fe
      ? true
      : !(
            (Oe = activeCellEditRef.current) != null && Oe.value.startsWith("=")
          ) && !commitCellEdit()
        ? false
        : (splitPaneSizes.activate(fe),
          (gridContainerRef.current = splitPaneSizes.refs.current[fe] ?? null),
          true);
  }
  return {
    runEditorCommand,
    handleGridKeyDown,
    commitDialogInput,
    activateGridPane,
  };
}

export { createGridInteractionActions };
