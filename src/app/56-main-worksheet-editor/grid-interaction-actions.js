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
  viewAndSelectionActions,
  persistenceActions,
  navigationActions,
  cellEditActions,
  dialogState,
  selectionState,
  navigationState,
  formulaEntryState,
  filterSortState,
  operationStatusState,
  importExportState,
  findReplaceState,
  clipboardState,
  getEditableCellText,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  selectionAreas,
  splitPaneSizes,
  rowVisibilityIndex,
  columnVisibilityIndex,
  defaultRowHeight,
  zoomScale,
  mergeRanges,
  activeCell,
  rowViewport,
}) {
  function runEditorCommand(fe) {
    if (fe === "xlsxSource") {
      viewAndSelectionActions.runGuardedAction(() =>
        importExportState.setXlsxConversionReview({
          snapshot: activeWorkbook,
          report: buildSourceImportRecord(activeWorkbook),
          fileName: activeWorkbook.xlsxSource.fileName,
          inspectOnly: true,
        }),
      );
      return;
    }
    if (cellEditActions.commitCellEdit()) {
      if (
        (fe === "populationFilter" &&
          filterSortState.setActiveFilterColumn(undefined),
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
        operationStatusState.setErrorNotice(
          "This command requires one rectangular range. Click a cell or enter a single range to continue.",
        );
        return;
      }
      if (fe === "calculateNow") {
        persistenceActions.runCalculationNow();
        return;
      }
      if (fe === "pasteSpecial" && clipboardState.clipboardSnapshotState) {
        clipboardState.setMoveCellsSource(
          clipboardState.clipboardSnapshotState.snapshot,
        );
        return;
      }
      (fe === "sheets" && navigationState.setTabBarSheetId(activeSheet.id),
        fe === "automation" && dialogState.setAutomationInspection(undefined),
        dialogState.setActiveDialogMode(fe));
    }
  }
  function handleGridKeyDown(fe) {
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
      operationStatusState.isOperationBusy &&
      ((Oe && ["z", "y", "b", "i", "u", "d", "r", "v"].includes(Be)) ||
        (!Oe &&
          (fe.key.length === 1 ||
            ["F2", "Delete", "Backspace"].includes(fe.key))))
    ) {
      (fe.preventDefault(),
        operationStatusState.setStatusNotice(
          "Wait for the current workbook operation to finish.",
        ));
      return;
    }
    if (
      (fe.key === "ContextMenu" && !fe.shiftKey && !Oe) ||
      (fe.key === "F10" && fe.shiftKey && !fe.ctrlKey && !fe.metaKey)
    ) {
      fe.preventDefault();
      const rr = (
        fe.currentTarget.querySelector(
          `[data-address="${selectionState.activeCellAddress}"]`,
        ) ?? fe.currentTarget
      ).getBoundingClientRect();
      selectionState.setContextMenuState({
        x: rr.left + Math.min(40, rr.width),
        y: rr.top + Math.min(24, rr.height),
        kind:
          parseRangeAddress(selectionState.activeSelectionRange).lastRow -
            parseRangeAddress(selectionState.activeSelectionRange).firstRow +
            1 ===
          activeSheet.rowCount
            ? "column"
            : parseRangeAddress(selectionState.activeSelectionRange).lastCol -
                  parseRangeAddress(selectionState.activeSelectionRange)
                    .firstCol +
                  1 ===
                activeSheet.colCount
              ? "row"
              : "cell",
      });
      return;
    }
    if (fe.key === "Escape" && clipboardState.clipboardSnapshotState) {
      (fe.preventDefault(),
        viewAndSelectionActions.clearCutSession(),
        operationStatusState.setStatusNotice(
          "Cut canceled; source cells remain unchanged.",
        ));
      return;
    }
    if (fe.key === "Escape" && selectionState.completedSelectionAreas.length) {
      (fe.preventDefault(),
        viewAndSelectionActions.selectSingleRange(
          selectionState.activeCellAddress,
        ),
        selectionState.setSelectionAnchorAddress(
          selectionState.activeCellAddress,
        ),
        (selectionState.activeCellAddressRef.current =
          selectionState.activeCellAddress));
      return;
    }
    if (fe.key === "Escape" && !fe.ctrlKey && !fe.metaKey && !fe.shiftKey) {
      (fe.preventDefault(),
        selectionState.appRootRef.current
          ?.querySelector('[role="tab"][aria-selected="true"]')
          ?.focus());
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
    if (
      selectionState.completedSelectionAreas.length &&
      ["Tab", "Enter"].includes(fe.key)
    ) {
      (fe.preventDefault(),
        viewAndSelectionActions.runGuardedAction(() => {
          const jt = enumerateSelectionCells(selectionAreas).filter((xr) => {
            const Sr = _t(xr);
            return (
              !rowVisibilityIndex.isHidden(Sr.row) &&
              !columnVisibilityIndex.isHidden(Sr.col)
            );
          });
          if (!jt.length) return;
          const rr = jt.indexOf(selectionState.activeCellAddress),
            Mr = fe.shiftKey ? -1 : 1,
            Br = jt[(rr + Mr + jt.length) % jt.length];
          cellEditActions.commitCellEdit() &&
            (selectionState.setActiveCellAddress(Br),
            selectionState.setSelectionAnchorAddress(Br),
            (selectionState.activeCellAddressRef.current = Br),
            viewAndSelectionActions.requestCellScroll(Br));
        }));
      return;
    }
    const Ue = _t(
      fe.shiftKey && fe.key.startsWith("Arrow")
        ? selectionState.activeCellAddressRef.current
        : selectionState.activeCellAddress,
    );
    let row = Ue.row,
      col = Ue.col;
    if ((fe.ctrlKey || fe.metaKey) && fe.key.toLowerCase() === "z") {
      (fe.preventDefault(), navigationActions.undoOrRedo(fe.shiftKey));
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
        (fe.preventDefault(),
          navigationActions.selectRowOrColumn("column", Ue.col, fe.shiftKey));
        return;
      }
      if (jt === "pageup" || jt === "pagedown") {
        fe.preventDefault();
        const rr = activeWorkbook.sheets.filter((sheet) => !sheet.hidden),
          Mr = rr.findIndex((xr) => xr.id === activeSheet.id),
          Br = rr[Mr + (jt === "pageup" ? -1 : 1)];
        Br &&
          cellEditActions.commitCellEdit() &&
          navigationActions.switchActiveWorksheet(Br.id);
        return;
      }
      if (jt.startsWith("arrow")) {
        fe.preventDefault();
        const rr = viewAndSelectionActions.jumpToDataEdge(
          Ue.row,
          Ue.col,
          jt === "arrowup" || jt === "arrowdown",
          jt === "arrowup" || jt === "arrowleft" ? -1 : 1,
        );
        navigationActions.selectCellAddress(
          `${columnIndexToLetters(rr.col)}${rr.row + 1}`,
          fe.shiftKey,
        );
        return;
      }
      if (jt === "a") {
        (fe.preventDefault(), navigationActions.selectAllCells());
        return;
      }
      if (["b", "i", "u"].includes(jt)) {
        if ((fe.preventDefault(), !isEditingLocked)) {
          const rr = jt === "b" ? "bold" : jt === "i" ? "italic" : "underline";
          cellEditActions.applyCellFormatPatch({
            style: { [rr]: !resolveEffectiveCellStyle(activeCell?.style)[rr] },
          });
        }
        return;
      }
      if (jt === "d") {
        (fe.preventDefault(),
          isEditingLocked ||
            viewAndSelectionActions.runGuardedAction(() => {
              (viewAndSelectionActions.requireSingleSelectionRange(),
                navigationActions.commitWorkbookEdit(
                  applyCellPreview(
                    activeWorkbook,
                    buildFillPreview(
                      activeWorkbook,
                      activeSheet.id,
                      selectionState.activeSelectionRange,
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
            viewAndSelectionActions.runGuardedAction(() => {
              (viewAndSelectionActions.requireSingleSelectionRange(),
                navigationActions.commitWorkbookEdit(
                  applyCellPreview(
                    activeWorkbook,
                    buildFillPreview(
                      activeWorkbook,
                      activeSheet.id,
                      selectionState.activeSelectionRange,
                      "right",
                    ),
                  ),
                  "Fill right",
                ));
            }));
        return;
      }
      if (jt === "y") {
        (fe.preventDefault(), navigationActions.undoOrRedo(true));
        return;
      }
      if (jt === "f" || jt === "h") {
        (fe.preventDefault(),
          findReplaceState.setIsFindBarVisible(true),
          requestAnimationFrame(() =>
            findReplaceState.findInputRef.current?.focus(),
          ));
        return;
      }
      if (jt === "home") {
        (fe.preventDefault(),
          navigationActions.selectCellAddress(
            `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`,
            fe.shiftKey,
          ));
        return;
      }
      if (jt === "end") {
        fe.preventDefault();
        const occupiedRangeBounds = findOccupiedRangeBounds(activeSheet);
        navigationActions.selectCellAddress(
          `${columnIndexToLetters(columnVisibilityIndex.next(occupiedRangeBounds?.lastCol ?? 0, -1, true) ?? columnVisibilityIndex.logicalIndex(0))}${(rowVisibilityIndex.next(occupiedRangeBounds?.lastRow ?? 0, -1, true) ?? rowVisibilityIndex.logicalIndex(0)) + 1}`,
          fe.shiftKey,
        );
        return;
      }
      return;
    }
    if (fe.shiftKey && fe.key === " ") {
      (fe.preventDefault(),
        navigationActions.selectRowOrColumn("row", Ue.row, false));
      return;
    }
    if (["Home", "PageUp", "PageDown"].includes(fe.key)) {
      (fe.preventDefault(),
        fe.key === "Home"
          ? (col = columnVisibilityIndex.logicalIndex(0))
          : (row = viewAndSelectionActions.stepVisibleRowIndex(
              row,
              Math.max(
                1,
                Math.floor(rowViewport / zoomScale / defaultRowHeight) - 1,
              ) * (fe.key === "PageUp" ? -1 : 1),
            )),
        navigationActions.selectCellAddress(
          `${columnIndexToLetters(col)}${row + 1}`,
          fe.shiftKey,
        ));
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
        fe.key === "ArrowDown" &&
          (row = viewAndSelectionActions.stepVisibleRowIndex(row, 1)),
        fe.key === "Enter" &&
          (row = viewAndSelectionActions.stepVisibleRowIndex(
            row,
            fe.shiftKey ? -1 : 1,
          )),
        fe.key === "ArrowUp" &&
          (row = viewAndSelectionActions.stepVisibleRowIndex(row, -1)),
        fe.key === "ArrowLeft" &&
          (col = viewAndSelectionActions.stepVisibleColumnIndex(col, -1)),
        (fe.key === "ArrowRight" || fe.key === "Tab") &&
          (col = viewAndSelectionActions.stepVisibleColumnIndex(
            col,
            fe.shiftKey && fe.key === "Tab" ? -1 : 1,
          )),
        !fe.shiftKey &&
          !filterSortState.filterValueSearch &&
          !filterSortState.activeColumnSort)
      ) {
        const mergeRange = mergeRanges.find((mergeRange) =>
          Tp(mergeRange, Ue.row, Ue.col),
        );
        mergeRange &&
          (col > Ue.col &&
            (col = columnVisibilityIndex.next(mergeRange.lastCol, 1) ?? Ue.col),
          col < Ue.col &&
            (col =
              columnVisibilityIndex.next(mergeRange.firstCol, -1) ?? Ue.col),
          row > Ue.row &&
            (row = rowVisibilityIndex.next(mergeRange.lastRow, 1) ?? Ue.row),
          row < Ue.row &&
            (row = rowVisibilityIndex.next(mergeRange.firstRow, -1) ?? Ue.row));
      }
      const jt = `${columnIndexToLetters(Math.max(0, Math.min(activeSheet.colCount - 1, col)))}${Math.max(1, Math.min(activeSheet.rowCount, row + 1))}`;
      navigationActions.selectCellAddress(
        jt,
        fe.shiftKey && !["Tab", "Enter"].includes(fe.key),
      );
      return;
    }
    (!isEditingLocked &&
      (fe.key === "Delete" || fe.key === "Backspace") &&
      (fe.preventDefault(),
      viewAndSelectionActions.runGuardedAction(() =>
        cellEditActions.applyCellPatches(
          viewAndSelectionActions.getSelectedCellAddresses().map((address) => ({
            sheetId: activeSheet.id,
            address: address,
            cell: { ...activeSheet.cells[address], raw: null },
          })),
          "Clear selection",
        ),
      )),
      !isEditingLocked &&
        (fe.key === "F2" ||
          (fe.key.length === 1 && !fe.ctrlKey && !fe.metaKey && !fe.altKey)) &&
        (fe.preventDefault(),
        (formulaEntryState.formulaPointerStateRef.current.mode =
          fe.key !== "F2"),
        formulaEntryState.setIsFormulaPointMode(
          formulaEntryState.formulaPointerStateRef.current.mode,
        ),
        cellEditActions.beginCellEdit(
          fe.key === "F2" ? getEditableCellText(activeCell) : fe.key,
          true,
        )));
  }
  function commitDialogInput() {
    dialogState.promptDialogState &&
      viewAndSelectionActions.runGuardedAction(() => {
        if (dialogState.promptDialogState.kind === "note")
          cellEditActions.applyCellFormatPatch({
            note: dialogState.promptDialogState.value,
          });
        else if (dialogState.promptDialogState.kind === "rename")
          navigationActions.replaceWorkbookSnapshot(
            renameWorkbook(activeWorkbook, dialogState.promptDialogState.value),
            "Rename workbook",
          );
        else {
          const workbook = addWorksheet(
              activeWorkbook,
              dialogState.promptDialogState.value,
            ),
            Oe = workbook.sheets.at(-1);
          (navigationActions.replaceWorkbookSnapshot(
            workbook,
            `Add sheet: ${Oe.name}`,
          ),
            navigationActions.switchActiveWorksheet(Oe.id));
        }
        dialogState.setPromptDialogState(null);
      });
  }
  function activateGridPane(fe) {
    return splitPaneSizes.active.current === fe
      ? true
      : !formulaEntryState.activeCellEditRef.current?.value.startsWith("=") &&
          !cellEditActions.commitCellEdit()
        ? false
        : (splitPaneSizes.activate(fe),
          (selectionState.gridContainerRef.current =
            splitPaneSizes.refs.current[fe] ?? null),
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
