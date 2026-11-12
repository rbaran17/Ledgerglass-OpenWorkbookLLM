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
  function runEditorCommand(commandName) {
    if (commandName === "xlsxSource") {
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
        (commandName === "populationFilter" &&
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
          ].includes(commandName))
      ) {
        operationStatusState.setErrorNotice(
          "This command requires one rectangular range. Click a cell or enter a single range to continue.",
        );
        return;
      }
      if (commandName === "calculateNow") {
        persistenceActions.runCalculationNow();
        return;
      }
      if (
        commandName === "pasteSpecial" &&
        clipboardState.clipboardSnapshotState
      ) {
        clipboardState.setMoveCellsSource(
          clipboardState.clipboardSnapshotState.snapshot,
        );
        return;
      }
      (commandName === "sheets" &&
        navigationState.setTabBarSheetId(activeSheet.id),
        commandName === "automation" &&
          dialogState.setAutomationInspection(undefined),
        dialogState.setActiveDialogMode(commandName));
    }
  }
  function handleGridKeyDown(event) {
    if (
      event.target !== event.currentTarget ||
      event.defaultPrevented ||
      event.nativeEvent.isComposing ||
      (event.altKey &&
        !(
          (event.ctrlKey || event.metaKey) &&
          event.key.toLowerCase() === "v"
        )) ||
      event.getModifierState("AltGraph")
    )
      return;
    const hasCommandModifier = event.ctrlKey || event.metaKey,
      lowerKey = event.key.toLowerCase();
    if (
      (hasCommandModifier &&
        ["tab", "enter", "f2", "f5", "f6", "delete", "backspace"].includes(
          lowerKey,
        )) ||
      (hasCommandModifier &&
        event.shiftKey &&
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
        ].includes(lowerKey))
    )
      return;
    if (
      operationStatusState.isOperationBusy &&
      ((hasCommandModifier &&
        ["z", "y", "b", "i", "u", "d", "r", "v"].includes(lowerKey)) ||
        (!hasCommandModifier &&
          (event.key.length === 1 ||
            ["F2", "Delete", "Backspace"].includes(event.key))))
    ) {
      (event.preventDefault(),
        operationStatusState.setStatusNotice(
          "Wait for the current workbook operation to finish.",
        ));
      return;
    }
    if (
      (event.key === "ContextMenu" && !event.shiftKey && !hasCommandModifier) ||
      (event.key === "F10" &&
        event.shiftKey &&
        !event.ctrlKey &&
        !event.metaKey)
    ) {
      event.preventDefault();
      const anchorRect = (
        event.currentTarget.querySelector(
          `[data-address="${selectionState.activeCellAddress}"]`,
        ) ?? event.currentTarget
      ).getBoundingClientRect();
      selectionState.setContextMenuState({
        x: anchorRect.left + Math.min(40, anchorRect.width),
        y: anchorRect.top + Math.min(24, anchorRect.height),
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
    if (event.key === "Escape" && clipboardState.clipboardSnapshotState) {
      (event.preventDefault(),
        viewAndSelectionActions.clearCutSession(),
        operationStatusState.setStatusNotice(
          "Cut canceled; source cells remain unchanged.",
        ));
      return;
    }
    if (
      event.key === "Escape" &&
      selectionState.completedSelectionAreas.length
    ) {
      (event.preventDefault(),
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
    if (
      event.key === "Escape" &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.shiftKey
    ) {
      (event.preventDefault(),
        selectionState.appRootRef.current
          ?.querySelector('[role="tab"][aria-selected="true"]')
          ?.focus());
      return;
    }
    if (
      (event.key === "F5" && !event.shiftKey && !hasCommandModifier) ||
      ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "g")
    ) {
      (event.preventDefault(), runEditorCommand("goToSpecial"));
      return;
    }
    if (
      (event.ctrlKey || event.metaKey) &&
      event.shiftKey &&
      event.key.toLowerCase() === "f"
    ) {
      (event.preventDefault(), runEditorCommand("search"));
      return;
    }
    if (
      selectionState.completedSelectionAreas.length &&
      ["Tab", "Enter"].includes(event.key)
    ) {
      (event.preventDefault(),
        viewAndSelectionActions.runGuardedAction(() => {
          const visibleCells = enumerateSelectionCells(selectionAreas).filter(
            (cellAddress) => {
              const cellPosition = _t(cellAddress);
              return (
                !rowVisibilityIndex.isHidden(cellPosition.row) &&
                !columnVisibilityIndex.isHidden(cellPosition.col)
              );
            },
          );
          if (!visibleCells.length) return;
          const activeIndex = visibleCells.indexOf(
              selectionState.activeCellAddress,
            ),
            direction = event.shiftKey ? -1 : 1,
            nextCellAddress =
              visibleCells[
                (activeIndex + direction + visibleCells.length) %
                  visibleCells.length
              ];
          cellEditActions.commitCellEdit() &&
            (selectionState.setActiveCellAddress(nextCellAddress),
            selectionState.setSelectionAnchorAddress(nextCellAddress),
            (selectionState.activeCellAddressRef.current = nextCellAddress),
            viewAndSelectionActions.requestCellScroll(nextCellAddress));
        }));
      return;
    }
    const originCell = _t(
      event.shiftKey && event.key.startsWith("Arrow")
        ? selectionState.activeCellAddressRef.current
        : selectionState.activeCellAddress,
    );
    let row = originCell.row,
      col = originCell.col;
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
      (event.preventDefault(), navigationActions.undoOrRedo(event.shiftKey));
      return;
    }
    if (event.ctrlKey || event.metaKey) {
      const shortcutKey = event.key.toLowerCase();
      if (shortcutKey === "1") {
        (event.preventDefault(), runEditorCommand("numberFormat"));
        return;
      }
      if (shortcutKey === "v" && event.altKey) {
        (event.preventDefault(), runEditorCommand("pasteSpecial"));
        return;
      }
      if (shortcutKey === " ") {
        (event.preventDefault(),
          navigationActions.selectRowOrColumn(
            "column",
            originCell.col,
            event.shiftKey,
          ));
        return;
      }
      if (shortcutKey === "pageup" || shortcutKey === "pagedown") {
        event.preventDefault();
        const visibleSheets = activeWorkbook.sheets.filter(
            (sheet) => !sheet.hidden,
          ),
          activeSheetIndex = visibleSheets.findIndex(
            (visibleSheet) => visibleSheet.id === activeSheet.id,
          ),
          targetSheet =
            visibleSheets[
              activeSheetIndex + (shortcutKey === "pageup" ? -1 : 1)
            ];
        targetSheet &&
          cellEditActions.commitCellEdit() &&
          navigationActions.switchActiveWorksheet(targetSheet.id);
        return;
      }
      if (shortcutKey.startsWith("arrow")) {
        event.preventDefault();
        const edgeCell = viewAndSelectionActions.jumpToDataEdge(
          originCell.row,
          originCell.col,
          shortcutKey === "arrowup" || shortcutKey === "arrowdown",
          shortcutKey === "arrowup" || shortcutKey === "arrowleft" ? -1 : 1,
        );
        navigationActions.selectCellAddress(
          `${columnIndexToLetters(edgeCell.col)}${edgeCell.row + 1}`,
          event.shiftKey,
        );
        return;
      }
      if (shortcutKey === "a") {
        (event.preventDefault(), navigationActions.selectAllCells());
        return;
      }
      if (["b", "i", "u"].includes(shortcutKey)) {
        if ((event.preventDefault(), !isEditingLocked)) {
          const styleProperty =
            shortcutKey === "b"
              ? "bold"
              : shortcutKey === "i"
                ? "italic"
                : "underline";
          cellEditActions.applyCellFormatPatch({
            style: {
              [styleProperty]: !resolveEffectiveCellStyle(activeCell?.style)[
                styleProperty
              ],
            },
          });
        }
        return;
      }
      if (shortcutKey === "d") {
        (event.preventDefault(),
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
      if (shortcutKey === "r") {
        (event.preventDefault(),
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
      if (shortcutKey === "y") {
        (event.preventDefault(), navigationActions.undoOrRedo(true));
        return;
      }
      if (shortcutKey === "f" || shortcutKey === "h") {
        (event.preventDefault(),
          findReplaceState.setIsFindBarVisible(true),
          requestAnimationFrame(() =>
            findReplaceState.findInputRef.current?.focus(),
          ));
        return;
      }
      if (shortcutKey === "home") {
        (event.preventDefault(),
          navigationActions.selectCellAddress(
            `${columnIndexToLetters(columnVisibilityIndex.logicalIndex(0))}${rowVisibilityIndex.logicalIndex(0) + 1}`,
            event.shiftKey,
          ));
        return;
      }
      if (shortcutKey === "end") {
        event.preventDefault();
        const occupiedRangeBounds = findOccupiedRangeBounds(activeSheet);
        navigationActions.selectCellAddress(
          `${columnIndexToLetters(columnVisibilityIndex.next(occupiedRangeBounds?.lastCol ?? 0, -1, true) ?? columnVisibilityIndex.logicalIndex(0))}${(rowVisibilityIndex.next(occupiedRangeBounds?.lastRow ?? 0, -1, true) ?? rowVisibilityIndex.logicalIndex(0)) + 1}`,
          event.shiftKey,
        );
        return;
      }
      return;
    }
    if (event.shiftKey && event.key === " ") {
      (event.preventDefault(),
        navigationActions.selectRowOrColumn("row", originCell.row, false));
      return;
    }
    if (["Home", "PageUp", "PageDown"].includes(event.key)) {
      (event.preventDefault(),
        event.key === "Home"
          ? (col = columnVisibilityIndex.logicalIndex(0))
          : (row = viewAndSelectionActions.stepVisibleRowIndex(
              row,
              Math.max(
                1,
                Math.floor(rowViewport / zoomScale / defaultRowHeight) - 1,
              ) * (event.key === "PageUp" ? -1 : 1),
            )),
        navigationActions.selectCellAddress(
          `${columnIndexToLetters(col)}${row + 1}`,
          event.shiftKey,
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
      ].includes(event.key)
    ) {
      if (
        (event.preventDefault(),
        event.key === "ArrowDown" &&
          (row = viewAndSelectionActions.stepVisibleRowIndex(row, 1)),
        event.key === "Enter" &&
          (row = viewAndSelectionActions.stepVisibleRowIndex(
            row,
            event.shiftKey ? -1 : 1,
          )),
        event.key === "ArrowUp" &&
          (row = viewAndSelectionActions.stepVisibleRowIndex(row, -1)),
        event.key === "ArrowLeft" &&
          (col = viewAndSelectionActions.stepVisibleColumnIndex(col, -1)),
        (event.key === "ArrowRight" || event.key === "Tab") &&
          (col = viewAndSelectionActions.stepVisibleColumnIndex(
            col,
            event.shiftKey && event.key === "Tab" ? -1 : 1,
          )),
        !event.shiftKey &&
          !filterSortState.filterValueSearch &&
          !filterSortState.activeColumnSort)
      ) {
        const mergeRange = mergeRanges.find((mergeRange) =>
          Tp(mergeRange, originCell.row, originCell.col),
        );
        mergeRange &&
          (col > originCell.col &&
            (col =
              columnVisibilityIndex.next(mergeRange.lastCol, 1) ??
              originCell.col),
          col < originCell.col &&
            (col =
              columnVisibilityIndex.next(mergeRange.firstCol, -1) ??
              originCell.col),
          row > originCell.row &&
            (row =
              rowVisibilityIndex.next(mergeRange.lastRow, 1) ?? originCell.row),
          row < originCell.row &&
            (row =
              rowVisibilityIndex.next(mergeRange.firstRow, -1) ??
              originCell.row));
      }
      const targetAddress = `${columnIndexToLetters(Math.max(0, Math.min(activeSheet.colCount - 1, col)))}${Math.max(1, Math.min(activeSheet.rowCount, row + 1))}`;
      navigationActions.selectCellAddress(
        targetAddress,
        event.shiftKey && !["Tab", "Enter"].includes(event.key),
      );
      return;
    }
    (!isEditingLocked &&
      (event.key === "Delete" || event.key === "Backspace") &&
      (event.preventDefault(),
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
        (event.key === "F2" ||
          (event.key.length === 1 &&
            !event.ctrlKey &&
            !event.metaKey &&
            !event.altKey)) &&
        (event.preventDefault(),
        (formulaEntryState.formulaPointerStateRef.current.mode =
          event.key !== "F2"),
        formulaEntryState.setIsFormulaPointMode(
          formulaEntryState.formulaPointerStateRef.current.mode,
        ),
        cellEditActions.beginCellEdit(
          event.key === "F2" ? getEditableCellText(activeCell) : event.key,
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
            newSheet = workbook.sheets.at(-1);
          (navigationActions.replaceWorkbookSnapshot(
            workbook,
            `Add sheet: ${newSheet.name}`,
          ),
            navigationActions.switchActiveWorksheet(newSheet.id));
        }
        dialogState.setPromptDialogState(null);
      });
  }
  function activateGridPane(paneId) {
    return splitPaneSizes.active.current === paneId
      ? true
      : !formulaEntryState.activeCellEditRef.current?.value.startsWith("=") &&
          !cellEditActions.commitCellEdit()
        ? false
        : (splitPaneSizes.activate(paneId),
          (selectionState.gridContainerRef.current =
            splitPaneSizes.refs.current[paneId] ?? null),
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
