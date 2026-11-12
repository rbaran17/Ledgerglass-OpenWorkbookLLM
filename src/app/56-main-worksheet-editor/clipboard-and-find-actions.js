import { x7 } from "../08-xlsx-import-export-and-preservation.js";
import { createCutSourceSnapshot } from "../47-fill-and-paste-special.js";
import {
  buildClipboardSnapshot,
  buildPastePreview,
  parseDateCellEntry,
} from "../04-cell-input-paste-and-fill.js";
import { _t, columnIndexToLetters } from "../03-workbook-model.js";
import { workbookLimits } from "../02-text-and-financial-functions.js";
import { wy } from "../55-external-source-refresh.js";

function createClipboardAndFindActions({
  viewAndSelectionActions,
  cellEditActions,
  navigationActions,
  selectionState,
  operationStatusState,
  findReplaceState,
  clipboardState,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  selectionAreasKey,
  workbookCalculator,
  rowVisibilityIndex,
  columnVisibilityIndex,
  latestEditStateRef,
  selectionRangeBounds,
}) {
  function selectionToTsv(isCut = false) {
    (viewAndSelectionActions.requireSingleSelectionRange(),
      viewAndSelectionActions.getSelectedCellAddresses());
    const startCell = {
        row: selectionRangeBounds.firstRow,
        col: selectionRangeBounds.firstCol,
      },
      endCell = {
        row: selectionRangeBounds.lastRow,
        col: selectionRangeBounds.lastCol,
      },
      rows = [];
    for (let row = startCell.row; row <= endCell.row; row++) {
      const rowValues = [];
      for (let col = startCell.col; col <= endCell.col; col++) {
        const cellAddress = `${columnIndexToLetters(col)}${row + 1}`,
          cellValue = workbookCalculator(activeSheet.id, cellAddress);
        rowValues.push(
          typeof cellValue == "string" && /^[=+@-]/.test(cellValue)
            ? `'${cellValue}`
            : String(cellValue ?? ""),
        );
      }
      rows.push(rowValues);
    }
    const tsvText = x7.unparse(rows, {
        delimiter: "	",
        newline: `
`,
      }),
      snapshot = isCut
        ? createCutSourceSnapshot(
            activeWorkbook,
            activeSheet.id,
            selectionState.activeSelectionRange,
          )
        : null;
    return (
      clipboardState.setClipboardSnapshotState(
        snapshot ? { snapshot: snapshot, text: tsvText } : null,
      ),
      (clipboardState.internalClipboardRef.current = {
        text: tsvText,
        snapshot: buildClipboardSnapshot(
          activeWorkbook,
          activeSheet.id,
          selectionState.activeSelectionRange,
        ),
      }),
      tsvText
    );
  }
  async function copyOrCutSelection(isCut) {
    try {
      if (isCut && isEditingLocked)
        throw new Error("This workbook is read-only.");
      const workbookAtCopy = activeWorkbook,
        tsvText = selectionToTsv(isCut);
      if ((await navigator.clipboard.writeText(tsvText), isCut)) {
        if (
          latestEditStateRef.current.readOnly ||
          latestEditStateRef.current.book !== workbookAtCopy ||
          latestEditStateRef.current.sheetId !== activeSheet.id ||
          latestEditStateRef.current.selectionKey !== selectionAreasKey
        )
          throw new Error("Workbook changed while copying. Copy again.");
        operationStatusState.setStatusNotice(
          "Cut is ready. Select a destination and paste to preview the move.",
        );
      }
      isCut || operationStatusState.setStatusNotice("Selection copied.");
    } catch (error) {
      (isCut && clipboardState.setClipboardSnapshotState(null),
        operationStatusState.setErrorNotice(
          error instanceof Error && error.name !== "NotAllowedError"
            ? error.message
            : "Clipboard access is unavailable. Use Ctrl/⌘ C or X with the grid focused.",
        ),
        selectionState.gridContainerRef.current?.focus());
    }
  }
  function pasteClipboardText(clipboardText) {
    viewAndSelectionActions.runGuardedAction(() => {
      if (
        (viewAndSelectionActions.requireSingleSelectionRange(),
        clipboardText.length > workbookLimits.maxFileBytes)
      )
        throw new Error("Clipboard text is too large.");
      const internalClipboard =
        clipboardState.internalClipboardRef.current?.text === clipboardText
          ? clipboardState.internalClipboardRef.current
          : null;
      if (
        internalClipboard &&
        clipboardState.clipboardSnapshotState?.text === clipboardText
      ) {
        clipboardState.setMoveCellsSource(
          clipboardState.clipboardSnapshotState.snapshot,
        );
        return;
      }
      if (internalClipboard) {
        const pastePreview = buildPastePreview(
          activeWorkbook,
          activeSheet.id,
          selectionState.activeSelectionRange,
          internalClipboard.snapshot,
          {
            mode: "all",
            deferValidationWarnings: true,
            readOnly: latestEditStateRef.current.readOnly,
          },
        );
        (cellEditActions.applyCellPatches(pastePreview.patches, "Paste cells"),
          viewAndSelectionActions.selectSingleRange(pastePreview.targetRange));
        return;
      }
      const parseResult = x7.parse(clipboardText.replace(/\r?\n$/, ""), {
        delimiter: "	",
        skipEmptyLines: false,
      });
      if (parseResult.errors.length)
        throw new Error("Clipboard text contains malformed quoted cells.");
      const data = parseResult.data,
        anchorCell = _t(selectionState.activeCellAddress),
        pendingPatches = [];
      if (
        !data.length ||
        data.some((dataRow) => dataRow.length !== data[0].length)
      )
        throw new Error("Paste requires a rectangular range.");
      if (data.length * data[0].length > workbookLimits.maxPatches)
        throw new Error(
          `Paste at most ${workbookLimits.maxPatches} cells at a time.`,
        );
      (data.forEach((rowValues, rowIndex) =>
        rowValues.forEach((rawValue, columnIndex) => {
          if (
            anchorCell.row + rowIndex >= activeSheet.rowCount ||
            anchorCell.col + columnIndex >= activeSheet.colCount
          )
            throw new Error(
              "Paste exceeds the Excel worksheet boundary (XFD1048576).",
            );
          const address = `${columnIndexToLetters(anchorCell.col + columnIndex)}${anchorCell.row + rowIndex + 1}`,
            dateCellEntry = parseDateCellEntry(
              rawValue,
              activeSheet.cells[address],
              latestEditStateRef.current.book.datePolicy?.dateSystem,
            );
          pendingPatches.push({
            sheetId: activeSheet.id,
            address: address,
            cell: dateCellEntry,
          });
        }),
      ),
        cellEditActions.applyCellPatches(pendingPatches, "Paste cells"),
        viewAndSelectionActions.selectSingleRange(
          wy(
            selectionState.activeCellAddress,
            `${columnIndexToLetters(anchorCell.col + data[0].length - 1)}${anchorCell.row + data.length}`,
          ),
        ));
    });
  }
  async function pasteFromClipboard() {
    try {
      const workbookAtPaste = activeWorkbook,
        clipboardText = await navigator.clipboard.readText();
      if (
        latestEditStateRef.current.book !== workbookAtPaste ||
        latestEditStateRef.current.sheetId !== activeSheet.id ||
        latestEditStateRef.current.selectionKey !== selectionAreasKey
      )
        throw new Error("Workbook changed while pasting. Try again.");
      pasteClipboardText(clipboardText);
    } catch {
      (operationStatusState.setErrorNotice(
        "Clipboard access is unavailable. Use Ctrl/⌘ V with the grid focused.",
      ),
        selectionState.gridContainerRef.current?.focus());
    }
  }
  function findNextTextMatch() {
    viewAndSelectionActions.runGuardedAction(() => {
      if (!findReplaceState.findQuery) return;
      const matchingAddresses = Object.keys(activeSheet.cells)
        .filter((cell) => {
          const cellPosition = _t(cell);
          return (
            !rowVisibilityIndex.isHidden(cellPosition.row) &&
            !columnVisibilityIndex.isHidden(cellPosition.col) &&
            String(activeSheet.cells[cell]?.raw ?? "")
              .toLowerCase()
              .includes(findReplaceState.findQuery.toLowerCase())
          );
        })
        .sort((leftAddress, rightAddress) => {
          const leftPosition = _t(leftAddress),
            rightPosition = _t(rightAddress);
          return (
            leftPosition.row - rightPosition.row ||
            leftPosition.col - rightPosition.col
          );
        });
      if (!matchingAddresses.length) {
        operationStatusState.setStatusNotice(
          `We couldn't find “${findReplaceState.findQuery}”.`,
        );
        return;
      }
      const activePosition = _t(selectionState.activeCellAddress),
        targetAddress =
          matchingAddresses.find((candidateAddress) => {
            const candidatePosition = _t(candidateAddress);
            return (
              candidatePosition.row > activePosition.row ||
              (candidatePosition.row === activePosition.row &&
                candidatePosition.col > activePosition.col)
            );
          }) ?? matchingAddresses[0];
      (navigationActions.selectCellAddress(targetAddress),
        operationStatusState.setStatusNotice(
          `${matchingAddresses.length} matching cell${matchingAddresses.length === 1 ? "" : "s"}.`,
        ));
    });
  }
  function replaceTextInSelection() {
    viewAndSelectionActions.runGuardedAction(() => {
      if (!findReplaceState.findQuery) throw new Error("Enter text to find.");
      const searchPattern = new RegExp(
          findReplaceState.findQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
          "gi",
        ),
        replacementPatches = viewAndSelectionActions
          .getSelectedCellAddresses()
          .filter(
            (cellAddress) =>
              typeof activeSheet.cells[cellAddress]?.raw == "string" &&
              String(activeSheet.cells[cellAddress].raw)
                .toLowerCase()
                .includes(findReplaceState.findQuery.toLowerCase()),
          )
          .map((address) => ({
            sheetId: activeSheet.id,
            address: address,
            cell: {
              ...activeSheet.cells[address],
              raw: String(activeSheet.cells[address].raw).replace(
                searchPattern,
                () => findReplaceState.replaceWithText,
              ),
            },
          }));
      (replacementPatches.length &&
        cellEditActions.applyCellPatches(
          replacementPatches,
          "Replace text in selection",
        ),
        operationStatusState.setStatusNotice(
          `Replaced text in ${replacementPatches.length} cells.`,
        ));
    });
  }
  return {
    selectionToTsv,
    copyOrCutSelection,
    pasteClipboardText,
    pasteFromClipboard,
    findNextTextMatch,
    replaceTextInSelection,
  };
}

export { createClipboardAndFindActions };
