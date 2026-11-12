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
  function selectionToTsv(fe = false) {
    (viewAndSelectionActions.requireSingleSelectionRange(),
      viewAndSelectionActions.getSelectedCellAddresses());
    const Oe = {
        row: selectionRangeBounds.firstRow,
        col: selectionRangeBounds.firstCol,
      },
      Be = {
        row: selectionRangeBounds.lastRow,
        col: selectionRangeBounds.lastCol,
      },
      Ue = [];
    for (let row = Oe.row; row <= Be.row; row++) {
      const wt = [];
      for (let col = Oe.col; col <= Be.col; col++) {
        const rr = `${columnIndexToLetters(col)}${row + 1}`,
          Mr = workbookCalculator(activeSheet.id, rr);
        wt.push(
          typeof Mr == "string" && /^[=+@-]/.test(Mr)
            ? `'${Mr}`
            : String(Mr ?? ""),
        );
      }
      Ue.push(wt);
    }
    const Xe = x7.unparse(Ue, {
        delimiter: "	",
        newline: `
`,
      }),
      snapshot = fe
        ? createCutSourceSnapshot(
            activeWorkbook,
            activeSheet.id,
            selectionState.activeSelectionRange,
          )
        : null;
    return (
      clipboardState.setClipboardSnapshotState(
        snapshot ? { snapshot: snapshot, text: Xe } : null,
      ),
      (clipboardState.internalClipboardRef.current = {
        text: Xe,
        snapshot: buildClipboardSnapshot(
          activeWorkbook,
          activeSheet.id,
          selectionState.activeSelectionRange,
        ),
      }),
      Xe
    );
  }
  async function copyOrCutSelection(fe) {
    try {
      if (fe && isEditingLocked) throw new Error("This workbook is read-only.");
      const Be = activeWorkbook,
        Ue = selectionToTsv(fe);
      if ((await navigator.clipboard.writeText(Ue), fe)) {
        if (
          latestEditStateRef.current.readOnly ||
          latestEditStateRef.current.book !== Be ||
          latestEditStateRef.current.sheetId !== activeSheet.id ||
          latestEditStateRef.current.selectionKey !== selectionAreasKey
        )
          throw new Error("Workbook changed while copying. Copy again.");
        operationStatusState.setStatusNotice(
          "Cut is ready. Select a destination and paste to preview the move.",
        );
      }
      fe || operationStatusState.setStatusNotice("Selection copied.");
    } catch (error) {
      (fe && clipboardState.setClipboardSnapshotState(null),
        operationStatusState.setErrorNotice(
          error instanceof Error && error.name !== "NotAllowedError"
            ? error.message
            : "Clipboard access is unavailable. Use Ctrl/⌘ C or X with the grid focused.",
        ),
        selectionState.gridContainerRef.current?.focus());
    }
  }
  function pasteClipboardText(fe) {
    viewAndSelectionActions.runGuardedAction(() => {
      if (
        (viewAndSelectionActions.requireSingleSelectionRange(),
        fe.length > workbookLimits.maxFileBytes)
      )
        throw new Error("Clipboard text is too large.");
      const Oe =
        clipboardState.internalClipboardRef.current?.text === fe
          ? clipboardState.internalClipboardRef.current
          : null;
      if (Oe && clipboardState.clipboardSnapshotState?.text === fe) {
        clipboardState.setMoveCellsSource(
          clipboardState.clipboardSnapshotState.snapshot,
        );
        return;
      }
      if (Oe) {
        const pastePreview = buildPastePreview(
          activeWorkbook,
          activeSheet.id,
          selectionState.activeSelectionRange,
          Oe.snapshot,
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
      const Be = x7.parse(fe.replace(/\r?\n$/, ""), {
        delimiter: "	",
        skipEmptyLines: false,
      });
      if (Be.errors.length)
        throw new Error("Clipboard text contains malformed quoted cells.");
      const data = Be.data,
        Xe = _t(selectionState.activeCellAddress),
        gt = [];
      if (!data.length || data.some((wt) => wt.length !== data[0].length))
        throw new Error("Paste requires a rectangular range.");
      if (data.length * data[0].length > workbookLimits.maxPatches)
        throw new Error(
          `Paste at most ${workbookLimits.maxPatches} cells at a time.`,
        );
      (data.forEach((wt, jt) =>
        wt.forEach((rr, Mr) => {
          if (
            Xe.row + jt >= activeSheet.rowCount ||
            Xe.col + Mr >= activeSheet.colCount
          )
            throw new Error(
              "Paste exceeds the Excel worksheet boundary (XFD1048576).",
            );
          const address = `${columnIndexToLetters(Xe.col + Mr)}${Xe.row + jt + 1}`,
            dateCellEntry = parseDateCellEntry(
              rr,
              activeSheet.cells[address],
              latestEditStateRef.current.book.datePolicy?.dateSystem,
            );
          gt.push({
            sheetId: activeSheet.id,
            address: address,
            cell: dateCellEntry,
          });
        }),
      ),
        cellEditActions.applyCellPatches(gt, "Paste cells"),
        viewAndSelectionActions.selectSingleRange(
          wy(
            selectionState.activeCellAddress,
            `${columnIndexToLetters(Xe.col + data[0].length - 1)}${Xe.row + data.length}`,
          ),
        ));
    });
  }
  async function pasteFromClipboard() {
    try {
      const Oe = activeWorkbook,
        Be = await navigator.clipboard.readText();
      if (
        latestEditStateRef.current.book !== Oe ||
        latestEditStateRef.current.sheetId !== activeSheet.id ||
        latestEditStateRef.current.selectionKey !== selectionAreasKey
      )
        throw new Error("Workbook changed while pasting. Try again.");
      pasteClipboardText(Be);
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
      const fe = Object.keys(activeSheet.cells)
        .filter((cell) => {
          const Xe = _t(cell);
          return (
            !rowVisibilityIndex.isHidden(Xe.row) &&
            !columnVisibilityIndex.isHidden(Xe.col) &&
            String(activeSheet.cells[cell]?.raw ?? "")
              .toLowerCase()
              .includes(findReplaceState.findQuery.toLowerCase())
          );
        })
        .sort((Ue, Xe) => {
          const gt = _t(Ue),
            Nt = _t(Xe);
          return gt.row - Nt.row || gt.col - Nt.col;
        });
      if (!fe.length) {
        operationStatusState.setStatusNotice(
          `We couldn't find “${findReplaceState.findQuery}”.`,
        );
        return;
      }
      const Oe = _t(selectionState.activeCellAddress),
        Be =
          fe.find((Ue) => {
            const Xe = _t(Ue);
            return Xe.row > Oe.row || (Xe.row === Oe.row && Xe.col > Oe.col);
          }) ?? fe[0];
      (navigationActions.selectCellAddress(Be),
        operationStatusState.setStatusNotice(
          `${fe.length} matching cell${fe.length === 1 ? "" : "s"}.`,
        ));
    });
  }
  function replaceTextInSelection() {
    viewAndSelectionActions.runGuardedAction(() => {
      if (!findReplaceState.findQuery) throw new Error("Enter text to find.");
      const fe = new RegExp(
          findReplaceState.findQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
          "gi",
        ),
        Oe = viewAndSelectionActions
          .getSelectedCellAddresses()
          .filter(
            (Be) =>
              typeof activeSheet.cells[Be]?.raw == "string" &&
              String(activeSheet.cells[Be].raw)
                .toLowerCase()
                .includes(findReplaceState.findQuery.toLowerCase()),
          )
          .map((address) => ({
            sheetId: activeSheet.id,
            address: address,
            cell: {
              ...activeSheet.cells[address],
              raw: String(activeSheet.cells[address].raw).replace(
                fe,
                () => findReplaceState.replaceWithText,
              ),
            },
          }));
      (Oe.length &&
        cellEditActions.applyCellPatches(Oe, "Replace text in selection"),
        operationStatusState.setStatusNotice(
          `Replaced text in ${Oe.length} cells.`,
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
