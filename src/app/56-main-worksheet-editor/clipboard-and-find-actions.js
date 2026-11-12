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
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  activeCellAddress,
  activeSelectionRange,
  selectSingleRange,
  selectionAreasKey,
  requireSingleSelectionRange,
  setStatusNotice,
  setErrorNotice,
  findQuery,
  replaceWithText,
  internalClipboardRef,
  clipboardSnapshotState,
  setClipboardSnapshotState,
  setMoveCellsSource,
  workbookCalculator,
  rowVisibilityIndex,
  columnVisibilityIndex,
  gridContainerRef,
  latestEditStateRef,
  selectionRangeBounds,
  getSelectedCellAddresses,
  runGuardedAction,
  applyCellPatches,
  selectCellAddress,
}) {
  function selectionToTsv(fe = false) {
    (requireSingleSelectionRange(), getSelectedCellAddresses());
    const Oe = {
        row: selectionRangeBounds.firstRow,
        col: selectionRangeBounds.firstCol,
      },
      Be = {
        row: selectionRangeBounds.lastRow,
        col: selectionRangeBounds.lastCol,
      },
      Ue = [];
    for (let Nt = Oe.row; Nt <= Be.row; Nt++) {
      const wt = [];
      for (let jt = Oe.col; jt <= Be.col; jt++) {
        const rr = `${columnIndexToLetters(jt)}${Nt + 1}`,
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
      gt = fe
        ? createCutSourceSnapshot(
            activeWorkbook,
            activeSheet.id,
            activeSelectionRange,
          )
        : null;
    return (
      setClipboardSnapshotState(gt ? { snapshot: gt, text: Xe } : null),
      (internalClipboardRef.current = {
        text: Xe,
        snapshot: buildClipboardSnapshot(
          activeWorkbook,
          activeSheet.id,
          activeSelectionRange,
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
        setStatusNotice(
          "Cut is ready. Select a destination and paste to preview the move.",
        );
      }
      fe || setStatusNotice("Selection copied.");
    } catch (Be) {
      (fe && setClipboardSnapshotState(null),
        setErrorNotice(
          Be instanceof Error && Be.name !== "NotAllowedError"
            ? Be.message
            : "Clipboard access is unavailable. Use Ctrl/⌘ C or X with the grid focused.",
        ),
        gridContainerRef.current?.focus());
    }
  }
  function pasteClipboardText(fe) {
    runGuardedAction(() => {
      if (
        (requireSingleSelectionRange(), fe.length > workbookLimits.maxFileBytes)
      )
        throw new Error("Clipboard text is too large.");
      const Oe =
        internalClipboardRef.current?.text === fe
          ? internalClipboardRef.current
          : null;
      if (Oe && clipboardSnapshotState?.text === fe) {
        setMoveCellsSource(clipboardSnapshotState.snapshot);
        return;
      }
      if (Oe) {
        const wt = buildPastePreview(
          activeWorkbook,
          activeSheet.id,
          activeSelectionRange,
          Oe.snapshot,
          {
            mode: "all",
            deferValidationWarnings: true,
            readOnly: latestEditStateRef.current.readOnly,
          },
        );
        (applyCellPatches(wt.patches, "Paste cells"),
          selectSingleRange(wt.targetRange));
        return;
      }
      const Be = x7.parse(fe.replace(/\r?\n$/, ""), {
        delimiter: "	",
        skipEmptyLines: false,
      });
      if (Be.errors.length)
        throw new Error("Clipboard text contains malformed quoted cells.");
      const Ue = Be.data,
        Xe = _t(activeCellAddress),
        gt = [];
      if (!Ue.length || Ue.some((wt) => wt.length !== Ue[0].length))
        throw new Error("Paste requires a rectangular range.");
      if (Ue.length * Ue[0].length > workbookLimits.maxPatches)
        throw new Error(
          `Paste at most ${workbookLimits.maxPatches} cells at a time.`,
        );
      (Ue.forEach((wt, jt) =>
        wt.forEach((rr, Mr) => {
          if (
            Xe.row + jt >= activeSheet.rowCount ||
            Xe.col + Mr >= activeSheet.colCount
          )
            throw new Error(
              "Paste exceeds the Excel worksheet boundary (XFD1048576).",
            );
          const Br = `${columnIndexToLetters(Xe.col + Mr)}${Xe.row + jt + 1}`,
            xr = parseDateCellEntry(
              rr,
              activeSheet.cells[Br],
              latestEditStateRef.current.book.datePolicy?.dateSystem,
            );
          gt.push({ sheetId: activeSheet.id, address: Br, cell: xr });
        }),
      ),
        applyCellPatches(gt, "Paste cells"),
        selectSingleRange(
          wy(
            activeCellAddress,
            `${columnIndexToLetters(Xe.col + Ue[0].length - 1)}${Xe.row + Ue.length}`,
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
      (setErrorNotice(
        "Clipboard access is unavailable. Use Ctrl/⌘ V with the grid focused.",
      ),
        gridContainerRef.current?.focus());
    }
  }
  function findNextTextMatch() {
    runGuardedAction(() => {
      if (!findQuery) return;
      const fe = Object.keys(activeSheet.cells)
        .filter((Ue) => {
          const Xe = _t(Ue);
          return (
            !rowVisibilityIndex.isHidden(Xe.row) &&
            !columnVisibilityIndex.isHidden(Xe.col) &&
            String(activeSheet.cells[Ue]?.raw ?? "")
              .toLowerCase()
              .includes(findQuery.toLowerCase())
          );
        })
        .sort((Ue, Xe) => {
          const gt = _t(Ue),
            Nt = _t(Xe);
          return gt.row - Nt.row || gt.col - Nt.col;
        });
      if (!fe.length) {
        setStatusNotice(`We couldn't find “${findQuery}”.`);
        return;
      }
      const Oe = _t(activeCellAddress),
        Be =
          fe.find((Ue) => {
            const Xe = _t(Ue);
            return Xe.row > Oe.row || (Xe.row === Oe.row && Xe.col > Oe.col);
          }) ?? fe[0];
      (selectCellAddress(Be),
        setStatusNotice(
          `${fe.length} matching cell${fe.length === 1 ? "" : "s"}.`,
        ));
    });
  }
  function replaceTextInSelection() {
    runGuardedAction(() => {
      if (!findQuery) throw new Error("Enter text to find.");
      const fe = new RegExp(
          findQuery.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
          "gi",
        ),
        Oe = getSelectedCellAddresses()
          .filter(
            (Be) =>
              typeof activeSheet.cells[Be]?.raw == "string" &&
              String(activeSheet.cells[Be].raw)
                .toLowerCase()
                .includes(findQuery.toLowerCase()),
          )
          .map((Be) => ({
            sheetId: activeSheet.id,
            address: Be,
            cell: {
              ...activeSheet.cells[Be],
              raw: String(activeSheet.cells[Be].raw).replace(
                fe,
                () => replaceWithText,
              ),
            },
          }));
      (Oe.length && applyCellPatches(Oe, "Replace text in selection"),
        setStatusNotice(`Replaced text in ${Oe.length} cells.`));
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
