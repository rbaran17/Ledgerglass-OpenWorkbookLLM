import i from "react/jsx-runtime";
import { renderFormulaBar } from "./render-formula-bar.js";
import {
  collectStyleFallbackWarnings,
  describeNamedStyleFallback,
  findDataValidationsInRange,
  getImportedNumberFormatWarning,
  setSplitPanes,
} from "../03-workbook-model.js";
import { LocalizedText } from "../04-cell-input-paste-and-fill.js";
import { Filter, Search, X } from "lucide-react";
import { PivotResultsStatusBar } from "../29-worksheet-tools-and-pivot-panels.js";
import { SplitPaneLayout } from "../53-find-replace-go-to.js";
import { renderSheetTabs } from "./render-sheet-tabs.js";
import { renderStatusBar } from "./render-status-bar.js";
import { saveBlobToFile } from "../13-consolidation-dialog.js";

function renderWorkbookGrid({
  readOnly,
  setIsWorkbookHealthDialogOpen,
  text,
  language,
  editorIdPrefix,
  activeWorkbook,
  activeSheetId,
  activeSheet,
  isEditingLocked,
  setActiveDialogMode,
  appRootRef,
  setTabBarSheetId,
  formulaInputRef,
  formulaSuggestionListRef,
  formulaCaretPosition,
  setFormulaCaretPosition,
  isFormulaBarFocused,
  setIsFormulaBarFocused,
  activeCellAddress,
  completedSelectionAreas,
  selectionAreas,
  goToReferenceText,
  setGoToReferenceText,
  cellEditText,
  activeCellEditRef,
  formulaPointerStateRef,
  activeFormulaReference,
  isFormulaPointMode,
  filterValueSearch,
  setFilterValueSearch,
  setFilterStateByScope,
  setAuditPopulationScope,
  filterStateKey,
  activeFilterState,
  activeColumnSort,
  sheetSplitPanes,
  splitPaneSizes,
  scrollActivePaneVertically,
  isOperationBusy,
  activeOperationControllerRef,
  setPromptDialogState,
  zoomPercent,
  isFilterBarVisible,
  setIsFilterBarVisible,
  isFindBarVisible,
  setIsFindBarVisible,
  findQuery,
  setFindQuery,
  replaceWithText,
  setReplaceWithText,
  fillDragState,
  clipboardSnapshotState,
  clearCutSession,
  workbookCalculator,
  filterApplication,
  filterResult,
  isFilterStale,
  gridContainerRef,
  findInputRef,
  latestEditStateRef,
  selectedCellCount,
  activeCellFormulaErrors,
  activeCellReferenceDiagnostics,
  activeCellNamedFormulaDiagnostics,
  selectedNumericValues,
  selectedNumericSum,
  activeCell,
  matchingRowCount,
  runEditorCommand,
  commitWorkbookEdit,
  runGuardedAction,
  applyCellPatches,
  syncFormulaEditorSelection,
  isFormulaEditingLocked,
  handleFormulaKeyDown,
  beginCellEdit,
  focusGridSurface,
  cancelCellEdit,
  commitCellEdit,
  applyZoomPercent,
  switchActiveWorksheet,
  commitNameBoxReference,
  findNextTextMatch,
  replaceTextInSelection,
  moveSelectionAfterCommit,
  renderGridPane,
}) {
  return i.jsxs("section", {
    className: "lg-workbook",
    "aria-label": text("Spreadsheet workbook"),
    lang: language("Spreadsheet workbook"),
    "data-ui-label": "Spreadsheet workbook",
    children: [
      renderFormulaBar({
        text,
        language,
        editorIdPrefix,
        activeWorkbook,
        activeSheet,
        isEditingLocked,
        formulaInputRef,
        formulaSuggestionListRef,
        formulaCaretPosition,
        setFormulaCaretPosition,
        isFormulaBarFocused,
        setIsFormulaBarFocused,
        goToReferenceText,
        setGoToReferenceText,
        cellEditText,
        activeCellEditRef,
        formulaPointerStateRef,
        gridContainerRef,
        syncFormulaEditorSelection,
        isFormulaEditingLocked,
        handleFormulaKeyDown,
        beginCellEdit,
        focusGridSurface,
        cancelCellEdit,
        commitCellEdit,
        commitNameBoxReference,
        moveSelectionAfterCommit,
      }),
      activeCellEditRef.current?.value.startsWith("=") &&
        i.jsxs("div", {
          className: "lg-formula-entry-help",
          id: `${editorIdPrefix}-formula-help`,
          lang: "en",
          children: [
            i.jsxs("span", {
              children: [
                "Editing",
                " ",
                activeWorkbook.sheets.find(
                  (fe) => fe.id === activeCellEditRef.current?.sheetId,
                )?.name,
                "!",
                activeCellEditRef.current?.address,
                " ·",
                " ",
                isFormulaPointMode ? "Point mode" : "Edit mode",
                " · Select cells to insert a reference. F2 switches arrow keys; F4 locks references. Enter applies; Esc cancels.",
              ],
            }),
            i.jsx("span", {
              role: "status",
              "aria-live": "polite",
              "aria-atomic": "true",
              "aria-label": "Formula reference",
              children: activeFormulaReference
                ? `Reference: ${activeWorkbook.sheets.find((fe) => fe.id === activeFormulaReference.sheetId)?.name}!${activeFormulaReference.range}`
                : "Formula not yet applied",
            }),
          ],
        }),
      findDataValidationsInRange(
        activeWorkbook,
        activeSheet.id,
        activeCellAddress,
      ).map((fe) =>
        i.jsxs(
          "div",
          {
            className: "lg-view-note",
            "aria-label": text("Validation input guidance"),
            lang: language("Validation input guidance"),
            "data-ui-label": "Validation input guidance",
            children: [
              fe.showInputMessage !== false &&
                (fe.inputTitle || fe.inputMessage) &&
                i.jsxs("span", {
                  children: [
                    i.jsx("strong", { children: fe.inputTitle }),
                    " ",
                    fe.inputMessage,
                    " ",
                  ],
                }),
              fe.kind === "list" &&
                fe.showDropdown !== false &&
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Choose a value for ",
                    }),
                    activeCellAddress,
                    " ",
                    i.jsxs("select", {
                      "aria-label": `Validation dropdown for ${activeCellAddress}`,
                      disabled: isEditingLocked,
                      value: "",
                      onChange: (Oe) => {
                        const Be = fe.list[Number(Oe.target.value)];
                        Be !== undefined &&
                          runGuardedAction(() => {
                            (cancelCellEdit(),
                              applyCellPatches(
                                [
                                  {
                                    sheetId: activeSheet.id,
                                    address: activeCellAddress,
                                    cell: {
                                      ...activeSheet.cells[activeCellAddress],
                                      raw: Be,
                                      literal: true,
                                    },
                                  },
                                ],
                                "Choose validation list value",
                              ));
                          });
                      },
                      children: [
                        i.jsx("option", {
                          value: "",
                          disabled: true,
                          children: "Select…",
                        }),
                        fe.list.map((Oe, Be) =>
                          i.jsx("option", { value: Be, children: Oe }, Be),
                        ),
                      ],
                    }),
                  ],
                }),
            ],
          },
          fe.id,
        ),
      ),
      workbookCalculator.getSpillDiagnostic(
        activeSheet.id,
        activeCellAddress,
      ) &&
        i.jsxs("div", {
          className: "lg-format-warning",
          role: "status",
          "aria-label": text("Array spill diagnostic"),
          lang: language("Array spill diagnostic"),
          "data-ui-label": "Array spill diagnostic",
          children: [
            activeCellAddress,
            ": ",
            workbookCalculator.getSpillDiagnostic(
              activeSheet.id,
              activeCellAddress,
            ).message,
          ],
        }),
      (workbookCalculator.getSpillAnchor(activeSheet.id, activeCellAddress) ||
        workbookCalculator.getSpillRange(activeSheet.id, activeCellAddress)) &&
        i.jsxs("div", {
          className: "lg-view-note",
          role: "status",
          "aria-label": text("Dynamic array range"),
          lang: language("Dynamic array range"),
          "data-ui-label": "Dynamic array range",
          children: [
            i.jsx(LocalizedText, { children: "Dynamic array" }),
            " ",
            workbookCalculator.getSpillRange(
              activeSheet.id,
              workbookCalculator.getSpillAnchor(
                activeSheet.id,
                activeCellAddress,
              ) ?? activeCellAddress,
            ),
            ".",
            workbookCalculator.getSpillAnchor(activeSheet.id, activeCellAddress)
              ? ` Edit the formula in ${workbookCalculator.getSpillAnchor(activeSheet.id, activeCellAddress)}; this cell is a calculated result.`
              : " Results resize when source data changes.",
          ],
        }),
      activeCellNamedFormulaDiagnostics.length > 0 &&
        i.jsxs("div", {
          className: "lg-format-warning",
          role: "status",
          "aria-label": text("Named formula diagnostic"),
          lang: language("Named formula diagnostic"),
          "data-ui-label": "Named formula diagnostic",
          children: [
            activeCellAddress,
            ":",
            " ",
            activeCellNamedFormulaDiagnostics.map((fe) => fe.message).join(" "),
          ],
        }),
      activeCellReferenceDiagnostics.length > 0 &&
        i.jsxs("div", {
          className: "lg-format-warning",
          role: "status",
          "aria-label": text("Formula reference diagnostic"),
          lang: language("Formula reference diagnostic"),
          "data-ui-label": "Formula reference diagnostic",
          children: [
            activeCellAddress,
            ":",
            " ",
            activeCellReferenceDiagnostics.map((fe) => fe.message).join(" "),
          ],
        }),
      activeCellFormulaErrors.length > 0 &&
        i.jsxs("div", {
          className: "lg-format-warning",
          role: "status",
          children: [
            activeCellAddress,
            ":",
            " ",
            activeCellFormulaErrors.map((fe) => fe.message).join(" "),
            " ",
            i.jsx("button", {
              type: "button",
              onClick: () => runEditorCommand("conditional"),
              children: i.jsx(LocalizedText, {
                children: "Review conditional rules",
              }),
            }),
          ],
        }),
      (describeNamedStyleFallback(activeCell?.style) ||
        activeWorkbook.styleCatalog?.missingTheme) &&
        i.jsxs("div", {
          className: "lg-format-warning",
          role: "status",
          children: [
            describeNamedStyleFallback(activeCell?.style) ||
              collectStyleFallbackWarnings(activeWorkbook).join(" "),
            i.jsx("button", {
              onClick: () => runEditorCommand("themes"),
              children: i.jsx(LocalizedText, {
                children: "Review themes and styles",
              }),
            }),
          ],
        }),
      getImportedNumberFormatWarning(activeSheet.cells[activeCellAddress]) &&
        i.jsxs("div", {
          className: "lg-view-note",
          role: "status",
          children: [
            activeCellAddress,
            ": ",
            getImportedNumberFormatWarning(
              activeSheet.cells[activeCellAddress],
            ),
          ],
        }),
      activeFilterState &&
        i.jsxs("div", {
          className: "lg-population-context",
          role: "status",
          "aria-label": text("Population filter context"),
          lang: language("Population filter context"),
          "data-ui-label": "Population filter context",
          children: [
            i.jsx(Filter, { size: 14 }),
            i.jsx("strong", {
              children: isFilterStale
                ? "Filter paused — workbook changed"
                : filterApplication.error
                  ? "Filter unavailable — source rows shown"
                  : `${filterResult?.context.visibleRows} visible / ${filterResult?.context.totalRows} total data rows`,
            }),
            filterResult &&
              i.jsxs("span", {
                children: [
                  filterResult.context.range,
                  " ·",
                  " ",
                  filterResult.context.filteredRows,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "filtered out ·",
                  }),
                  " ",
                  filterResult.context.hiddenMatchingRows,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "matching rows manually hidden",
                  }),
                ],
              }),
            i.jsx("button", {
              onClick: () => runEditorCommand("populationFilter"),
              children: isFilterStale ? "Review and reapply" : "Edit filters",
            }),
            i.jsx("button", {
              onClick: () => {
                (setFilterStateByScope((fe) => {
                  const Oe = { ...fe };
                  return (delete Oe[filterStateKey], Oe);
                }),
                  setAuditPopulationScope("full"));
              },
              children: i.jsx(LocalizedText, {
                children: "Clear population filters",
              }),
            }),
            filterResult &&
              i.jsx("button", {
                onClick: () =>
                  saveBlobToFile(
                    new Blob([JSON.stringify(filterResult.context, null, 2)], {
                      type: "application/json",
                    }),
                    "Population-filter-context.json",
                  ),
                children: i.jsx(LocalizedText, {
                  children: "Save filter context",
                }),
              }),
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(LocalizedText, {
                    children: "Population scope",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Column filters intersect; advanced groups use OR. These are session views; named criteria can be saved in the native workbook through Edit filters. Native JSON, XLSX, print, other analysis tools and AI selections include full source data in their chosen ranges. CSV/TSV and Audit Checks offer full or visible scope. View filtering does not change SUM or ordinary formulas; SUBTOTAL/AGGREGATE retain their existing worksheet visibility semantics.",
                  }),
                }),
              ],
            }),
            filterApplication.error &&
              i.jsx("span", {
                role: "alert",
                children: filterApplication.error,
              }),
          ],
        }),
      filterResult?.context.visibleRows === 0 &&
        i.jsxs("div", {
          className: "lg-population-empty",
          role: "status",
          children: [
            i.jsx(LocalizedText, {
              children: "No matching visible records. All",
            }),
            " ",
            filterResult.context.totalRows,
            " ",
            i.jsx(LocalizedText, {
              children:
                "source data rows are retained. Clear population filters to show them again; manually hidden rows remain hidden.",
            }),
          ],
        }),
      isFilterBarVisible &&
        i.jsxs("div", {
          className: "lg-find-bar",
          children: [
            i.jsx(Filter, { size: 14 }),
            i.jsx("input", {
              "aria-label": text("Filter rows"),
              placeholder: text("Filter rows (first row is a header)"),
              value: filterValueSearch,
              onChange: (fe) => {
                (setFilterValueSearch(fe.target.value),
                  setFilterStateByScope((Oe) => {
                    const Be = { ...Oe };
                    return (delete Be[filterStateKey], Be);
                  }),
                  scrollActivePaneVertically(0),
                  gridContainerRef.current &&
                    (gridContainerRef.current.scrollTop = 0));
              },
              lang: language("Filter rows"),
              "data-ui-label": "Filter rows",
            }),
            i.jsx("button", {
              "aria-label": text("Close filter"),
              onClick: () => {
                (setIsFilterBarVisible(false),
                  setFilterValueSearch(""),
                  focusGridSurface());
              },
              lang: language("Close filter"),
              "data-ui-label": "Close filter",
              children: i.jsx(X, { size: 14 }),
            }),
          ],
        }),
      i.jsx(PivotResultsStatusBar, {
        workbook: activeWorkbook,
        sheetId: activeSheetId,
        onOpen: () => runEditorCommand("pivot"),
        onExplore: () => runEditorCommand("pivotExplore"),
      }),
      isFindBarVisible &&
        i.jsxs("div", {
          className: "lg-find-bar",
          children: [
            i.jsx(Search, { size: 14 }),
            i.jsx("input", {
              ref: findInputRef,
              "aria-label": text("Find text"),
              placeholder: text("Find"),
              value: findQuery,
              onChange: (fe) => setFindQuery(fe.target.value),
              onKeyDown: (fe) => {
                fe.key === "Enter" && findNextTextMatch();
              },
              lang: language("Find text"),
              "data-ui-label": "Find text",
            }),
            i.jsx("button", {
              onClick: findNextTextMatch,
              children: i.jsx(LocalizedText, {
                children: "Find Next",
              }),
            }),
            i.jsx("input", {
              "aria-label": text("Replace text"),
              placeholder: text("Replace with"),
              value: replaceWithText,
              onChange: (fe) => setReplaceWithText(fe.target.value),
              lang: language("Replace text"),
              "data-ui-label": "Replace text",
            }),
            i.jsx("button", {
              disabled: isEditingLocked,
              onClick: replaceTextInSelection,
              children: i.jsx(LocalizedText, {
                children: "Replace in selection",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close find"),
              onClick: () => {
                (setIsFindBarVisible(false), focusGridSurface());
              },
              lang: language("Close find"),
              "data-ui-label": "Close find",
              children: i.jsx(X, { size: 14 }),
            }),
          ],
        }),
      fillDragState &&
        i.jsxs("div", {
          className: "lg-view-note",
          role: "status",
          children: [
            i.jsx(LocalizedText, { children: "Fill preview · " }),
            fillDragState.targetRange,
            " ",
            i.jsx(LocalizedText, {
              children: "· Release to review, or press Escape to cancel",
            }),
          ],
        }),
      filterValueSearch &&
        matchingRowCount === 0 &&
        i.jsx("div", {
          className: "lg-population-empty",
          role: "status",
          children: i.jsx(LocalizedText, {
            children:
              "No rows match this search. Source data is unchanged; clear the search to restore the view.",
          }),
        }),
      (filterValueSearch || activeColumnSort) &&
        i.jsxs("div", {
          className: "lg-view-note",
          children: [
            i.jsx(Filter, { size: 12 }),
            " ",
            i.jsx(LocalizedText, {
              children:
                "View only · source row numbers and formula references stay fixed",
            }),
            filterValueSearch && ` · ${matchingRowCount} matching rows`,
          ],
        }),
      clipboardSnapshotState &&
        i.jsxs("div", {
          className: "lg-view-note",
          role: "status",
          children: [
            i.jsx(LocalizedText, { children: "Cut ready:" }),
            " ",
            activeWorkbook.sheets.find(
              (fe) => fe.id === clipboardSnapshotState.snapshot.sheetId,
            )?.name ?? "Deleted worksheet",
            "!",
            clipboardSnapshotState.snapshot.range,
            " ",
            i.jsx(LocalizedText, {
              children: "· Paste to review the move.",
            }),
            " ",
            i.jsx("button", {
              type: "button",
              onClick: clearCutSession,
              children: i.jsx(LocalizedText, {
                children: "Cancel cut",
              }),
            }),
          ],
        }),
      i.jsx(SplitPaneLayout, {
        split: sheetSplitPanes,
        snapshot: activeWorkbook,
        readOnly: isEditingLocked,
        active: splitPaneSizes.activeId,
        onResize: (fe) =>
          runGuardedAction(() =>
            commitWorkbookEdit(
              setSplitPanes(
                latestEditStateRef.current.book,
                activeSheet.id,
                fe,
              ),
              "Resize split panes",
            ),
          ),
        renderPane: renderGridPane,
      }),
      renderSheetTabs({
        readOnly,
        text,
        language,
        activeWorkbook,
        activeSheet,
        setActiveDialogMode,
        appRootRef,
        setTabBarSheetId,
        activeCellEditRef,
        setPromptDialogState,
        runEditorCommand,
        commitCellEdit,
        switchActiveWorksheet,
      }),
      renderStatusBar({
        setIsWorkbookHealthDialogOpen,
        text,
        language,
        isEditingLocked,
        completedSelectionAreas,
        selectionAreas,
        isOperationBusy,
        activeOperationControllerRef,
        zoomPercent,
        selectedCellCount,
        selectedNumericValues,
        selectedNumericSum,
        applyZoomPercent,
      }),
    ],
  });
}

export { renderWorkbookGrid };
