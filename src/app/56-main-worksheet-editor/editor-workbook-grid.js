import i from "react/jsx-runtime";
import { EditorFormulaBar } from "./editor-formula-bar.js";
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
import { EditorSheetTabs } from "./editor-sheet-tabs.js";
import { EditorStatusBar } from "./editor-status-bar.js";
import { saveBlobToFile } from "../13-consolidation-dialog.js";

function EditorWorkbookGrid({
  viewAndSelectionActions,
  gridInteractionActions,
  navigationActions,
  cellEditActions,
  formulaEntryActions,
  clipboardAndFindActions,
  dialogState,
  selectionState,
  navigationState,
  formulaEntryState,
  filterSortState,
  auditState,
  operationStatusState,
  viewDisplayState,
  findReplaceState,
  dragFillState,
  clipboardState,
  readOnly,
  text,
  language,
  editorIdPrefix,
  activeWorkbook,
  activeSheetId,
  activeSheet,
  isEditingLocked,
  selectionAreas,
  filterStateKey,
  activeFilterState,
  sheetSplitPanes,
  splitPaneSizes,
  workbookCalculator,
  filterApplication,
  filterResult,
  isFilterStale,
  latestEditStateRef,
  selectedCellCount,
  activeCellFormulaErrors,
  activeCellReferenceDiagnostics,
  activeCellNamedFormulaDiagnostics,
  selectedNumericValues,
  selectedNumericSum,
  activeCell,
  matchingRowCount,
  renderGridPane,
}) {
  return i.jsxs("section", {
    className: "lg-workbook",
    "aria-label": text("Spreadsheet workbook"),
    lang: language("Spreadsheet workbook"),
    "data-ui-label": "Spreadsheet workbook",
    children: [
      i.jsx(EditorFormulaBar, {
        cellEditActions,
        formulaEntryActions,
        viewAndSelectionActions,
        navigationActions,
        formulaEntryState,
        navigationState,
        selectionState,
        text,
        language,
        editorIdPrefix,
        activeWorkbook,
        activeSheet,
        isEditingLocked,
      }),
      formulaEntryState.activeCellEditRef.current?.value.startsWith("=") &&
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
                  (sheet) =>
                    sheet.id ===
                    formulaEntryState.activeCellEditRef.current?.sheetId,
                )?.name,
                "!",
                formulaEntryState.activeCellEditRef.current?.address,
                " ·",
                " ",
                formulaEntryState.isFormulaPointMode
                  ? "Point mode"
                  : "Edit mode",
                " · Select cells to insert a reference. F2 switches arrow keys; F4 locks references. Enter applies; Esc cancels.",
              ],
            }),
            i.jsx("span", {
              role: "status",
              "aria-live": "polite",
              "aria-atomic": "true",
              "aria-label": "Formula reference",
              children: formulaEntryState.activeFormulaReference
                ? `Reference: ${activeWorkbook.sheets.find((sheet) => sheet.id === formulaEntryState.activeFormulaReference.sheetId)?.name}!${formulaEntryState.activeFormulaReference.range}`
                : "Formula not yet applied",
            }),
          ],
        }),
      findDataValidationsInRange(
        activeWorkbook,
        activeSheet.id,
        selectionState.activeCellAddress,
      ).map((dataValidation) =>
        i.jsxs(
          "div",
          {
            className: "lg-view-note",
            "aria-label": text("Validation input guidance"),
            lang: language("Validation input guidance"),
            "data-ui-label": "Validation input guidance",
            children: [
              dataValidation.showInputMessage !== false &&
                (dataValidation.inputTitle || dataValidation.inputMessage) &&
                i.jsxs("span", {
                  children: [
                    i.jsx("strong", { children: dataValidation.inputTitle }),
                    " ",
                    dataValidation.inputMessage,
                    " ",
                  ],
                }),
              dataValidation.kind === "list" &&
                dataValidation.showDropdown !== false &&
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Choose a value for ",
                    }),
                    selectionState.activeCellAddress,
                    " ",
                    i.jsxs("select", {
                      "aria-label": `Validation dropdown for ${selectionState.activeCellAddress}`,
                      disabled: isEditingLocked,
                      value: "",
                      onChange: (event) => {
                        const raw =
                          dataValidation.list[Number(event.target.value)];
                        raw !== undefined &&
                          viewAndSelectionActions.runGuardedAction(() => {
                            (cellEditActions.cancelCellEdit(),
                              cellEditActions.applyCellPatches(
                                [
                                  {
                                    sheetId: activeSheet.id,
                                    address: selectionState.activeCellAddress,
                                    cell: {
                                      ...activeSheet.cells[
                                        selectionState.activeCellAddress
                                      ],
                                      raw: raw,
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
                        dataValidation.list.map((optionLabel, optionIndex) =>
                          i.jsx(
                            "option",
                            { value: optionIndex, children: optionLabel },
                            optionIndex,
                          ),
                        ),
                      ],
                    }),
                  ],
                }),
            ],
          },
          dataValidation.id,
        ),
      ),
      workbookCalculator.getSpillDiagnostic(
        activeSheet.id,
        selectionState.activeCellAddress,
      ) &&
        i.jsxs("div", {
          className: "lg-format-warning",
          role: "status",
          "aria-label": text("Array spill diagnostic"),
          lang: language("Array spill diagnostic"),
          "data-ui-label": "Array spill diagnostic",
          children: [
            selectionState.activeCellAddress,
            ": ",
            workbookCalculator.getSpillDiagnostic(
              activeSheet.id,
              selectionState.activeCellAddress,
            ).message,
          ],
        }),
      (workbookCalculator.getSpillAnchor(
        activeSheet.id,
        selectionState.activeCellAddress,
      ) ||
        workbookCalculator.getSpillRange(
          activeSheet.id,
          selectionState.activeCellAddress,
        )) &&
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
                selectionState.activeCellAddress,
              ) ?? selectionState.activeCellAddress,
            ),
            ".",
            workbookCalculator.getSpillAnchor(
              activeSheet.id,
              selectionState.activeCellAddress,
            )
              ? ` Edit the formula in ${workbookCalculator.getSpillAnchor(activeSheet.id, selectionState.activeCellAddress)}; this cell is a calculated result.`
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
            selectionState.activeCellAddress,
            ":",
            " ",
            activeCellNamedFormulaDiagnostics
              .map(
                (activeCellNamedFormulaDiagnostic) =>
                  activeCellNamedFormulaDiagnostic.message,
              )
              .join(" "),
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
            selectionState.activeCellAddress,
            ":",
            " ",
            activeCellReferenceDiagnostics
              .map(
                (activeCellReferenceDiagnostic) =>
                  activeCellReferenceDiagnostic.message,
              )
              .join(" "),
          ],
        }),
      activeCellFormulaErrors.length > 0 &&
        i.jsxs("div", {
          className: "lg-format-warning",
          role: "status",
          children: [
            selectionState.activeCellAddress,
            ":",
            " ",
            activeCellFormulaErrors
              .map((activeCellFormulaError) => activeCellFormulaError.message)
              .join(" "),
            " ",
            i.jsx("button", {
              type: "button",
              onClick: () =>
                gridInteractionActions.runEditorCommand("conditional"),
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
              onClick: () => gridInteractionActions.runEditorCommand("themes"),
              children: i.jsx(LocalizedText, {
                children: "Review themes and styles",
              }),
            }),
          ],
        }),
      getImportedNumberFormatWarning(
        activeSheet.cells[selectionState.activeCellAddress],
      ) &&
        i.jsxs("div", {
          className: "lg-view-note",
          role: "status",
          children: [
            selectionState.activeCellAddress,
            ": ",
            getImportedNumberFormatWarning(
              activeSheet.cells[selectionState.activeCellAddress],
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
              onClick: () =>
                gridInteractionActions.runEditorCommand("populationFilter"),
              children: isFilterStale ? "Review and reapply" : "Edit filters",
            }),
            i.jsx("button", {
              onClick: () => {
                (filterSortState.setFilterStateByScope((filterStateMap) => {
                  const remainingFilterStates = { ...filterStateMap };
                  return (
                    delete remainingFilterStates[filterStateKey],
                    remainingFilterStates
                  );
                }),
                  auditState.setAuditPopulationScope("full"));
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
      filterSortState.isFilterBarVisible &&
        i.jsxs("div", {
          className: "lg-find-bar",
          children: [
            i.jsx(Filter, { size: 14 }),
            i.jsx("input", {
              "aria-label": text("Filter rows"),
              placeholder: text("Filter rows (first row is a header)"),
              value: filterSortState.filterValueSearch,
              onChange: (event) => {
                (filterSortState.setFilterValueSearch(event.target.value),
                  filterSortState.setFilterStateByScope(
                    (previousFilterStates) => {
                      const updatedFilterStates = { ...previousFilterStates };
                      return (
                        delete updatedFilterStates[filterStateKey],
                        updatedFilterStates
                      );
                    },
                  ),
                  viewAndSelectionActions.scrollActivePaneVertically(0),
                  selectionState.gridContainerRef.current &&
                    (selectionState.gridContainerRef.current.scrollTop = 0));
              },
              lang: language("Filter rows"),
              "data-ui-label": "Filter rows",
            }),
            i.jsx("button", {
              "aria-label": text("Close filter"),
              onClick: () => {
                (filterSortState.setIsFilterBarVisible(false),
                  filterSortState.setFilterValueSearch(""),
                  viewAndSelectionActions.focusGridSurface());
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
        onOpen: () => gridInteractionActions.runEditorCommand("pivot"),
        onExplore: () =>
          gridInteractionActions.runEditorCommand("pivotExplore"),
      }),
      findReplaceState.isFindBarVisible &&
        i.jsxs("div", {
          className: "lg-find-bar",
          children: [
            i.jsx(Search, { size: 14 }),
            i.jsx("input", {
              ref: findReplaceState.findInputRef,
              "aria-label": text("Find text"),
              placeholder: text("Find"),
              value: findReplaceState.findQuery,
              onChange: (event) =>
                findReplaceState.setFindQuery(event.target.value),
              onKeyDown: (event) => {
                event.key === "Enter" &&
                  clipboardAndFindActions.findNextTextMatch();
              },
              lang: language("Find text"),
              "data-ui-label": "Find text",
            }),
            i.jsx("button", {
              onClick: clipboardAndFindActions.findNextTextMatch,
              children: i.jsx(LocalizedText, {
                children: "Find Next",
              }),
            }),
            i.jsx("input", {
              "aria-label": text("Replace text"),
              placeholder: text("Replace with"),
              value: findReplaceState.replaceWithText,
              onChange: (event) =>
                findReplaceState.setReplaceWithText(event.target.value),
              lang: language("Replace text"),
              "data-ui-label": "Replace text",
            }),
            i.jsx("button", {
              disabled: isEditingLocked,
              onClick: clipboardAndFindActions.replaceTextInSelection,
              children: i.jsx(LocalizedText, {
                children: "Replace in selection",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close find"),
              onClick: () => {
                (findReplaceState.setIsFindBarVisible(false),
                  viewAndSelectionActions.focusGridSurface());
              },
              lang: language("Close find"),
              "data-ui-label": "Close find",
              children: i.jsx(X, { size: 14 }),
            }),
          ],
        }),
      dragFillState.fillDragState &&
        i.jsxs("div", {
          className: "lg-view-note",
          role: "status",
          children: [
            i.jsx(LocalizedText, { children: "Fill preview · " }),
            dragFillState.fillDragState.targetRange,
            " ",
            i.jsx(LocalizedText, {
              children: "· Release to review, or press Escape to cancel",
            }),
          ],
        }),
      filterSortState.filterValueSearch &&
        matchingRowCount === 0 &&
        i.jsx("div", {
          className: "lg-population-empty",
          role: "status",
          children: i.jsx(LocalizedText, {
            children:
              "No rows match this search. Source data is unchanged; clear the search to restore the view.",
          }),
        }),
      (filterSortState.filterValueSearch || filterSortState.activeColumnSort) &&
        i.jsxs("div", {
          className: "lg-view-note",
          children: [
            i.jsx(Filter, { size: 12 }),
            " ",
            i.jsx(LocalizedText, {
              children:
                "View only · source row numbers and formula references stay fixed",
            }),
            filterSortState.filterValueSearch &&
              ` · ${matchingRowCount} matching rows`,
          ],
        }),
      clipboardState.clipboardSnapshotState &&
        i.jsxs("div", {
          className: "lg-view-note",
          role: "status",
          children: [
            i.jsx(LocalizedText, { children: "Cut ready:" }),
            " ",
            activeWorkbook.sheets.find(
              (sheet) =>
                sheet.id ===
                clipboardState.clipboardSnapshotState.snapshot.sheetId,
            )?.name ?? "Deleted worksheet",
            "!",
            clipboardState.clipboardSnapshotState.snapshot.range,
            " ",
            i.jsx(LocalizedText, {
              children: "· Paste to review the move.",
            }),
            " ",
            i.jsx("button", {
              type: "button",
              onClick: viewAndSelectionActions.clearCutSession,
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
        onResize: (splitPanes) =>
          viewAndSelectionActions.runGuardedAction(() =>
            navigationActions.commitWorkbookEdit(
              setSplitPanes(
                latestEditStateRef.current.book,
                activeSheet.id,
                splitPanes,
              ),
              "Resize split panes",
            ),
          ),
        renderPane: renderGridPane,
      }),
      i.jsx(EditorSheetTabs, {
        gridInteractionActions,
        cellEditActions,
        navigationActions,
        dialogState,
        selectionState,
        navigationState,
        formulaEntryState,
        readOnly,
        text,
        language,
        activeWorkbook,
        activeSheet,
      }),
      i.jsx(EditorStatusBar, {
        navigationActions,
        dialogState,
        selectionState,
        operationStatusState,
        viewDisplayState,
        text,
        language,
        isEditingLocked,
        selectionAreas,
        selectedCellCount,
        selectedNumericValues,
        selectedNumericSum,
      }),
    ],
  });
}

export { EditorWorkbookGrid };
