import i from "react/jsx-runtime";
import {
  AutomationScriptsDialog,
  preserveAutomationArtifact,
  removeAutomationArtifact,
  runApprovedMigration,
} from "../41-automation-and-scripts.js";
import { FindReplaceDialog } from "../53-find-replace-go-to.js";
import {
  ConditionalFormattingRulesManager,
  FormatCellsDialog,
} from "../23-format-cells-and-conditional-formatting.js";
import { ThemesAndCellStylesDialog } from "../22-themes-and-cell-styles.js";
import { MergeCellsPreviewDialog } from "../24-merge-audit-and-ribbon-helpers.js";
import { M4e } from "../48-move-cells-and-reference-repair.js";
import {
  OutlineToolsPanel,
  ReviewToolsPanel,
} from "../49-review-tools-and-outlines.js";
import { PasteSpecialDialog } from "../47-fill-and-paste-special.js";
import {
  CustomSortDialog,
  WorksheetTabBar,
} from "../45-sorting-and-worksheets.js";
import { WorkbookManagerDialog } from "../46-workbook-manager.js";
import { TableToolsDialog } from "../30-table-tools-dialog.js";
import { PivotExplorerDialog } from "../31-pivot-explorer-and-charts.js";
import { PicturePropertiesDialog } from "../32-drawing-objects-and-printing.js";
import { PageSetupDialog } from "../33-page-setup-and-print.js";
import {
  CalculationInspectorDialog,
  EvaluateFormulaDialog,
  FormulaCompatibilityDialog,
} from "../35-formula-inspection-and-compatibility.js";
import { FunctionAndNameDialog } from "../34-function-browser-and-names.js";
import { FilterViewsDialog } from "../19-filter-views-dialog.js";
import {
  ExternalDataRefreshDialog,
  chainMaybePromise,
  isPromiseLike,
} from "../55-external-source-refresh.js";
import { DataAnalysisDialog } from "../38-data-analysis-dialog.js";
import { AddInsAndCustomFunctionsDialog } from "../05-add-ins-and-custom-functions.js";
import { WorkbookLinksDialog } from "../09-workbook-links.js";
import { SourceImportDialog } from "../10-source-import-dialog.js";
import { FormatConversionDialog } from "../11-local-workbook-store-and-versions.js";
import { ReportOutputsDialog } from "../39-report-outputs-and-what-if.js";
import { DataModelDialog } from "../15-data-model-dialog.js";
import { PatternCheckDialog } from "../51-pattern-checks-and-citations.js";
import {
  applyFilterConditions,
  formatOccupiedRangeAddress,
} from "../03-workbook-model.js";
import { AgingScheduleScreen } from "../21-aging-and-materiality-screens.js";
import { MaterialityRiskScreen } from "../50-materiality-screen-and-pattern-checks.js";
import { WorkbookComparisonDialog } from "../52-workbook-comparison.js";
import { WhatIfAnalysisDialog } from "../40-what-if-and-delimited-dialogs.js";
import {
  FormulaWatchWindow,
  GoalSeekDialog,
} from "../36-watch-window-and-goal-seek.js";
import { CalculationOptionsDialog } from "../20-calculation-options-and-assistant.js";
import { DataToolDialogRouter } from "../29-worksheet-tools-and-pivot-panels.js";
import {
  ConsolidationDialog,
  saveBlobToFile,
} from "../13-consolidation-dialog.js";
import { encodeNativeWorkbook } from "../08-xlsx-import-export-and-preservation.js";
import { importAutomationFile } from "../07-file-import-and-worker-dispatch.js";
import { TransformationRecipesDialog } from "../18-transformation-recipes-dialog.js";
import {
  CombinePopulationsDialog,
  TextToColumnsDialog,
} from "../42-combine-populations-and-text-columns.js";
import { TextCleanupDialog } from "../43-text-cleanup.js";
import { CleanAndReshapeDialog } from "../44-clean-and-reshape.js";
import { DuplicateReviewDialog } from "../28-duplicate-review-and-structure.js";

function EditorDialogs({
  persistenceState,
  dialogState,
  navigationState,
  formulaEntryState,
  selectionState,
  filterSortState,
  auditState,
  operationStatusState,
  importExportState,
  clipboardState,
  extensions,
  allowLinkedWorkbookFiles,
  documentExtractor,
  externalSourceAdapter,
  consolidationDefinition,
  allowConsolidationFiles,
  actor,
  readOnly,
  persistenceKey,
  workbookStore,
  workbookStorePersistent,
  activeReliabilityMonitor,
  activeWorkbook,
  effectiveLinkedWorkbooks,
  setLinkedWorkbooks,
  effectiveConsolidationWorkbooks,
  setConsolidationWorkbooks,
  activeSheet,
  workbookPersistence,
  isEditingLocked,
  selectSingleRange,
  selectionAreas,
  filterStateKey,
  activeFilterState,
  scrollActivePaneVertically,
  operationScope,
  calculationRun,
  workbookCalculator,
  filterResult,
  latestEditStateRef,
  handleCalculationCompleted,
  runCalculationNow,
  applyWorkbookView,
  reportSaveStatus,
  enqueueSaveOperation,
  acknowledgeSavedWorkbook,
  persistWorkbookSnapshot,
  commitWorkbookChange,
  commitWorkbookEdit,
  replaceWorkbookSnapshot,
  commitWorkbookWithNewSheet,
  beginCellEdit,
  switchActiveWorksheet,
  navigateToSheetCell,
  selectSheetRanges,
}) {
  return (
    dialogState.activeDialogMode &&
    (() => {
      const fe = {
        mode: dialogState.activeDialogMode,
        workbook: activeWorkbook,
        sheetId: activeSheet.id,
        range: selectionState.activeSelectionRange,
        readOnly:
          dialogState.activeDialogMode === "protection"
            ? readOnly
            : isEditingLocked,
        onClose: () => dialogState.setActiveDialogMode(null),
        onWorkbookChange: commitWorkbookEdit,
      };
      if (dialogState.activeDialogMode === "automation")
        return i.jsx(
          AutomationScriptsDialog,
          {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            range: selectionState.activeSelectionRange,
            readOnly: isEditingLocked,
            initialInspection: dialogState.automationInspection,
            onClose: fe.onClose,
            onPreserve: (Ue) => {
              const Xe = latestEditStateRef.current.book,
                gt = preserveAutomationArtifact(Xe, Ue, {
                  actor,
                  readOnly: latestEditStateRef.current.readOnly,
                });
              (gt !== Xe && commitWorkbookChange(gt),
                operationStatusState.setStatusNotice(
                  "Original automation bytes preserved in native JSON. No source code was executed.",
                ));
            },
            onRemove: (Ue) =>
              commitWorkbookChange(
                removeAutomationArtifact(latestEditStateRef.current.book, Ue, {
                  actor,
                  readOnly: latestEditStateRef.current.readOnly,
                }),
              ),
            onApply: (Ue, Xe, gt) => {
              const Nt = runApprovedMigration(
                latestEditStateRef.current.book,
                Ue,
                {
                  actor,
                  scope: Xe,
                  readOnly: latestEditStateRef.current.readOnly,
                  retainSource: gt,
                },
              );
              (commitWorkbookChange(Nt),
                operationStatusState.setStatusNotice(
                  "Reviewed migration applied in one Undo step; source hashes and run provenance recorded.",
                ));
            },
            onDownload: (Ue, Xe) =>
              saveBlobToFile(
                new Blob([Uint8Array.from(Ue)], {
                  type: "application/octet-stream",
                }),
                Xe,
              ),
          },
          dialogState.automationInspection?.sha256 ?? "automation",
        );
      if (
        dialogState.activeDialogMode === "search" ||
        dialogState.activeDialogMode === "goToSpecial"
      )
        return i.jsx(
          FindReplaceDialog,
          {
            ...fe,
            ranges: selectionAreas,
            active: selectionState.activeCellAddress,
            readOnly,
            initialTab:
              dialogState.activeDialogMode === "goToSpecial"
                ? "special"
                : "find",
            onWorkbookChange: replaceWorkbookSnapshot,
            onNavigate: navigateToSheetCell,
            onSelectRanges: selectSheetRanges,
          },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "cellStyle")
        return i.jsx(FormatCellsDialog, {
          ...fe,
          ranges: selectionAreas,
          active: selectionState.activeCellAddress,
        });
      if (dialogState.activeDialogMode === "themes")
        return i.jsx(ThemesAndCellStylesDialog, {
          ...fe,
          ranges: selectionAreas,
        });
      if (dialogState.activeDialogMode === "conditional")
        return i.jsx(ConditionalFormattingRulesManager, { ...fe });
      if (dialogState.activeDialogMode === "merge")
        return i.jsx(MergeCellsPreviewDialog, { ...fe });
      if (dialogState.activeDialogMode === "numberFormat")
        return i.jsx(M4e, {
          ...fe,
          ranges: selectionAreas,
          active: selectionState.activeCellAddress,
        });
      if (dialogState.activeDialogMode === "outlines")
        return i.jsx(OutlineToolsPanel, { ...fe });
      if (
        dialogState.activeDialogMode === "comments" ||
        dialogState.activeDialogMode === "hyperlink"
      )
        return i.jsx(ReviewToolsPanel, {
          ...fe,
          range: selectionState.activeCellAddress,
          mode: dialogState.activeDialogMode,
          active: selectionState.activeCellAddress,
          actor,
          onNavigate: navigateToSheetCell,
        });
      if (
        ["pasteSpecial", "fillSeries", "viewSettings"].includes(
          dialogState.activeDialogMode,
        )
      )
        return i.jsx(
          PasteSpecialDialog,
          {
            ...fe,
            mode: dialogState.activeDialogMode,
            clipboard: clipboardState.internalClipboardRef.current?.snapshot,
            onSelectionChange: selectSingleRange,
          },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "sheets")
        return i.jsx(WorksheetTabBar, {
          workbook: activeWorkbook,
          sheetId: navigationState.tabBarSheetId ?? activeSheet.id,
          readOnly,
          onWorkbookChange: replaceWorkbookSnapshot,
          onSelectSheet: (Ue) => {
            (dialogState.setActiveDialogMode(null), switchActiveWorksheet(Ue));
          },
          onClose: fe.onClose,
        });
      if (
        ["workbookProperties", "versions", "template"].includes(
          dialogState.activeDialogMode,
        )
      )
        return i.jsx(WorkbookManagerDialog, {
          mode:
            dialogState.activeDialogMode === "workbookProperties"
              ? "properties"
              : dialogState.activeDialogMode,
          workbook: activeWorkbook,
          readOnly,
          persistent: !!(
            persistenceKey ||
            (workbookStore && workbookStorePersistent)
          ),
          checkpoints: (
            persistenceState.persistedWorkbookRecord?.checkpoints ?? []
          ).filter((Ue) => Ue.workbookId === activeWorkbook.id),
          onClose: fe.onClose,
          onWorkbookChange: replaceWorkbookSnapshot,
          onCreateWorkbook: (Ue) => {
            if (readOnly) throw new Error("This workbook is read-only.");
            (saveBlobToFile(
              new Blob([encodeNativeWorkbook(Ue)], {
                type: "application/json",
              }),
              `${Ue.name.replace(/[^a-zA-Z0-9 _.-]/g, "_")}.json`,
            ),
              operationStatusState.setStatusNotice(
                "Created an independent workbook download. The template remains open.",
              ));
          },
          onCheckpoint: (Ue) => {
            if (readOnly) throw new Error("This workbook is read-only.");
            const Xe = persistenceState.persistenceSessionRef.current,
              gt = latestEditStateRef.current.book;
            return chainMaybePromise(persistWorkbookSnapshot(gt), () =>
              enqueueSaveOperation((Nt) => {
                if (latestEditStateRef.current.book !== gt)
                  throw new Error(
                    "The workbook changed. Save the checkpoint again for the current revision.",
                  );
                return chainMaybePromise(
                  workbookPersistence.checkpoint(
                    Ue,
                    persistenceState.persistedWorkbookRecordRef.current.guard,
                    { signal: Nt },
                  ),
                  (wt) => {
                    acknowledgeSavedWorkbook(wt, Xe);
                  },
                );
              }),
            );
          },
          onCompare: (Ue) =>
            workbookPersistence.compareCheckpoint(
              Ue,
              latestEditStateRef.current.book,
            ),
          onRestore: (Ue) => {
            if (readOnly) throw new Error("This workbook is read-only.");
            if (!persistenceState.persistedWorkbookRecordRef.current)
              throw new Error(
                "Save this workbook before restoring a checkpoint.",
              );
            const Xe = latestEditStateRef.current.book,
              gt = persistenceState.persistenceSessionRef.current;
            persistenceState.saveAbortControllerRef.current?.abort();
            const Nt = new AbortController();
            return (
              (persistenceState.saveAbortControllerRef.current = Nt),
              enqueueSaveOperation((jt) => {
                if (latestEditStateRef.current.book !== Xe || Nt.signal.aborted)
                  throw new Error(
                    "The workbook changed. Compare the checkpoint again before restoring.",
                  );
                const rr = () => Nt.abort();
                jt.addEventListener("abort", rr, { once: true });
                const Mr = (xr) => {
                    if ((jt.removeEventListener("abort", rr), gt.conflict))
                      throw gt.conflict;
                    if (
                      !gt.active ||
                      gt !== persistenceState.persistenceSessionRef.current ||
                      Nt.signal.aborted ||
                      latestEditStateRef.current.book !== Xe ||
                      formulaEntryState.activeCellEditRef.current ||
                      latestEditStateRef.current.readOnly
                    )
                      throw new Error(
                        "The editor changed before the restore completed. Your current local changes were retained.",
                      );
                    (commitWorkbookChange(xr.workbook, true, false, xr),
                      applyWorkbookView(xr.view),
                      operationStatusState.setStatusNotice(
                        "Restored checkpoint as a new revision.",
                      ),
                      reportSaveStatus("saved", xr.workbook));
                  },
                  Br = workbookPersistence.restoreCheckpoint(
                    Ue,
                    Xe,
                    persistenceState.persistedWorkbookRecordRef.current.guard,
                    actor,
                    {
                      signal: Nt.signal,
                    },
                  );
                return isPromiseLike(Br)
                  ? Br.then(Mr).finally(() =>
                      jt.removeEventListener("abort", rr),
                    )
                  : Mr(Br);
              })
            );
          },
        });
      if (["table", "pivot", "query"].includes(dialogState.activeDialogMode))
        return i.jsx(
          TableToolsDialog,
          { ...fe, mode: dialogState.activeDialogMode },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "pivotExplore")
        return i.jsx(PivotExplorerDialog, { ...fe });
      if (["picture", "shape", "chart"].includes(dialogState.activeDialogMode))
        return i.jsx(
          PicturePropertiesDialog,
          { ...fe },
          dialogState.activeDialogMode,
        );
      if (
        [
          "margins",
          "orientation",
          "size",
          "printArea",
          "printTitles",
          "pageBreaks",
          "fit",
          "pageLayout",
        ].includes(dialogState.activeDialogMode)
      )
        return i.jsx(PageSetupDialog, { ...fe }, dialogState.activeDialogMode);
      if (dialogState.activeDialogMode === "functionCompatibility")
        return i.jsx(FormulaCompatibilityDialog, {
          workbook: activeWorkbook,
          onClose: () => dialogState.setActiveDialogMode(null),
          onNavigate: navigateToSheetCell,
        });
      if (["function", "names"].includes(dialogState.activeDialogMode))
        return i.jsx(
          FunctionAndNameDialog,
          {
            ...fe,
            mode: dialogState.activeDialogMode,
            onInsertFormula: (Ue) => {
              (dialogState.setActiveDialogMode(null),
                beginCellEdit(Ue),
                formulaEntryState.setFormulaCaretPosition(Ue.length),
                formulaEntryState.setIsFormulaBarFocused(true),
                requestAnimationFrame(() => {
                  (formulaEntryState.formulaInputRef.current?.focus(),
                    formulaEntryState.formulaInputRef.current?.setSelectionRange(
                      Ue.length,
                      Ue.length,
                    ));
                }));
            },
          },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "populationFilter")
        return i.jsx(FilterViewsDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          column: filterSortState.activeFilterColumn,
          definition: activeFilterState?.definition,
          readOnly,
          onWorkbookChange: (Ue, Xe) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error(
                "The workbook changed. Reopen the filter tool before saving a view.",
              );
            replaceWorkbookSnapshot(Ue, Xe);
          },
          onClose: fe.onClose,
          onApply: (Ue) => {
            try {
              (applyFilterConditions(
                activeWorkbook,
                activeSheet.id,
                Ue,
                workbookCalculator,
              ),
                filterSortState.setFilterStateByScope((Xe) => ({
                  ...Xe,
                  [filterStateKey]: {
                    snapshot: activeWorkbook,
                    definition: Ue,
                  },
                })),
                filterSortState.setFilterValueSearch(""),
                filterSortState.setActiveColumnSort(null),
                auditState.setAuditPopulationScope("full"),
                scrollActivePaneVertically(0),
                dialogState.setActiveDialogMode(null),
                operationStatusState.setStatusNotice(
                  "Filter view applied. Source rows, formulas and history are unchanged.",
                ));
            } catch (Xe) {
              operationStatusState.setErrorNotice(
                Xe instanceof Error
                  ? Xe.message
                  : "Recalculate before filtering.",
              );
            }
          },
        });
      if (dialogState.activeDialogMode === "externalSources")
        return i.jsx(ExternalDataRefreshDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          adapter: externalSourceAdapter,
          reliability: activeReliabilityMonitor,
          readOnly,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the source again.");
            commitWorkbookWithNewSheet(Ue, Xe);
          },
        });
      if (
        dialogState.activeDialogMode === "solver" ||
        dialogState.activeDialogMode === "statistics"
      )
        return i.jsx(DataAnalysisDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          initialKind: dialogState.activeDialogMode,
          readOnly,
          allowHidden: true,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the analysis again.");
            commitWorkbookWithNewSheet(Ue, Xe);
          },
        });
      if (dialogState.activeDialogMode === "extensions")
        return i.jsx(AddInsAndCustomFunctionsDialog, {
          workbook: activeWorkbook,
          extensions,
          readOnly,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Review extensions again.",
              );
            replaceWorkbookSnapshot(Ue, Xe);
          },
        });
      if (dialogState.activeDialogMode === "workbookLinks")
        return i.jsx(WorkbookLinksDialog, {
          workbook: activeWorkbook,
          readOnly,
          approvedSources: effectiveLinkedWorkbooks,
          onApprovedSourcesChange: setLinkedWorkbooks,
          allowFileSources: allowLinkedWorkbookFiles,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Preview links again.",
              );
            replaceWorkbookSnapshot(Ue, Xe);
          },
        });
      if (dialogState.activeDialogMode === "importSources")
        return i.jsx(SourceImportDialog, {
          workbook: activeWorkbook,
          readOnly,
          extractor: documentExtractor,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Preview the source again.",
              );
            commitWorkbookWithNewSheet(Ue, Xe);
          },
        });
      if (dialogState.activeDialogMode === "workbookFormats")
        return i.jsx(FormatConversionDialog, {
          workbook: activeWorkbook,
          readOnly: isEditingLocked,
          initialFile: importExportState.pendingImportFile,
          onClose: () => {
            (importExportState.setPendingImportFile(undefined), fe.onClose());
          },
          onWorkbookChange: (Ue) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Review the conversion again.",
              );
            (commitWorkbookChange(Ue),
              switchActiveWorksheet(Ue.sheets.find((Xe) => !Xe.hidden).id));
          },
          onAutomation: async (Ue) => {
            const Xe = await importAutomationFile(Ue, operationScope);
            (dialogState.setAutomationInspection(Xe),
              importExportState.setPendingImportFile(undefined),
              dialogState.setActiveDialogMode("automation"));
          },
        });
      if (dialogState.activeDialogMode === "reportOutputs")
        return i.jsx(ReportOutputsDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          filter: filterResult?.context.definition,
          readOnly,
          allowHidden: true,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the report again.");
            replaceWorkbookSnapshot(Ue, Xe);
          },
        });
      if (dialogState.activeDialogMode === "dataModel")
        return i.jsx(DataModelDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          readOnly,
          allowHidden: true,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the model again.");
            commitWorkbookWithNewSheet(Ue, Xe);
          },
        });
      if (dialogState.activeDialogMode === "sort")
        return i.jsx(CustomSortDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          readOnly,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            (replaceWorkbookSnapshot(Ue, Xe),
              filterSortState.setActiveColumnSort(null),
              filterSortState.setFilterValueSearch(""));
          },
        });
      if (
        [
          "combinePopulations",
          "transformationRecipes",
          "consolidation",
          "textToColumns",
          "textCleanup",
          "fieldTransform",
          "duplicates",
        ].includes(dialogState.activeDialogMode)
      ) {
        const Ue = {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range:
            dialogState.activeDialogMode === "duplicates" &&
            !selectionState.activeSelectionRange.includes(":")
              ? formatOccupiedRangeAddress(activeSheet)
              : selectionState.activeSelectionRange,
          readOnly,
          allowHidden: true,
          onClose: fe.onClose,
          onWorkbookChange: commitWorkbookWithNewSheet,
        };
        return dialogState.activeDialogMode === "consolidation"
          ? i.jsx(ConsolidationDialog, {
              ...Ue,
              initialDefinition:
                activeSheet.consolidation?.definition ??
                consolidationDefinition,
              approvedWorkbooks: effectiveConsolidationWorkbooks,
              onApprovedWorkbooksChange: setConsolidationWorkbooks,
              allowFileSources:
                allowConsolidationFiles && !!setConsolidationWorkbooks,
            })
          : dialogState.activeDialogMode === "transformationRecipes"
            ? i.jsx(TransformationRecipesDialog, { ...Ue })
            : dialogState.activeDialogMode === "combinePopulations"
              ? i.jsx(CombinePopulationsDialog, { ...Ue })
              : dialogState.activeDialogMode === "textToColumns"
                ? i.jsx(TextToColumnsDialog, { ...Ue })
                : dialogState.activeDialogMode === "textCleanup"
                  ? i.jsx(TextCleanupDialog, { ...Ue })
                  : dialogState.activeDialogMode === "fieldTransform"
                    ? i.jsx(CleanAndReshapeDialog, { ...Ue })
                    : i.jsx(DuplicateReviewDialog, { ...Ue });
      }
      return dialogState.activeDialogMode === "patterns"
        ? i.jsx(PatternCheckDialog, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            range:
              auditState.auditSourceRange ||
              (selectionState.activeSelectionRange.includes(":")
                ? selectionState.activeSelectionRange
                : formatOccupiedRangeAddress(activeSheet)),
            readOnly,
            allowHidden: true,
            onClose: fe.onClose,
          })
        : dialogState.activeDialogMode === "aging"
          ? i.jsx(AgingScheduleScreen, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range:
                auditState.auditSourceRange ||
                (selectionState.activeSelectionRange.includes(":")
                  ? selectionState.activeSelectionRange
                  : formatOccupiedRangeAddress(activeSheet)),
              readOnly,
              allowHidden: true,
              onClose: fe.onClose,
            })
          : dialogState.activeDialogMode === "materiality"
            ? i.jsx(MaterialityRiskScreen, {
                workbook: activeWorkbook,
                sheetId: activeSheet.id,
                range:
                  auditState.auditSourceRange ||
                  (selectionState.activeSelectionRange.includes(":")
                    ? selectionState.activeSelectionRange
                    : formatOccupiedRangeAddress(activeSheet)),
                readOnly,
                allowHidden: true,
                onClose: fe.onClose,
              })
            : dialogState.activeDialogMode === "workbookComparison"
              ? i.jsx(WorkbookComparisonDialog, {
                  workbook: activeWorkbook,
                  sheetId: activeSheet.id,
                  range: selectionState.activeSelectionRange.includes(":")
                    ? selectionState.activeSelectionRange
                    : formatOccupiedRangeAddress(activeSheet),
                  readOnly,
                  allowHidden: true,
                  onClose: fe.onClose,
                })
              : dialogState.activeDialogMode === "scenarios" ||
                  dialogState.activeDialogMode === "dataTable"
                ? i.jsx(WhatIfAnalysisDialog, {
                    workbook: activeWorkbook,
                    sheetId: activeSheet.id,
                    address: selectionState.activeCellAddress,
                    initialKind:
                      dialogState.activeDialogMode === "scenarios"
                        ? "scenarios"
                        : "one-variable",
                    readOnly,
                    allowHidden: true,
                    onClose: fe.onClose,
                    onWorkbookChange: commitWorkbookWithNewSheet,
                  })
                : dialogState.activeDialogMode === "goalSeek"
                  ? i.jsx(GoalSeekDialog, {
                      workbook: activeWorkbook,
                      sheetId: activeSheet.id,
                      address: selectionState.activeCellAddress,
                      readOnly,
                      allowHidden: true,
                      onClose: fe.onClose,
                      onWorkbookChange: replaceWorkbookSnapshot,
                    })
                  : dialogState.activeDialogMode === "calculation"
                    ? i.jsx(CalculationOptionsDialog, {
                        workbook: activeWorkbook,
                        readOnly,
                        onClose: fe.onClose,
                        onWorkbookChange: replaceWorkbookSnapshot,
                        onCalculated: handleCalculationCompleted,
                        onCalculate: runCalculationNow,
                        onCancelCalculation: calculationRun.cancel,
                        onNavigate: navigateToSheetCell,
                      })
                    : dialogState.activeDialogMode === "formulaWatches"
                      ? i.jsx(FormulaWatchWindow, {
                          workbook: activeWorkbook,
                          sheetId: activeSheet.id,
                          range: selectionState.activeSelectionRange,
                          ranges: selectionAreas,
                          readOnly,
                          allowHidden: true,
                          onClose: fe.onClose,
                          onWorkbookChange: replaceWorkbookSnapshot,
                          onNavigate: navigateToSheetCell,
                        })
                      : dialogState.activeDialogMode === "evaluateFormula"
                        ? i.jsx(EvaluateFormulaDialog, {
                            workbook: activeWorkbook,
                            sheetId: activeSheet.id,
                            address: selectionState.activeCellAddress,
                            allowHidden: true,
                            onClose: fe.onClose,
                            onNavigate: navigateToSheetCell,
                          })
                        : dialogState.activeDialogMode === "tracePrecedents" ||
                            dialogState.activeDialogMode === "traceDependents"
                          ? i.jsx(CalculationInspectorDialog, {
                              workbook: activeWorkbook,
                              sheetId: activeSheet.id,
                              address: selectionState.activeCellAddress,
                              initialDirection:
                                dialogState.activeDialogMode ===
                                "tracePrecedents"
                                  ? "precedents"
                                  : "dependents",
                              allowHidden: true,
                              onClose: fe.onClose,
                              onNavigate: navigateToSheetCell,
                            })
                          : i.jsx(
                              DataToolDialogRouter,
                              { ...fe, onNavigate: navigateToSheetCell },
                              dialogState.activeDialogMode,
                            );
    })()
  );
}

export { EditorDialogs };
