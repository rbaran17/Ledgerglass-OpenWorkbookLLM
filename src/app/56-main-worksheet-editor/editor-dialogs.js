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
  viewAndSelectionActions,
  persistenceActions,
  navigationActions,
  cellEditActions,
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
  selectionAreas,
  filterStateKey,
  activeFilterState,
  operationScope,
  calculationRun,
  workbookCalculator,
  filterResult,
  latestEditStateRef,
  handleCalculationCompleted,
}) {
  return (
    dialogState.activeDialogMode &&
    (() => {
      const dialogProps = {
        mode: dialogState.activeDialogMode,
        workbook: activeWorkbook,
        sheetId: activeSheet.id,
        range: selectionState.activeSelectionRange,
        readOnly:
          dialogState.activeDialogMode === "protection"
            ? readOnly
            : isEditingLocked,
        onClose: () => dialogState.setActiveDialogMode(null),
        onWorkbookChange: navigationActions.commitWorkbookEdit,
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
            onClose: dialogProps.onClose,
            onPreserve: (inspection) => {
              const book = latestEditStateRef.current.book,
                preservedBook = preserveAutomationArtifact(book, inspection, {
                  actor,
                  readOnly: latestEditStateRef.current.readOnly,
                });
              (preservedBook !== book &&
                persistenceActions.commitWorkbookChange(preservedBook),
                operationStatusState.setStatusNotice(
                  "Original automation bytes preserved in native JSON. No source code was executed.",
                ));
            },
            onRemove: (artifactId) =>
              persistenceActions.commitWorkbookChange(
                removeAutomationArtifact(
                  latestEditStateRef.current.book,
                  artifactId,
                  {
                    actor,
                    readOnly: latestEditStateRef.current.readOnly,
                  },
                ),
              ),
            onApply: (migration, scope, retainSource) => {
              const migratedBook = runApprovedMigration(
                latestEditStateRef.current.book,
                migration,
                {
                  actor,
                  scope: scope,
                  readOnly: latestEditStateRef.current.readOnly,
                  retainSource: retainSource,
                },
              );
              (persistenceActions.commitWorkbookChange(migratedBook),
                operationStatusState.setStatusNotice(
                  "Reviewed migration applied in one Undo step; source hashes and run provenance recorded.",
                ));
            },
            onDownload: (artifactBytes, fileName) =>
              saveBlobToFile(
                new Blob([Uint8Array.from(artifactBytes)], {
                  type: "application/octet-stream",
                }),
                fileName,
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
            ...dialogProps,
            ranges: selectionAreas,
            active: selectionState.activeCellAddress,
            readOnly,
            initialTab:
              dialogState.activeDialogMode === "goToSpecial"
                ? "special"
                : "find",
            onWorkbookChange: navigationActions.replaceWorkbookSnapshot,
            onNavigate: navigationActions.navigateToSheetCell,
            onSelectRanges: navigationActions.selectSheetRanges,
          },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "cellStyle")
        return i.jsx(FormatCellsDialog, {
          ...dialogProps,
          ranges: selectionAreas,
          active: selectionState.activeCellAddress,
        });
      if (dialogState.activeDialogMode === "themes")
        return i.jsx(ThemesAndCellStylesDialog, {
          ...dialogProps,
          ranges: selectionAreas,
        });
      if (dialogState.activeDialogMode === "conditional")
        return i.jsx(ConditionalFormattingRulesManager, { ...dialogProps });
      if (dialogState.activeDialogMode === "merge")
        return i.jsx(MergeCellsPreviewDialog, { ...dialogProps });
      if (dialogState.activeDialogMode === "numberFormat")
        return i.jsx(M4e, {
          ...dialogProps,
          ranges: selectionAreas,
          active: selectionState.activeCellAddress,
        });
      if (dialogState.activeDialogMode === "outlines")
        return i.jsx(OutlineToolsPanel, { ...dialogProps });
      if (
        dialogState.activeDialogMode === "comments" ||
        dialogState.activeDialogMode === "hyperlink"
      )
        return i.jsx(ReviewToolsPanel, {
          ...dialogProps,
          range: selectionState.activeCellAddress,
          mode: dialogState.activeDialogMode,
          active: selectionState.activeCellAddress,
          actor,
          onNavigate: navigationActions.navigateToSheetCell,
        });
      if (
        ["pasteSpecial", "fillSeries", "viewSettings"].includes(
          dialogState.activeDialogMode,
        )
      )
        return i.jsx(
          PasteSpecialDialog,
          {
            ...dialogProps,
            mode: dialogState.activeDialogMode,
            clipboard: clipboardState.internalClipboardRef.current?.snapshot,
            onSelectionChange: viewAndSelectionActions.selectSingleRange,
          },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "sheets")
        return i.jsx(WorksheetTabBar, {
          workbook: activeWorkbook,
          sheetId: navigationState.tabBarSheetId ?? activeSheet.id,
          readOnly,
          onWorkbookChange: navigationActions.replaceWorkbookSnapshot,
          onSelectSheet: (sheetId) => {
            (dialogState.setActiveDialogMode(null),
              navigationActions.switchActiveWorksheet(sheetId));
          },
          onClose: dialogProps.onClose,
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
          ).filter((checkpoint) => checkpoint.workbookId === activeWorkbook.id),
          onClose: dialogProps.onClose,
          onWorkbookChange: navigationActions.replaceWorkbookSnapshot,
          onCreateWorkbook: (workbook) => {
            if (readOnly) throw new Error("This workbook is read-only.");
            (saveBlobToFile(
              new Blob([encodeNativeWorkbook(workbook)], {
                type: "application/json",
              }),
              `${workbook.name.replace(/[^a-zA-Z0-9 _.-]/g, "_")}.json`,
            ),
              operationStatusState.setStatusNotice(
                "Created an independent workbook download. The template remains open.",
              ));
          },
          onCheckpoint: (Ue) => {
            if (readOnly) throw new Error("This workbook is read-only.");
            const persistenceSession =
                persistenceState.persistenceSessionRef.current,
              book = latestEditStateRef.current.book;
            return chainMaybePromise(
              persistenceActions.persistWorkbookSnapshot(book),
              () =>
                persistenceActions.enqueueSaveOperation((signal) => {
                  if (latestEditStateRef.current.book !== book)
                    throw new Error(
                      "The workbook changed. Save the checkpoint again for the current revision.",
                    );
                  return chainMaybePromise(
                    workbookPersistence.checkpoint(
                      Ue,
                      persistenceState.persistedWorkbookRecordRef.current.guard,
                      { signal: signal },
                    ),
                    (savedRecord) => {
                      persistenceActions.acknowledgeSavedWorkbook(
                        savedRecord,
                        persistenceSession,
                      );
                    },
                  );
                }),
            );
          },
          onCompare: (targetCheckpoint) =>
            workbookPersistence.compareCheckpoint(
              targetCheckpoint,
              latestEditStateRef.current.book,
            ),
          onRestore: (checkpointToRestore) => {
            if (readOnly) throw new Error("This workbook is read-only.");
            if (!persistenceState.persistedWorkbookRecordRef.current)
              throw new Error(
                "Save this workbook before restoring a checkpoint.",
              );
            const book = latestEditStateRef.current.book,
              restoreSession = persistenceState.persistenceSessionRef.current;
            persistenceState.saveAbortControllerRef.current?.abort();
            const restoreAbortController = new AbortController();
            return (
              (persistenceState.saveAbortControllerRef.current =
                restoreAbortController),
              persistenceActions.enqueueSaveOperation((restoreSignal) => {
                if (
                  latestEditStateRef.current.book !== book ||
                  restoreAbortController.signal.aborted
                )
                  throw new Error(
                    "The workbook changed. Compare the checkpoint again before restoring.",
                  );
                const abortRestore = () => restoreAbortController.abort();
                restoreSignal.addEventListener("abort", abortRestore, {
                  once: true,
                });
                const applyRestoreResult = (restoreResult) => {
                    if (
                      (restoreSignal.removeEventListener("abort", abortRestore),
                      restoreSession.conflict)
                    )
                      throw restoreSession.conflict;
                    if (
                      !restoreSession.active ||
                      restoreSession !==
                        persistenceState.persistenceSessionRef.current ||
                      restoreAbortController.signal.aborted ||
                      latestEditStateRef.current.book !== book ||
                      formulaEntryState.activeCellEditRef.current ||
                      latestEditStateRef.current.readOnly
                    )
                      throw new Error(
                        "The editor changed before the restore completed. Your current local changes were retained.",
                      );
                    (persistenceActions.commitWorkbookChange(
                      restoreResult.workbook,
                      true,
                      false,
                      restoreResult,
                    ),
                      viewAndSelectionActions.applyWorkbookView(
                        restoreResult.view,
                      ),
                      operationStatusState.setStatusNotice(
                        "Restored checkpoint as a new revision.",
                      ),
                      persistenceActions.reportSaveStatus(
                        "saved",
                        restoreResult.workbook,
                      ));
                  },
                  restoreOutcome = workbookPersistence.restoreCheckpoint(
                    checkpointToRestore,
                    book,
                    persistenceState.persistedWorkbookRecordRef.current.guard,
                    actor,
                    {
                      signal: restoreAbortController.signal,
                    },
                  );
                return isPromiseLike(restoreOutcome)
                  ? restoreOutcome
                      .then(applyRestoreResult)
                      .finally(() =>
                        restoreSignal.removeEventListener(
                          "abort",
                          abortRestore,
                        ),
                      )
                  : applyRestoreResult(restoreOutcome);
              })
            );
          },
        });
      if (["table", "pivot", "query"].includes(dialogState.activeDialogMode))
        return i.jsx(
          TableToolsDialog,
          { ...dialogProps, mode: dialogState.activeDialogMode },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "pivotExplore")
        return i.jsx(PivotExplorerDialog, { ...dialogProps });
      if (["picture", "shape", "chart"].includes(dialogState.activeDialogMode))
        return i.jsx(
          PicturePropertiesDialog,
          { ...dialogProps },
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
        return i.jsx(
          PageSetupDialog,
          { ...dialogProps },
          dialogState.activeDialogMode,
        );
      if (dialogState.activeDialogMode === "functionCompatibility")
        return i.jsx(FormulaCompatibilityDialog, {
          workbook: activeWorkbook,
          onClose: () => dialogState.setActiveDialogMode(null),
          onNavigate: navigationActions.navigateToSheetCell,
        });
      if (["function", "names"].includes(dialogState.activeDialogMode))
        return i.jsx(
          FunctionAndNameDialog,
          {
            ...dialogProps,
            mode: dialogState.activeDialogMode,
            onInsertFormula: (formulaText) => {
              (dialogState.setActiveDialogMode(null),
                cellEditActions.beginCellEdit(formulaText),
                formulaEntryState.setFormulaCaretPosition(formulaText.length),
                formulaEntryState.setIsFormulaBarFocused(true),
                requestAnimationFrame(() => {
                  (formulaEntryState.formulaInputRef.current?.focus(),
                    formulaEntryState.formulaInputRef.current?.setSelectionRange(
                      formulaText.length,
                      formulaText.length,
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
          onWorkbookChange: (filterViewWorkbook, filterViewNotice) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error(
                "The workbook changed. Reopen the filter tool before saving a view.",
              );
            navigationActions.replaceWorkbookSnapshot(
              filterViewWorkbook,
              filterViewNotice,
            );
          },
          onClose: dialogProps.onClose,
          onApply: (definition) => {
            try {
              (applyFilterConditions(
                activeWorkbook,
                activeSheet.id,
                definition,
                workbookCalculator,
              ),
                filterSortState.setFilterStateByScope(
                  (previousFilterStates) => ({
                    ...previousFilterStates,
                    [filterStateKey]: {
                      snapshot: activeWorkbook,
                      definition: definition,
                    },
                  }),
                ),
                filterSortState.setFilterValueSearch(""),
                filterSortState.setActiveColumnSort(null),
                auditState.setAuditPopulationScope("full"),
                viewAndSelectionActions.scrollActivePaneVertically(0),
                dialogState.setActiveDialogMode(null),
                operationStatusState.setStatusNotice(
                  "Filter view applied. Source rows, formulas and history are unchanged.",
                ));
            } catch (error) {
              operationStatusState.setErrorNotice(
                error instanceof Error
                  ? error.message
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
          onClose: dialogProps.onClose,
          onWorkbookChange: (refreshedWorkbook, refreshNotice) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the source again.");
            navigationActions.commitWorkbookWithNewSheet(
              refreshedWorkbook,
              refreshNotice,
            );
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
          onClose: dialogProps.onClose,
          onWorkbookChange: (analysisWorkbook, analysisNotice) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the analysis again.");
            navigationActions.commitWorkbookWithNewSheet(
              analysisWorkbook,
              analysisNotice,
            );
          },
        });
      if (dialogState.activeDialogMode === "extensions")
        return i.jsx(AddInsAndCustomFunctionsDialog, {
          workbook: activeWorkbook,
          extensions,
          readOnly,
          onClose: dialogProps.onClose,
          onWorkbookChange: (extendedWorkbook, extensionNotice) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Review extensions again.",
              );
            navigationActions.replaceWorkbookSnapshot(
              extendedWorkbook,
              extensionNotice,
            );
          },
        });
      if (dialogState.activeDialogMode === "workbookLinks")
        return i.jsx(WorkbookLinksDialog, {
          workbook: activeWorkbook,
          readOnly,
          approvedSources: effectiveLinkedWorkbooks,
          onApprovedSourcesChange: setLinkedWorkbooks,
          allowFileSources: allowLinkedWorkbookFiles,
          onClose: dialogProps.onClose,
          onWorkbookChange: (linkedWorkbook, linkNotice) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Preview links again.",
              );
            navigationActions.replaceWorkbookSnapshot(
              linkedWorkbook,
              linkNotice,
            );
          },
        });
      if (dialogState.activeDialogMode === "importSources")
        return i.jsx(SourceImportDialog, {
          workbook: activeWorkbook,
          readOnly,
          extractor: documentExtractor,
          onClose: dialogProps.onClose,
          onWorkbookChange: (importedWorkbook, importNotice) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Preview the source again.",
              );
            navigationActions.commitWorkbookWithNewSheet(
              importedWorkbook,
              importNotice,
            );
          },
        });
      if (dialogState.activeDialogMode === "workbookFormats")
        return i.jsx(FormatConversionDialog, {
          workbook: activeWorkbook,
          readOnly: isEditingLocked,
          initialFile: importExportState.pendingImportFile,
          onClose: () => {
            (importExportState.setPendingImportFile(undefined),
              dialogProps.onClose());
          },
          onWorkbookChange: (workbook) => {
            if (
              latestEditStateRef.current.book !== activeWorkbook ||
              latestEditStateRef.current.readOnly
            )
              throw Error(
                "The workbook or permissions changed. Review the conversion again.",
              );
            (persistenceActions.commitWorkbookChange(workbook),
              navigationActions.switchActiveWorksheet(
                workbook.sheets.find((sheet) => !sheet.hidden).id,
              ));
          },
          onAutomation: async (automationFile) => {
            const automationInspection = await importAutomationFile(
              automationFile,
              operationScope,
            );
            (dialogState.setAutomationInspection(automationInspection),
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
          onClose: dialogProps.onClose,
          onWorkbookChange: (reportWorkbook, reportNotice) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the report again.");
            navigationActions.replaceWorkbookSnapshot(
              reportWorkbook,
              reportNotice,
            );
          },
        });
      if (dialogState.activeDialogMode === "dataModel")
        return i.jsx(DataModelDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          readOnly,
          allowHidden: true,
          onClose: dialogProps.onClose,
          onWorkbookChange: (modelWorkbook, modelNotice) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the model again.");
            navigationActions.commitWorkbookWithNewSheet(
              modelWorkbook,
              modelNotice,
            );
          },
        });
      if (dialogState.activeDialogMode === "sort")
        return i.jsx(CustomSortDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: selectionState.activeSelectionRange,
          readOnly,
          onClose: dialogProps.onClose,
          onWorkbookChange: (sortedWorkbook, sortNotice) => {
            (navigationActions.replaceWorkbookSnapshot(
              sortedWorkbook,
              sortNotice,
            ),
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
        const sharedDialogProps = {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range:
            dialogState.activeDialogMode === "duplicates" &&
            !selectionState.activeSelectionRange.includes(":")
              ? formatOccupiedRangeAddress(activeSheet)
              : selectionState.activeSelectionRange,
          readOnly,
          allowHidden: true,
          onClose: dialogProps.onClose,
          onWorkbookChange: navigationActions.commitWorkbookWithNewSheet,
        };
        return dialogState.activeDialogMode === "consolidation"
          ? i.jsx(ConsolidationDialog, {
              ...sharedDialogProps,
              initialDefinition:
                activeSheet.consolidation?.definition ??
                consolidationDefinition,
              approvedWorkbooks: effectiveConsolidationWorkbooks,
              onApprovedWorkbooksChange: setConsolidationWorkbooks,
              allowFileSources:
                allowConsolidationFiles && !!setConsolidationWorkbooks,
            })
          : dialogState.activeDialogMode === "transformationRecipes"
            ? i.jsx(TransformationRecipesDialog, { ...sharedDialogProps })
            : dialogState.activeDialogMode === "combinePopulations"
              ? i.jsx(CombinePopulationsDialog, { ...sharedDialogProps })
              : dialogState.activeDialogMode === "textToColumns"
                ? i.jsx(TextToColumnsDialog, { ...sharedDialogProps })
                : dialogState.activeDialogMode === "textCleanup"
                  ? i.jsx(TextCleanupDialog, { ...sharedDialogProps })
                  : dialogState.activeDialogMode === "fieldTransform"
                    ? i.jsx(CleanAndReshapeDialog, { ...sharedDialogProps })
                    : i.jsx(DuplicateReviewDialog, { ...sharedDialogProps });
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
            onClose: dialogProps.onClose,
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
              onClose: dialogProps.onClose,
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
                onClose: dialogProps.onClose,
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
                  onClose: dialogProps.onClose,
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
                    onClose: dialogProps.onClose,
                    onWorkbookChange:
                      navigationActions.commitWorkbookWithNewSheet,
                  })
                : dialogState.activeDialogMode === "goalSeek"
                  ? i.jsx(GoalSeekDialog, {
                      workbook: activeWorkbook,
                      sheetId: activeSheet.id,
                      address: selectionState.activeCellAddress,
                      readOnly,
                      allowHidden: true,
                      onClose: dialogProps.onClose,
                      onWorkbookChange:
                        navigationActions.replaceWorkbookSnapshot,
                    })
                  : dialogState.activeDialogMode === "calculation"
                    ? i.jsx(CalculationOptionsDialog, {
                        workbook: activeWorkbook,
                        readOnly,
                        onClose: dialogProps.onClose,
                        onWorkbookChange:
                          navigationActions.replaceWorkbookSnapshot,
                        onCalculated: handleCalculationCompleted,
                        onCalculate: persistenceActions.runCalculationNow,
                        onCancelCalculation: calculationRun.cancel,
                        onNavigate: navigationActions.navigateToSheetCell,
                      })
                    : dialogState.activeDialogMode === "formulaWatches"
                      ? i.jsx(FormulaWatchWindow, {
                          workbook: activeWorkbook,
                          sheetId: activeSheet.id,
                          range: selectionState.activeSelectionRange,
                          ranges: selectionAreas,
                          readOnly,
                          allowHidden: true,
                          onClose: dialogProps.onClose,
                          onWorkbookChange:
                            navigationActions.replaceWorkbookSnapshot,
                          onNavigate: navigationActions.navigateToSheetCell,
                        })
                      : dialogState.activeDialogMode === "evaluateFormula"
                        ? i.jsx(EvaluateFormulaDialog, {
                            workbook: activeWorkbook,
                            sheetId: activeSheet.id,
                            address: selectionState.activeCellAddress,
                            allowHidden: true,
                            onClose: dialogProps.onClose,
                            onNavigate: navigationActions.navigateToSheetCell,
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
                              onClose: dialogProps.onClose,
                              onNavigate: navigationActions.navigateToSheetCell,
                            })
                          : i.jsx(
                              DataToolDialogRouter,
                              {
                                ...dialogProps,
                                onNavigate:
                                  navigationActions.navigateToSheetCell,
                              },
                              dialogState.activeDialogMode,
                            );
    })()
  );
}

export { EditorDialogs };
