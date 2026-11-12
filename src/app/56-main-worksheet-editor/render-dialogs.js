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
} from "../23-format-cells-dialog.js";
import { ThemesAndCellStylesDialog } from "../22-themes-and-cell-styles.js";
import { MergeCellsPreviewDialog } from "../24-merge-and-audit-analysis.js";
import { M4e } from "../48-format-cells-and-move.js";
import {
  OutlineToolsPanel,
  ReviewToolsPanel,
} from "../49-outlines-and-comments.js";
import { PasteSpecialDialog } from "../47-fill-and-paste-special.js";
import {
  CustomSortDialog,
  WorksheetTabBar,
} from "../45-sorting-and-worksheets.js";
import { WorkbookManagerDialog } from "../46-workbook-manager.js";
import { TableToolsDialog } from "../30-table-and-pivot-builder.js";
import { PivotExplorerDialog } from "../31-pivot-explore-and-slicers.js";
import { PicturePropertiesDialog } from "../32-drawing-and-chart-export.js";
import { PageSetupDialog } from "../33-page-setup-and-print.js";
import {
  CalculationInspectorDialog,
  EvaluateFormulaDialog,
  FormulaCompatibilityDialog,
} from "../35-formula-compatibility-report.js";
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
import { FormatConversionDialog } from "../11-workbook-formats-dialog.js";
import { ReportOutputsDialog } from "../39-report-outputs.js";
import { DataModelDialog } from "../15-data-model-dialog.js";
import { PatternCheckDialog } from "../51-pattern-checks-and-citations.js";
import {
  applyFilterConditions,
  formatOccupiedRangeAddress,
} from "../03-workbook-model.js";
import { AgingScheduleScreen } from "../21-aging-and-materiality-screens.js";
import { MaterialityRiskScreen } from "../50-materiality-screen.js";
import { WorkbookComparisonDialog } from "../52-workbook-comparison.js";
import { WhatIfAnalysisDialog } from "../40-what-if-analysis.js";
import {
  FormulaWatchWindow,
  GoalSeekDialog,
} from "../36-watch-window-and-goal-seek.js";
import { CalculationOptionsDialog } from "../20-calculation-options-and-sources.js";
import { DataToolDialogRouter } from "../29-conditional-formatting-dialog.js";
import {
  ConsolidationDialog,
  saveBlobToFile,
} from "../13-consolidation-dialog.js";
import { encodeNativeWorkbook } from "../08-xlsx-export-and-preservation.js";
import { importAutomationFile } from "../07-xlsx-archive-reader.js";
import { TransformationRecipesDialog } from "../18-transformation-recipes-dialog.js";
import {
  CombinePopulationsDialog,
  TextToColumnsDialog,
} from "../42-combine-populations.js";
import { TextCleanupDialog } from "../43-text-cleanup.js";
import { CleanAndReshapeDialog } from "../44-clean-and-reshape.js";
import { DuplicateReviewDialog } from "../28-duplicate-review-and-structure.js";

function renderEditorDialogs({
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
  persistedWorkbookRecordRef,
  persistedWorkbookRecord,
  persistenceSessionRef,
  saveAbortControllerRef,
  isEditingLocked,
  activeDialogMode,
  setActiveDialogMode,
  automationInspection,
  setAutomationInspection,
  tabBarSheetId,
  formulaInputRef,
  setFormulaCaretPosition,
  setIsFormulaBarFocused,
  activeCellAddress,
  activeSelectionRange,
  selectSingleRange,
  selectionAreas,
  activeCellEditRef,
  setFilterValueSearch,
  setFilterStateByScope,
  activeFilterColumn,
  setAuditPopulationScope,
  filterStateKey,
  activeFilterState,
  setActiveColumnSort,
  scrollActivePaneVertically,
  setStatusNotice,
  setErrorNotice,
  operationScope,
  pendingImportFile,
  setPendingImportFile,
  auditSourceRange,
  internalClipboardRef,
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
    activeDialogMode &&
    (() => {
      const fe = {
        mode: activeDialogMode,
        workbook: activeWorkbook,
        sheetId: activeSheet.id,
        range: activeSelectionRange,
        readOnly:
          activeDialogMode === "protection" ? readOnly : isEditingLocked,
        onClose: () => setActiveDialogMode(null),
        onWorkbookChange: commitWorkbookEdit,
      };
      if (activeDialogMode === "automation")
        return i.jsx(
          AutomationScriptsDialog,
          {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            range: activeSelectionRange,
            readOnly: isEditingLocked,
            initialInspection: automationInspection,
            onClose: fe.onClose,
            onPreserve: (Ue) => {
              const Xe = latestEditStateRef.current.book,
                gt = preserveAutomationArtifact(Xe, Ue, {
                  actor,
                  readOnly: latestEditStateRef.current.readOnly,
                });
              (gt !== Xe && commitWorkbookChange(gt),
                setStatusNotice(
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
                setStatusNotice(
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
          automationInspection?.sha256 ?? "automation",
        );
      if (activeDialogMode === "search" || activeDialogMode === "goToSpecial")
        return i.jsx(
          FindReplaceDialog,
          {
            ...fe,
            ranges: selectionAreas,
            active: activeCellAddress,
            readOnly,
            initialTab: activeDialogMode === "goToSpecial" ? "special" : "find",
            onWorkbookChange: replaceWorkbookSnapshot,
            onNavigate: navigateToSheetCell,
            onSelectRanges: selectSheetRanges,
          },
          activeDialogMode,
        );
      if (activeDialogMode === "cellStyle")
        return i.jsx(FormatCellsDialog, {
          ...fe,
          ranges: selectionAreas,
          active: activeCellAddress,
        });
      if (activeDialogMode === "themes")
        return i.jsx(ThemesAndCellStylesDialog, {
          ...fe,
          ranges: selectionAreas,
        });
      if (activeDialogMode === "conditional")
        return i.jsx(ConditionalFormattingRulesManager, { ...fe });
      if (activeDialogMode === "merge")
        return i.jsx(MergeCellsPreviewDialog, { ...fe });
      if (activeDialogMode === "numberFormat")
        return i.jsx(M4e, {
          ...fe,
          ranges: selectionAreas,
          active: activeCellAddress,
        });
      if (activeDialogMode === "outlines")
        return i.jsx(OutlineToolsPanel, { ...fe });
      if (activeDialogMode === "comments" || activeDialogMode === "hyperlink")
        return i.jsx(ReviewToolsPanel, {
          ...fe,
          range: activeCellAddress,
          mode: activeDialogMode,
          active: activeCellAddress,
          actor,
          onNavigate: navigateToSheetCell,
        });
      if (
        ["pasteSpecial", "fillSeries", "viewSettings"].includes(
          activeDialogMode,
        )
      )
        return i.jsx(
          PasteSpecialDialog,
          {
            ...fe,
            mode: activeDialogMode,
            clipboard: internalClipboardRef.current?.snapshot,
            onSelectionChange: selectSingleRange,
          },
          activeDialogMode,
        );
      if (activeDialogMode === "sheets")
        return i.jsx(WorksheetTabBar, {
          workbook: activeWorkbook,
          sheetId: tabBarSheetId ?? activeSheet.id,
          readOnly,
          onWorkbookChange: replaceWorkbookSnapshot,
          onSelectSheet: (Ue) => {
            (setActiveDialogMode(null), switchActiveWorksheet(Ue));
          },
          onClose: fe.onClose,
        });
      if (
        ["workbookProperties", "versions", "template"].includes(
          activeDialogMode,
        )
      )
        return i.jsx(WorkbookManagerDialog, {
          mode:
            activeDialogMode === "workbookProperties"
              ? "properties"
              : activeDialogMode,
          workbook: activeWorkbook,
          readOnly,
          persistent: !!(
            persistenceKey ||
            (workbookStore && workbookStorePersistent)
          ),
          checkpoints: (persistedWorkbookRecord?.checkpoints ?? []).filter(
            (Ue) => Ue.workbookId === activeWorkbook.id,
          ),
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
              setStatusNotice(
                "Created an independent workbook download. The template remains open.",
              ));
          },
          onCheckpoint: (Ue) => {
            if (readOnly) throw new Error("This workbook is read-only.");
            const Xe = persistenceSessionRef.current,
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
                    persistedWorkbookRecordRef.current.guard,
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
            if (!persistedWorkbookRecordRef.current)
              throw new Error(
                "Save this workbook before restoring a checkpoint.",
              );
            const Xe = latestEditStateRef.current.book,
              gt = persistenceSessionRef.current;
            saveAbortControllerRef.current?.abort();
            const Nt = new AbortController();
            return (
              (saveAbortControllerRef.current = Nt),
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
                      gt !== persistenceSessionRef.current ||
                      Nt.signal.aborted ||
                      latestEditStateRef.current.book !== Xe ||
                      activeCellEditRef.current ||
                      latestEditStateRef.current.readOnly
                    )
                      throw new Error(
                        "The editor changed before the restore completed. Your current local changes were retained.",
                      );
                    (commitWorkbookChange(xr.workbook, true, false, xr),
                      applyWorkbookView(xr.view),
                      setStatusNotice("Restored checkpoint as a new revision."),
                      reportSaveStatus("saved", xr.workbook));
                  },
                  Br = workbookPersistence.restoreCheckpoint(
                    Ue,
                    Xe,
                    persistedWorkbookRecordRef.current.guard,
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
      if (["table", "pivot", "query"].includes(activeDialogMode))
        return i.jsx(
          TableToolsDialog,
          { ...fe, mode: activeDialogMode },
          activeDialogMode,
        );
      if (activeDialogMode === "pivotExplore")
        return i.jsx(PivotExplorerDialog, { ...fe });
      if (["picture", "shape", "chart"].includes(activeDialogMode))
        return i.jsx(PicturePropertiesDialog, { ...fe }, activeDialogMode);
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
        ].includes(activeDialogMode)
      )
        return i.jsx(PageSetupDialog, { ...fe }, activeDialogMode);
      if (activeDialogMode === "functionCompatibility")
        return i.jsx(FormulaCompatibilityDialog, {
          workbook: activeWorkbook,
          onClose: () => setActiveDialogMode(null),
          onNavigate: navigateToSheetCell,
        });
      if (["function", "names"].includes(activeDialogMode))
        return i.jsx(
          FunctionAndNameDialog,
          {
            ...fe,
            mode: activeDialogMode,
            onInsertFormula: (Ue) => {
              (setActiveDialogMode(null),
                beginCellEdit(Ue),
                setFormulaCaretPosition(Ue.length),
                setIsFormulaBarFocused(true),
                requestAnimationFrame(() => {
                  (formulaInputRef.current?.focus(),
                    formulaInputRef.current?.setSelectionRange(
                      Ue.length,
                      Ue.length,
                    ));
                }));
            },
          },
          activeDialogMode,
        );
      if (activeDialogMode === "populationFilter")
        return i.jsx(FilterViewsDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: activeSelectionRange,
          column: activeFilterColumn,
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
                setFilterStateByScope((Xe) => ({
                  ...Xe,
                  [filterStateKey]: {
                    snapshot: activeWorkbook,
                    definition: Ue,
                  },
                })),
                setFilterValueSearch(""),
                setActiveColumnSort(null),
                setAuditPopulationScope("full"),
                scrollActivePaneVertically(0),
                setActiveDialogMode(null),
                setStatusNotice(
                  "Filter view applied. Source rows, formulas and history are unchanged.",
                ));
            } catch (Xe) {
              setErrorNotice(
                Xe instanceof Error
                  ? Xe.message
                  : "Recalculate before filtering.",
              );
            }
          },
        });
      if (activeDialogMode === "externalSources")
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
      if (activeDialogMode === "solver" || activeDialogMode === "statistics")
        return i.jsx(DataAnalysisDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: activeSelectionRange,
          initialKind: activeDialogMode,
          readOnly,
          allowHidden: true,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the analysis again.");
            commitWorkbookWithNewSheet(Ue, Xe);
          },
        });
      if (activeDialogMode === "extensions")
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
      if (activeDialogMode === "workbookLinks")
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
      if (activeDialogMode === "importSources")
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
      if (activeDialogMode === "workbookFormats")
        return i.jsx(FormatConversionDialog, {
          workbook: activeWorkbook,
          readOnly: isEditingLocked,
          initialFile: pendingImportFile,
          onClose: () => {
            (setPendingImportFile(undefined), fe.onClose());
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
            (setAutomationInspection(Xe),
              setPendingImportFile(undefined),
              setActiveDialogMode("automation"));
          },
        });
      if (activeDialogMode === "reportOutputs")
        return i.jsx(ReportOutputsDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: activeSelectionRange,
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
      if (activeDialogMode === "dataModel")
        return i.jsx(DataModelDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: activeSelectionRange,
          readOnly,
          allowHidden: true,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            if (latestEditStateRef.current.book !== activeWorkbook)
              throw Error("The workbook changed. Preview the model again.");
            commitWorkbookWithNewSheet(Ue, Xe);
          },
        });
      if (activeDialogMode === "sort")
        return i.jsx(CustomSortDialog, {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range: activeSelectionRange,
          readOnly,
          onClose: fe.onClose,
          onWorkbookChange: (Ue, Xe) => {
            (replaceWorkbookSnapshot(Ue, Xe),
              setActiveColumnSort(null),
              setFilterValueSearch(""));
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
        ].includes(activeDialogMode)
      ) {
        const Ue = {
          workbook: activeWorkbook,
          sheetId: activeSheet.id,
          range:
            activeDialogMode === "duplicates" &&
            !activeSelectionRange.includes(":")
              ? formatOccupiedRangeAddress(activeSheet)
              : activeSelectionRange,
          readOnly,
          allowHidden: true,
          onClose: fe.onClose,
          onWorkbookChange: commitWorkbookWithNewSheet,
        };
        return activeDialogMode === "consolidation"
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
          : activeDialogMode === "transformationRecipes"
            ? i.jsx(TransformationRecipesDialog, { ...Ue })
            : activeDialogMode === "combinePopulations"
              ? i.jsx(CombinePopulationsDialog, { ...Ue })
              : activeDialogMode === "textToColumns"
                ? i.jsx(TextToColumnsDialog, { ...Ue })
                : activeDialogMode === "textCleanup"
                  ? i.jsx(TextCleanupDialog, { ...Ue })
                  : activeDialogMode === "fieldTransform"
                    ? i.jsx(CleanAndReshapeDialog, { ...Ue })
                    : i.jsx(DuplicateReviewDialog, { ...Ue });
      }
      return activeDialogMode === "patterns"
        ? i.jsx(PatternCheckDialog, {
            workbook: activeWorkbook,
            sheetId: activeSheet.id,
            range:
              auditSourceRange ||
              (activeSelectionRange.includes(":")
                ? activeSelectionRange
                : formatOccupiedRangeAddress(activeSheet)),
            readOnly,
            allowHidden: true,
            onClose: fe.onClose,
          })
        : activeDialogMode === "aging"
          ? i.jsx(AgingScheduleScreen, {
              workbook: activeWorkbook,
              sheetId: activeSheet.id,
              range:
                auditSourceRange ||
                (activeSelectionRange.includes(":")
                  ? activeSelectionRange
                  : formatOccupiedRangeAddress(activeSheet)),
              readOnly,
              allowHidden: true,
              onClose: fe.onClose,
            })
          : activeDialogMode === "materiality"
            ? i.jsx(MaterialityRiskScreen, {
                workbook: activeWorkbook,
                sheetId: activeSheet.id,
                range:
                  auditSourceRange ||
                  (activeSelectionRange.includes(":")
                    ? activeSelectionRange
                    : formatOccupiedRangeAddress(activeSheet)),
                readOnly,
                allowHidden: true,
                onClose: fe.onClose,
              })
            : activeDialogMode === "workbookComparison"
              ? i.jsx(WorkbookComparisonDialog, {
                  workbook: activeWorkbook,
                  sheetId: activeSheet.id,
                  range: activeSelectionRange.includes(":")
                    ? activeSelectionRange
                    : formatOccupiedRangeAddress(activeSheet),
                  readOnly,
                  allowHidden: true,
                  onClose: fe.onClose,
                })
              : activeDialogMode === "scenarios" ||
                  activeDialogMode === "dataTable"
                ? i.jsx(WhatIfAnalysisDialog, {
                    workbook: activeWorkbook,
                    sheetId: activeSheet.id,
                    address: activeCellAddress,
                    initialKind:
                      activeDialogMode === "scenarios"
                        ? "scenarios"
                        : "one-variable",
                    readOnly,
                    allowHidden: true,
                    onClose: fe.onClose,
                    onWorkbookChange: commitWorkbookWithNewSheet,
                  })
                : activeDialogMode === "goalSeek"
                  ? i.jsx(GoalSeekDialog, {
                      workbook: activeWorkbook,
                      sheetId: activeSheet.id,
                      address: activeCellAddress,
                      readOnly,
                      allowHidden: true,
                      onClose: fe.onClose,
                      onWorkbookChange: replaceWorkbookSnapshot,
                    })
                  : activeDialogMode === "calculation"
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
                    : activeDialogMode === "formulaWatches"
                      ? i.jsx(FormulaWatchWindow, {
                          workbook: activeWorkbook,
                          sheetId: activeSheet.id,
                          range: activeSelectionRange,
                          ranges: selectionAreas,
                          readOnly,
                          allowHidden: true,
                          onClose: fe.onClose,
                          onWorkbookChange: replaceWorkbookSnapshot,
                          onNavigate: navigateToSheetCell,
                        })
                      : activeDialogMode === "evaluateFormula"
                        ? i.jsx(EvaluateFormulaDialog, {
                            workbook: activeWorkbook,
                            sheetId: activeSheet.id,
                            address: activeCellAddress,
                            allowHidden: true,
                            onClose: fe.onClose,
                            onNavigate: navigateToSheetCell,
                          })
                        : activeDialogMode === "tracePrecedents" ||
                            activeDialogMode === "traceDependents"
                          ? i.jsx(CalculationInspectorDialog, {
                              workbook: activeWorkbook,
                              sheetId: activeSheet.id,
                              address: activeCellAddress,
                              initialDirection:
                                activeDialogMode === "tracePrecedents"
                                  ? "precedents"
                                  : "dependents",
                              allowHidden: true,
                              onClose: fe.onClose,
                              onNavigate: navigateToSheetCell,
                            })
                          : i.jsx(
                              DataToolDialogRouter,
                              { ...fe, onNavigate: navigateToSheetCell },
                              activeDialogMode,
                            );
    })()
  );
}

export { renderEditorDialogs };
