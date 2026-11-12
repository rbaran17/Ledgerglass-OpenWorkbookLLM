import {
  FileProcessingError,
  importAutomationFile,
  importDelimitedFile,
  importWorkbookFile,
  toFileProcessingError,
} from "../07-file-import-and-worker-dispatch.js";
import { workbookLimits } from "../02-text-and-financial-functions.js";
import { saveBlobToFile } from "../13-consolidation-dialog.js";
import {
  encodeNativeWorkbook,
  exportWorkbookToXlsx,
} from "../08-xlsx-import-export-and-preservation.js";
import { runAuditAnalysis } from "../24-merge-audit-and-ribbon-helpers.js";
import { formatOccupiedRangeAddress } from "../03-workbook-model.js";
import {
  buildAssistantSelectionContext,
  validateProposedCalculation,
} from "../20-calculation-options-and-assistant.js";
import { requestAssistantAnswer } from "../21-aging-and-materiality-screens.js";
import { buildAnswerCitations } from "../51-pattern-checks-and-citations.js";

function createFileAndAnalysisActions({
  dialogState,
  auditState,
  operationStatusState,
  importExportState,
  assistantState,
  adapter,
  actor,
  activeReliabilityMonitor,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  selectionAreas,
  activeFilterState,
  operationScope,
  filterResult,
  latestEditStateRef,
  assistantSelection,
  commitWorkbookChange,
  runGuardedAction,
  switchActiveWorksheet,
}) {
  async function importWorkbookFromFile(fe) {
    if (!fe) return;
    operationStatusState.activeOperationControllerRef.current?.abort();
    const Oe = new AbortController(),
      Be = ++operationStatusState.latestOperationIdRef.current;
    operationStatusState.activeOperationControllerRef.current = Oe;
    const Ue = {
        ...operationScope,
        signal: operationScope.signal
          ? AbortSignal.any([Oe.signal, operationScope.signal])
          : Oe.signal,
      },
      Xe = () => Be === operationStatusState.latestOperationIdRef.current;
    (operationStatusState.setIsOperationBusy(true),
      operationStatusState.setErrorNotice(""));
    try {
      if (isEditingLocked) throw new Error("This workbook is read-only.");
      if (/\.(xls|xlsb|xltx|ods)$/i.test(fe.name)) {
        (importExportState.setPendingImportFile(fe),
          dialogState.setActiveDialogMode("workbookFormats"));
        return;
      }
      if (/\.(bas|vba|cls|ts|osts|xlsm|xlam|xltm)$/i.test(fe.name)) {
        if (fe.size > workbookLimits.maxFileBytes)
          throw new Error("Files must be 10 MB or smaller.");
        const wt = await importAutomationFile(fe, Ue);
        if (!Xe()) return;
        if (
          latestEditStateRef.current.book !== activeWorkbook ||
          latestEditStateRef.current.readOnly
        )
          throw new Error(
            "The workbook changed while the file was being read. Open the file again.",
          );
        (dialogState.setAutomationInspection(wt),
          dialogState.setActiveDialogMode("automation"));
        return;
      }
      if (/\.(csv|tsv)$/i.test(fe.name)) {
        if (fe.size > workbookLimits.maxFileBytes)
          throw new Error("Files must be 10 MB or smaller.");
        const wt = await importDelimitedFile(fe, {}, Ue);
        if (!Xe()) return;
        if (
          latestEditStateRef.current.book !== activeWorkbook ||
          latestEditStateRef.current.readOnly
        )
          throw new Error(
            "The workbook changed while the file was being read. Open the file again.",
          );
        importExportState.setDelimitedImportExportSource({
          snapshot: activeWorkbook,
          bytes: new Uint8Array(),
          file: fe,
          initial: wt,
          fileName: fe.name,
        });
        return;
      }
      const Nt = await importWorkbookFile(fe, Ue);
      if (!Xe()) return;
      if (
        latestEditStateRef.current.book !== activeWorkbook ||
        latestEditStateRef.current.readOnly
      )
        throw new Error(
          "The workbook changed while the file was being read. Import again to replace the current revision.",
        );
      if (Nt.report) {
        importExportState.setXlsxConversionReview({
          snapshot: activeWorkbook,
          report: Nt.report,
          imported: Nt.workbook,
          fileName: fe.name,
        });
        return;
      }
      (commitWorkbookChange(Nt.workbook),
        switchActiveWorksheet(Nt.workbook.sheets.find((wt) => !wt.hidden).id),
        operationStatusState.setStatusNotice(
          Nt.warnings.length ? Nt.warnings.join(" ") : "Workbook imported.",
        ));
    } catch (Nt) {
      if (!Xe()) return;
      if (
        !Oe.signal.aborted &&
        (!(Nt instanceof FileProcessingError) ||
          ["FILE_ACTIVE", "FILE_INVALID"].includes(Nt.code)) &&
        /\.xlsx$/i.test(fe.name) &&
        fe.size <= workbookLimits.maxFileBytes &&
        latestEditStateRef.current.book === activeWorkbook
      )
        try {
          const wt = await importAutomationFile(fe, Ue);
          if (
            Xe() &&
            !latestEditStateRef.current.readOnly &&
            wt.items.length &&
            latestEditStateRef.current.book === activeWorkbook
          ) {
            (dialogState.setAutomationInspection(wt),
              dialogState.setActiveDialogMode("automation"));
            return;
          }
        } catch {}
      Xe() &&
        operationStatusState.setErrorNotice(toFileProcessingError(Nt).message);
    } finally {
      Xe() &&
        ((operationStatusState.activeOperationControllerRef.current = null),
        operationStatusState.setIsOperationBusy(false),
        importExportState.importFileInputRef.current &&
          (importExportState.importFileInputRef.current.value = ""));
    }
  }
  async function exportWorkbookAs(fe) {
    (operationStatusState.setIsOperationBusy(true),
      operationStatusState.setErrorNotice(""));
    try {
      const Oe =
        activeWorkbook.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) ||
        "Workbook";
      if (fe === "xlsx") {
        const Be = await exportWorkbookToXlsx(activeWorkbook);
        (Be.report.warnings.push(
          "Full workbook export includes all filtered and hidden rows. Session filter views are not saved in XLSX. Use CSV/TSV with Visible population for a filtered extract.",
        ),
          importExportState.setXlsxConversionReview({
            snapshot: activeWorkbook,
            report: { ...Be.report, fileName: `${Oe}-derived.xlsx` },
            blob: Be.blob,
            fileName: `${Oe}-derived.xlsx`,
          }));
      }
      (fe === "csv" &&
        importExportState.setDelimitedImportExportSource({
          snapshot: activeWorkbook,
          sheetId: activeSheet.id,
          ...(filterResult
            ? { population: filterResult.context.definition }
            : {}),
        }),
        fe === "json" &&
          (operationStatusState.setStatusNotice(
            "Native JSON exports the full workbook, including all filtered and hidden rows and any named filter views. Active session masks are not saved; load named criteria, preview and apply after reopening.",
          ),
          saveBlobToFile(
            new Blob([encodeNativeWorkbook(activeWorkbook)], {
              type: "application/json",
            }),
            `${Oe}.json`,
          )));
    } catch (Oe) {
      operationStatusState.setErrorNotice(
        Oe instanceof Error ? Oe.message : "Export failed.",
      );
    } finally {
      operationStatusState.setIsOperationBusy(false);
    }
  }
  function runCurrentAuditAnalysis() {
    runGuardedAction(() => {
      if (auditState.auditPopulationScope === "visible" && !filterResult)
        throw Error(
          "Review and apply a current population filter before visible-only analysis.",
        );
      const fe = runAuditAnalysis(
        activeWorkbook,
        activeSheet.id,
        auditState.selectedAuditKind,
        {
          range:
            auditState.auditPopulationScope === "visible" && filterResult
              ? filterResult.context.range
              : auditState.auditSourceRange ||
                formatOccupiedRangeAddress(activeSheet),
          ...(auditState.auditPopulationScope === "visible"
            ? {
                population: {
                  definition:
                    filterResult?.context.definition ??
                    activeFilterState?.definition,
                  scope: "visible",
                },
              }
            : {}),
          threshold: auditState.materialityThreshold,
          sampleSize: auditState.auditSampleSize,
          seed: auditState.samplingSeed,
          asOf: auditState.agingAsOfDate,
        },
      );
      (auditState.setAuditAnalysisResult(fe),
        auditState.setAgingReportPayload({
          schemaVersion: 1,
          workbookId: activeWorkbook.id,
          workbookName: activeWorkbook.name,
          revision: activeWorkbook.revision,
          sheetId: activeSheet.id,
          sheetName: activeSheet.name,
          generatedAt: new Date().toISOString(),
          actor,
          populationScope: auditState.auditPopulationScope,
          ...(filterResult ? { filterContext: filterResult.context } : {}),
          kind: auditState.selectedAuditKind,
          parameters: {
            range:
              auditState.auditPopulationScope === "visible" && filterResult
                ? filterResult.context.range
                : auditState.auditSourceRange ||
                  formatOccupiedRangeAddress(activeSheet),
            threshold: auditState.materialityThreshold,
            sampleSize: auditState.auditSampleSize,
            seed: auditState.samplingSeed,
            asOf: auditState.agingAsOfDate,
          },
          result: fe,
        }),
        operationStatusState.setStatusNotice(
          auditState.auditPopulationScope === "visible"
            ? "Analysis completed on visible population rows; citations retain original worksheet addresses."
            : "Analysis completed on the full selected source range, including hidden and filtered rows.",
        ));
    });
  }
  async function askAssistantQuestion(fe, Oe) {
    if (!fe.trim()) return null;
    if (selectionAreas.length > 1)
      throw (
        operationStatusState.setErrorNotice(
          "Select one rectangular range before asking the assistant.",
        ),
        new Error("Select one rectangular range before asking the assistant.")
      );
    assistantState.assistantAbortControllerRef.current?.abort();
    const Be = ++assistantState.assistantRequestIdRef.current;
    (assistantState.setIsAssistantBusy(true),
      assistantState.setAssistantProposal(null),
      assistantState.setAnswerCitations(null),
      operationStatusState.setErrorNotice(""),
      operationStatusState.setStatusNotice(""));
    const Ue = new AbortController();
    assistantState.assistantAbortControllerRef.current = Ue;
    try {
      const gt = buildAssistantSelectionContext(
          activeWorkbook,
          assistantSelection,
        ),
        Nt = await requestAssistantAnswer(adapter, fe, gt, {
          conversation: Oe,
          signal: Ue.signal,
          reliability: activeReliabilityMonitor,
        });
      if (
        Ue.signal.aborted ||
        Be !== assistantState.assistantRequestIdRef.current
      )
        return null;
      const wt = validateProposedCalculation(
        activeWorkbook,
        assistantSelection,
        Nt,
      );
      return (
        assistantState.setAnswerCitations(
          buildAnswerCitations(activeWorkbook, assistantSelection, wt, gt),
        ),
        assistantState.setAssistantProposal(wt),
        wt
      );
    } catch (gt) {
      if (
        (Be === assistantState.assistantRequestIdRef.current &&
          !Ue.signal.aborted &&
          operationStatusState.setErrorNotice(
            gt instanceof Error
              ? gt.message
              : "The assistant could not complete the request.",
          ),
        Ue.signal.aborted ||
          Be !== assistantState.assistantRequestIdRef.current)
      )
        return null;
      throw gt;
    } finally {
      Be === assistantState.assistantRequestIdRef.current &&
        assistantState.setIsAssistantBusy(false);
    }
  }
  return {
    importWorkbookFromFile,
    exportWorkbookAs,
    runCurrentAuditAnalysis,
    askAssistantQuestion,
  };
}

export { createFileAndAnalysisActions };
