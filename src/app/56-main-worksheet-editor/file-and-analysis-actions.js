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
  adapter,
  actor,
  activeReliabilityMonitor,
  activeWorkbook,
  activeSheet,
  isEditingLocked,
  setActiveDialogMode,
  setAutomationInspection,
  selectionAreas,
  auditPopulationScope,
  activeFilterState,
  setStatusNotice,
  setErrorNotice,
  setIsOperationBusy,
  operationScope,
  latestOperationIdRef,
  activeOperationControllerRef,
  setDelimitedImportExportSource,
  setPendingImportFile,
  setXlsxConversionReview,
  selectedAuditKind,
  setAuditAnalysisResult,
  materialityThreshold,
  auditSampleSize,
  samplingSeed,
  setAgingReportPayload,
  agingAsOfDate,
  auditSourceRange,
  setAssistantProposal,
  setAnswerCitations,
  setIsAssistantBusy,
  filterResult,
  importFileInputRef,
  assistantAbortControllerRef,
  assistantRequestIdRef,
  latestEditStateRef,
  assistantSelection,
  commitWorkbookChange,
  runGuardedAction,
  switchActiveWorksheet,
}) {
  async function importWorkbookFromFile(fe) {
    if (!fe) return;
    activeOperationControllerRef.current?.abort();
    const Oe = new AbortController(),
      Be = ++latestOperationIdRef.current;
    activeOperationControllerRef.current = Oe;
    const Ue = {
        ...operationScope,
        signal: operationScope.signal
          ? AbortSignal.any([Oe.signal, operationScope.signal])
          : Oe.signal,
      },
      Xe = () => Be === latestOperationIdRef.current;
    (setIsOperationBusy(true), setErrorNotice(""));
    try {
      if (isEditingLocked) throw new Error("This workbook is read-only.");
      if (/\.(xls|xlsb|xltx|ods)$/i.test(fe.name)) {
        (setPendingImportFile(fe), setActiveDialogMode("workbookFormats"));
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
        (setAutomationInspection(wt), setActiveDialogMode("automation"));
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
        setDelimitedImportExportSource({
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
        setXlsxConversionReview({
          snapshot: activeWorkbook,
          report: Nt.report,
          imported: Nt.workbook,
          fileName: fe.name,
        });
        return;
      }
      (commitWorkbookChange(Nt.workbook),
        switchActiveWorksheet(Nt.workbook.sheets.find((wt) => !wt.hidden).id),
        setStatusNotice(
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
            (setAutomationInspection(wt), setActiveDialogMode("automation"));
            return;
          }
        } catch {}
      Xe() && setErrorNotice(toFileProcessingError(Nt).message);
    } finally {
      Xe() &&
        ((activeOperationControllerRef.current = null),
        setIsOperationBusy(false),
        importFileInputRef.current && (importFileInputRef.current.value = ""));
    }
  }
  async function exportWorkbookAs(fe) {
    (setIsOperationBusy(true), setErrorNotice(""));
    try {
      const Oe =
        activeWorkbook.name.replace(/[^a-zA-Z0-9 _-]/g, "").slice(0, 80) ||
        "Workbook";
      if (fe === "xlsx") {
        const Be = await exportWorkbookToXlsx(activeWorkbook);
        (Be.report.warnings.push(
          "Full workbook export includes all filtered and hidden rows. Session filter views are not saved in XLSX. Use CSV/TSV with Visible population for a filtered extract.",
        ),
          setXlsxConversionReview({
            snapshot: activeWorkbook,
            report: { ...Be.report, fileName: `${Oe}-derived.xlsx` },
            blob: Be.blob,
            fileName: `${Oe}-derived.xlsx`,
          }));
      }
      (fe === "csv" &&
        setDelimitedImportExportSource({
          snapshot: activeWorkbook,
          sheetId: activeSheet.id,
          ...(filterResult
            ? { population: filterResult.context.definition }
            : {}),
        }),
        fe === "json" &&
          (setStatusNotice(
            "Native JSON exports the full workbook, including all filtered and hidden rows and any named filter views. Active session masks are not saved; load named criteria, preview and apply after reopening.",
          ),
          saveBlobToFile(
            new Blob([encodeNativeWorkbook(activeWorkbook)], {
              type: "application/json",
            }),
            `${Oe}.json`,
          )));
    } catch (Oe) {
      setErrorNotice(Oe instanceof Error ? Oe.message : "Export failed.");
    } finally {
      setIsOperationBusy(false);
    }
  }
  function runCurrentAuditAnalysis() {
    runGuardedAction(() => {
      if (auditPopulationScope === "visible" && !filterResult)
        throw Error(
          "Review and apply a current population filter before visible-only analysis.",
        );
      const fe = runAuditAnalysis(
        activeWorkbook,
        activeSheet.id,
        selectedAuditKind,
        {
          range:
            auditPopulationScope === "visible" && filterResult
              ? filterResult.context.range
              : auditSourceRange || formatOccupiedRangeAddress(activeSheet),
          ...(auditPopulationScope === "visible"
            ? {
                population: {
                  definition:
                    filterResult?.context.definition ??
                    activeFilterState?.definition,
                  scope: "visible",
                },
              }
            : {}),
          threshold: materialityThreshold,
          sampleSize: auditSampleSize,
          seed: samplingSeed,
          asOf: agingAsOfDate,
        },
      );
      (setAuditAnalysisResult(fe),
        setAgingReportPayload({
          schemaVersion: 1,
          workbookId: activeWorkbook.id,
          workbookName: activeWorkbook.name,
          revision: activeWorkbook.revision,
          sheetId: activeSheet.id,
          sheetName: activeSheet.name,
          generatedAt: new Date().toISOString(),
          actor,
          populationScope: auditPopulationScope,
          ...(filterResult ? { filterContext: filterResult.context } : {}),
          kind: selectedAuditKind,
          parameters: {
            range:
              auditPopulationScope === "visible" && filterResult
                ? filterResult.context.range
                : auditSourceRange || formatOccupiedRangeAddress(activeSheet),
            threshold: materialityThreshold,
            sampleSize: auditSampleSize,
            seed: samplingSeed,
            asOf: agingAsOfDate,
          },
          result: fe,
        }),
        setStatusNotice(
          auditPopulationScope === "visible"
            ? "Analysis completed on visible population rows; citations retain original worksheet addresses."
            : "Analysis completed on the full selected source range, including hidden and filtered rows.",
        ));
    });
  }
  async function askAssistantQuestion(fe, Oe) {
    if (!fe.trim()) return null;
    if (selectionAreas.length > 1)
      throw (
        setErrorNotice(
          "Select one rectangular range before asking the assistant.",
        ),
        new Error("Select one rectangular range before asking the assistant.")
      );
    assistantAbortControllerRef.current?.abort();
    const Be = ++assistantRequestIdRef.current;
    (setIsAssistantBusy(true),
      setAssistantProposal(null),
      setAnswerCitations(null),
      setErrorNotice(""),
      setStatusNotice(""));
    const Ue = new AbortController();
    assistantAbortControllerRef.current = Ue;
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
      if (Ue.signal.aborted || Be !== assistantRequestIdRef.current)
        return null;
      const wt = validateProposedCalculation(
        activeWorkbook,
        assistantSelection,
        Nt,
      );
      return (
        setAnswerCitations(
          buildAnswerCitations(activeWorkbook, assistantSelection, wt, gt),
        ),
        setAssistantProposal(wt),
        wt
      );
    } catch (gt) {
      if (
        (Be === assistantRequestIdRef.current &&
          !Ue.signal.aborted &&
          setErrorNotice(
            gt instanceof Error
              ? gt.message
              : "The assistant could not complete the request.",
          ),
        Ue.signal.aborted || Be !== assistantRequestIdRef.current)
      )
        return null;
      throw gt;
    } finally {
      Be === assistantRequestIdRef.current && setIsAssistantBusy(false);
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
