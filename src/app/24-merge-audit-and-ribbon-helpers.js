import { X, ChevronDown } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import __calculationWorkerSource from "../workers/calculation.worker.js?worker-source";
import {
  resolveCalculationBudget,
  estimateWorkbookMemoryBytes,
  isDeepFrozenValue,
  getExtensionCacheEntries,
  estimateValueFootprintBytes,
  workbookHasVolatileFunctions,
  readMonotonicClockMs,
  defineDataPropertyByKey,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  getWorkbookCalculation,
  getCurrentCalculatorRecord,
  assertRangeCellBudget,
  getIsolatedCalculation,
  findSheetOrThrow,
  carryCalculatorRecordForward,
  formatOccupiedRangeAddress,
  calculateWorkbookWithSettings,
  runIterativeCalculation,
  _t,
  assertFieldTransformsUnchanged,
  applyFilterConditions,
  touchWorkbookModifiedAt,
  assertTextCleanupEvidenceUnchanged,
  createWorkbookMetadata,
  applyTableSourceEvidence,
  normalizeWorkbookInput,
  refreshCalculationReport,
  validateRowOrigins,
  acceptWorkerCalculationResult,
  columnIndexToLetters,
  checkDataValidationBatch,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  previewMergeCells,
  assertMergePreviewCurrent,
} from "./23-format-cells-and-conditional-formatting.js";

function MergeCellsPreviewDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    snapshotRef = ne.useRef({
      book: structuredClone(props.workbook),
      sheetId: props.sheetId,
      range: props.range,
    }),
    latestPropsRef = ne.useRef(props);
  latestPropsRef.current = props;
  const [unmerge, setUnmerge] = ne.useState(false),
    [pageIndex, setPageIndex] = ne.useState(0),
    [actionError, setActionError] = ne.useState(""),
    snapshotBookJson = JSON.stringify(snapshotRef.current.book),
    isSnapshotStale =
      snapshotBookJson !== JSON.stringify(props.workbook) ||
      snapshotRef.current.sheetId !== props.sheetId ||
      snapshotRef.current.range !== props.range,
    previewState = ne.useMemo(() => {
      try {
        return {
          plan: previewMergeCells(
            snapshotRef.current.book,
            snapshotRef.current.sheetId,
            snapshotRef.current.range,
            unmerge,
          ),
          error: "",
        };
      } catch (error) {
        return {
          plan: null,
          error:
            error instanceof Error ? error.message : "Unable to preview merge.",
        };
      }
    }, [unmerge]),
    plan = previewState.plan,
    pageCount = Math.max(1, Math.ceil((plan?.cells.length ?? 0) / 12)),
    currentPage = Math.min(pageIndex, pageCount - 1);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  function onClick() {
    try {
      const latestProps = latestPropsRef.current;
      if (
        snapshotBookJson !== JSON.stringify(latestProps.workbook) ||
        snapshotRef.current.sheetId !== latestProps.sheetId ||
        snapshotRef.current.range !== latestProps.range
      )
        throw Error(
          "The workbook or selection changed. Close and reopen this preview.",
        );
      if (!plan) throw Error(previewState.error);
      (latestProps.onWorkbookChange(
        assertMergePreviewCurrent(latestProps.workbook, plan, {
          readOnly: latestProps.readOnly,
        }),
        unmerge ? "Unmerge cells" : "Merge cells",
      ),
        latestProps.onClose());
    } catch (error) {
      setActionError(
        error instanceof Error ? error.message : "Unable to apply merge.",
      );
    }
  }
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(), event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey && document.activeElement === firstFocusable
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "et-header",
          children: [
            i.jsx("h2", {
              id: titleId,
              children: unmerge ? "Unmerge cells" : "Merge & Center",
            }),
            i.jsx("button", {
              "aria-label": text("Close merge preview"),
              onClick: props.onClose,
              lang: language("Close merge preview"),
              "data-ui-label": "Close merge preview",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "et-body",
          children: [
            i.jsx("p", {
              children: i.jsxs("strong", {
                children: [
                  plan?.sheetName ?? "Worksheet",
                  "!",
                  snapshotRef.current.range,
                ],
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Merge shows and centers the upper-left cell. Covered values, formulas, notes and review evidence stay in the native workbook and reappear on Unmerge. Calculated previews are informational; volatile formulas may recalculate.",
              }),
            }),
            i.jsxs("div", {
              className: "et-actions",
              children: [
                i.jsx("button", {
                  onClick: () => {
                    (setUnmerge(false), setPageIndex(0), setActionError(""));
                  },
                  "aria-pressed": !unmerge,
                  children: i.jsx(LocalizedText, { children: "Preview merge" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (setUnmerge(true), setPageIndex(0), setActionError(""));
                  },
                  "aria-pressed": unmerge,
                  children: i.jsx(LocalizedText, {
                    children: "Unmerge selected cells",
                  }),
                }),
              ],
            }),
            plan &&
              i.jsxs("section", {
                "aria-label": text("Merge preview"),
                lang: language("Merge preview"),
                "data-ui-label": "Merge preview",
                children: [
                  i.jsxs("p", {
                    children: [
                      plan.affectedRanges.join(", "),
                      " · ",
                      plan.hiddenValues,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "populated cells hidden by this merge.",
                      }),
                    ],
                  }),
                  plan.xlsxBlocked &&
                    i.jsx("p", {
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Excel export would discard covered values, so XLSX export is blocked until this range is unmerged. Native JSON retains every cell.",
                      }),
                    }),
                  i.jsxs("table", {
                    className: "et-preview-table",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Cell",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Value and evidence retained",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Display effect",
                              }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: plan.cells
                          .slice(currentPage * 12, currentPage * 12 + 12)
                          .map((previewCell) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", {
                                    children: previewCell.address,
                                  }),
                                  i.jsxs("td", {
                                    style: { overflowWrap: "anywhere" },
                                    children: [
                                      i.jsx("div", {
                                        children:
                                          previewCell.display || "(blank)",
                                      }),
                                      i.jsxs("div", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Raw: ",
                                          }),
                                          String(
                                            previewCell.cell.raw ?? "(blank)",
                                          ),
                                        ],
                                      }),
                                      previewCell.cell.note !== undefined &&
                                        i.jsxs("div", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Note: ",
                                            }),
                                            previewCell.cell.note,
                                          ],
                                        }),
                                      previewCell.cell.tickmark !== undefined &&
                                        i.jsxs("div", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Tickmark: ",
                                            }),
                                            previewCell.cell.tickmark,
                                          ],
                                        }),
                                      previewCell.cell.hyperlink &&
                                        i.jsxs("div", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Link:",
                                            }),
                                            " ",
                                            previewCell.cell.hyperlink.kind ===
                                            "external"
                                              ? previewCell.cell.hyperlink.url
                                              : `${previewCell.cell.hyperlink.sheetId}!${previewCell.cell.hyperlink.address}`,
                                          ],
                                        }),
                                      previewCell.discussions > 0 &&
                                        i.jsxs("div", {
                                          children: [
                                            previewCell.discussions,
                                            " ",
                                            i.jsx(LocalizedText, {
                                              children:
                                                "review discussions retained",
                                            }),
                                          ],
                                        }),
                                      Object.keys(previewCell.cell).some(
                                        (key) =>
                                          ![
                                            "raw",
                                            "note",
                                            "tickmark",
                                            "hyperlink",
                                          ].includes(key),
                                      ) &&
                                        i.jsxs("details", {
                                          children: [
                                            i.jsx("summary", {
                                              children: i.jsx(LocalizedText, {
                                                children:
                                                  "Formatting and cell metadata",
                                              }),
                                            }),
                                            i.jsx("pre", {
                                              style: { whiteSpace: "pre-wrap" },
                                              children: JSON.stringify(
                                                previewCell.cell,
                                                null,
                                                2,
                                              ),
                                            }),
                                          ],
                                        }),
                                    ],
                                  }),
                                  i.jsx("td", { children: previewCell.effect }),
                                ],
                              },
                              previewCell.address,
                            ),
                          ),
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "et-pagination",
                    children: [
                      i.jsx("button", {
                        disabled: currentPage === 0,
                        onClick: () => setPageIndex(currentPage - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous preview page",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          currentPage + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          pageCount,
                        ],
                      }),
                      i.jsx("button", {
                        disabled: currentPage === pageCount - 1,
                        onClick: () => setPageIndex(currentPage + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next preview page",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
            (actionError ||
              previewState.error ||
              isSnapshotStale ||
              props.readOnly) &&
              i.jsx("p", {
                role: "alert",
                children:
                  actionError ||
                  previewState.error ||
                  (isSnapshotStale
                    ? "The workbook or selection changed. Close and reopen this preview."
                    : "This workbook is read-only."),
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "et-primary",
              disabled: !plan || isSnapshotStale || props.readOnly,
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Apply" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function applyWorkbookChange(
  workbook,
  nextWorkbook,
  actor,
  actionLabel,
  options = {},
) {
  if (
    ((nextWorkbook = assertTextCleanupEvidenceUnchanged(
      workbook,
      nextWorkbook,
    )),
    (nextWorkbook = assertFieldTransformsUnchanged(workbook, nextWorkbook)),
    (nextWorkbook = applyTableSourceEvidence(workbook, nextWorkbook)),
    validateRowOrigins(workbook, nextWorkbook),
    workbook.id !== nextWorkbook.id ||
      workbook.revision !== nextWorkbook.revision)
  )
    throw new Error(
      "The workbook changed while this tool was open. Reopen the tool and try again.",
    );
  if (workbook.structureProtection?.enabled) {
    const readSheetStructure = (workbook) =>
      workbook.sheets.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        hidden: !!sheet.hidden,
        veryHidden: !!sheet.veryHidden,
        tabColor: sheet.tabColor ?? null,
      }));
    if (
      JSON.stringify(readSheetStructure(workbook)) !==
      JSON.stringify(readSheetStructure(nextWorkbook))
    )
      throw new Error(
        "Workbook structure is protected. Unprotect it before adding, deleting, renaming, moving or hiding worksheets.",
      );
  }
  const cellChanges = [];
  for (const sheetId of new Set([
    ...workbook.sheets.map((sheet) => sheet.id),
    ...nextWorkbook.sheets.map((sheet) => sheet.id),
  ])) {
    const sheet = nextWorkbook.sheets.find((sheet) => sheet.id === sheetId),
      beforeSheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
    if (beforeSheet?.cells !== sheet?.cells)
      for (const address of new Set([
        ...Object.keys(beforeSheet?.cells ?? {}),
        ...Object.keys(sheet?.cells ?? {}),
      ])) {
        const before = beforeSheet?.cells[address],
          after = sheet?.cells[address];
        if (
          !(
            before === after || JSON.stringify(before) === JSON.stringify(after)
          ) &&
          (cellChanges.push({
            sheetId: sheetId,
            address: address,
            ...(before ? { before: before } : {}),
            ...(after ? { after: after } : {}),
          }),
          cellChanges.length > workbookLimits.maxCells)
        )
          throw new Error(
            "This operation exceeds the 100,000 cell-change history limit.",
          );
      }
  }
  if (
    workbook.events.reduce(
      (total, historyEvent) => total + historyEvent.changes.length,
      0,
    ) +
      cellChanges.length >
    workbookLimits.maxCells
  )
    throw new Error(
      "Change history has reached its 100,000 cell-change limit. Export this workbook before continuing.",
    );
  const patchCount = Math.max(
    1,
    Math.ceil(cellChanges.length / workbookLimits.maxPatches),
  );
  if (workbook.events.length + patchCount > 2e3)
    throw new Error(
      "Change history is full. Export this workbook before continuing.",
    );
  const restoredChangeKeys = new Set(),
    changeKey = (changeSheetId, changeAddress, cellSnapshot) =>
      JSON.stringify([
        changeSheetId,
        changeAddress,
        cellSnapshot ?? { raw: null },
      ]);
  if (options.restoringHistory)
    for (const pastEvent of workbook.events)
      for (const change of pastEvent.changes)
        (restoredChangeKeys.add(
          changeKey(change.sheetId, change.address, change.before),
        ),
          restoredChangeKeys.add(
            changeKey(change.sheetId, change.address, change.after),
          ));
  checkDataValidationBatch(
    nextWorkbook,
    cellChanges.filter(
      (cellChange) =>
        nextWorkbook.sheets.some((sheet) => sheet.id === cellChange.sheetId) &&
        !(
          options.restoringHistory &&
          restoredChangeKeys.has(
            changeKey(cellChange.sheetId, cellChange.address, cellChange.after),
          )
        ) &&
        ((cellChange.before?.raw ?? null) !== (cellChange.after?.raw ?? null) ||
          !!cellChange.before?.literal != !!cellChange.after?.literal),
    ),
  );
  const historyEvents = Array.from({ length: patchCount }, (f, patchIndex) => ({
      id: crypto.randomUUID(),
      timestamp: new Date().toISOString(),
      actor: actor,
      action:
        patchCount > 1
          ? `${actionLabel} (${patchIndex + 1}/${patchCount})`
          : actionLabel,
      changes: cellChanges.slice(
        patchIndex * workbookLimits.maxPatches,
        (patchIndex + 1) * workbookLimits.maxPatches,
      ),
    })),
    updatedWorkbook = {
      ...touchWorkbookModifiedAt(nextWorkbook),
      revision: workbook.revision + 1,
      events: [...workbook.events, ...historyEvents],
    };
  return (
    carryCalculatorRecordForward(workbook, updatedWorkbook),
    updatedWorkbook
  );
}
function buildWorkbookSnapshot(workbook) {
  return {
    schemaVersion: workbook.schemaVersion,
    id: workbook.id,
    name: workbook.name,
    revision: workbook.revision,
    events: [],
    calculation: workbook.calculation,
    datePolicy: workbook.datePolicy,
    definedNames: workbook.definedNames,
    workbookLinks: workbook.workbookLinks,
    extensions: workbook.extensions,
    sheets: workbook.sheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      rowCount: sheet.rowCount,
      colCount: sheet.colCount,
      cells: Object.fromEntries(
        Object.entries(sheet.cells).map(([address, cell]) => [
          address,
          { raw: cell.raw, ...(cell.literal ? { literal: true } : {}) },
        ]),
      ),
      hidden: sheet.hidden,
      veryHidden: sheet.veryHidden,
      hiddenRows: sheet.hiddenRows,
      hiddenCols: sheet.hiddenCols,
      tables: sheet.tables,
      merges: sheet.merges,
      outlineGroups: sheet.outlineGroups,
    })),
  };
}
function buildCalculationRequest(workbook, completed, requestId, force) {
  if (
    !completed ||
    !isDeepFrozenValue(workbook) ||
    !isDeepFrozenValue(completed) ||
    completed.id !== workbook.id
  )
    return {
      id: requestId,
      force: force,
      workbook: buildWorkbookSnapshot(workbook),
      extensionAvailability: getExtensionCacheEntries(workbook),
    };
  const removedCells = {},
    sheets = workbook.sheets.map((sheet) => {
      const completedSheet = completed.sheets.find(
          (candidateSheet) => candidateSheet.id === sheet.id,
        ),
        cells = {};
      if (completedSheet?.cells !== sheet.cells) {
        for (const cellAddress in sheet.cells) {
          const cell = sheet.cells[cellAddress],
            completedCell = completedSheet?.cells[cellAddress];
          (!completedCell ||
            cell.raw !== completedCell.raw ||
            !!cell.literal != !!completedCell.literal) &&
            (cells[cellAddress] = {
              raw: cell.raw,
              ...(cell.literal ? { literal: true } : {}),
            });
        }
        completedSheet &&
          (removedCells[sheet.id] = Object.keys(completedSheet.cells).filter(
            (cell) => !Object.prototype.hasOwnProperty.call(sheet.cells, cell),
          ));
      }
      return { ...sheet, cells: cells };
    });
  return {
    id: requestId,
    force: force,
    extensionAvailability: getExtensionCacheEntries(workbook),
    baseRevision: completed.revision,
    removedCells: removedCells,
    workbook: buildWorkbookSnapshot({ ...workbook, sheets: sheets }),
  };
}
class CalculationWorkerSession {
  constructor(factory, reliability) {
    defineDataPropertyByKey(this, "worker");
    defineDataPropertyByKey(this, "nextId", 0);
    defineDataPropertyByKey(this, "completed");
    defineDataPropertyByKey(this, "active");
    ((this.factory = factory), (this.reliability = reliability));
  }
  terminateWorker() {
    const worker = this.worker;
    ((this.worker = undefined), (this.completed = undefined));
    try {
      worker && Promise.resolve(worker.terminate()).catch(() => {});
    } catch {}
  }
  cancel(
    reason = "Calculation cancelled. Inputs are intact; use Calculate Now to retry.",
  ) {
    this.active?.cancel(reason);
  }
  dispose() {
    (this.cancel(), this.terminateWorker());
  }
  run(workbook, options = {}) {
    this.cancel("A newer calculation replaced this request.");
    const report = getIsolatedCalculation(workbook).report,
      runId = ++this.nextId,
      monotonicClockMs = readMonotonicClockMs(),
      { memoryBudgetBytes } = resolveCalculationBudget(workbook.calculation),
      reliabilityRun = this.reliability?.start("calculation", "run");
    return new Promise((resolvePromise) => {
      let isSettled = false,
        timeoutId;
      const settle = (result, shouldTerminateWorker = false) => {
          isSettled ||
            ((isSettled = true),
            timeoutId && clearTimeout(timeoutId),
            options.signal?.removeEventListener("abort", cancelFromSignal),
            (this.active = undefined),
            this.worker &&
              ((this.worker.onmessage = null),
              (this.worker.onerror = null),
              shouldTerminateWorker && this.terminateWorker()),
            resolvePromise(result));
        },
        settleWithSentinel = (errorSentinel, message) => {
          const failureKind =
            errorSentinel === "#TIMEOUT!"
              ? "timeout"
              : errorSentinel === "#LIMIT!"
                ? "limit"
                : errorSentinel === "#PENDING!"
                  ? "cancelled"
                  : "unavailable";
          (reliabilityRun?.finish(
            failureKind === "cancelled" ? "cancelled" : "failed",
            failureKind,
          ),
            settle(
              runIterativeCalculation(
                workbook,
                report,
                errorSentinel,
                message,
                readMonotonicClockMs() - monotonicClockMs,
              ),
              true,
            ));
        },
        cancelFromSignal = () =>
          settleWithSentinel(
            "#PENDING!",
            "Calculation cancelled. Inputs are intact; use Calculate Now to retry.",
          );
      if (
        ((this.active = {
          cancel: (cancelReason) =>
            settleWithSentinel(
              "#PENDING!",
              cancelReason ??
                "Calculation cancelled; source inputs are intact.",
            ),
        }),
        options.signal?.aborted)
      ) {
        cancelFromSignal();
        return;
      }
      if (
        (options.signal?.addEventListener("abort", cancelFromSignal, {
          once: true,
        }),
        estimateWorkbookMemoryBytes(workbook, memoryBudgetBytes) >
          memoryBudgetBytes)
      ) {
        settleWithSentinel(
          "#LIMIT!",
          "Calculation inputs exceed the working-data budget. Reduce the data or increase the budget in Calculation Options.",
        );
        return;
      }
      timeoutId = setTimeout(
        () =>
          settleWithSentinel(
            "#TIMEOUT!",
            "Calculation exceeded its time budget and its worker was terminated. Inputs are intact; simplify formulas or increase the budget, then retry.",
          ),
        workbook.calculation?.timeBudgetMs ?? 2500,
      );
      try {
        ((this.worker ??= this.factory()),
          (this.worker.onmessage = ({ data }) => {
            if (isSettled) return;
            const workerMessage = data;
            if (!(!workerMessage || workerMessage.id !== runId)) {
              if (
                readMonotonicClockMs() - monotonicClockMs >
                (workbook.calculation?.timeBudgetMs ?? 2500)
              ) {
                settleWithSentinel(
                  "#TIMEOUT!",
                  "The worker result arrived after the calculation deadline and was discarded.",
                );
                return;
              }
              if (!workerMessage.result) {
                settleWithSentinel(
                  typeof workerMessage.error == "string" &&
                    workerMessage.error.startsWith("#LIMIT!")
                    ? "#LIMIT!"
                    : "#BLOCKED!",
                  "The calculation worker returned no usable result. Inputs are intact; retry the calculation.",
                );
                return;
              }
              if (
                estimateValueFootprintBytes(
                  workerMessage.result,
                  memoryBudgetBytes,
                ) > memoryBudgetBytes
              ) {
                settleWithSentinel(
                  "#LIMIT!",
                  "Calculation results exceed the working-data budget; no partial results were published.",
                );
                return;
              }
              try {
                const completedReport = {
                    ...workerMessage.result.report,
                    execution: {
                      mode: "worker",
                      state:
                        workerMessage.result.report.execution?.state ===
                        "pending"
                          ? "complete"
                          : (workerMessage.result.report.execution?.state ??
                            "complete"),
                      message: workerMessage.result.report.resources?.stopped
                        ? (workerMessage.result.report.resources.reason ??
                          "Calculation stopped at a resource limit.")
                        : workerMessage.result.report.pending > 0
                          ? "Manual calculation is pending. Use Calculate Now to refresh results. Source inputs were not modified."
                          : "Calculated in an isolated worker. Source inputs were not modified.",
                      wallElapsedMs: readMonotonicClockMs() - monotonicClockMs,
                    },
                    limits: {
                      ...workerMessage.result.report.limits,
                      hardInterrupt: true,
                    },
                  },
                  acceptedCalculation = acceptWorkerCalculationResult(
                    workbook,
                    report,
                    {
                      ...workerMessage.result,
                      report: completedReport,
                    },
                  );
                acceptedCalculation && (this.completed = workbook);
                const stoppedSentinel =
                  acceptedCalculation?.report.resources?.stopped;
                (reliabilityRun?.finish(
                  acceptedCalculation
                    ? stoppedSentinel
                      ? "failed"
                      : "succeeded"
                    : "cancelled",
                  stoppedSentinel
                    ? stoppedSentinel === "#TIMEOUT!"
                      ? "timeout"
                      : "limit"
                    : undefined,
                ),
                  settle(acceptedCalculation));
              } catch {
                settleWithSentinel(
                  "#BLOCKED!",
                  "The calculation worker returned an invalid result. Inputs are intact; retry the calculation.",
                );
              }
            }
          }),
          (this.worker.onerror = () =>
            settleWithSentinel(
              "#BLOCKED!",
              "The calculation worker could not run. Check the host's worker policy, then retry; inputs are intact.",
            )),
          this.worker.postMessage(
            buildCalculationRequest(
              workbook,
              this.completed,
              runId,
              !!options.force,
            ),
          ));
      } catch {
        settleWithSentinel(
          "#BLOCKED!",
          "Isolated calculation is unavailable. The host must allow or provide a calculation worker; source inputs are intact.",
        );
      }
    });
  }
}
const calculationWorkerSourceText = __calculationWorkerSource;
function createCalculationWorker() {
  const workerUrl = URL.createObjectURL(
    new Blob([calculationWorkerSourceText], { type: "text/javascript" }),
  );
  try {
    const worker = new Worker(workerUrl, { name: "workbook-calculation" }),
      workerHandle = {
        postMessage: (message) => worker.postMessage(message),
        terminate: () => {
          (worker.terminate(), URL.revokeObjectURL(workerUrl));
        },
        onmessage: null,
        onerror: null,
      };
    return (
      (worker.onmessage = (messageEvent) =>
        workerHandle.onmessage?.({ data: messageEvent.data })),
      (worker.onerror = (errorEvent) =>
        workerHandle.onerror?.({ message: errorEvent.message })),
      workerHandle
    );
  } catch (error) {
    throw (URL.revokeObjectURL(workerUrl), error);
  }
}
function useCalculationRun(
  activeWorkbook,
  calculationWorkerFactory,
  handleCalculationCompleted,
  activeReliabilityMonitor,
) {
  const resolvedWorkerFactory =
      calculationWorkerFactory === null
        ? undefined
        : (calculationWorkerFactory ??
          (typeof Worker == "function" ? createCalculationWorker : undefined)),
    workerSession = ne.useMemo(
      () =>
        resolvedWorkerFactory
          ? new CalculationWorkerSession(
              resolvedWorkerFactory,
              activeReliabilityMonitor,
            )
          : undefined,
      [resolvedWorkerFactory, activeReliabilityMonitor],
    ),
    [calculationVersion, setCalculationVersion] = ne.useState(0),
    completedHandlerRef = ne.useRef(handleCalculationCompleted);
  completedHandlerRef.current = handleCalculationCompleted;
  const debounceTimerRef = ne.useRef(undefined),
    isolatedWorkbookRef = ne.useRef(undefined);
  workerSession &&
    isolatedWorkbookRef.current !== activeWorkbook &&
    (getIsolatedCalculation(activeWorkbook),
    (isolatedWorkbookRef.current = activeWorkbook));
  const readCalculation = (shouldRefreshReport = false) => {
      const reliabilityRun = activeReliabilityMonitor?.start(
        "calculation",
        "run",
      );
      try {
        const workbookCalculation = shouldRefreshReport
            ? refreshCalculationReport(activeWorkbook)
            : getWorkbookCalculation(activeWorkbook),
          stoppedSentinel = workbookCalculation.report.resources?.stopped;
        return (
          reliabilityRun?.finish(
            stoppedSentinel ? "failed" : "succeeded",
            stoppedSentinel === "#TIMEOUT!"
              ? "timeout"
              : stoppedSentinel
                ? "limit"
                : undefined,
          ),
          workbookCalculation
        );
      } catch (error) {
        throw (reliabilityRun?.fail(error), error);
      }
    },
    calculation = ne.useMemo(
      () =>
        workerSession
          ? (getCurrentCalculatorRecord(activeWorkbook) ??
            getIsolatedCalculation(activeWorkbook))
          : readCalculation(),
      [activeWorkbook, workerSession, calculationVersion],
    );
  return (
    ne.useEffect(() => () => workerSession?.dispose(), [workerSession]),
    ne.useEffect(() => {
      if (!workerSession) return;
      const abortController = new AbortController(),
        debounceTimeoutId = setTimeout(() => {
          ((debounceTimerRef.current = undefined),
            workerSession
              .run(activeWorkbook, { signal: abortController.signal })
              .then((runResult) => {
                !abortController.signal.aborted &&
                  runResult &&
                  (setCalculationVersion(
                    (previousVersion) => previousVersion + 1,
                  ),
                  completedHandlerRef.current());
              }));
        }, 60);
      return (
        (debounceTimerRef.current = debounceTimeoutId),
        () => {
          (clearTimeout(debounceTimeoutId),
            debounceTimerRef.current === debounceTimeoutId &&
              (debounceTimerRef.current = undefined),
            abortController.abort());
        }
      );
    }, [activeWorkbook, workerSession]),
    {
      calculation: calculation,
      run: () => {
        if (
          (clearTimeout(debounceTimerRef.current),
          (debounceTimerRef.current = undefined),
          !workerSession)
        ) {
          (readCalculation(true),
            setCalculationVersion((priorVersion) => priorVersion + 1),
            completedHandlerRef.current());
          return;
        }
        (workerSession
          .run(activeWorkbook, { force: true })
          .then((runOutcome) => {
            runOutcome &&
              (setCalculationVersion((currentVersion) => currentVersion + 1),
              completedHandlerRef.current());
          }),
          setCalculationVersion((priorVersion) => priorVersion + 1),
          completedHandlerRef.current());
      },
      cancel: () => {
        if (debounceTimerRef.current !== undefined && workerSession) {
          (clearTimeout(debounceTimerRef.current),
            (debounceTimerRef.current = undefined));
          const abortController = new AbortController();
          (abortController.abort(),
            workerSession
              .run(activeWorkbook, { signal: abortController.signal })
              .then(() =>
                setCalculationVersion((currentVersion) => currentVersion + 1),
              ));
        } else workerSession?.cancel();
        (setCalculationVersion((priorVersion) => priorVersion + 1),
          completedHandlerRef.current());
      },
      isolated: !!workerSession,
    }
  );
}
const headerCellStyle = { bold: true, fill: "#163c45", color: "#ffffff" };
function buildSampleSheetCells(sheetId, sheetName, rows) {
  const cells = {};
  return (
    rows.forEach((row, rowIndex) =>
      row.forEach((raw, columnIndex) => {
        cells[`${columnIndexToLetters(columnIndex)}${rowIndex + 1}`] = {
          raw: raw,
          ...(rowIndex === 0 ? { style: headerCellStyle } : {}),
        };
      }),
    ),
    {
      id: sheetId,
      name: sheetName,
      cells: cells,
      rowCount: 100,
      colCount: 10,
      frozenRows: 1,
    }
  );
}
/**
 * The sample workbook behind `?demo` and the ribbon's sample command: a year of orders for an
 * invented shop, a summary by region and by product built with formulas across sheets, and a
 * sheet of notes. Written for the rebuilt project; the data is made up.
 */
function buildSampleWorkbook() {
  const orders = [
      ["2025-01-08", "North", "Desk lamp", 24, 39.5, "R. Okafor", "Online"],
      ["2025-01-21", "West", "Standing desk", 6, 489, "M. Lindqvist", "Store"],
      ["2025-02-04", "East", "Office chair", 15, 219, "R. Okafor", "Online"],
      [
        "2025-02-19",
        "South",
        "Monitor arm",
        32,
        84.9,
        "T. Nakamura",
        "Partner",
      ],
      ["2025-03-03", "North", "Office chair", 9, 219, "M. Lindqvist", "Store"],
      ["2025-03-17", "East", "Desk lamp", 40, 39.5, "T. Nakamura", "Online"],
      ["2025-04-02", "West", "Monitor arm", 18, 84.9, "R. Okafor", "Online"],
      [
        "2025-04-24",
        "South",
        "Standing desk",
        4,
        489,
        "M. Lindqvist",
        "Partner",
      ],
      ["2025-05-12", "North", "Monitor arm", 27, 84.9, "T. Nakamura", "Online"],
      ["2025-05-29", "East", "Standing desk", 8, 489, "R. Okafor", "Store"],
      ["2025-06-10", "West", "Desk lamp", 55, 39.5, "M. Lindqvist", "Online"],
      ["2025-06-26", "South", "Office chair", 12, 219, "T. Nakamura", "Store"],
      ["2025-07-15", "North", "Standing desk", 5, 489, "R. Okafor", "Partner"],
      ["2025-08-05", "East", "Monitor arm", 21, 84.9, "M. Lindqvist", "Online"],
      ["2025-08-28", "West", "Office chair", 14, 219, "T. Nakamura", "Online"],
      ["2025-09-16", "South", "Desk lamp", 36, 39.5, "R. Okafor", "Store"],
      ["2025-10-07", "North", "Desk lamp", 48, 39.5, "M. Lindqvist", "Online"],
      ["2025-10-30", "East", "Office chair", 11, 219, "T. Nakamura", "Partner"],
      ["2025-11-18", "West", "Standing desk", 7, 489, "R. Okafor", "Online"],
      [
        "2025-12-09",
        "South",
        "Monitor arm",
        29,
        84.9,
        "M. Lindqvist",
        "Online",
      ],
    ],
    orderRows = [
      [
        "Order ID",
        "Date",
        "Region",
        "Product",
        "Units",
        "Unit price",
        "Revenue",
        "Sales rep",
        "Channel",
        "Status",
      ],
    ];
  orders.forEach(
    (
      [orderDate, region, product, units, unitPrice, salesRep, channel],
      index,
    ) => {
      const rowNumber = orderRows.length + 1;
      orderRows.push([
        `SO-25-${String(index + 1).padStart(3, "0")}`,
        orderDate,
        region,
        product,
        units,
        unitPrice,
        `=E${rowNumber}*F${rowNumber}`,
        salesRep,
        channel,
        index === 17 || index === 19 ? "Open" : "Shipped",
      ]);
    },
  );
  const lastOrderRow = orderRows.length,
    orderTotalRow = lastOrderRow + 1;
  orderRows.push([
    "Total",
    null,
    null,
    null,
    `=SUM(E2:E${lastOrderRow})`,
    null,
    `=SUM(G2:G${lastOrderRow})`,
    null,
    null,
    null,
  ]);
  const ordersSheet = buildSampleSheetCells("orders", "Orders", orderRows);
  for (let rowNumber = 2; rowNumber <= orderTotalRow; rowNumber++)
    for (const amountColumn of ["F", "G"])
      if (ordersSheet.cells[`${amountColumn}${rowNumber}`]?.raw != null)
        ordersSheet.cells[`${amountColumn}${rowNumber}`].format = "number";
  ((ordersSheet.cells.E12.note =
    "Largest order of the year by units. Bulk price agreed with the customer."),
    (ordersSheet.cells.E12.style = { fill: "#fff0d1" }));
  for (const totalColumn of ["A", "E", "G"])
    ordersSheet.cells[`${totalColumn}${orderTotalRow}`].style = {
      bold: true,
      fill: "#e4eef0",
    };

  const orderRange = (column) =>
    `Orders!$${column}$2:$${column}$${lastOrderRow}`;
  /** A summary of the orders by the values of one column, with a total and a check row. */
  const buildSummaryRows = (heading, keyColumn, keys) => {
    const rows = [[heading, "Orders", "Units", "Revenue", "Share of revenue"]],
      lastKeyRow = keys.length + 1,
      totalRow = lastKeyRow + 1;
    keys.forEach((key, keyIndex) => {
      const rowNumber = keyIndex + 2;
      rows.push([
        key,
        `=COUNTIF(${orderRange(keyColumn)},A${rowNumber})`,
        `=SUMIF(${orderRange(keyColumn)},A${rowNumber},${orderRange("E")})`,
        `=SUMIF(${orderRange(keyColumn)},A${rowNumber},${orderRange("G")})`,
        `=D${rowNumber}/$D$${totalRow}`,
      ]);
    });
    rows.push([
      "Total",
      `=SUM(B2:B${lastKeyRow})`,
      `=SUM(C2:C${lastKeyRow})`,
      `=SUM(D2:D${lastKeyRow})`,
      `=IF(ABS(D${totalRow}-Orders!G${orderTotalRow})<0.01,"Matches orders","Check")`,
    ]);
    return { rows, lastKeyRow, totalRow };
  };
  const formatSummarySheet = (sheet, { lastKeyRow, totalRow }) => {
    for (let rowNumber = 2; rowNumber <= totalRow; rowNumber++)
      sheet.cells[`D${rowNumber}`].format = "number";
    for (let rowNumber = 2; rowNumber <= lastKeyRow; rowNumber++)
      sheet.cells[`E${rowNumber}`].format = "percent";
    for (const totalColumn of ["A", "B", "C", "D", "E"])
      sheet.cells[`${totalColumn}${totalRow}`].style = {
        bold: true,
        fill: "#e4eef0",
      };
    return sheet;
  };
  const uniqueSorted = (columnIndex) =>
      [...new Set(orders.map((order) => order[columnIndex]))].sort(),
    regionSummary = buildSummaryRows("Region", "C", uniqueSorted(1)),
    productSummary = buildSummaryRows("Product", "D", uniqueSorted(2)),
    regionSheet = formatSummarySheet(
      buildSampleSheetCells("by-region", "By region", regionSummary.rows),
      regionSummary,
    ),
    productSheet = formatSummarySheet(
      buildSampleSheetCells("by-product", "By product", productSummary.rows),
      productSummary,
    ),
    notesSheet = buildSampleSheetCells("notes", "Notes", [
      ["Sheet", "What it shows", "Try this"],
      [
        "Orders",
        "Twenty orders with a revenue formula on every row",
        "Change a unit count and watch the totals update",
      ],
      [
        "By region",
        "Orders, units and revenue by region, built with COUNTIF and SUMIF across sheets",
        "Select a range and read the statistics in the status bar",
      ],
      [
        "By product",
        "The same summary by product, with each product's share of revenue",
        "Insert a chart or a PivotTable from the Insert tab",
      ],
      [
        "Check",
        "Each summary's last cell compares its total with the Orders total",
        "Open the AI assistant and ask about the selected range",
      ],
      [
        "Sample data",
        "Invented figures for trying the application",
        "Not a real company's records",
      ],
    ]);
  return normalizeWorkbookInput({
    schemaVersion: 1,
    id: "demo-sample-sales-2025",
    name: "Sample sales · 2025",
    revision: 0,
    metadata: createWorkbookMetadata(),
    sheets: [ordersSheet, regionSheet, productSheet, notesSheet],
    events: [],
  });
}
const formatAmountTwoDecimals = (amount) =>
    new Intl.NumberFormat("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(amount),
  normalizeHeaderLabel = (label) =>
    String(label ?? "")
      .trim()
      .toLowerCase()
      .replace(/[ _-]+/g, " "),
  parseDateOnlyText = (dateText) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) return NaN;
    const timestamp = Date.parse(`${dateText}T00:00:00Z`);
    return Number.isFinite(timestamp) &&
      new Date(timestamp).toISOString().slice(0, 10) === dateText
      ? timestamp
      : NaN;
  };
function runAuditAnalysis(
  activeWorkbook,
  sheetRef,
  selectedAuditKind,
  analysisOptions = {},
) {
  const sheetOrThrow = findSheetOrThrow(activeWorkbook, sheetRef),
    workbookWithSettings = calculateWorkbookWithSettings(activeWorkbook),
    occupiedRangeAddress = formatOccupiedRangeAddress(sheetOrThrow, false),
    analysisRange =
      analysisOptions.range ??
      analysisOptions.population?.definition.range ??
      occupiedRangeAddress;
  if (
    analysisOptions.population &&
    analysisOptions.population.scope !== "full" &&
    analysisOptions.population.scope !== "visible"
  )
    throw Error("Choose full or visible analysis scope.");
  const filterContext = analysisOptions.population
    ? applyFilterConditions(
        activeWorkbook,
        sheetRef,
        analysisOptions.population.definition,
        workbookWithSettings,
      ).context
    : undefined;
  if (
    filterContext &&
    analysisRange !== filterContext.range &&
    analysisRange !== analysisOptions.population.definition.range
  )
    throw Error(
      "The analysis range must match the declared filter population.",
    );
  const visibleRowSet =
      analysisOptions.population?.scope === "visible"
        ? new Set(filterContext.visibleSourceRows)
        : null,
    rangeAddresses = assertRangeCellBudget(analysisRange);
  if (
    rangeAddresses.some((cellAddress) => {
      const cellPosition = _t(cellAddress);
      return (
        cellPosition.row >= sheetOrThrow.rowCount ||
        cellPosition.col >= sheetOrThrow.colCount
      );
    })
  )
    throw new Error("Audit range exceeds sheet dimensions.");
  const rowsByIndex = new Map();
  for (const rangeAddress of rangeAddresses) {
    const row = _t(rangeAddress).row;
    (rowsByIndex.has(row) ||
      rowsByIndex.set(row, { row: row, values: [], addresses: [] }),
      rowsByIndex
        .get(row)
        .values.push(workbookWithSettings(sheetRef, rangeAddress)),
      rowsByIndex.get(row).addresses.push(rangeAddress));
  }
  const rowGroups = [...rowsByIndex.values()],
    headerRow = rowGroups.shift(),
    headerLabels = headerRow.values.map(normalizeHeaderLabel),
    nonemptyDataRows = rowGroups.filter(
      (rowGroup) =>
        (!visibleRowSet || visibleRowSet.has(rowGroup.row + 1)) &&
        rowGroup.values.some((value) => value !== null && value !== ""),
    ),
    citeCell = (address) => ({
      workbookId: activeWorkbook.id,
      sheetId: sheetOrThrow.id,
      sheetName: sheetOrThrow.name,
      address: address,
      value: workbookWithSettings(sheetRef, address),
    }),
    citeRow = (rowToCite) => rowToCite.addresses.map(citeCell),
    findHeaderIndex = (...labels) =>
      headerLabels.findIndex((headerLabel) => labels.includes(headerLabel)),
    debitColumn = findHeaderIndex("debit", "debits"),
    creditColumn = findHeaderIndex("credit", "credits"),
    amountColumn = findHeaderIndex(
      "net",
      "net amount",
      "amount",
      "balance",
      "outstanding",
      "outstanding amount",
    ),
    amountColumns =
      debitColumn >= 0 && creditColumn >= 0
        ? [debitColumn, creditColumn]
        : amountColumn >= 0
          ? [amountColumn]
          : debitColumn >= 0
            ? [debitColumn]
            : creditColumn >= 0
              ? [creditColumn]
              : [];
  if (!amountColumns.length) {
    const firstNumericColumn = headerRow.values.findIndex(
      (value, columnIndex) =>
        nonemptyDataRows.some(
          (dataRow) => typeof dataRow.values[columnIndex] == "number",
        ),
    );
    firstNumericColumn >= 0 && amountColumns.push(firstNumericColumn);
  }
  const amountLabel =
      debitColumn >= 0 && creditColumn >= 0
        ? "Debit minus Credit"
        : amountColumns.length
          ? `${String(headerRow.values[amountColumns[0]] ?? "unnamed column")} (${columnIndexToLetters(_t(headerRow.addresses[amountColumns[0]]).col)})${amountColumn < 0 && debitColumn < 0 && creditColumn < 0 ? "; first numeric column fallback" : ""}`
          : "no numeric amount column",
    readAmount = (amountRow) => {
      if (!amountColumns.length) return null;
      const amountValues = amountColumns.map(
        (amountColumnIndex) => amountRow.values[amountColumnIndex],
      );
      return amountValues.some(
        (amountValue) =>
          amountValue !== null &&
          amountValue !== "" &&
          typeof amountValue != "number",
      ) ||
        amountValues.every(
          (amountItem) => amountItem === null || amountItem === "",
        )
        ? null
        : amountColumns.length === 2
          ? Number(amountValues[0] ?? 0) - Number(amountValues[1] ?? 0)
          : Number(amountValues[0]);
    },
    citeAmounts = (rowForAmounts) =>
      amountColumns.map((amountColumnPosition) =>
        citeCell(rowForAmounts.addresses[amountColumnPosition]),
      ),
    findings = [],
    scopeNote = `${sheetOrThrow.name}!${analysisRange}; first selected row is the header; ${nonemptyDataRows.length} nonempty data rows. Population scope: ${analysisOptions.population?.scope === "visible" ? `visible filter population (${filterContext.visibleRows} visible / ${filterContext.totalRows} total data rows); original source addresses retained.` : "full selected source range, including hidden and filtered rows."}`,
    analysisRows = nonemptyDataRows.filter(
      (candidateRow) =>
        !["total", "grand total", "totals"].includes(
          normalizeHeaderLabel(candidateRow.values[0]),
        ),
    ),
    excludedTotalCount = nonemptyDataRows.length - analysisRows.length,
    scopeSummary = `${scopeNote}${excludedTotalCount ? ` Excluded ${excludedTotalCount} row(s) labelled Total, Totals, or Grand total in the first selected column.` : ""}`,
    rowsWithAmounts = analysisRows.map((row) => ({
      row: row,
      amount: readAmount(row),
    })),
    numericRows = rowsWithAmounts.filter(
      (rowWithAmount) => rowWithAmount.amount !== null,
    ),
    excludedAmountCount = rowsWithAmounts.length - numericRows.length,
    amountNote = ` Amount: ${amountLabel}.${excludedAmountCount ? ` Excluded ${excludedAmountCount} row(s) with blank, nonnumeric, or error amount cells.` : ""}`;
  if (selectedAuditKind === "duplicates") {
    const identifierColumn = findHeaderIndex(
        "invoice id",
        "invoice number",
        "invoice no",
        "journal id",
        "je id",
        "transaction id",
        "document number",
      ),
      groupsByKey = new Map();
    for (const analysisRow of analysisRows) {
      if (
        identifierColumn >= 0 &&
        (analysisRow.values[identifierColumn] === null ||
          analysisRow.values[identifierColumn] === "")
      )
        continue;
      const groupKey =
        identifierColumn >= 0
          ? String(analysisRow.values[identifierColumn]).trim().toLowerCase()
          : JSON.stringify(analysisRow.values);
      (groupsByKey.has(groupKey) || groupsByKey.set(groupKey, []),
        groupsByKey.get(groupKey).push(analysisRow));
    }
    for (const [duplicateKey, duplicateRows] of groupsByKey)
      duplicateRows.length > 1 &&
        findings.push({
          id: `duplicate-${duplicateRows[0].row + 1}`,
          title: `${identifierColumn >= 0 ? String(headerRow.values[identifierColumn]) : "Repeated row"}: ${duplicateKey}`,
          detail: `${duplicateRows.length} occurrences on rows ${duplicateRows.map((rowEntry) => rowEntry.row + 1).join(", ")}. Repeated journal IDs can be valid for multi-line journals; inspect the source before concluding an error.`,
          severity: "warning",
          citations: duplicateRows.flatMap((duplicateRow) =>
            identifierColumn >= 0
              ? [
                  citeCell(duplicateRow.addresses[identifierColumn]),
                  ...citeAmounts(duplicateRow),
                ]
              : citeRow(duplicateRow),
          ),
        });
    return {
      title: "Duplicate scan",
      summary: `${scopeSummary} ${findings.length} repeated ${identifierColumn >= 0 ? `${String(headerRow.values[identifierColumn])} group(s), matching trimmed identifiers without case sensitivity` : "entire-row group(s), matching exact evaluated values"}.`,
      findings: findings,
    };
  }
  if (selectedAuditKind === "materiality") {
    const materialityThreshold = analysisOptions.threshold ?? 5e4;
    if (!Number.isFinite(materialityThreshold) || materialityThreshold < 0)
      throw new Error(
        "Materiality threshold must be a nonnegative finite number.",
      );
    if (!amountColumns.length)
      throw new Error(
        "No numeric amount column found. Include a Debit/Credit, Amount, Net, or Balance column.",
      );
    for (const { row, amount } of numericRows)
      Math.abs(amount) >= materialityThreshold &&
        findings.push({
          id: `materiality-${row.row + 1}`,
          title: `Row ${row.row + 1} · ${formatAmountTwoDecimals(amount)}`,
          detail: `Absolute row amount ${formatAmountTwoDecimals(Math.abs(amount))} meets or exceeds the screening threshold ${formatAmountTwoDecimals(materialityThreshold)}. This is a row-level screen; it does not set audit materiality or aggregate a journal.`,
          severity: "high",
          citations: citeRow(row),
        });
    return {
      title: "Materiality screen",
      summary: `${scopeSummary}${amountNote} ${findings.length} row(s) meet the absolute-value threshold of ${formatAmountTwoDecimals(materialityThreshold)}.`,
      findings: findings,
    };
  }
  if (selectedAuditKind === "sample") {
    const sampleSize = analysisOptions.sampleSize ?? 8,
      sampleSeed = analysisOptions.seed ?? "audit-2025";
    if (!Number.isInteger(sampleSize) || sampleSize < 1 || sampleSize > 500)
      throw new Error("Sample size must be a whole number from 1 to 500.");
    if (sampleSeed.length > 200)
      throw new Error("Sample seed must contain at most 200 characters.");
    let seedHash = 2166136261;
    for (let seedIndex = 0; seedIndex < sampleSeed.length; seedIndex++)
      ((seedHash ^= sampleSeed.charCodeAt(seedIndex)),
        (seedHash = Math.imul(seedHash, 16777619)));
    const nextRandom = () => {
        seedHash += 1831565813;
        let mixedValue = seedHash;
        return (
          (mixedValue = Math.imul(
            mixedValue ^ (mixedValue >>> 15),
            mixedValue | 1,
          )),
          (mixedValue ^=
            mixedValue +
            Math.imul(mixedValue ^ (mixedValue >>> 7), mixedValue | 61)),
          ((mixedValue ^ (mixedValue >>> 14)) >>> 0) / 4294967296
        );
      },
      shuffledRows = [...analysisRows];
    for (
      let shuffleIndex = shuffledRows.length - 1;
      shuffleIndex > 0;
      shuffleIndex--
    ) {
      const swapIndex = Math.floor(nextRandom() * (shuffleIndex + 1));
      [shuffledRows[shuffleIndex], shuffledRows[swapIndex]] = [
        shuffledRows[swapIndex],
        shuffledRows[shuffleIndex],
      ];
    }
    return (
      shuffledRows
        .slice(0, sampleSize)
        .sort((leftRow, rightRow) => leftRow.row - rightRow.row)
        .forEach((selectedRow) =>
          findings.push({
            id: `sample-${selectedRow.row + 1}`,
            title: `Selected row ${selectedRow.row + 1}`,
            detail: `Selected without replacement using seed “${sampleSeed}”. Preserve the workbook revision, population range, row order, and seed to reproduce.`,
            severity: "info",
            citations: citeRow(selectedRow),
          }),
        ),
      {
        title: "Reproducible sample",
        summary: `${scopeSummary} Selected ${findings.length} of ${analysisRows.length} rows (requested ${sampleSize}), seed “${sampleSeed}”, FNV-1a seed + Mulberry32/Fisher–Yates. Simple random row selection; no statistical confidence or population conclusion is inferred.${workbookHasVolatileFunctions(activeWorkbook) ? " This workbook contains volatile formulas. The seed reproduces row selection for the same population and order, not changing formula values; preserve a reviewed value snapshot for evidence." : ""}`,
        findings: findings,
      }
    );
  }
  if (selectedAuditKind === "summary") {
    if (!amountColumns.length)
      throw new Error(
        "No numeric amount column found. Include a Debit/Credit, Amount, Net, or Balance column.",
      );
    const accountColumn = findHeaderIndex(
        "account",
        "account name",
        "gl account",
        "category",
      ),
      totalsByAccount = new Map();
    for (const { row, amount } of numericRows) {
      const accountLabel =
        accountColumn < 0
          ? "All rows"
          : String(row.values[accountColumn] ?? "(blank account)");
      totalsByAccount.has(accountLabel) ||
        totalsByAccount.set(accountLabel, { total: 0, rows: [] });
      const accountTotal = totalsByAccount.get(accountLabel);
      ((accountTotal.total += amount), accountTotal.rows.push(row));
    }
    return (
      [...totalsByAccount]
        .sort(([leftAccount], [rightAccount]) =>
          leftAccount.localeCompare(rightAccount),
        )
        .forEach(([accountName, accountGroup], groupIndex) =>
          findings.push({
            id: `summary-${groupIndex}`,
            title: `${accountName} · ${formatAmountTwoDecimals(accountGroup.total)}`,
            detail: `${accountGroup.rows.length} row(s); signed ${amountLabel} totals ${formatAmountTwoDecimals(accountGroup.total)}.`,
            severity: "info",
            citations: accountGroup.rows.flatMap((row) => [
              ...(accountColumn >= 0
                ? [citeCell(row.addresses[accountColumn])]
                : []),
              ...citeAmounts(row),
            ]),
          }),
        ),
      {
        title: "Account summary",
        summary: `${scopeSummary}${amountNote} ${totalsByAccount.size} group(s); signed total ${formatAmountTwoDecimals(numericRows.reduce((runningTotal, signedEntry) => runningTotal + signedEntry.amount, 0))}. Groups use exact evaluated account labels.`,
        findings: findings,
      }
    );
  }
  if (selectedAuditKind === "aging") {
    if (
      !analysisOptions.asOf ||
      !Number.isFinite(parseDateOnlyText(analysisOptions.asOf))
    )
      throw new Error(
        "Aging requires a valid as-of date in YYYY-MM-DD format.",
      );
    if (!amountColumns.length)
      throw new Error("Aging requires a numeric amount column.");
    const dateColumn = findHeaderIndex(
      "due date",
      "invoice date",
      "date",
      "posting date",
    );
    if (dateColumn < 0)
      throw new Error(
        "Aging requires a Date, Due date, Invoice date, or Posting date column.",
      );
    const asOfTime = Date.parse(`${analysisOptions.asOf}T00:00:00Z`),
      agingBuckets = new Map(
        [
          "Future / not due",
          "0–30 days",
          "31–60 days",
          "61–90 days",
          "Over 90 days",
        ].map((bucketLabel) => [bucketLabel, { total: 0, rows: [] }]),
      );
    let excludedDateCount = 0;
    for (const { row, amount } of numericRows) {
      const dateValue = row.values[dateColumn],
        dateTimestamp =
          typeof dateValue == "number"
            ? Date.UTC(1899, 11, 31) +
              (Math.floor(dateValue) > 60
                ? Math.floor(dateValue) - 1
                : Math.floor(dateValue)) *
                864e5
            : typeof dateValue == "string"
              ? parseDateOnlyText(dateValue)
              : NaN;
      if (!Number.isFinite(dateTimestamp) || dateValue === 60) {
        excludedDateCount++;
        continue;
      }
      const ageDays = Math.floor((asOfTime - dateTimestamp) / 864e5),
        ageBucket =
          ageDays < 0
            ? "Future / not due"
            : ageDays <= 30
              ? "0–30 days"
              : ageDays <= 60
                ? "31–60 days"
                : ageDays <= 90
                  ? "61–90 days"
                  : "Over 90 days",
        bucketEntry = agingBuckets.get(ageBucket);
      ((bucketEntry.total += amount), bucketEntry.rows.push(row));
    }
    return (
      [...agingBuckets].forEach(([bucketName, bucketGroup], bucketIndex) =>
        findings.push({
          id: `aging-${bucketIndex}`,
          title: `${bucketName} · ${formatAmountTwoDecimals(bucketGroup.total)}`,
          detail: `${bucketGroup.rows.length} row(s) as of ${analysisOptions.asOf}, based on ${String(headerRow.values[dateColumn])}. ${headerLabels[dateColumn] !== "due date" ? "Age since transaction date; this is not a days-past-due assessment." : "Days relative to due date."}`,
          severity:
            bucketName === "Over 90 days" && bucketGroup.rows.length
              ? "warning"
              : "info",
          citations: bucketGroup.rows.flatMap((row) => [
            citeCell(row.addresses[dateColumn]),
            ...citeAmounts(row),
          ]),
        }),
      ),
      {
        title: "Aging analysis",
        summary: `${scopeSummary}${amountNote} As of ${analysisOptions.asOf}; ${excludedDateCount} additional row(s) excluded for invalid or missing dates. Signed amounts are used; for receivables aging select an open-item population and outstanding amounts.`,
        findings: findings,
      }
    );
  }
  throw new Error("Unsupported audit analysis.");
}
function recordCrossWorkbookImport(
  activeWorkbook,
  sourceWorkbook,
  actor,
  historyDirection,
) {
  if (activeWorkbook.id === sourceWorkbook.id)
    throw Error("Use the ordinary transaction path for the same workbook.");
  const workbookInput = normalizeWorkbookInput(structuredClone(sourceWorkbook)),
    revision = Math.max(activeWorkbook.revision, workbookInput.revision) + 1;
  if (!Number.isSafeInteger(revision))
    throw Error("The workbook revision limit has been reached.");
  if (workbookInput.events.length >= 2e3)
    throw Error(
      "The target workbook history is full. Export a saved copy before continuing.",
    );
  return normalizeWorkbookInput(
    touchWorkbookModifiedAt({
      ...workbookInput,
      revision: revision,
      events: [
        ...workbookInput.events,
        {
          id: crypto.randomUUID(),
          timestamp: new Date().toISOString(),
          actor: actor,
          action: `${historyDirection === "undo" ? "Undo" : "Redo"} file import; previous ID [${activeWorkbook.id}] rev ${activeWorkbook.revision}; source rev ${sourceWorkbook.revision}`,
          changes: [],
        },
      ],
    }),
  );
}
const keyboardShortcutGroups = [
  ["F1", "Open commands and shortcuts", "Workbook, outside dialogs"],
  [
    "Ctrl / Command + F1",
    "Collapse or expand the ribbon",
    "Workbook, outside dialogs",
  ],
  [
    "Tab / Shift + Tab",
    "Move between controls; Enter or Space activates buttons",
    "Ribbon, menus and dialogs",
  ],
  ["Left / Right, Home / End", "Choose a ribbon tab", "Ribbon tabs"],
  ["Down", "Focus the first available command", "Ribbon tabs"],
  [
    "Shift + F10 / Context Menu key",
    "Open cell, row or column commands",
    "Grid",
  ],
  [
    "Up / Down, Home / End",
    "Move between menu commands; Enter activates",
    "Context and File menus",
  ],
  [
    "Escape",
    "Close the current menu/dialog; in the grid, cancel a cut or extra selection first, then focus the ribbon",
    "Workbook",
  ],
  ["Shift + Tab from File", "Return to the preceding host control", "Ribbon"],
  ["Arrow keys / Shift + Arrow keys", "Move / extend the selection", "Grid"],
  [
    "Ctrl / Command + Arrow keys",
    "Jump to an occupied run or worksheet edge",
    "Grid",
  ],
  [
    "Tab / Shift + Tab, Enter / Shift + Enter",
    "Move through cells forward / backward",
    "Grid",
  ],
  [
    "Home, Page Up / Page Down",
    "Move to first visible column or by a page",
    "Grid",
  ],
  [
    "Ctrl / Command + Home / End",
    "Go to first visible cell / last used cell",
    "Grid",
  ],
  ["Ctrl / Command + Page Up / Page Down", "Switch visible worksheets", "Grid"],
  ["Ctrl / Command + A", "Select all cells", "Grid"],
  ["Shift + Space / Ctrl + Space", "Select row / column", "Grid"],
  ["F2 / typing", "Edit the active cell", "Editable grid"],
  ["Delete / Backspace", "Clear cell contents", "Editable grid"],
  [
    "Ctrl / Command + C / X / V",
    "Copy / cut / paste; browser clipboard permission may be required",
    "Grid",
  ],
  ["Ctrl / Command + Alt + V", "Open Paste Special", "Grid"],
  [
    "Ctrl / Command + Z, Ctrl / Command + Y or Shift + Z",
    "Undo / redo",
    "Editable grid",
  ],
  ["Ctrl / Command + B / I / U", "Bold / italic / underline", "Editable grid"],
  ["Ctrl / Command + D / R", "Fill down / right", "Editable grid"],
  ["Ctrl / Command + F / H", "Find / replace", "Grid"],
  ["Ctrl / Command + Shift + F", "Open advanced Find", "Grid"],
  ["F5 / Ctrl / Command + G", "Go To Special", "Grid"],
  ["Ctrl / Command + 1", "Open Format Cells", "Grid"],
  ["F6 / Shift + F6", "Cycle worksheet panes", "Split grid"],
];
function handleDialogKeyDown(keyboardEvent, onClose) {
  if (!keyboardEvent.defaultPrevented) {
    if (keyboardEvent.key === "Escape")
      (keyboardEvent.preventDefault(),
        keyboardEvent.stopPropagation(),
        onClose());
    else if (keyboardEvent.key === "Tab") {
      const focusableElements = Array.from(
          keyboardEvent.currentTarget.querySelectorAll(
            'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
          ),
        ).filter((focusableElement) => !focusableElement.closest("[hidden]")),
        firstFocusable = focusableElements[0],
        lastFocusable = focusableElements.at(-1);
      keyboardEvent.shiftKey && document.activeElement === firstFocusable
        ? (keyboardEvent.preventDefault(),
          lastFocusable == null || lastFocusable.focus())
        : !keyboardEvent.shiftKey &&
          document.activeElement === lastFocusable &&
          (keyboardEvent.preventDefault(),
          firstFocusable == null || firstFocusable.focus());
    }
  }
}
function CommandReferenceDialog({ commands, onClose, onLocate }) {
  const [searchQuery, setSearchQuery] = ne.useState(""),
    searchInputRef = ne.useRef(null),
    dialogTitleId = ne.useId();
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      searchInputRef.current?.focus(),
      () => {
        activeElement?.isConnected &&
          activeElement.focus({ preventScroll: true });
      }
    );
  }, []);
  const matchesSearch = (text) =>
      text.toLowerCase().includes(searchQuery.trim().toLowerCase()),
    matchingCommands = commands.filter((command) =>
      matchesSearch(
        `${command.name} ${command.location} ${command.reason ?? ""}`,
      ),
    ),
    matchingShortcutGroups = keyboardShortcutGroups.filter(
      (keyboardShortcutGroup) => matchesSearch(keyboardShortcutGroup.join(" ")),
    );
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      className: "lg-command-reference",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": dialogTitleId,
      onKeyDown: (keyEvent) => handleDialogKeyDown(keyEvent, onClose),
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: dialogTitleId,
              children: i.jsx(LocalizedText, {
                children: "Commands & shortcuts",
              }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: onClose,
              children: i.jsx(LocalizedText, {
                children: "Close command reference",
              }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Find a command from any ribbon tab or File. “Show command” opens its location and focuses the control. Use Tab to move and Enter or Space to activate. Unavailable commands explain the current condition below.",
          }),
        }),
        i.jsxs("label", {
          children: [
            i.jsx(LocalizedText, { children: "Search commands and shortcuts" }),
            i.jsx("input", {
              ref: searchInputRef,
              value: searchQuery,
              onChange: (event) => setSearchQuery(event.target.value),
            }),
          ],
        }),
        i.jsxs("p", {
          role: "status",
          children: [
            matchingCommands.length,
            " ",
            i.jsx(LocalizedText, { children: "commands and " }),
            matchingShortcutGroups.length,
            " ",
            i.jsx(LocalizedText, { children: "shortcuts" }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-command-results",
          children: [
            matchingShortcutGroups.length > 0 &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Keyboard reference",
                    }),
                  }),
                  i.jsxs("table", {
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Keys",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Action",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Where",
                              }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: matchingShortcutGroups.map(
                          ([shortcutKeys, shortcutAction, shortcutWhere]) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", { children: shortcutKeys }),
                                  i.jsx("td", { children: shortcutAction }),
                                  i.jsx("td", { children: shortcutWhere }),
                                ],
                              },
                              shortcutKeys,
                            ),
                        ),
                      }),
                    ],
                  }),
                ],
              }),
            matchingCommands.length > 0 &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, { children: "Commands" }),
                  }),
                  i.jsx("ul", {
                    children: matchingCommands.map((commandEntry) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            i.jsxs("div", {
                              children: [
                                i.jsx("strong", {
                                  children: commandEntry.name,
                                }),
                                i.jsx("span", {
                                  children: commandEntry.location,
                                }),
                                commandEntry.reason &&
                                  i.jsxs("p", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Unavailable: ",
                                      }),
                                      commandEntry.reason,
                                    ],
                                  }),
                              ],
                            }),
                            !commandEntry.reason &&
                              i.jsx("button", {
                                type: "button",
                                "aria-label": `Show ${commandEntry.name} in ${commandEntry.location}`,
                                onClick: () => onLocate(commandEntry),
                                children: i.jsx(LocalizedText, {
                                  children: "Show command",
                                }),
                              }),
                          ],
                        },
                        commandEntry.id,
                      ),
                    ),
                  }),
                ],
              }),
            !matchingCommands.length &&
              !matchingShortcutGroups.length &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "No matching commands. Try a ribbon tab, action or shortcut name.",
                }),
              }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Shortcuts are scoped to this workbook. Text fields keep native editing shortcuts. Browser, operating-system and host shortcuts not listed here pass through; a host may reserve a shortcut by cancelling it before the workbook handles it. Escape from the grid returns to the ribbon; Shift+Tab from File returns to the host.",
          }),
        }),
      ],
    }),
  });
}
function extractElementText(children) {
  return ne.Children.toArray(children)
    .map((child) =>
      typeof child == "string" || typeof child == "number"
        ? String(child)
        : ne.isValidElement(child)
          ? extractElementText(child.props.children)
          : "",
    )
    .filter(Boolean)
    .join(" ")
    .trim();
}
function collectRibbonCommands(
  ribbonContent,
  tabName,
  componentTypes,
  getDisabledReason,
) {
  const commands = [];
  function mapCommandChildren(
    children,
    locationLabel = tabName,
    inheritedLabel,
  ) {
    return ne.Children.map(children, (childElement) => {
      if (!ne.isValidElement(childElement)) return childElement;
      const props = childElement.props;
      if (props["data-workbook-command"]) return childElement;
      let commandLabel = "",
        disabled = props.disabled,
        injectedProps = {};
      if (
        (childElement.type === componentTypes.group
          ? ((locationLabel = `${tabName} → ${props.label}`),
            props.onDialog &&
              ((commandLabel = props.dialogLabel ?? `${props.label} settings`),
              (disabled = props.dialogDisabled)))
          : childElement.type === componentTypes.command ||
              childElement.type === componentTypes.icon
            ? (commandLabel = props.label ?? "")
            : childElement.type === componentTypes.color
              ? (commandLabel =
                  props.kind === "fill" ? "Fill Color" : "Font Color")
              : ["button", "select", "input"].includes(childElement.type) &&
                props.role !== "tab" &&
                (commandLabel =
                  props["aria-label"] ??
                  props.title ??
                  inheritedLabel ??
                  extractElementText(props.children)),
        commandLabel)
      ) {
        commandLabel = commandLabel.replace(/ \(Ctrl\+[A-Z]\)$/, "");
        const commandId = `${tabName}:${commands.length}`,
          reason = disabled ? getDisabledReason(commandLabel) : undefined;
        (commands.push({
          id: commandId,
          name: commandLabel,
          location: locationLabel,
          reason: reason,
        }),
          childElement.type === componentTypes.group
            ? (injectedProps = { commandId: commandId, reason: reason })
            : typeof childElement.type == "string"
              ? (injectedProps = {
                  "data-workbook-command": commandId,
                  "aria-label": props["aria-label"] ?? commandLabel,
                  "aria-description": reason,
                  ...(reason ? { title: `${commandLabel}: ${reason}` } : {}),
                })
              : (injectedProps = { commandId: commandId, reason: reason }));
      }
      return ne.cloneElement(
        childElement,
        injectedProps,
        props.children === undefined
          ? undefined
          : mapCommandChildren(
              props.children,
              locationLabel,
              childElement.type === "label"
                ? extractElementText(props.children)
                : inheritedLabel,
            ),
      );
    });
  }
  return { content: mapCommandChildren(ribbonContent), commands: commands };
}
const Uh = [
    "Home",
    "Insert",
    "Page Layout",
    "Formulas",
    "Data",
    "Review",
    "View",
  ],
  Wde = [
    "#FFFFFF",
    "#000000",
    "#E7E6E6",
    "#44546A",
    "#4472C4",
    "#ED7D31",
    "#A5A5A5",
    "#FFC000",
    "#5B9BD5",
    "#70AD47",
    "#F2F2F2",
    "#7F7F7F",
    "#D9E2F3",
    "#D9EAD3",
    "#FFF2CC",
    "#FCE4D6",
    "#F4CCCC",
    "#D9D2E9",
    "#CFE2F3",
    "#D0E0E3",
    "#C00000",
    "#FF0000",
    "#FFC000",
    "#FFFF00",
    "#92D050",
    "#00B050",
    "#00B0F0",
    "#0070C0",
    "#002060",
    "#7030A0",
  ],
  resolveRibbonCommandLabel = (label) =>
    ({
      "New Note": "Add note",
      "New / Edit Note": "Add note",
      "Mark Reviewed": "Mark reviewed",
      "Audit Checks": "Audit checks",
      "Show Changes": "Change history",
      "Fill Down": "Fill down",
      "Sort A to Z": "Sort ascending view",
      "Sort Z to A": "Sort descending view",
      "AI Assistant": "AI assistant",
      "Ask AI Assistant": "AI assistant",
    })[label] ?? label.replace(/ \(Ctrl\+[A-Z]\)$/, "");
function RibbonGroup({
  commandId,
  reason,
  label,
  children,
  className = "",
  onDialog,
  dialogLabel,
  dialogDisabled,
}) {
  return i.jsxs("div", {
    className: `xl-ribbon-group ${className}`,
    children: [
      i.jsx("div", { className: "xl-group-commands", children }),
      i.jsxs("div", {
        className: "xl-group-label",
        style: { position: "relative" },
        children: [
          label,
          onDialog &&
            i.jsx("button", {
              "data-workbook-command": commandId,
              "aria-description": reason,
              type: "button",
              title: reason ? `${dialogLabel}: ${reason}` : dialogLabel,
              "aria-label": dialogLabel,
              disabled: dialogDisabled,
              onClick: onDialog,
              style: {
                position: "absolute",
                right: 0,
                bottom: 0,
                border: 0,
                background: "transparent",
                fontSize: 12,
                padding: "0 3px",
                lineHeight: 1.2,
              },
              children: "↘",
            }),
        ],
      }),
    ],
  });
}
function RibbonCommandButton({
  commandId,
  reason,
  icon,
  label,
  onClick,
  disabled,
  large,
  active,
  title,
  className = "",
  dropdown = false,
}) {
  return i.jsxs("button", {
    "data-workbook-command": commandId,
    "aria-description": reason,
    type: "button",
    className: `xl-command ${large ? "xl-command-large" : ""} ${active ? "xl-command-active" : ""} ${className}`,
    onClick,
    disabled,
    title: reason ? `${label}: ${reason}` : (title ?? label),
    "aria-label": resolveRibbonCommandLabel(label),
    "aria-pressed": active === undefined ? undefined : active,
    children: [
      i.jsx("span", { className: "xl-command-icon", children: icon }),
      i.jsx("span", { className: "xl-command-label", children: label }),
      dropdown &&
        i.jsx(ChevronDown, { className: "xl-command-chevron", size: 10 }),
    ],
  });
}
function RibbonIconButton({
  commandId,
  reason,
  icon,
  label,
  onClick,
  disabled,
  active,
  className = "",
}) {
  return i.jsx("button", {
    "data-workbook-command": commandId,
    "aria-description": reason,
    type: "button",
    className: `xl-icon-button ${active ? "xl-command-active" : ""} ${className}`,
    "aria-label": resolveRibbonCommandLabel(label),
    title: reason ? `${label}: ${reason}` : label,
    onClick,
    disabled,
    "aria-pressed": active === undefined ? undefined : active,
    children: icon,
  });
}

export {
  useCalculationRun,
  MergeCellsPreviewDialog,
  buildSampleWorkbook,
  collectRibbonCommands,
  RibbonGroup,
  runAuditAnalysis,
  recordCrossWorkbookImport,
  Uh,
  CommandReferenceDialog,
  Wde,
  RibbonIconButton,
  RibbonCommandButton,
  applyWorkbookChange,
};
