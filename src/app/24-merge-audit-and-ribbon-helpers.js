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

function MergeCellsPreviewDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef({
      book: structuredClone(e.workbook),
      sheetId: e.sheetId,
      range: e.range,
    }),
    s = ne.useRef(e);
  s.current = e;
  const [unmerge, u] = ne.useState(false),
    [l, c] = ne.useState(0),
    [h, f] = ne.useState(""),
    p = JSON.stringify(a.current.book),
    m =
      p !== JSON.stringify(e.workbook) ||
      a.current.sheetId !== e.sheetId ||
      a.current.range !== e.range,
    v = ne.useMemo(() => {
      try {
        return {
          plan: previewMergeCells(
            a.current.book,
            a.current.sheetId,
            a.current.range,
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
    plan = v.plan,
    b = Math.max(1, Math.ceil((plan?.cells.length ?? 0) / 12)),
    y = Math.min(l, b - 1);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  function onClick() {
    try {
      const E = s.current;
      if (
        p !== JSON.stringify(E.workbook) ||
        a.current.sheetId !== E.sheetId ||
        a.current.range !== E.range
      )
        throw Error(
          "The workbook or selection changed. Close and reopen this preview.",
        );
      if (!plan) throw Error(v.error);
      (E.onWorkbookChange(
        assertMergePreviewCurrent(E.workbook, plan, { readOnly: E.readOnly }),
        unmerge ? "Unmerge cells" : "Merge cells",
      ),
        E.onClose());
    } catch (error) {
      f(error instanceof Error ? error.message : "Unable to apply merge.");
    }
  }
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: o,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(), event.stopPropagation(), e.onClose()),
          event.key === "Tab")
        ) {
          const k = Array.from(
              o.current?.querySelectorAll(
                'button:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
            F = k[0],
            I = k.at(-1);
          event.shiftKey && document.activeElement === F
            ? (event.preventDefault(), I == null || I.focus())
            : !event.shiftKey &&
              document.activeElement === I &&
              (event.preventDefault(), F == null || F.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "et-header",
          children: [
            i.jsx("h2", {
              id: n,
              children: unmerge ? "Unmerge cells" : "Merge & Center",
            }),
            i.jsx("button", {
              "aria-label": text("Close merge preview"),
              onClick: e.onClose,
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
                  a.current.range,
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
                    (u(false), c(0), f(""));
                  },
                  "aria-pressed": !unmerge,
                  children: i.jsx(LocalizedText, { children: "Preview merge" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (u(true), c(0), f(""));
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
                          .slice(y * 12, y * 12 + 12)
                          .map((E) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", { children: E.address }),
                                  i.jsxs("td", {
                                    style: { overflowWrap: "anywhere" },
                                    children: [
                                      i.jsx("div", {
                                        children: E.display || "(blank)",
                                      }),
                                      i.jsxs("div", {
                                        children: [
                                          i.jsx(LocalizedText, {
                                            children: "Raw: ",
                                          }),
                                          String(E.cell.raw ?? "(blank)"),
                                        ],
                                      }),
                                      E.cell.note !== undefined &&
                                        i.jsxs("div", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Note: ",
                                            }),
                                            E.cell.note,
                                          ],
                                        }),
                                      E.cell.tickmark !== undefined &&
                                        i.jsxs("div", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Tickmark: ",
                                            }),
                                            E.cell.tickmark,
                                          ],
                                        }),
                                      E.cell.hyperlink &&
                                        i.jsxs("div", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Link:",
                                            }),
                                            " ",
                                            E.cell.hyperlink.kind === "external"
                                              ? E.cell.hyperlink.url
                                              : `${E.cell.hyperlink.sheetId}!${E.cell.hyperlink.address}`,
                                          ],
                                        }),
                                      E.discussions > 0 &&
                                        i.jsxs("div", {
                                          children: [
                                            E.discussions,
                                            " ",
                                            i.jsx(LocalizedText, {
                                              children:
                                                "review discussions retained",
                                            }),
                                          ],
                                        }),
                                      Object.keys(E.cell).some(
                                        (A) =>
                                          ![
                                            "raw",
                                            "note",
                                            "tickmark",
                                            "hyperlink",
                                          ].includes(A),
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
                                                E.cell,
                                                null,
                                                2,
                                              ),
                                            }),
                                          ],
                                        }),
                                    ],
                                  }),
                                  i.jsx("td", { children: E.effect }),
                                ],
                              },
                              E.address,
                            ),
                          ),
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "et-pagination",
                    children: [
                      i.jsx("button", {
                        disabled: y === 0,
                        onClick: () => c(y - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous preview page",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          y + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          b,
                        ],
                      }),
                      i.jsx("button", {
                        disabled: y === b - 1,
                        onClick: () => c(y + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next preview page",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
            (h || v.error || m || e.readOnly) &&
              i.jsx("p", {
                role: "alert",
                children:
                  h ||
                  v.error ||
                  (m
                    ? "The workbook or selection changed. Close and reopen this preview."
                    : "This workbook is read-only."),
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "et-primary",
              disabled: !plan || m || e.readOnly,
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Apply" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function applyWorkbookChange(workbook, r, actor, n, o = {}) {
  if (
    ((r = assertTextCleanupEvidenceUnchanged(workbook, r)),
    (r = assertFieldTransformsUnchanged(workbook, r)),
    (r = applyTableSourceEvidence(workbook, r)),
    validateRowOrigins(workbook, r),
    workbook.id !== r.id || workbook.revision !== r.revision)
  )
    throw new Error(
      "The workbook changed while this tool was open. Reopen the tool and try again.",
    );
  if (workbook.structureProtection?.enabled) {
    const f = (workbook) =>
      workbook.sheets.map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        hidden: !!sheet.hidden,
        veryHidden: !!sheet.veryHidden,
        tabColor: sheet.tabColor ?? null,
      }));
    if (JSON.stringify(f(workbook)) !== JSON.stringify(f(r)))
      throw new Error(
        "Workbook structure is protected. Unprotect it before adding, deleting, renaming, moving or hiding worksheets.",
      );
  }
  const a = [];
  for (const sheetId of new Set([
    ...workbook.sheets.map((sheet) => sheet.id),
    ...r.sheets.map((sheet) => sheet.id),
  ])) {
    const sheet = r.sheets.find((sheet) => sheet.id === sheetId),
      m = workbook.sheets.find((sheet) => sheet.id === sheetId);
    if (m?.cells !== sheet?.cells)
      for (const address of new Set([
        ...Object.keys(m?.cells ?? {}),
        ...Object.keys(sheet?.cells ?? {}),
      ])) {
        const before = m?.cells[address],
          after = sheet?.cells[address];
        if (
          !(
            before === after || JSON.stringify(before) === JSON.stringify(after)
          ) &&
          (a.push({
            sheetId: sheetId,
            address: address,
            ...(before ? { before: before } : {}),
            ...(after ? { after: after } : {}),
          }),
          a.length > workbookLimits.maxCells)
        )
          throw new Error(
            "This operation exceeds the 100,000 cell-change history limit.",
          );
      }
  }
  if (
    workbook.events.reduce((f, p) => f + p.changes.length, 0) + a.length >
    workbookLimits.maxCells
  )
    throw new Error(
      "Change history has reached its 100,000 cell-change limit. Export this workbook before continuing.",
    );
  const s = Math.max(1, Math.ceil(a.length / workbookLimits.maxPatches));
  if (workbook.events.length + s > 2e3)
    throw new Error(
      "Change history is full. Export this workbook before continuing.",
    );
  const d = new Set(),
    u = (f, p, m) => JSON.stringify([f, p, m ?? { raw: null }]);
  if (o.restoringHistory)
    for (const f of workbook.events)
      for (const change of f.changes)
        (d.add(u(change.sheetId, change.address, change.before)),
          d.add(u(change.sheetId, change.address, change.after)));
  checkDataValidationBatch(
    r,
    a.filter(
      (f) =>
        r.sheets.some((sheet) => sheet.id === f.sheetId) &&
        !(o.restoringHistory && d.has(u(f.sheetId, f.address, f.after))) &&
        ((f.before?.raw ?? null) !== (f.after?.raw ?? null) ||
          !!f.before?.literal != !!f.after?.literal),
    ),
  );
  const l = Array.from({ length: s }, (f, p) => ({
      id: crypto.randomUUID(),
      timestamp: new Date().toISOString(),
      actor: actor,
      action: s > 1 ? `${n} (${p + 1}/${s})` : n,
      changes: a.slice(
        p * workbookLimits.maxPatches,
        (p + 1) * workbookLimits.maxPatches,
      ),
    })),
    c = {
      ...touchWorkbookModifiedAt(r),
      revision: workbook.revision + 1,
      events: [...workbook.events, ...l],
    };
  return (carryCalculatorRecordForward(workbook, c), c);
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
        Object.entries(sheet.cells).map(([t, cell]) => [
          t,
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
function buildCalculationRequest(workbook, completed, t, force) {
  if (
    !completed ||
    !isDeepFrozenValue(workbook) ||
    !isDeepFrozenValue(completed) ||
    completed.id !== workbook.id
  )
    return {
      id: t,
      force: force,
      workbook: buildWorkbookSnapshot(workbook),
      extensionAvailability: getExtensionCacheEntries(workbook),
    };
  const removedCells = {},
    sheets = workbook.sheets.map((sheet) => {
      const d = completed.sheets.find((l) => l.id === sheet.id),
        cells = {};
      if (d?.cells !== sheet.cells) {
        for (const l in sheet.cells) {
          const cell = sheet.cells[l],
            h = d?.cells[l];
          (!h || cell.raw !== h.raw || !!cell.literal != !!h.literal) &&
            (cells[l] = {
              raw: cell.raw,
              ...(cell.literal ? { literal: true } : {}),
            });
        }
        d &&
          (removedCells[sheet.id] = Object.keys(d.cells).filter(
            (cell) => !Object.prototype.hasOwnProperty.call(sheet.cells, cell),
          ));
      }
      return { ...sheet, cells: cells };
    });
  return {
    id: t,
    force: force,
    extensionAvailability: getExtensionCacheEntries(workbook),
    baseRevision: completed.revision,
    removedCells: removedCells,
    workbook: buildWorkbookSnapshot({ ...workbook, sheets: sheets }),
  };
}
class CalculationWorkerSession {
  constructor(r, t) {
    defineDataPropertyByKey(this, "worker");
    defineDataPropertyByKey(this, "nextId", 0);
    defineDataPropertyByKey(this, "completed");
    defineDataPropertyByKey(this, "active");
    ((this.factory = r), (this.reliability = t));
  }
  terminateWorker() {
    const worker = this.worker;
    ((this.worker = undefined), (this.completed = undefined));
    try {
      worker && Promise.resolve(worker.terminate()).catch(() => {});
    } catch {}
  }
  cancel(
    r = "Calculation cancelled. Inputs are intact; use Calculate Now to retry.",
  ) {
    this.active?.cancel(r);
  }
  dispose() {
    (this.cancel(), this.terminateWorker());
  }
  run(workbook, t = {}) {
    this.cancel("A newer calculation replaced this request.");
    const report = getIsolatedCalculation(workbook).report,
      o = ++this.nextId,
      monotonicClockMs = readMonotonicClockMs(),
      { memoryBudgetBytes } = resolveCalculationBudget(workbook.calculation),
      d = this.reliability?.start("calculation", "run");
    return new Promise((l) => {
      let c = false,
        h;
      const f = (y, w = false) => {
          c ||
            ((c = true),
            h && clearTimeout(h),
            t.signal?.removeEventListener("abort", m),
            (this.active = undefined),
            this.worker &&
              ((this.worker.onmessage = null),
              (this.worker.onerror = null),
              w && this.terminateWorker()),
            l(y));
        },
        p = (y, w) => {
          const E =
            y === "#TIMEOUT!"
              ? "timeout"
              : y === "#LIMIT!"
                ? "limit"
                : y === "#PENDING!"
                  ? "cancelled"
                  : "unavailable";
          (d?.finish(E === "cancelled" ? "cancelled" : "failed", E),
            f(
              runIterativeCalculation(
                workbook,
                report,
                y,
                w,
                readMonotonicClockMs() - monotonicClockMs,
              ),
              true,
            ));
        },
        m = () =>
          p(
            "#PENDING!",
            "Calculation cancelled. Inputs are intact; use Calculate Now to retry.",
          );
      if (
        ((this.active = {
          cancel: (y) =>
            p(
              "#PENDING!",
              y ?? "Calculation cancelled; source inputs are intact.",
            ),
        }),
        t.signal?.aborted)
      ) {
        m();
        return;
      }
      if (
        (t.signal?.addEventListener("abort", m, { once: true }),
        estimateWorkbookMemoryBytes(workbook, memoryBudgetBytes) >
          memoryBudgetBytes)
      ) {
        p(
          "#LIMIT!",
          "Calculation inputs exceed the working-data budget. Reduce the data or increase the budget in Calculation Options.",
        );
        return;
      }
      h = setTimeout(
        () =>
          p(
            "#TIMEOUT!",
            "Calculation exceeded its time budget and its worker was terminated. Inputs are intact; simplify formulas or increase the budget, then retry.",
          ),
        workbook.calculation?.timeBudgetMs ?? 2500,
      );
      try {
        ((this.worker ??= this.factory()),
          (this.worker.onmessage = ({ data }) => {
            if (c) return;
            const w = data;
            if (!(!w || w.id !== o)) {
              if (
                readMonotonicClockMs() - monotonicClockMs >
                (workbook.calculation?.timeBudgetMs ?? 2500)
              ) {
                p(
                  "#TIMEOUT!",
                  "The worker result arrived after the calculation deadline and was discarded.",
                );
                return;
              }
              if (!w.result) {
                p(
                  typeof w.error == "string" && w.error.startsWith("#LIMIT!")
                    ? "#LIMIT!"
                    : "#BLOCKED!",
                  "The calculation worker returned no usable result. Inputs are intact; retry the calculation.",
                );
                return;
              }
              if (
                estimateValueFootprintBytes(w.result, memoryBudgetBytes) >
                memoryBudgetBytes
              ) {
                p(
                  "#LIMIT!",
                  "Calculation results exceed the working-data budget; no partial results were published.",
                );
                return;
              }
              try {
                const O = {
                    ...w.result.report,
                    execution: {
                      mode: "worker",
                      state:
                        w.result.report.execution?.state === "pending"
                          ? "complete"
                          : (w.result.report.execution?.state ?? "complete"),
                      message: w.result.report.resources?.stopped
                        ? (w.result.report.resources.reason ??
                          "Calculation stopped at a resource limit.")
                        : w.result.report.pending > 0
                          ? "Manual calculation is pending. Use Calculate Now to refresh results. Source inputs were not modified."
                          : "Calculated in an isolated worker. Source inputs were not modified.",
                      wallElapsedMs: readMonotonicClockMs() - monotonicClockMs,
                    },
                    limits: { ...w.result.report.limits, hardInterrupt: true },
                  },
                  j = acceptWorkerCalculationResult(workbook, report, {
                    ...w.result,
                    report: O,
                  });
                j && (this.completed = workbook);
                const z = j?.report.resources?.stopped;
                (d?.finish(
                  j ? (z ? "failed" : "succeeded") : "cancelled",
                  z ? (z === "#TIMEOUT!" ? "timeout" : "limit") : undefined,
                ),
                  f(j));
              } catch {
                p(
                  "#BLOCKED!",
                  "The calculation worker returned an invalid result. Inputs are intact; retry the calculation.",
                );
              }
            }
          }),
          (this.worker.onerror = () =>
            p(
              "#BLOCKED!",
              "The calculation worker could not run. Check the host's worker policy, then retry; inputs are intact.",
            )),
          this.worker.postMessage(
            buildCalculationRequest(workbook, this.completed, o, !!t.force),
          ));
      } catch {
        p(
          "#BLOCKED!",
          "Isolated calculation is unavailable. The host must allow or provide a calculation worker; source inputs are intact.",
        );
      }
    });
  }
}
const calculationWorkerSourceText = __calculationWorkerSource;
function createCalculationWorker() {
  const e = URL.createObjectURL(
    new Blob([calculationWorkerSourceText], { type: "text/javascript" }),
  );
  try {
    const r = new Worker(e, { name: "workbook-calculation" }),
      t = {
        postMessage: (n) => r.postMessage(n),
        terminate: () => {
          (r.terminate(), URL.revokeObjectURL(e));
        },
        onmessage: null,
        onerror: null,
      };
    return (
      (r.onmessage = (n) => t.onmessage?.({ data: n.data })),
      (r.onerror = (n) => t.onerror?.({ message: n.message })),
      t
    );
  } catch (error) {
    throw (URL.revokeObjectURL(e), error);
  }
}
function useCalculationRun(
  activeWorkbook,
  calculationWorkerFactory,
  handleCalculationCompleted,
  activeReliabilityMonitor,
) {
  const a =
      calculationWorkerFactory === null
        ? undefined
        : (calculationWorkerFactory ??
          (typeof Worker == "function" ? createCalculationWorker : undefined)),
    s = ne.useMemo(
      () =>
        a
          ? new CalculationWorkerSession(a, activeReliabilityMonitor)
          : undefined,
      [a, activeReliabilityMonitor],
    ),
    [d, u] = ne.useState(0),
    l = ne.useRef(handleCalculationCompleted);
  l.current = handleCalculationCompleted;
  const c = ne.useRef(undefined),
    h = ne.useRef(undefined);
  s &&
    h.current !== activeWorkbook &&
    (getIsolatedCalculation(activeWorkbook), (h.current = activeWorkbook));
  const f = (g = false) => {
      const b = activeReliabilityMonitor?.start("calculation", "run");
      try {
        const w = g
            ? refreshCalculationReport(activeWorkbook)
            : getWorkbookCalculation(activeWorkbook),
          E = w.report.resources?.stopped;
        return (
          b?.finish(
            E ? "failed" : "succeeded",
            E === "#TIMEOUT!" ? "timeout" : E ? "limit" : undefined,
          ),
          w
        );
      } catch (error) {
        throw (b?.fail(error), error);
      }
    },
    calculation = ne.useMemo(
      () =>
        s
          ? (getCurrentCalculatorRecord(activeWorkbook) ??
            getIsolatedCalculation(activeWorkbook))
          : f(),
      [activeWorkbook, s, d],
    );
  return (
    ne.useEffect(() => () => s?.dispose(), [s]),
    ne.useEffect(() => {
      if (!s) return;
      const g = new AbortController(),
        b = setTimeout(() => {
          ((c.current = undefined),
            s.run(activeWorkbook, { signal: g.signal }).then((y) => {
              !g.signal.aborted && y && (u((w) => w + 1), l.current());
            }));
        }, 60);
      return (
        (c.current = b),
        () => {
          (clearTimeout(b),
            c.current === b && (c.current = undefined),
            g.abort());
        }
      );
    }, [activeWorkbook, s]),
    {
      calculation: calculation,
      run: () => {
        if ((clearTimeout(c.current), (c.current = undefined), !s)) {
          (f(true), u((g) => g + 1), l.current());
          return;
        }
        (s.run(activeWorkbook, { force: true }).then((g) => {
          g && (u((b) => b + 1), l.current());
        }),
          u((g) => g + 1),
          l.current());
      },
      cancel: () => {
        if (c.current !== undefined && s) {
          (clearTimeout(c.current), (c.current = undefined));
          const g = new AbortController();
          (g.abort(),
            s
              .run(activeWorkbook, { signal: g.signal })
              .then(() => u((b) => b + 1)));
        } else s?.cancel();
        (u((g) => g + 1), l.current());
      },
      isolated: !!s,
    }
  );
}
const headerCellStyle = { bold: true, fill: "#163c45", color: "#ffffff" };
function buildSampleSheetCells(e, r, t) {
  const cells = {};
  return (
    t.forEach((o, a) =>
      o.forEach((raw, d) => {
        cells[`${columnIndexToLetters(d)}${a + 1}`] = {
          raw: raw,
          ...(a === 0 ? { style: headerCellStyle } : {}),
        };
      }),
    ),
    { id: e, name: r, cells: cells, rowCount: 100, colCount: 10, frozenRows: 1 }
  );
}
function buildSampleLedgerWorkbook() {
  const e = [
      [
        "2025-01-03",
        "Cash",
        "Equity",
        "Opening capital contribution",
        25e4,
        "CAP-001",
      ],
      [
        "2025-01-07",
        "Equipment",
        "Cash",
        "Network equipment acquisition",
        185e3,
        "INV-4821",
      ],
      [
        "2025-01-12",
        "Rent expense",
        "Cash",
        "January office lease",
        12500,
        "LEASE-01",
      ],
      [
        "2025-01-18",
        "Accounts receivable",
        "Revenue",
        "Advisory services – Westbridge",
        64e3,
        "SALES-1041",
      ],
      [
        "2025-01-25",
        "Payroll expense",
        "Cash",
        "January payroll",
        43820,
        "PAY-0125",
      ],
      [
        "2025-02-03",
        "Cash",
        "Accounts receivable",
        "Westbridge remittance",
        64e3,
        "RCPT-2041",
      ],
      [
        "2025-02-11",
        "Professional fees",
        "Accounts payable",
        "Legal contract review",
        8750,
        "INV-4956",
      ],
      [
        "2025-02-19",
        "Accounts receivable",
        "Revenue",
        "Advisory services – Harbor",
        32500,
        "SALES-1042",
      ],
      [
        "2025-03-02",
        "Rent expense",
        "Cash",
        "March office lease",
        12500,
        "LEASE-03",
      ],
      [
        "2025-03-17",
        "Software expense",
        "Cash",
        "Annual analytics subscription",
        16800,
        "INV-5200",
      ],
      [
        "2025-04-04",
        "Accounts payable",
        "Cash",
        "Legal invoice settlement",
        8750,
        "PMT-3191",
      ],
      [
        "2025-04-22",
        "Accounts receivable",
        "Revenue",
        "Advisory services – Cedar",
        91200,
        "SALES-1043",
      ],
      [
        "2025-05-09",
        "Travel expense",
        "Cash",
        "Fieldwork travel reimbursement",
        3240.75,
        "EXP-0638",
      ],
      [
        "2025-06-14",
        "Cash",
        "Accounts receivable",
        "Harbor remittance",
        32500,
        "RCPT-2042",
      ],
      [
        "2025-07-30",
        "Payroll expense",
        "Cash",
        "July payroll",
        45210,
        "PAY-0730",
      ],
      [
        "2025-08-21",
        "Insurance expense",
        "Cash",
        "Professional indemnity premium",
        14200,
        "POL-2025",
      ],
      [
        "2025-09-12",
        "Accounts receivable",
        "Revenue",
        "Advisory services – Elm",
        48750,
        "SALES-1044",
      ],
      [
        "2025-10-20",
        "Professional fees",
        "Accounts payable",
        "Specialist valuation support",
        21750,
        "INV-6782",
      ],
      [
        "2025-11-28",
        "Cash",
        "Accounts receivable",
        "Cedar partial remittance",
        6e4,
        "RCPT-2043",
      ],
      [
        "2025-12-31",
        "Depreciation expense",
        "Accumulated depreciation",
        "Year-end depreciation charge",
        30833.33,
        "FA-2025",
      ],
    ],
    r = [
      [
        "Journal ID",
        "Date",
        "Account",
        "Description",
        "Debit",
        "Credit",
        "Net",
        "Prepared by",
        "Source",
        "Review",
      ],
    ];
  (e.forEach(([l, c, h, f, p, m], v) => {
    const g = `JE-25-${String(v + 1).padStart(3, "0")}`;
    for (const b of [0, 1]) {
      const y = r.length + 1;
      r.push([
        `${g}/${b + 1}`,
        l,
        b === 0 ? c : h,
        f,
        b === 0 ? p : 0,
        b === 1 ? p : 0,
        `=E${y}-F${y}`,
        v % 2 === 0 ? "A. Patel" : "J. Morgan",
        m,
        v === 1 || v === 8 ? "Pending" : "Reviewed",
      ]);
    }
  }),
    (r[17][0] = r[15][0]),
    r.push([
      "Total",
      null,
      null,
      null,
      "=SUM(E2:E41)",
      "=SUM(F2:F41)",
      "=SUM(G2:G41)",
      null,
      null,
      null,
    ]));
  const sampleSheetCells = buildSampleSheetCells(
    "general-ledger",
    "General ledger",
    r,
  );
  for (let l = 2; l <= 42; l++)
    for (const c of ["E", "F", "G"])
      sampleSheetCells.cells[`${c}${l}`].format = "number";
  ((sampleSheetCells.cells.A18.note =
    "Identifier repeats A16. Inspect source documentation and posting sequence before resolving."),
    (sampleSheetCells.cells.A18.style = { fill: "#fff0d1" }),
    (sampleSheetCells.cells.E4.note =
      "Agreed acquisition amount to equipment invoice INV-4821."),
    (sampleSheetCells.cells.E4.tickmark = "V"));
  for (const l of ["A", "E", "F", "G"])
    sampleSheetCells.cells[`${l}42`].style = { bold: true, fill: "#e4eef0" };
  const n = [...new Set(e.flatMap((l) => [l[1], l[2]]))].sort(),
    o = [
      ["Account", "Debit activity", "Credit activity", "Net balance", "Status"],
    ];
  n.forEach((l, c) => {
    const h = c + 2;
    o.push([
      l,
      `=SUMIF('General ledger'!$C$2:$C$41,A${h},'General ledger'!$E$2:$E$41)`,
      `=SUMIF('General ledger'!$C$2:$C$41,A${h},'General ledger'!$F$2:$F$41)`,
      `=B${h}-C${h}`,
      "Ready for review",
    ]);
  });
  const a = n.length + 1,
    s = a + 1;
  o.push([
    "Total",
    `=SUM(B2:B${a})`,
    `=SUM(C2:C${a})`,
    `=SUM(D2:D${a})`,
    `=IF(ABS(D${s})<0.01,"Balanced","Investigate")`,
  ]);
  const d = buildSampleSheetCells("trial-balance", "Trial balance", o);
  for (let l = 2; l <= s; l++)
    for (const c of ["B", "C", "D"]) d.cells[`${c}${l}`].format = "number";
  for (const l of ["A", "B", "C", "D", "E"])
    d.cells[`${l}${s}`].style = { bold: true, fill: "#e4eef0" };
  const u = buildSampleSheetCells("audit-plan", "Audit plan", [
    ["Procedure", "Purpose", "Owner", "Status", "Reference"],
    [
      "Ledger reconciliation",
      "Confirm debit and credit control totals",
      "A. Patel",
      "Complete",
      "General ledger!E42:G42",
    ],
    [
      "Duplicate scan",
      "Inspect repeated journal-line identifiers",
      "J. Morgan",
      "Pending",
      "General ledger!A1:J41",
    ],
    [
      "Materiality screen",
      "Screen absolute row amounts at 50,000",
      "A. Patel",
      "Pending",
      "General ledger!A1:J41",
    ],
    [
      "Sample selection",
      "Select 8 reproducible rows for inspection",
      "J. Morgan",
      "Pending",
      "Seed: audit-2025",
    ],
    [
      "Source vouching",
      "Inspect invoice and posting evidence for selected rows",
      "A. Patel",
      "In progress",
      "Tickmark V = vouched",
    ],
    [
      "Account summary",
      "Reconcile signed movement by account",
      "J. Morgan",
      "Pending",
      "Trial balance",
    ],
    [
      "Demo data",
      "Synthetic journal data for product evaluation",
      "System",
      "Information",
      "Not a real client workbook",
    ],
  ]);
  return normalizeWorkbookInput({
    schemaVersion: 1,
    id: "demo-northstar-fy2025",
    name: "Northstar Advisory · FY2025",
    revision: 0,
    metadata: createWorkbookMetadata(),
    sheets: [sampleSheetCells, d, u],
    events: [],
  });
}
const formatAmountTwoDecimals = (e) =>
    new Intl.NumberFormat("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    }).format(e),
  normalizeHeaderLabel = (e) =>
    String(e ?? "")
      .trim()
      .toLowerCase()
      .replace(/[ _-]+/g, " "),
  parseDateOnlyText = (e) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(e)) return NaN;
    const r = Date.parse(`${e}T00:00:00Z`);
    return Number.isFinite(r) && new Date(r).toISOString().slice(0, 10) === e
      ? r
      : NaN;
  };
function runAuditAnalysis(activeWorkbook, r, selectedAuditKind, n = {}) {
  const sheetOrThrow = findSheetOrThrow(activeWorkbook, r),
    workbookWithSettings = calculateWorkbookWithSettings(activeWorkbook),
    occupiedRangeAddress = formatOccupiedRangeAddress(sheetOrThrow, false),
    d = n.range ?? n.population?.definition.range ?? occupiedRangeAddress;
  if (
    n.population &&
    n.population.scope !== "full" &&
    n.population.scope !== "visible"
  )
    throw Error("Choose full or visible analysis scope.");
  const u = n.population
    ? applyFilterConditions(
        activeWorkbook,
        r,
        n.population.definition,
        workbookWithSettings,
      ).context
    : undefined;
  if (u && d !== u.range && d !== n.population.definition.range)
    throw Error(
      "The analysis range must match the declared filter population.",
    );
  const l =
      n.population?.scope === "visible" ? new Set(u.visibleSourceRows) : null,
    c = assertRangeCellBudget(d);
  if (
    c.some((U) => {
      const G = _t(U);
      return G.row >= sheetOrThrow.rowCount || G.col >= sheetOrThrow.colCount;
    })
  )
    throw new Error("Audit range exceeds sheet dimensions.");
  const h = new Map();
  for (const U of c) {
    const row = _t(U).row;
    (h.has(row) || h.set(row, { row: row, values: [], addresses: [] }),
      h.get(row).values.push(workbookWithSettings(r, U)),
      h.get(row).addresses.push(U));
  }
  const f = [...h.values()],
    p = f.shift(),
    m = p.values.map(normalizeHeaderLabel),
    v = f.filter(
      (U) =>
        (!l || l.has(U.row + 1)) &&
        U.values.some((value) => value !== null && value !== ""),
    ),
    g = (address) => ({
      workbookId: activeWorkbook.id,
      sheetId: sheetOrThrow.id,
      sheetName: sheetOrThrow.name,
      address: address,
      value: workbookWithSettings(r, address),
    }),
    b = (U) => U.addresses.map(g),
    y = (...U) => m.findIndex((G) => U.includes(G)),
    w = y("debit", "debits"),
    E = y("credit", "credits"),
    A = y(
      "net",
      "net amount",
      "amount",
      "balance",
      "outstanding",
      "outstanding amount",
    ),
    k =
      w >= 0 && E >= 0
        ? [w, E]
        : A >= 0
          ? [A]
          : w >= 0
            ? [w]
            : E >= 0
              ? [E]
              : [];
  if (!k.length) {
    const U = p.values.findIndex((value, H) =>
      v.some((Q) => typeof Q.values[H] == "number"),
    );
    U >= 0 && k.push(U);
  }
  const F =
      w >= 0 && E >= 0
        ? "Debit minus Credit"
        : k.length
          ? `${String(p.values[k[0]] ?? "unnamed column")} (${columnIndexToLetters(_t(p.addresses[k[0]]).col)})${A < 0 && w < 0 && E < 0 ? "; first numeric column fallback" : ""}`
          : "no numeric amount column",
    I = (U) => {
      if (!k.length) return null;
      const G = k.map((H) => U.values[H]);
      return G.some((H) => H !== null && H !== "" && typeof H != "number") ||
        G.every((H) => H === null || H === "")
        ? null
        : k.length === 2
          ? Number(G[0] ?? 0) - Number(G[1] ?? 0)
          : Number(G[0]);
    },
    O = (U) => k.map((G) => g(U.addresses[G])),
    findings = [],
    z = `${sheetOrThrow.name}!${d}; first selected row is the header; ${v.length} nonempty data rows. Population scope: ${n.population?.scope === "visible" ? `visible filter population (${u.visibleRows} visible / ${u.totalRows} total data rows); original source addresses retained.` : "full selected source range, including hidden and filtered rows."}`,
    V = v.filter(
      (U) =>
        !["total", "grand total", "totals"].includes(
          normalizeHeaderLabel(U.values[0]),
        ),
    ),
    D = v.length - V.length,
    B = `${z}${D ? ` Excluded ${D} row(s) labelled Total, Totals, or Grand total in the first selected column.` : ""}`,
    _ = V.map((row) => ({ row: row, amount: I(row) })),
    C = _.filter((U) => U.amount !== null),
    T = _.length - C.length,
    N = ` Amount: ${F}.${T ? ` Excluded ${T} row(s) with blank, nonnumeric, or error amount cells.` : ""}`;
  if (selectedAuditKind === "duplicates") {
    const U = y(
        "invoice id",
        "invoice number",
        "invoice no",
        "journal id",
        "je id",
        "transaction id",
        "document number",
      ),
      G = new Map();
    for (const H of V) {
      if (U >= 0 && (H.values[U] === null || H.values[U] === "")) continue;
      const Q =
        U >= 0
          ? String(H.values[U]).trim().toLowerCase()
          : JSON.stringify(H.values);
      (G.has(Q) || G.set(Q, []), G.get(Q).push(H));
    }
    for (const [H, Q] of G)
      Q.length > 1 &&
        findings.push({
          id: `duplicate-${Q[0].row + 1}`,
          title: `${U >= 0 ? String(p.values[U]) : "Repeated row"}: ${H}`,
          detail: `${Q.length} occurrences on rows ${Q.map((X) => X.row + 1).join(", ")}. Repeated journal IDs can be valid for multi-line journals; inspect the source before concluding an error.`,
          severity: "warning",
          citations: Q.flatMap((X) =>
            U >= 0 ? [g(X.addresses[U]), ...O(X)] : b(X),
          ),
        });
    return {
      title: "Duplicate scan",
      summary: `${B} ${findings.length} repeated ${U >= 0 ? `${String(p.values[U])} group(s), matching trimmed identifiers without case sensitivity` : "entire-row group(s), matching exact evaluated values"}.`,
      findings: findings,
    };
  }
  if (selectedAuditKind === "materiality") {
    const U = n.threshold ?? 5e4;
    if (!Number.isFinite(U) || U < 0)
      throw new Error(
        "Materiality threshold must be a nonnegative finite number.",
      );
    if (!k.length)
      throw new Error(
        "No numeric amount column found. Include a Debit/Credit, Amount, Net, or Balance column.",
      );
    for (const { row, amount } of C)
      Math.abs(amount) >= U &&
        findings.push({
          id: `materiality-${row.row + 1}`,
          title: `Row ${row.row + 1} · ${formatAmountTwoDecimals(amount)}`,
          detail: `Absolute row amount ${formatAmountTwoDecimals(Math.abs(amount))} meets or exceeds the screening threshold ${formatAmountTwoDecimals(U)}. This is a row-level screen; it does not set audit materiality or aggregate a journal.`,
          severity: "high",
          citations: b(row),
        });
    return {
      title: "Materiality screen",
      summary: `${B}${N} ${findings.length} row(s) meet the absolute-value threshold of ${formatAmountTwoDecimals(U)}.`,
      findings: findings,
    };
  }
  if (selectedAuditKind === "sample") {
    const U = n.sampleSize ?? 8,
      G = n.seed ?? "audit-2025";
    if (!Number.isInteger(U) || U < 1 || U > 500)
      throw new Error("Sample size must be a whole number from 1 to 500.");
    if (G.length > 200)
      throw new Error("Sample seed must contain at most 200 characters.");
    let H = 2166136261;
    for (let J = 0; J < G.length; J++)
      ((H ^= G.charCodeAt(J)), (H = Math.imul(H, 16777619)));
    const Q = () => {
        H += 1831565813;
        let J = H;
        return (
          (J = Math.imul(J ^ (J >>> 15), J | 1)),
          (J ^= J + Math.imul(J ^ (J >>> 7), J | 61)),
          ((J ^ (J >>> 14)) >>> 0) / 4294967296
        );
      },
      X = [...V];
    for (let J = X.length - 1; J > 0; J--) {
      const q = Math.floor(Q() * (J + 1));
      [X[J], X[q]] = [X[q], X[J]];
    }
    return (
      X.slice(0, U)
        .sort((J, q) => J.row - q.row)
        .forEach((J) =>
          findings.push({
            id: `sample-${J.row + 1}`,
            title: `Selected row ${J.row + 1}`,
            detail: `Selected without replacement using seed “${G}”. Preserve the workbook revision, population range, row order, and seed to reproduce.`,
            severity: "info",
            citations: b(J),
          }),
        ),
      {
        title: "Reproducible sample",
        summary: `${B} Selected ${findings.length} of ${V.length} rows (requested ${U}), seed “${G}”, FNV-1a seed + Mulberry32/Fisher–Yates. Simple random row selection; no statistical confidence or population conclusion is inferred.${workbookHasVolatileFunctions(activeWorkbook) ? " This workbook contains volatile formulas. The seed reproduces row selection for the same population and order, not changing formula values; preserve a reviewed value snapshot for evidence." : ""}`,
        findings: findings,
      }
    );
  }
  if (selectedAuditKind === "summary") {
    if (!k.length)
      throw new Error(
        "No numeric amount column found. Include a Debit/Credit, Amount, Net, or Balance column.",
      );
    const U = y("account", "account name", "gl account", "category"),
      G = new Map();
    for (const { row, amount } of C) {
      const X = U < 0 ? "All rows" : String(row.values[U] ?? "(blank account)");
      G.has(X) || G.set(X, { total: 0, rows: [] });
      const J = G.get(X);
      ((J.total += amount), J.rows.push(row));
    }
    return (
      [...G]
        .sort(([H], [Q]) => H.localeCompare(Q))
        .forEach(([H, Q], X) =>
          findings.push({
            id: `summary-${X}`,
            title: `${H} · ${formatAmountTwoDecimals(Q.total)}`,
            detail: `${Q.rows.length} row(s); signed ${F} totals ${formatAmountTwoDecimals(Q.total)}.`,
            severity: "info",
            citations: Q.rows.flatMap((row) => [
              ...(U >= 0 ? [g(row.addresses[U])] : []),
              ...O(row),
            ]),
          }),
        ),
      {
        title: "Account summary",
        summary: `${B}${N} ${G.size} group(s); signed total ${formatAmountTwoDecimals(C.reduce((H, Q) => H + Q.amount, 0))}. Groups use exact evaluated account labels.`,
        findings: findings,
      }
    );
  }
  if (selectedAuditKind === "aging") {
    if (!n.asOf || !Number.isFinite(parseDateOnlyText(n.asOf)))
      throw new Error(
        "Aging requires a valid as-of date in YYYY-MM-DD format.",
      );
    if (!k.length) throw new Error("Aging requires a numeric amount column.");
    const U = y("due date", "invoice date", "date", "posting date");
    if (U < 0)
      throw new Error(
        "Aging requires a Date, Due date, Invoice date, or Posting date column.",
      );
    const G = Date.parse(`${n.asOf}T00:00:00Z`),
      H = new Map(
        [
          "Future / not due",
          "0–30 days",
          "31–60 days",
          "61–90 days",
          "Over 90 days",
        ].map((X) => [X, { total: 0, rows: [] }]),
      );
    let Q = 0;
    for (const { row, amount } of C) {
      const q = row.values[U],
        W =
          typeof q == "number"
            ? Date.UTC(1899, 11, 31) +
              (Math.floor(q) > 60 ? Math.floor(q) - 1 : Math.floor(q)) * 864e5
            : typeof q == "string"
              ? parseDateOnlyText(q)
              : NaN;
      if (!Number.isFinite(W) || q === 60) {
        Q++;
        continue;
      }
      const K = Math.floor((G - W) / 864e5),
        te =
          K < 0
            ? "Future / not due"
            : K <= 30
              ? "0–30 days"
              : K <= 60
                ? "31–60 days"
                : K <= 90
                  ? "61–90 days"
                  : "Over 90 days",
        ae = H.get(te);
      ((ae.total += amount), ae.rows.push(row));
    }
    return (
      [...H].forEach(([X, J], q) =>
        findings.push({
          id: `aging-${q}`,
          title: `${X} · ${formatAmountTwoDecimals(J.total)}`,
          detail: `${J.rows.length} row(s) as of ${n.asOf}, based on ${String(p.values[U])}. ${m[U] !== "due date" ? "Age since transaction date; this is not a days-past-due assessment." : "Days relative to due date."}`,
          severity: X === "Over 90 days" && J.rows.length ? "warning" : "info",
          citations: J.rows.flatMap((row) => [g(row.addresses[U]), ...O(row)]),
        }),
      ),
      {
        title: "Aging analysis",
        summary: `${B}${N} As of ${n.asOf}; ${Q} additional row(s) excluded for invalid or missing dates. Signed amounts are used; for receivables aging select an open-item population and outstanding amounts.`,
        findings: findings,
      }
    );
  }
  throw new Error("Unsupported audit analysis.");
}
function recordCrossWorkbookImport(activeWorkbook, r, actor, n) {
  if (activeWorkbook.id === r.id)
    throw Error("Use the ordinary transaction path for the same workbook.");
  const workbookInput = normalizeWorkbookInput(structuredClone(r)),
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
          action: `${n === "undo" ? "Undo" : "Redo"} file import; previous ID [${activeWorkbook.id}] rev ${activeWorkbook.revision}; source rev ${r.revision}`,
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
function handleDialogKeyDown(e, onClose) {
  if (!e.defaultPrevented) {
    if (e.key === "Escape")
      (e.preventDefault(), e.stopPropagation(), onClose());
    else if (e.key === "Tab") {
      const t = Array.from(
          e.currentTarget.querySelectorAll(
            'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
          ),
        ).filter((a) => !a.closest("[hidden]")),
        n = t[0],
        o = t.at(-1);
      e.shiftKey && document.activeElement === n
        ? (e.preventDefault(), o == null || o.focus())
        : !e.shiftKey &&
          document.activeElement === o &&
          (e.preventDefault(), n == null || n.focus());
    }
  }
}
function CommandReferenceDialog({ commands, onClose, onLocate }) {
  const [n, o] = ne.useState(""),
    a = ne.useRef(null),
    s = ne.useId();
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      a.current?.focus(),
      () => {
        activeElement?.isConnected &&
          activeElement.focus({ preventScroll: true });
      }
    );
  }, []);
  const d = (c) => c.toLowerCase().includes(n.trim().toLowerCase()),
    u = commands.filter((command) =>
      d(`${command.name} ${command.location} ${command.reason ?? ""}`),
    ),
    l = keyboardShortcutGroups.filter((keyboardShortcutGroup) =>
      d(keyboardShortcutGroup.join(" ")),
    );
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      className: "lg-command-reference",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": s,
      onKeyDown: (c) => handleDialogKeyDown(c, onClose),
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: s,
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
              ref: a,
              value: n,
              onChange: (event) => o(event.target.value),
            }),
          ],
        }),
        i.jsxs("p", {
          role: "status",
          children: [
            u.length,
            " ",
            i.jsx(LocalizedText, { children: "commands and " }),
            l.length,
            " ",
            i.jsx(LocalizedText, { children: "shortcuts" }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-command-results",
          children: [
            l.length > 0 &&
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
                        children: l.map(([c, h, f]) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("th", { children: c }),
                                i.jsx("td", { children: h }),
                                i.jsx("td", { children: f }),
                              ],
                            },
                            c,
                          ),
                        ),
                      }),
                    ],
                  }),
                ],
              }),
            u.length > 0 &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, { children: "Commands" }),
                  }),
                  i.jsx("ul", {
                    children: u.map((c) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            i.jsxs("div", {
                              children: [
                                i.jsx("strong", { children: c.name }),
                                i.jsx("span", { children: c.location }),
                                c.reason &&
                                  i.jsxs("p", {
                                    children: [
                                      i.jsx(LocalizedText, {
                                        children: "Unavailable: ",
                                      }),
                                      c.reason,
                                    ],
                                  }),
                              ],
                            }),
                            !c.reason &&
                              i.jsx("button", {
                                type: "button",
                                "aria-label": `Show ${c.name} in ${c.location}`,
                                onClick: () => onLocate(c),
                                children: i.jsx(LocalizedText, {
                                  children: "Show command",
                                }),
                              }),
                          ],
                        },
                        c.id,
                      ),
                    ),
                  }),
                ],
              }),
            !u.length &&
              !l.length &&
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
function extractElementText(e) {
  return ne.Children.toArray(e)
    .map((r) =>
      typeof r == "string" || typeof r == "number"
        ? String(r)
        : ne.isValidElement(r)
          ? extractElementText(r.props.children)
          : "",
    )
    .filter(Boolean)
    .join(" ")
    .trim();
}
function collectRibbonCommands(e, r, t, n) {
  const commands = [];
  function a(s, d = r, u) {
    return ne.Children.map(s, (l) => {
      if (!ne.isValidElement(l)) return l;
      const props = l.props;
      if (props["data-workbook-command"]) return l;
      let h = "",
        disabled = props.disabled,
        p = {};
      if (
        (l.type === t.group
          ? ((d = `${r} → ${props.label}`),
            props.onDialog &&
              ((h = props.dialogLabel ?? `${props.label} settings`),
              (disabled = props.dialogDisabled)))
          : l.type === t.command || l.type === t.icon
            ? (h = props.label ?? "")
            : l.type === t.color
              ? (h = props.kind === "fill" ? "Fill Color" : "Font Color")
              : ["button", "select", "input"].includes(l.type) &&
                props.role !== "tab" &&
                (h =
                  props["aria-label"] ??
                  props.title ??
                  u ??
                  extractElementText(props.children)),
        h)
      ) {
        h = h.replace(/ \(Ctrl\+[A-Z]\)$/, "");
        const commandId = `${r}:${commands.length}`,
          reason = disabled ? n(h) : undefined;
        (commands.push({ id: commandId, name: h, location: d, reason: reason }),
          l.type === t.group
            ? (p = { commandId: commandId, reason: reason })
            : typeof l.type == "string"
              ? (p = {
                  "data-workbook-command": commandId,
                  "aria-label": props["aria-label"] ?? h,
                  "aria-description": reason,
                  ...(reason ? { title: `${h}: ${reason}` } : {}),
                })
              : (p = { commandId: commandId, reason: reason }));
      }
      return ne.cloneElement(
        l,
        p,
        props.children === undefined
          ? undefined
          : a(
              props.children,
              d,
              l.type === "label" ? extractElementText(props.children) : u,
            ),
      );
    });
  }
  return { content: a(e), commands: commands };
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
  resolveRibbonCommandLabel = (e) =>
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
    })[e] ?? e.replace(/ \(Ctrl\+[A-Z]\)$/, "");
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
  buildSampleLedgerWorkbook,
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
