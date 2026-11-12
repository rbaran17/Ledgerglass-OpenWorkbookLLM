import {
  LLMRequestError,
  answerGroundingLimits,
  assertCondition,
  assertNoUnexpectedFields,
  assistantRequestLimits,
  attachmentMediaTypes,
  buildAssistantSelectionContext,
  buildLocalDemoProposal,
  getFileExtension,
  hasOnlyAllowedKeys,
  isPlainRecord,
  isSafeAttachmentFilename,
  matchesFileSignature,
  parseAssistantChatRequest,
  parseCellAddress,
  textAttachmentExtensions,
  validateAssistantFormula,
  validateAssistantProposal,
  validateCitationRange,
  validateProposedCalculation,
} from "./19b-assistant-protocol.js";

import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { string, array, discriminatedUnion, literal, object } from "zod";
import {
  walkFormulaNodes,
  calculationBudgetDefaults,
  tokenizeFormulaText,
  workbookUsesNowOrToday,
  normalizeBindingName,
  resolveLocalDateTime,
  normalizeDatePolicy,
  calculationFeatureSupport,
  defineDataPropertyByKey,
} from "./01-schemas-date-functions-and-validation.js";
import {
  workbookLimitTable,
  workbookLimits,
  supportedFontFamilies,
} from "./02-text-and-financial-functions.js";
import {
  evaluateWorkbookFormulas,
  builtinFunctionNames,
  getWorkbookCalculation,
  withCalculationSettings,
  calculateWorkbookWithSettings,
  buildCalculationEvidence,
  parseRangeAddress,
  refreshCalculationReport,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { numericPrecisionPolicy } from "./08-xlsx-import-export-and-preservation.js";

function CalculationOptionsDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    [, setRecalcVersion] = ne.useState(0),
    { report } = getWorkbookCalculation(props.workbook),
    [draftSettings, setDraftSettings] = ne.useState(report.settings),
    [actionError, setActionError] = ne.useState(""),
    [formulaFilter, setFormulaFilter] = ne.useState(""),
    [pageIndex, setPageIndex] = ne.useState(0),
    [datePolicy, setDatePolicy] = ne.useState(() =>
      normalizeDatePolicy(props.workbook.datePolicy),
    );
  (ne.useEffect(() => {
    (setDraftSettings(getWorkbookCalculation(props.workbook).report.settings),
      setDatePolicy(normalizeDatePolicy(props.workbook.datePolicy)),
      setPageIndex(0));
  }, [props.workbook]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current
          ?.querySelector('[data-ui-label="Close Calculation Options"]')
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const runAction = (action) => {
      try {
        (action(), setActionError(""));
      } catch (error) {
        setActionError(
          error instanceof Error ? error.message : "Calculation action failed.",
        );
      }
    },
    onClick = () =>
      runAction(() => {
        if (props.onCalculate) {
          props.onCalculate();
          return;
        }
        (refreshCalculationReport(props.workbook),
          setRecalcVersion((previousVersion) => previousVersion + 1),
          props.onCalculated?.());
      }),
    applySettings = () =>
      runAction(() => {
        if (props.readOnly)
          throw Error(
            "This host is read-only. Calculation settings cannot be changed.",
          );
        props.onWorkbookChange(
          {
            ...withCalculationSettings(props.workbook, draftSettings),
            datePolicy: normalizeDatePolicy(datePolicy),
          },
          `Set calculation to ${draftSettings.mode}; iteration ${draftSettings.iterative ? "enabled" : "disabled"}; dates ${datePolicy.dateSystem} ${datePolicy.textDateOrder}, time zone ${datePolicy.timeZone}`,
        );
      }),
    downloadEvidence = () =>
      runAction(() => {
        const calculationEvidence = buildCalculationEvidence(
          props.workbook,
          report,
        );
        if (props.onExport) {
          props.onExport(calculationEvidence);
          return;
        }
        const evidenceUrl = URL.createObjectURL(
            new Blob([calculationEvidence], { type: "application/json" }),
          ),
          downloadLink = document.createElement("a");
        ((downloadLink.href = evidenceUrl),
          (downloadLink.download = "Calculation-evidence.json"),
          downloadLink.click(),
          setTimeout(() => URL.revokeObjectURL(evidenceUrl), 1e3));
      }),
    matchingCells = report.cells.filter((cell) =>
      `${props.workbook.sheets.find((sheet) => sheet.id === cell.sheetId)?.name}!${cell.address} ${cell.formula} ${cell.status} ${cell.result.value}`
        .toLowerCase()
        .includes(formulaFilter.toLowerCase()),
    ),
    pageCells = matchingCells.slice(pageIndex * 100, (pageIndex + 1) * 100),
    onKeyDown = (event) => {
      var lastFocusable, firstFocusable;
      if (
        (event.key === "Escape" && (event.stopPropagation(), props.onClose()),
        event.key !== "Tab")
      )
        return;
      const focusableElements = [
        ...dialogRef.current.querySelectorAll(
          'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
        ),
      ];
      event.shiftKey && document.activeElement === focusableElements[0]
        ? (event.preventDefault(),
          (lastFocusable = focusableElements.at(-1)) == null ||
            lastFocusable.focus())
        : !event.shiftKey &&
          document.activeElement === focusableElements.at(-1) &&
          (event.preventDefault(),
          (firstFocusable = focusableElements[0]) == null ||
            firstFocusable.focus());
    };
  return i.jsx("div", {
    className: "fw-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "fw-dialog rc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${titleId}-title`,
      onKeyDown: onKeyDown,
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${titleId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Calculation Options",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Review current results, dependency changes and separately retained imported caches.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Calculation Options"),
              onClick: props.onClose,
              lang: language("Close Calculation Options"),
              "data-ui-label": "Close Calculation Options",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("fieldset", {
          disabled: props.readOnly,
          className: "rc-settings",
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, { children: "Workbook settings" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Calculation mode" }),
                i.jsxs("select", {
                  "aria-label": text("Calculation mode"),
                  value: draftSettings.mode,
                  onChange: (event) =>
                    setDraftSettings({
                      ...draftSettings,
                      mode: event.target.value,
                    }),
                  lang: language("Calculation mode"),
                  "data-ui-label": "Calculation mode",
                  children: [
                    i.jsx("option", {
                      value: "automatic",
                      children: "Automatic",
                    }),
                    i.jsx("option", { value: "manual", children: "Manual" }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: draftSettings.iterative,
                  onChange: (event) =>
                    setDraftSettings({
                      ...draftSettings,
                      iterative: event.target.checked,
                    }),
                }),
                i.jsx(LocalizedText, {
                  children: "Enable iterative calculation",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Maximum iterations" }),
                i.jsx(LocalizedNumberField, {
                  "aria-label": text("Maximum iterations"),
                  min: 1,
                  max: 1e3,
                  value: draftSettings.maxIterations,
                  onValueChange: (value) =>
                    setDraftSettings({
                      ...draftSettings,
                      maxIterations: Number(value),
                    }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Maximum change" }),
                i.jsx(LocalizedNumberField, {
                  "aria-label": text("Maximum change"),
                  min: 0,
                  max: 1,
                  step: "0.000001",
                  value: draftSettings.maxChange,
                  onValueChange: (value) =>
                    setDraftSettings({
                      ...draftSettings,
                      maxChange: Number(value),
                    }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Time budget (ms)" }),
                i.jsx(LocalizedNumberField, {
                  "aria-label": text("Time budget (ms)"),
                  min: 1,
                  max: 1e4,
                  value: draftSettings.timeBudgetMs,
                  onValueChange: (value) =>
                    setDraftSettings({
                      ...draftSettings,
                      timeBudgetMs: Number(value),
                    }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Working-data budget (MiB)" }),
                i.jsx(LocalizedNumberField, {
                  "aria-label": text("Working-data budget (MiB)"),
                  min: 1,
                  max: 128,
                  value:
                    draftSettings.memoryBudgetMiB ??
                    calculationBudgetDefaults.defaultMemoryBudgetMiB,
                  onValueChange: (value) =>
                    setDraftSettings({
                      ...draftSettings,
                      memoryBudgetMiB: Number(value),
                    }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Calculation work budget" }),
                i.jsx(LocalizedNumberField, {
                  "aria-label": text("Calculation work budget"),
                  min: 1e3,
                  max: 1e7,
                  step: 1e3,
                  value:
                    draftSettings.workBudget ??
                    calculationBudgetDefaults.defaultWorkBudget,
                  onValueChange: (value) =>
                    setDraftSettings({
                      ...draftSettings,
                      workBudget: Number(value),
                    }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Workbook date system" }),
                i.jsxs("select", {
                  "aria-label": text("Workbook date system"),
                  value: datePolicy.dateSystem,
                  onChange: (event) =>
                    setDatePolicy({
                      ...datePolicy,
                      dateSystem: event.target.value,
                    }),
                  lang: language("Workbook date system"),
                  "data-ui-label": "Workbook date system",
                  children: [
                    i.jsx("option", {
                      value: "1900",
                      children: "1900 (Windows default)",
                    }),
                    i.jsx("option", {
                      value: "1904",
                      children: "1904 (legacy Mac)",
                    }),
                  ],
                }),
              ],
            }),
            datePolicy.dateSystem !== report.dateContext.dateSystem &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Changing the date system keeps stored numbers and formulas. Existing date serials will be interpreted using the new system; ordinary corresponding dates differ by 1,462 days. The 1900 system includes the fictitious serial 60 as a special boundary. Review this change before applying. Undo restores the previous setting.",
                }),
              }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Text date order" }),
                i.jsxs("select", {
                  "aria-label": text("Text date order"),
                  value: datePolicy.textDateOrder,
                  onChange: (event) =>
                    setDatePolicy({
                      ...datePolicy,
                      textDateOrder: event.target.value,
                    }),
                  lang: language("Text date order"),
                  "data-ui-label": "Text date order",
                  children: [
                    i.jsx("option", {
                      value: "MDY",
                      children: "Month/day/year",
                    }),
                    i.jsx("option", {
                      value: "DMY",
                      children: "Day/month/year",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Calculation time zone" }),
                i.jsx("input", {
                  "aria-label": text("Calculation time zone"),
                  value: datePolicy.timeZone,
                  onChange: (event) =>
                    setDatePolicy({
                      ...datePolicy,
                      timeZone: event.target.value,
                    }),
                  placeholder: text("UTC or America/New_York"),
                  lang: language("Calculation time zone"),
                  "data-ui-label": "Calculation time zone",
                }),
              ],
            }),
            i.jsx("button", {
              onClick: applySettings,
              children: i.jsx(LocalizedText, {
                children: "Apply calculation settings",
              }),
            }),
          ],
        }),
        props.readOnly &&
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Read-only host: settings are locked. Calculate Now refreshes local results without editing source cells or history.",
            }),
          }),
        i.jsxs("div", {
          className: "fw-toolbar",
          children: [
            i.jsx("button", {
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Calculate Now" }),
            }),
            report.execution?.mode === "worker" &&
              report.execution.state === "pending" &&
              props.onCancelCalculation &&
              i.jsx("button", {
                onClick: props.onCancelCalculation,
                children: i.jsx(LocalizedText, {
                  children: "Cancel calculation",
                }),
              }),
            i.jsx("button", {
              onClick: downloadEvidence,
              children: i.jsx(LocalizedText, {
                children: "Download calculation evidence",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Find formula" }),
                i.jsx("input", {
                  "aria-label": text("Find formula results"),
                  value: formulaFilter,
                  onChange: (event) => {
                    (setFormulaFilter(event.target.value), setPageIndex(0));
                  },
                  lang: language("Find formula results"),
                  "data-ui-label": "Find formula results",
                }),
              ],
            }),
          ],
        }),
        actionError && i.jsx("p", { role: "alert", children: actionError }),
        i.jsxs("p", {
          role: "status",
          children: [
            i.jsx(LocalizedText, { children: "Revision " }),
            report.revision,
            " · ",
            report.settings.mode,
            " · ",
            report.formulaCount,
            " ",
            i.jsx(LocalizedText, { children: "formulas · " }),
            report.evaluated,
            " ",
            i.jsx(LocalizedText, { children: "evaluated ·" }),
            " ",
            report.reused,
            " ",
            i.jsx(LocalizedText, { children: "reused · " }),
            report.pending,
            " ",
            i.jsx(LocalizedText, { children: "pending ·" }),
            " ",
            report.dependencyEdges,
            " ",
            i.jsx(LocalizedText, { children: "dependency edges" }),
          ],
        }),
        i.jsxs("p", {
          children: [report.reason, ". ", report.iteration.message],
        }),
        i.jsxs("details", {
          className: "rc-date-evidence",
          open: true,
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Calculation limits and volatile formulas",
              }),
            }),
            i.jsxs("p", {
              role: "status",
              children: [
                report.execution?.message,
                " ",
                report.limits.hardInterrupt
                  ? "Worker deadline and cancellation are enabled."
                  : "Cooperative execution: this host does not interrupt an individual running function.",
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, { children: "Working data:" }),
                " ",
                (report.resources.accountedBytes / 1048576).toFixed(2),
                " /",
                " ",
                (report.resources.memoryBudgetBytes / 1048576).toFixed(0),
                " ",
                i.jsx(LocalizedText, { children: "MiB accounted. Work: " }),
                report.resources.workUnits.toLocaleString(),
                " /",
                " ",
                report.resources.workBudget.toLocaleString(),
                " ",
                i.jsx(LocalizedText, { children: "units. Time budget:" }),
                " ",
                report.settings.timeBudgetMs,
                " ",
                i.jsx(LocalizedText, { children: "ms." }),
              ],
            }),
            report.resources.reason &&
              i.jsxs("p", {
                role: "alert",
                children: [
                  report.resources.reason,
                  i.jsx(LocalizedText, {
                    children:
                      ". Committed inputs remain intact. Adjust the data or budget and choose Calculate Now to retry.",
                  }),
                ],
              }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "The memory setting bounds accounted input, intermediate and result data. It does not measure the browser's total memory; parser and runtime overhead are separate. Worker deadlines depend on the browser delivering timer events.",
              }),
            }),
            i.jsx("p", { children: calculationFeatureSupport.refresh }),
            i.jsx("p", { children: calculationFeatureSupport.random }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Volatile functions in this workbook:",
                }),
                " ",
                report.volatile.functions.join(", ") || "None",
                i.jsx(LocalizedText, {
                  children:
                    ". Unsupported volatile functions require a compatibility review.",
                }),
              ],
            }),
          ],
        }),
        i.jsxs("details", {
          className: "rc-date-evidence",
          open: true,
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Precision and monetary rounding",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, { children: "Numbers use " }),
                numericPrecisionPolicy.storage,
                ". ",
                numericPrecisionPolicy.display,
                " ",
                numericPrecisionPolicy.rounding,
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Binary arithmetic can produce 0.30000000000000004 for 0.1 + 0.2. General displays 0.3; ROUND(0.1+0.2,2) explicitly produces 0.3. Store identifiers longer than 15 digits as Text. Native JSON retains stored numbers; XLSX may expose Excel precision limits. SUM uses decimal accumulation of stored numbers, so it can differ from a chain of binary additions.",
              }),
            }),
          ],
        }),
        i.jsxs("details", {
          className: "rc-date-evidence",
          open: true,
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Date and business-calendar policy",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, { children: "Gregorian calendar · " }),
                report.dateContext.dateSystem,
                " ",
                i.jsx(LocalizedText, {
                  children:
                    "serial system · ISO YYYY-MM-DD within the supported date range accepted · slash dates use ",
                }),
                report.dateContext.textDateOrder,
                i.jsx(LocalizedText, {
                  children:
                    ". Other calendars are unsupported. XLSX retains the selected serial system; native JSON also retains text order and time zone. Holiday ranges and weekend masks are supplied in WORKDAY / NETWORKDAYS formulas.",
                }),
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, { children: "Calculation snapshot:" }),
                " ",
                i.jsx("time", {
                  dateTime: report.dateContext.calculatedAt,
                  children: report.dateContext.calculatedAt,
                }),
                " ",
                i.jsx(LocalizedText, { children: "· time zone " }),
                report.dateContext.timeZone,
                " ",
                i.jsx(LocalizedText, { children: "· UTC offset" }),
                " ",
                report.dateContext.utcOffsetMinutes,
                " ",
                i.jsx(LocalizedText, { children: "minutes · local clock" }),
                " ",
                report.dateContext.localDateTime,
                ".",
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "TODAY and NOW share this captured instant. Calculate Now refreshes it; there is no background clock. Pending formulas have no current time result. Calendar policy persists in native JSON; XLSX does not retain this time zone or text-date policy.",
              }),
            }),
          ],
        }),
        !!report.cycles.length &&
          i.jsxs("p", {
            role: "alert",
            children: [
              i.jsx(LocalizedText, { children: "Circular references:" }),
              " ",
              report.cycles
                .slice(0, 10)
                .map((cycle) =>
                  cycle
                    .map((cycleNode) => {
                      const separatorIndex = cycleNode.lastIndexOf(":");
                      return `${props.workbook.sheets.find((sheet) => sheet.id === cycleNode.slice(0, separatorIndex))?.name ?? "(removed sheet)"}!${cycleNode.slice(separatorIndex + 1)}`;
                    })
                    .join(" → "),
                )
                .join("; "),
              report.cycles.length > 10 ? " (more in evidence)" : "",
            ],
          }),
        !!report.pending &&
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Pending formulas display #PENDING!. Historical and imported values are never substituted as current results. Use Calculate Now before relying on or sharing them.",
            }),
          }),
        i.jsx("div", {
          className: "fw-table-wrap",
          children: i.jsxs("table", {
            children: [
              i.jsx("thead", {
                children: i.jsxs("tr", {
                  children: [
                    i.jsx("th", {
                      children: i.jsx(LocalizedText, { children: "Source" }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(LocalizedText, { children: "Formula" }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(LocalizedText, {
                        children: "Current result",
                      }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(LocalizedText, {
                        children: "Status / evaluations",
                      }),
                    }),
                    i.jsx("th", {
                      children: i.jsx(LocalizedText, {
                        children: "Historical result",
                      }),
                    }),
                  ],
                }),
              }),
              i.jsx("tbody", {
                children: pageCells.map((reportCell) =>
                  i.jsxs(
                    "tr",
                    {
                      children: [
                        i.jsx("td", {
                          children: i.jsxs("button", {
                            onClick: () => {
                              (props.onNavigate?.(
                                reportCell.sheetId,
                                reportCell.address,
                              ),
                                props.onClose());
                            },
                            children: [
                              props.workbook.sheets.find(
                                (sheet) => sheet.id === reportCell.sheetId,
                              )?.name,
                              "!",
                              reportCell.address,
                            ],
                          }),
                        }),
                        i.jsx("td", { children: reportCell.formula }),
                        i.jsxs("td", {
                          children: [
                            String(reportCell.result.value ?? "(blank)"),
                            " (",
                            reportCell.result.type,
                            ")",
                          ],
                        }),
                        i.jsxs("td", {
                          children: [
                            reportCell.status,
                            " / ",
                            reportCell.evaluations,
                            reportCell.financial?.map(
                              (financialDetail, detailIndex) =>
                                i.jsxs(
                                  "details",
                                  {
                                    children: [
                                      i.jsxs("summary", {
                                        children: [
                                          financialDetail.functionName,
                                          " ",
                                          i.jsx(LocalizedText, {
                                            children: "financial details — ",
                                          }),
                                          financialDetail.status,
                                        ],
                                      }),
                                      i.jsx("p", {
                                        children: financialDetail.message,
                                      }),
                                      i.jsx("p", {
                                        children: financialDetail.convention,
                                      }),
                                      financialDetail.solver &&
                                        i.jsxs(i.Fragment, {
                                          children: [
                                            i.jsxs("p", {
                                              children: [
                                                i.jsx(LocalizedText, {
                                                  children: "Guess ",
                                                }),
                                                financialDetail.guess,
                                                "; ",
                                                financialDetail.solver
                                                  .evaluations,
                                                " ",
                                                i.jsx(LocalizedText, {
                                                  children:
                                                    "equation evaluations. Bounded search: ",
                                                }),
                                                financialDetail.solver.status,
                                                i.jsx(LocalizedText, {
                                                  children:
                                                    ". Candidate roots:",
                                                }),
                                                " ",
                                                financialDetail.solver.candidates
                                                  .map((candidate) =>
                                                    candidate.rate.toPrecision(
                                                      12,
                                                    ),
                                                  )
                                                  .join(", ") || "None",
                                                ".",
                                              ],
                                            }),
                                            financialDetail.solver.selected &&
                                              i.jsxs("p", {
                                                children: [
                                                  i.jsx(LocalizedText, {
                                                    children: "Selected rate",
                                                  }),
                                                  " ",
                                                  financialDetail.solver.selected.rate.toPrecision(
                                                    12,
                                                  ),
                                                  i.jsx(LocalizedText, {
                                                    children:
                                                      "; equation residual",
                                                  }),
                                                  " ",
                                                  financialDetail.solver.selected.residual.toPrecision(
                                                    6,
                                                  ),
                                                  i.jsx(LocalizedText, {
                                                    children:
                                                      "; relative residual",
                                                  }),
                                                  " ",
                                                  financialDetail.solver.selected.relativeResidual.toExponential(
                                                    3,
                                                  ),
                                                  ".",
                                                ],
                                              }),
                                            i.jsx("p", {
                                              children: i.jsx(LocalizedText, {
                                                children:
                                                  "Root discovery is not exhaustive. Repeated calls of the same function in this cell retain the last diagnostic; use Evaluate Formula for the expression trace.",
                                              }),
                                            }),
                                          ],
                                        }),
                                    ],
                                  },
                                  `${financialDetail.functionName}-${detailIndex}`,
                                ),
                            ),
                          ],
                        }),
                        i.jsx("td", {
                          children: reportCell.lastKnown
                            ? `${String(reportCell.lastKnown.result.value)} at revision ${reportCell.lastKnown.revision} — historical only`
                            : "None retained",
                        }),
                      ],
                    },
                    `${reportCell.sheetId}:${reportCell.address}`,
                  ),
                ),
              }),
            ],
          }),
        }),
        i.jsxs("div", {
          className: "fw-toolbar",
          children: [
            i.jsx("button", {
              disabled: pageIndex === 0,
              onClick: () => setPageIndex((currentPage) => currentPage - 1),
              children: i.jsx(LocalizedText, { children: "Previous results" }),
            }),
            i.jsxs("span", {
              children: [
                matchingCells.length ? pageIndex * 100 + 1 : 0,
                "–",
                Math.min((pageIndex + 1) * 100, matchingCells.length),
                " ",
                i.jsx(LocalizedText, { children: "of " }),
                matchingCells.length,
              ],
            }),
            i.jsx("button", {
              disabled: (pageIndex + 1) * 100 >= matchingCells.length,
              onClick: () => setPageIndex((currentPage) => currentPage + 1),
              children: i.jsx(LocalizedText, { children: "Next results" }),
            }),
          ],
        }),
        props.workbook.importedCalculationCaches &&
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(LocalizedText, {
                    children: "Imported cache evidence —",
                  }),
                  " ",
                  props.workbook.importedCalculationCaches.entries.length,
                  " ",
                  i.jsx(LocalizedText, { children: "source formulas" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  props.workbook.importedCalculationCaches.fileName,
                  " ",
                  i.jsx(LocalizedText, { children: "· SHA-256" }),
                  " ",
                  props.workbook.importedCalculationCaches.sourceSha256,
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Imported " }),
                  props.workbook.importedCalculationCaches.importedAt,
                  i.jsx(LocalizedText, { children: "; source mode " }),
                  props.workbook.importedCalculationCaches.sourceSettings.mode,
                  i.jsx(LocalizedText, {
                    children:
                      ". Source caches and settings remain historical, unverified evidence. Download calculation evidence for every typed cache comparison, including missing caches and changed formulas. Original XLSX bytes are not retained by this field.",
                  }),
                ],
              }),
            ],
          }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Scalar dependency changes reuse unaffected results. Names, structure and arrays trigger a conservative recalculation. Iteration starts from zero and requires numeric convergence; it does not promise Excel-identical iteration order. Worker execution can be terminated; synchronous SDK and diagnostic calculations use cooperative checks. Formula inspection is a separate diagnostic calculation.",
          }),
        }),
      ],
    }),
  });
}
function GridContextMenu({
  x: left,
  y: menuTop,
  label,
  actions,
  workbook,
  selection,
  onClose,
}) {
  const menuRef = ne.useRef(null),
    openedContext = ne.useRef({
      workbook,
      snapshot: JSON.stringify(workbook),
      selection,
    }),
    [position, setPosition] = ne.useState({ left: left, top: menuTop }),
    staleReason =
      openedContext.current.workbook !== workbook ||
      openedContext.current.snapshot !== JSON.stringify(workbook) ||
      openedContext.current.selection !== selection
        ? "The workbook or selection changed. Close and reopen this menu."
        : undefined;
  return (
    ne.useLayoutEffect(() => {
      const menuRect = menuRef.current?.getBoundingClientRect();
      menuRect &&
        setPosition({
          left: Math.max(
            8,
            Math.min(left, window.innerWidth - menuRect.width - 8),
          ),
          top: Math.max(
            8,
            Math.min(menuTop, window.innerHeight - menuRect.height - 8),
          ),
        });
    }, [left, menuTop]),
    ne.useEffect(() => {
      let hasInsideInteraction = false;
      const focusFrameId = requestAnimationFrame(() =>
          menuRef.current?.querySelector("button")?.focus(),
        ),
        handleDocumentEvent = (event) => {
          menuRef.current?.contains(event.target)
            ? (hasInsideInteraction = true)
            : (event.type !== "focusin" || hasInsideInteraction) &&
              onClose(false);
        };
      return (
        document.addEventListener("pointerdown", handleDocumentEvent),
        document.addEventListener("focusin", handleDocumentEvent),
        () => {
          (cancelAnimationFrame(focusFrameId),
            document.removeEventListener("pointerdown", handleDocumentEvent),
            document.removeEventListener("focusin", handleDocumentEvent));
        }
      );
    }, [onClose]),
    i.jsxs("div", {
      ref: menuRef,
      className: "lg-grid-context",
      role: "menu",
      "aria-label": `${label} commands`,
      style: position,
      onContextMenu: (event) => event.preventDefault(),
      onKeyDown: (event) => {
        if (
          event.defaultPrevented ||
          event.ctrlKey ||
          event.metaKey ||
          event.altKey
        )
          return;
        const menuItems = Array.from(
            event.currentTarget.querySelectorAll('[role="menuitem"]'),
          ),
          activeIndex = menuItems.indexOf(document.activeElement);
        if (["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) {
          (event.preventDefault(), event.stopPropagation());
          const nextIndex =
            event.key === "Home"
              ? 0
              : event.key === "End"
                ? menuItems.length - 1
                : (activeIndex +
                    (event.key === "ArrowDown" ? 1 : menuItems.length - 1)) %
                  menuItems.length;
          menuItems[nextIndex]?.focus();
        } else
          (event.key === "Escape" || event.key === "Tab") &&
            (event.preventDefault(), event.stopPropagation(), onClose(true));
      },
      children: [
        i.jsx("strong", { children: label }),
        actions.map((action, actionIndex) => {
          const disabledReason = staleReason ?? action.reason;
          return i.jsxs(
            "button",
            {
              type: "button",
              role: "menuitem",
              tabIndex: actionIndex === 0 ? 0 : -1,
              "aria-label": action.name,
              "aria-disabled": !!disabledReason,
              "aria-description": disabledReason,
              title: disabledReason
                ? `${action.name}: ${disabledReason}`
                : action.name,
              onClick: () => {
                disabledReason || (onClose(true), action.run());
              },
              children: [
                action.name,
                disabledReason && i.jsx("small", { children: disabledReason }),
              ],
            },
            action.name,
          );
        }),
      ],
    })
  );
}
function WorkbookCapacityLimits() {
  const { text, language } = useWorkbookMessages();
  return i.jsxs("section", {
    "aria-label": text("Workbook capacity limits"),
    lang: language("Workbook capacity limits"),
    "data-ui-label": "Workbook capacity limits",
    children: [
      i.jsx("h4", { children: i.jsx(LocalizedText, { children: "Capacity" }) }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(LocalizedText, {
          children:
            "The worksheet address space matches Excel. Processing and file limits below are this application's limits.",
        }),
      }),
      i.jsx("dl", {
        className: "lg-capacity-list",
        children: workbookLimitTable.map(([limitLabel, limitValue]) =>
          i.jsxs(
            "div",
            {
              children: [
                i.jsx("dt", { children: limitLabel }),
                i.jsx("dd", { children: limitValue }),
              ],
            },
            limitLabel,
          ),
        ),
      }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(LocalizedText, {
          children:
            "Text units use UTF-16: some symbols, including many emoji, count as two. File limits count bytes. Excel allows 8,192 formula characters; this editor allows 4,096 text units.",
        }),
      }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(LocalizedText, {
          children:
            "Imports that exceed a limit are refused before replacing your workbook. Empty rows between distant cells are not filled in. Available memory and browser storage quota can impose lower limits.",
        }),
      }),
      i.jsx("p", {
        className: "lg-muted",
        children: i.jsx(LocalizedText, {
          children:
            "Uploaded files are parsed in isolated workers with a 30-second default deadline and at most two simultaneous requests. Cancel file processing stops a pending import. A blocked worker never falls back to parsing in the editor. Error messages omit source details; keep the original file for repairs. Hosts can provide a trusted worker and a deadline up to 60 seconds.",
        }),
      }),
    ],
  });
}
const attachmentAcceptAttribute = Object.keys(attachmentMediaTypes)
  .map((attachmentMediaType) => `.${attachmentMediaType}`)
  .join(",");

function readAttachmentBytes(file, signal) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    let isSettled = false;
    const settle = (buffer, error) => {
        isSettled ||
          ((isSettled = true),
          clearTimeout(timeoutId),
          signal?.removeEventListener("abort", handleAbort),
          error ? reject(error) : resolve(new Uint8Array(buffer)));
      },
      handleAbort = () => {
        (settle(
          undefined,
          new DOMException("File preparation cancelled.", "AbortError"),
        ),
          reader.abort());
      },
      timeoutId = setTimeout(() => {
        (settle(
          undefined,
          new Error("Reading the attachment timed out. Try a smaller file."),
        ),
          reader.abort());
      }, 15e3);
    ((reader.onload = () => settle(reader.result)),
      (reader.onerror = () =>
        settle(
          undefined,
          new Error("The attachment could not be read. Select it again."),
        )),
      (reader.onabort = () =>
        settle(
          undefined,
          new DOMException("File preparation cancelled.", "AbortError"),
        )),
      signal?.addEventListener("abort", handleAbort, { once: true }),
      signal?.aborted ? handleAbort() : reader.readAsArrayBuffer(file));
  });
}
async function prepareChatAttachment(file, signal) {
  assertCondition(
    isSafeAttachmentFilename(file.name),
    "Use a filename of 1–180 characters without path or control characters.",
  );
  const fileExtension = getFileExtension(file.name);
  (assertCondition(
    Object.hasOwn(attachmentMediaTypes, fileExtension),
    "Supported attachments: TXT, Markdown, CSV, TSV, JSON, PDF, DOCX, XLSX, PNG, JPEG and WebP.",
  ),
    assertCondition(
      file.size > 0 && file.size <= assistantRequestLimits.fileBytes,
      "Each attachment must be nonempty and no larger than 5 MB.",
    ));
  const attachmentBytes = await readAttachmentBytes(file, signal);
  assertCondition(
    attachmentBytes.length === file.size &&
      matchesFileSignature(attachmentBytes, fileExtension),
    "The file content does not match its extension. Choose the original supported file.",
  );
  let attachmentData;
  const encoding = textAttachmentExtensions.has(fileExtension)
    ? "utf8"
    : "base64";
  if (encoding === "utf8") {
    try {
      attachmentData = new TextDecoder("utf-8", {
        fatal: true,
        ignoreBOM: true,
      }).decode(attachmentBytes);
    } catch {
      throw new Error(
        "Text attachments must use UTF-8 encoding. Save the file as UTF-8 and try again.",
      );
    }
    assertCondition(
      !attachmentData.includes("\0"),
      "This appears to be a binary file. Choose a supported document or UTF-8 text file.",
    );
  } else {
    let binaryString = "";
    for (
      let byteOffset = 0;
      byteOffset < attachmentBytes.length;
      byteOffset += 8192
    )
      binaryString += String.fromCharCode(
        ...attachmentBytes.subarray(byteOffset, byteOffset + 8192),
      );
    attachmentData = btoa(binaryString);
  }
  return {
    id: crypto.randomUUID(),
    name: file.name,
    mediaType: attachmentMediaTypes[fileExtension],
    size: attachmentBytes.length,
    encoding: encoding,
    data: attachmentData,
  };
}

function isAbortOrTimeoutError(error) {
  return error instanceof LLMRequestError
    ? new LLMRequestError(error.code)
    : error instanceof Error && error.name === "TimeoutError"
      ? new LLMRequestError("timeout")
      : error instanceof Error && error.name === "AbortError"
        ? new LLMRequestError("cancelled")
        : new LLMRequestError("unavailable");
}
async function raceWithAbortSignal(operation, signal, timeoutMs = 3e4) {
  if (!Number.isFinite(timeoutMs) || timeoutMs < 1 || timeoutMs > 12e4)
    throw new LLMRequestError("invalid_request");
  if (signal?.aborted) throw new LLMRequestError("cancelled");
  const controller = new AbortController();
  let rejectRace;
  const abortPromise = new Promise((resolve, reject) => {
      rejectRace = reject;
    }),
    abortRace = (reasonCode) => {
      const abortError = new LLMRequestError(reasonCode);
      (rejectRace(abortError), controller.abort(abortError));
    },
    handleAbort = () => abortRace("cancelled");
  signal?.addEventListener("abort", handleAbort, { once: true });
  const timeoutId = setTimeout(() => abortRace("timeout"), timeoutMs);
  try {
    return await Promise.race([abortPromise, operation(controller.signal)]);
  } catch (error) {
    throw isAbortOrTimeoutError(error);
  } finally {
    (clearTimeout(timeoutId),
      signal?.removeEventListener("abort", handleAbort));
  }
}

function validateCalculationClock(assistantSelectionContext) {
  try {
    (assertCondition(isPlainRecord(assistantSelectionContext), "context"),
      assertNoUnexpectedFields(assistantSelectionContext, [
        "workbookId",
        "revision",
        "sheetId",
        "sheetName",
        "range",
        "cells",
        "datePolicy",
        "calculationClock",
      ]));
    const datePolicy = normalizeDatePolicy(
      assistantSelectionContext.datePolicy,
    );
    if (assistantSelectionContext.calculationClock !== undefined) {
      const calculationClock = assistantSelectionContext.calculationClock;
      (assertCondition(isPlainRecord(calculationClock), "clock"),
        assertCondition(
          typeof calculationClock.calculatedAt == "string" &&
            /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(
              calculationClock.calculatedAt,
            ),
          "clock timestamp",
        ));
      const localDateTime = resolveLocalDateTime(
        datePolicy,
        new Date(calculationClock.calculatedAt),
      );
      (assertNoUnexpectedFields(calculationClock, Object.keys(localDateTime)),
        assertCondition(
          Object.entries(localDateTime).every(
            ([fieldName, expectedValue]) =>
              calculationClock[fieldName] === expectedValue,
          ),
          "clock policy/value mismatch",
        ));
    }
    for (const identityField of ["workbookId", "sheetId", "sheetName"])
      assertCondition(
        typeof assistantSelectionContext[identityField] == "string" &&
          assistantSelectionContext[identityField].length > 0 &&
          assistantSelectionContext[identityField].length <= 200,
        "identity",
      );
    (assertCondition(
      Number.isSafeInteger(assistantSelectionContext.revision) &&
        assistantSelectionContext.revision >= 0,
      "revision",
    ),
      assertCondition(
        typeof assistantSelectionContext.range == "string" &&
          assistantSelectionContext.range.length <= 64,
        "range",
      ));
    const rangeParts = assistantSelectionContext.range.split(":");
    assertCondition(rangeParts.length <= 2, "range");
    const cellAddress = parseCellAddress(rangeParts[0]),
      endAddress = parseCellAddress(rangeParts[1] ?? rangeParts[0]),
      cellCount =
        (endAddress.row - cellAddress.row + 1) *
        (endAddress.col - cellAddress.col + 1);
    (assertCondition(
      endAddress.row >= cellAddress.row &&
        endAddress.col >= cellAddress.col &&
        cellCount <= workbookLimits.maxLLMCells,
      "bounds",
    ),
      assertCondition(
        Array.isArray(assistantSelectionContext.cells) &&
          assistantSelectionContext.cells.length === cellCount,
        "population",
      ));
    const seenAddresses = new Set();
    for (const cell of assistantSelectionContext.cells) {
      (assertCondition(isPlainRecord(cell), "cell"),
        assertNoUnexpectedFields(cell, ["address", "raw", "value"]),
        assertCondition(
          typeof cell.address == "string" && !seenAddresses.has(cell.address),
          "address",
        ));
      const parsedAddress = parseCellAddress(cell.address);
      (assertCondition(
        parsedAddress.row >= cellAddress.row &&
          parsedAddress.row <= endAddress.row &&
          parsedAddress.col >= cellAddress.col &&
          parsedAddress.col <= endAddress.col,
        "scope",
      ),
        seenAddresses.add(cell.address));
      for (const scalarValue of [cell.raw, cell.value])
        assertCondition(
          scalarValue === null ||
            typeof scalarValue == "boolean" ||
            (typeof scalarValue == "number" && Number.isFinite(scalarValue)) ||
            (typeof scalarValue == "string" && scalarValue.length <= 32767),
          "value",
        );
    }
  } catch {
    throw new LLMRequestError("invalid_request");
  }
}

export {
  parseAssistantChatRequest,
  GridContextMenu,
  validateAssistantFormula,
  validateAssistantProposal,
  validateCitationRange,
  validateProposedCalculation,
  WorkbookCapacityLimits,
  CalculationOptionsDialog,
  assistantRequestLimits,
  attachmentAcceptAttribute,
  LLMRequestError,
  prepareChatAttachment,
  raceWithAbortSignal,
  buildLocalDemoProposal,
  validateCalculationClock,
  answerGroundingLimits,
  buildAssistantSelectionContext,
};
export { assertCondition };
