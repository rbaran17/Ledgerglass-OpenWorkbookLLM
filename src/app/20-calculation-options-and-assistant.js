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

function CalculationOptionsDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    [, a] = ne.useState(0),
    { report } = getWorkbookCalculation(e.workbook),
    [d, u] = ne.useState(report.settings),
    [l, c] = ne.useState(""),
    [h, f] = ne.useState(""),
    [p, m] = ne.useState(0),
    [v, g] = ne.useState(() => normalizeDatePolicy(e.workbook.datePolicy));
  (ne.useEffect(() => {
    (u(getWorkbookCalculation(e.workbook).report.settings),
      g(normalizeDatePolicy(e.workbook.datePolicy)),
      m(0));
  }, [e.workbook]),
    ne.useEffect(() => {
      const j = document.activeElement;
      return (
        o.current
          ?.querySelector('[data-ui-label="Close Calculation Options"]')
          ?.focus(),
        () => {
          j?.isConnected && j.focus();
        }
      );
    }, []));
  const b = (j) => {
      try {
        (j(), c(""));
      } catch (z) {
        c(z instanceof Error ? z.message : "Calculation action failed.");
      }
    },
    y = () =>
      b(() => {
        if (e.onCalculate) {
          e.onCalculate();
          return;
        }
        (refreshCalculationReport(e.workbook),
          a((z) => z + 1),
          e.onCalculated?.());
      }),
    w = () =>
      b(() => {
        if (e.readOnly)
          throw Error(
            "This host is read-only. Calculation settings cannot be changed.",
          );
        e.onWorkbookChange(
          {
            ...withCalculationSettings(e.workbook, d),
            datePolicy: normalizeDatePolicy(v),
          },
          `Set calculation to ${d.mode}; iteration ${d.iterative ? "enabled" : "disabled"}; dates ${v.dateSystem} ${v.textDateOrder}, time zone ${v.timeZone}`,
        );
      }),
    E = () =>
      b(() => {
        const j = buildCalculationEvidence(e.workbook, report);
        if (e.onExport) {
          e.onExport(j);
          return;
        }
        const z = URL.createObjectURL(
            new Blob([j], { type: "application/json" }),
          ),
          V = document.createElement("a");
        ((V.href = z),
          (V.download = "Calculation-evidence.json"),
          V.click(),
          setTimeout(() => URL.revokeObjectURL(z), 1e3));
      }),
    A = report.cells.filter((j) =>
      `${e.workbook.sheets.find((V) => V.id === j.sheetId)?.name}!${j.address} ${j.formula} ${j.status} ${j.result.value}`
        .toLowerCase()
        .includes(h.toLowerCase()),
    ),
    k = A.slice(p * 100, (p + 1) * 100),
    F = (j) => {
      var V, D;
      if (
        (j.key === "Escape" && (j.stopPropagation(), e.onClose()),
        j.key !== "Tab")
      )
        return;
      const z = [
        ...o.current.querySelectorAll(
          'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
        ),
      ];
      j.shiftKey && document.activeElement === z[0]
        ? (j.preventDefault(), (V = z.at(-1)) == null || V.focus())
        : !j.shiftKey &&
          document.activeElement === z.at(-1) &&
          (j.preventDefault(), (D = z[0]) == null || D.focus());
    };
  return i.jsx("div", {
    className: "fw-backdrop",
    children: i.jsxs("section", {
      ref: o,
      className: "fw-dialog rc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: F,
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
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
              onClick: e.onClose,
              lang: language("Close Calculation Options"),
              "data-ui-label": "Close Calculation Options",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("fieldset", {
          disabled: e.readOnly,
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
                  value: d.mode,
                  onChange: (j) => u({ ...d, mode: j.target.value }),
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
                  checked: d.iterative,
                  onChange: (j) => u({ ...d, iterative: j.target.checked }),
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
                  value: d.maxIterations,
                  onValueChange: (j) => u({ ...d, maxIterations: Number(j) }),
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
                  value: d.maxChange,
                  onValueChange: (j) => u({ ...d, maxChange: Number(j) }),
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
                  value: d.timeBudgetMs,
                  onValueChange: (j) => u({ ...d, timeBudgetMs: Number(j) }),
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
                    d.memoryBudgetMiB ??
                    calculationBudgetDefaults.defaultMemoryBudgetMiB,
                  onValueChange: (j) => u({ ...d, memoryBudgetMiB: Number(j) }),
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
                    d.workBudget ?? calculationBudgetDefaults.defaultWorkBudget,
                  onValueChange: (j) => u({ ...d, workBudget: Number(j) }),
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Workbook date system" }),
                i.jsxs("select", {
                  "aria-label": text("Workbook date system"),
                  value: v.dateSystem,
                  onChange: (j) => g({ ...v, dateSystem: j.target.value }),
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
            v.dateSystem !== report.dateContext.dateSystem &&
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
                  value: v.textDateOrder,
                  onChange: (j) => g({ ...v, textDateOrder: j.target.value }),
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
                  value: v.timeZone,
                  onChange: (j) => g({ ...v, timeZone: j.target.value }),
                  placeholder: text("UTC or America/New_York"),
                  lang: language("Calculation time zone"),
                  "data-ui-label": "Calculation time zone",
                }),
              ],
            }),
            i.jsx("button", {
              onClick: w,
              children: i.jsx(LocalizedText, {
                children: "Apply calculation settings",
              }),
            }),
          ],
        }),
        e.readOnly &&
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
              onClick: y,
              children: i.jsx(LocalizedText, { children: "Calculate Now" }),
            }),
            report.execution?.mode === "worker" &&
              report.execution.state === "pending" &&
              e.onCancelCalculation &&
              i.jsx("button", {
                onClick: e.onCancelCalculation,
                children: i.jsx(LocalizedText, {
                  children: "Cancel calculation",
                }),
              }),
            i.jsx("button", {
              onClick: E,
              children: i.jsx(LocalizedText, {
                children: "Download calculation evidence",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Find formula" }),
                i.jsx("input", {
                  "aria-label": text("Find formula results"),
                  value: h,
                  onChange: (j) => {
                    (f(j.target.value), m(0));
                  },
                  lang: language("Find formula results"),
                  "data-ui-label": "Find formula results",
                }),
              ],
            }),
          ],
        }),
        l && i.jsx("p", { role: "alert", children: l }),
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
                .map((j) =>
                  j
                    .map((z) => {
                      const V = z.lastIndexOf(":");
                      return `${e.workbook.sheets.find((B) => B.id === z.slice(0, V))?.name ?? "(removed sheet)"}!${z.slice(V + 1)}`;
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
                children: k.map((j) =>
                  i.jsxs(
                    "tr",
                    {
                      children: [
                        i.jsx("td", {
                          children: i.jsxs("button", {
                            onClick: () => {
                              (e.onNavigate?.(j.sheetId, j.address),
                                e.onClose());
                            },
                            children: [
                              e.workbook.sheets.find((D) => D.id === j.sheetId)
                                ?.name,
                              "!",
                              j.address,
                            ],
                          }),
                        }),
                        i.jsx("td", { children: j.formula }),
                        i.jsxs("td", {
                          children: [
                            String(j.result.value ?? "(blank)"),
                            " (",
                            j.result.type,
                            ")",
                          ],
                        }),
                        i.jsxs("td", {
                          children: [
                            j.status,
                            " / ",
                            j.evaluations,
                            j.financial?.map((D, B) =>
                              i.jsxs(
                                "details",
                                {
                                  children: [
                                    i.jsxs("summary", {
                                      children: [
                                        D.functionName,
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children: "financial details — ",
                                        }),
                                        D.status,
                                      ],
                                    }),
                                    i.jsx("p", { children: D.message }),
                                    i.jsx("p", { children: D.convention }),
                                    D.solver &&
                                      i.jsxs(i.Fragment, {
                                        children: [
                                          i.jsxs("p", {
                                            children: [
                                              i.jsx(LocalizedText, {
                                                children: "Guess ",
                                              }),
                                              D.guess,
                                              "; ",
                                              D.solver.evaluations,
                                              " ",
                                              i.jsx(LocalizedText, {
                                                children:
                                                  "equation evaluations. Bounded search: ",
                                              }),
                                              D.solver.status,
                                              i.jsx(LocalizedText, {
                                                children: ". Candidate roots:",
                                              }),
                                              " ",
                                              D.solver.candidates
                                                .map((_) =>
                                                  _.rate.toPrecision(12),
                                                )
                                                .join(", ") || "None",
                                              ".",
                                            ],
                                          }),
                                          D.solver.selected &&
                                            i.jsxs("p", {
                                              children: [
                                                i.jsx(LocalizedText, {
                                                  children: "Selected rate",
                                                }),
                                                " ",
                                                D.solver.selected.rate.toPrecision(
                                                  12,
                                                ),
                                                i.jsx(LocalizedText, {
                                                  children:
                                                    "; equation residual",
                                                }),
                                                " ",
                                                D.solver.selected.residual.toPrecision(
                                                  6,
                                                ),
                                                i.jsx(LocalizedText, {
                                                  children:
                                                    "; relative residual",
                                                }),
                                                " ",
                                                D.solver.selected.relativeResidual.toExponential(
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
                                `${D.functionName}-${B}`,
                              ),
                            ),
                          ],
                        }),
                        i.jsx("td", {
                          children: j.lastKnown
                            ? `${String(j.lastKnown.result.value)} at revision ${j.lastKnown.revision} — historical only`
                            : "None retained",
                        }),
                      ],
                    },
                    `${j.sheetId}:${j.address}`,
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
              disabled: p === 0,
              onClick: () => m((j) => j - 1),
              children: i.jsx(LocalizedText, { children: "Previous results" }),
            }),
            i.jsxs("span", {
              children: [
                A.length ? p * 100 + 1 : 0,
                "–",
                Math.min((p + 1) * 100, A.length),
                " ",
                i.jsx(LocalizedText, { children: "of " }),
                A.length,
              ],
            }),
            i.jsx("button", {
              disabled: (p + 1) * 100 >= A.length,
              onClick: () => m((j) => j + 1),
              children: i.jsx(LocalizedText, { children: "Next results" }),
            }),
          ],
        }),
        e.workbook.importedCalculationCaches &&
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(LocalizedText, {
                    children: "Imported cache evidence —",
                  }),
                  " ",
                  e.workbook.importedCalculationCaches.entries.length,
                  " ",
                  i.jsx(LocalizedText, { children: "source formulas" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  e.workbook.importedCalculationCaches.fileName,
                  " ",
                  i.jsx(LocalizedText, { children: "· SHA-256" }),
                  " ",
                  e.workbook.importedCalculationCaches.sourceSha256,
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Imported " }),
                  e.workbook.importedCalculationCaches.importedAt,
                  i.jsx(LocalizedText, { children: "; source mode " }),
                  e.workbook.importedCalculationCaches.sourceSettings.mode,
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
  x,
  y,
  label,
  actions,
  workbook,
  selection,
  onClose,
}) {
  const d = ne.useRef(null),
    u = ne.useRef({ workbook, snapshot: JSON.stringify(workbook), selection }),
    [l, c] = ne.useState({ left: x, top: y }),
    f =
      u.current.workbook !== workbook ||
      u.current.snapshot !== JSON.stringify(workbook) ||
      u.current.selection !== selection
        ? "The workbook or selection changed. Close and reopen this menu."
        : undefined;
  return (
    ne.useLayoutEffect(() => {
      const p = d.current?.getBoundingClientRect();
      p &&
        c({
          left: Math.max(8, Math.min(x, window.innerWidth - p.width - 8)),
          top: Math.max(8, Math.min(y, window.innerHeight - p.height - 8)),
        });
    }, [x, y]),
    ne.useEffect(() => {
      let p = false;
      const m = requestAnimationFrame(() =>
          d.current?.querySelector("button")?.focus(),
        ),
        v = (g) => {
          d.current?.contains(g.target)
            ? (p = true)
            : (g.type !== "focusin" || p) && onClose(false);
        };
      return (
        document.addEventListener("pointerdown", v),
        document.addEventListener("focusin", v),
        () => {
          (cancelAnimationFrame(m),
            document.removeEventListener("pointerdown", v),
            document.removeEventListener("focusin", v));
        }
      );
    }, [onClose]),
    i.jsxs("div", {
      ref: d,
      className: "lg-grid-context",
      role: "menu",
      "aria-label": `${label} commands`,
      style: l,
      onContextMenu: (p) => p.preventDefault(),
      onKeyDown: (p) => {
        if (p.defaultPrevented || p.ctrlKey || p.metaKey || p.altKey) return;
        const m = Array.from(
            p.currentTarget.querySelectorAll('[role="menuitem"]'),
          ),
          v = m.indexOf(document.activeElement);
        if (["ArrowDown", "ArrowUp", "Home", "End"].includes(p.key)) {
          (p.preventDefault(), p.stopPropagation());
          const b =
            p.key === "Home"
              ? 0
              : p.key === "End"
                ? m.length - 1
                : (v + (p.key === "ArrowDown" ? 1 : m.length - 1)) % m.length;
          m[b]?.focus();
        } else
          (p.key === "Escape" || p.key === "Tab") &&
            (p.preventDefault(), p.stopPropagation(), onClose(true));
      },
      children: [
        i.jsx("strong", { children: label }),
        actions.map((p, m) => {
          const v = f ?? p.reason;
          return i.jsxs(
            "button",
            {
              type: "button",
              role: "menuitem",
              tabIndex: m === 0 ? 0 : -1,
              "aria-label": p.name,
              "aria-disabled": !!v,
              "aria-description": v,
              title: v ? `${p.name}: ${v}` : p.name,
              onClick: () => {
                v || (onClose(true), p.run());
              },
              children: [p.name, v && i.jsx("small", { children: v })],
            },
            p.name,
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
        children: workbookLimitTable.map(([t, n]) =>
          i.jsxs(
            "div",
            {
              children: [
                i.jsx("dt", { children: t }),
                i.jsx("dd", { children: n }),
              ],
            },
            t,
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
  .map((e) => `.${e}`)
  .join(",");

function readAttachmentBytes(e, r) {
  return new Promise((t, n) => {
    const o = new FileReader();
    let a = false;
    const s = (l, c) => {
        a ||
          ((a = true),
          clearTimeout(u),
          r?.removeEventListener("abort", d),
          c ? n(c) : t(new Uint8Array(l)));
      },
      d = () => {
        (s(
          undefined,
          new DOMException("File preparation cancelled.", "AbortError"),
        ),
          o.abort());
      },
      u = setTimeout(() => {
        (s(
          undefined,
          new Error("Reading the attachment timed out. Try a smaller file."),
        ),
          o.abort());
      }, 15e3);
    ((o.onload = () => s(o.result)),
      (o.onerror = () =>
        s(
          undefined,
          new Error("The attachment could not be read. Select it again."),
        )),
      (o.onabort = () =>
        s(
          undefined,
          new DOMException("File preparation cancelled.", "AbortError"),
        )),
      r?.addEventListener("abort", d, { once: true }),
      r?.aborted ? d() : o.readAsArrayBuffer(e));
  });
}
async function prepareChatAttachment(e, r) {
  assertCondition(
    isSafeAttachmentFilename(e.name),
    "Use a filename of 1–180 characters without path or control characters.",
  );
  const t = getFileExtension(e.name);
  (assertCondition(
    Object.hasOwn(attachmentMediaTypes, t),
    "Supported attachments: TXT, Markdown, CSV, TSV, JSON, PDF, DOCX, XLSX, PNG, JPEG and WebP.",
  ),
    assertCondition(
      e.size > 0 && e.size <= assistantRequestLimits.fileBytes,
      "Each attachment must be nonempty and no larger than 5 MB.",
    ));
  const n = await readAttachmentBytes(e, r);
  assertCondition(
    n.length === e.size && matchesFileSignature(n, t),
    "The file content does not match its extension. Choose the original supported file.",
  );
  let o;
  const a = textAttachmentExtensions.has(t) ? "utf8" : "base64";
  if (a === "utf8") {
    try {
      o = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(n);
    } catch {
      throw new Error(
        "Text attachments must use UTF-8 encoding. Save the file as UTF-8 and try again.",
      );
    }
    assertCondition(
      !o.includes("\0"),
      "This appears to be a binary file. Choose a supported document or UTF-8 text file.",
    );
  } else {
    let s = "";
    for (let d = 0; d < n.length; d += 8192)
      s += String.fromCharCode(...n.subarray(d, d + 8192));
    o = btoa(s);
  }
  return {
    id: crypto.randomUUID(),
    name: e.name,
    mediaType: attachmentMediaTypes[t],
    size: n.length,
    encoding: a,
    data: o,
  };
}

function isAbortOrTimeoutError(e) {
  return e instanceof LLMRequestError
    ? new LLMRequestError(e.code)
    : e instanceof Error && e.name === "TimeoutError"
      ? new LLMRequestError("timeout")
      : e instanceof Error && e.name === "AbortError"
        ? new LLMRequestError("cancelled")
        : new LLMRequestError("unavailable");
}
async function raceWithAbortSignal(e, r, t = 3e4) {
  if (!Number.isFinite(t) || t < 1 || t > 12e4)
    throw new LLMRequestError("invalid_request");
  if (r?.aborted) throw new LLMRequestError("cancelled");
  const n = new AbortController();
  let o;
  const a = new Promise((l, c) => {
      o = c;
    }),
    s = (l) => {
      const c = new LLMRequestError(l);
      (o(c), n.abort(c));
    },
    d = () => s("cancelled");
  r?.addEventListener("abort", d, { once: true });
  const u = setTimeout(() => s("timeout"), t);
  try {
    return await Promise.race([a, e(n.signal)]);
  } catch (l) {
    throw isAbortOrTimeoutError(l);
  } finally {
    (clearTimeout(u), r?.removeEventListener("abort", d));
  }
}

function validateCalculationClock(e) {
  try {
    (assertCondition(isPlainRecord(e), "context"),
      assertNoUnexpectedFields(e, [
        "workbookId",
        "revision",
        "sheetId",
        "sheetName",
        "range",
        "cells",
        "datePolicy",
        "calculationClock",
      ]));
    const r = normalizeDatePolicy(e.datePolicy);
    if (e.calculationClock !== undefined) {
      const d = e.calculationClock;
      (assertCondition(isPlainRecord(d), "clock"),
        assertCondition(
          typeof d.calculatedAt == "string" &&
            /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(d.calculatedAt),
          "clock timestamp",
        ));
      const u = resolveLocalDateTime(r, new Date(d.calculatedAt));
      (assertNoUnexpectedFields(d, Object.keys(u)),
        assertCondition(
          Object.entries(u).every(([l, c]) => d[l] === c),
          "clock policy/value mismatch",
        ));
    }
    for (const d of ["workbookId", "sheetId", "sheetName"])
      assertCondition(
        typeof e[d] == "string" && e[d].length > 0 && e[d].length <= 200,
        "identity",
      );
    (assertCondition(
      Number.isSafeInteger(e.revision) && e.revision >= 0,
      "revision",
    ),
      assertCondition(
        typeof e.range == "string" && e.range.length <= 64,
        "range",
      ));
    const t = e.range.split(":");
    assertCondition(t.length <= 2, "range");
    const n = parseCellAddress(t[0]),
      o = parseCellAddress(t[1] ?? t[0]),
      a = (o.row - n.row + 1) * (o.col - n.col + 1);
    (assertCondition(
      o.row >= n.row && o.col >= n.col && a <= workbookLimits.maxLLMCells,
      "bounds",
    ),
      assertCondition(
        Array.isArray(e.cells) && e.cells.length === a,
        "population",
      ));
    const s = new Set();
    for (const d of e.cells) {
      (assertCondition(isPlainRecord(d), "cell"),
        assertNoUnexpectedFields(d, ["address", "raw", "value"]),
        assertCondition(
          typeof d.address == "string" && !s.has(d.address),
          "address",
        ));
      const u = parseCellAddress(d.address);
      (assertCondition(
        u.row >= n.row && u.row <= o.row && u.col >= n.col && u.col <= o.col,
        "scope",
      ),
        s.add(d.address));
      for (const l of [d.raw, d.value])
        assertCondition(
          l === null ||
            typeof l == "boolean" ||
            (typeof l == "number" && Number.isFinite(l)) ||
            (typeof l == "string" && l.length <= 32767),
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
