import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { defineDataPropertyByKey } from "./01-schemas-date-functions-and-validation.js";
import { automationLimits, columnIndexToLetters } from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  localizeDecimalPoint,
} from "./04-cell-input-paste-and-fill.js";
import { importDelimitedFile } from "./07-file-import-and-worker-dispatch.js";
import {
  delimitedImportDefaults,
  exportWorksheetToCsv,
  getWorksheetUsedRange,
} from "./08-xlsx-import-export-and-preservation.js";
import { useOperationScope } from "./09-workbook-links.js";
import {
  buildWhatIfPreview,
  computeWhatIfResult,
  applyWhatIfScenario,
  saveWhatIfAnalysis,
  findSavedWhatIfAnalysis,
  buildWhatIfReportRows,
  isWhatIfAnalysisInScope,
  parseAssumptionNumbers,
} from "./39-report-outputs-and-what-if.js";

function WhatIfAnalysisDialog(e) {
  const { text, language, entryLocale } = useWorkbookMessages(),
    o = ne.useRef(entryLocale).current,
    a = o === "de-DE" || o === "fr-FR",
    s = (ge) =>
      ge.map((se) => localizeDecimalPoint(se, o)).join(a ? "; " : ", "),
    d = ne.useId(),
    u = ne.useRef(null),
    [l, c] = ne.useState(e.initialKind ?? "scenarios"),
    [h, f] = ne.useState("Sensitivity analysis"),
    [sheetId, m] = ne.useState(e.sheetId),
    [v, g] = ne.useState(""),
    [b, y] = ne.useState(e.address),
    [orientation, E] = ne.useState("column"),
    [A, k] = ne.useState(""),
    [F, I] = ne.useState(""),
    [O, j] = ne.useState([
      { name: "Downside", values: "" },
      { name: "Upside", values: "" },
    ]),
    [checked, V] = ne.useState(false),
    [D, B] = ne.useState(""),
    [_, C] = ne.useState("0"),
    [T, N] = ne.useState(null),
    [R, L] = ne.useState(""),
    $ = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && checked,
      allowedSheetIds: e.allowedSheetIds,
    },
    U = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    G = e.workbook.sheets.filter(
      (sheet) =>
        U.includes(sheet.id) &&
        ($.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    H = JSON.stringify([U, $.includeHidden]),
    Q = ne.useRef(H),
    X$1 =
      !!T &&
      T.snapshot === JSON.stringify(e.workbook) &&
      T.context.includeHidden === $.includeHidden &&
      JSON.stringify(T.context.allowedSheetIds) === JSON.stringify(U),
    J = (e.workbook.whatIfAnalyses ?? []).filter((ge) =>
      isWhatIfAnalysisInScope(e.workbook, ge, $),
    );
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      u.current?.querySelector('[data-ui-label="Analysis name"]')?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      Q.current !== H &&
        (N(null),
        L(""),
        B(""),
        g(""),
        y(""),
        k(""),
        I(""),
        j([{ name: "Scenario 1", values: "" }]),
        (Q.current = H));
    }, [H]));
  const q = (ge) => {
      (ge(), N(null), L(""));
    },
    W = (ge) => {
      try {
        (ge(), L(""));
      } catch (error) {
        L(
          error instanceof Error
            ? error.message
            : "What-if analysis could not complete.",
        );
      }
    },
    K = (ge) => {
      const se = ge.match(/'(?:[^']|'')*'![^,\n]+|[^,\n]+/g) ?? [];
      if (!se.length || /[,\n]\s*[,\n]|[,\n]\s*$/.test(ge))
        throw Error("Enter cell addresses separated by commas or new lines.");
      return se.map((we) => {
        const address = we.trim(),
          Ae = address.lastIndexOf("!");
        if (Ae < 0) return { sheetId: sheetId, address: address };
        const ye = address
            .slice(0, Ae)
            .replace(/^'(.*)'$/, "$1")
            .replaceAll("''", "'"),
          Te = G.find((Pe) => Pe.name.toLowerCase() === ye.toLowerCase());
        if (!Te)
          throw Error(
            "A referenced worksheet is missing or outside the current context.",
          );
        return { sheetId: Te.id, address: address.slice(Ae + 1) };
      });
    },
    onClick = () =>
      W(() => {
        N(null);
        const ge = { name: h, inputs: K(v), outputs: K(b) },
          request =
            l === "scenarios"
              ? {
                  ...ge,
                  kind: l,
                  scenarios: O.map((we) => ({
                    name: we.name,
                    values: parseAssumptionNumbers(we.values, o),
                  })),
                }
              : l === "one-variable"
                ? {
                    ...ge,
                    kind: l,
                    orientation: orientation,
                    values: parseAssumptionNumbers(A, o),
                  }
                : {
                    ...ge,
                    kind: l,
                    rowValues: parseAssumptionNumbers(A, o),
                    columnValues: parseAssumptionNumbers(F, o),
                  };
        (N(computeWhatIfResult(e.workbook, request, $)), C("0"));
      }),
    ae = (ge, se = false) =>
      W(() => {
        const we = J.find((ye) => ye.id === ge);
        if ((N(null), B(ge), !we)) return;
        if (
          we.assumptions.some(
            (assumption) =>
              !e.workbook.sheets.some(
                (sheet) => sheet.id === assumption.sheetId,
              ),
          )
        )
          throw Error(
            "A saved source worksheet was deleted. This definition cannot reconnect by name. Remove the definition or create a new analysis with current source cells; its historical report is retained.",
          );
        const request = we.request,
          Ae = (ye) =>
            ye
              .map(
                (Te) =>
                  `'${e.workbook.sheets.find((sheet) => sheet.id === Te.sheetId).name.replaceAll("'", "''")}'!${Te.address}`,
              )
              .join(", ");
        (f(request.name),
          g(Ae(request.inputs)),
          y(Ae(request.outputs)),
          c(se ? "scenarios" : request.kind),
          se
            ? j([
                {
                  name: "Restore saved baseline",
                  values: s(we.baseline.inputs),
                },
              ])
            : request.kind === "scenarios"
              ? j(
                  request.scenarios.map((scenario) => ({
                    name: scenario.name,
                    values: s(scenario.values),
                  })),
                )
              : request.kind === "one-variable"
                ? (k(s(request.values)), E(request.orientation))
                : (k(s(request.rowValues)), I(s(request.columnValues))));
      }),
    oe = () =>
      W(() => {
        T &&
          (e.onWorkbookChange(
            saveWhatIfAnalysis(e.workbook, T, $, D || undefined),
            `Save what-if report: ${T.request.name}`,
          ),
          N(null));
      }),
    de = () =>
      W(() => {
        if (!T) return;
        const ge = applyWhatIfScenario(e.workbook, T, Number(_), $);
        (ge !== e.workbook &&
          e.onWorkbookChange(ge, `Apply scenario: ${T.trials[Number(_)].name}`),
          N(null));
      }),
    he = () =>
      W(() => {
        if (!T) return;
        const ge = JSON.stringify(
          buildWhatIfPreview(e.workbook, T, $),
          null,
          2,
        );
        if (e.onExport) e.onExport(ge);
        else {
          const se = URL.createObjectURL(
              new Blob([ge], { type: "application/json" }),
            ),
            we = document.createElement("a");
          ((we.href = se),
            (we.download = "what-if-analysis.json"),
            we.click(),
            setTimeout(() => URL.revokeObjectURL(se), 1e3));
        }
      }),
    Se = X$1 && T ? buildWhatIfReportRows(T) : [],
    le = Se.findIndex((ge) => ge[0] === "Source assumptions at preview");
  return i.jsx("div", {
    className: "gs-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: u,
      className: "gs-dialog wi-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${d}-title`,
      onKeyDown: (event) => {
        var we, Ee;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const Ae = [
            ...(u.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
            ) ?? []),
          ].filter(
            (ye) =>
              !ye.hidden &&
              (!ye.closest("details:not([open])") || ye.tagName === "SUMMARY"),
          );
          event.shiftKey && document.activeElement === Ae[0]
            ? (event.preventDefault(), (we = Ae.at(-1)) == null || we.focus())
            : !event.shiftKey &&
              document.activeElement === Ae.at(-1) &&
              (event.preventDefault(), (Ee = Ae[0]) == null || Ee.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${d}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "What-if Analysis",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Compare assumptions for estimates and valuation sensitivity.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close What-if Analysis"),
              onClick: e.onClose,
              lang: language("Close What-if Analysis"),
              "data-ui-label": "Close What-if Analysis",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "gs-fields",
          children: [
            i.jsx("p", {
              role: "note",
              children: text(
                a
                  ? "Separate numeric assumptions with semicolons or new lines. Use a comma for decimals."
                  : "Separate numeric assumptions with commas, semicolons or new lines.",
              ),
            }),
            entryLocale !== o &&
              i.jsx("p", {
                role: "status",
                children: text(
                  "The current edit keeps its original numeric locale until you apply or cancel it.",
                ),
              }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Analysis name" }),
                i.jsx("input", {
                  "aria-label": text("Analysis name"),
                  value: h,
                  maxLength: 100,
                  onChange: (event) => q(() => f(event.target.value)),
                  lang: language("Analysis name"),
                  "data-ui-label": "Analysis name",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Analysis type" }),
                i.jsxs("select", {
                  "aria-label": text("Analysis type"),
                  value: l,
                  onChange: (event) =>
                    q(() => {
                      (c(event.target.value), B(""));
                    }),
                  lang: language("Analysis type"),
                  "data-ui-label": "Analysis type",
                  children: [
                    i.jsx("option", {
                      value: "scenarios",
                      children: "Named scenarios",
                    }),
                    i.jsx("option", {
                      value: "one-variable",
                      children: "One-variable data table",
                    }),
                    i.jsx("option", {
                      value: "two-variable",
                      children: "Two-variable data table",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Default worksheet" }),
                i.jsxs("select", {
                  "aria-label": text("Default worksheet"),
                  value: G.some((ge) => ge.id === sheetId) ? sheetId : "",
                  onChange: (event) => q(() => m(event.target.value)),
                  lang: language("Default worksheet"),
                  "data-ui-label": "Default worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    G.map((ge) =>
                      i.jsx(
                        "option",
                        { value: ge.id, children: ge.name },
                        ge.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Saved analyses" }),
                i.jsxs("select", {
                  "aria-label": text("Saved analyses"),
                  value: J.some((ge) => ge.id === D) ? D : "",
                  onChange: (event) => ae(event.target.value),
                  lang: language("Saved analyses"),
                  "data-ui-label": "Saved analyses",
                  children: [
                    i.jsx("option", { value: "", children: "New analysis" }),
                    J.map((ge) =>
                      i.jsx(
                        "option",
                        { value: ge.id, children: ge.request.name },
                        ge.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Changing cells" }),
                i.jsx("input", {
                  "aria-label": text("Changing cells"),
                  placeholder:
                    l === "two-variable"
                      ? "B2, B3 — across first, down second"
                      : "B2, B3",
                  value: v,
                  onChange: (event) => q(() => g(event.target.value)),
                  lang: language("Changing cells"),
                  "data-ui-label": "Changing cells",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Result cells" }),
                i.jsx("input", {
                  "aria-label": text("Result cells"),
                  placeholder: text("B5, B6"),
                  value: b,
                  onChange: (event) => q(() => y(event.target.value)),
                  lang: language("Result cells"),
                  "data-ui-label": "Result cells",
                }),
              ],
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Use commas between cells. Cross-sheet references may use 'Sheet name'!B2. Result cells must contain formulas.",
          }),
        }),
        D &&
          i.jsxs("div", {
            className: "wi-actions",
            children: [
              i.jsx("button", {
                onClick: () => ae(D, true),
                children: i.jsx(LocalizedText, {
                  children: "Review saved baseline",
                }),
              }),
              i.jsx("button", {
                disabled: e.readOnly,
                onClick: () =>
                  W(() => {
                    (e.onWorkbookChange(
                      findSavedWhatIfAnalysis(e.workbook, D, $),
                      "Remove saved what-if definition",
                    ),
                      B(""),
                      N(null));
                  }),
                children: i.jsx(LocalizedText, {
                  children: "Remove saved definition",
                }),
              }),
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "Saved locations are historical. Check them after moving or deleting source rows; preview again before any action. Removing a definition keeps its report sheets.",
                }),
              }),
            ],
          }),
        l === "scenarios"
          ? i.jsxs("fieldset", {
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "Named scenarios",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Enter values in the same order as the changing cells. Up to 32 inputs and 100 named scenarios.",
                  }),
                }),
                O.map((ge, se) =>
                  i.jsxs(
                    "div",
                    {
                      className: "wi-scenario",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Scenario " }),
                            se + 1,
                            " ",
                            i.jsx(LocalizedText, { children: "name" }),
                            i.jsx("input", {
                              "aria-label": `Scenario ${se + 1} name`,
                              value: ge.name,
                              maxLength: 100,
                              onChange: (event) =>
                                q(() =>
                                  j(
                                    O.map((Ee, Ae) =>
                                      Ae === se
                                        ? { ...Ee, name: event.target.value }
                                        : Ee,
                                    ),
                                  ),
                                ),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Scenario " }),
                            se + 1,
                            " ",
                            i.jsx(LocalizedText, { children: "values" }),
                            i.jsx("input", {
                              "aria-label": `Scenario ${se + 1} values`,
                              placeholder: a ? "90000; 45000" : "90000, 45000",
                              value: ge.values,
                              onChange: (event) =>
                                q(() =>
                                  j(
                                    O.map((Ee, Ae) =>
                                      Ae === se
                                        ? { ...Ee, values: event.target.value }
                                        : Ee,
                                    ),
                                  ),
                                ),
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: O.length === 1,
                          "aria-label": `Remove scenario ${se + 1}`,
                          onClick: () =>
                            q(() => j(O.filter((we, Ee) => Ee !== se))),
                          children: i.jsx(LocalizedText, {
                            children: "Remove",
                          }),
                        }),
                      ],
                    },
                    se,
                  ),
                ),
                i.jsx("button", {
                  disabled: O.length >= 100,
                  onClick: () =>
                    q(() =>
                      j([
                        ...O,
                        { name: `Scenario ${O.length + 1}`, values: "" },
                      ]),
                    ),
                  children: i.jsx(LocalizedText, { children: "Add scenario" }),
                }),
              ],
            })
          : i.jsxs("div", {
              className: "gs-fields",
              children: [
                i.jsxs("label", {
                  children: [
                    l === "two-variable"
                      ? "Across-row input values (first changing cell)"
                      : "Trial input values",
                    i.jsx("textarea", {
                      "aria-label":
                        l === "two-variable"
                          ? "Across-row input values"
                          : "Trial input values",
                      placeholder: a
                        ? "80000; 100000; 120000"
                        : "80000, 100000, 120000",
                      value: A,
                      onChange: (event) => q(() => k(event.target.value)),
                    }),
                  ],
                }),
                l === "two-variable"
                  ? i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children:
                            "Down-column input values (second changing cell)",
                        }),
                        i.jsx("textarea", {
                          "aria-label": text("Down-column input values"),
                          placeholder: a
                            ? "40000; 50000; 60000"
                            : "40000, 50000, 60000",
                          value: F,
                          onChange: (event) => q(() => I(event.target.value)),
                          lang: language("Down-column input values"),
                          "data-ui-label": "Down-column input values",
                        }),
                      ],
                    })
                  : i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Values arranged" }),
                        i.jsxs("select", {
                          "aria-label": text("Values arranged"),
                          value: orientation,
                          onChange: (event) => q(() => E(event.target.value)),
                          lang: language("Values arranged"),
                          "data-ui-label": "Values arranged",
                          children: [
                            i.jsx("option", {
                              value: "column",
                              children: "Down a column",
                            }),
                            i.jsx("option", {
                              value: "row",
                              children: "Across a row",
                            }),
                          ],
                        }),
                      ],
                    }),
              ],
            }),
        e.allowHidden &&
          i.jsxs("label", {
            className: "wi-check",
            children: [
              i.jsx("input", {
                type: "checkbox",
                checked: checked,
                onChange: (event) => q(() => V(event.target.checked)),
              }),
              i.jsx(LocalizedText, { children: "Include hidden sources" }),
            ],
          }),
        R && i.jsx("p", { role: "alert", className: "gs-error", children: R }),
        i.jsx("button", {
          className: "gs-primary",
          onClick: onClick,
          children: i.jsx(LocalizedText, { children: "Preview analysis" }),
        }),
        T &&
          i.jsxs("section", {
            className: "gs-preview",
            "aria-label": text("What-if results"),
            lang: language("What-if results"),
            "data-ui-label": "What-if results",
            children: [
              X$1
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("h3", {
                        children: i.jsx(LocalizedText, {
                          children: "Review results",
                        }),
                      }),
                      i.jsxs("p", {
                        role: "status",
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Baseline inputs are unchanged.",
                          }),
                          " ",
                          T.trials.filter((trial) => !trial.error).length,
                          " ",
                          i.jsx(LocalizedText, { children: "of" }),
                          " ",
                          T.trials.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "trials succeeded. ",
                          }),
                          T.dependencyCells,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "dependency cells; ",
                          }),
                          T.workUnits,
                          " ",
                          i.jsx(LocalizedText, { children: "work units." }),
                        ],
                      }),
                      T.baseline.error &&
                        i.jsxs("p", {
                          role: "alert",
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Baseline calculation failed: ",
                            }),
                            T.baseline.error,
                          ],
                        }),
                      i.jsx("div", {
                        className: "wi-scroll",
                        children: i.jsx("table", {
                          "aria-label": text("Sensitivity results"),
                          lang: language("Sensitivity results"),
                          "data-ui-label": "Sensitivity results",
                          children: i.jsx("tbody", {
                            children: Se.slice(4, le - 1).map((ge, se) =>
                              i.jsx(
                                "tr",
                                {
                                  children: ge.map((we, Ee) =>
                                    Ee === 0
                                      ? i.jsx(
                                          "th",
                                          {
                                            scope: "row",
                                            children: String(we ?? ""),
                                          },
                                          Ee,
                                        )
                                      : i.jsx(
                                          "td",
                                          { children: String(we ?? "") },
                                          Ee,
                                        ),
                                  ),
                                },
                                se,
                              ),
                            ),
                          }),
                        }),
                      }),
                      i.jsxs("details", {
                        children: [
                          i.jsx("summary", {
                            children: i.jsx(LocalizedText, {
                              children:
                                "Inspect source assumptions and formulas",
                            }),
                          }),
                          i.jsx("div", {
                            className: "wi-scroll",
                            children: i.jsxs("table", {
                              "aria-label": text("Source assumptions"),
                              lang: language("Source assumptions"),
                              "data-ui-label": "Source assumptions",
                              children: [
                                i.jsx("thead", {
                                  children: i.jsxs("tr", {
                                    children: [
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Source",
                                        }),
                                      }),
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Original value / formula",
                                        }),
                                      }),
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Literal text",
                                        }),
                                      }),
                                    ],
                                  }),
                                }),
                                i.jsx("tbody", {
                                  children: T.assumptions.map(
                                    (assumption, se) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsxs("th", {
                                              scope: "row",
                                              children: [
                                                assumption.sheetName,
                                                "!",
                                                assumption.address,
                                              ],
                                            }),
                                            i.jsx("td", {
                                              children: String(assumption.raw),
                                            }),
                                            i.jsx("td", {
                                              children: assumption.literal
                                                ? "Yes"
                                                : "No",
                                            }),
                                          ],
                                        },
                                        se,
                                      ),
                                  ),
                                }),
                              ],
                            }),
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "wi-actions",
                        children: [
                          i.jsx("button", {
                            onClick: he,
                            children: i.jsx(LocalizedText, {
                              children: "Export reviewed evidence",
                            }),
                          }),
                          i.jsx("button", {
                            disabled: e.readOnly,
                            onClick: oe,
                            children: D
                              ? "Update analysis and create report"
                              : "Save analysis and create report",
                          }),
                        ],
                      }),
                      T.request.kind === "scenarios" &&
                        i.jsxs("div", {
                          className: "wi-actions",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Scenario to apply",
                                }),
                                i.jsx("select", {
                                  "aria-label": text("Scenario to apply"),
                                  value: _,
                                  onChange: (event) => C(event.target.value),
                                  lang: language("Scenario to apply"),
                                  "data-ui-label": "Scenario to apply",
                                  children: T.trials.map((trial, se) =>
                                    i.jsxs(
                                      "option",
                                      {
                                        value: se,
                                        disabled: !!trial.error,
                                        children: [
                                          trial.name,
                                          trial.error ? " — failed" : "",
                                        ],
                                      },
                                      se,
                                    ),
                                  ),
                                }),
                              ],
                            }),
                            i.jsx("button", {
                              className: "gs-primary",
                              disabled:
                                e.readOnly || !!T.trials[Number(_)]?.error,
                              onClick: de,
                              children: i.jsx(LocalizedText, {
                                children: "Apply selected scenario",
                              }),
                            }),
                          ],
                        }),
                    ],
                  })
                : i.jsx("p", {
                    role: "status",
                    children: i.jsx(LocalizedText, {
                      children:
                        "This preview is stale. Results and assumptions are withheld; preview again using the current workbook and permissions.",
                    }),
                  }),
              i.jsx("button", {
                onClick: () => N(null),
                children: i.jsx(LocalizedText, { children: "Discard preview" }),
              }),
            ],
          }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Supported scope and limits",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Numeric inputs and deterministic scalar formulas. One-variable tables support up to 16 result formulas; two-variable tables support one. Up to 100 values per axis, 400 combinations and 50,000 dependency-cell work units. Invalid assumptions and failed calculations never change the source.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Reports are snapshots. Refresh by loading a saved definition and previewing again; previous reports remain intact. Native JSON retains definitions and evidence. XLSX retains report cells, without native Excel scenario or data-table definitions. Volatile, dynamic, array, blank and incomplete dependencies are excluded. Multivariable Solver is separate.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: e.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
function DelimitedImportExportDialog({
  source,
  blocked,
  onClose,
  onImport,
  onExport,
  onReport,
}) {
  const { text, language } = useWorkbookMessages(),
    u = "bytes" in source,
    l = ne.useId(),
    c = ne.useRef(null),
    operationScope = useOperationScope(),
    f = ne.useRef(0),
    p = ne.useRef(null),
    [m, v] = ne.useState(() => ({
      ...delimitedImportDefaults,
      columns: {},
      includeHeader: true,
      bom: false,
      quoteAll: false,
      dateOutput: "calendar",
      ...(u
        ? {}
        : {
            delimiter: ",",
            encoding: "utf-8",
            newline: `\r
`,
            range:
              source.population?.range ??
              getWorksheetUsedRange(source.snapshot, source.sheetId),
            ...(source.population
              ? {
                  population: {
                    definition: source.population,
                    scope: "full",
                  },
                }
              : {}),
            hasHeader: source.population
              ? true
              : (source.snapshot.sheets.find(
                  (sheet) => sheet.id === source.sheetId,
                )?.delimitedSource?.hasHeader ?? true),
          }),
    })),
    g = (R) => {
      try {
        return {
          result: u
            ? undefined
            : exportWorksheetToCsv(source.snapshot, source.sheetId, R),
          error: "",
        };
      } catch (error) {
        return {
          result: undefined,
          error: error instanceof Error ? error.message : "Preview failed.",
        };
      }
    },
    [b, y] = ne.useState(() =>
      u && source.initial ? { result: source.initial, error: "" } : g(m),
    ),
    [disabled, E] = ne.useState(u && !source.initial),
    [A, k] = ne.useState(false),
    [F, I] = ne.useState(0),
    [O, j] = ne.useState(""),
    z = b.result?.report,
    V = Math.max(1, Math.ceil((z?.columns ?? 1) / 20));
  function D(R, L) {
    (v((U) => ({ ...U, [R]: L })),
      k(true),
      f.current++,
      p.current?.abort(),
      E(false),
      j(""));
  }
  async function onClick() {
    const R = ++f.current;
    p.current?.abort();
    const L = new AbortController();
    if (((p.current = L), u)) {
      E(true);
      try {
        const U = await importDelimitedFile(
          source.file ?? new File([source.bytes], source.fileName),
          m,
          {
            ...operationScope,
            signal: operationScope.signal
              ? AbortSignal.any([L.signal, operationScope.signal])
              : L.signal,
          },
        );
        if (f.current !== R) return;
        y({ result: U, error: "" });
      } catch (error) {
        if (f.current !== R) return;
        y({
          error: error instanceof Error ? error.message : "Preview failed.",
        });
      } finally {
        f.current === R && E(false);
      }
    } else y(g(m));
    (k(false), I(0), j(""));
  }
  (ne.useEffect(
    () => (
      u && !source.initial && onClick(),
      () => {
        (f.current++, p.current?.abort());
      }
    ),
    [source],
  ),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        c.current?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const _ = (R, L, $) =>
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: R }),
          i.jsx("select", {
            value: String(m[L]),
            onChange: (event) => D(L, event.target.value),
            children: $.map(([U, G]) =>
              i.jsx("option", { value: U, children: text(G) }, U),
            ),
          }),
        ],
      }),
    C = (R, L) =>
      i.jsxs("label", {
        className: "lg-delimited-check",
        children: [
          i.jsx("input", {
            type: "checkbox",
            checked: m[L],
            onChange: (event) => D(L, event.target.checked),
          }),
          R,
        ],
      });
  return i.jsx("div", {
    className: "lg-delimited-overlay",
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": l,
      tabIndex: -1,
      ref: c,
      className: "lg-delimited-review",
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), onClose()),
          event.key === "Tab")
        ) {
          const $ = [
              ...(c.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
              ) ?? []),
            ],
            U = $[0],
            G = $.at(-1);
          event.shiftKey &&
          (document.activeElement === U || document.activeElement === c.current)
            ? (event.preventDefault(), G == null || G.focus())
            : !event.shiftKey &&
              document.activeElement === G &&
              (event.preventDefault(), U == null || U.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: l,
              children: i.jsx(LocalizedText, {
                children: u
                  ? "Review delimited import"
                  : "Review delimited export",
              }),
            }),
            i.jsxs("p", {
              children: [
                u ? source.fileName : source.snapshot.name,
                " ",
                i.jsx(LocalizedText, { children: "· CSV / TSV" }),
              ],
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-delimited-scroll",
          children: [
            i.jsx("p", {
              children: u
                ? "Opening replaces the workbook after review. Cancel keeps it unchanged; Undo restores it after importing."
                : `Export calculated values from revision ${source.snapshot.revision}. Choose full or visible scope when a column filter is active; otherwise the rectangle includes hidden and filtered rows. CSV/TSV cannot retain workbook formatting or formulas.`,
            }),
            !u &&
              source.population &&
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Export population" }),
                  i.jsxs("select", {
                    "aria-label": text("Export population"),
                    value: m.population?.scope ?? "full",
                    onChange: (event) =>
                      D("population", {
                        definition: source.population,
                        scope: event.target.value,
                      }),
                    lang: language("Export population"),
                    "data-ui-label": "Export population",
                    children: [
                      i.jsx("option", {
                        value: "full",
                        children:
                          "Full filter range, including hidden and filtered rows",
                      }),
                      i.jsx("option", {
                        value: "visible",
                        children: "Visible population only",
                      }),
                    ],
                  }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Preview again after changing scope. Save the schema report to retain criteria and original row identities.",
                    }),
                  }),
                ],
              }),
            i.jsxs("fieldset", {
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "File and data controls",
                  }),
                }),
                i.jsxs("div", {
                  className: "lg-delimited-controls",
                  children: [
                    _("Delimiter", "delimiter", [
                      ...(u ? [["auto", "Automatic"]] : []),
                      [",", "Comma (CSV)"],
                      ["	", "Tab (TSV)"],
                      [";", "Semicolon"],
                      ["|", "Pipe"],
                    ]),
                    _("Encoding", "encoding", [
                      ...(u ? [["auto", "Automatic: BOM or UTF-8"]] : []),
                      ["utf-8", "UTF-8"],
                      ["utf-16le", "UTF-16 little-endian"],
                      ["utf-16be", "UTF-16 big-endian"],
                      ["windows-1252", "Windows-1252"],
                    ]),
                    _("Quote character", "quote", [
                      ['"', 'Double quote (")'],
                      ["'", "Single quote (')"],
                      ["none", "No quotes"],
                    ]),
                    _("Quote escape", "escape", [
                      ["double", "Double the quote"],
                      ["backslash", "Backslash"],
                    ]),
                    _("Record newline", "newline", [
                      ...(u ? [["auto", "Automatic"]] : []),
                      [
                        `\r
`,
                        "CRLF",
                      ],
                      [
                        `
`,
                        "LF",
                      ],
                      ["\r", "CR"],
                    ]),
                    _("Decimal separator", "decimal", [
                      [".", "Dot"],
                      [",", "Comma"],
                    ]),
                    _("Grouping separator (input text)", "grouping", [
                      ["none", "None"],
                      [",", "Comma"],
                      [".", "Dot"],
                      [" ", "Space"],
                    ]),
                    _("Date order", "dateOrder", [
                      ["ymd", "Year / month / day"],
                      ["dmy", "Day / month / year"],
                      ["mdy", "Month / day / year"],
                    ]),
                    _("Empty fields", "empty", [
                      ["blank", "Blank cell"],
                      ["text", "Empty text"],
                    ]),
                    C("First record is a header", "hasHeader"),
                    C("Trim surrounding whitespace", "trim"),
                    !u &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Export range",
                              }),
                              i.jsx("input", {
                                value: m.range ?? "",
                                onChange: (event) =>
                                  D("range", event.target.value),
                              }),
                            ],
                          }),
                          _("Date output", "dateOutput", [
                            ["calendar", "Calendar date/time (nearest second)"],
                            ["serial", "Native serial number"],
                          ]),
                          C("Include header in output", "includeHeader"),
                          C("Include byte order mark (BOM)", "bom"),
                          C("Quote all fields", "quoteAll"),
                        ],
                      }),
                  ],
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Auto preserves leading-zero identifiers, long numbers, dates and Boolean text. Choose a column type to convert it. Dates need four-digit years. Empty records retain their source row positions.",
              }),
            }),
            z &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Column schema",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      z.records.toLocaleString(),
                      " ",
                      i.jsx(LocalizedText, { children: "records ·" }),
                      " ",
                      z.columns.toLocaleString(),
                      " ",
                      i.jsx(LocalizedText, { children: "columns ·" }),
                      " ",
                      z.storedCells.toLocaleString(),
                      " ",
                      i.jsx(LocalizedText, { children: "stored cells ·" }),
                      " ",
                      (z.fileBytes / 1024).toFixed(1),
                      " ",
                      i.jsx(LocalizedText, { children: "KiB" }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "lg-delimited-schema",
                    children: z.schema.slice(F * 20, F * 20 + 20).map((R) =>
                      i.jsxs(
                        "label",
                        {
                          children: [
                            columnIndexToLetters(R.index),
                            " — ",
                            R.label || "(blank header)",
                            i.jsx("select", {
                              "aria-label": `Type for column ${columnIndexToLetters(R.index)}`,
                              value: m.columns[R.index] ?? "auto",
                              onChange: (event) =>
                                D("columns", {
                                  ...m.columns,
                                  [R.index]: event.target.value,
                                }),
                              children: [
                                "auto",
                                "text",
                                "number",
                                "integer",
                                "boolean",
                                "date",
                              ].map((L) =>
                                i.jsx("option", { value: L, children: L }, L),
                              ),
                            }),
                            i.jsxs("small", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Observed: ",
                                }),
                                R.observed.join(", ") || "header only",
                              ],
                            }),
                          ],
                        },
                        R.index,
                      ),
                    ),
                  }),
                  V > 1 &&
                    i.jsxs("div", {
                      children: [
                        i.jsx("button", {
                          disabled: F === 0,
                          onClick: () => I((R) => R - 1),
                          children: i.jsx(LocalizedText, {
                            children: "Previous columns",
                          }),
                        }),
                        " ",
                        i.jsx(LocalizedText, { children: "Page " }),
                        F + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "of " }),
                        V,
                        " ",
                        i.jsx("button", {
                          disabled: F + 1 >= V,
                          onClick: () => I((R) => R + 1),
                          children: i.jsx(LocalizedText, {
                            children: "Next columns",
                          }),
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("button", {
              onClick: onClick,
              disabled: disabled,
              children: i.jsx(LocalizedText, { children: "Refresh preview" }),
            }),
            disabled &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "Processing file in an isolated worker…",
                }),
              }),
            A &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Settings changed. Refresh preview before continuing.",
                }),
              }),
            b.error && i.jsx("p", { role: "alert", children: b.error }),
            z &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Data preview",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, {
                        children:
                          "First 25 logical records and 200 columns. Validation checks the entire file or export range. ",
                      }),
                      z.issueCount,
                      " ",
                      i.jsx(LocalizedText, { children: "issues;" }),
                      " ",
                      z.changeCount,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "conversions or text changes, including",
                      }),
                      " ",
                      z.safeTextChanges,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "safe-export changes.",
                      }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "lg-delimited-table",
                    children: i.jsxs("table", {
                      "aria-label": text("Delimited data preview"),
                      lang: language("Delimited data preview"),
                      "data-ui-label": "Delimited data preview",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Record",
                                }),
                              }),
                              z.schema.slice(0, 200).map((R) =>
                                i.jsxs(
                                  "th",
                                  {
                                    children: [
                                      columnIndexToLetters(R.index),
                                      " — ",
                                      R.label,
                                    ],
                                  },
                                  R.index,
                                ),
                              ),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: z.preview.map((cell) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", { children: cell.record }),
                                  cell.values.map((value, $) =>
                                    i.jsx(
                                      "td",
                                      {
                                        title: cell.raw[$],
                                        children:
                                          value === null
                                            ? "(blank)"
                                            : value === ""
                                              ? "(empty text)"
                                              : String(value),
                                      },
                                      $,
                                    ),
                                  ),
                                ],
                              },
                              cell.record,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  !!z.issueCount &&
                    i.jsxs("div", {
                      role: "alert",
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(LocalizedText, {
                            children: "Resolve issues before continuing",
                          }),
                        }),
                        i.jsx("ul", {
                          children: z.issues.map((issue, L) =>
                            i.jsxs(
                              "li",
                              {
                                children: [
                                  i.jsx(LocalizedText, { children: "Record " }),
                                  issue.record,
                                  issue.column
                                    ? `, column ${issue.column}`
                                    : "",
                                  ": ",
                                  issue.message,
                                  issue.value !== undefined
                                    ? ` Value: ${issue.value}`
                                    : "",
                                ],
                              },
                              L,
                            ),
                          ),
                        }),
                        z.issueCount > z.issues.length &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Showing the first ",
                              }),
                              z.issues.length,
                              " ",
                              i.jsx(LocalizedText, { children: "of" }),
                              " ",
                              z.issueCount,
                              " ",
                              i.jsx(LocalizedText, { children: "issues." }),
                            ],
                          }),
                      ],
                    }),
                  i.jsxs("details", {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Conversions and changed text (",
                          }),
                          z.changeCount,
                          ")",
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "The saved report contains every changed value; the dialog shows the first 100.",
                        }),
                      }),
                      i.jsxs("table", {
                        "aria-label": text("Delimited changes"),
                        lang: language("Delimited changes"),
                        "data-ui-label": "Delimited changes",
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
                                    children: "Before",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "After",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Reason",
                                  }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: z.changes.slice(0, 100).map((R, L) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("td", { children: R.address }),
                                    i.jsx("td", { children: R.before }),
                                    i.jsx("td", { children: R.after }),
                                    i.jsx("td", { children: R.reason }),
                                  ],
                                },
                                L,
                              ),
                            ),
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Exchange policy",
                    }),
                  }),
                  i.jsx("ul", {
                    children: z.warnings.map((warning, L) =>
                      i.jsx("li", { children: warning }, L),
                    ),
                  }),
                ],
              }),
          ],
        }),
        (blocked || O) && i.jsx("p", { role: "alert", children: blocked || O }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: !z || A || disabled,
              onClick: () => z && onReport(z),
              children: i.jsx(LocalizedText, {
                children: "Save schema and changes",
              }),
            }),
            i.jsx("button", {
              disabled: A || disabled || !!blocked || !z?.valid,
              onClick: () => {
                var R, L;
                if (!(A || disabled || blocked || !z?.valid))
                  try {
                    u
                      ? (R = b.result) != null &&
                        R.workbook &&
                        onImport(b.result.workbook, z)
                      : (L = b.result) != null &&
                        L.bytes &&
                        onExport(b.result.bytes, z);
                  } catch (error) {
                    j(
                      error instanceof Error
                        ? error.message
                        : "Exchange failed.",
                    );
                  }
              },
              children: i.jsx(LocalizedText, {
                children: u ? "Import reviewed data" : "Download reviewed data",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
function tokenizeVbaSource(e) {
  const r = [];
  let t = 0,
    line = 1;
  const o = (a) => {
    throw { line: line, message: a };
  };
  for (; t < e.length;) {
    const a = e[t];
    if (/\s/.test(a)) {
      (a ===
        `
` && line++,
        t++);
      continue;
    }
    if (e.slice(t, t + 2) === "//") {
      for (
        ;
        t < e.length &&
        e[t] !==
          `
`;
      )
        t++;
      continue;
    }
    if (e.slice(t, t + 2) === "/*") {
      for (t += 2; t < e.length && e.slice(t, t + 2) !== "*/";)
        e[t++] ===
          `
` && line++;
      (t === e.length && o("Unclosed block comment."), (t += 2));
      continue;
    }
    const s = line;
    if (a === '"' || a === "'") {
      let d = "";
      const u = a;
      for (t++; t < e.length && e[t] !== u;) {
        let l = e[t++];
        if (
          ((l ===
            `
` ||
            l === "\r") &&
            o("Use an escaped newline inside a string literal."),
          l === "\\")
        ) {
          const c = e[t++],
            h = {
              n: `
`,
              r: "\r",
              t: "	",
              b: "\b",
              f: "\f",
              v: "\v",
              "\\": "\\",
              '"': '"',
              "'": "'",
            };
          if (c === "u" || c === "x") {
            const f = c === "u" ? 4 : 2,
              p = e.slice(t, t + f);
            (new RegExp(`^[a-f0-9]{${f}}$`, "i").test(p) ||
              o("Invalid string escape."),
              (l = String.fromCharCode(parseInt(p, 16))),
              (t += f));
          } else
            Object.hasOwn(h, c) ? (l = h[c]) : o("Unsupported string escape.");
        }
        d += l;
      }
      (e[t++] !== u && o("Unclosed string literal."),
        r.push({ type: "str", value: d, line: s }));
    } else {
      const d = e.slice(t),
        u = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(d),
        l = /^[A-Za-z_$][\w$]*/.exec(d);
      u
        ? (r.push({ type: "number", value: u[0], line: s }), (t += u[0].length))
        : l
          ? (r.push({ type: "id", value: l[0], line: s }), (t += l[0].length))
          : "().,;:=[]{}+-".includes(a)
            ? (r.push({ type: "symbol", value: a, line: s }), t++)
            : o(
                `Unsupported source token ${a}. Templates, expressions and arbitrary JavaScript are outside this adapter.`,
              );
    }
    r.length > automationLimits.maxTokens &&
      o("Source exceeds the 20,000-token migration limit.");
  }
  return r;
}
class VbaAssignmentParser {
  constructor(r) {
    defineDataPropertyByKey(this, "pos", 0);
    defineDataPropertyByKey(this, "bindings", new Map());
    defineDataPropertyByKey(this, "commands", []);
    defineDataPropertyByKey(this, "issues", []);
    this.ts = r;
  }
  peek(r) {
    return r === undefined ? this.ts[this.pos] : this.ts[this.pos]?.value === r;
  }
  error(r) {
    throw {
      line: this.ts[this.pos]?.line ?? this.ts.at(-1)?.line ?? 1,
      message: r,
    };
  }
  take(r) {
    const t = this.ts[this.pos++];
    return (
      (!t ||
        (r !== undefined &&
          (t.value !== r ||
            t.type !== ("().,;:=[]{}+-".includes(r) ? "symbol" : "id")))) &&
        (this.pos--,
        this.error(
          `Expected ${r ?? "source token"}; source is outside the literal-range grammar.`,
        )),
      t
    );
  }
  id() {
    const r = this.take();
    return (
      r.type !== "id" && this.error("Expected a plain identifier."),
      r.value
    );
  }
  literal() {
    if (this.peek("[")) {
      this.take("[");
      const r = [];
      for (; !this.peek("]");) {
        this.take("[");
        const t = [];
        for (; !this.peek("]");) {
          const n = this.scalar();
          if (
            (t.push(n),
            t.length > automationLimits.maxCells &&
              this.error("Matrix exceeds 500 cells."),
            this.peek(","))
          )
            this.take(",");
          else break;
        }
        if (
          (this.take("]"),
          r.push(t),
          (r.length > automationLimits.maxCells ||
            r.reduce((n, o) => n + o.length, 0) > 500) &&
            this.error("Matrix exceeds 500 cells."),
          this.peek(","))
        )
          this.take(",");
        else break;
      }
      return (this.take("]"), r);
    }
    return this.scalar();
  }
  scalar() {
    let r = 1;
    const t = this.peek("-") || this.peek("+");
    this.peek("-")
      ? ((r = -1), this.take("-"))
      : this.peek("+") && this.take("+");
    const n = this.take();
    if (n.type === "number") {
      const o = r * Number(n.value);
      return (
        Number.isFinite(o) || this.error("Numeric literal must be finite."),
        o
      );
    }
    if (
      (t && this.error("Only numeric literals may have a unary sign."),
      n.type === "str")
    )
      return n.value;
    if (n.type === "id" && ["true", "false", "null"].includes(n.value))
      return n.value === "null" ? null : n.value === "true";
    this.error(
      "Only literal text, numbers, booleans and rectangular literal matrices are supported.",
    );
  }
  argument() {
    const r = this.peek();
    if (
      (r || this.error("Missing argument."),
      r.type === "id" && !["true", "false", "null"].includes(r.value))
    ) {
      const t = this.bindings.get(this.id());
      return (t || this.error("Unknown or dynamic binding."), t);
    }
    return { kind: "value", value: this.literal() };
  }
  expression() {
    let r = this.argument();
    for (; this.peek(".");) {
      this.take(".");
      const t = this.id(),
        line = this.ts[this.pos - 1].line;
      if ((this.take("("), r.kind === "workbook" && t === "getActiveWorksheet"))
        (this.take(")"), (r = { kind: "sheet" }));
      else if (r.kind === "workbook" && t === "getWorksheet") {
        const o = this.argument();
        (this.take(")"),
          (o.kind !== "value" || typeof o.value != "string") &&
            this.error("Worksheet name must be literal text."),
          (r = { kind: "sheet", sheet: o.value }));
      } else if (r.kind === "sheet" && t === "getRange") {
        const o = this.argument();
        (this.take(")"),
          (o.kind !== "value" || typeof o.value != "string") &&
            this.error("Range address must be literal text."),
          (r = { kind: "range", sheet: r.sheet, range: o.value }));
      } else if (
        r.kind === "range" &&
        ["setValue", "setValues", "setFormula", "setFormulas"].includes(t)
      ) {
        const o = this.argument();
        (this.take(")"),
          o.kind !== "value" &&
            this.error("Range write arguments must be literals."));
        const a = Array.isArray(o.value);
        (a !== ["setValues", "setFormulas"].includes(t) &&
          this.error(`${t} requires ${a ? "a scalar" : "a matrix"}.`),
          (r = {
            kind: "command",
            command: {
              line: line,
              kind: t.includes("Formula") ? "set-formulas" : "set-values",
              sheet: r.sheet,
              range: r.range,
              values: o.value,
            },
          }));
      } else if (r.kind === "range" && t === "clear") {
        for (const o of [
          "ExcelScript",
          ".",
          "ClearApplyTo",
          ".",
          "contents",
          ")",
        ])
          this.take(o);
        r = {
          kind: "command",
          command: {
            line: line,
            kind: "clear-contents",
            sheet: r.sheet,
            range: r.range,
          },
        };
      } else
        this.error(
          `Unsupported Office Scripts API: ${t}. No part of this script can run until every issue is resolved.`,
        );
    }
    return r;
  }
  parse() {
    try {
      (this.take("function"), this.take("main"), this.take("("));
      const r = this.id();
      for (
        this.take(":"),
          this.take("ExcelScript"),
          this.take("."),
          this.take("Workbook"),
          this.take(")"),
          this.take("{"),
          this.bindings.set(r, { kind: "workbook" });
        this.pos < this.ts.length && !this.peek("}");
      )
        try {
          if (this.peek("const") || this.peek("let")) {
            this.take();
            const t = this.id();
            (this.bindings.has(t) &&
              this.error("Binding redeclaration is not supported."),
              this.take("="));
            const n = this.expression();
            (n.kind === "command" &&
              this.error("A command cannot be assigned to a binding."),
              this.take(";"),
              this.bindings.set(t, n));
          } else {
            const t = this.expression();
            (t.kind !== "command" &&
              this.error("Only literal range commands may be executed."),
              this.take(";"),
              this.commands.push(t.command));
          }
        } catch (error) {
          for (
            this.issues.push(createVbaSyntaxError(error));
            this.pos < this.ts.length && !this.peek(";") && !this.peek("}");
          )
            this.pos++;
          this.peek(";") && this.pos++;
        }
      (this.take("}"),
        this.pos !== this.ts.length &&
          this.error(
            "Additional procedures or top-level executable statements are unsupported.",
          ));
    } catch (error) {
      this.issues.push(createVbaSyntaxError(error));
    }
  }
}
function createVbaSyntaxError(error) {
  const r = error;
  return {
    line: r.line ?? 1,
    message: r.message ?? "Unsupported source syntax.",
  };
}
const vbaStringLiteralPattern = '"(?:[^"]|"")*"';
function parseVbaAssignmentValue(e) {
  if (new RegExp("^" + vbaStringLiteralPattern + "$").test(e))
    return e.slice(1, -1).replaceAll('""', '"');
  if (/^(true|false)$/i.test(e)) return /^true$/i.test(e);
  if (
    /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(e) &&
    Number.isFinite(Number(e))
  )
    return Number(e);
  throw Error(
    "Only literal text, finite numbers or booleans are supported in VBA assignments.",
  );
}
function splitVbaStatements(e) {
  let r = "",
    t = false;
  for (let n = 0; n < e.length; n++) {
    const o = e[n];
    if (o === '"') {
      if (((r += o), t && e[n + 1] === '"')) {
        r += e[++n];
        continue;
      }
      t = !t;
    } else {
      if (o === "'" && !t) break;
      if (o === ":" && !t)
        throw Error("Multiple colon-separated VBA statements are unsupported.");
      r += o;
    }
  }
  if (t) throw Error("Unclosed VBA string.");
  return r.trim();
}

export {
  WhatIfAnalysisDialog,
  splitVbaStatements,
  vbaStringLiteralPattern,
  tokenizeVbaSource,
  VbaAssignmentParser,
  createVbaSyntaxError,
  parseVbaAssignmentValue,
  DelimitedImportExportDialog,
};
