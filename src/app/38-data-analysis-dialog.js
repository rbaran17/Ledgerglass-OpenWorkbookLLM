import { LabeledField } from "./14-dax-data-model-engine.js";

import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import c_ from "react-dom/server";
import {
  deepFreezeValue,
  fingerprintCanonicalValue,
} from "./01-schemas-date-functions-and-validation.js";
import {
  assertWorkbookStructureEditable,
  formatNumberWithPattern,
  validateObjectReportRequest,
  reportOutputFormatVersion,
  parseReportRange,
  inspectPictureDataUrl,
  formatOccupiedRangeAddress,
  reportOutputLimits,
  hashRecordExcludingHash,
  parseReportSnapshot,
  inspectChartSource,
  _t,
  formatDiagnosticsMessage,
  cn,
  parseRangeAddress,
  applyFilterConditions,
  normalizeWorkbookInput,
  columnIndexToLetters,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { readSourceRectangles } from "./12-consolidation-and-dependency-inspection.js";
import { ChartReportSvg } from "./31-pivot-explorer-and-charts.js";
import { normalizeAnalysisPermissions } from "./36-watch-window-and-goal-seek.js";
import {
  readSavedAnalysisReport,
  commitAnalysisResult,
  runAnalysisRequest,
  parseNumericText,
} from "./37-statistics-and-linear-solver.js";

function DataAnalysisDialog(e) {
  const { text } = useWorkbookMessages(),
    t = ne.useId(),
    n = ne.useRef(null),
    o = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    analysisReport = o.analysisReport,
    [s, d] = ne.useState(
      analysisReport?.request.kind ?? e.initialKind ?? "statistics",
    ),
    [u, l] = ne.useState("Setup"),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState(null),
    [checked, v] = ne.useState(analysisReport?.includeHidden ?? false),
    [g, b] = ne.useState(0),
    y = () => {
      let te = 1;
      for (
        ;
        e.workbook.sheets.some(
          (sheet) => sheet.name.toLowerCase() === `analysis ${te}`,
        );
      )
        te++;
      return `Analysis ${te}`;
    },
    [outputName, E] = ne.useState(y),
    [A, k] = ne.useState(() =>
      analysisReport?.request.kind === "solver"
        ? analysisReport.request
        : {
            kind: "solver",
            outputName: outputName,
            objective: {
              sheetId: e.sheetId,
              address:
                Object.entries(o.cells).find(
                  ([, cell]) =>
                    !cell.literal &&
                    typeof cell.raw == "string" &&
                    cell.raw.startsWith("="),
                )?.[0] ?? "C2",
            },
            direction: "maximize",
            variables: Object.entries(o.cells)
              .filter(([, cell]) => typeof cell.raw == "number")
              .slice(0, 2)
              .map(([address]) => ({
                cell: { sheetId: e.sheetId, address: address },
                lower: 0,
                upper: 100,
                domain: "continuous",
              })),
            constraints: [],
            tolerance: 1e-8,
            maxCandidates: 5e4,
          },
    ),
    [F, I] = ne.useState(() =>
      analysisReport?.request.kind === "statistics"
        ? analysisReport.request
        : {
            kind: "statistics",
            outputName: outputName,
            sheetId: e.sheetId,
            range: e.range.includes(":")
              ? e.range
              : formatOccupiedRangeAddress(o),
            method: "descriptive",
            missing: "reject",
          },
    ),
    [O, j] = ne.useState(F.bins?.join(", ") ?? "10, 20, 30"),
    [z, V] = ne.useState(0),
    [D, B] = ne.useState(-1),
    [_, C] = ne.useState(null),
    [T, N] = ne.useState(""),
    R = JSON.stringify([
      e.readOnly,
      e.allowedSheetIds,
      !!e.allowHidden && checked,
    ]),
    L = !!f && D === z && _ === e.workbook && T === R,
    $ = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      includeHidden: !!e.allowHidden && checked,
    },
    U = () => {
      (V((te) => te + 1), p(null), h(""), b(0));
    },
    G = (te) => {
      (k((ae) => ({ ...ae, ...te })), U());
    },
    H = (te) => {
      (I((ae) => ({ ...ae, ...te })), U());
    },
    Q = e.workbook.sheets.filter(
      (sheet) =>
        !sheet.analysisReport &&
        !sheet.dataModelReport &&
        (!e.allowedSheetIds || e.allowedSheetIds.includes(sheet.id)) &&
        ($.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    X$1 = (te, ae, oe) =>
      i.jsx("select", {
        "aria-label": text(oe),
        value: te,
        onChange: (event) => ae(event.target.value),
        children: Q.map((de) =>
          i.jsx("option", { value: de.id, children: de.name }, de.id),
        ),
      }),
    J = (te, ae, oe) =>
      i.jsx(LocalizedNumberField, {
        "aria-label": text(te),
        value: ae === undefined || !Number.isFinite(ae) ? "" : ae,
        onValueChange: (de) => oe(parseNumericText(de)),
      });
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      n.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      const te = n.current?.querySelector(".lg-model-body");
      te && (te.scrollTop = 0);
    }, [u, f]));
  const onClick = () => {
      (p(null), h(""));
      try {
        const request =
            s === "solver"
              ? { ...A, outputName: outputName }
              : {
                  ...F,
                  outputName: outputName,
                  ...(F.method === "histogram"
                    ? { bins: O.split(",").map(parseNumericText) }
                    : {}),
                },
          ae = runAnalysisRequest(e.workbook, request, $);
        (p(ae), B(z), C(e.workbook), N(R), l("Results"), b(0));
      } catch (error) {
        h(error instanceof Error ? error.message : "Analysis failed.");
      }
    },
    W = (te) => {
      try {
        if (!L || !f)
          throw Error("Preview the current inputs and permissions first.");
        const ae = commitAnalysisResult(e.workbook, f, te, $);
        (e.onWorkbookChange(
          ae,
          te === "report"
            ? "Save analysis report"
            : "Apply Solver solution and report",
        ),
          p(null),
          e.onClose());
      } catch (error) {
        h(
          error instanceof Error
            ? error.message
            : "The analysis was not saved.",
        );
      }
    };
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      ref: n,
      className: "lg-data-model lg-analysis",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": t,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), e.onClose()),
          event.key === "Tab")
        ) {
          const ae = [
              ...n.current.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary,[tabindex="0"]',
              ),
            ].filter(
              (he) =>
                !he.hidden &&
                (he.getClientRects().length ||
                  navigator.userAgent.includes("jsdom")),
            ),
            oe = ae[0],
            de = ae.at(-1);
          event.shiftKey && document.activeElement === oe
            ? (event.preventDefault(), de == null || de.focus())
            : !event.shiftKey &&
              document.activeElement === de &&
              (event.preventDefault(), oe == null || oe.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", {
                  className: "eyebrow",
                  children: i.jsx(LocalizedText, {
                    children: "Data · Analysis",
                  }),
                }),
                i.jsx("h2", {
                  id: t,
                  children: i.jsx(LocalizedText, {
                    children: s === "solver" ? "Solver" : "Data Analysis",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      s === "solver"
                        ? "Define a bounded linear model, review the result, then choose whether to apply it."
                        : "Choose a statistical method and review the sample before saving a report.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close analysis"),
              onClick: e.onClose,
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("nav", {
          "aria-label": text("Analysis sections"),
          children: [
            ["Setup", "Results"].map((te) =>
              i.jsx(
                "button",
                {
                  "aria-pressed": u === te,
                  onClick: () => l(te),
                  children: i.jsx(LocalizedText, { children: te }),
                },
                te,
              ),
            ),
            i.jsx("span", {
              className: "lg-analysis-mode",
              children: LabeledField(
                "Tool",
                i.jsxs("select", {
                  value: s,
                  onChange: (event) => {
                    (d(event.target.value), U(), l("Setup"));
                  },
                  children: [
                    i.jsx("option", { value: "solver", children: "Solver" }),
                    i.jsx("option", {
                      value: "statistics",
                      children: "Data Analysis",
                    }),
                  ],
                }),
              ),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-model-body",
          children: [
            analysisReport &&
              i.jsx("p", {
                className: "lg-analysis-hint",
                children: i.jsx(LocalizedText, {
                  children:
                    "Loaded the saved definition. A new preview and a new report name are required to refresh it.",
                }),
              }),
            c &&
              i.jsx("p", {
                role: "alert",
                className: "lg-analysis-error",
                children: c,
              }),
            u === "Setup"
              ? i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("div", {
                      className: "lg-model-controls",
                      children: [
                        LabeledField(
                          "Output worksheet",
                          i.jsx("input", {
                            value: outputName,
                            onChange: (event) => {
                              (E(event.target.value), U());
                            },
                          }),
                        ),
                        e.allowHidden &&
                          LabeledField(
                            "Include hidden source cells",
                            i.jsx("input", {
                              type: "checkbox",
                              checked: checked,
                              onChange: (event) => {
                                (v(event.target.checked), U());
                              },
                            }),
                          ),
                      ],
                    }),
                    s === "solver"
                      ? i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Objective",
                                  }),
                                }),
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    LabeledField(
                                      "Objective worksheet",
                                      X$1(
                                        A.objective.sheetId,
                                        (sheetId) =>
                                          G({
                                            objective: {
                                              ...A.objective,
                                              sheetId: sheetId,
                                            },
                                          }),
                                        "Objective worksheet",
                                      ),
                                    ),
                                    LabeledField(
                                      "Objective formula cell",
                                      i.jsx("input", {
                                        value: A.objective.address,
                                        onChange: (event) =>
                                          G({
                                            objective: {
                                              ...A.objective,
                                              address: event.target.value
                                                .replace(/\$/g, "")
                                                .toUpperCase(),
                                            },
                                          }),
                                      }),
                                    ),
                                    LabeledField(
                                      "Direction",
                                      i.jsxs("select", {
                                        value: A.direction,
                                        onChange: (event) =>
                                          G({ direction: event.target.value }),
                                        children: [
                                          i.jsx("option", {
                                            value: "maximize",
                                            children: "Maximize",
                                          }),
                                          i.jsx("option", {
                                            value: "minimize",
                                            children: "Minimize",
                                          }),
                                          i.jsx("option", {
                                            value: "target",
                                            children: "Target value",
                                          }),
                                        ],
                                      }),
                                    ),
                                    A.direction === "target" &&
                                      LabeledField(
                                        "Target value",
                                        J("Target value", A.target, (target) =>
                                          G({ target: target }),
                                        ),
                                      ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Changing cells",
                                  }),
                                }),
                                i.jsx("p", {
                                  className: "lg-analysis-hint",
                                  children: i.jsx(LocalizedText, {
                                    children:
                                      "Up to six numeric inputs. Give every input finite lower and upper bounds. Use continuous inputs together, or integer and binary inputs together.",
                                  }),
                                }),
                                A.variables.map((variable, ae) =>
                                  i.jsxs(
                                    "div",
                                    {
                                      className:
                                        "lg-model-controls lg-analysis-row",
                                      children: [
                                        LabeledField(
                                          `Variable ${ae + 1} worksheet`,
                                          X$1(
                                            variable.cell.sheetId,
                                            (sheetId) =>
                                              G({
                                                variables: A.variables.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          cell: {
                                                            ...de.cell,
                                                            sheetId: sheetId,
                                                          },
                                                        }
                                                      : de,
                                                ),
                                              }),
                                            `Variable ${ae + 1} worksheet`,
                                          ),
                                        ),
                                        LabeledField(
                                          `Variable ${ae + 1} cell`,
                                          i.jsx("input", {
                                            value: variable.cell.address,
                                            onChange: (event) =>
                                              G({
                                                variables: A.variables.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          cell: {
                                                            ...de.cell,
                                                            address:
                                                              event.target.value
                                                                .replace(
                                                                  /\$/g,
                                                                  "",
                                                                )
                                                                .toUpperCase(),
                                                          },
                                                        }
                                                      : de,
                                                ),
                                              }),
                                          }),
                                        ),
                                        ["lower", "upper"].map((oe) =>
                                          LabeledField(
                                            `Variable ${ae + 1} ${oe}`,
                                            J(
                                              `Variable ${ae + 1} ${oe}`,
                                              variable[oe],
                                              (de) =>
                                                G({
                                                  variables: A.variables.map(
                                                    (he, Se) =>
                                                      Se === ae
                                                        ? { ...he, [oe]: de }
                                                        : he,
                                                  ),
                                                }),
                                            ),
                                          ),
                                        ),
                                        LabeledField(
                                          `Variable ${ae + 1} domain`,
                                          i.jsxs("select", {
                                            value: variable.domain,
                                            onChange: (event) =>
                                              G({
                                                variables: A.variables.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          domain:
                                                            event.target.value,
                                                          ...(event.target
                                                            .value === "binary"
                                                            ? {
                                                                lower: 0,
                                                                upper: 1,
                                                              }
                                                            : {}),
                                                        }
                                                      : de,
                                                ),
                                              }),
                                            children: [
                                              i.jsx("option", {
                                                value: "continuous",
                                                children: "Continuous",
                                              }),
                                              i.jsx("option", {
                                                value: "integer",
                                                children: "Integer",
                                              }),
                                              i.jsx("option", {
                                                value: "binary",
                                                children: "Binary",
                                              }),
                                            ],
                                          }),
                                        ),
                                        i.jsx("button", {
                                          "aria-label": text(
                                            `Remove variable ${ae + 1}`,
                                          ),
                                          onClick: () =>
                                            G({
                                              variables: A.variables.filter(
                                                (oe, de) => de !== ae,
                                              ),
                                            }),
                                          children: "×",
                                        }),
                                      ],
                                    },
                                    ae,
                                  ),
                                ),
                                i.jsx("button", {
                                  disabled: A.variables.length >= 6,
                                  onClick: () =>
                                    G({
                                      variables: [
                                        ...A.variables,
                                        {
                                          cell: {
                                            sheetId: Q[0]?.id ?? e.sheetId,
                                            address: "A2",
                                          },
                                          lower: 0,
                                          upper: 100,
                                          domain: "continuous",
                                        },
                                      ],
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Add changing cell",
                                  }),
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Constraints",
                                  }),
                                }),
                                A.constraints.map((constraint, ae) =>
                                  i.jsxs(
                                    "div",
                                    {
                                      className:
                                        "lg-model-controls lg-analysis-row",
                                      children: [
                                        LabeledField(
                                          `Constraint ${ae + 1} worksheet`,
                                          X$1(
                                            constraint.left.sheetId,
                                            (sheetId) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          left: {
                                                            ...de.left,
                                                            sheetId: sheetId,
                                                          },
                                                        }
                                                      : de,
                                                ),
                                              }),
                                            `Constraint ${ae + 1} worksheet`,
                                          ),
                                        ),
                                        LabeledField(
                                          `Constraint ${ae + 1} cell`,
                                          i.jsx("input", {
                                            value: constraint.left.address,
                                            onChange: (event) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          left: {
                                                            ...de.left,
                                                            address:
                                                              event.target.value
                                                                .replace(
                                                                  /\$/g,
                                                                  "",
                                                                )
                                                                .toUpperCase(),
                                                          },
                                                        }
                                                      : de,
                                                ),
                                              }),
                                          }),
                                        ),
                                        LabeledField(
                                          `Constraint ${ae + 1} relation`,
                                          i.jsx("select", {
                                            value: constraint.operator,
                                            onChange: (event) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          operator:
                                                            event.target.value,
                                                        }
                                                      : de,
                                                ),
                                              }),
                                            children: ["<=", ">=", "="].map(
                                              (oe) =>
                                                i.jsx(
                                                  "option",
                                                  { children: oe },
                                                  oe,
                                                ),
                                            ),
                                          }),
                                        ),
                                        LabeledField(
                                          `Constraint ${ae + 1} right type`,
                                          i.jsxs("select", {
                                            value:
                                              typeof constraint.right ==
                                              "number"
                                                ? "number"
                                                : "cell",
                                            onChange: (event) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          right:
                                                            event.target
                                                              .value ===
                                                            "number"
                                                              ? 0
                                                              : {
                                                                  sheetId:
                                                                    constraint
                                                                      .left
                                                                      .sheetId,
                                                                  address: "A2",
                                                                },
                                                        }
                                                      : de,
                                                ),
                                              }),
                                            children: [
                                              i.jsx("option", {
                                                value: "number",
                                                children: "Number",
                                              }),
                                              i.jsx("option", {
                                                value: "cell",
                                                children: "Cell",
                                              }),
                                            ],
                                          }),
                                        ),
                                        typeof constraint.right == "number"
                                          ? LabeledField(
                                              `Constraint ${ae + 1} value`,
                                              J(
                                                `Constraint ${ae + 1} value`,
                                                constraint.right,
                                                (right) =>
                                                  G({
                                                    constraints:
                                                      A.constraints.map(
                                                        (de, he) =>
                                                          he === ae
                                                            ? {
                                                                ...de,
                                                                right: right,
                                                              }
                                                            : de,
                                                      ),
                                                  }),
                                              ),
                                            )
                                          : i.jsxs(i.Fragment, {
                                              children: [
                                                LabeledField(
                                                  `Constraint ${ae + 1} right worksheet`,
                                                  X$1(
                                                    constraint.right.sheetId,
                                                    (sheetId) =>
                                                      G({
                                                        constraints:
                                                          A.constraints.map(
                                                            (de, he) =>
                                                              he === ae
                                                                ? {
                                                                    ...de,
                                                                    right: {
                                                                      ...constraint.right,
                                                                      sheetId:
                                                                        sheetId,
                                                                    },
                                                                  }
                                                                : de,
                                                          ),
                                                      }),
                                                    `Constraint ${ae + 1} right worksheet`,
                                                  ),
                                                ),
                                                LabeledField(
                                                  `Constraint ${ae + 1} right cell`,
                                                  i.jsx("input", {
                                                    value:
                                                      constraint.right.address,
                                                    onChange: (event) =>
                                                      G({
                                                        constraints:
                                                          A.constraints.map(
                                                            (de, he) =>
                                                              he === ae
                                                                ? {
                                                                    ...de,
                                                                    right: {
                                                                      ...constraint.right,
                                                                      address:
                                                                        event.target.value
                                                                          .replace(
                                                                            /\$/g,
                                                                            "",
                                                                          )
                                                                          .toUpperCase(),
                                                                    },
                                                                  }
                                                                : de,
                                                          ),
                                                      }),
                                                  }),
                                                ),
                                              ],
                                            }),
                                        i.jsx("button", {
                                          "aria-label": text(
                                            `Remove constraint ${ae + 1}`,
                                          ),
                                          onClick: () =>
                                            G({
                                              constraints: A.constraints.filter(
                                                (oe, de) => de !== ae,
                                              ),
                                            }),
                                          children: "×",
                                        }),
                                      ],
                                    },
                                    ae,
                                  ),
                                ),
                                i.jsx("button", {
                                  disabled: A.constraints.length >= 16,
                                  onClick: () =>
                                    G({
                                      constraints: [
                                        ...A.constraints,
                                        {
                                          left: {
                                            sheetId: A.objective.sheetId,
                                            address: A.objective.address,
                                          },
                                          operator: "<=",
                                          right: 100,
                                        },
                                      ],
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Add constraint",
                                  }),
                                }),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "lg-model-controls",
                              children: [
                                LabeledField(
                                  "Feasibility tolerance",
                                  J(
                                    "Feasibility tolerance",
                                    A.tolerance,
                                    (tolerance) => G({ tolerance: tolerance }),
                                  ),
                                ),
                                LabeledField(
                                  "Maximum candidates",
                                  J(
                                    "Maximum candidates",
                                    A.maxCandidates,
                                    (maxCandidates) =>
                                      G({ maxCandidates: maxCandidates }),
                                  ),
                                ),
                              ],
                            }),
                            i.jsx("p", {
                              className: "lg-analysis-hint",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Supported formulas: numeric cell references, defined names, +, −, ×, ÷, SUM, AVERAGE and linear SUMPRODUCT. Nonlinear and mixed integer models require another engine.",
                              }),
                            }),
                          ],
                        })
                      : i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Source and method",
                                  }),
                                }),
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    LabeledField(
                                      "Source worksheet",
                                      X$1(
                                        F.sheetId,
                                        (sheetId) => H({ sheetId: sheetId }),
                                        "Source worksheet",
                                      ),
                                    ),
                                    LabeledField(
                                      "Headed source range",
                                      i.jsx("input", {
                                        value: F.range,
                                        onChange: (event) =>
                                          H({
                                            range: event.target.value
                                              .replace(/\$/g, "")
                                              .toUpperCase(),
                                          }),
                                      }),
                                    ),
                                    LabeledField(
                                      "Analysis method",
                                      i.jsx("select", {
                                        value: F.method,
                                        onChange: (event) =>
                                          H({ method: event.target.value }),
                                        children: [
                                          [
                                            "descriptive",
                                            "Descriptive statistics",
                                          ],
                                          ["correlation", "Correlation"],
                                          ["covariance", "Covariance"],
                                          ["regression", "Regression"],
                                          ["histogram", "Histogram"],
                                          [
                                            "distribution",
                                            "Probability distributions",
                                          ],
                                          ["forecast", "Forecasting"],
                                        ].map(([te, ae]) =>
                                          i.jsx(
                                            "option",
                                            { value: te, children: ae },
                                            te,
                                          ),
                                        ),
                                      }),
                                    ),
                                    LabeledField(
                                      "Incomplete rows",
                                      i.jsxs("select", {
                                        value: F.missing,
                                        onChange: (event) =>
                                          H({ missing: event.target.value }),
                                        children: [
                                          i.jsx("option", {
                                            value: "reject",
                                            children: "Reject analysis",
                                          }),
                                          i.jsx("option", {
                                            value: "exclude",
                                            children:
                                              "Exclude and list affected rows",
                                          }),
                                        ],
                                      }),
                                    ),
                                  ],
                                }),
                                i.jsx("p", {
                                  className: "lg-analysis-hint",
                                  children: i.jsx(LocalizedText, {
                                    children:
                                      "Include text column headers and numeric data. Up to 500 data rows, six columns and 2,000 cells. Formula inputs are limited to 50 deterministic scalar formulas with bounded dependencies.",
                                  }),
                                }),
                              ],
                            }),
                            F.method === "regression" &&
                              i.jsx("p", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Put the response Y in the first column and predictors X in the remaining columns. The model includes an intercept, coefficient confidence intervals and residuals.",
                                }),
                              }),
                            F.method === "histogram" &&
                              LabeledField(
                                "Bin upper bounds",
                                i.jsx("input", {
                                  value: O,
                                  onChange: (event) => {
                                    (j(event.target.value), U());
                                  },
                                }),
                              ),
                            F.method === "distribution" &&
                              i.jsxs("fieldset", {
                                children: [
                                  i.jsx("legend", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Distribution parameters",
                                    }),
                                  }),
                                  i.jsxs("div", {
                                    className: "lg-model-controls",
                                    children: [
                                      LabeledField(
                                        "Distribution",
                                        i.jsxs("select", {
                                          value: F.distribution ?? "normal",
                                          onChange: (event) =>
                                            H({
                                              distribution: event.target.value,
                                            }),
                                          children: [
                                            i.jsx("option", {
                                              value: "normal",
                                              children: "Normal",
                                            }),
                                            i.jsx("option", {
                                              value: "binomial",
                                              children: "Binomial",
                                            }),
                                            i.jsx("option", {
                                              value: "poisson",
                                              children: "Poisson",
                                            }),
                                          ],
                                        }),
                                      ),
                                      !F.distribution ||
                                      F.distribution === "normal"
                                        ? i.jsxs(i.Fragment, {
                                            children: [
                                              LabeledField(
                                                "Mean",
                                                J("Mean", F.mean, (mean) =>
                                                  H({ mean: mean }),
                                                ),
                                              ),
                                              LabeledField(
                                                "Standard deviation",
                                                J(
                                                  "Standard deviation",
                                                  F.standardDeviation,
                                                  (standardDeviation) =>
                                                    H({
                                                      standardDeviation:
                                                        standardDeviation,
                                                    }),
                                                ),
                                              ),
                                            ],
                                          })
                                        : F.distribution === "binomial"
                                          ? i.jsxs(i.Fragment, {
                                              children: [
                                                LabeledField(
                                                  "Trials",
                                                  J(
                                                    "Trials",
                                                    F.trials,
                                                    (trials) =>
                                                      H({ trials: trials }),
                                                  ),
                                                ),
                                                LabeledField(
                                                  "Success probability",
                                                  J(
                                                    "Success probability",
                                                    F.probability,
                                                    (probability) =>
                                                      H({
                                                        probability:
                                                          probability,
                                                      }),
                                                  ),
                                                ),
                                              ],
                                            })
                                          : LabeledField(
                                              "Rate (lambda)",
                                              J(
                                                "Rate (lambda)",
                                                F.lambda,
                                                (lambda) =>
                                                  H({ lambda: lambda }),
                                              ),
                                            ),
                                    ],
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Select one column of evaluation points. Parameters are supplied explicitly; they are not estimated from that column.",
                                    }),
                                  }),
                                ],
                              }),
                            F.method === "forecast" &&
                              i.jsxs("fieldset", {
                                children: [
                                  i.jsx("legend", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Forecast settings",
                                    }),
                                  }),
                                  i.jsxs("div", {
                                    className: "lg-model-controls",
                                    children: [
                                      LabeledField(
                                        "Forecast method",
                                        i.jsxs("select", {
                                          value: F.forecast ?? "linear",
                                          onChange: (event) =>
                                            H({ forecast: event.target.value }),
                                          children: [
                                            i.jsx("option", {
                                              value: "linear",
                                              children: "Linear trend",
                                            }),
                                            i.jsx("option", {
                                              value: "moving-average",
                                              children: "Moving average",
                                            }),
                                            i.jsx("option", {
                                              value: "exponential",
                                              children: "Exponential smoothing",
                                            }),
                                          ],
                                        }),
                                      ),
                                      LabeledField(
                                        "Forecast periods",
                                        J(
                                          "Forecast periods",
                                          F.horizon ?? 6,
                                          (horizon) => H({ horizon: horizon }),
                                        ),
                                      ),
                                      F.forecast === "moving-average" &&
                                        LabeledField(
                                          "Moving average window",
                                          J(
                                            "Moving average window",
                                            F.window ?? 3,
                                            (te) => H({ window: te }),
                                          ),
                                        ),
                                      F.forecast === "exponential" &&
                                        LabeledField(
                                          "Smoothing alpha",
                                          J(
                                            "Smoothing alpha",
                                            F.alpha ?? 0.3,
                                            (alpha) => H({ alpha: alpha }),
                                          ),
                                        ),
                                    ],
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Select numeric period then value. Periods must be increasing and equally spaced. No seasonal model or prediction interval is claimed.",
                                    }),
                                  }),
                                ],
                              }),
                          ],
                        }),
                  ],
                })
              : i.jsx(i.Fragment, {
                  children: f
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("div", {
                            className: `lg-analysis-summary ${f.result.success ? "" : "failed"}`,
                            role: "status",
                            children: [
                              i.jsx("strong", {
                                children: f.result.status.toUpperCase(),
                              }),
                              i.jsx("span", { children: f.result.message }),
                              i.jsx("span", { children: f.result.algorithm }),
                              i.jsxs("span", {
                                children: [
                                  "Included ",
                                  f.result.included,
                                  " / ",
                                  f.result.inspected,
                                  " ",
                                  "· Excluded ",
                                  f.result.exclusions.length,
                                  " · Candidates",
                                  " ",
                                  f.result.iterations,
                                ],
                              }),
                              f.result.objective !== undefined &&
                                i.jsxs("b", {
                                  children: [
                                    "Objective:",
                                    " ",
                                    f.result.objective.toLocaleString(
                                      undefined,
                                      {
                                        maximumFractionDigits: 10,
                                      },
                                    ),
                                  ],
                                }),
                            ],
                          }),
                          !L &&
                            i.jsx("p", {
                              role: "alert",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "The workbook or permissions changed. Preview again.",
                              }),
                            }),
                          f.result.rows.length > 0 &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsx("div", {
                                  className: "lg-analysis-results",
                                  tabIndex: 0,
                                  "aria-label": text("Analysis results"),
                                  children: i.jsx("table", {
                                    children: i.jsx("tbody", {
                                      children: f.result.rows
                                        .slice(g * 20, g * 20 + 20)
                                        .map((te, ae) =>
                                          i.jsx(
                                            "tr",
                                            {
                                              children: te.map((oe, de) =>
                                                i.jsx(
                                                  "td",
                                                  {
                                                    children:
                                                      oe === null
                                                        ? "Undefined"
                                                        : typeof oe == "number"
                                                          ? oe !== 0 &&
                                                            (Math.abs(oe) <
                                                              1e-6 ||
                                                              Math.abs(oe) >=
                                                                1e12)
                                                            ? oe.toExponential(
                                                                6,
                                                              )
                                                            : oe.toLocaleString(
                                                                undefined,
                                                                {
                                                                  maximumSignificantDigits: 12,
                                                                },
                                                              )
                                                          : String(oe),
                                                  },
                                                  de,
                                                ),
                                              ),
                                            },
                                            ae,
                                          ),
                                        ),
                                    }),
                                  }),
                                }),
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    i.jsx("button", {
                                      disabled: g === 0,
                                      onClick: () => b((te) => te - 1),
                                      children: i.jsx(LocalizedText, {
                                        children: "Previous results",
                                      }),
                                    }),
                                    i.jsxs("span", {
                                      children: [
                                        "Page ",
                                        g + 1,
                                        " /",
                                        " ",
                                        Math.max(
                                          1,
                                          Math.ceil(f.result.rows.length / 20),
                                        ),
                                      ],
                                    }),
                                    i.jsx("button", {
                                      disabled:
                                        (g + 1) * 20 >= f.result.rows.length,
                                      onClick: () => b((te) => te + 1),
                                      children: i.jsx(LocalizedText, {
                                        children: "Next results",
                                      }),
                                    }),
                                  ],
                                }),
                              ],
                            }),
                          i.jsxs("details", {
                            open: true,
                            children: [
                              i.jsx("summary", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Method, assumptions and reference checks",
                                }),
                              }),
                              i.jsx("ul", {
                                children: f.result.assumptions.map(
                                  (assumption, ae) =>
                                    i.jsx("li", { children: assumption }, ae),
                                ),
                              }),
                            ],
                          }),
                          i.jsxs("details", {
                            children: [
                              i.jsxs("summary", {
                                children: [
                                  "Sample exclusions (",
                                  f.result.exclusions.length,
                                  ")",
                                ],
                              }),
                              f.result.exclusions.map((exclusion) =>
                                i.jsxs(
                                  "p",
                                  {
                                    children: [
                                      exclusion.locator,
                                      " — ",
                                      exclusion.reason,
                                    ],
                                  },
                                  exclusion.row,
                                ),
                              ),
                            ],
                          }),
                        ],
                      })
                    : i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Choose the inputs in Setup, then preview the analysis.",
                        }),
                      }),
                }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              className: "lg-analysis-license",
              children: i.jsx(LocalizedText, {
                children:
                  "Local computation · MIT engines · Preview leaves source cells unchanged",
              }),
            }),
            i.jsx("button", {
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Preview analysis" }),
            }),
            i.jsx("button", {
              disabled: e.readOnly || !L || !f?.result.success,
              onClick: () => W("report"),
              children: i.jsx(LocalizedText, { children: "Save report" }),
            }),
            s === "solver" &&
              i.jsx("button", {
                className: "primary",
                disabled: e.readOnly || !L || !f?.result.success,
                onClick: () => W("apply-solution"),
                children: i.jsx(LocalizedText, {
                  children: "Apply solution and report",
                }),
              }),
          ],
        }),
      ],
    }),
  });
}
function AnalysisSnapshotNotice({ workbook, sheetId, onReview }) {
  const n = workbook.sheets.find(
    (sheet) => sheet.id === sheetId,
  )?.analysisReport;
  if (!n) return null;
  const savedAnalysisReport = readSavedAnalysisReport(workbook, sheetId, {
    includeHidden: n.includeHidden,
  });
  return i.jsxs("div", {
    className: "lg-notice",
    role: "status",
    children: [
      i.jsxs("strong", {
        children: ["Analysis snapshot · ", savedAnalysisReport.status],
      }),
      i.jsx("span", { children: savedAnalysisReport.message }),
      i.jsx("button", {
        onClick: onReview,
        children: i.jsx(LocalizedText, { children: "Review analysis" }),
      }),
    ],
  });
}
function captureAnalysisWorkbookState(workbookInput) {
  return fingerprintCanonicalValue({
    id: workbookInput.id,
    name: workbookInput.name,
    sheets: workbookInput.sheets,
    datePolicy: workbookInput.datePolicy ?? null,
    calculation: workbookInput.calculation ?? null,
    definedNames: workbookInput.definedNames ?? null,
    styleCatalog: workbookInput.styleCatalog ?? null,
  });
}
const isCellHiddenInSheet = (e, r, t) =>
  !!(
    e.hidden ||
    e.veryHidden ||
    cn(e, "row").some((n) => n.start <= r && n.end >= r) ||
    cn(e, "column").some((n) => n.start <= t && n.end >= t)
  );
function buildReportSourceSheet(reportSnapshot) {
  return {
    id: reportSnapshot.request.sheetId,
    name: reportSnapshot.sheetName,
    rowCount: 1048576,
    colCount: 16384,
    cells: Object.fromEntries(
      reportSnapshot.rows.flatMap((row) =>
        row.cells.map((cell) => [
          cell.address,
          { raw: cell.value, literal: true, style: cell.style },
        ]),
      ),
    ),
  };
}
function buildReportSnapshot(workbook, r, t = {}) {
  const request = validateObjectReportRequest(r),
    analysisPermissions = normalizeAnalysisPermissions(workbook, t),
    workbookInput = normalizeWorkbookInput(workbook),
    s = workbookInput.sheets.find((sheet) => sheet.id === request.sheetId);
  if (!s || !analysisPermissions.allowedSheetIds.includes(s.id))
    throw Error("The report source is outside the authorized host context.");
  if (!analysisPermissions.includeHidden && (s.hidden || s.veryHidden))
    throw Error("The report source worksheet is hidden.");
  if (
    workbookInput.calculation?.mode === "manual" ||
    workbookInput.calculation?.iterative
  )
    throw Error(
      "Report capture requires automatic, non-iterative calculation. Recalculate and retry.",
    );
  const d =
    request.kind === "object"
      ? structuredClone(
          s.objects?.find((object) => object.id === request.objectId),
        )
      : undefined;
  if (request.kind === "object" && !d)
    throw Error("The source drawing no longer exists.");
  if (
    d &&
    parseReportRange(d.kind === "chart" ? d.range : d.anchor) !== request.range
  )
    throw Error("The report range must match the drawing source.");
  if (d && d.kind !== "chart" && request.filter)
    throw Error("Pictures and shapes do not use population filters.");
  const rangeAddress = parseRangeAddress(request.range),
    columns = Array.from(
      { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
      (j, z) => rangeAddress.firstCol + z,
    );
  let c,
    h = [];
  if (d && d.kind !== "chart") {
    if (
      !analysisPermissions.includeHidden &&
      isCellHiddenInSheet(s, rangeAddress.firstRow, rangeAddress.firstCol)
    )
      throw Error(
        "The drawing anchor is hidden. Explicit authorization is required.",
      );
    ((c = [[null]]), (h = [{ sheetId: s.id, address: d.anchor }]));
  } else {
    const sourceRectangles = readSourceRectangles(
      workbookInput,
      [{ sheetId: s.id, range: request.range }],
      t,
    );
    ((c = sourceRectangles.grids[0].values),
      (h = sourceRectangles.dependencyCells));
  }
  const f = new Map();
  c.forEach((j, z) =>
    j.forEach((V, D) =>
      f.set(
        `${columnIndexToLetters(rangeAddress.firstCol + D)}${rangeAddress.firstRow + z + 1}`,
        V,
      ),
    ),
  );
  const p = (j, z) => {
      if (j !== s.id || !f.has(z))
        throw Error("Report calculation is outside the captured range.");
      return f.get(z);
    },
    m = request.filter
      ? applyFilterConditions(workbookInput, s.id, request.filter, p)
      : undefined;
  m && (request.filter = m.context.definition);
  const v = c.map((j, z) => rangeAddress.firstRow + z + 1),
    g =
      request.population === "visible"
        ? v.filter(
            (j) =>
              !cn(s, "row").some((z) => z.start <= j - 1 && z.end >= j - 1) &&
              (!m ||
                j === rangeAddress.firstRow + 1 ||
                m.context.matchingSourceRows.includes(j)),
          )
        : v;
  let b = g;
  if (d?.kind === "chart") {
    const j = d.chartOptions?.includedRows,
      z = g.filter(
        (D) =>
          (d.headers !== false && D === rangeAddress.firstRow + 1) ||
          !j ||
          j.includes(D),
      );
    d.chartOptions = {
      ...d.chartOptions,
      includedRows: z.filter(
        (D) => !(d.headers !== false && D === rangeAddress.firstRow + 1),
      ),
    };
    const V = inspectChartSource(s, d, p);
    (formatDiagnosticsMessage(V), (b = z));
  }
  if (!b.length) throw Error("No rows remain in the report population.");
  const warnings = [
    "Fixed values and formatting captured locally; links refresh only after host authorization and explicit review. No network access or remote refresh is performed.",
    "Table images use readable, wrapped cell text with portable fonts. Merges are refused; conditional formatting, layout fidelity and native Office OLE links are not represented.",
  ];
  ((s.analysisReport ||
    s.dataModelReport ||
    s.pivots?.length ||
    s.queries?.length) &&
    warnings.push(
      "This worksheet contains generated report output. Captured cells are historical source values; verify the underlying analysis, model, pivot or query freshness separately.",
    ),
    request.filter &&
      warnings.push(
        request.population === "full"
          ? "Filter criteria are recorded, but full population was selected: filtered rows are included."
          : "Only matching visible rows are retained; excluded rows have locators but no values in this snapshot.",
      ),
    d?.kind === "chart" &&
      warnings.push(
        "Chart row selections are preserved. Explicitly excluded chart rows are omitted from retained cells.",
      ));
  const w = {
    engine: reportOutputFormatVersion,
    id: crypto.randomUUID(),
    request: request,
    workbookId: workbookInput.id,
    workbookName: workbookInput.name,
    revision: workbookInput.revision,
    sheetName: s.name,
    calculatedAt: new Date().toISOString(),
    sourceHash: captureAnalysisWorkbookState(workbookInput),
    includeHidden: analysisPermissions.includeHidden,
    authorizedSheetIds: [...new Set([s.id, ...h.map((j) => j.sheetId)])].sort(),
    sourceCells: h.map((j) => ({ sheetId: j.sheetId, address: j.address })),
    rows: b.map((row) => ({
      row: row,
      cells: columns.map((column, V) => {
        const address = `${columnIndexToLetters(column)}${row}`,
          B = c[row - rangeAddress.firstRow - 1][V],
          cell = s.cells[address];
        return {
          address: address,
          value: B,
          text: formatNumberWithPattern(
            B,
            cell,
            workbookInput.datePolicy?.dateSystem,
          ),
          ...(cell?.style
            ? { style: resolveEffectiveCellStyle(cell.style) }
            : {}),
        };
      }),
    })),
    columns: columns,
    widths: columns.map((column) => s.columnWidths?.[String(column)] ?? 84),
    excludedRows: v.filter((j) => !b.includes(j)),
    ...(d ? { object: d } : {}),
    warnings: warnings,
    hash: "",
  };
  return (
    (w.hash = hashRecordExcludingHash(w)),
    deepFreezeValue(parseReportSnapshot(w))
  );
}
const qT = new WeakMap();
function refreshReportOutput(workbook, request, t = {}, refreshId) {
  if (refreshId) {
    const d = workbook.reportOutputs?.find(
      (reportOutput) => reportOutput.id === refreshId,
    );
    if (!d || d.snapshot.request.mode !== "linked")
      throw Error("Only an existing linked output can be refreshed.");
    if (
      fingerprintCanonicalValue(validateObjectReportRequest(request)) !==
      fingerprintCanonicalValue(d.snapshot.request)
    )
      throw Error(
        "Refresh uses the saved definition. Publish a new output to change it.",
      );
    if (d.snapshot.workbookId !== workbook.id)
      throw Error(
        "This link belongs to a different workbook. Publish a new output after reviewing the source.",
      );
  }
  const reportSnapshot = buildReportSnapshot(workbook, request, t),
    a = {
      request: reportSnapshot.request,
      snapshot: reportSnapshot,
      workbookHash: fingerprintCanonicalValue(workbook),
      contextHash: fingerprintCanonicalValue(
        normalizeAnalysisPermissions(workbook, t),
      ),
      ...(refreshId ? { refreshId: refreshId } : {}),
    };
  return (qT.set(a, fingerprintCanonicalValue(a)), a);
}
function publishReportOutput(workbook, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    qT.get(r) !== fingerprintCanonicalValue(r) ||
      r.workbookHash !== fingerprintCanonicalValue(workbook) ||
      r.contextHash !==
        fingerprintCanonicalValue(normalizeAnalysisPermissions(workbook, t)))
  )
    throw Error(
      "The preview, workbook or permissions changed. Preview the report again.",
    );
  const workbookInput = normalizeWorkbookInput(workbook),
    reportOutputs = [...(workbookInput.reportOutputs ?? [])];
  if (r.refreshId) {
    const s = reportOutputs.findIndex(
      (reportOutput) => reportOutput.id === r.refreshId,
    );
    if (s < 0) throw Error("The linked output was removed.");
    reportOutputs[s] = {
      ...reportOutputs[s],
      version: reportOutputs[s].version + 1,
      snapshot: structuredClone(r.snapshot),
    };
  } else {
    if (
      reportOutputs.some(
        (reportOutput) =>
          reportOutput.snapshot.request.name.toLowerCase() ===
          r.request.name.toLowerCase(),
      )
    )
      throw Error("A report with this name already exists.");
    reportOutputs.push({
      id: crypto.randomUUID(),
      version: 1,
      snapshot: structuredClone(r.snapshot),
    });
  }
  const a = normalizeWorkbookInput({
    ...workbookInput,
    reportOutputs: reportOutputs,
  });
  return (qT.delete(r), a);
}
function removeReportOutput(workbook, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(workbook);
  const n = workbook.reportOutputs?.find(
    (reportOutput) => reportOutput.id === r,
  );
  if (!n) throw Error("Report output not found.");
  return (
    validateReportSourceScope(workbook, n.snapshot, t),
    normalizeWorkbookInput({
      ...workbook,
      reportOutputs: workbook.reportOutputs.filter(
        (reportOutput) => reportOutput.id !== r,
      ),
    })
  );
}
function validateReportSourceScope(workbook, snapshot, t = {}) {
  const analysisPermissions = normalizeAnalysisPermissions(workbook, t);
  if (
    snapshot.workbookId !== workbook.id ||
    snapshot.authorizedSheetIds.some(
      (authorizedSheetId) =>
        !analysisPermissions.allowedSheetIds.includes(authorizedSheetId),
    ) ||
    (snapshot.includeHidden && !analysisPermissions.includeHidden)
  )
    throw Error("The report is outside the current authorized source context.");
  for (const sourceCell of snapshot.sourceCells) {
    const a = workbook.sheets.find((sheet) => sheet.id === sourceCell.sheetId),
      s = _t(sourceCell.address);
    if (
      !a ||
      !analysisPermissions.allowedSheetIds.includes(sourceCell.sheetId) ||
      (!analysisPermissions.includeHidden &&
        isCellHiddenInSheet(a, s.row, s.col))
    )
      throw Error(
        "A report source cell is missing, hidden or outside the authorized context.",
      );
  }
  for (const authorizedSheetId of snapshot.authorizedSheetIds) {
    const a = workbook.sheets.find((sheet) => sheet.id === authorizedSheetId);
    if (
      !a ||
      (!analysisPermissions.includeHidden && (a.hidden || a.veryHidden))
    )
      throw Error("The report source is missing or hidden.");
  }
}
function describeReportOutputStatus(workbook, r, t = {}) {
  try {
    const o = workbook.reportOutputs?.find(
      (reportOutput) => reportOutput.id === r,
    );
    if (!o) throw Error("Report output not found.");
    validateReportSourceScope(workbook, o.snapshot, t);
    const reportSnapshot = parseReportSnapshot(o.snapshot);
    return reportSnapshot.request.mode === "snapshot"
      ? {
          status: "snapshot",
          message: `Fixed snapshot of revision ${reportSnapshot.revision}. Later edits do not change it.`,
        }
      : captureAnalysisWorkbookState(workbook) !== reportSnapshot.sourceHash
        ? {
            status: "stale",
            message:
              "Source content or settings changed. Preview and review a refresh; the saved snapshot remains unchanged.",
          }
        : {
            status: "current",
            message:
              "Source fingerprint matches. The rendered content remains the saved snapshot; use a trusted read for current values.",
          };
  } catch (error) {
    return {
      status: "unavailable",
      message: error instanceof Error ? error.message : String(error),
    };
  }
}
const escapeHtmlText = (e) =>
    String(e).replace(
      /[&<>"']/g,
      (r) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&apos;",
        })[r],
    ),
  toSafeFileName = (e) =>
    e.replace(/[\x00-\x1f\\/:*?"<>|]/g, "_").slice(0, 100) || "Report",
  formatReportCaption = (e) =>
    `${e.workbookName} · revision ${e.revision} · ${e.sheetName}!${e.request.range} · calculated ${e.calculatedAt}`;
function wrapTextByCharacters(e, r) {
  return e
    .replace(
      /\r\n?/g,
      `
`,
    )
    .split(
      `
`,
    )
    .flatMap((t) => {
      const n = Array.from(t),
        o = [];
      for (let a = 0; a < n.length; a += r) o.push(n.slice(a, a + r).join(""));
      return o.length ? o : [""];
    });
}
function renderReportSvg(e) {
  const reportSnapshot = parseReportSnapshot(e),
    object = reportSnapshot.object;
  if (
    [
      reportSnapshot.request.name,
      reportSnapshot.workbookName,
      reportSnapshot.sheetName,
      object?.label ?? "",
      ...reportSnapshot.rows.flatMap((row) =>
        row.cells.map((cell) => cell.text),
      ),
    ].some((h) =>
      new RegExp(
        "[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\uFFFE\\uFFFF]|\\p{Surrogate}",
        "u",
      ).test(h),
    )
  )
    throw Error(
      "Source text contains characters SVG cannot represent. Clean the text or export JSON/HTML.",
    );
  let width = 640,
    height = 0,
    a = "";
  if (object?.kind === "chart") {
    const reportSourceSheet = buildReportSourceSheet(reportSnapshot),
      report = inspectChartSource(
        reportSourceSheet,
        object,
        (m, v) => reportSourceSheet.cells[v]?.raw ?? null,
      );
    formatDiagnosticsMessage(report);
    const p = c_.renderToStaticMarkup(
      i.jsx(ChartReportSvg, { report: report }),
    );
    ((width = 640),
      (height = object.chart === "sparkline" ? 128 : 380),
      (a = `<svg x="18" y="112" width="${width}" height="${height}" viewBox="${object.chart === "sparkline" ? "0 0 200 40" : "0 0 640 380"}">${p}</svg>`));
  } else if (object?.kind === "picture")
    (inspectPictureDataUrl(object.dataURL),
      (width = Math.max(640, object.width)),
      (height = object.height),
      (a = `<image x="18" y="112" width="${object.width}" height="${height}" href="${escapeHtmlText(object.dataURL)}"/>`));
  else if (object?.kind === "shape") {
    ((width = Math.max(640, object.width)),
      (height = Math.max(80, object.height)));
    const h = `fill="${object.fill}" stroke="${object.stroke}" stroke-width="2"`,
      f =
        object.shape === "ellipse"
          ? `<ellipse cx="${object.width / 2}" cy="${height / 2}" rx="${object.width / 2 - 2}" ry="${height / 2 - 2}" ${h}/>`
          : object.shape === "arrow"
            ? `<path d="M2 20 H${object.width - 30} V2 L${object.width - 2} ${height / 2} L${object.width - 30} ${height - 2} V${height - 20} H2 Z" ${h}/>`
            : object.shape === "text"
              ? ""
              : `<rect x="1" y="1" width="${object.width - 2}" height="${height - 2}" ${h}/>`,
      p = wrapTextByCharacters(
        object.label,
        Math.max(4, Math.floor((object.width - 16) / 8)),
      );
    if (p.length * 18 + 16 > height)
      throw Error(
        "The shape label does not fit the report image. Enlarge the source shape before publishing.",
      );
    a = `<g transform="translate(18 112)">${f}${p.map((m, v) => `<text x="8" y="${24 + v * 18}" font-size="13" fill="#202334">${escapeHtmlText(m)}</text>`).join("")}</g>`;
  } else {
    const h = reportSnapshot.columns.map((column, m) =>
      Math.min(
        600,
        Math.max(
          84,
          reportSnapshot.widths[m],
          ...reportSnapshot.rows.map((row) =>
            typeof row.cells[m].value == "number"
              ? Array.from(row.cells[m].text).length * 8 + 20
              : 84,
          ),
        ),
      ),
    );
    width = Math.max(
      640,
      h.reduce((p, m) => p + m, 0),
    );
    let f = 112;
    for (const row of reportSnapshot.rows) {
      const m = row.cells.map((cell, y) =>
          wrapTextByCharacters(
            cell.text,
            Math.max(1, Math.floor((h[y] - 16) / 8)),
          ),
        ),
        v = Math.max(28, ...m.map((b) => b.length * 18 + 12));
      let g = 18;
      (row.cells.forEach((cell, y) => {
        const w = cell.style ?? {},
          E = h[y];
        ((a += `<rect x="${g}" y="${f}" width="${E}" height="${v}" fill="${escapeHtmlText(w.fill ?? "#ffffff")}" stroke="#ccd6fa"/><text font-family="Courier New,monospace" font-size="13" fill="${escapeHtmlText(w.color ?? "#202334")}"${w.bold ? ' font-weight="bold"' : ""}${w.italic ? ' font-style="italic"' : ""}${w.underline ? ' text-decoration="underline"' : ""}>${m[y].map((A, k) => `<tspan x="${g + 8}" y="${f + 20 + k * 18}">${escapeHtmlText(A)}</tspan>`).join("")}</text>`),
          (g += E));
      }),
        (f += v));
    }
    height = f - 112;
  }
  if (
    ((width += 36),
    (height += 162),
    width > reportOutputLimits.imageWidth ||
      height > reportOutputLimits.imageHeight ||
      width * height > reportOutputLimits.imagePixels)
  )
    throw Error(
      "Report image exceeds 4,096 × 16,384 / 16 megapixels. Narrow the range or export JSON/HTML instead.",
    );
  const s = `${reportSnapshot.request.population} population · ${reportSnapshot.rows.length} retained / ${reportSnapshot.rows.length + reportSnapshot.excludedRows.length} range rows · ${reportSnapshot.request.filter ? "filter criteria recorded" : "no population filter"}`,
    d = wrapTextByCharacters(
      formatReportCaption(reportSnapshot),
      Math.floor((width - 36) / 7.4),
    ),
    u = [
      ...wrapTextByCharacters(
        reportSnapshot.request.name,
        Math.floor((width - 36) / 9),
      ),
      ...d,
      ...wrapTextByCharacters(s, Math.floor((width - 36) / 7.4)),
    ],
    l = Math.max(0, u.length * 18 + 24 - 112);
  if (
    ((height += l),
    (a = `<g transform="translate(0 ${l})">${a}</g>`),
    height > reportOutputLimits.imageHeight ||
      width * height > reportOutputLimits.imagePixels)
  )
    throw Error(
      "Report image exceeds its pixel budget. Narrow the range or export JSON/HTML.",
    );
  const content = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img"><title>${escapeHtmlText(reportSnapshot.request.name)}</title><desc>${escapeHtmlText(formatReportCaption(reportSnapshot))}. ${escapeHtmlText(s)}. Fixed snapshot; no automatic refresh.</desc><metadata>${escapeHtmlText(JSON.stringify({ engine: reportSnapshot.engine, hash: reportSnapshot.hash, request: reportSnapshot.request, workbookId: reportSnapshot.workbookId, revision: reportSnapshot.revision, calculatedAt: reportSnapshot.calculatedAt }))}</metadata><rect width="100%" height="100%" fill="white"/><g font-family="Arial,sans-serif">${u.map((h, f) => `<text x="18" y="${25 + f * 18}" font-size="${f === 0 ? 16 : 12}" fill="#362c63">${escapeHtmlText(h)}</text>`).join("")}${a}<text x="18" y="${height - 17}" font-size="11" fill="#545969">Fixed snapshot · ${escapeHtmlText(reportSnapshot.id)} · full values and context in the JSON export</text></g></svg>`;
  return {
    filename: toSafeFileName(reportSnapshot.request.name) + ".svg",
    mime: "image/svg+xml",
    content: content,
    width: width,
    height: height,
  };
}
function buildReportSvgFile(e) {
  const reportSnapshot = parseReportSnapshot(e);
  return {
    filename: toSafeFileName(reportSnapshot.request.name) + ".json",
    mime: "application/json",
    content: JSON.stringify(reportSnapshot, null, 2),
  };
}
function ReportOutputView({ snapshot, showData = true }) {
  const reportSnapshot = parseReportSnapshot(snapshot);
  let n,
    o = "";
  try {
    n = renderReportSvg(reportSnapshot);
  } catch (error) {
    o = error.message;
  }
  return i.jsxs("figure", {
    className: "lg-report-output-view",
    style: { margin: 0, color: "#202334", fontFamily: "Arial,sans-serif" },
    children: [
      i.jsxs("figcaption", {
        children: [
          i.jsx("strong", { children: reportSnapshot.request.name }),
          i.jsx("p", { children: formatReportCaption(reportSnapshot) }),
          i.jsxs("p", {
            children: [
              "Fixed snapshot · ",
              reportSnapshot.request.population,
              " population · ",
              reportSnapshot.rows.length,
              " ",
              "retained / ",
              reportSnapshot.rows.length + reportSnapshot.excludedRows.length,
              " range rows.",
              " ",
              reportSnapshot.request.filter
                ? "Filter criteria retained in context."
                : "No population filter.",
            ],
          }),
        ],
      }),
      n
        ? i.jsx("img", {
            alt:
              reportSnapshot.request.name +
              " — " +
              formatReportCaption(reportSnapshot),
            src:
              "data:image/svg+xml;charset=utf-8," +
              encodeURIComponent(n.content),
            width: n.width,
            height: n.height,
            style: { maxWidth: "100%", height: "auto", display: "block" },
          })
        : i.jsx("p", { role: "status", children: o }),
      showData &&
        i.jsxs("details", {
          open: !n,
          children: [
            i.jsx("summary", {
              children: "Accessible values and report context",
            }),
            i.jsxs("table", {
              style: { borderCollapse: "collapse" },
              children: [
                i.jsxs("caption", {
                  children: [
                    reportSnapshot.sheetName,
                    "!",
                    reportSnapshot.request.range,
                  ],
                }),
                i.jsx("thead", {
                  children: i.jsxs("tr", {
                    children: [
                      i.jsx("th", { children: "Source row" }),
                      reportSnapshot.columns.map((column) =>
                        i.jsxs(
                          "th",
                          {
                            children: [
                              columnIndexToLetters(column),
                              " · column ",
                              column + 1,
                            ],
                          },
                          column,
                        ),
                      ),
                    ],
                  }),
                }),
                i.jsx("tbody", {
                  children: reportSnapshot.rows.map((row) =>
                    i.jsxs(
                      "tr",
                      {
                        children: [
                          i.jsx("th", { scope: "row", children: row.row }),
                          row.cells.map((cell) =>
                            i.jsx(
                              "td",
                              {
                                style: {
                                  border: "1px solid #ccd6fa",
                                  padding: 8,
                                  whiteSpace: "pre-wrap",
                                },
                                children: cell.text,
                              },
                              cell.address,
                            ),
                          ),
                        ],
                      },
                      row.row,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("p", {
              children: [
                "Excluded source rows: ",
                reportSnapshot.excludedRows.join(", ") || "none",
                ".",
              ],
            }),
            reportSnapshot.request.filter &&
              i.jsx("pre", {
                style: { whiteSpace: "pre-wrap" },
                children: JSON.stringify(
                  reportSnapshot.request.filter,
                  null,
                  2,
                ),
              }),
            reportSnapshot.warnings.map((warning, s) =>
              i.jsx("p", { children: warning }, s),
            ),
          ],
        }),
    ],
  });
}
function buildReportHtmlFile(e) {
  const reportSnapshot = parseReportSnapshot(e),
    t = c_.renderToStaticMarkup(
      i.jsx(ReportOutputView, { snapshot: reportSnapshot }),
    );
  return {
    filename: toSafeFileName(reportSnapshot.request.name) + ".html",
    mime: "text/html",
    content: `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"><title>${escapeHtmlText(reportSnapshot.request.name)}</title></head><body style="padding:18px;overflow:auto">${t}</body></html>`,
  };
}
async function renderReportPngBlob(e) {
  const r = renderReportSvg(e),
    t = URL.createObjectURL(new Blob([r.content], { type: "image/svg+xml" }));
  try {
    const n = new Image();
    await new Promise((d, u) => {
      const l = setTimeout(() => {
        ((n.src = ""), u(Error("Report image rendering timed out.")));
      }, 15e3);
      ((n.onload = () => {
        (clearTimeout(l), d());
      }),
        (n.onerror = () => {
          (clearTimeout(l),
            u(
              Error(
                "The browser could not render this report image. Export SVG or HTML.",
              ),
            ));
        }),
        (n.src = t));
    });
    const o = document.createElement("canvas");
    ((o.width = r.width), (o.height = r.height));
    const a = o.getContext("2d");
    if (!a) throw Error("PNG rendering is unavailable in this browser.");
    a.drawImage(n, 0, 0);
    const blob = await new Promise((d, u) => {
      const l = setTimeout(() => u(Error("PNG encoding timed out.")), 15e3);
      o.toBlob((c) => {
        (clearTimeout(l), c ? d(c) : u(Error("PNG encoding failed.")));
      }, "image/png");
    });
    if (blob.size > 16 * 1024 * 1024)
      throw Error("PNG exceeds 16 MB. Export SVG or narrow the report.");
    return { filename: toSafeFileName(e.request.name) + ".png", blob: blob };
  } finally {
    URL.revokeObjectURL(t);
  }
}

export {
  describeReportOutputStatus,
  publishReportOutput,
  buildReportHtmlFile,
  buildReportSvgFile,
  refreshReportOutput,
  AnalysisSnapshotNotice,
  ReportOutputView,
  removeReportOutput,
  renderReportSvg,
  renderReportPngBlob,
  DataAnalysisDialog,
  validateReportSourceScope,
};
