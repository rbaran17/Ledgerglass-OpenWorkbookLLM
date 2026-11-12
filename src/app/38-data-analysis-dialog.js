import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import c_ from "react-dom/server";
import { Lg, hr } from "./01-core-schemas-and-hashing.js";
import {
  $a,
  Af,
  BR,
  CU,
  FR,
  P1,
  Ta,
  Vv,
  W5,
  Wg,
  Ww,
  _t,
  aR,
  cn,
  ct,
  f0,
  rn,
  ut,
  ys,
} from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import { J$ } from "./12-formula-dependency-inspector.js";
import { LI } from "./31-pivot-explore-and-slicers.js";
import { lm } from "./36-watch-window-and-goal-seek.js";
import { Axe, Cxe, Exe, Oi, aj } from "./37-linear-solver-engine.js";

function kxe(e) {
  var K;
  const { text: r } = Dr(),
    t = ne.useId(),
    n = ne.useRef(null),
    o = e.workbook.sheets.find((te) => te.id === e.sheetId),
    a = o.analysisReport,
    [s, d] = ne.useState(
      (a == null ? void 0 : a.request.kind) ?? e.initialKind ?? "statistics",
    ),
    [u, l] = ne.useState("Setup"),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState((a == null ? void 0 : a.includeHidden) ?? false),
    [g, b] = ne.useState(0),
    y = () => {
      let te = 1;
      for (
        ;
        e.workbook.sheets.some(
          (ae) => ae.name.toLowerCase() === `analysis ${te}`,
        );
      )
        te++;
      return `Analysis ${te}`;
    },
    [w, E] = ne.useState(y),
    [A, k] = ne.useState(() => {
      var te;
      return (a == null ? void 0 : a.request.kind) === "solver"
        ? a.request
        : {
            kind: "solver",
            outputName: w,
            objective: {
              sheetId: e.sheetId,
              address:
                ((te = Object.entries(o.cells).find(
                  ([, ae]) =>
                    !ae.literal &&
                    typeof ae.raw == "string" &&
                    ae.raw.startsWith("="),
                )) == null
                  ? void 0
                  : te[0]) ?? "C2",
            },
            direction: "maximize",
            variables: Object.entries(o.cells)
              .filter(([, ae]) => typeof ae.raw == "number")
              .slice(0, 2)
              .map(([ae]) => ({
                cell: { sheetId: e.sheetId, address: ae },
                lower: 0,
                upper: 100,
                domain: "continuous",
              })),
            constraints: [],
            tolerance: 1e-8,
            maxCandidates: 5e4,
          };
    }),
    [F, I] = ne.useState(() =>
      (a == null ? void 0 : a.request.kind) === "statistics"
        ? a.request
        : {
            kind: "statistics",
            outputName: w,
            sheetId: e.sheetId,
            range: e.range.includes(":") ? e.range : Ta(o),
            method: "descriptive",
            missing: "reject",
          },
    ),
    [O, j] = ne.useState(
      ((K = F.bins) == null ? void 0 : K.join(", ")) ?? "10, 20, 30",
    ),
    [z, V] = ne.useState(0),
    [D, B] = ne.useState(-1),
    [_, C] = ne.useState(null),
    [T, N] = ne.useState(""),
    R = JSON.stringify([e.readOnly, e.allowedSheetIds, !!e.allowHidden && m]),
    L = !!f && D === z && _ === e.workbook && T === R,
    $ = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      includeHidden: !!e.allowHidden && m,
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
      (te) =>
        !te.analysisReport &&
        !te.dataModelReport &&
        (!e.allowedSheetIds || e.allowedSheetIds.includes(te.id)) &&
        ($.includeHidden || (!te.hidden && !te.veryHidden)),
    ),
    X$1 = (te, ae, oe) =>
      i.jsx("select", {
        "aria-label": r(oe),
        value: te,
        onChange: (de) => ae(de.target.value),
        children: Q.map((de) =>
          i.jsx("option", { value: de.id, children: de.name }, de.id),
        ),
      }),
    J = (te, ae, oe) =>
      i.jsx(ci, {
        "aria-label": r(te),
        value: ae === void 0 || !Number.isFinite(ae) ? "" : ae,
        onValueChange: (de) => oe(aj(de)),
      });
  (ne.useEffect(() => {
    var ae, oe;
    const te = document.activeElement;
    return (
      (oe = (ae = n.current) == null ? void 0 : ae.querySelector("button")) ==
        null || oe.focus(),
      () => {
        te != null && te.isConnected && te.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      var ae;
      const te =
        (ae = n.current) == null ? void 0 : ae.querySelector(".lg-model-body");
      te && (te.scrollTop = 0);
    }, [u, f]));
  const q = () => {
      (p(null), h(""));
      try {
        const te =
            s === "solver"
              ? { ...A, outputName: w }
              : {
                  ...F,
                  outputName: w,
                  ...(F.method === "histogram"
                    ? { bins: O.split(",").map(aj) }
                    : {}),
                },
          ae = Exe(e.workbook, te, $);
        (p(ae), B(z), C(e.workbook), N(R), l("Results"), b(0));
      } catch (te) {
        h(te instanceof Error ? te.message : "Analysis failed.");
      }
    },
    W = (te) => {
      try {
        if (!L || !f)
          throw Error("Preview the current inputs and permissions first.");
        const ae = Cxe(e.workbook, f, te, $);
        (e.onWorkbookChange(
          ae,
          te === "report"
            ? "Save analysis report"
            : "Apply Solver solution and report",
        ),
          p(null),
          e.onClose());
      } catch (ae) {
        h(ae instanceof Error ? ae.message : "The analysis was not saved.");
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
      onKeyDown: (te) => {
        if (
          (te.key === "Escape" && (te.stopPropagation(), e.onClose()),
          te.key === "Tab")
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
          te.shiftKey && document.activeElement === oe
            ? (te.preventDefault(), de == null || de.focus())
            : !te.shiftKey &&
              document.activeElement === de &&
              (te.preventDefault(), oe == null || oe.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", {
                  className: "eyebrow",
                  children: i.jsx(M, { children: "Data · Analysis" }),
                }),
                i.jsx("h2", {
                  id: t,
                  children: i.jsx(M, {
                    children: s === "solver" ? "Solver" : "Data Analysis",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      s === "solver"
                        ? "Define a bounded linear model, review the result, then choose whether to apply it."
                        : "Choose a statistical method and review the sample before saving a report.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close analysis"),
              onClick: e.onClose,
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("nav", {
          "aria-label": r("Analysis sections"),
          children: [
            ["Setup", "Results"].map((te) =>
              i.jsx(
                "button",
                {
                  "aria-pressed": u === te,
                  onClick: () => l(te),
                  children: i.jsx(M, { children: te }),
                },
                te,
              ),
            ),
            i.jsx("span", {
              className: "lg-analysis-mode",
              children: Oi(
                "Tool",
                i.jsxs("select", {
                  value: s,
                  onChange: (te) => {
                    (d(te.target.value), U(), l("Setup"));
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
            a &&
              i.jsx("p", {
                className: "lg-analysis-hint",
                children: i.jsx(M, {
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
                        Oi(
                          "Output worksheet",
                          i.jsx("input", {
                            value: w,
                            onChange: (te) => {
                              (E(te.target.value), U());
                            },
                          }),
                        ),
                        e.allowHidden &&
                          Oi(
                            "Include hidden source cells",
                            i.jsx("input", {
                              type: "checkbox",
                              checked: m,
                              onChange: (te) => {
                                (v(te.target.checked), U());
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
                                  children: i.jsx(M, { children: "Objective" }),
                                }),
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    Oi(
                                      "Objective worksheet",
                                      X$1(
                                        A.objective.sheetId,
                                        (te) =>
                                          G({
                                            objective: {
                                              ...A.objective,
                                              sheetId: te,
                                            },
                                          }),
                                        "Objective worksheet",
                                      ),
                                    ),
                                    Oi(
                                      "Objective formula cell",
                                      i.jsx("input", {
                                        value: A.objective.address,
                                        onChange: (te) =>
                                          G({
                                            objective: {
                                              ...A.objective,
                                              address: te.target.value
                                                .replace(/\$/g, "")
                                                .toUpperCase(),
                                            },
                                          }),
                                      }),
                                    ),
                                    Oi(
                                      "Direction",
                                      i.jsxs("select", {
                                        value: A.direction,
                                        onChange: (te) =>
                                          G({ direction: te.target.value }),
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
                                      Oi(
                                        "Target value",
                                        J("Target value", A.target, (te) =>
                                          G({ target: te }),
                                        ),
                                      ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(M, {
                                    children: "Changing cells",
                                  }),
                                }),
                                i.jsx("p", {
                                  className: "lg-analysis-hint",
                                  children: i.jsx(M, {
                                    children:
                                      "Up to six numeric inputs. Give every input finite lower and upper bounds. Use continuous inputs together, or integer and binary inputs together.",
                                  }),
                                }),
                                A.variables.map((te, ae) =>
                                  i.jsxs(
                                    "div",
                                    {
                                      className:
                                        "lg-model-controls lg-analysis-row",
                                      children: [
                                        Oi(
                                          `Variable ${ae + 1} worksheet`,
                                          X$1(
                                            te.cell.sheetId,
                                            (oe) =>
                                              G({
                                                variables: A.variables.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          cell: {
                                                            ...de.cell,
                                                            sheetId: oe,
                                                          },
                                                        }
                                                      : de,
                                                ),
                                              }),
                                            `Variable ${ae + 1} worksheet`,
                                          ),
                                        ),
                                        Oi(
                                          `Variable ${ae + 1} cell`,
                                          i.jsx("input", {
                                            value: te.cell.address,
                                            onChange: (oe) =>
                                              G({
                                                variables: A.variables.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          cell: {
                                                            ...de.cell,
                                                            address:
                                                              oe.target.value
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
                                          Oi(
                                            `Variable ${ae + 1} ${oe}`,
                                            J(
                                              `Variable ${ae + 1} ${oe}`,
                                              te[oe],
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
                                        Oi(
                                          `Variable ${ae + 1} domain`,
                                          i.jsxs("select", {
                                            value: te.domain,
                                            onChange: (oe) =>
                                              G({
                                                variables: A.variables.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          domain:
                                                            oe.target.value,
                                                          ...(oe.target
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
                                          "aria-label": r(
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
                                  onClick: () => {
                                    var te;
                                    return G({
                                      variables: [
                                        ...A.variables,
                                        {
                                          cell: {
                                            sheetId:
                                              ((te = Q[0]) == null
                                                ? void 0
                                                : te.id) ?? e.sheetId,
                                            address: "A2",
                                          },
                                          lower: 0,
                                          upper: 100,
                                          domain: "continuous",
                                        },
                                      ],
                                    });
                                  },
                                  children: i.jsx(M, {
                                    children: "Add changing cell",
                                  }),
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(M, {
                                    children: "Constraints",
                                  }),
                                }),
                                A.constraints.map((te, ae) =>
                                  i.jsxs(
                                    "div",
                                    {
                                      className:
                                        "lg-model-controls lg-analysis-row",
                                      children: [
                                        Oi(
                                          `Constraint ${ae + 1} worksheet`,
                                          X$1(
                                            te.left.sheetId,
                                            (oe) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          left: {
                                                            ...de.left,
                                                            sheetId: oe,
                                                          },
                                                        }
                                                      : de,
                                                ),
                                              }),
                                            `Constraint ${ae + 1} worksheet`,
                                          ),
                                        ),
                                        Oi(
                                          `Constraint ${ae + 1} cell`,
                                          i.jsx("input", {
                                            value: te.left.address,
                                            onChange: (oe) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          left: {
                                                            ...de.left,
                                                            address:
                                                              oe.target.value
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
                                        Oi(
                                          `Constraint ${ae + 1} relation`,
                                          i.jsx("select", {
                                            value: te.operator,
                                            onChange: (oe) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          operator:
                                                            oe.target.value,
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
                                        Oi(
                                          `Constraint ${ae + 1} right type`,
                                          i.jsxs("select", {
                                            value:
                                              typeof te.right == "number"
                                                ? "number"
                                                : "cell",
                                            onChange: (oe) =>
                                              G({
                                                constraints: A.constraints.map(
                                                  (de, he) =>
                                                    he === ae
                                                      ? {
                                                          ...de,
                                                          right:
                                                            oe.target.value ===
                                                            "number"
                                                              ? 0
                                                              : {
                                                                  sheetId:
                                                                    te.left
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
                                        typeof te.right == "number"
                                          ? Oi(
                                              `Constraint ${ae + 1} value`,
                                              J(
                                                `Constraint ${ae + 1} value`,
                                                te.right,
                                                (oe) =>
                                                  G({
                                                    constraints:
                                                      A.constraints.map(
                                                        (de, he) =>
                                                          he === ae
                                                            ? {
                                                                ...de,
                                                                right: oe,
                                                              }
                                                            : de,
                                                      ),
                                                  }),
                                              ),
                                            )
                                          : i.jsxs(i.Fragment, {
                                              children: [
                                                Oi(
                                                  `Constraint ${ae + 1} right worksheet`,
                                                  X$1(
                                                    te.right.sheetId,
                                                    (oe) =>
                                                      G({
                                                        constraints:
                                                          A.constraints.map(
                                                            (de, he) =>
                                                              he === ae
                                                                ? {
                                                                    ...de,
                                                                    right: {
                                                                      ...te.right,
                                                                      sheetId:
                                                                        oe,
                                                                    },
                                                                  }
                                                                : de,
                                                          ),
                                                      }),
                                                    `Constraint ${ae + 1} right worksheet`,
                                                  ),
                                                ),
                                                Oi(
                                                  `Constraint ${ae + 1} right cell`,
                                                  i.jsx("input", {
                                                    value: te.right.address,
                                                    onChange: (oe) =>
                                                      G({
                                                        constraints:
                                                          A.constraints.map(
                                                            (de, he) =>
                                                              he === ae
                                                                ? {
                                                                    ...de,
                                                                    right: {
                                                                      ...te.right,
                                                                      address:
                                                                        oe.target.value
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
                                          "aria-label": r(
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
                                  children: i.jsx(M, {
                                    children: "Add constraint",
                                  }),
                                }),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "lg-model-controls",
                              children: [
                                Oi(
                                  "Feasibility tolerance",
                                  J(
                                    "Feasibility tolerance",
                                    A.tolerance,
                                    (te) => G({ tolerance: te }),
                                  ),
                                ),
                                Oi(
                                  "Maximum candidates",
                                  J(
                                    "Maximum candidates",
                                    A.maxCandidates,
                                    (te) => G({ maxCandidates: te }),
                                  ),
                                ),
                              ],
                            }),
                            i.jsx("p", {
                              className: "lg-analysis-hint",
                              children: i.jsx(M, {
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
                                  children: i.jsx(M, {
                                    children: "Source and method",
                                  }),
                                }),
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    Oi(
                                      "Source worksheet",
                                      X$1(
                                        F.sheetId,
                                        (te) => H({ sheetId: te }),
                                        "Source worksheet",
                                      ),
                                    ),
                                    Oi(
                                      "Headed source range",
                                      i.jsx("input", {
                                        value: F.range,
                                        onChange: (te) =>
                                          H({
                                            range: te.target.value
                                              .replace(/\$/g, "")
                                              .toUpperCase(),
                                          }),
                                      }),
                                    ),
                                    Oi(
                                      "Analysis method",
                                      i.jsx("select", {
                                        value: F.method,
                                        onChange: (te) =>
                                          H({ method: te.target.value }),
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
                                    Oi(
                                      "Incomplete rows",
                                      i.jsxs("select", {
                                        value: F.missing,
                                        onChange: (te) =>
                                          H({ missing: te.target.value }),
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
                                  children: i.jsx(M, {
                                    children:
                                      "Include text column headers and numeric data. Up to 500 data rows, six columns and 2,000 cells. Formula inputs are limited to 50 deterministic scalar formulas with bounded dependencies.",
                                  }),
                                }),
                              ],
                            }),
                            F.method === "regression" &&
                              i.jsx("p", {
                                children: i.jsx(M, {
                                  children:
                                    "Put the response Y in the first column and predictors X in the remaining columns. The model includes an intercept, coefficient confidence intervals and residuals.",
                                }),
                              }),
                            F.method === "histogram" &&
                              Oi(
                                "Bin upper bounds",
                                i.jsx("input", {
                                  value: O,
                                  onChange: (te) => {
                                    (j(te.target.value), U());
                                  },
                                }),
                              ),
                            F.method === "distribution" &&
                              i.jsxs("fieldset", {
                                children: [
                                  i.jsx("legend", {
                                    children: i.jsx(M, {
                                      children: "Distribution parameters",
                                    }),
                                  }),
                                  i.jsxs("div", {
                                    className: "lg-model-controls",
                                    children: [
                                      Oi(
                                        "Distribution",
                                        i.jsxs("select", {
                                          value: F.distribution ?? "normal",
                                          onChange: (te) =>
                                            H({
                                              distribution: te.target.value,
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
                                              Oi(
                                                "Mean",
                                                J("Mean", F.mean, (te) =>
                                                  H({ mean: te }),
                                                ),
                                              ),
                                              Oi(
                                                "Standard deviation",
                                                J(
                                                  "Standard deviation",
                                                  F.standardDeviation,
                                                  (te) =>
                                                    H({
                                                      standardDeviation: te,
                                                    }),
                                                ),
                                              ),
                                            ],
                                          })
                                        : F.distribution === "binomial"
                                          ? i.jsxs(i.Fragment, {
                                              children: [
                                                Oi(
                                                  "Trials",
                                                  J("Trials", F.trials, (te) =>
                                                    H({ trials: te }),
                                                  ),
                                                ),
                                                Oi(
                                                  "Success probability",
                                                  J(
                                                    "Success probability",
                                                    F.probability,
                                                    (te) =>
                                                      H({ probability: te }),
                                                  ),
                                                ),
                                              ],
                                            })
                                          : Oi(
                                              "Rate (lambda)",
                                              J(
                                                "Rate (lambda)",
                                                F.lambda,
                                                (te) => H({ lambda: te }),
                                              ),
                                            ),
                                    ],
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(M, {
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
                                    children: i.jsx(M, {
                                      children: "Forecast settings",
                                    }),
                                  }),
                                  i.jsxs("div", {
                                    className: "lg-model-controls",
                                    children: [
                                      Oi(
                                        "Forecast method",
                                        i.jsxs("select", {
                                          value: F.forecast ?? "linear",
                                          onChange: (te) =>
                                            H({ forecast: te.target.value }),
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
                                      Oi(
                                        "Forecast periods",
                                        J(
                                          "Forecast periods",
                                          F.horizon ?? 6,
                                          (te) => H({ horizon: te }),
                                        ),
                                      ),
                                      F.forecast === "moving-average" &&
                                        Oi(
                                          "Moving average window",
                                          J(
                                            "Moving average window",
                                            F.window ?? 3,
                                            (te) => H({ window: te }),
                                          ),
                                        ),
                                      F.forecast === "exponential" &&
                                        Oi(
                                          "Smoothing alpha",
                                          J(
                                            "Smoothing alpha",
                                            F.alpha ?? 0.3,
                                            (te) => H({ alpha: te }),
                                          ),
                                        ),
                                    ],
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(M, {
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
                              f.result.objective !== void 0 &&
                                i.jsxs("b", {
                                  children: [
                                    "Objective:",
                                    " ",
                                    f.result.objective.toLocaleString(void 0, {
                                      maximumFractionDigits: 10,
                                    }),
                                  ],
                                }),
                            ],
                          }),
                          !L &&
                            i.jsx("p", {
                              role: "alert",
                              children: i.jsx(M, {
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
                                  "aria-label": r("Analysis results"),
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
                                                                void 0,
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
                                      children: i.jsx(M, {
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
                                      children: i.jsx(M, {
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
                                children: i.jsx(M, {
                                  children:
                                    "Method, assumptions and reference checks",
                                }),
                              }),
                              i.jsx("ul", {
                                children: f.result.assumptions.map((te, ae) =>
                                  i.jsx("li", { children: te }, ae),
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
                              f.result.exclusions.map((te) =>
                                i.jsxs(
                                  "p",
                                  { children: [te.locator, " — ", te.reason] },
                                  te.row,
                                ),
                              ),
                            ],
                          }),
                        ],
                      })
                    : i.jsx("p", {
                        children: i.jsx(M, {
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
              children: i.jsx(M, {
                children:
                  "Local computation · MIT engines · Preview leaves source cells unchanged",
              }),
            }),
            i.jsx("button", {
              onClick: q,
              children: i.jsx(M, { children: "Preview analysis" }),
            }),
            i.jsx("button", {
              disabled: e.readOnly || !L || !(f != null && f.result.success),
              onClick: () => W("report"),
              children: i.jsx(M, { children: "Save report" }),
            }),
            s === "solver" &&
              i.jsx("button", {
                className: "primary",
                disabled: e.readOnly || !L || !(f != null && f.result.success),
                onClick: () => W("apply-solution"),
                children: i.jsx(M, { children: "Apply solution and report" }),
              }),
          ],
        }),
      ],
    }),
  });
}
function Sxe({ workbook: e, sheetId: r, onReview: t }) {
  var a;
  const n =
    (a = e.sheets.find((s) => s.id === r)) == null ? void 0 : a.analysisReport;
  if (!n) return null;
  const o = Axe(e, r, { includeHidden: n.includeHidden });
  return i.jsxs("div", {
    className: "lg-notice",
    role: "status",
    children: [
      i.jsxs("strong", { children: ["Analysis snapshot · ", o.status] }),
      i.jsx("span", { children: o.message }),
      i.jsx("button", {
        onClick: t,
        children: i.jsx(M, { children: "Review analysis" }),
      }),
    ],
  });
}
function yW(e) {
  return hr({
    id: e.id,
    name: e.name,
    sheets: e.sheets,
    datePolicy: e.datePolicy ?? null,
    calculation: e.calculation ?? null,
    definedNames: e.definedNames ?? null,
    styleCatalog: e.styleCatalog ?? null,
  });
}
const wW = (e, r, t) =>
  !!(
    e.hidden ||
    e.veryHidden ||
    cn(e, "row").some((n) => n.start <= r && n.end >= r) ||
    cn(e, "column").some((n) => n.start <= t && n.end >= t)
  );
function Txe(e) {
  return {
    id: e.request.sheetId,
    name: e.sheetName,
    rowCount: 1048576,
    colCount: 16384,
    cells: Object.fromEntries(
      e.rows.flatMap((r) =>
        r.cells.map((t) => [
          t.address,
          { raw: t.value, literal: true, style: t.style },
        ]),
      ),
    ),
  };
}
function Nxe(e, r, t = {}) {
  var E, A, k, F, I, O;
  const n = BR(r),
    o = lm(e, t),
    a = rn(e),
    s = a.sheets.find((j) => j.id === n.sheetId);
  if (!s || !o.allowedSheetIds.includes(s.id))
    throw Error("The report source is outside the authorized host context.");
  if (!o.includeHidden && (s.hidden || s.veryHidden))
    throw Error("The report source worksheet is hidden.");
  if (
    ((E = a.calculation) == null ? void 0 : E.mode) === "manual" ||
    ((A = a.calculation) != null && A.iterative)
  )
    throw Error(
      "Report capture requires automatic, non-iterative calculation. Recalculate and retry.",
    );
  const d =
    n.kind === "object"
      ? structuredClone(
          (k = s.objects) == null ? void 0 : k.find((j) => j.id === n.objectId),
        )
      : void 0;
  if (n.kind === "object" && !d)
    throw Error("The source drawing no longer exists.");
  if (d && FR(d.kind === "chart" ? d.range : d.anchor) !== n.range)
    throw Error("The report range must match the drawing source.");
  if (d && d.kind !== "chart" && n.filter)
    throw Error("Pictures and shapes do not use population filters.");
  const u = ct(n.range),
    l = Array.from(
      { length: u.lastCol - u.firstCol + 1 },
      (j, z) => u.firstCol + z,
    );
  let c,
    h = [];
  if (d && d.kind !== "chart") {
    if (!o.includeHidden && wW(s, u.firstRow, u.firstCol))
      throw Error(
        "The drawing anchor is hidden. Explicit authorization is required.",
      );
    ((c = [[null]]), (h = [{ sheetId: s.id, address: d.anchor }]));
  } else {
    const j = J$(a, [{ sheetId: s.id, range: n.range }], t);
    ((c = j.grids[0].values), (h = j.dependencyCells));
  }
  const f = new Map();
  c.forEach((j, z) =>
    j.forEach((V, D) => f.set(`${ut(u.firstCol + D)}${u.firstRow + z + 1}`, V)),
  );
  const p = (j, z) => {
      if (j !== s.id || !f.has(z))
        throw Error("Report calculation is outside the captured range.");
      return f.get(z);
    },
    m = n.filter ? f0(a, s.id, n.filter, p) : void 0;
  m && (n.filter = m.context.definition);
  const v = c.map((j, z) => u.firstRow + z + 1),
    g =
      n.population === "visible"
        ? v.filter(
            (j) =>
              !cn(s, "row").some((z) => z.start <= j - 1 && z.end >= j - 1) &&
              (!m ||
                j === u.firstRow + 1 ||
                m.context.matchingSourceRows.includes(j)),
          )
        : v;
  let b = g;
  if ((d == null ? void 0 : d.kind) === "chart") {
    const j = (F = d.chartOptions) == null ? void 0 : F.includedRows,
      z = g.filter(
        (D) =>
          (d.headers !== false && D === u.firstRow + 1) || !j || j.includes(D),
      );
    d.chartOptions = {
      ...d.chartOptions,
      includedRows: z.filter(
        (D) => !(d.headers !== false && D === u.firstRow + 1),
      ),
    };
    const V = Ww(s, d, p);
    (aR(V), (b = z));
  }
  if (!b.length) throw Error("No rows remain in the report population.");
  const y = [
    "Fixed values and formatting captured locally; links refresh only after host authorization and explicit review. No network access or remote refresh is performed.",
    "Table images use readable, wrapped cell text with portable fonts. Merges are refused; conditional formatting, layout fidelity and native Office OLE links are not represented.",
  ];
  ((s.analysisReport ||
    s.dataModelReport ||
    ((I = s.pivots) != null && I.length) ||
    ((O = s.queries) != null && O.length)) &&
    y.push(
      "This worksheet contains generated report output. Captured cells are historical source values; verify the underlying analysis, model, pivot or query freshness separately.",
    ),
    n.filter &&
      y.push(
        n.population === "full"
          ? "Filter criteria are recorded, but full population was selected: filtered rows are included."
          : "Only matching visible rows are retained; excluded rows have locators but no values in this snapshot.",
      ),
    (d == null ? void 0 : d.kind) === "chart" &&
      y.push(
        "Chart row selections are preserved. Explicitly excluded chart rows are omitted from retained cells.",
      ));
  const w = {
    engine: CU,
    id: crypto.randomUUID(),
    request: n,
    workbookId: a.id,
    workbookName: a.name,
    revision: a.revision,
    sheetName: s.name,
    calculatedAt: new Date().toISOString(),
    sourceHash: yW(a),
    includeHidden: o.includeHidden,
    authorizedSheetIds: [...new Set([s.id, ...h.map((j) => j.sheetId)])].sort(),
    sourceCells: h.map((j) => ({ sheetId: j.sheetId, address: j.address })),
    rows: b.map((j) => ({
      row: j,
      cells: l.map((z, V) => {
        var C;
        const D = `${ut(z)}${j}`,
          B = c[j - u.firstRow - 1][V],
          _ = s.cells[D];
        return {
          address: D,
          value: B,
          text: Af(B, _, (C = a.datePolicy) == null ? void 0 : C.dateSystem),
          ...(_ != null && _.style ? { style: ys(_.style) } : {}),
        };
      }),
    })),
    columns: l,
    widths: l.map((j) => {
      var z;
      return ((z = s.columnWidths) == null ? void 0 : z[String(j)]) ?? 84;
    }),
    excludedRows: v.filter((j) => !b.includes(j)),
    ...(d ? { object: d } : {}),
    warnings: y,
    hash: "",
  };
  return ((w.hash = W5(w)), Lg(Wg(w)));
}
const qT = new WeakMap();
function Rxe(e, r, t = {}, n) {
  var s;
  if (n) {
    const d =
      (s = e.reportOutputs) == null ? void 0 : s.find((u) => u.id === n);
    if (!d || d.snapshot.request.mode !== "linked")
      throw Error("Only an existing linked output can be refreshed.");
    if (hr(BR(r)) !== hr(d.snapshot.request))
      throw Error(
        "Refresh uses the saved definition. Publish a new output to change it.",
      );
    if (d.snapshot.workbookId !== e.id)
      throw Error(
        "This link belongs to a different workbook. Publish a new output after reviewing the source.",
      );
  }
  const o = Nxe(e, r, t),
    a = {
      request: o.request,
      snapshot: o,
      workbookHash: hr(e),
      contextHash: hr(lm(e, t)),
      ...(n ? { refreshId: n } : {}),
    };
  return (qT.set(a, hr(a)), a);
}
function Ixe(e, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (
    ($a(e),
    qT.get(r) !== hr(r) ||
      r.workbookHash !== hr(e) ||
      r.contextHash !== hr(lm(e, t)))
  )
    throw Error(
      "The preview, workbook or permissions changed. Preview the report again.",
    );
  const n = rn(e),
    o = [...(n.reportOutputs ?? [])];
  if (r.refreshId) {
    const s = o.findIndex((d) => d.id === r.refreshId);
    if (s < 0) throw Error("The linked output was removed.");
    o[s] = {
      ...o[s],
      version: o[s].version + 1,
      snapshot: structuredClone(r.snapshot),
    };
  } else {
    if (
      o.some(
        (s) =>
          s.snapshot.request.name.toLowerCase() ===
          r.request.name.toLowerCase(),
      )
    )
      throw Error("A report with this name already exists.");
    o.push({
      id: crypto.randomUUID(),
      version: 1,
      snapshot: structuredClone(r.snapshot),
    });
  }
  const a = rn({ ...n, reportOutputs: o });
  return (qT.delete(r), a);
}
function _xe(e, r, t = {}) {
  var o;
  if (t.readOnly) throw Error("This workbook is read-only.");
  $a(e);
  const n = (o = e.reportOutputs) == null ? void 0 : o.find((a) => a.id === r);
  if (!n) throw Error("Report output not found.");
  return (
    m_(e, n.snapshot, t),
    rn({ ...e, reportOutputs: e.reportOutputs.filter((a) => a.id !== r) })
  );
}
function m_(e, r, t = {}) {
  const n = lm(e, t);
  if (
    r.workbookId !== e.id ||
    r.authorizedSheetIds.some((o) => !n.allowedSheetIds.includes(o)) ||
    (r.includeHidden && !n.includeHidden)
  )
    throw Error("The report is outside the current authorized source context.");
  for (const o of r.sourceCells) {
    const a = e.sheets.find((d) => d.id === o.sheetId),
      s = _t(o.address);
    if (
      !a ||
      !n.allowedSheetIds.includes(o.sheetId) ||
      (!n.includeHidden && wW(a, s.row, s.col))
    )
      throw Error(
        "A report source cell is missing, hidden or outside the authorized context.",
      );
  }
  for (const o of r.authorizedSheetIds) {
    const a = e.sheets.find((s) => s.id === o);
    if (!a || (!n.includeHidden && (a.hidden || a.veryHidden)))
      throw Error("The report source is missing or hidden.");
  }
}
function Dxe(e, r, t = {}) {
  var n;
  try {
    const o =
      (n = e.reportOutputs) == null ? void 0 : n.find((s) => s.id === r);
    if (!o) throw Error("Report output not found.");
    m_(e, o.snapshot, t);
    const a = Wg(o.snapshot);
    return a.request.mode === "snapshot"
      ? {
          status: "snapshot",
          message: `Fixed snapshot of revision ${a.revision}. Later edits do not change it.`,
        }
      : yW(e) !== a.sourceHash
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
  } catch (o) {
    return {
      status: "unavailable",
      message: o instanceof Error ? o.message : String(o),
    };
  }
}
const cd = (e) =>
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
  nA = (e) =>
    e.replace(/[\x00-\x1f\\/:*?"<>|]/g, "_").slice(0, 100) || "Report",
  T3 = (e) =>
    `${e.workbookName} · revision ${e.revision} · ${e.sheetName}!${e.request.range} · calculated ${e.calculatedAt}`;
function gy(e, r) {
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
function g_(e) {
  const r = Wg(e),
    t = r.object;
  if (
    [
      r.request.name,
      r.workbookName,
      r.sheetName,
      (t == null ? void 0 : t.label) ?? "",
      ...r.rows.flatMap((h) => h.cells.map((f) => f.text)),
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
  let n = 640,
    o = 0,
    a = "";
  if ((t == null ? void 0 : t.kind) === "chart") {
    const h = Txe(r),
      f = Ww(h, t, (m, v) => {
        var g;
        return ((g = h.cells[v]) == null ? void 0 : g.raw) ?? null;
      });
    aR(f);
    const p = c_.renderToStaticMarkup(i.jsx(LI, { report: f }));
    ((n = 640),
      (o = t.chart === "sparkline" ? 128 : 380),
      (a = `<svg x="18" y="112" width="${n}" height="${o}" viewBox="${t.chart === "sparkline" ? "0 0 200 40" : "0 0 640 380"}">${p}</svg>`));
  } else if ((t == null ? void 0 : t.kind) === "picture")
    (P1(t.dataURL),
      (n = Math.max(640, t.width)),
      (o = t.height),
      (a = `<image x="18" y="112" width="${t.width}" height="${o}" href="${cd(t.dataURL)}"/>`));
  else if ((t == null ? void 0 : t.kind) === "shape") {
    ((n = Math.max(640, t.width)), (o = Math.max(80, t.height)));
    const h = `fill="${t.fill}" stroke="${t.stroke}" stroke-width="2"`,
      f =
        t.shape === "ellipse"
          ? `<ellipse cx="${t.width / 2}" cy="${o / 2}" rx="${t.width / 2 - 2}" ry="${o / 2 - 2}" ${h}/>`
          : t.shape === "arrow"
            ? `<path d="M2 20 H${t.width - 30} V2 L${t.width - 2} ${o / 2} L${t.width - 30} ${o - 2} V${o - 20} H2 Z" ${h}/>`
            : t.shape === "text"
              ? ""
              : `<rect x="1" y="1" width="${t.width - 2}" height="${o - 2}" ${h}/>`,
      p = gy(t.label, Math.max(4, Math.floor((t.width - 16) / 8)));
    if (p.length * 18 + 16 > o)
      throw Error(
        "The shape label does not fit the report image. Enlarge the source shape before publishing.",
      );
    a = `<g transform="translate(18 112)">${f}${p.map((m, v) => `<text x="8" y="${24 + v * 18}" font-size="13" fill="#202334">${cd(m)}</text>`).join("")}</g>`;
  } else {
    const h = r.columns.map((p, m) =>
      Math.min(
        600,
        Math.max(
          84,
          r.widths[m],
          ...r.rows.map((v) =>
            typeof v.cells[m].value == "number"
              ? Array.from(v.cells[m].text).length * 8 + 20
              : 84,
          ),
        ),
      ),
    );
    n = Math.max(
      640,
      h.reduce((p, m) => p + m, 0),
    );
    let f = 112;
    for (const p of r.rows) {
      const m = p.cells.map((b, y) =>
          gy(b.text, Math.max(1, Math.floor((h[y] - 16) / 8))),
        ),
        v = Math.max(28, ...m.map((b) => b.length * 18 + 12));
      let g = 18;
      (p.cells.forEach((b, y) => {
        const w = b.style ?? {},
          E = h[y];
        ((a += `<rect x="${g}" y="${f}" width="${E}" height="${v}" fill="${cd(w.fill ?? "#ffffff")}" stroke="#ccd6fa"/><text font-family="Courier New,monospace" font-size="13" fill="${cd(w.color ?? "#202334")}"${w.bold ? ' font-weight="bold"' : ""}${w.italic ? ' font-style="italic"' : ""}${w.underline ? ' text-decoration="underline"' : ""}>${m[y].map((A, k) => `<tspan x="${g + 8}" y="${f + 20 + k * 18}">${cd(A)}</tspan>`).join("")}</text>`),
          (g += E));
      }),
        (f += v));
    }
    o = f - 112;
  }
  if (
    ((n += 36),
    (o += 162),
    n > Vv.imageWidth || o > Vv.imageHeight || n * o > Vv.imagePixels)
  )
    throw Error(
      "Report image exceeds 4,096 × 16,384 / 16 megapixels. Narrow the range or export JSON/HTML instead.",
    );
  const s = `${r.request.population} population · ${r.rows.length} retained / ${r.rows.length + r.excludedRows.length} range rows · ${r.request.filter ? "filter criteria recorded" : "no population filter"}`,
    d = gy(T3(r), Math.floor((n - 36) / 7.4)),
    u = [
      ...gy(r.request.name, Math.floor((n - 36) / 9)),
      ...d,
      ...gy(s, Math.floor((n - 36) / 7.4)),
    ],
    l = Math.max(0, u.length * 18 + 24 - 112);
  if (
    ((o += l),
    (a = `<g transform="translate(0 ${l})">${a}</g>`),
    o > Vv.imageHeight || n * o > Vv.imagePixels)
  )
    throw Error(
      "Report image exceeds its pixel budget. Narrow the range or export JSON/HTML.",
    );
  const c = `<svg xmlns="http://www.w3.org/2000/svg" width="${n}" height="${o}" viewBox="0 0 ${n} ${o}" role="img"><title>${cd(r.request.name)}</title><desc>${cd(T3(r))}. ${cd(s)}. Fixed snapshot; no automatic refresh.</desc><metadata>${cd(JSON.stringify({ engine: r.engine, hash: r.hash, request: r.request, workbookId: r.workbookId, revision: r.revision, calculatedAt: r.calculatedAt }))}</metadata><rect width="100%" height="100%" fill="white"/><g font-family="Arial,sans-serif">${u.map((h, f) => `<text x="18" y="${25 + f * 18}" font-size="${f === 0 ? 16 : 12}" fill="#362c63">${cd(h)}</text>`).join("")}${a}<text x="18" y="${o - 17}" font-size="11" fill="#545969">Fixed snapshot · ${cd(r.id)} · full values and context in the JSON export</text></g></svg>`;
  return {
    filename: nA(r.request.name) + ".svg",
    mime: "image/svg+xml",
    content: c,
    width: n,
    height: o,
  };
}
function Oxe(e) {
  const r = Wg(e);
  return {
    filename: nA(r.request.name) + ".json",
    mime: "application/json",
    content: JSON.stringify(r, null, 2),
  };
}
function XT({ snapshot: e, showData: r = true }) {
  const t = Wg(e);
  let n,
    o = "";
  try {
    n = g_(t);
  } catch (a) {
    o = a.message;
  }
  return i.jsxs("figure", {
    className: "lg-report-output-view",
    style: { margin: 0, color: "#202334", fontFamily: "Arial,sans-serif" },
    children: [
      i.jsxs("figcaption", {
        children: [
          i.jsx("strong", { children: t.request.name }),
          i.jsx("p", { children: T3(t) }),
          i.jsxs("p", {
            children: [
              "Fixed snapshot · ",
              t.request.population,
              " population · ",
              t.rows.length,
              " ",
              "retained / ",
              t.rows.length + t.excludedRows.length,
              " range rows.",
              " ",
              t.request.filter
                ? "Filter criteria retained in context."
                : "No population filter.",
            ],
          }),
        ],
      }),
      n
        ? i.jsx("img", {
            alt: t.request.name + " — " + T3(t),
            src:
              "data:image/svg+xml;charset=utf-8," +
              encodeURIComponent(n.content),
            width: n.width,
            height: n.height,
            style: { maxWidth: "100%", height: "auto", display: "block" },
          })
        : i.jsx("p", { role: "status", children: o }),
      r &&
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
                  children: [t.sheetName, "!", t.request.range],
                }),
                i.jsx("thead", {
                  children: i.jsxs("tr", {
                    children: [
                      i.jsx("th", { children: "Source row" }),
                      t.columns.map((a) =>
                        i.jsxs(
                          "th",
                          { children: [ut(a), " · column ", a + 1] },
                          a,
                        ),
                      ),
                    ],
                  }),
                }),
                i.jsx("tbody", {
                  children: t.rows.map((a) =>
                    i.jsxs(
                      "tr",
                      {
                        children: [
                          i.jsx("th", { scope: "row", children: a.row }),
                          a.cells.map((s) =>
                            i.jsx(
                              "td",
                              {
                                style: {
                                  border: "1px solid #ccd6fa",
                                  padding: 8,
                                  whiteSpace: "pre-wrap",
                                },
                                children: s.text,
                              },
                              s.address,
                            ),
                          ),
                        ],
                      },
                      a.row,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("p", {
              children: [
                "Excluded source rows: ",
                t.excludedRows.join(", ") || "none",
                ".",
              ],
            }),
            t.request.filter &&
              i.jsx("pre", {
                style: { whiteSpace: "pre-wrap" },
                children: JSON.stringify(t.request.filter, null, 2),
              }),
            t.warnings.map((a, s) => i.jsx("p", { children: a }, s)),
          ],
        }),
    ],
  });
}
function Mxe(e) {
  const r = Wg(e),
    t = c_.renderToStaticMarkup(i.jsx(XT, { snapshot: r }));
  return {
    filename: nA(r.request.name) + ".html",
    mime: "text/html",
    content: `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"><title>${cd(r.request.name)}</title></head><body style="padding:18px;overflow:auto">${t}</body></html>`,
  };
}
async function jxe(e) {
  const r = g_(e),
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
    const s = await new Promise((d, u) => {
      const l = setTimeout(() => u(Error("PNG encoding timed out.")), 15e3);
      o.toBlob((c) => {
        (clearTimeout(l), c ? d(c) : u(Error("PNG encoding failed.")));
      }, "image/png");
    });
    if (s.size > 16 * 1024 * 1024)
      throw Error("PNG exceeds 16 MB. Export SVG or narrow the report.");
    return { filename: nA(e.request.name) + ".png", blob: s };
  } finally {
    URL.revokeObjectURL(t);
  }
}
const sj = ({ filename: e, blob: r }) => {
  const t = URL.createObjectURL(r),
    n = document.createElement("a");
  ((n.href = t),
    (n.download = e),
    n.click(),
    setTimeout(() => URL.revokeObjectURL(t), 1e3));
};

export { Dxe, Ixe, Mxe, Oxe, Rxe, Sxe, XT, _xe, g_, jxe, kxe, m_, sj };
