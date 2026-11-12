import { stringifyJsonValue } from "./42-combine-populations-and-text-columns.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  normalizeDatePolicy,
  dateSystemMaxSerial,
  parseDateText,
} from "./01-schemas-date-functions-and-validation.js";
import {
  findSheetOrThrow,
  calculateWorkbookWithSettings,
  _t,
  parseRangeAddress,
  assertSheetUnprotected,
  matrixResultFunctionNames,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { assertSelectionInAuthorizedSheet } from "./15-data-model-dialog.js";
import {
  commitApprovedTransformation,
  verifyTransformationPreview,
  selectExceptionFreeRows,
  buildTransformationPreview,
} from "./16-transformation-recipe-engine.js";
import {
  textCleanupStepLabels,
  describeSplitFieldStep,
  parseFieldPositions,
} from "./43-text-cleanup.js";

function CleanAndReshapeDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(false),
    s = ne.useRef([e.sheetId, e.range]),
    [d, u] = ne.useState(e.initialRequest?.sourceSheetId ?? e.sheetId),
    [l, c] = ne.useState(e.initialRequest?.sourceRange ?? e.range),
    [h, f] = ne.useState(e.initialRequest?.header ?? true),
    [p, m] = ne.useState(e.initialRequest?.outputName ?? "Reshaped fields"),
    [v, g] = ne.useState(false),
    [b, y] = ne.useState(() => structuredClone(e.initialRequest?.steps ?? [])),
    [w, E] = ne.useState(null),
    [A, k] = ne.useState("trim"),
    [F, I] = ne.useState("1"),
    [O, j] = ne.useState("collapse"),
    [z, V] = ne.useState("upper"),
    [D, B] = ne.useState(""),
    [_, C] = ne.useState(""),
    [T, N] = ne.useState("delimited"),
    [R, L] = ne.useState(",;"),
    [$, U] = ne.useState(false),
    [G, H] = ne.useState("5, 10"),
    [Q, X$1] = ne.useState(" "),
    [J, q] = ne.useState("Combined"),
    [W, K] = ne.useState(false),
    [te, ae] = ne.useState("number"),
    [oe, de] = ne.useState("."),
    [he, Se] = ne.useState(","),
    [le, ce] = ne.useState("MDY"),
    [ge, se] = ne.useState("down"),
    [we, Ee] = ne.useState(""),
    [Ae, ye] = ne.useState("text"),
    [Te, Pe] = ne.useState(""),
    [qe, ht] = ne.useState("2, 3"),
    [Re, Me] = ne.useState(""),
    [ke, je] = ne.useState([]),
    [$e, Je] = ne.useState(""),
    [ze, at] = ne.useState(""),
    [_e, Ze] = ne.useState(null),
    [lt, be] = ne.useState(null),
    [ft, pt] = ne.useState(0),
    [Fe, nt] = ne.useState(0),
    [mt, Wt] = ne.useState(false),
    Zt = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((Qe) => Qe.id)),
    ].sort(),
    xt = {
      readOnly: e.readOnly,
      allowedSheetIds: Zt,
      includeHidden: !!e.allowHidden && v,
      allowNewSheet: e.allowNewSheet ?? true,
    },
    or = e.workbook.sheets.filter(
      (Qe) =>
        Zt.includes(Qe.id) &&
        (xt.includeHidden || (!Qe.hidden && !Qe.veryHidden)),
    ),
    Wr = {
      sourceSheetId: d,
      sourceRange: l,
      header: h,
      outputName: p,
      steps: b,
    },
    Rr = JSON.stringify({
      request: Wr,
      options: xt,
      selection: [e.sheetId, e.range],
      editing: w,
      operation: A,
      columns: F,
      trimMode: O,
      caseMode: z,
      find: D,
      replacement: _,
      splitMode: T,
      delimiters: R,
      collapse: $,
      breaks: G,
      separator: Q,
      fieldName: J,
      skipBlank: W,
      convertTo: te,
      decimal: oe,
      group: he,
      dateOrder: le,
      direction: ge,
      fillValue: we,
      fillType: Ae,
      groupColumns: Te,
      exampleRows: qe,
      exampleValues: Re,
      candidate: $e,
    }),
    qt =
      _e?.key === Rr && _e.plan.snapshot === JSON.stringify(e.workbook)
        ? _e.plan
        : null,
    Ye =
      lt?.key === Rr && lt.snapshot === JSON.stringify(e.workbook)
        ? lt.message
        : null,
    bt = JSON.stringify({
      sourceSheetId: d,
      sourceRange: l,
      header: h,
      outputName: p,
      steps: b,
      editing: w,
      columns: F,
      exampleRows: qe,
      exampleValues: Re,
      options: xt,
      selection: [e.sheetId, e.range],
      snapshot: e.workbook,
    }),
    [St, cr] = ne.useState(""),
    dr = St === bt ? ke : [];
  (ne.useEffect(() => {
    const Qe = document.activeElement;
    return (
      o.current
        ?.querySelector('[data-ui-label="Transform source range"]')
        ?.focus(),
      () => {
        Qe?.isConnected && Qe.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      (s.current[0] === e.sheetId && s.current[1] === e.range) ||
        ((s.current = [e.sheetId, e.range]),
        u(e.sheetId),
        c(e.range),
        Ze(null),
        je([]),
        Je(""),
        be(null),
        Wt(false));
    }, [e.sheetId, e.range]));
  const Ut = (Qe) =>
      be({
        message:
          Qe instanceof Error
            ? Qe.message
            : "Review the transformation settings.",
        key: Rr,
        snapshot: JSON.stringify(e.workbook),
      }),
    Jt = () => {
      (Ze(null), be(null), Wt(false), (a.current = false));
    },
    fr = () => {
      if (!/^\s*\d+(?:\s*,\s*\d+)+\s*$/.test(qe))
        throw Error(
          "Enter at least two worksheet row numbers, separated by commas.",
        );
      const Qe = qe.split(",").map((It) => Number(It)),
        Dt = Re.split(`
`);
      if (Qe.length !== Dt.length)
        throw Error("Provide one example value per row, on separate lines.");
      return Qe.map((It, ar) => ({ row: It, value: Dt[ar] }));
    },
    br = () => {
      Jt();
      try {
        const Qe = selectExceptionFreeRows(
          e.workbook,
          { ...Wr, steps: w === null ? b : b.slice(0, w) },
          fr(),
          parseFieldPositions(F),
          xt,
        );
        (je(Qe.candidates),
          Je(""),
          cr(bt),
          at(
            `${Qe.candidates.length} matching patterns from ${Qe.searched} candidates. ${Qe.truncated ? "Results are capped. " : ""}Search is bounded; choose a pattern, add the step, and review every resulting row.`,
          ));
      } catch (Qe) {
        Ut(Qe);
      }
    },
    rt = () => {
      const Qe = parseFieldPositions(F);
      if (A === "trim") return { kind: A, columns: Qe, mode: O };
      if (A === "clean" || A === "select") return { kind: A, columns: Qe };
      if (A === "case") return { kind: A, columns: Qe, mode: z };
      if (A === "replace")
        return { kind: A, columns: Qe, find: D, replacement: _ };
      if (A === "split") {
        if (Qe.length !== 1) throw Error("Split one field at a time.");
        if (T === "fixed" && !/^\s*\d+(?:\s*,\s*\d+)*\s*$/.test(G))
          throw Error(
            "Enter increasing fixed-width break positions separated by commas.",
          );
        const It = T === "fixed" ? G.split(",").map(Number) : [];
        return {
          kind: A,
          column: Qe[0],
          mode: T,
          delimiters: R,
          collapse: $,
          breaks: It,
        };
      }
      if (A === "concat")
        return { kind: A, columns: Qe, separator: Q, name: J, skipBlank: W };
      if (A === "convert")
        return {
          kind: A,
          columns: Qe,
          to: te,
          decimal: oe,
          group: he,
          dateOrder: le,
        };
      if (A === "fill") {
        let It = we;
        if (ge === "value" && Ae === "number") {
          if (!we.trim() || !Number.isFinite(Number(we)))
            throw Error("Enter a finite fill number.");
          It = Number(we);
        }
        if (ge === "value" && Ae === "boolean") {
          if (!/^(TRUE|FALSE)$/i.test(we))
            throw Error("Enter TRUE or FALSE for Boolean fill.");
          It = we.toUpperCase() === "TRUE";
        }
        return {
          kind: A,
          columns: Qe,
          direction: ge,
          value: It,
          groupColumns: parseFieldPositions(Te, true),
        };
      }
      const Dt = dr.find((It) => It.id === $e);
      if (!Dt)
        throw Error(
          "Suggest patterns and explicitly choose one before adding the inferred step.",
        );
      return {
        kind: "pattern",
        name: J,
        pattern: Dt.pattern,
        examples: fr(),
        inferenceColumns: Qe,
      };
    },
    Le = () => {
      Jt();
      try {
        const Qe = rt(),
          Dt = [...b];
        (w === null ? Dt.push(Qe) : (Dt[w] = Qe),
          buildTransformationPreview(e.workbook, { ...Wr, steps: Dt }, xt),
          y(Dt),
          E(null),
          je([]),
          Je(""));
      } catch (Qe) {
        Ut(Qe);
      }
    },
    et = (Qe) => {
      Jt();
      const Dt = b[Qe];
      (E(Qe),
        k(Dt.kind),
        I(
          ("columns" in Dt
            ? Dt.columns
            : "column" in Dt
              ? [Dt.column]
              : Dt.inferenceColumns
          )
            .map((It) => It + 1)
            .join(", "),
        ),
        Dt.kind === "trim" && j(Dt.mode),
        Dt.kind === "case" && V(Dt.mode),
        Dt.kind === "replace" && (B(Dt.find), C(Dt.replacement)),
        Dt.kind === "split" &&
          (N(Dt.mode),
          L(Dt.delimiters),
          U(Dt.collapse),
          H(Dt.breaks.join(", "))),
        Dt.kind === "concat" &&
          (q(Dt.name), X$1(Dt.separator), K(Dt.skipBlank)),
        Dt.kind === "convert" &&
          (ae(Dt.to), de(Dt.decimal), Se(Dt.group), ce(Dt.dateOrder)),
        Dt.kind === "fill" &&
          (se(Dt.direction),
          Ee(Dt.value === null ? "" : String(Dt.value)),
          ye(
            typeof Dt.value == "number"
              ? "number"
              : typeof Dt.value == "boolean"
                ? "boolean"
                : "text",
          ),
          Pe(Dt.groupColumns.map((It) => It + 1).join(", "))),
        Dt.kind === "pattern" &&
          (q(Dt.name),
          ht(Dt.examples.map((It) => It.row).join(", ")),
          Me(
            Dt.examples.map((It) => It.value).join(`
`),
          ),
          je([]),
          Je("")));
    },
    st = () => {
      Jt();
      try {
        const Qe = buildTransformationPreview(e.workbook, Wr, xt);
        (pt(0), nt(0), Ze({ plan: Qe, key: Rr }));
      } catch (Qe) {
        Ut(Qe);
      }
    },
    kt = () => {
      if (!(!qt || e.readOnly || a.current))
        try {
          const Qe = commitApprovedTransformation(e.workbook, qt, {
            ...xt,
            approved: true,
            acceptExceptions: mt,
          });
          ((a.current = true),
            e.onWorkbookChange(
              Qe,
              `Clean & Reshape: ${qt.sourceSheetName}!${qt.request.sourceRange} → ${qt.request.outputName} (${qt.exceptions.length} retained parse exceptions)`,
            ),
            Ze(null));
        } catch (Qe) {
          Ut(Qe);
        }
    },
    gr = (Qe, Dt) => {
      const It = URL.createObjectURL(
          new Blob([Qe], { type: "application/json" }),
        ),
        ar = document.createElement("a");
      ((ar.href = It), (ar.download = Dt), ar.click(), URL.revokeObjectURL(It));
    },
    Ir = (e.workbook.fieldTransforms ?? []).filter((Qe) => {
      try {
        const Dt = e.workbook.sheets.find((It) => It.id === Qe.sourceSheetId);
        return Dt
          ? (assertSelectionInAuthorizedSheet(Dt, Qe.request.sourceRange, {
              allowedSheetIds: Zt,
              includeHidden: xt.includeHidden,
              allowNewSheet: xt.allowNewSheet,
            }),
            true)
          : false;
      } catch {
        return false;
      }
    });
  return i.jsx("div", {
    className: "fr-backdrop",
    children: i.jsxs("section", {
      className: "fr-dialog",
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (Qe) => {
        var It, ar;
        if (
          (Qe.stopPropagation(),
          Qe.key === "Escape" && (Qe.preventDefault(), e.onClose()),
          Qe.key === "Tab")
        ) {
          const Lr = [
            ...(o.current?.querySelectorAll(
              "button,input,select,textarea,summary",
            ) ?? []),
          ].filter(
            (Pr) =>
              !Pr.matches(":disabled") &&
              !Pr.hidden &&
              !Pr.closest("details:not([open]) > :not(summary)"),
          );
          Qe.shiftKey && document.activeElement === Lr[0]
            ? (Qe.preventDefault(), (It = Lr.at(-1)) == null || It.focus())
            : !Qe.shiftKey &&
              document.activeElement === Lr.at(-1) &&
              (Qe.preventDefault(), (ar = Lr[0]) == null || ar.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Clean & Reshape",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Review typed results and exceptions before creating a separate worksheet.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Clean & Reshape"),
              onClick: e.onClose,
              lang: language("Close Clean & Reshape"),
              "data-ui-label": "Close Clean & Reshape",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fr-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Original cells stay intact. Steps run in order using the field positions from the preceding step. Inference never applies automatically.",
              }),
            }),
            e.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Read-only: saved evidence is available; recipe preview and Apply require write access.",
                }),
              }),
            i.jsxs("fieldset", {
              disabled: e.readOnly,
              className: "fr-fields",
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "Source and output",
                  }),
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Source worksheet" }),
                    i.jsx("select", {
                      "aria-label": text("Transform source worksheet"),
                      value: d,
                      onChange: (Qe) => u(Qe.target.value),
                      lang: language("Transform source worksheet"),
                      "data-ui-label": "Transform source worksheet",
                      children: or.map((Qe) =>
                        i.jsx(
                          "option",
                          { value: Qe.id, children: Qe.name },
                          Qe.id,
                        ),
                      ),
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Source range" }),
                    i.jsx("input", {
                      "aria-label": text("Transform source range"),
                      value: l,
                      onChange: (Qe) => c(Qe.target.value),
                      lang: language("Transform source range"),
                      "data-ui-label": "Transform source range",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Output worksheet" }),
                    i.jsx("input", {
                      "aria-label": text("Transform output worksheet"),
                      value: p,
                      onChange: (Qe) => m(Qe.target.value),
                      lang: language("Transform output worksheet"),
                      "data-ui-label": "Transform output worksheet",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  className: "fr-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: h,
                      onChange: (Qe) => f(Qe.target.checked),
                    }),
                    i.jsx(LocalizedText, {
                      children: "First source row contains headings",
                    }),
                  ],
                }),
                e.allowHidden &&
                  i.jsxs("label", {
                    className: "fr-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: v,
                        onChange: (Qe) => g(Qe.target.checked),
                      }),
                      i.jsx(LocalizedText, {
                        children: "Include hidden source cells",
                      }),
                    ],
                  }),
              ],
            }),
            i.jsxs("section", {
              "aria-label": text("Transformation recipe"),
              lang: language("Transformation recipe"),
              "data-ui-label": "Transformation recipe",
              children: [
                i.jsxs("h3", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Transformation steps (",
                    }),
                    b.length,
                    "/12)",
                  ],
                }),
                !b.length &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Add one or more steps, then preview the complete result.",
                    }),
                  }),
                i.jsx("ol", {
                  className: "fr-steps",
                  children: b.map((Qe, Dt) =>
                    i.jsxs(
                      "li",
                      {
                        children: [
                          i.jsx("span", {
                            children: describeSplitFieldStep(Qe),
                          }),
                          i.jsxs("div", {
                            children: [
                              i.jsxs("button", {
                                disabled: e.readOnly,
                                onClick: () => et(Dt),
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Edit step ",
                                  }),
                                  Dt + 1,
                                ],
                              }),
                              i.jsx("button", {
                                disabled: e.readOnly || Dt === 0,
                                "aria-label": `Move step ${Dt + 1} up`,
                                onClick: () => {
                                  Jt();
                                  const It = [...b];
                                  (([It[Dt - 1], It[Dt]] = [
                                    It[Dt],
                                    It[Dt - 1],
                                  ]),
                                    y(It),
                                    E(null));
                                },
                                children: "↑",
                              }),
                              i.jsx("button", {
                                disabled: e.readOnly || Dt === b.length - 1,
                                "aria-label": `Move step ${Dt + 1} down`,
                                onClick: () => {
                                  Jt();
                                  const It = [...b];
                                  (([It[Dt + 1], It[Dt]] = [
                                    It[Dt],
                                    It[Dt + 1],
                                  ]),
                                    y(It),
                                    E(null));
                                },
                                children: "↓",
                              }),
                              i.jsxs("button", {
                                disabled: e.readOnly,
                                onClick: () => {
                                  (Jt(),
                                    y(b.filter((It, ar) => ar !== Dt)),
                                    E(null));
                                },
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Remove step ",
                                  }),
                                  Dt + 1,
                                ],
                              }),
                            ],
                          }),
                        ],
                      },
                      Dt,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("fieldset", {
              disabled: e.readOnly,
              className: "fr-fields",
              children: [
                i.jsx("legend", {
                  children: w === null ? "Add a step" : `Edit step ${w + 1}`,
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Operation" }),
                    i.jsx("select", {
                      "aria-label": text("Field operation"),
                      value: A,
                      onChange: (Qe) => k(Qe.target.value),
                      lang: language("Field operation"),
                      "data-ui-label": "Field operation",
                      children: textCleanupStepLabels.map(([Qe, Dt]) =>
                        i.jsx("option", { value: Qe, children: Dt }, Qe),
                      ),
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Field positions" }),
                    i.jsx("input", {
                      "aria-label": text("Field positions"),
                      value: F,
                      onChange: (Qe) => I(Qe.target.value),
                      placeholder: "1, 3",
                      lang: language("Field positions"),
                      "data-ui-label": "Field positions",
                    }),
                  ],
                }),
                A === "trim" &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Whitespace" }),
                      i.jsxs("select", {
                        "aria-label": text("Whitespace handling"),
                        value: O,
                        onChange: (Qe) => j(Qe.target.value),
                        lang: language("Whitespace handling"),
                        "data-ui-label": "Whitespace handling",
                        children: [
                          i.jsx("option", {
                            value: "collapse",
                            children: "Trim edges and collapse spaces",
                          }),
                          i.jsx("option", {
                            value: "edges",
                            children: "Trim edges only",
                          }),
                        ],
                      }),
                    ],
                  }),
                A === "case" &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Case" }),
                      i.jsxs("select", {
                        "aria-label": text("Case conversion"),
                        value: z,
                        onChange: (Qe) => V(Qe.target.value),
                        lang: language("Case conversion"),
                        "data-ui-label": "Case conversion",
                        children: [
                          i.jsx("option", {
                            value: "upper",
                            children: "UPPER",
                          }),
                          i.jsx("option", {
                            value: "lower",
                            children: "lower",
                          }),
                          i.jsx("option", {
                            value: "proper",
                            children: "Proper",
                          }),
                          i.jsx("option", {
                            value: "none",
                            children: "Keep case",
                          }),
                        ],
                      }),
                    ],
                  }),
                A === "replace" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Find literal text",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Find literal text"),
                            value: D,
                            onChange: (Qe) => B(Qe.target.value),
                            lang: language("Find literal text"),
                            "data-ui-label": "Find literal text",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Replace with" }),
                          i.jsx("input", {
                            "aria-label": text("Replacement text"),
                            value: _,
                            onChange: (Qe) => C(Qe.target.value),
                            lang: language("Replacement text"),
                            "data-ui-label": "Replacement text",
                          }),
                        ],
                      }),
                    ],
                  }),
                A === "split" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Split mode" }),
                          i.jsxs("select", {
                            "aria-label": text("Field split mode"),
                            value: T,
                            onChange: (Qe) => N(Qe.target.value),
                            lang: language("Field split mode"),
                            "data-ui-label": "Field split mode",
                            children: [
                              i.jsx("option", {
                                value: "delimited",
                                children: "Delimited",
                              }),
                              i.jsx("option", {
                                value: "fixed",
                                children: "Fixed width",
                              }),
                            ],
                          }),
                        ],
                      }),
                      T === "delimited"
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Delimiter characters",
                                  }),
                                  i.jsx("input", {
                                    "aria-label": text("Delimiter characters"),
                                    value: R,
                                    onChange: (Qe) => L(Qe.target.value),
                                    lang: language("Delimiter characters"),
                                    "data-ui-label": "Delimiter characters",
                                  }),
                                ],
                              }),
                              i.jsx("button", {
                                onClick: () => L(R + "	"),
                                children: i.jsx(LocalizedText, {
                                  children: "Add tab delimiter",
                                }),
                              }),
                              i.jsxs("label", {
                                className: "fr-check",
                                children: [
                                  i.jsx("input", {
                                    type: "checkbox",
                                    checked: $,
                                    onChange: (Qe) => U(Qe.target.checked),
                                  }),
                                  i.jsx(LocalizedText, {
                                    children: "Collapse consecutive delimiters",
                                  }),
                                ],
                              }),
                            ],
                          })
                        : i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Break after character positions",
                              }),
                              i.jsx("input", {
                                "aria-label": text("Fixed-width breaks"),
                                value: G,
                                onChange: (Qe) => H(Qe.target.value),
                                lang: language("Fixed-width breaks"),
                                "data-ui-label": "Fixed-width breaks",
                              }),
                            ],
                          }),
                    ],
                  }),
                (A === "concat" || A === "pattern") &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "New field name" }),
                      i.jsx("input", {
                        "aria-label": text("New field name"),
                        value: J,
                        onChange: (Qe) => q(Qe.target.value),
                        lang: language("New field name"),
                        "data-ui-label": "New field name",
                      }),
                    ],
                  }),
                A === "concat" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Separator" }),
                          i.jsx("input", {
                            "aria-label": text("Concatenation separator"),
                            value: Q,
                            onChange: (Qe) => X$1(Qe.target.value),
                            lang: language("Concatenation separator"),
                            "data-ui-label": "Concatenation separator",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        className: "fr-check",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: W,
                            onChange: (Qe) => K(Qe.target.checked),
                          }),
                          i.jsx(LocalizedText, {
                            children: "Skip blank fields",
                          }),
                        ],
                      }),
                    ],
                  }),
                A === "convert" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Result type" }),
                          i.jsx("select", {
                            "aria-label": text("Conversion type"),
                            value: te,
                            onChange: (Qe) => ae(Qe.target.value),
                            lang: language("Conversion type"),
                            "data-ui-label": "Conversion type",
                            children: [
                              "text",
                              "number",
                              "integer",
                              "boolean",
                              "date",
                            ].map((Qe) =>
                              i.jsx("option", { value: Qe, children: Qe }, Qe),
                            ),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Decimal separator",
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Transform decimal separator"),
                            value: oe,
                            onChange: (Qe) => de(Qe.target.value),
                            lang: language("Transform decimal separator"),
                            "data-ui-label": "Transform decimal separator",
                            children: [
                              i.jsx("option", { value: ".", children: "Dot" }),
                              i.jsx("option", {
                                value: ",",
                                children: "Comma",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Grouping separator",
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Transform grouping separator"),
                            value: he,
                            onChange: (Qe) => Se(Qe.target.value),
                            lang: language("Transform grouping separator"),
                            "data-ui-label": "Transform grouping separator",
                            children: [
                              i.jsx("option", { value: "", children: "None" }),
                              i.jsx("option", {
                                value: ",",
                                children: "Comma",
                              }),
                              i.jsx("option", { value: ".", children: "Dot" }),
                              i.jsx("option", {
                                value: " ",
                                children: "Space",
                              }),
                              i.jsx("option", {
                                value: "'",
                                children: "Apostrophe",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Slash date order",
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Transform date order"),
                            value: le,
                            onChange: (Qe) => ce(Qe.target.value),
                            lang: language("Transform date order"),
                            "data-ui-label": "Transform date order",
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
                    ],
                  }),
                A === "fill" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Fill direction" }),
                          i.jsxs("select", {
                            "aria-label": text("Fill direction"),
                            value: ge,
                            onChange: (Qe) => se(Qe.target.value),
                            lang: language("Fill direction"),
                            "data-ui-label": "Fill direction",
                            children: [
                              i.jsx("option", {
                                value: "down",
                                children: "Down from previous value",
                              }),
                              i.jsx("option", {
                                value: "up",
                                children: "Up from next value",
                              }),
                              i.jsx("option", {
                                value: "value",
                                children: "Constant value",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Group field positions (optional)",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Fill group fields"),
                            value: Te,
                            onChange: (Qe) => Pe(Qe.target.value),
                            lang: language("Fill group fields"),
                            "data-ui-label": "Fill group fields",
                          }),
                        ],
                      }),
                      ge === "value" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Fill type" }),
                                i.jsxs("select", {
                                  "aria-label": text("Fill value type"),
                                  value: Ae,
                                  onChange: (Qe) => ye(Qe.target.value),
                                  lang: language("Fill value type"),
                                  "data-ui-label": "Fill value type",
                                  children: [
                                    i.jsx("option", {
                                      value: "text",
                                      children: "Text",
                                    }),
                                    i.jsx("option", {
                                      value: "number",
                                      children: "Number",
                                    }),
                                    i.jsx("option", {
                                      value: "boolean",
                                      children: "Boolean",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Fill value",
                                }),
                                i.jsx("input", {
                                  "aria-label": text("Fill value"),
                                  value: we,
                                  onChange: (Qe) => Ee(Qe.target.value),
                                  lang: language("Fill value"),
                                  "data-ui-label": "Fill value",
                                }),
                              ],
                            }),
                          ],
                        }),
                    ],
                  }),
                A === "pattern" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Example worksheet rows",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Inference example rows"),
                            value: qe,
                            onChange: (Qe) => ht(Qe.target.value),
                            lang: language("Inference example rows"),
                            "data-ui-label": "Inference example rows",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Expected text, one example per line",
                          }),
                          i.jsx("textarea", {
                            "aria-label": text("Inference example values"),
                            value: Re,
                            onChange: (Qe) => Me(Qe.target.value),
                            rows: 3,
                            lang: language("Inference example values"),
                            "data-ui-label": "Inference example values",
                          }),
                        ],
                      }),
                      i.jsx("button", {
                        onClick: br,
                        children: i.jsx(LocalizedText, {
                          children: "Suggest patterns",
                        }),
                      }),
                      St === bt &&
                        i.jsx("p", {
                          role: "status",
                          className: "fr-wide",
                          children: ze,
                        }),
                      dr.length > 0 &&
                        i.jsxs("label", {
                          className: "fr-wide",
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Choose a matching pattern",
                            }),
                            i.jsxs("select", {
                              "aria-label": text("Inferred pattern"),
                              value: $e,
                              onChange: (Qe) => Je(Qe.target.value),
                              lang: language("Inferred pattern"),
                              "data-ui-label": "Inferred pattern",
                              children: [
                                i.jsx("option", {
                                  value: "",
                                  children: "Choose explicitly…",
                                }),
                                dr.map((Qe) =>
                                  i.jsx(
                                    "option",
                                    { value: Qe.id, children: Qe.description },
                                    Qe.id,
                                  ),
                                ),
                              ],
                            }),
                          ],
                        }),
                    ],
                  }),
                i.jsxs("div", {
                  className: "fr-wide fr-actions",
                  children: [
                    i.jsx("button", {
                      onClick: Le,
                      disabled: w === null && b.length >= 12,
                      children:
                        w === null
                          ? "Add transformation step"
                          : "Save transformation step",
                    }),
                    w !== null &&
                      i.jsx("button", {
                        onClick: () => {
                          (Jt(), E(null));
                        },
                        children: i.jsx(LocalizedText, {
                          children: "Cancel step edit",
                        }),
                      }),
                  ],
                }),
              ],
            }),
            Ye &&
              i.jsx("p", {
                role: "alert",
                className: "fr-error",
                children: Ye,
              }),
            _e &&
              !qt &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Settings, source or permissions changed. Preview again before applying.",
                }),
              }),
            i.jsxs("div", {
              className: "fr-actions",
              children: [
                i.jsx("button", {
                  disabled: e.readOnly || !b.length || w !== null,
                  onClick: st,
                  children: i.jsx(LocalizedText, {
                    children: "Preview transformation",
                  }),
                }),
                qt &&
                  i.jsx("button", {
                    onClick: () => {
                      try {
                        gr(
                          verifyTransformationPreview(e.workbook, qt, xt),
                          "Field-transformation-preview.json",
                        );
                      } catch (Qe) {
                        Ut(Qe);
                      }
                    },
                    children: i.jsx(LocalizedText, {
                      children: "Download transformation preview",
                    }),
                  }),
              ],
            }),
            qt &&
              i.jsxs("section", {
                className: "fr-preview",
                "aria-label": text("Transformation preview"),
                lang: language("Transformation preview"),
                "data-ui-label": "Transformation preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Review before Apply",
                    }),
                  }),
                  i.jsxs("p", {
                    role: "status",
                    children: [
                      qt.counts.sourceRows,
                      " ",
                      i.jsx(LocalizedText, { children: "source rows · " }),
                      qt.counts.outputCells,
                      " ",
                      i.jsx(LocalizedText, { children: "output cells · " }),
                      qt.counts.changedCells,
                      " ",
                      i.jsx(LocalizedText, { children: "changed values ·" }),
                      " ",
                      qt.counts.exceptions,
                      " ",
                      i.jsx(LocalizedText, { children: "parse exceptions in" }),
                      " ",
                      qt.counts.exceptionRows,
                      " ",
                      i.jsx(LocalizedText, { children: "rows." }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Original: " }),
                      qt.sourceSheetName,
                      "!",
                      qt.request.sourceRange,
                      i.jsx(LocalizedText, { children: ", revision " }),
                      e.workbook.revision,
                      i.jsx(LocalizedText, { children: ". Output:" }),
                      " ",
                      qt.request.outputName,
                      "!",
                      qt.targetRange,
                      ".",
                    ],
                  }),
                  i.jsxs("details", {
                    children: [
                      i.jsx("summary", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Step results and resulting field positions",
                        }),
                      }),
                      qt.stages.map((Qe) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsx(LocalizedText, { children: "Step " }),
                              Qe.step,
                              " (",
                              Qe.kind,
                              "):",
                              " ",
                              Qe.fields
                                .map((Dt, It) => `${It + 1}: ${Dt}`)
                                .join(" · "),
                              ".",
                              " ",
                              Qe.changedCells,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "affected cells; ",
                              }),
                              Qe.exceptions,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "new exceptions.",
                              }),
                            ],
                          },
                          Qe.step,
                        ),
                      ),
                    ],
                  }),
                  i.jsx("div", {
                    className: "fr-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("caption", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Before/after values and original source locators",
                          }),
                        }),
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Output",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Field",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Original source / raw value",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "After",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Type",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Review",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: qt.outputs
                            .slice(ft * 20, (ft + 1) * 20)
                            .map((Qe) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("td", { children: Qe.address }),
                                    i.jsx("td", { children: Qe.field }),
                                    i.jsx("td", {
                                      children: i.jsxs("details", {
                                        children: [
                                          i.jsx("summary", {
                                            children: Qe.sources
                                              .map(
                                                (Dt) =>
                                                  `${qt.sourceSheetName}!${Dt}`,
                                              )
                                              .join(", "),
                                          }),
                                          Qe.sources.map((Dt) =>
                                            i.jsxs(
                                              "pre",
                                              {
                                                children: [
                                                  Dt,
                                                  ":",
                                                  " ",
                                                  stringifyJsonValue(
                                                    qt.sources.find(
                                                      (Lr) => Lr.address === Dt,
                                                    )?.original?.raw ?? null,
                                                  ),
                                                ],
                                              },
                                              Dt,
                                            ),
                                          ),
                                        ],
                                      }),
                                    }),
                                    i.jsx("td", {
                                      children: i.jsx("code", {
                                        children: stringifyJsonValue(Qe.value),
                                      }),
                                    }),
                                    i.jsx("td", { children: Qe.type }),
                                    i.jsx("td", {
                                      children: Qe.exception
                                        ? "Exception retained"
                                        : "Ready",
                                    }),
                                  ],
                                },
                                Qe.address,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "fr-actions",
                    children: [
                      i.jsx("button", {
                        disabled: ft === 0,
                        onClick: () => pt(ft - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous results",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Results " }),
                          ft * 20 + 1,
                          "–",
                          Math.min((ft + 1) * 20, qt.outputs.length),
                          " ",
                          i.jsx(LocalizedText, { children: "of" }),
                          " ",
                          qt.outputs.length,
                        ],
                      }),
                      i.jsx("button", {
                        disabled: (ft + 1) * 20 >= qt.outputs.length,
                        onClick: () => pt(ft + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next results",
                        }),
                      }),
                    ],
                  }),
                  qt.exceptions.length > 0 &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(LocalizedText, {
                            children: "Retained parse exceptions",
                          }),
                        }),
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Raw values remain in the derived output and an exceptions appendix. Later steps do not convert a failed field. Source cells remain unchanged.",
                          }),
                        }),
                        i.jsx("div", {
                          className: "fr-scroll",
                          children: i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Row / field",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Step",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Raw value",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Source",
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
                                children: qt.exceptions
                                  .slice(Fe * 20, (Fe + 1) * 20)
                                  .map((Qe, Dt) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [Qe.row, " / ", Qe.field],
                                          }),
                                          i.jsx("td", { children: Qe.step }),
                                          i.jsx("td", {
                                            children: i.jsx("code", {
                                              children: stringifyJsonValue(
                                                Qe.raw,
                                              ),
                                            }),
                                          }),
                                          i.jsx("td", {
                                            children: Qe.sources.join(", "),
                                          }),
                                          i.jsx("td", { children: Qe.reason }),
                                        ],
                                      },
                                      Dt,
                                    ),
                                  ),
                              }),
                            ],
                          }),
                        }),
                        i.jsxs("div", {
                          className: "fr-actions",
                          children: [
                            i.jsx("button", {
                              disabled: !Fe,
                              onClick: () => nt(Fe - 1),
                              children: i.jsx(LocalizedText, {
                                children: "Previous exceptions",
                              }),
                            }),
                            i.jsxs("span", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Exceptions ",
                                }),
                                Fe * 20 + 1,
                                "–",
                                Math.min((Fe + 1) * 20, qt.exceptions.length),
                                " ",
                                i.jsx(LocalizedText, { children: "of " }),
                                qt.exceptions.length,
                              ],
                            }),
                            i.jsx("button", {
                              disabled: (Fe + 1) * 20 >= qt.exceptions.length,
                              onClick: () => nt(Fe + 1),
                              children: i.jsx(LocalizedText, {
                                children: "Next exceptions",
                              }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          className: "fr-check",
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: mt,
                              onChange: (Qe) => Wt(Qe.target.checked),
                            }),
                            i.jsx(LocalizedText, {
                              children:
                                "Keep raw values and the exceptions appendix",
                            }),
                          ],
                        }),
                      ],
                    }),
                ],
              }),
            Ir.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Saved transformation evidence (",
                      }),
                      Ir.length,
                      ")",
                    ],
                  }),
                  Ir.map((Qe) =>
                    i.jsxs(
                      "article",
                      {
                        className: "fr-record",
                        children: [
                          i.jsx("h3", { children: Qe.outputSheetName }),
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, { children: "Source " }),
                              Qe.sourceSheetName,
                              "!",
                              Qe.request.sourceRange,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "· original revision ",
                              }),
                              Qe.sourceRevision,
                              " · ",
                              Qe.createdAt,
                              " · ",
                              Qe.exceptions.length,
                              " ",
                              i.jsx(LocalizedText, {
                                children:
                                  "exceptions. These are historical snapshot locators.",
                              }),
                            ],
                          }),
                          i.jsx("ol", {
                            children: Qe.request.steps.map((Dt, It) =>
                              i.jsx(
                                "li",
                                { children: describeSplitFieldStep(Dt) },
                                It,
                              ),
                            ),
                          }),
                          i.jsxs("button", {
                            onClick: () =>
                              gr(
                                JSON.stringify(Qe, null, 2),
                                "Field-transformation-record.json",
                              ),
                            children: [
                              i.jsx(LocalizedText, { children: "Download " }),
                              Qe.outputSheetName,
                              " ",
                              i.jsx(LocalizedText, { children: "evidence" }),
                            ],
                          }),
                          i.jsxs("button", {
                            disabled: e.readOnly,
                            onClick: () => {
                              (Jt(),
                                u(Qe.request.sourceSheetId),
                                c(Qe.request.sourceRange),
                                f(Qe.request.header),
                                y(structuredClone(Qe.request.steps)),
                                m(`${Qe.outputSheetName.slice(0, 26)} copy`),
                                E(null),
                                je([]),
                                Je(""));
                            },
                            children: [
                              i.jsx(LocalizedText, { children: "Reuse " }),
                              Qe.outputSheetName,
                              " ",
                              i.jsx(LocalizedText, { children: "recipe" }),
                            ],
                          }),
                        ],
                      },
                      Qe.id,
                    ),
                  ),
                ],
              }),
            i.jsx("p", {
              className: "fr-boundary",
              children: i.jsx(LocalizedText, {
                children:
                  "Up to 500 source cells, 12 steps and 5,000 output cells including exceptions. Sources must be reviewed values. Patterns are deterministic and bounded. Native JSON retains recipes and original evidence; XLSX retains current cells and notes. Local history is editable outside this workflow.",
              }),
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "fr-primary",
              disabled:
                !qt ||
                e.readOnly ||
                (qt.exceptions.length > 0 && !mt) ||
                w !== null,
              onClick: kt,
              children: i.jsx(LocalizedText, {
                children: "Apply reviewed transformation",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const sortEngineLimits = Object.freeze({
    rows: 1e4,
    cells: 1e5,
    keys: 64,
    customItems: 100,
    formulaCells: 5e3,
    formulaCharacters: 1e6,
    rangeRewriteWork: 5e6,
    pageSize: 25,
  }),
  stableSortPolicy = Object.freeze({
    ties: "Stable ties retain their physical source-row order at preview time, in both directions.",
    records:
      "Whole worksheet records move across every populated column in the chosen row interval. A partial selection requires expansion; cancelling preserves the source.",
    types:
      "Automatic uses numbers, text, booleans, errors, then blanks. Descending reverses nonblank order; blanks stay last. Text uses en-US collation with case sensitivity selectable. Numeric-looking text stays text unless Number is explicitly selected. Date accepts numeric serials or the workbook's strict ISO/MDY/DMY policy.",
    references:
      "Direct references follow record identities, including absolute markers. Contiguous range membership is preserved; a range that would split is refused. Order-sensitive calculations may change after sorting.",
    scope:
      "All source rows in the range participate, including empty-key rows. Hidden rows require explicit inclusion. Temporary view filters are separate. Color/icon and left-to-right sorts are outside this contract.",
  }),
  MW = new WeakMap(),
  formatSortRangeAddress = (e) => {
    const r = parseRangeAddress(e),
      t = formatCellAddress(r.firstRow, r.firstCol),
      n = formatCellAddress(r.lastRow, r.lastCol);
    return t === n ? t : `${t}:${n}`;
  },
  isNonBlankCellEntry = (e) =>
    e.raw !== null || Object.keys(e).some((r) => r !== "raw"),
  N3 = (e) => !e.literal && typeof e.raw == "string" && e.raw.startsWith("="),
  buildAllowedSheetKey = (e, r) =>
    JSON.stringify(
      [...(r.allowedSheetIds ?? e.sheets.map((t) => t.id))].sort(),
    );
function assertSortWorkbookAuthorized(e, r, t) {
  if (t.readOnly)
    throw Error(
      "This workbook is read-only. Sorting cannot change its records.",
    );
  if (t.allowedSheetIds && !t.allowedSheetIds.includes(r))
    throw Error("The sort source is outside the host-approved worksheets.");
  const n = findSheetOrThrow(e, r);
  if (n.hidden || n.veryHidden)
    throw Error("Unhide this worksheet before sorting its records.");
  return (assertSheetUnprotected(n), n);
}
function resolveSortSourceRange(e, r, t, n = false) {
  const o = findSheetOrThrow(e, r),
    a = parseRangeAddress(t);
  if (a.lastRow >= o.rowCount || a.lastCol >= o.colCount)
    throw Error("The sort selection is outside the worksheet.");
  let s = a.firstRow,
    d = a.lastRow,
    u = a.firstCol,
    l = a.lastCol;
  const c = Object.entries(o.cells)
    .filter(([, m]) => isNonBlankCellEntry(m))
    .map(([m]) => _t(m));
  if (s === d && u === l) {
    const m = o.tables
      ?.map((v) => ({
        bounds: parseRangeAddress(v.range),
        totals: v.showTotals,
      }))
      .find(
        (v) =>
          s >= v.bounds.firstRow &&
          s <= v.bounds.lastRow &&
          u >= v.bounds.firstCol &&
          u <= v.bounds.lastCol,
      );
    if (m)
      ((s = m.bounds.firstRow),
        (d = m.bounds.lastRow - (m.totals ? 1 : 0)),
        (u = m.bounds.firstCol),
        (l = m.bounds.lastCol));
    else {
      const v = new Set(c.map((g) => g.row));
      if (!v.has(s))
        throw Error("Select a populated record or an explicit range to sort.");
      for (; s > 0 && v.has(s - 1);) s--;
      for (; d + 1 < o.rowCount && v.has(d + 1);) d++;
    }
  }
  for (const m of c)
    m.row >= s &&
      m.row <= d &&
      ((u = Math.min(u, m.col)), (l = Math.max(l, m.col)));
  const h = formatSortRangeAddress(
      `${formatCellAddress(s, u)}:${formatCellAddress(d, l)}`,
    ),
    f = formatSortRangeAddress(t);
  return {
    selectedRange: f,
    expandedRange: h,
    expansionRequired: f !== h,
    firstDataRow: s + (n ? 1 : 0),
    lastRow: d,
  };
}
function normalizeSortSpecification(e, r) {
  if (
    !r ||
    typeof r.headers != "boolean" ||
    !Array.isArray(r.keys) ||
    !r.keys.length ||
    r.keys.length > sortEngineLimits.keys ||
    [r.caseSensitive, r.expandSelection, r.includeHidden].some(
      (o) => o !== undefined && typeof o != "boolean",
    )
  )
    throw Error("Choose headers and 1–64 sort levels with valid options.");
  const t = structuredClone(r);
  t.range = formatSortRangeAddress(r.range);
  const n = new Set();
  for (const o of t.keys) {
    if (
      !Number.isInteger(o.column) ||
      o.column < 0 ||
      o.column >= findSheetOrThrow(e, r.sheetId).colCount ||
      n.has(o.column) ||
      !["ascending", "descending"].includes(o.direction) ||
      !["auto", "number", "date", "text"].includes(o.type)
    )
      throw Error(
        "Choose distinct valid sort columns, directions and data types.",
      );
    if ((n.add(o.column), o.customOrder !== undefined)) {
      const a = (s) =>
        t.caseSensitive
          ? s.normalize("NFC")
          : s.normalize("NFC").toLocaleLowerCase("en-US");
      if (
        !Array.isArray(o.customOrder) ||
        !o.customOrder.length ||
        o.customOrder.length > sortEngineLimits.customItems ||
        o.type !== "text" ||
        o.customOrder.some(
          (s) => typeof s != "string" || !s.length || s.length > 255,
        ) ||
        new Set(o.customOrder.map(a)).size !== o.customOrder.length
      )
        throw Error(
          "Custom orders require Text, 1–100 distinct nonempty labels of at most 255 characters.",
        );
    }
  }
  return t;
}
function buildSortPlan(e, r, t) {
  const n = parseRangeAddress(t.expandedRange),
    o = calculateWorkbookWithSettings(e),
    a = normalizeDatePolicy(e.datePolicy);
  return Array.from({ length: t.lastRow - t.firstDataRow + 1 }, (s, d) => {
    const u = t.firstDataRow + d,
      l = r.keys.map((c) => {
        if (c.column < n.firstCol || c.column > n.lastCol)
          throw Error(
            "Every sort key must be inside the complete record range.",
          );
        const h = formatCellAddress(u, c.column),
          f = findSheetOrThrow(e, r.sheetId).cells[h],
          p = o(r.sheetId, h);
        if (
          f &&
          N3(f) &&
          typeof p == "string" &&
          ["#PENDING!", "#TIMEOUT!", "#LIMIT!", "#BLOCKED!"].includes(p)
        )
          throw Error(
            `Calculate ${h} successfully before sorting. Its current result is ${p}.`,
          );
        const m =
          !!f && N3(f) && o.getTypedValue(r.sheetId, h).type === "error";
        let v = p;
        if (p !== null && p !== "" && !m) {
          if ((c.type === "text" && (v = String(p)), c.type === "number")) {
            if (typeof p == "number") v = p;
            else if (
              typeof p == "string" &&
              /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(p.trim())
            )
              v = Number(p);
            else
              throw Error(
                `${h} cannot be interpreted as a Number. Choose Automatic or correct the source.`,
              );
            if (!Number.isFinite(v))
              throw Error(`${h} is outside the finite numeric range.`);
          }
          if (
            c.type === "date" &&
            ((v =
              typeof p == "number"
                ? p
                : typeof p == "string"
                  ? (parseDateText(p, a) ?? null)
                  : null),
            v === null ||
              typeof v != "number" ||
              v < 0 ||
              v >= dateSystemMaxSerial(a.dateSystem) + 1)
          )
            throw Error(
              `${h} is not a valid date under the workbook's date policy.`,
            );
        }
        return { value: p, compare: v, error: m };
      });
    return { row: u, keys: l };
  });
}
function assertSortSafeFormulas(e, r, t) {
  let n = 0,
    o = 0,
    a = 0;
  const s = (d) => {
    (n++, (o += d.length));
    const u = d.replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'/g, "");
    if (/\b(?:INDIRECT|OFFSET)\s*\(/i.test(u))
      throw Error(
        "Sorting cannot safely rewrite dynamic INDIRECT/OFFSET references. Replace them with explicit references first.",
      );
    if (
      /[{}#]/.test(u) ||
      matrixResultFunctionNames.some((l) =>
        new RegExp(`\\b${l.replaceAll(".", "\\.")}\\s*\\(`, "i").test(u),
      )
    )
      throw Error(
        "Sorting cannot safely relocate array/spill formulas or existing reference-error tokens. Resolve or separate those formulas first.",
      );
    a += (u.match(/:/g) ?? []).length;
  };
  for (const d of e.sheets) {
    for (const u of Object.values(d.cells)) N3(u) && s(u.raw);
    for (const u of [...(d.validations ?? []), ...(d.conditionalFormats ?? [])])
      u.formula && s(u.formula);
  }
  for (const d of e.definedNames ?? []) s(d.refersTo);
  if (
    n > sortEngineLimits.formulaCells ||
    o > sortEngineLimits.formulaCharacters ||
    a * r > sortEngineLimits.rangeRewriteWork
  )
    throw Error(
      "The sort exceeds the bounded formula-reference review budget. Reduce or separate the workbook before sorting.",
    );
  if (
    t.allowedSheetIds &&
    e.sheets.some((d) => !t.allowedSheetIds.includes(d.id))
  )
    throw Error(
      "A structural sort requires the host to authorize all worksheets so cross-sheet reference changes can be reviewed.",
    );
}

export {
  assertSortWorkbookAuthorized,
  buildAllowedSheetKey,
  buildSortPlan,
  MW,
  N3,
  stableSortPolicy,
  resolveSortSourceRange,
  assertSortSafeFormulas,
  isNonBlankCellEntry,
  CleanAndReshapeDialog,
  normalizeSortSpecification,
  sortEngineLimits,
};
export { formatSortRangeAddress };
