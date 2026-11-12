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
    [sourceSheetId, u] = ne.useState(
      e.initialRequest?.sourceSheetId ?? e.sheetId,
    ),
    [sourceRange, c] = ne.useState(e.initialRequest?.sourceRange ?? e.range),
    [h, f] = ne.useState(e.initialRequest?.header ?? true),
    [outputName, m] = ne.useState(
      e.initialRequest?.outputName ?? "Reshaped fields",
    ),
    [checked, g] = ne.useState(false),
    [steps, y] = ne.useState(() =>
      structuredClone(e.initialRequest?.steps ?? []),
    ),
    [editing, E] = ne.useState(null),
    [A, k] = ne.useState("trim"),
    [columns, I] = ne.useState("1"),
    [O, j] = ne.useState("collapse"),
    [z, V] = ne.useState("upper"),
    [find, B] = ne.useState(""),
    [replacement, C] = ne.useState(""),
    [T, N] = ne.useState("delimited"),
    [delimiters, L] = ne.useState(",;"),
    [$, U] = ne.useState(false),
    [breaks, H] = ne.useState("5, 10"),
    [separator, X$1] = ne.useState(" "),
    [J, q] = ne.useState("Combined"),
    [W, K] = ne.useState(false),
    [te, ae] = ne.useState("number"),
    [decimal, de] = ne.useState("."),
    [group, Se] = ne.useState(","),
    [dateOrder, ce] = ne.useState("MDY"),
    [direction, se] = ne.useState("down"),
    [fillValue, Ee] = ne.useState(""),
    [fillType, ye] = ne.useState("text"),
    [groupColumns, Pe] = ne.useState(""),
    [exampleRows, ht] = ne.useState("2, 3"),
    [exampleValues, Me] = ne.useState(""),
    [ke, je] = ne.useState([]),
    [candidate, Je] = ne.useState(""),
    [ze, at] = ne.useState(""),
    [_e, Ze] = ne.useState(null),
    [lt, be] = ne.useState(null),
    [ft, pt] = ne.useState(0),
    [Fe, nt] = ne.useState(0),
    [mt, Wt] = ne.useState(false),
    allowedSheetIds = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    xt = {
      readOnly: e.readOnly,
      allowedSheetIds: allowedSheetIds,
      includeHidden: !!e.allowHidden && checked,
      allowNewSheet: e.allowNewSheet ?? true,
    },
    or = e.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        (xt.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    request = {
      sourceSheetId: sourceSheetId,
      sourceRange: sourceRange,
      header: h,
      outputName: outputName,
      steps: steps,
    },
    Rr = JSON.stringify({
      request: request,
      options: xt,
      selection: [e.sheetId, e.range],
      editing: editing,
      operation: A,
      columns: columns,
      trimMode: O,
      caseMode: z,
      find: find,
      replacement: replacement,
      splitMode: T,
      delimiters: delimiters,
      collapse: $,
      breaks: breaks,
      separator: separator,
      fieldName: J,
      skipBlank: W,
      convertTo: te,
      decimal: decimal,
      group: group,
      dateOrder: dateOrder,
      direction: direction,
      fillValue: fillValue,
      fillType: fillType,
      groupColumns: groupColumns,
      exampleRows: exampleRows,
      exampleValues: exampleValues,
      candidate: candidate,
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
      sourceSheetId: sourceSheetId,
      sourceRange: sourceRange,
      header: h,
      outputName: outputName,
      steps: steps,
      editing: editing,
      columns: columns,
      exampleRows: exampleRows,
      exampleValues: exampleValues,
      options: xt,
      selection: [e.sheetId, e.range],
      snapshot: e.workbook,
    }),
    [St, cr] = ne.useState(""),
    dr = St === bt ? ke : [];
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current
        ?.querySelector('[data-ui-label="Transform source range"]')
        ?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
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
      if (!/^\s*\d+(?:\s*,\s*\d+)+\s*$/.test(exampleRows))
        throw Error(
          "Enter at least two worksheet row numbers, separated by commas.",
        );
      const Qe = exampleRows.split(",").map((It) => Number(It)),
        Dt = exampleValues.split(`
`);
      if (Qe.length !== Dt.length)
        throw Error("Provide one example value per row, on separate lines.");
      return Qe.map((row, ar) => ({ row: row, value: Dt[ar] }));
    },
    onClick = () => {
      Jt();
      try {
        const exceptionFreeRows = selectExceptionFreeRows(
          e.workbook,
          {
            ...request,
            steps: editing === null ? steps : steps.slice(0, editing),
          },
          fr(),
          parseFieldPositions(columns),
          xt,
        );
        (je(exceptionFreeRows.candidates),
          Je(""),
          cr(bt),
          at(
            `${exceptionFreeRows.candidates.length} matching patterns from ${exceptionFreeRows.searched} candidates. ${exceptionFreeRows.truncated ? "Results are capped. " : ""}Search is bounded; choose a pattern, add the step, and review every resulting row.`,
          ));
      } catch (error) {
        Ut(error);
      }
    },
    rt = () => {
      const fieldPositions = parseFieldPositions(columns);
      if (A === "trim") return { kind: A, columns: fieldPositions, mode: O };
      if (A === "clean" || A === "select")
        return { kind: A, columns: fieldPositions };
      if (A === "case") return { kind: A, columns: fieldPositions, mode: z };
      if (A === "replace")
        return {
          kind: A,
          columns: fieldPositions,
          find: find,
          replacement: replacement,
        };
      if (A === "split") {
        if (fieldPositions.length !== 1)
          throw Error("Split one field at a time.");
        if (T === "fixed" && !/^\s*\d+(?:\s*,\s*\d+)*\s*$/.test(breaks))
          throw Error(
            "Enter increasing fixed-width break positions separated by commas.",
          );
        const It = T === "fixed" ? breaks.split(",").map(Number) : [];
        return {
          kind: A,
          column: fieldPositions[0],
          mode: T,
          delimiters: delimiters,
          collapse: $,
          breaks: It,
        };
      }
      if (A === "concat")
        return {
          kind: A,
          columns: fieldPositions,
          separator: separator,
          name: J,
          skipBlank: W,
        };
      if (A === "convert")
        return {
          kind: A,
          columns: fieldPositions,
          to: te,
          decimal: decimal,
          group: group,
          dateOrder: dateOrder,
        };
      if (A === "fill") {
        let It = fillValue;
        if (direction === "value" && fillType === "number") {
          if (!fillValue.trim() || !Number.isFinite(Number(fillValue)))
            throw Error("Enter a finite fill number.");
          It = Number(fillValue);
        }
        if (direction === "value" && fillType === "boolean") {
          if (!/^(TRUE|FALSE)$/i.test(fillValue))
            throw Error("Enter TRUE or FALSE for Boolean fill.");
          It = fillValue.toUpperCase() === "TRUE";
        }
        return {
          kind: A,
          columns: fieldPositions,
          direction: direction,
          value: It,
          groupColumns: parseFieldPositions(groupColumns, true),
        };
      }
      const Dt = dr.find((It) => It.id === candidate);
      if (!Dt)
        throw Error(
          "Suggest patterns and explicitly choose one before adding the inferred step.",
        );
      return {
        kind: "pattern",
        name: J,
        pattern: Dt.pattern,
        examples: fr(),
        inferenceColumns: fieldPositions,
      };
    },
    Le = () => {
      Jt();
      try {
        const Qe = rt(),
          Dt = [...steps];
        (editing === null ? Dt.push(Qe) : (Dt[editing] = Qe),
          buildTransformationPreview(e.workbook, { ...request, steps: Dt }, xt),
          y(Dt),
          E(null),
          je([]),
          Je(""));
      } catch (error) {
        Ut(error);
      }
    },
    et = (Qe) => {
      Jt();
      const Dt = steps[Qe];
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
          Pe(Dt.groupColumns.map((groupColumn) => groupColumn + 1).join(", "))),
        Dt.kind === "pattern" &&
          (q(Dt.name),
          ht(Dt.examples.map((example) => example.row).join(", ")),
          Me(
            Dt.examples.map((example) => example.value).join(`
`),
          ),
          je([]),
          Je("")));
    },
    st = () => {
      Jt();
      try {
        const transformationPreview = buildTransformationPreview(
          e.workbook,
          request,
          xt,
        );
        (pt(0), nt(0), Ze({ plan: transformationPreview, key: Rr }));
      } catch (error) {
        Ut(error);
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
        } catch (error) {
          Ut(error);
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
        const Dt = e.workbook.sheets.find(
          (sheet) => sheet.id === Qe.sourceSheetId,
        );
        return Dt
          ? (assertSelectionInAuthorizedSheet(Dt, Qe.request.sourceRange, {
              allowedSheetIds: allowedSheetIds,
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
      onKeyDown: (event) => {
        var It, ar;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
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
          event.shiftKey && document.activeElement === Lr[0]
            ? (event.preventDefault(), (It = Lr.at(-1)) == null || It.focus())
            : !event.shiftKey &&
              document.activeElement === Lr.at(-1) &&
              (event.preventDefault(), (ar = Lr[0]) == null || ar.focus());
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
                      value: sourceSheetId,
                      onChange: (event) => u(event.target.value),
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
                      value: sourceRange,
                      onChange: (event) => c(event.target.value),
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
                      value: outputName,
                      onChange: (event) => m(event.target.value),
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
                      onChange: (event) => f(event.target.checked),
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
                        checked: checked,
                        onChange: (event) => g(event.target.checked),
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
                    steps.length,
                    "/12)",
                  ],
                }),
                !steps.length &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Add one or more steps, then preview the complete result.",
                    }),
                  }),
                i.jsx("ol", {
                  className: "fr-steps",
                  children: steps.map((step, Dt) =>
                    i.jsxs(
                      "li",
                      {
                        children: [
                          i.jsx("span", {
                            children: describeSplitFieldStep(step),
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
                                  const It = [...steps];
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
                                disabled: e.readOnly || Dt === steps.length - 1,
                                "aria-label": `Move step ${Dt + 1} down`,
                                onClick: () => {
                                  Jt();
                                  const It = [...steps];
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
                                    y(steps.filter((It, ar) => ar !== Dt)),
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
                  children:
                    editing === null
                      ? "Add a step"
                      : `Edit step ${editing + 1}`,
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Operation" }),
                    i.jsx("select", {
                      "aria-label": text("Field operation"),
                      value: A,
                      onChange: (event) => k(event.target.value),
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
                      value: columns,
                      onChange: (event) => I(event.target.value),
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
                        onChange: (event) => j(event.target.value),
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
                        onChange: (event) => V(event.target.value),
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
                            value: find,
                            onChange: (event) => B(event.target.value),
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
                            value: replacement,
                            onChange: (event) => C(event.target.value),
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
                            onChange: (event) => N(event.target.value),
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
                                    value: delimiters,
                                    onChange: (event) => L(event.target.value),
                                    lang: language("Delimiter characters"),
                                    "data-ui-label": "Delimiter characters",
                                  }),
                                ],
                              }),
                              i.jsx("button", {
                                onClick: () => L(delimiters + "	"),
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
                                    onChange: (event) =>
                                      U(event.target.checked),
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
                                value: breaks,
                                onChange: (event) => H(event.target.value),
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
                        onChange: (event) => q(event.target.value),
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
                            value: separator,
                            onChange: (event) => X$1(event.target.value),
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
                            onChange: (event) => K(event.target.checked),
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
                            onChange: (event) => ae(event.target.value),
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
                            value: decimal,
                            onChange: (event) => de(event.target.value),
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
                            value: group,
                            onChange: (event) => Se(event.target.value),
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
                            value: dateOrder,
                            onChange: (event) => ce(event.target.value),
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
                            value: direction,
                            onChange: (event) => se(event.target.value),
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
                            value: groupColumns,
                            onChange: (event) => Pe(event.target.value),
                            lang: language("Fill group fields"),
                            "data-ui-label": "Fill group fields",
                          }),
                        ],
                      }),
                      direction === "value" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Fill type" }),
                                i.jsxs("select", {
                                  "aria-label": text("Fill value type"),
                                  value: fillType,
                                  onChange: (event) => ye(event.target.value),
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
                                  value: fillValue,
                                  onChange: (event) => Ee(event.target.value),
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
                            value: exampleRows,
                            onChange: (event) => ht(event.target.value),
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
                            value: exampleValues,
                            onChange: (event) => Me(event.target.value),
                            rows: 3,
                            lang: language("Inference example values"),
                            "data-ui-label": "Inference example values",
                          }),
                        ],
                      }),
                      i.jsx("button", {
                        onClick: onClick,
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
                              value: candidate,
                              onChange: (event) => Je(event.target.value),
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
                      disabled: editing === null && steps.length >= 12,
                      children:
                        editing === null
                          ? "Add transformation step"
                          : "Save transformation step",
                    }),
                    editing !== null &&
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
                  disabled: e.readOnly || !steps.length || editing !== null,
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
                      } catch (error) {
                        Ut(error);
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
                      qt.stages.map((stage) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsx(LocalizedText, { children: "Step " }),
                              stage.step,
                              " (",
                              stage.kind,
                              "):",
                              " ",
                              stage.fields
                                .map((field, It) => `${It + 1}: ${field}`)
                                .join(" · "),
                              ".",
                              " ",
                              stage.changedCells,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "affected cells; ",
                              }),
                              stage.exceptions,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "new exceptions.",
                              }),
                            ],
                          },
                          stage.step,
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
                                                (source) =>
                                                  `${qt.sourceSheetName}!${source}`,
                                              )
                                              .join(", "),
                                          }),
                                          Qe.sources.map((source) =>
                                            i.jsxs(
                                              "pre",
                                              {
                                                children: [
                                                  source,
                                                  ":",
                                                  " ",
                                                  stringifyJsonValue(
                                                    qt.sources.find(
                                                      (Lr) =>
                                                        Lr.address === source,
                                                    )?.original?.raw ?? null,
                                                  ),
                                                ],
                                              },
                                              source,
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
                                  .map((cell, Dt) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              cell.row,
                                              " / ",
                                              cell.field,
                                            ],
                                          }),
                                          i.jsx("td", { children: cell.step }),
                                          i.jsx("td", {
                                            children: i.jsx("code", {
                                              children: stringifyJsonValue(
                                                cell.raw,
                                              ),
                                            }),
                                          }),
                                          i.jsx("td", {
                                            children: cell.sources.join(", "),
                                          }),
                                          i.jsx("td", {
                                            children: cell.reason,
                                          }),
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
                              onChange: (event) => Wt(event.target.checked),
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
                            children: Qe.request.steps.map((step, It) =>
                              i.jsx(
                                "li",
                                { children: describeSplitFieldStep(step) },
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
                editing !== null,
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
    const rangeAddress = parseRangeAddress(e),
      cellAddress = formatCellAddress(
        rangeAddress.firstRow,
        rangeAddress.firstCol,
      ),
      n = formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol);
    return cellAddress === n ? cellAddress : `${cellAddress}:${n}`;
  },
  isNonBlankCellEntry = (cell) =>
    cell.raw !== null || Object.keys(cell).some((r) => r !== "raw"),
  N3 = (cell) =>
    !cell.literal && typeof cell.raw == "string" && cell.raw.startsWith("="),
  buildAllowedSheetKey = (workbook, r) =>
    JSON.stringify(
      [
        ...(r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id)),
      ].sort(),
    );
function assertSortWorkbookAuthorized(workbook, sheetId, t) {
  if (t.readOnly)
    throw Error(
      "This workbook is read-only. Sorting cannot change its records.",
    );
  if (t.allowedSheetIds && !t.allowedSheetIds.includes(sheetId))
    throw Error("The sort source is outside the host-approved worksheets.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  if (sheetOrThrow.hidden || sheetOrThrow.veryHidden)
    throw Error("Unhide this worksheet before sorting its records.");
  return (assertSheetUnprotected(sheetOrThrow), sheetOrThrow);
}
function resolveSortSourceRange(workbook, sheetId, range, headers = false) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(range);
  if (
    rangeAddress.lastRow >= sheetOrThrow.rowCount ||
    rangeAddress.lastCol >= sheetOrThrow.colCount
  )
    throw Error("The sort selection is outside the worksheet.");
  let firstRow = rangeAddress.firstRow,
    lastRow = rangeAddress.lastRow,
    firstCol = rangeAddress.firstCol,
    lastCol = rangeAddress.lastCol;
  const c = Object.entries(sheetOrThrow.cells)
    .filter(([, m]) => isNonBlankCellEntry(m))
    .map(([m]) => _t(m));
  if (firstRow === lastRow && firstCol === lastCol) {
    const m = sheetOrThrow.tables
      ?.map((table) => ({
        bounds: parseRangeAddress(table.range),
        totals: table.showTotals,
      }))
      .find(
        (v) =>
          firstRow >= v.bounds.firstRow &&
          firstRow <= v.bounds.lastRow &&
          firstCol >= v.bounds.firstCol &&
          firstCol <= v.bounds.lastCol,
      );
    if (m)
      ((firstRow = m.bounds.firstRow),
        (lastRow = m.bounds.lastRow - (m.totals ? 1 : 0)),
        (firstCol = m.bounds.firstCol),
        (lastCol = m.bounds.lastCol));
    else {
      const v = new Set(c.map((g) => g.row));
      if (!v.has(firstRow))
        throw Error("Select a populated record or an explicit range to sort.");
      for (; firstRow > 0 && v.has(firstRow - 1);) firstRow--;
      for (; lastRow + 1 < sheetOrThrow.rowCount && v.has(lastRow + 1);)
        lastRow++;
    }
  }
  for (const m of c)
    m.row >= firstRow &&
      m.row <= lastRow &&
      ((firstCol = Math.min(firstCol, m.col)),
      (lastCol = Math.max(lastCol, m.col)));
  const sortRangeAddress = formatSortRangeAddress(
      `${formatCellAddress(firstRow, firstCol)}:${formatCellAddress(lastRow, lastCol)}`,
    ),
    f = formatSortRangeAddress(range);
  return {
    selectedRange: f,
    expandedRange: sortRangeAddress,
    expansionRequired: f !== sortRangeAddress,
    firstDataRow: firstRow + (headers ? 1 : 0),
    lastRow: lastRow,
  };
}
function normalizeSortSpecification(workbook, r) {
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
  for (const key of t.keys) {
    if (
      !Number.isInteger(key.column) ||
      key.column < 0 ||
      key.column >= findSheetOrThrow(workbook, r.sheetId).colCount ||
      n.has(key.column) ||
      !["ascending", "descending"].includes(key.direction) ||
      !["auto", "number", "date", "text"].includes(key.type)
    )
      throw Error(
        "Choose distinct valid sort columns, directions and data types.",
      );
    if ((n.add(key.column), key.customOrder !== undefined)) {
      const a = (s) =>
        t.caseSensitive
          ? s.normalize("NFC")
          : s.normalize("NFC").toLocaleLowerCase("en-US");
      if (
        !Array.isArray(key.customOrder) ||
        !key.customOrder.length ||
        key.customOrder.length > sortEngineLimits.customItems ||
        key.type !== "text" ||
        key.customOrder.some(
          (s) => typeof s != "string" || !s.length || s.length > 255,
        ) ||
        new Set(key.customOrder.map(a)).size !== key.customOrder.length
      )
        throw Error(
          "Custom orders require Text, 1–100 distinct nonempty labels of at most 255 characters.",
        );
    }
  }
  return t;
}
function buildSortPlan(workbook, r, t) {
  const rangeAddress = parseRangeAddress(t.expandedRange),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    datePolicy = normalizeDatePolicy(workbook.datePolicy);
  return Array.from({ length: t.lastRow - t.firstDataRow + 1 }, (s, d) => {
    const row = t.firstDataRow + d,
      keys = r.keys.map((key) => {
        if (
          key.column < rangeAddress.firstCol ||
          key.column > rangeAddress.lastCol
        )
          throw Error(
            "Every sort key must be inside the complete record range.",
          );
        const cellAddress = formatCellAddress(row, key.column),
          f = findSheetOrThrow(workbook, r.sheetId).cells[cellAddress],
          p = workbookWithSettings(r.sheetId, cellAddress);
        if (
          f &&
          N3(f) &&
          typeof p == "string" &&
          ["#PENDING!", "#TIMEOUT!", "#LIMIT!", "#BLOCKED!"].includes(p)
        )
          throw Error(
            `Calculate ${cellAddress} successfully before sorting. Its current result is ${p}.`,
          );
        const m =
          !!f &&
          N3(f) &&
          workbookWithSettings.getTypedValue(r.sheetId, cellAddress).type ===
            "error";
        let compare = p;
        if (p !== null && p !== "" && !m) {
          if (
            (key.type === "text" && (compare = String(p)),
            key.type === "number")
          ) {
            if (typeof p == "number") compare = p;
            else if (
              typeof p == "string" &&
              /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(p.trim())
            )
              compare = Number(p);
            else
              throw Error(
                `${cellAddress} cannot be interpreted as a Number. Choose Automatic or correct the source.`,
              );
            if (!Number.isFinite(compare))
              throw Error(
                `${cellAddress} is outside the finite numeric range.`,
              );
          }
          if (
            key.type === "date" &&
            ((compare =
              typeof p == "number"
                ? p
                : typeof p == "string"
                  ? (parseDateText(p, datePolicy) ?? null)
                  : null),
            compare === null ||
              typeof compare != "number" ||
              compare < 0 ||
              compare >= dateSystemMaxSerial(datePolicy.dateSystem) + 1)
          )
            throw Error(
              `${cellAddress} is not a valid date under the workbook's date policy.`,
            );
        }
        return { value: p, compare: compare, error: m };
      });
    return { row: row, keys: keys };
  });
}
function assertSortSafeFormulas(workbook, r, t) {
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
      matrixResultFunctionNames.some((matrixResultFunctionName) =>
        new RegExp(
          `\\b${matrixResultFunctionName.replaceAll(".", "\\.")}\\s*\\(`,
          "i",
        ).test(u),
      )
    )
      throw Error(
        "Sorting cannot safely relocate array/spill formulas or existing reference-error tokens. Resolve or separate those formulas first.",
      );
    a += (u.match(/:/g) ?? []).length;
  };
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells)) N3(cell) && s(cell.raw);
    for (const u of [
      ...(sheet.validations ?? []),
      ...(sheet.conditionalFormats ?? []),
    ])
      u.formula && s(u.formula);
  }
  for (const d of workbook.definedNames ?? []) s(d.refersTo);
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
    workbook.sheets.some((sheet) => !t.allowedSheetIds.includes(sheet.id))
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
