import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { textCleanupFunctionNames, cn } from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { describeInferredTextRule } from "./16-transformation-recipe-engine.js";
import {
  createTextFunctionField,
  previewTextCleanup,
  getTextFunctionSpec,
  createTextCleanupOutput,
  stringifyEscapedJson,
} from "./42-combine-populations-and-text-columns.js";

function TextCleanupDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(false),
    [sourceRange, d] = ne.useState(e.range),
    [outputName, l] = ne.useState("Cleaned text"),
    [checked, h] = ne.useState(false),
    [steps, p] = ne.useState([
      getTextFunctionSpec("CLEAN"),
      getTextFunctionSpec("TRIM"),
    ]),
    [m, v] = ne.useState("SUBSTITUTE"),
    [g, b] = ne.useState(null),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(0),
    [k, F] = ne.useState(e.workbook.textCleanups?.at(-1)?.id ?? ""),
    [I, O] = ne.useState(0),
    allowedSheetIds = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    z = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && checked,
      allowedSheetIds: allowedSheetIds,
      allowNewSheet: e.allowNewSheet ?? true,
    },
    V = JSON.stringify({
      sourceRange: sourceRange,
      outputName: outputName,
      steps: steps,
      options: z,
      selection: [e.sheetId, e.range],
    }),
    D = JSON.stringify(e.workbook),
    B = g?.key === V && g.plan.snapshot === D ? g.plan : null,
    _ = ne.useRef({ p: e, key: V, options: z });
  _.current = { p: e, key: V, options: z };
  const C = () => {
    (b(null), w(""), (a.current = false), A(0));
  };
  (ne.useEffect(() => {
    (d(e.range), b(null), w(""), (a.current = false));
  }, [e.sheetId, e.range]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        o.current?.querySelector("input:not(:disabled),button")?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const T = (H) =>
      w(
        H instanceof Error
          ? H.message
          : "Cleanup failed. Review the source and steps.",
      ),
    onClick = () => {
      C();
      try {
        const H = steps.map((step) => ({
          function: step.name,
          args: step.fields
            .filter((field) => !(field.optional && field.value === ""))
            .map((X) =>
              X.kind === "number"
                ? X.value.trim()
                  ? Number(X.value)
                  : NaN
                : X.kind === "boolean"
                  ? X.value === "TRUE"
                  : X.value,
            ),
        }));
        b({
          key: V,
          plan: previewTextCleanup(
            e.workbook,
            {
              sourceSheetId: e.sheetId,
              sourceRange: sourceRange,
              outputName: outputName,
              steps: H,
            },
            z,
          ),
        });
      } catch (error) {
        T(error);
      }
    },
    R = () => {
      if (!(!B || a.current))
        try {
          if (_.current.key !== g?.key)
            throw Error("The cleanup settings changed. Preview again.");
          const textCleanupOutput = createTextCleanupOutput(
            _.current.p.workbook,
            B,
            _.current.options,
          );
          ((a.current = true),
            _.current.p.onWorkbookChange(
              textCleanupOutput,
              `Text cleanup: ${B.sourceSheetName}!${B.request.sourceRange} → ${B.request.outputName}`,
            ),
            b(null),
            F(textCleanupOutput.textCleanups.at(-1).id),
            O(0));
        } catch (error) {
          T(error);
        }
    },
    L = (e.workbook.textCleanups ?? []).filter((H) => {
      const sheet = e.workbook.sheets.find(
        (sheet) => sheet.id === H.sourceSheetId,
      );
      return (
        (e.allowedSheetIds === undefined ||
          allowedSheetIds.includes(H.sourceSheetId)) &&
        (z.includeHidden ||
          !sheet ||
          (!sheet.hidden &&
            !sheet.veryHidden &&
            !cn(sheet, "row").length &&
            !cn(sheet, "column").length))
      );
    }),
    $ = L.find((H) => H.id === k) ?? L.at(-1);
  return i.jsx("div", {
    className: "ttc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      className: "ttc-dialog cleanup-dialog",
      ref: o,
      onKeyDown: (event) => {
        var X, J;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const q = [
            ...(o.current?.querySelectorAll("button,input,select,summary") ??
              []),
          ].filter(
            (W) =>
              !W.matches(":disabled") &&
              !W.closest("details:not([open]) > :not(summary)"),
          );
          event.shiftKey && document.activeElement === q[0]
            ? (event.preventDefault(), (X = q.at(-1)) == null || X.focus())
            : !event.shiftKey &&
              document.activeElement === q.at(-1) &&
              (event.preventDefault(), (J = q[0]) == null || J.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: `${n}-title`,
              children: i.jsx(LocalizedText, { children: "Clean up text" }),
            }),
            i.jsx("button", {
              "aria-label": text("Close text cleanup"),
              onClick: e.onClose,
              lang: language("Close text cleanup"),
              "data-ui-label": "Close text cleanup",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Review text changes before creating a separate worksheet. Original cells, evidence notes and source identities stay in place.",
          }),
        }),
        i.jsxs("fieldset", {
          className: "ttc-fields",
          disabled: e.readOnly || e.allowNewSheet === false,
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, { children: "Source and output" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Cleanup source range"),
                  value: sourceRange,
                  onChange: (event) => {
                    (d(event.target.value), C());
                  },
                  lang: language("Cleanup source range"),
                  "data-ui-label": "Cleanup source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "New worksheet name" }),
                i.jsx("input", {
                  value: outputName,
                  onChange: (event) => {
                    (l(event.target.value), C());
                  },
                }),
              ],
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "ttc-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (h(event.target.checked), C());
                    },
                  }),
                  i.jsx(LocalizedText, {
                    children: "Include hidden source cells",
                  }),
                ],
              }),
          ],
        }),
        e.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "This workbook is read-only. Saved cleanup evidence can still be reviewed.",
            }),
          }),
        i.jsxs("fieldset", {
          disabled: e.readOnly || e.allowNewSheet === false,
          className: "cleanup-steps",
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, { children: "Steps in order" }),
            }),
            steps.map((step, Q) =>
              i.jsxs(
                "div",
                {
                  className: "cleanup-step",
                  children: [
                    i.jsxs("strong", { children: [Q + 1, ". ", step.name] }),
                    step.fields.map((field, J) =>
                      i.jsxs(
                        "label",
                        {
                          children: [
                            field.label,
                            field.kind === "boolean"
                              ? i.jsxs("select", {
                                  "aria-label": `Step ${Q + 1} ${field.label}`,
                                  value: field.value,
                                  onChange: (event) => {
                                    (p(
                                      steps.map((step, K) =>
                                        K === Q
                                          ? {
                                              ...step,
                                              fields: step.fields.map(
                                                (field, ae) =>
                                                  ae === J
                                                    ? {
                                                        ...field,
                                                        value:
                                                          event.target.value,
                                                      }
                                                    : field,
                                              ),
                                            }
                                          : step,
                                      ),
                                    ),
                                      C());
                                  },
                                  children: [
                                    i.jsx("option", { children: "TRUE" }),
                                    i.jsx("option", { children: "FALSE" }),
                                  ],
                                })
                              : i.jsx("input", {
                                  "aria-label": `Step ${Q + 1} ${field.label}`,
                                  value: field.value,
                                  onChange: (event) => {
                                    (p(
                                      steps.map((step, K) =>
                                        K === Q
                                          ? {
                                              ...step,
                                              fields: step.fields.map(
                                                (field, ae) =>
                                                  ae === J
                                                    ? {
                                                        ...field,
                                                        value:
                                                          event.target.value,
                                                      }
                                                    : field,
                                              ),
                                            }
                                          : step,
                                      ),
                                    ),
                                      C());
                                  },
                                }),
                          ],
                        },
                        J,
                      ),
                    ),
                    i.jsx("button", {
                      "aria-label": `Move step ${Q + 1} up`,
                      disabled: Q === 0,
                      onClick: () => {
                        const X = [...steps];
                        (([X[Q - 1], X[Q]] = [X[Q], X[Q - 1]]), p(X), C());
                      },
                      children: "↑",
                    }),
                    i.jsx("button", {
                      "aria-label": `Remove step ${Q + 1}`,
                      onClick: () => {
                        (p(steps.filter((step, J) => Q !== J)), C());
                      },
                      children: i.jsx(LocalizedText, { children: "Remove" }),
                    }),
                  ],
                },
                Q,
              ),
            ),
            i.jsxs("div", {
              className: "cleanup-add",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Add function" }),
                    i.jsx("select", {
                      value: m,
                      onChange: (event) => v(event.target.value),
                      children: textCleanupFunctionNames.map(
                        (textCleanupFunctionName) =>
                          i.jsx(
                            "option",
                            { children: textCleanupFunctionName },
                            textCleanupFunctionName,
                          ),
                      ),
                    }),
                  ],
                }),
                i.jsx("button", {
                  disabled: steps.length >= 12,
                  onClick: () => {
                    (p([...steps, getTextFunctionSpec(m)]), C());
                  },
                  children: i.jsx(LocalizedText, { children: "Add step" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (p([
                      ...steps,
                      {
                        name: "SUBSTITUTE",
                        fields: [
                          createTextFunctionField("Old text", " "),
                          createTextFunctionField("New text", " "),
                          createTextFunctionField(
                            "Occurrence (blank = all)",
                            "",
                            "number",
                            true,
                          ),
                        ],
                      },
                    ]),
                      C());
                  },
                  disabled: steps.length >= 12,
                  children: i.jsx(LocalizedText, {
                    children: "Replace nonbreaking spaces",
                  }),
                }),
              ],
            }),
          ],
        }),
        i.jsx("p", {
          className: "ttc-boundary",
          children: i.jsx(LocalizedText, {
            children:
              "TRIM removes ordinary spaces; CLEAN removes ASCII control characters. Nonbreaking spaces need an explicit SUBSTITUTE step before TRIM. Character positions count Unicode code points. VALUE uses US numeric notation; NUMBERVALUE accepts explicit decimal and group separators. Numeric conversion removes leading zeros.",
          }),
        }),
        i.jsx("button", {
          onClick: onClick,
          disabled: e.readOnly || e.allowNewSheet === false,
          children: i.jsx(LocalizedText, { children: "Preview cleanup" }),
        }),
        y && i.jsx("p", { className: "ttc-error", role: "alert", children: y }),
        g &&
          !B &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook or settings changed. Preview again before applying.",
            }),
          }),
        B &&
          i.jsxs("section", {
            className: "ttc-preview",
            "aria-label": text("Text cleanup preview"),
            lang: language("Text cleanup preview"),
            "data-ui-label": "Text cleanup preview",
            children: [
              i.jsxs("h3", {
                children: [
                  B.rows.length,
                  " ",
                  i.jsx(LocalizedText, { children: "results · " }),
                  B.errors.length,
                  " ",
                  i.jsx(LocalizedText, { children: "errors" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Source revision " }),
                  e.workbook.revision,
                  i.jsx(LocalizedText, {
                    children:
                      ". Quoted values expose spaces; nonbreaking spaces appear as \\u00a0.",
                  }),
                ],
              }),
              i.jsx("code", { children: B.formulas[0] }),
              i.jsx("div", {
                className: "ttc-scroll",
                children: i.jsxs("table", {
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source cell",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Original evidence value",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Result",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Type / error",
                            }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: B.rows.slice(E * 20, E * 20 + 20).map((H) =>
                        i.jsxs(
                          "tr",
                          {
                            children: [
                              i.jsx("td", { children: H.sourceAddress }),
                              i.jsx("td", {
                                children: i.jsx("code", {
                                  children: stringifyEscapedJson(
                                    H.original?.raw ?? null,
                                  ),
                                }),
                              }),
                              i.jsx("td", {
                                children: i.jsx("code", {
                                  children: stringifyEscapedJson(H.result),
                                }),
                              }),
                              i.jsx("td", {
                                children:
                                  B.errors.find(
                                    (error) =>
                                      error.sourceAddress === H.sourceAddress,
                                  )?.error ?? typeof H.result,
                              }),
                            ],
                          },
                          H.sourceAddress,
                        ),
                      ),
                    }),
                  ],
                }),
              }),
              i.jsxs("div", {
                className: "ttc-pages",
                children: [
                  i.jsx("button", {
                    disabled: E === 0,
                    onClick: () => A(E - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous results",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      E + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      Math.ceil(B.rows.length / 20),
                    ],
                  }),
                  i.jsx("button", {
                    disabled: (E + 1) * 20 >= B.rows.length,
                    onClick: () => A(E + 1),
                    children: i.jsx(LocalizedText, {
                      children: "Next results",
                    }),
                  }),
                ],
              }),
              B.errors.length > 0 &&
                i.jsx("p", {
                  role: "alert",
                  children: i.jsx(LocalizedText, {
                    children:
                      "Resolve all errors before applying. No partial output will be created.",
                  }),
                }),
            ],
          }),
        L.length > 0 &&
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(LocalizedText, {
                    children: "Saved text cleanup evidence (",
                  }),
                  L.length,
                  ")",
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Cleanup record" }),
                  i.jsx("select", {
                    value: $?.id ?? "",
                    onChange: (event) => {
                      (F(event.target.value), O(0));
                    },
                    children: L.map((H) =>
                      i.jsxs(
                        "option",
                        {
                          value: H.id,
                          children: [
                            H.outputSheetName,
                            " · source revision ",
                            H.sourceRevision,
                            " ·",
                            " ",
                            H.createdAt,
                          ],
                        },
                        H.id,
                      ),
                    ),
                  }),
                ],
              }),
              $ &&
                i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("p", {
                      children: [
                        i.jsx(LocalizedText, { children: "Source: " }),
                        $.sourceSheetName,
                        " ",
                        i.jsx(LocalizedText, { children: "· workbook" }),
                        " ",
                        $.workbookId,
                        " ",
                        i.jsx(LocalizedText, { children: "· revision " }),
                        $.sourceRevision,
                        i.jsx(LocalizedText, {
                          children:
                            ". Addresses and values are snapshots from the cleanup, even if cells later move or are deleted.",
                        }),
                      ],
                    }),
                    i.jsx("code", { children: stringifyEscapedJson($.steps) }),
                    i.jsx("div", {
                      className: "ttc-scroll",
                      children: i.jsxs("table", {
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Source at cleanup",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Original value and note",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Result at cleanup",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Original file row",
                                  }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: $.rows
                              .slice(I * 20, I * 20 + 20)
                              .map((H) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("td", {
                                        children: H.sourceAddress,
                                      }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: stringifyEscapedJson(
                                            H.original ?? { raw: null },
                                          ),
                                        }),
                                      }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: stringifyEscapedJson(
                                            H.result,
                                          ),
                                        }),
                                      }),
                                      i.jsx("td", {
                                        children: H.origin
                                          ? `${H.origin.sourceName} · ${H.origin.sourceSheet} · row ${H.origin.originalRow}`
                                          : "Editor value snapshot",
                                      }),
                                    ],
                                  },
                                  H.sourceAddress,
                                ),
                              ),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("button", {
                      disabled: I === 0,
                      onClick: () => O(I - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous evidence",
                      }),
                    }),
                    i.jsx("button", {
                      disabled: (I + 1) * 20 >= $.rows.length,
                      onClick: () => O(I + 1),
                      children: i.jsx(LocalizedText, {
                        children: "Next evidence",
                      }),
                    }),
                  ],
                }),
            ],
          }),
        i.jsx("p", {
          className: "ttc-boundary",
          children: i.jsx(LocalizedText, {
            children:
              "Up to 500 ordinary source cells, 200,000 source characters and 12 steps. Keep a native JSON backup to retain cleanup history; CSV/XLSX do not preserve that history.",
          }),
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
            i.jsx("button", {
              className: "ttc-primary",
              disabled: !B || !!B.errors.length || e.readOnly || a.current,
              onClick: R,
              children: i.jsx(LocalizedText, {
                children: "Create cleaned worksheet",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const textCleanupStepLabels = [
    ["trim", "Trim whitespace"],
    ["clean", "Remove control characters"],
    ["case", "Change case"],
    ["replace", "Replace text"],
    ["split", "Split fields"],
    ["concat", "Concatenate fields"],
    ["convert", "Convert type"],
    ["fill", "Fill gaps"],
    ["select", "Select and reorder fields"],
    ["pattern", "Infer from examples"],
  ],
  parseFieldPositions = (e, r = false) => {
    if (!e.trim() && r) return [];
    if (!/^\s*\d+(?:\s*,\s*\d+)*\s*$/.test(e))
      throw Error("Enter field positions separated by commas, such as 1, 3.");
    const t = e.split(",").map((n) => Number(n) - 1);
    if (t.some((n) => n < 0 || n > 49) || new Set(t).size !== t.length)
      throw Error("Choose distinct field positions from 1 to 50.");
    return t;
  };
function describeSplitFieldStep(step) {
  if (step.kind === "pattern")
    return `Infer “${step.name}”: ${describeInferredTextRule(step.pattern)} (${step.examples.length} examples)`;
  if (step.kind === "split")
    return `Split field ${step.column + 1}: ${step.mode === "fixed" ? "breaks " + step.breaks.join(", ") : "delimiters " + JSON.stringify(step.delimiters) + (step.collapse ? ", collapse consecutive" : "")}`;
  const r = step.columns.map((column) => column + 1).join(", ");
  return `${textCleanupStepLabels.find((textCleanupStepLabel) => textCleanupStepLabel[0] === step.kind)?.[1]} · fields ${r}${step.kind === "convert" ? ` → ${step.to} (decimal ${step.decimal}, groups ${step.group || "none"}, dates ${step.dateOrder})` : step.kind === "fill" ? ` → ${step.direction}${step.groupColumns.length ? " within fields " + step.groupColumns.map((groupColumn) => groupColumn + 1).join(", ") : ""}` : step.kind === "concat" ? ` → ${step.name}; separator ${JSON.stringify(step.separator)}` : step.kind === "case" || step.kind === "trim" ? ` → ${step.mode}` : step.kind === "replace" ? ` ${JSON.stringify(step.find)} → ${JSON.stringify(step.replacement)}` : ""}`;
}

export {
  textCleanupStepLabels,
  describeSplitFieldStep,
  parseFieldPositions,
  TextCleanupDialog,
};
