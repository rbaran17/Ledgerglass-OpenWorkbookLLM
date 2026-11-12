import { failWithMessage } from "./03-workbook-model.js";
import {
  buildCompositeKey,
  isFormulaInputValue,
} from "./11-local-workbook-store-and-versions.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  normalizeHiddenIntervals,
  resolveValidationInspection,
  parsePrintTitles,
  resolveMovedRange,
  mapIndexAfterMove,
  summarizeValidationInspection,
  calculateWorkbookWithSettings,
  normalizeDataValidationRules,
  _t,
  checkStructuralEditImpact,
  cn,
  parseRangeAddress,
  shiftFormulaWatchesOnAxisEdit,
  shiftPrintSettingsOnAxisEdit,
  assertSheetUnprotected,
  resolveDefinedNameValue,
  validateRowColumnMove,
  normalizeWorkbookInput,
  shiftReviewTargetsOnAxisEdit,
  measureChartSelection,
  columnIndexToLetters,
  mapRangeAfterMove,
  applyAxisStructureEdit,
  shiftFormulaReferences,
  matrixResultFunctionNames,
  rewriteStructuredReferences,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { inspectFormulaDependencies } from "./12-consolidation-and-dependency-inspection.js";

function DataValidationDialog({
  workbook,
  sheetId,
  range,
  readOnly,
  onClose,
  onWorkbookChange,
  onNavigate,
}) {
  const { text, language } = useWorkbookMessages(),
    l = workbook.sheets.find((ce) => ce.id === sheetId),
    [c, h] = ne.useState(),
    [f, p] = ne.useState(range),
    [m, v] = ne.useState("list"),
    [g, b] = ne.useState("between"),
    [y, w] = ne.useState("Pending,Complete,Reviewed"),
    [E, A] = ne.useState("0"),
    [k, F] = ne.useState("100"),
    [I, O] = ne.useState(`=LEN(${range.split(":")[0]})>0`),
    [j, z] = ne.useState(true),
    [V, D] = ne.useState(true),
    [B, _] = ne.useState(true),
    [C, T] = ne.useState(""),
    [N, R] = ne.useState(""),
    [L, $] = ne.useState("stop"),
    [U, G] = ne.useState("Invalid value"),
    [H, Q] = ne.useState("Enter a value that satisfies this rule."),
    [X$1, J] = ne.useState(""),
    [q, W] = ne.useState(),
    K = ne.useRef(null);
  ne.useEffect(() => {
    const ce = document.activeElement;
    return (
      K.current?.querySelector("button,input")?.focus(),
      () => ce?.focus()
    );
  }, []);
  const te = readOnly || !!l.protection?.enabled,
    ae = (ce) => {
      J("");
      try {
        ce();
      } catch (ge) {
        J(
          ge instanceof Error
            ? ge.message
            : "Validation could not be completed.",
        );
      }
    },
    oe = () =>
      ae(() => {
        if (te) throw Error("This worksheet is read-only or protected.");
        if (
          (assertSheetUnprotected(l),
          !["list", "custom"].includes(m) && (!E.trim() || !k.trim()))
        )
          throw Error("Enter finite validation bounds.");
        const ce = {
            id: c ?? crypto.randomUUID(),
            range: f,
            kind: m,
            operator: g,
            min: Number(E),
            max: Number(k),
            list:
              m === "list"
                ? y.trim().startsWith("[")
                  ? JSON.parse(y)
                  : y.split(",").map((se) => se.trim())
                : [],
            allowBlank: j,
            message: H,
            errorStyle: L,
            errorTitle: U,
            showDropdown: V,
            showInputMessage: B,
            inputTitle: C,
            inputMessage: N,
            ...(m === "custom" ? { formula: I } : {}),
          },
          ge = normalizeDataValidationRules(
            c
              ? (l.validations ?? []).map((se) => (se.id === c ? ce : se))
              : [...(l.validations ?? []), ce],
          );
        (onWorkbookChange(
          {
            ...workbook,
            sheets: workbook.sheets.map((se) =>
              se.id === sheetId ? { ...se, validations: ge } : se,
            ),
          },
          c ? "Update data validation" : "Add data validation",
        ),
          onClose());
      }),
    de = (ce) => {
      (h(ce.id),
        p(ce.range),
        v(ce.kind),
        b(ce.operator),
        w(JSON.stringify(ce.list)),
        A(String(ce.min)),
        F(String(ce.max)),
        O(ce.formula ?? "=TRUE()"),
        z(ce.allowBlank),
        D(ce.showDropdown !== false),
        _(ce.showInputMessage !== false),
        T(ce.inputTitle ?? ""),
        R(ce.inputMessage ?? ""),
        $(ce.errorStyle ?? "stop"),
        G(ce.errorTitle ?? "Invalid value"),
        Q(ce.message));
    };
  return i.jsx("div", {
    className: "wt-backdrop",
    onKeyDown: (ce) => {
      if (
        (ce.key === "Escape" && (ce.stopPropagation(), onClose()),
        ce.key === "Tab")
      ) {
        const se = [
            ...(K.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled)",
            ) ?? []),
          ],
          we = se[0],
          Ee = se.at(-1);
        ce.shiftKey && document.activeElement === we
          ? (ce.preventDefault(), Ee == null || Ee.focus())
          : !ce.shiftKey &&
            document.activeElement === Ee &&
            (ce.preventDefault(), we == null || we.focus());
      }
    },
    children: i.jsxs("div", {
      className: "wt-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": text("Data Validation"),
      ref: K,
      style: { width: "min(760px, 100%)" },
      lang: language("Data Validation"),
      "data-ui-label": "Data Validation",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children: i.jsx(LocalizedText, { children: "Data Validation" }),
            }),
            i.jsx("button", {
              "aria-label": text("Close data validation"),
              onClick: onClose,
              lang: language("Close data validation"),
              "data-ui-label": "Close data validation",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wt-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Rules apply to typed values, paste, SDK edits and approved assistant changes. Stop rejects the entire batch; Warning and Information require explicit review. Existing values are retained.",
              }),
            }),
            i.jsx("button", {
              onClick: () =>
                ae(() =>
                  W({
                    value: summarizeValidationInspection(workbook, sheetId),
                    snapshot: JSON.stringify(workbook),
                  }),
                ),
              children: i.jsx(LocalizedText, {
                children: "Inspect existing values",
              }),
            }),
            q &&
              i.jsxs("section", {
                "aria-label": text("Existing validation findings"),
                lang: language("Existing validation findings"),
                "data-ui-label": "Existing validation findings",
                children: [
                  i.jsx("button", {
                    disabled: q.snapshot !== JSON.stringify(workbook),
                    onClick: () =>
                      ae(() => {
                        const ce = resolveValidationInspection(
                            workbook,
                            q.value,
                          ),
                          ge = URL.createObjectURL(
                            new Blob([ce], { type: "application/json" }),
                          ),
                          se = document.createElement("a");
                        ((se.href = ge),
                          (se.download = "Validation-findings.json"),
                          se.click(),
                          setTimeout(() => URL.revokeObjectURL(ge), 1e3));
                      }),
                    children: i.jsx(LocalizedText, {
                      children: "Download validation findings",
                    }),
                  }),
                  i.jsx("p", {
                    role: "status",
                    children:
                      q.snapshot === JSON.stringify(workbook)
                        ? q.value.message
                        : "Workbook changed. Inspect existing values again.",
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Inspection recalculates locally, including in Manual mode; it does not publish current calculation results or change source data.",
                    }),
                  }),
                  q.snapshot === JSON.stringify(workbook) &&
                    i.jsx("div", {
                      style: { maxHeight: 200, overflow: "auto" },
                      children: q.value.issues.map((ce, ge) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsxs("button", {
                                onClick: () => {
                                  (onNavigate?.(ce.sheetId, ce.address),
                                    onNavigate && onClose());
                                },
                                disabled: !onNavigate,
                                children: [ce.sheetName, "!", ce.address],
                              }),
                              " ",
                              ce.style,
                              ": ",
                              ce.message,
                              " ",
                              ce.reason,
                            ],
                          },
                          ge,
                        ),
                      ),
                    }),
                ],
              }),
            i.jsxs("details", {
              children: [
                i.jsxs("summary", {
                  children: [
                    l.validations?.length ?? 0,
                    " ",
                    i.jsx(LocalizedText, { children: "existing rules" }),
                  ],
                }),
                l.validations?.map((ce) =>
                  i.jsxs(
                    "div",
                    {
                      className: "wt-existing",
                      children: [
                        i.jsxs("span", {
                          children: [
                            ce.range,
                            " · ",
                            ce.kind,
                            " · ",
                            ce.errorStyle ?? "stop",
                          ],
                        }),
                        i.jsxs("button", {
                          disabled: te,
                          onClick: () => de(ce),
                          children: [
                            i.jsx(LocalizedText, { children: "Edit " }),
                            ce.range,
                          ],
                        }),
                        i.jsxs("button", {
                          disabled: te,
                          onClick: () =>
                            ae(() => {
                              (assertSheetUnprotected(l),
                                onWorkbookChange(
                                  {
                                    ...workbook,
                                    sheets: workbook.sheets.map((ge) =>
                                      ge.id === sheetId
                                        ? {
                                            ...ge,
                                            validations: ge.validations?.filter(
                                              (we) => we.id !== ce.id,
                                            ),
                                          }
                                        : ge,
                                    ),
                                  },
                                  "Remove data validation",
                                ),
                                onClose());
                            }),
                          children: [
                            i.jsx(LocalizedText, { children: "Remove " }),
                            ce.range,
                          ],
                        }),
                      ],
                    },
                    ce.id,
                  ),
                ),
              ],
            }),
            i.jsxs("fieldset", {
              disabled: te,
              style: { border: 0, padding: 0, display: "grid", gap: 12 },
              children: [
                i.jsx("legend", { children: c ? "Edit rule" : "New rule" }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Apply to" }),
                    i.jsx("input", {
                      "aria-label": text("Validation range"),
                      value: f,
                      onChange: (ce) => p(ce.target.value),
                      lang: language("Validation range"),
                      "data-ui-label": "Validation range",
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "wt-columns",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Allow" }),
                        i.jsx("select", {
                          "aria-label": text("Validation type"),
                          value: m,
                          onChange: (ce) => v(ce.target.value),
                          lang: language("Validation type"),
                          "data-ui-label": "Validation type",
                          children: [
                            ["list", "List"],
                            ["number", "Decimal"],
                            ["integer", "Whole number"],
                            ["date", "Date"],
                            ["time", "Time"],
                            ["textLength", "Text length"],
                            ["custom", "Custom formula"],
                          ].map(([ce, ge]) =>
                            i.jsx("option", { value: ce, children: ge }, ce),
                          ),
                        }),
                      ],
                    }),
                    !["list", "custom"].includes(m) &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Condition" }),
                          i.jsx("select", {
                            "aria-label": text("Validation operator"),
                            value: g,
                            onChange: (ce) => b(ce.target.value),
                            lang: language("Validation operator"),
                            "data-ui-label": "Validation operator",
                            children: [
                              ["between", "Between"],
                              ["notBetween", "Not between"],
                              ["equal", "Equal to"],
                              ["notEqual", "Not equal to"],
                              ["greater", "Greater than"],
                              ["greaterEqual", "Greater than or equal"],
                              ["less", "Less than"],
                              ["lessEqual", "Less than or equal"],
                            ].map(([ce, ge]) =>
                              i.jsx("option", { value: ce, children: ge }, ce),
                            ),
                          }),
                        ],
                      }),
                  ],
                }),
                m === "list"
                  ? i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children:
                            "Values, separated by commas (or a JSON array for exact text)",
                        }),
                        i.jsx("input", {
                          "aria-label": text("Validation values"),
                          value: y,
                          onChange: (ce) => w(ce.target.value),
                          lang: language("Validation values"),
                          "data-ui-label": "Validation values",
                        }),
                      ],
                    })
                  : m === "custom"
                    ? i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Formula relative to the top-left cell",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Validation formula"),
                            value: I,
                            onChange: (ce) => O(ce.target.value),
                            lang: language("Validation formula"),
                            "data-ui-label": "Validation formula",
                          }),
                          i.jsx("small", {
                            children: i.jsx(LocalizedText, {
                              children:
                                "Use TRUE/FALSE or a finite number. Relative references move with each cell; $ references stay fixed. Formula errors stop entry.",
                            }),
                          }),
                        ],
                      })
                    : i.jsxs("div", {
                        className: "wt-columns",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Minimum / comparison value",
                              }),
                              i.jsx(LocalizedNumberField, {
                                "aria-label": text("Validation values"),
                                value: E,
                                onValueChange: A,
                                lang: language("Validation values"),
                                "data-ui-label": "Validation values",
                              }),
                            ],
                          }),
                          ["between", "notBetween"].includes(g) &&
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Maximum" }),
                                i.jsx(LocalizedNumberField, {
                                  "aria-label": text("Validation maximum"),
                                  value: k,
                                  onValueChange: F,
                                  lang: language("Validation maximum"),
                                  "data-ui-label": "Validation maximum",
                                }),
                              ],
                            }),
                        ],
                      }),
                (m === "date" || m === "time") &&
                  i.jsx("p", {
                    children:
                      m === "date"
                        ? "Bounds use Gregorian/1900 date serials. Entered text follows the workbook's explicit MDY/DMY policy."
                        : "Bounds and numeric inputs use day fractions: 0.5 is noon. Time text such as 09:30 is supported.",
                  }),
                i.jsxs("label", {
                  className: "wt-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: j,
                      onChange: (ce) => z(ce.target.checked),
                    }),
                    i.jsx(LocalizedText, { children: "Allow blank cells" }),
                  ],
                }),
                m === "list" &&
                  i.jsxs("label", {
                    className: "wt-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: V,
                        onChange: (ce) => D(ce.target.checked),
                      }),
                      i.jsx(LocalizedText, { children: "Show dropdown" }),
                    ],
                  }),
                i.jsxs("label", {
                  className: "wt-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: B,
                      onChange: (ce) => _(ce.target.checked),
                    }),
                    i.jsx(LocalizedText, {
                      children: "Show input message when selected",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Input title" }),
                    i.jsx("input", {
                      "aria-label": text("Validation input title"),
                      maxLength: 100,
                      value: C,
                      onChange: (ce) => T(ce.target.value),
                      lang: language("Validation input title"),
                      "data-ui-label": "Validation input title",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Input message" }),
                    i.jsx("input", {
                      "aria-label": text("Validation input message"),
                      maxLength: 500,
                      value: N,
                      onChange: (ce) => R(ce.target.value),
                      lang: language("Validation input message"),
                      "data-ui-label": "Validation input message",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Error style" }),
                    i.jsxs("select", {
                      "aria-label": text("Validation error style"),
                      value: L,
                      onChange: (ce) => $(ce.target.value),
                      lang: language("Validation error style"),
                      "data-ui-label": "Validation error style",
                      children: [
                        i.jsx("option", { value: "stop", children: "Stop" }),
                        i.jsx("option", {
                          value: "warning",
                          children: "Warning",
                        }),
                        i.jsx("option", {
                          value: "information",
                          children: "Information",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Error title" }),
                    i.jsx("input", {
                      "aria-label": text("Validation error title"),
                      value: U,
                      onChange: (ce) => G(ce.target.value),
                      lang: language("Validation error title"),
                      "data-ui-label": "Validation error title",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Error message" }),
                    i.jsx("input", {
                      "aria-label": text("Validation message"),
                      maxLength: 500,
                      value: H,
                      onChange: (ce) => Q(ce.target.value),
                      lang: language("Validation message"),
                      "data-ui-label": "Validation message",
                    }),
                  ],
                }),
              ],
            }),
            X$1 && i.jsx("p", { role: "alert", children: X$1 }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: te,
              onClick: oe,
              children: c ? "Save rule" : "Apply",
            }),
          ],
        }),
      ],
    }),
  });
}
function shiftRangeForAxisEdit(e, r, t, n = false) {
  const o = parseRangeAddress(e),
    a = r.axis;
  return (
    resolveMovedRange(
      a === "row" ? o.firstRow : o.firstCol,
      a === "row" ? o.lastRow : o.lastCol,
      r,
      t,
      n,
    ),
    assertRangeWithinSheetBounds(e, r)
  );
}
function assertMoveKeepsTablesIntact(e, r, t) {
  const n = t,
    o = t.axis,
    a = structuredClone(e);
  for (const s of a.sheets) {
    for (const d of s.tables ?? []) {
      if (d.columns.some((h) => h.calculatedFormula))
        throw Error(
          "Clear saved calculated-column rules before moving worksheet structures, keeping their cell formulas; then reapply the rules.",
        );
      if (s.id !== r) continue;
      const u = parseRangeAddress(d.range),
        l = shiftRangeForAxisEdit(d.range, t, `table ${d.name}`),
        c = parseRangeAddress(l);
      if (o === "row") {
        if (
          mapIndexAfterMove(u.firstRow, n) !== c.firstRow ||
          (d.showTotals && mapIndexAfterMove(u.lastRow, n) !== c.lastRow)
        )
          throw Error(
            "This move would displace a table header or totals row. Move the complete table or only its data rows.",
          );
      } else
        d.columns = d.columns
          .map((h, f) => ({
            column: h,
            index: mapIndexAfterMove(u.firstCol + f, n),
          }))
          .sort((h, f) => h.index - f.index)
          .map((h) => h.column);
      d.range = l;
    }
    for (const d of [...(s.pivots ?? []), ...(s.queries ?? [])])
      if (
        (s.id === r &&
          (d.outputRange = shiftRangeForAxisEdit(
            d.outputRange,
            t,
            "generated analysis output",
            true,
          )),
        d.source.sheetId === r)
      ) {
        const u = parseRangeAddress(d.source.range),
          l = shiftRangeForAxisEdit(d.source.range, t, "analysis source");
        if (
          o === "row" &&
          mapIndexAfterMove(u.firstRow, n) !== parseRangeAddress(l).firstRow
        )
          throw Error(
            "This move would displace an analysis source header. Move the whole source or only its data rows.",
          );
        d.source.range = l;
      }
  }
  return a;
}
function collectAxisStructureConflicts(e, r) {
  const t = r.axis,
    n = r,
    o = t === "row" ? "hiddenRows" : "hiddenCols",
    a = t === "row" ? e.rowCount : e.colCount,
    s = (e[o] ?? []).flatMap((l) => mapRangeAfterMove(l.start, l.end, n)),
    d = e.outlineGroups?.map((l) =>
      l.axis !== t
        ? l
        : {
            ...l,
            ...resolveMovedRange(l.start, l.end, n, "outline group", true),
          },
    );
  return {
    [o]: normalizeHiddenIntervals(s, a),
    ...(d ? { outlineGroups: d } : {}),
  };
}
function shiftPrintLayoutForAxis(e, r) {
  if (!e.pageLayout) return;
  const t = { ...e.pageLayout },
    n = r.axis,
    o = r,
    a = n === "row" ? "repeatRows" : "repeatColumns",
    s = n === "row" ? "rowBreaks" : "columnBreaks",
    d = parsePrintTitles(t[a], n, e);
  if (d) {
    const u = resolveMovedRange(d.start, d.end, o, "print titles", true);
    t[a] =
      n === "row"
        ? `${u.start + 1}:${u.end + 1}`
        : `${columnIndexToLetters(u.start)}:${columnIndexToLetters(u.end)}`;
  }
  return (
    (t[s] &&= [...new Set(t[s].map((u) => mapIndexAfterMove(u, o)))].sort(
      (u, l) => u - l,
    )),
    t
  );
}
function remapAxisIndex(e, r) {
  return r.moveBefore !== undefined
    ? mapIndexAfterMove(e, r)
    : r.remove && e >= r.index && e < r.index + r.count
      ? null
      : e >= r.index + (r.remove ? r.count : 0)
        ? e + (r.remove ? -r.count : r.count)
        : e;
}
function assertInsertionWithinSheet(e, r) {
  const t = _t(e),
    n = remapAxisIndex(r.axis === "row" ? t.row : t.col, r);
  if (n === null) return null;
  if (
    n < 0 ||
    n >= (r.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols)
  )
    throw Error(
      "Insertion would move existing content beyond the Excel worksheet boundary.",
    );
  const o = /^(\$?)([A-Z]+)(\$?)(\d+)$/i.exec(e);
  return r.axis === "row"
    ? `${o[1]}${o[2]}${o[3]}${n + 1}`
    : `${o[1]}${columnIndexToLetters(n)}${o[3]}${o[4]}`;
}
function assertRangeWithinSheetBounds(e, r) {
  const t = parseRangeAddress(e);
  let n = r.axis === "row" ? t.firstRow : t.firstCol,
    o = r.axis === "row" ? t.lastRow : t.lastCol;
  if (r.moveBefore !== undefined) {
    const d = resolveMovedRange(n, o, r, `reference ${e}`);
    ((n = d.start), (o = d.end));
  } else if (r.remove) {
    const d = r.index + r.count - 1;
    if (n >= r.index && o <= d) return null;
    ((n = n >= r.index && n <= d ? r.index : remapAxisIndex(n, r)),
      (o = o >= r.index && o <= d ? r.index - 1 : remapAxisIndex(o, r)));
  } else ((n = remapAxisIndex(n, r)), (o = remapAxisIndex(o, r)));
  if (o >= (r.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols))
    throw Error("The range would exceed the Excel worksheet boundary.");
  const a =
      r.axis === "row"
        ? `${columnIndexToLetters(t.firstCol)}${n + 1}`
        : `${columnIndexToLetters(n)}${t.firstRow + 1}`,
    s =
      r.axis === "row"
        ? `${columnIndexToLetters(t.lastCol)}${o + 1}`
        : `${columnIndexToLetters(o)}${t.lastRow + 1}`;
  return a === s ? a : `${a}:${s}`;
}
function rewriteFormulaForStructuralEdit(e, r, t, n) {
  if (!e.startsWith("=")) return e;
  const o = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\\][A-Za-z0-9_.\\]*)`,
    a = new RegExp(
      `^(${o}!)?(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?::\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?[1-9]\\d*:\\$?[1-9]\\d*)`,
    ),
    s = new RegExp(`^${o}:${o}!`);
  let d = "",
    u = 0;
  for (; u < e.length;) {
    const l = e[u],
      c = e.slice(u);
    if (l === '"') {
      const p = u++;
      for (; u < e.length;)
        if (e[u++] === '"')
          if (e[u] === '"') u++;
          else break;
      d += e.slice(p, u);
      continue;
    }
    const h = u === 0 || !/[\w.$]/.test(e[u - 1]);
    if (h && s.test(c))
      throw Error(
        "Structural edits cannot safely update 3-D worksheet references. Replace the 3-D reference with explicit worksheet references first.",
      );
    if (l === "[") {
      const p = u++;
      let m = 1;
      for (; u < e.length && m;)
        (e[u] === "[" && m++, e[u] === "]" && m--, u++);
      if (h && /^[^+*/(),]*!/.test(e.slice(u)))
        throw Error(
          "Structural edits cannot safely update external workbook references.",
        );
      d += e.slice(p, u);
      continue;
    }
    const f = h ? a.exec(c) : null;
    if (
      f &&
      !/[\w.!:]/.test(e[u + f[0].length] ?? "") &&
      !/^\s*\(/.test(e.slice(u + f[0].length))
    ) {
      const p = f[1] ?? "",
        m = p ? p.slice(0, -1).replace(/^'|'$/g, "").replace(/''/g, "'") : r;
      if (m.includes(":"))
        throw Error(
          "Structural edits cannot safely update 3-D worksheet references.",
        );
      if (m.includes("["))
        throw Error(
          "Structural edits cannot safely update external workbook references.",
        );
      let v = f[0];
      if (m.toLowerCase() === t.toLowerCase()) {
        const g = f[2],
          b = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(g),
          y = /^\$?[1-9]\d*:\$?[1-9]\d*$/.test(g);
        try {
          if (b || y) {
            if ((b && n.axis === "column") || (y && n.axis === "row")) {
              const [w, E] = g.split(":"),
                A = w.replace(/\$/g, ""),
                k = E.replace(/\$/g, ""),
                F = assertRangeWithinSheetBounds(
                  b ? `${A}1:${k}1` : `A${A}:A${k}`,
                  n,
                );
              if (!F) v = p + "#REF!";
              else {
                const I = parseRangeAddress(F),
                  O = b
                    ? _t(A + "1").col > _t(k + "1").col
                    : Number(A) > Number(k),
                  j = b
                    ? columnIndexToLetters(I.firstCol)
                    : String(I.firstRow + 1),
                  z = b
                    ? columnIndexToLetters(I.lastCol)
                    : String(I.lastRow + 1);
                v =
                  p +
                  (w.startsWith("$") ? "$" : "") +
                  (O ? z : j) +
                  ":" +
                  (E.startsWith("$") ? "$" : "") +
                  (O ? j : z);
              }
            }
          } else if (g.includes(":")) {
            const [w, E] = g.split(":"),
              A = assertRangeWithinSheetBounds(g, n);
            if (!A) v = p + "#REF!";
            else {
              const k = parseRangeAddress(A),
                F = _t(w),
                I = _t(E),
                O = /^(\$?)[A-Z]+(\$?)\d+$/i.exec(w),
                j = /^(\$?)[A-Z]+(\$?)\d+$/i.exec(E),
                z = (V, D) => {
                  const B = F.col <= I.col === V ? k.firstCol : k.lastCol,
                    _ = F.row <= I.row === V ? k.firstRow : k.lastRow;
                  return `${D[1]}${columnIndexToLetters(B)}${D[2]}${_ + 1}`;
                };
              v = p + z(true, O) + ":" + z(false, j);
            }
          } else v = p + (assertInsertionWithinSheet(g, n) ?? "#REF!");
        } catch (w) {
          if (n.moveBefore !== undefined) throw w;
          v = p + "#REF!";
        }
      }
      ((d += v), (u += f[0].length));
      continue;
    }
    ((d += l), u++);
  }
  return d;
}
function applyStructuralEdit(e, r, t) {
  if (t.permutation !== undefined && t.moveBefore === undefined)
    throw Error(
      "A structural permutation requires an explicit move operation.",
    );
  if (
    !["row", "column"].includes(t.axis) ||
    !Number.isInteger(t.index) ||
    t.index < 0 ||
    !Number.isInteger(t.count) ||
    t.count < 1 ||
    t.count > 1e4 ||
    t.index + t.count >
      (t.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols)
  )
    throw Error("Choose a valid position and 1–10,000 rows or columns.");
  if (t.moveBefore !== undefined) {
    if (t.remove)
      throw Error(
        "A structural operation cannot delete and move at the same time.",
      );
    validateRowColumnMove(
      t,
      t.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols,
    );
  }
  const n = e.sheets.find((c) => c.id === r);
  if (!n) throw Error("The worksheet was not found.");
  assertSheetUnprotected(n);
  const o =
    t.moveBefore !== undefined
      ? assertMoveKeepsTablesIntact(e, r, t)
      : checkStructuralEditImpact(e, r, t.axis, t.index, t.count, t.remove);
  let a = 0;
  const s = o.sheets.map((c) => {
      const h = {};
      for (const [O, j] of Object.entries(c.cells)) {
        const z = c.id === r ? assertInsertionWithinSheet(O, t) : O;
        if (!z) continue;
        const V =
          typeof j.raw == "string" && j.raw.startsWith("=") && !j.literal
            ? rewriteFormulaForStructuralEdit(j.raw, c.name, n.name, t)
            : j.raw;
        (V !== j.raw && assertSheetUnprotected(c), (h[z] = { ...j, raw: V }));
      }
      const f = (O) => {
          if (!["formula", "custom"].includes(O.kind) || !O.formula) return O;
          const j = O.kind === "custom" ? "validation" : "conditional";
          if (
            [O.formula, ...(e.definedNames ?? []).map((T) => T.refersTo)].some(
              (T) =>
                /\b(?:INDIRECT|OFFSET)\s*\(/i.test(
                  T.replace(/"(?:[^"]|"")*"/g, '""'),
                ),
            )
          )
            throw Error(
              `Structural edits cannot repair dynamic ${j}-rule references. Remove the formula rule first.`,
            );
          if (c.id === r && t.remove) {
            const T = O.range.split(":")[0];
            if (assertInsertionWithinSheet(T, t) === null)
              throw Error(
                `This edit removes a ${j} formula's anchor. Relocate or remove the rule first.`,
              );
          }
          const V = rewriteFormulaForStructuralEdit(
              O.formula,
              c.name,
              n.name,
              t,
            ),
            D = t.axis === "row" ? 1 : 0,
            B = t.axis === "column" ? 1 : 0,
            _ = parseRangeAddress(O.range);
          if (
            shiftFormulaReferences(O.formula, D, B) !== O.formula &&
            (t.axis === "row" ? _.lastRow > _.firstRow : _.lastCol > _.firstCol)
          ) {
            const T =
              (_.lastRow - _.firstRow + 1) * (_.lastCol - _.firstCol + 1);
            if (a + T > 5e3)
              throw Error(
                `Structural edits can verify at most 5,000 applied cells for relative ${j} formulas. Narrow or remove the rules first.`,
              );
            a += T;
            const N =
              c.id === r ? assertRangeWithinSheetBounds(O.range, t) : O.range;
            if (!N)
              throw Error(
                `This edit removes a ${j} formula range. Remove the rule first.`,
              );
            const R = parseRangeAddress(N);
            for (let L = _.firstRow; L <= _.lastRow; L++)
              for (let $ = _.firstCol; $ <= _.lastCol; $++) {
                const U = `${columnIndexToLetters($)}${L + 1}`,
                  G = c.id === r ? assertInsertionWithinSheet(U, t) : U;
                if (!G) continue;
                const H = _t(G),
                  Q = rewriteFormulaForStructuralEdit(
                    shiftFormulaReferences(
                      O.formula,
                      L - _.firstRow,
                      $ - _.firstCol,
                    ),
                    c.name,
                    n.name,
                    t,
                  ),
                  X = shiftFormulaReferences(
                    V,
                    H.row - R.firstRow,
                    H.col - R.firstCol,
                  );
                if (Q !== X)
                  throw Error(
                    `This structural edit requires different relative ${j} formulas within one rule. Split, relocate, or remove the rule first.`,
                  );
              }
          }
          return (
            V !== O.formula && assertSheetUnprotected(c),
            { ...O, formula: V }
          );
        },
        p = c.conditionalFormats?.map(f),
        m = c.validations?.map(f);
      if (c.id !== r)
        return {
          ...c,
          cells: h,
          ...(p ? { conditionalFormats: p } : {}),
          ...(m ? { validations: m } : {}),
        };
      const v = t.axis === "row" ? "rowHeights" : "columnWidths",
        g = {};
      for (const [O, j] of Object.entries(c[v] ?? {})) {
        const z = remapAxisIndex(Number(O), t);
        if (z !== null) {
          if (
            z >=
            (t.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols)
          )
            throw Error(
              "Insertion would discard explicit sizing at the worksheet boundary.",
            );
          g[z] = j;
        }
      }
      const b = (c.merges ?? [])
          .map((O) =>
            t.moveBefore !== undefined
              ? shiftRangeForAxisEdit(O, t, "merged range", true)
              : assertRangeWithinSheetBounds(O, t),
          )
          .filter((O) => !!O),
        y = (m ?? []).flatMap((O) => {
          const j = assertRangeWithinSheetBounds(O.range, t);
          return j ? [{ ...O, range: j }] : [];
        }),
        w = (p ?? []).flatMap((O) => {
          const j = assertRangeWithinSheetBounds(O.range, t);
          return j ? [{ ...O, range: j }] : [];
        }),
        E = c.objects?.map((O) => {
          const j = assertInsertionWithinSheet(O.anchor, t);
          if (!j)
            throw Error(
              `This edit would remove the anchor of ${O.label || O.kind}. Move or delete that drawing first.`,
            );
          if (O.kind !== "chart" || !O.range) return { ...O, anchor: j };
          const z = parseRangeAddress(O.range),
            V = t.index + t.count - 1;
          if (
            t.remove &&
            t.axis === "row" &&
            O.headers &&
            t.index <= z.firstRow &&
            V >= z.firstRow
          )
            throw Error(
              "This edit would remove a chart header row. Adjust or delete the chart first.",
            );
          if (
            t.remove &&
            t.axis === "column" &&
            z.lastCol > z.firstCol &&
            t.index <= z.firstCol &&
            V >= z.firstCol
          )
            throw Error(
              "This edit would remove chart category labels. Adjust or delete the chart first.",
            );
          const D = assertRangeWithinSheetBounds(O.range, t);
          if (!D)
            throw Error(
              "This edit would remove the entire chart source. Adjust or delete the chart first.",
            );
          const B = measureChartSelection(D, c);
          if (
            t.moveBefore !== undefined &&
            ((t.axis === "row" &&
              O.headers &&
              remapAxisIndex(z.firstRow, t) !== B.firstRow) ||
              (t.axis === "column" &&
                z.lastCol > z.firstCol &&
                remapAxisIndex(z.firstCol, t) !== B.firstCol))
          )
            throw Error(
              "This move would displace chart headers or category labels. Move the complete chart source or only its data rows or value columns.",
            );
          if (
            (O.headers && B.lastRow === B.firstRow) ||
            (z.lastCol > z.firstCol && B.lastCol === B.firstCol)
          )
            throw Error(
              "This edit would remove all chart data. Adjust or delete the chart first.",
            );
          const _ = O.chartOptions
            ? structuredClone(O.chartOptions)
            : undefined;
          if (_) {
            const C = (T, N) => {
              const R = remapAxisIndex(T - 1, t);
              if (R === null)
                throw Error(
                  `This edit would remove a selected chart ${N}. Reset the chart selection before deleting it.`,
                );
              return R + 1;
            };
            (t.axis === "row" &&
              _.includedRows &&
              (_.includedRows = _.includedRows.map((T) => C(T, "filter row"))),
              t.axis === "column" &&
                _.series &&
                (_.series = _.series.map((T) => ({
                  ...T,
                  column: C(T.column, "series"),
                }))));
          }
          return {
            ...O,
            anchor: j,
            range: D,
            ...(_ ? { chartOptions: _ } : {}),
          };
        }),
        A =
          t.moveBefore !== undefined
            ? shiftPrintLayoutForAxis(c, t)
            : c.pageLayout
              ? shiftPrintSettingsOnAxisEdit(
                  c.pageLayout,
                  c,
                  t.axis,
                  t.index,
                  t.count,
                  t.remove,
                )
              : undefined;
      if (A) {
        const O = assertRangeWithinSheetBounds(A.printArea, t);
        if (!O)
          throw Error(
            "This edit would remove the entire print area. Change the print area first.",
          );
        A.printArea = O;
      }
      return {
        ...c,
        cells: h,
        ...(t.moveBefore !== undefined
          ? collectAxisStructureConflicts(c, t)
          : applyAxisStructureEdit(
              { ...c, merges: b },
              t.axis,
              t.index,
              t.count,
              t.remove,
            )),
        ...(t.axis === "row" && c.rowOriginIds
          ? {
              rowOriginIds: Object.fromEntries(
                Object.entries(c.rowOriginIds).flatMap(([O, j]) => {
                  const z = remapAxisIndex(Number(O), t);
                  if (z === null) return [];
                  if (z >= workbookLimits.maxRows)
                    throw Error(
                      "Insertion would discard source-row provenance at the worksheet boundary.",
                    );
                  return [[String(z), j]];
                }),
              ),
            }
          : {}),
        [v]: g,
        merges: b,
        validations: y,
        conditionalFormats: w,
        ...(E ? { objects: E } : {}),
        ...(A ? { pageLayout: A } : {}),
      };
    }),
    d = o.definedNames?.map((c) => {
      const h = c.scopeSheetId
          ? (o.sheets.find((m) => m.id === c.scopeSheetId)?.name ?? "")
          : "",
        f = rewriteFormulaForStructuralEdit(c.refersTo, h, n.name, t);
      if (
        !c.scopeSheetId &&
        rewriteFormulaForStructuralEdit(c.refersTo, n.name, n.name, t) !== f
      )
        throw Error(
          `Workbook-scoped name ${c.name} has ambiguous unqualified worksheet references. Qualify those references before inserting or deleting rows or columns.`,
        );
      return { ...c, refersTo: f };
    }),
    u = shiftReviewTargetsOnAxisEdit(
      shiftFormulaWatchesOnAxisEdit(
        { ...o, sheets: s, ...(d ? { definedNames: d } : {}) },
        r,
        t,
      ),
      r,
      t,
    );
  for (const c of u.sheets) {
    const h = e.sheets.find((f) => f.id === c.id);
    c.id !== r &&
      JSON.stringify(c.cells) !== JSON.stringify(h.cells) &&
      assertSheetUnprotected(h);
  }
  return normalizeWorkbookInput(u);
}
const structuralEditLimits = {
  sourceCells: 500,
  dependencyCells: 1e3,
  formulaCells: 5e3,
  depth: 8,
  pageSize: 20,
};

function resolveDuplicateReviewScope(e, r) {
  return (
    r.includeHidden !== undefined &&
      typeof r.includeHidden != "boolean" &&
      failWithMessage("Include hidden must be boolean."),
    r.allowedSheetIds !== undefined &&
      (!Array.isArray(r.allowedSheetIds) ||
        r.allowedSheetIds.some(
          (t) => typeof t != "string" || !e.sheets.some((n) => n.id === t),
        )) &&
      failWithMessage("The duplicate review host context is invalid."),
    {
      includeHidden: !!r.includeHidden,
      allowedSheetIds: [
        ...new Set(r.allowedSheetIds ?? e.sheets.map((t) => t.id)),
      ].sort(),
    }
  );
}
function Jde(e, r, t, n) {
  let o = "",
    a = 0;
  for (; a < r.length;) {
    if (r[a] === "'") {
      const u = a++;
      for (; a < r.length;)
        if (r[a++] === "'")
          if (r[a] === "'") a++;
          else break;
      o += r.slice(u, a);
      continue;
    }
    const s = a,
      d = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(r.slice(a));
    if (d && (a === 0 || !/[A-Za-z0-9_.$]/.test(r[a - 1]))) {
      a += d[0].length;
      let u = 1;
      for (; a < r.length && u;)
        (r[a] === "[" && u++, r[a] === "]" && u--, a++);
      const l = r.slice(s, a),
        c =
          !u &&
          !/^[A-Za-z0-9_!]/.test(r.slice(a)) &&
          rewriteStructuredReferences(e, t, l, n);
      o += c ? "0" : l;
      continue;
    }
    o += r[a++];
  }
  return o;
}
function Qde(e, r) {
  let t = 0,
    n = false;
  const o = (a, s, d) => {
    if (!a) return;
    ((t += a.length),
      t > 1e6 &&
        failWithMessage(
          "Output-name safety inspection exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ));
    const u = Jde(e, a.replace(/"(?:[^"]|"")*"/g, ""), s, d);
    for (const l of u.matchAll(
      /('(?:[^']|'')*'|[A-Za-z_.\d\u007F-\uFFFF]+)\s*!/g,
    )) {
      const c = l[1];
      (c.startsWith("'")
        ? c.slice(1, -1).replace(/''/g, "'")
        : c
      ).toLowerCase() === r.toLowerCase() && (n = true);
    }
  };
  for (const a of e.sheets) {
    for (const [s, d] of Object.entries(a.cells))
      isFormulaInputValue(d) && o(String(d.raw), a.id, s);
    for (const s of a.conditionalFormats ?? [])
      if (s.kind === "formula") {
        const d = parseRangeAddress(s.range);
        o(s.formula, a.id, formatCellAddress(d.firstRow, d.firstCol));
      }
    for (const s of a.tables ?? []) {
      const d = parseRangeAddress(s.range);
      s.columns.forEach((u, l) =>
        o(
          u.calculatedFormula,
          a.id,
          formatCellAddress(
            Math.min(d.firstRow + 1, d.lastRow),
            d.firstCol + l,
          ),
        ),
      );
    }
  }
  for (const a of e.definedNames ?? [])
    o(a.refersTo, a.scopeSheetId ?? e.sheets[0].id);
  return n;
}
function normalizeDuplicateReviewRequest(e, r) {
  ((!r ||
    typeof r != "object" ||
    Object.keys(r).some(
      (s) =>
        ![
          "sheetId",
          "range",
          "keyColumns",
          "headers",
          "comparison",
          "caseSensitive",
          "trimText",
          "blankPolicy",
          "outputSheetName",
        ].includes(s),
    )) &&
    failWithMessage("Invalid duplicate review request."),
    (typeof r.sheetId != "string" || typeof r.range != "string") &&
      failWithMessage("Choose a source worksheet and one rectangular range."));
  const t = parseRangeAddress(r.range),
    n = t.lastCol - t.firstCol + 1,
    o = t.lastRow - t.firstRow + 1;
  (n * o > structuralEditLimits.sourceCells &&
    failWithMessage("Select at most 500 source cells for duplicate review."),
    (!Array.isArray(r.keyColumns) ||
      !r.keyColumns.length ||
      r.keyColumns.some(
        (s) => !Number.isInteger(s) || s < t.firstCol || s > t.lastCol,
      )) &&
      failWithMessage(
        "Choose one or more key columns inside the source range.",
      ),
    ([r.headers, r.caseSensitive, r.trimText].some(
      (s) => typeof s != "boolean",
    ) ||
      !["raw", "calculated"].includes(r.comparison) ||
      !["match", "keep"].includes(r.blankPolicy)) &&
      failWithMessage(
        "Choose valid comparison, header, case, trim and blank rules.",
      ),
    o <= (r.headers ? 1 : 0) &&
      failWithMessage(
        "The source must contain at least one data row after its header.",
      ),
    typeof r.outputSheetName != "string" &&
      failWithMessage("Choose a new output worksheet name."));
  const a = r.outputSheetName.trim();
  return (
    (!a ||
      a.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(a) ||
      /^'|'$/.test(a)) &&
      failWithMessage(
        "The output sheet needs a valid Excel worksheet name of 1–31 characters.",
      ),
    e.sheets.some((s) => s.name.toLowerCase() === a.toLowerCase()) &&
      failWithMessage("A worksheet with that output name already exists."),
    Qde(e, a) &&
      failWithMessage(
        "That output name is already referenced by workbook formulas or formula metadata. Choose a different name to avoid reconnecting an unresolved source reference.",
      ),
    {
      ...r,
      range: `${formatCellAddress(t.firstRow, t.firstCol)}:${formatCellAddress(t.lastRow, t.lastCol)}`,
      keyColumns: [...new Set(r.keyColumns)].sort((s, d) => s - d),
      outputSheetName: a,
    }
  );
}
function Pz(e, r) {
  const t = new Set(e.sheets.map((a) => a.id)),
    n = (a) => {
      if (!(!a || typeof a != "object"))
        for (const [s, d] of Object.entries(a))
          /sheetId$/i.test(s) && typeof d == "string"
            ? t.add(d)
            : typeof d == "object" && n(d);
    };
  n(e);
  let o = 1;
  for (; t.has(`${r}-${o}`);) o++;
  return `${r}-${o}`;
}
function runDuplicateInspection(e, r, t, n) {
  const o = parseRangeAddress(t),
    a = [];
  for (let v = o.firstRow; v <= o.lastRow; v++)
    for (let g = o.firstCol; g <= o.lastCol; g++)
      a.push(formatCellAddress(v, g));
  if (
    (a.some((v) => !r.cells[v] || r.cells[v].raw === null) &&
      failWithMessage(
        "Calculated comparison requires populated source cells; blank or possible spill-child cells are outside this subset. Use raw comparison or a populated selection.",
      ),
    !a.some((v) => isFormulaInputValue(r.cells[v])))
  )
    return new Map(a.map((v) => [v, r.cells[v].raw]));
  const s = {
    id: Pz(e, "duplicate-inspection"),
    name: "Duplicate inspection",
    rowCount: workbookLimits.maxRows,
    colCount: workbookLimits.maxCols,
    cells: { A1: { raw: `=SUM('${r.name.replace(/'/g, "''")}'!${t})` } },
  };
  let d = false;
  for (let v = 0; v < 4; v++) {
    s.name = `_dup_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`;
    const g = s.name.toLowerCase();
    if (
      !e.sheets.some(
        (b) =>
          b.name.toLowerCase() === g ||
          Object.values(b.cells).some(
            (y) => typeof y.raw == "string" && y.raw.toLowerCase().includes(g),
          ),
      ) &&
      !e.definedNames?.some((b) => b.refersTo.toLowerCase().includes(g))
    ) {
      d = true;
      break;
    }
  }
  d ||
    failWithMessage(
      "A temporary duplicate inspection context could not be isolated.",
    );
  const u = inspectFormulaDependencies(
      {
        ...e,
        sheets: [
          s,
          ...e.sheets.map((v) => (v.veryHidden ? { ...v, hidden: true } : v)),
        ],
      },
      {
        sheetId: s.id,
        address: "A1",
        direction: "precedents",
        depth: structuralEditLimits.depth,
        includeHidden: n.includeHidden,
        allowedSheetIds: [...n.allowedSheetIds, s.id],
        limits: {
          maxNodes: structuralEditLimits.dependencyCells,
          maxFormulaCells: structuralEditLimits.formulaCells,
          maxReferences: 1e4,
        },
      },
    ),
    l = new Set(u.nodes.map((v) => v.key));
  (u.coverage.truncated ||
    u.coverage.depthLimited ||
    u.cycles.length ||
    u.issues.some((v) => v.from === u.target.key || l.has(v.from))) &&
    failWithMessage(
      "Calculated comparison requires a complete supported static dependency closure within 1,000 cells, 5,000 scanned formulas and depth 8; hidden, dynamic, unsupported or cyclic dependencies are excluded.",
    );
  const c = new Set([
    ...matrixResultFunctionNames,
    "ISFORMULA",
    "INDEX",
    "CHOOSE",
    "INDIRECT",
    "OFFSET",
    "LET",
    "LAMBDA",
    "RAND",
    "RANDBETWEEN",
    "NOW",
    "TODAY",
    "CELL",
    "INFO",
    "RTD",
    "WEBSERVICE",
    "STOCKHISTORY",
    "EVALUATE",
    "GET.CELL",
  ]);
  for (const v of u.nodes) {
    const g = e.sheets.find((y) => y.id === v.sheetId),
      b = g.cells[v.address];
    if (
      ((!b || b.raw === null) &&
        failWithMessage(
          "Calculated comparison requires populated static dependencies; blank or possible spill-child dependencies are unsupported.",
        ),
      isFormulaInputValue(b))
    ) {
      const y = resolveDefinedNameValue(String(b.raw), e, g.id).replace(
        /"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g,
        "",
      );
      (/[{}#]/.test(y) ||
        [...y.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)].some((w) =>
          c.has(
            w[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, ""),
          ),
        )) &&
        failWithMessage(
          "Calculated comparison excludes volatile, dynamic-reference and array/spill formulas. Raw comparison can preserve their formula text without execution.",
        );
    }
  }
  const h = {
      ...e,
      sheets: e.sheets
        .filter((v) => n.allowedSheetIds.includes(v.id))
        .map((v) => ({
          ...v,
          cells: Object.fromEntries(
            Object.entries(v.cells).filter(([g]) =>
              l.has(buildCompositeKey(v.id, g)),
            ),
          ),
        })),
    },
    f = calculateWorkbookWithSettings(h),
    p = new Map();
  for (const v of u.nodes) {
    const g = f(v.sheetId, v.address);
    (f.getOperationalErrors().length &&
      failWithMessage(
        `An operational calculation error (${f.getOperationalErrors().join(", ")}) prevents duplicate review, even when a formula handler conceals it.`,
      ),
      f.getSpillRange(v.sheetId, v.address) &&
        failWithMessage(
          "Calculated comparison requires scalar cells; a dependency produced an array/spill result.",
        ),
      isFormulaInputValue(
        e.sheets.find((b) => b.id === v.sheetId).cells[v.address],
      ) &&
        typeof g == "string" &&
        /^#(?:[A-Z0-9_/.]+[!?]?)/.test(g) &&
        failWithMessage(
          "Calculated comparison cannot group formula errors or error-like formula strings; correct them or use raw comparison.",
        ),
      typeof g == "number" &&
        !Number.isFinite(g) &&
        failWithMessage("Nonfinite values cannot be used in duplicate review."),
      v.sheetId === r.id && a.includes(v.address) && p.set(v.address, g));
  }
  return (
    (f.getOperationalErrors().length || p.size !== a.length) &&
      failWithMessage(
        "Calculated source coverage is incomplete or encountered an operational error.",
      ),
    p
  );
}
function createDuplicateReviewSheet(e, r, t = {}) {
  (t.readOnly &&
    failWithMessage(
      "This workbook is read-only. A derived worksheet cannot be created.",
    ),
    assertWorkbookStructureEditable(e));
  const n = normalizeDuplicateReviewRequest(e, r),
    o = resolveDuplicateReviewScope(e, t),
    a = e.sheets.find((j) => j.id === n.sheetId);
  a || failWithMessage("The source worksheet no longer exists.");
  const s = parseRangeAddress(n.range),
    d = s.lastCol - s.firstCol + 1;
  ((!o.allowedSheetIds.includes(a.id) ||
    (!o.includeHidden &&
      (a.hidden ||
        a.veryHidden ||
        cn(a, "row").some((j) => j.start <= s.lastRow && j.end >= s.firstRow) ||
        cn(a, "column").some(
          (j) => j.start <= s.lastCol && j.end >= s.firstCol,
        )))) &&
    failWithMessage(
      "The source range is outside the visible, authorized host context. Include hidden only within permitted host scope.",
    ),
    (a.merges ?? []).some((j) => {
      const z = parseRangeAddress(j);
      return (
        z.firstRow <= s.lastRow &&
        z.lastRow >= s.firstRow &&
        z.firstCol <= s.lastCol &&
        z.lastCol >= s.firstCol
      );
    }) &&
      failWithMessage(
        "Unmerge the selected source before duplicate review so row boundaries are unambiguous.",
      ));
  const u =
      n.comparison === "calculated"
        ? runDuplicateInspection(e, a, n.range, o)
        : null,
    l = (j, z) => {
      const V = formatCellAddress(j, z),
        D = structuredClone(a.cells[V] ?? { raw: null }),
        B = u ? u.get(V) : D.raw;
      return (
        typeof B == "number" &&
          !Number.isFinite(B) &&
          failWithMessage(
            "Nonfinite values cannot be used in duplicate review.",
          ),
        { cell: D, value: B }
      );
    },
    c = Array.from({ length: d }, (j, z) =>
      n.headers
        ? l(s.firstRow, s.firstCol + z).cell
        : { raw: `Column ${columnIndexToLetters(s.firstCol + z)}` },
    ),
    h = [],
    f = new Map();
  for (let j = s.firstRow + (n.headers ? 1 : 0); j <= s.lastRow; j++) {
    const z = Array.from({ length: d }, (L, $) => l(j, s.firstCol + $)),
      V = z.map((L) => L.cell),
      D = z.map((L) => L.value),
      B = n.keyColumns.map((L) => {
        const $ = L - s.firstCol,
          U = V[$];
        let G = D[$];
        return n.comparison === "raw" && isFormulaInputValue(U)
          ? ["formula", G]
          : (typeof G == "string" &&
              (n.trimText && (G = G.trim()),
              n.caseSensitive || (G = G.toLowerCase())),
            G === null || G === "" ? ["blank"] : [typeof G, G]);
      }),
      _ = B.some((L) => L[0] === "blank"),
      C = JSON.stringify(B),
      T = {
        sourceRow: j + 1,
        sourceRange: `${formatCellAddress(j, s.firstCol)}:${formatCellAddress(j, s.lastCol)}`,
        cells: V,
        values: D,
        normalizedKey: C,
        blankKey: _,
        disposition: "retained",
        groupId: "",
        firstSourceRow: j + 1,
        groupSize: 1,
        outputRow: 0,
      },
      N = _ && n.blankPolicy === "keep" ? `blank-row:${j}` : C,
      R = f.get(N) ?? [];
    (R.push(T), f.set(N, R), h.push(T));
  }
  const p = [];
  for (const j of f.values())
    if (j.length > 1) {
      const z = `G${p.length + 1}`,
        V = j[0].sourceRow;
      p.push({
        id: z,
        firstSourceRow: V,
        memberRows: j.map((D) => D.sourceRow),
        normalizedKey: j[0].normalizedKey,
      });
      for (let D = 0; D < j.length; D++)
        Object.assign(j[D], {
          groupId: z,
          firstSourceRow: V,
          groupSize: j.length,
          disposition: D === 0 ? "retained" : "excluded",
        });
    }
  const m = h.filter((j) => j.disposition === "retained"),
    v = h.filter((j) => j.disposition === "excluded"),
    g = {
      input: h.length,
      retained: m.length,
      excluded: v.length,
      duplicateGroups: p.length,
      blankKeyRows: h.filter((j) => j.blankKey).length,
    },
    b = {
      id: Pz(e, "duplicate-review"),
      name: n.outputSheetName,
      rowCount: workbookLimits.maxRows,
      colCount: workbookLimits.maxCols,
      cells: {},
      frozenRows: 3,
    },
    y = (j) => ({
      raw: j,
      ...(typeof j == "string" && j.startsWith("=") ? { literal: true } : {}),
    }),
    w = (j, z, V) => {
      b.cells[formatCellAddress(j - 1, z)] = V;
    };
  (w(1, 0, y("Unique population — static source snapshot")),
    w(
      2,
      0,
      y(
        `${g.input} source rows; ${g.retained} retained; ${g.excluded} excluded from population. ${n.comparison} typed keys ${n.keyColumns.map(columnIndexToLetters).join(", ")}; ${n.headers ? "source header" : "no source header"}; ${n.caseSensitive ? "case-sensitive" : "ignore case"}; ${n.trimText ? "trim outer text" : "preserve whitespace"}; blanks ${n.blankPolicy}. First source row wins. Source unchanged.`,
      ),
    ));
  const E = (j) => {
      (c.forEach((z, V) => {
        const D =
          n.headers && u
            ? u.get(formatCellAddress(s.firstRow, s.firstCol + V))
            : z.raw;
        w(j, V, {
          ...structuredClone(z),
          ...y(D),
          literal: typeof D == "string" && D.startsWith("=") ? true : z.literal,
        });
      }),
        [
          "Source locator",
          "Duplicate group",
          "First source row",
          "Group members",
        ].forEach((z, V) => w(j, d + V, y(z))));
    },
    A = (j, z) => {
      ((j.outputRow = z),
        j.cells.forEach((D, B) => {
          const _ = j.values[B];
          w(z, B, {
            ...structuredClone(D),
            raw: _,
            ...(typeof _ == "string" && _.startsWith("=")
              ? { literal: true }
              : {}),
          });
        }),
        [
          JSON.stringify({
            workbookId: e.id,
            revision: e.revision,
            sheetId: a.id,
            sheetName: a.name,
            range: j.sourceRange,
          }),
          j.groupId,
          j.firstSourceRow,
          j.groupSize,
        ].forEach((D, B) => w(z, d + B, y(D))));
    };
  (E(3), m.forEach((j, z) => A(j, 4 + z)));
  const k = `A3:${columnIndexToLetters(d + 3)}${3 + m.length}`;
  let F = null;
  if (v.length) {
    const j = m.length + 6;
    (w(
      j,
      0,
      y("Excluded duplicate rows — retained here and unchanged on the source"),
    ),
      E(j + 1),
      v.forEach((z, V) => A(z, j + 2 + V)),
      (F = `A${j + 1}:${columnIndexToLetters(d + 3)}${j + 1 + v.length}`));
  }
  e.sheets.reduce((j, z) => j + Object.keys(z.cells).length, 0) +
    Object.keys(b.cells).length >
    workbookLimits.maxCells &&
    failWithMessage(
      "The derived worksheet would exceed the workbook's stored-cell budget.",
    );
  const O = structuredClone(e);
  return (
    O.sheets.push(b),
    new TextEncoder().encode(JSON.stringify(O)).byteLength >
      workbookLimits.maxNativeBytes &&
      failWithMessage(
        "The derived workbook would exceed the native file budget.",
      ),
    {
      snapshot: JSON.stringify(e),
      request: n,
      context: o,
      headerCells: c,
      rows: h,
      groups: p,
      counts: g,
      outputSheetId: b.id,
      populationRange: k,
      excludedRange: F,
      next: O,
    }
  );
}
function applyDuplicateReviewPreview(e, r, t = {}) {
  (t.readOnly && failWithMessage("This workbook is read-only."),
    JSON.stringify(e) !== r.snapshot &&
      failWithMessage(
        "The workbook changed after duplicate review. Preview again before Apply.",
      ));
  const n = createDuplicateReviewSheet(e, r.request, t);
  return (
    JSON.stringify(n) !== JSON.stringify(r) &&
      failWithMessage(
        "The duplicate preview or host context changed. Preview and review again.",
      ),
    n.next
  );
}
const describeCellValueType = (e) =>
    e === null
      ? "Blank"
      : typeof e == "string"
        ? JSON.stringify(e)
        : `${typeof e}: ${String(e)}`,
  qk = (e) => {
    try {
      const r = parseRangeAddress(e);
      return r.lastCol - r.firstCol + 1 <= 500
        ? Array.from(
            { length: r.lastCol - r.firstCol + 1 },
            (t, n) => r.firstCol + n,
          )
        : [];
    } catch {
      return [];
    }
  };

export {
  applyStructuralEdit,
  createDuplicateReviewSheet,
  describeCellValueType,
  assertInsertionWithinSheet,
  structuralEditLimits,
  DataValidationDialog,
  qk,
  applyDuplicateReviewPreview,
};
