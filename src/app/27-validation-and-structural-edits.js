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
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    [c, h] = ne.useState(),
    [f, p] = ne.useState(range),
    [m, v] = ne.useState("list"),
    [operator, b] = ne.useState("between"),
    [y, w] = ne.useState("Pending,Complete,Reviewed"),
    [E, onValueChange] = ne.useState("0"),
    [k, F] = ne.useState("100"),
    [formula, O] = ne.useState(`=LEN(${range.split(":")[0]})>0`),
    [j, z] = ne.useState(true),
    [V, D] = ne.useState(true),
    [B, _] = ne.useState(true),
    [inputTitle, T] = ne.useState(""),
    [inputMessage, R] = ne.useState(""),
    [errorStyle, $] = ne.useState("stop"),
    [errorTitle, G] = ne.useState("Invalid value"),
    [H, Q] = ne.useState("Enter a value that satisfies this rule."),
    [X$1, J] = ne.useState(""),
    [q, W] = ne.useState(),
    K = ne.useRef(null);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      K.current?.querySelector("button,input")?.focus(),
      () => activeElement?.focus()
    );
  }, []);
  const disabled = readOnly || !!sheet.protection?.enabled,
    ae = (ce) => {
      J("");
      try {
        ce();
      } catch (error) {
        J(
          error instanceof Error
            ? error.message
            : "Validation could not be completed.",
        );
      }
    },
    onClick = () =>
      ae(() => {
        if (disabled) throw Error("This worksheet is read-only or protected.");
        if (
          (assertSheetUnprotected(sheet),
          !["list", "custom"].includes(m) && (!E.trim() || !k.trim()))
        )
          throw Error("Enter finite validation bounds.");
        const ce = {
            id: c ?? crypto.randomUUID(),
            range: f,
            kind: m,
            operator: operator,
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
            errorStyle: errorStyle,
            errorTitle: errorTitle,
            showDropdown: V,
            showInputMessage: B,
            inputTitle: inputTitle,
            inputMessage: inputMessage,
            ...(m === "custom" ? { formula: formula } : {}),
          },
          dataValidationRules = normalizeDataValidationRules(
            c
              ? (sheet.validations ?? []).map((se) => (se.id === c ? ce : se))
              : [...(sheet.validations ?? []), ce],
          );
        (onWorkbookChange(
          {
            ...workbook,
            sheets: workbook.sheets.map((sheet) =>
              sheet.id === sheetId
                ? { ...sheet, validations: dataValidationRules }
                : sheet,
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
        onValueChange(String(ce.min)),
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
    onKeyDown: (event) => {
      if (
        (event.key === "Escape" && (event.stopPropagation(), onClose()),
        event.key === "Tab")
      ) {
        const se = [
            ...(K.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled)",
            ) ?? []),
          ],
          we = se[0],
          Ee = se.at(-1);
        event.shiftKey && document.activeElement === we
          ? (event.preventDefault(), Ee == null || Ee.focus())
          : !event.shiftKey &&
            document.activeElement === Ee &&
            (event.preventDefault(), we == null || we.focus());
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
                        const validationInspection =
                            resolveValidationInspection(workbook, q.value),
                          ge = URL.createObjectURL(
                            new Blob([validationInspection], {
                              type: "application/json",
                            }),
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
                      children: q.value.issues.map((issue, ge) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsxs("button", {
                                onClick: () => {
                                  (onNavigate?.(issue.sheetId, issue.address),
                                    onNavigate && onClose());
                                },
                                disabled: !onNavigate,
                                children: [issue.sheetName, "!", issue.address],
                              }),
                              " ",
                              issue.style,
                              ": ",
                              issue.message,
                              " ",
                              issue.reason,
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
                    sheet.validations?.length ?? 0,
                    " ",
                    i.jsx(LocalizedText, { children: "existing rules" }),
                  ],
                }),
                sheet.validations?.map((validation) =>
                  i.jsxs(
                    "div",
                    {
                      className: "wt-existing",
                      children: [
                        i.jsxs("span", {
                          children: [
                            validation.range,
                            " · ",
                            validation.kind,
                            " · ",
                            validation.errorStyle ?? "stop",
                          ],
                        }),
                        i.jsxs("button", {
                          disabled: disabled,
                          onClick: () => de(validation),
                          children: [
                            i.jsx(LocalizedText, { children: "Edit " }),
                            validation.range,
                          ],
                        }),
                        i.jsxs("button", {
                          disabled: disabled,
                          onClick: () =>
                            ae(() => {
                              (assertSheetUnprotected(sheet),
                                onWorkbookChange(
                                  {
                                    ...workbook,
                                    sheets: workbook.sheets.map((sheet) =>
                                      sheet.id === sheetId
                                        ? {
                                            ...sheet,
                                            validations:
                                              sheet.validations?.filter(
                                                (we) => we.id !== validation.id,
                                              ),
                                          }
                                        : sheet,
                                    ),
                                  },
                                  "Remove data validation",
                                ),
                                onClose());
                            }),
                          children: [
                            i.jsx(LocalizedText, { children: "Remove " }),
                            validation.range,
                          ],
                        }),
                      ],
                    },
                    validation.id,
                  ),
                ),
              ],
            }),
            i.jsxs("fieldset", {
              disabled: disabled,
              style: { border: 0, padding: 0, display: "grid", gap: 12 },
              children: [
                i.jsx("legend", { children: c ? "Edit rule" : "New rule" }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Apply to" }),
                    i.jsx("input", {
                      "aria-label": text("Validation range"),
                      value: f,
                      onChange: (event) => p(event.target.value),
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
                          onChange: (event) => v(event.target.value),
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
                            value: operator,
                            onChange: (event) => b(event.target.value),
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
                          onChange: (event) => w(event.target.value),
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
                            value: formula,
                            onChange: (event) => O(event.target.value),
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
                                onValueChange: onValueChange,
                                lang: language("Validation values"),
                                "data-ui-label": "Validation values",
                              }),
                            ],
                          }),
                          ["between", "notBetween"].includes(operator) &&
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
                      onChange: (event) => z(event.target.checked),
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
                        onChange: (event) => D(event.target.checked),
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
                      onChange: (event) => _(event.target.checked),
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
                      value: inputTitle,
                      onChange: (event) => T(event.target.value),
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
                      value: inputMessage,
                      onChange: (event) => R(event.target.value),
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
                      value: errorStyle,
                      onChange: (event) => $(event.target.value),
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
                      value: errorTitle,
                      onChange: (event) => G(event.target.value),
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
                      onChange: (event) => Q(event.target.value),
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
              disabled: disabled,
              onClick: onClick,
              children: c ? "Save rule" : "Apply",
            }),
          ],
        }),
      ],
    }),
  });
}
function shiftRangeForAxisEdit(e, r, t, n = false) {
  const rangeAddress = parseRangeAddress(e),
    axis = r.axis;
  return (
    resolveMovedRange(
      axis === "row" ? rangeAddress.firstRow : rangeAddress.firstCol,
      axis === "row" ? rangeAddress.lastRow : rangeAddress.lastCol,
      r,
      t,
      n,
    ),
    assertRangeWithinSheetBounds(e, r)
  );
}
function assertMoveKeepsTablesIntact(e, sheetId, t) {
  const n = t,
    axis = t.axis,
    workbook = structuredClone(e);
  for (const sheet of workbook.sheets) {
    for (const d of sheet.tables ?? []) {
      if (d.columns.some((column) => column.calculatedFormula))
        throw Error(
          "Clear saved calculated-column rules before moving worksheet structures, keeping their cell formulas; then reapply the rules.",
        );
      if (sheet.id !== sheetId) continue;
      const rangeAddress = parseRangeAddress(d.range),
        l = shiftRangeForAxisEdit(d.range, t, `table ${d.name}`),
        c = parseRangeAddress(l);
      if (axis === "row") {
        if (
          mapIndexAfterMove(rangeAddress.firstRow, n) !== c.firstRow ||
          (d.showTotals &&
            mapIndexAfterMove(rangeAddress.lastRow, n) !== c.lastRow)
        )
          throw Error(
            "This move would displace a table header or totals row. Move the complete table or only its data rows.",
          );
      } else
        d.columns = d.columns
          .map((column, f) => ({
            column: column,
            index: mapIndexAfterMove(rangeAddress.firstCol + f, n),
          }))
          .sort((h, f) => h.index - f.index)
          .map((h) => h.column);
      d.range = l;
    }
    for (const d of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
      if (
        (sheet.id === sheetId &&
          (d.outputRange = shiftRangeForAxisEdit(
            d.outputRange,
            t,
            "generated analysis output",
            true,
          )),
        d.source.sheetId === sheetId)
      ) {
        const rangeAddress = parseRangeAddress(d.source.range),
          l = shiftRangeForAxisEdit(d.source.range, t, "analysis source");
        if (
          axis === "row" &&
          mapIndexAfterMove(rangeAddress.firstRow, n) !==
            parseRangeAddress(l).firstRow
        )
          throw Error(
            "This move would displace an analysis source header. Move the whole source or only its data rows.",
          );
        d.source.range = l;
      }
  }
  return workbook;
}
function collectAxisStructureConflicts(sheet, r) {
  const axis = r.axis,
    n = r,
    o = axis === "row" ? "hiddenRows" : "hiddenCols",
    a = axis === "row" ? sheet.rowCount : sheet.colCount,
    s = (sheet[o] ?? []).flatMap((l) => mapRangeAfterMove(l.start, l.end, n)),
    outlineGroups = sheet.outlineGroups?.map((outlineGroup) =>
      outlineGroup.axis !== axis
        ? outlineGroup
        : {
            ...outlineGroup,
            ...resolveMovedRange(
              outlineGroup.start,
              outlineGroup.end,
              n,
              "outline group",
              true,
            ),
          },
    );
  return {
    [o]: normalizeHiddenIntervals(s, a),
    ...(outlineGroups ? { outlineGroups: outlineGroups } : {}),
  };
}
function shiftPrintLayoutForAxis(sheet, r) {
  if (!sheet.pageLayout) return;
  const t = { ...sheet.pageLayout },
    axis = r.axis,
    o = r,
    a = axis === "row" ? "repeatRows" : "repeatColumns",
    s = axis === "row" ? "rowBreaks" : "columnBreaks",
    printTitles = parsePrintTitles(t[a], axis, sheet);
  if (printTitles) {
    const movedRange = resolveMovedRange(
      printTitles.start,
      printTitles.end,
      o,
      "print titles",
      true,
    );
    t[a] =
      axis === "row"
        ? `${movedRange.start + 1}:${movedRange.end + 1}`
        : `${columnIndexToLetters(movedRange.start)}:${columnIndexToLetters(movedRange.end)}`;
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
  const rangeAddress = parseRangeAddress(e);
  let n = r.axis === "row" ? rangeAddress.firstRow : rangeAddress.firstCol,
    o = r.axis === "row" ? rangeAddress.lastRow : rangeAddress.lastCol;
  if (r.moveBefore !== undefined) {
    const movedRange = resolveMovedRange(n, o, r, `reference ${e}`);
    ((n = movedRange.start), (o = movedRange.end));
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
        ? `${columnIndexToLetters(rangeAddress.firstCol)}${n + 1}`
        : `${columnIndexToLetters(n)}${rangeAddress.firstRow + 1}`,
    s =
      r.axis === "row"
        ? `${columnIndexToLetters(rangeAddress.lastCol)}${o + 1}`
        : `${columnIndexToLetters(o)}${rangeAddress.lastRow + 1}`;
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
                const rangeAddress = parseRangeAddress(F),
                  O = b
                    ? _t(A + "1").col > _t(k + "1").col
                    : Number(A) > Number(k),
                  j = b
                    ? columnIndexToLetters(rangeAddress.firstCol)
                    : String(rangeAddress.firstRow + 1),
                  z = b
                    ? columnIndexToLetters(rangeAddress.lastCol)
                    : String(rangeAddress.lastRow + 1);
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
              const rangeAddress = parseRangeAddress(A),
                F = _t(w),
                I = _t(E),
                O = /^(\$?)[A-Z]+(\$?)\d+$/i.exec(w),
                j = /^(\$?)[A-Z]+(\$?)\d+$/i.exec(E),
                z = (V, D) => {
                  const B =
                      F.col <= I.col === V
                        ? rangeAddress.firstCol
                        : rangeAddress.lastCol,
                    _ =
                      F.row <= I.row === V
                        ? rangeAddress.firstRow
                        : rangeAddress.lastRow;
                  return `${D[1]}${columnIndexToLetters(B)}${D[2]}${_ + 1}`;
                };
              v = p + z(true, O) + ":" + z(false, j);
            }
          } else v = p + (assertInsertionWithinSheet(g, n) ?? "#REF!");
        } catch (error) {
          if (n.moveBefore !== undefined) throw error;
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
function applyStructuralEdit(workbook, sheetId, t) {
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
  const n = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!n) throw Error("The worksheet was not found.");
  assertSheetUnprotected(n);
  const o =
    t.moveBefore !== undefined
      ? assertMoveKeepsTablesIntact(workbook, sheetId, t)
      : checkStructuralEditImpact(
          workbook,
          sheetId,
          t.axis,
          t.index,
          t.count,
          t.remove,
        );
  let a = 0;
  const sheets = o.sheets.map((sheet) => {
      const cells = {};
      for (const [O, cell] of Object.entries(sheet.cells)) {
        const z = sheet.id === sheetId ? assertInsertionWithinSheet(O, t) : O;
        if (!z) continue;
        const raw =
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          !cell.literal
            ? rewriteFormulaForStructuralEdit(cell.raw, sheet.name, n.name, t)
            : cell.raw;
        (raw !== cell.raw && assertSheetUnprotected(sheet),
          (cells[z] = { ...cell, raw: raw }));
      }
      const f = (O) => {
          if (!["formula", "custom"].includes(O.kind) || !O.formula) return O;
          const j = O.kind === "custom" ? "validation" : "conditional";
          if (
            [
              O.formula,
              ...(workbook.definedNames ?? []).map((T) => T.refersTo),
            ].some((T) =>
              /\b(?:INDIRECT|OFFSET)\s*\(/i.test(
                T.replace(/"(?:[^"]|"")*"/g, '""'),
              ),
            )
          )
            throw Error(
              `Structural edits cannot repair dynamic ${j}-rule references. Remove the formula rule first.`,
            );
          if (sheet.id === sheetId && t.remove) {
            const T = O.range.split(":")[0];
            if (assertInsertionWithinSheet(T, t) === null)
              throw Error(
                `This edit removes a ${j} formula's anchor. Relocate or remove the rule first.`,
              );
          }
          const formula = rewriteFormulaForStructuralEdit(
              O.formula,
              sheet.name,
              n.name,
              t,
            ),
            D = t.axis === "row" ? 1 : 0,
            B = t.axis === "column" ? 1 : 0,
            rangeAddress = parseRangeAddress(O.range);
          if (
            shiftFormulaReferences(O.formula, D, B) !== O.formula &&
            (t.axis === "row"
              ? rangeAddress.lastRow > rangeAddress.firstRow
              : rangeAddress.lastCol > rangeAddress.firstCol)
          ) {
            const T =
              (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
              (rangeAddress.lastCol - rangeAddress.firstCol + 1);
            if (a + T > 5e3)
              throw Error(
                `Structural edits can verify at most 5,000 applied cells for relative ${j} formulas. Narrow or remove the rules first.`,
              );
            a += T;
            const N =
              sheet.id === sheetId
                ? assertRangeWithinSheetBounds(O.range, t)
                : O.range;
            if (!N)
              throw Error(
                `This edit removes a ${j} formula range. Remove the rule first.`,
              );
            const R = parseRangeAddress(N);
            for (
              let firstRow = rangeAddress.firstRow;
              firstRow <= rangeAddress.lastRow;
              firstRow++
            )
              for (
                let firstCol = rangeAddress.firstCol;
                firstCol <= rangeAddress.lastCol;
                firstCol++
              ) {
                const U = `${columnIndexToLetters(firstCol)}${firstRow + 1}`,
                  G =
                    sheet.id === sheetId ? assertInsertionWithinSheet(U, t) : U;
                if (!G) continue;
                const H = _t(G),
                  Q = rewriteFormulaForStructuralEdit(
                    shiftFormulaReferences(
                      O.formula,
                      firstRow - rangeAddress.firstRow,
                      firstCol - rangeAddress.firstCol,
                    ),
                    sheet.name,
                    n.name,
                    t,
                  ),
                  X = shiftFormulaReferences(
                    formula,
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
            formula !== O.formula && assertSheetUnprotected(sheet),
            { ...O, formula: formula }
          );
        },
        conditionalFormats = sheet.conditionalFormats?.map(f),
        validations = sheet.validations?.map(f);
      if (sheet.id !== sheetId)
        return {
          ...sheet,
          cells: cells,
          ...(conditionalFormats
            ? { conditionalFormats: conditionalFormats }
            : {}),
          ...(validations ? { validations: validations } : {}),
        };
      const v = t.axis === "row" ? "rowHeights" : "columnWidths",
        g = {};
      for (const [O, j] of Object.entries(sheet[v] ?? {})) {
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
      const merges = (sheet.merges ?? [])
          .map((O) =>
            t.moveBefore !== undefined
              ? shiftRangeForAxisEdit(O, t, "merged range", true)
              : assertRangeWithinSheetBounds(O, t),
          )
          .filter((O) => !!O),
        y = (validations ?? []).flatMap((O) => {
          const range = assertRangeWithinSheetBounds(O.range, t);
          return range ? [{ ...O, range: range }] : [];
        }),
        w = (conditionalFormats ?? []).flatMap((O) => {
          const range = assertRangeWithinSheetBounds(O.range, t);
          return range ? [{ ...O, range: range }] : [];
        }),
        objects = sheet.objects?.map((object) => {
          const anchor = assertInsertionWithinSheet(object.anchor, t);
          if (!anchor)
            throw Error(
              `This edit would remove the anchor of ${object.label || object.kind}. Move or delete that drawing first.`,
            );
          if (object.kind !== "chart" || !object.range)
            return { ...object, anchor: anchor };
          const rangeAddress = parseRangeAddress(object.range),
            V = t.index + t.count - 1;
          if (
            t.remove &&
            t.axis === "row" &&
            object.headers &&
            t.index <= rangeAddress.firstRow &&
            V >= rangeAddress.firstRow
          )
            throw Error(
              "This edit would remove a chart header row. Adjust or delete the chart first.",
            );
          if (
            t.remove &&
            t.axis === "column" &&
            rangeAddress.lastCol > rangeAddress.firstCol &&
            t.index <= rangeAddress.firstCol &&
            V >= rangeAddress.firstCol
          )
            throw Error(
              "This edit would remove chart category labels. Adjust or delete the chart first.",
            );
          const range = assertRangeWithinSheetBounds(object.range, t);
          if (!range)
            throw Error(
              "This edit would remove the entire chart source. Adjust or delete the chart first.",
            );
          const B = measureChartSelection(range, sheet);
          if (
            t.moveBefore !== undefined &&
            ((t.axis === "row" &&
              object.headers &&
              remapAxisIndex(rangeAddress.firstRow, t) !== B.firstRow) ||
              (t.axis === "column" &&
                rangeAddress.lastCol > rangeAddress.firstCol &&
                remapAxisIndex(rangeAddress.firstCol, t) !== B.firstCol))
          )
            throw Error(
              "This move would displace chart headers or category labels. Move the complete chart source or only its data rows or value columns.",
            );
          if (
            (object.headers && B.lastRow === B.firstRow) ||
            (rangeAddress.lastCol > rangeAddress.firstCol &&
              B.lastCol === B.firstCol)
          )
            throw Error(
              "This edit would remove all chart data. Adjust or delete the chart first.",
            );
          const chartOptions = object.chartOptions
            ? structuredClone(object.chartOptions)
            : undefined;
          if (chartOptions) {
            const C = (T, N) => {
              const R = remapAxisIndex(T - 1, t);
              if (R === null)
                throw Error(
                  `This edit would remove a selected chart ${N}. Reset the chart selection before deleting it.`,
                );
              return R + 1;
            };
            (t.axis === "row" &&
              chartOptions.includedRows &&
              (chartOptions.includedRows = chartOptions.includedRows.map(
                (includedRow) => C(includedRow, "filter row"),
              )),
              t.axis === "column" &&
                chartOptions.series &&
                (chartOptions.series = chartOptions.series.map((sery) => ({
                  ...sery,
                  column: C(sery.column, "series"),
                }))));
          }
          return {
            ...object,
            anchor: anchor,
            range: range,
            ...(chartOptions ? { chartOptions: chartOptions } : {}),
          };
        }),
        pageLayout =
          t.moveBefore !== undefined
            ? shiftPrintLayoutForAxis(sheet, t)
            : sheet.pageLayout
              ? shiftPrintSettingsOnAxisEdit(
                  sheet.pageLayout,
                  sheet,
                  t.axis,
                  t.index,
                  t.count,
                  t.remove,
                )
              : undefined;
      if (pageLayout) {
        const O = assertRangeWithinSheetBounds(pageLayout.printArea, t);
        if (!O)
          throw Error(
            "This edit would remove the entire print area. Change the print area first.",
          );
        pageLayout.printArea = O;
      }
      return {
        ...sheet,
        cells: cells,
        ...(t.moveBefore !== undefined
          ? collectAxisStructureConflicts(sheet, t)
          : applyAxisStructureEdit(
              { ...sheet, merges: merges },
              t.axis,
              t.index,
              t.count,
              t.remove,
            )),
        ...(t.axis === "row" && sheet.rowOriginIds
          ? {
              rowOriginIds: Object.fromEntries(
                Object.entries(sheet.rowOriginIds).flatMap(([O, j]) => {
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
        merges: merges,
        validations: y,
        conditionalFormats: w,
        ...(objects ? { objects: objects } : {}),
        ...(pageLayout ? { pageLayout: pageLayout } : {}),
      };
    }),
    definedNames = o.definedNames?.map((definedName) => {
      const h = definedName.scopeSheetId
          ? (o.sheets.find((sheet) => sheet.id === definedName.scopeSheetId)
              ?.name ?? "")
          : "",
        refersTo = rewriteFormulaForStructuralEdit(
          definedName.refersTo,
          h,
          n.name,
          t,
        );
      if (
        !definedName.scopeSheetId &&
        rewriteFormulaForStructuralEdit(
          definedName.refersTo,
          n.name,
          n.name,
          t,
        ) !== refersTo
      )
        throw Error(
          `Workbook-scoped name ${definedName.name} has ambiguous unqualified worksheet references. Qualify those references before inserting or deleting rows or columns.`,
        );
      return { ...definedName, refersTo: refersTo };
    }),
    u = shiftReviewTargetsOnAxisEdit(
      shiftFormulaWatchesOnAxisEdit(
        {
          ...o,
          sheets: sheets,
          ...(definedNames ? { definedNames: definedNames } : {}),
        },
        sheetId,
        t,
      ),
      sheetId,
      t,
    );
  for (const sheet of u.sheets) {
    const h = workbook.sheets.find((f) => f.id === sheet.id);
    sheet.id !== sheetId &&
      JSON.stringify(sheet.cells) !== JSON.stringify(h.cells) &&
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

function resolveDuplicateReviewScope(workbook, r) {
  return (
    r.includeHidden !== undefined &&
      typeof r.includeHidden != "boolean" &&
      failWithMessage("Include hidden must be boolean."),
    r.allowedSheetIds !== undefined &&
      (!Array.isArray(r.allowedSheetIds) ||
        r.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The duplicate review host context is invalid."),
    {
      includeHidden: !!r.includeHidden,
      allowedSheetIds: [
        ...new Set(
          r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    }
  );
}
function Jde(workbook, r, t, address) {
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
          rewriteStructuredReferences(workbook, t, l, address);
      o += c ? "0" : l;
      continue;
    }
    o += r[a++];
  }
  return o;
}
function Qde(workbook, outputSheetName) {
  let t = 0,
    n = false;
  const o = (a, s, address) => {
    if (!a) return;
    ((t += a.length),
      t > 1e6 &&
        failWithMessage(
          "Output-name safety inspection exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ));
    const u = Jde(workbook, a.replace(/"(?:[^"]|"")*"/g, ""), s, address);
    for (const l of u.matchAll(
      /('(?:[^']|'')*'|[A-Za-z_.\d\u007F-\uFFFF]+)\s*!/g,
    )) {
      const c = l[1];
      (c.startsWith("'")
        ? c.slice(1, -1).replace(/''/g, "'")
        : c
      ).toLowerCase() === outputSheetName.toLowerCase() && (n = true);
    }
  };
  for (const sheet of workbook.sheets) {
    for (const [s, cell] of Object.entries(sheet.cells))
      isFormulaInputValue(cell) && o(String(cell.raw), sheet.id, s);
    for (const s of sheet.conditionalFormats ?? [])
      if (s.kind === "formula") {
        const rangeAddress = parseRangeAddress(s.range);
        o(
          s.formula,
          sheet.id,
          formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol),
        );
      }
    for (const s of sheet.tables ?? []) {
      const rangeAddress = parseRangeAddress(s.range);
      s.columns.forEach((column, l) =>
        o(
          column.calculatedFormula,
          sheet.id,
          formatCellAddress(
            Math.min(rangeAddress.firstRow + 1, rangeAddress.lastRow),
            rangeAddress.firstCol + l,
          ),
        ),
      );
    }
  }
  for (const a of workbook.definedNames ?? [])
    o(a.refersTo, a.scopeSheetId ?? workbook.sheets[0].id);
  return n;
}
function normalizeDuplicateReviewRequest(workbook, request) {
  ((!request ||
    typeof request != "object" ||
    Object.keys(request).some(
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
    (typeof request.sheetId != "string" || typeof request.range != "string") &&
      failWithMessage("Choose a source worksheet and one rectangular range."));
  const rangeAddress = parseRangeAddress(request.range),
    n = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    o = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  (n * o > structuralEditLimits.sourceCells &&
    failWithMessage("Select at most 500 source cells for duplicate review."),
    (!Array.isArray(request.keyColumns) ||
      !request.keyColumns.length ||
      request.keyColumns.some(
        (keyColumn) =>
          !Number.isInteger(keyColumn) ||
          keyColumn < rangeAddress.firstCol ||
          keyColumn > rangeAddress.lastCol,
      )) &&
      failWithMessage(
        "Choose one or more key columns inside the source range.",
      ),
    ([request.headers, request.caseSensitive, request.trimText].some(
      (s) => typeof s != "boolean",
    ) ||
      !["raw", "calculated"].includes(request.comparison) ||
      !["match", "keep"].includes(request.blankPolicy)) &&
      failWithMessage(
        "Choose valid comparison, header, case, trim and blank rules.",
      ),
    o <= (request.headers ? 1 : 0) &&
      failWithMessage(
        "The source must contain at least one data row after its header.",
      ),
    typeof request.outputSheetName != "string" &&
      failWithMessage("Choose a new output worksheet name."));
  const outputSheetName = request.outputSheetName.trim();
  return (
    (!outputSheetName ||
      outputSheetName.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(outputSheetName) ||
      /^'|'$/.test(outputSheetName)) &&
      failWithMessage(
        "The output sheet needs a valid Excel worksheet name of 1–31 characters.",
      ),
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === outputSheetName.toLowerCase(),
    ) && failWithMessage("A worksheet with that output name already exists."),
    Qde(workbook, outputSheetName) &&
      failWithMessage(
        "That output name is already referenced by workbook formulas or formula metadata. Choose a different name to avoid reconnecting an unresolved source reference.",
      ),
    {
      ...request,
      range: `${formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol)}:${formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol)}`,
      keyColumns: [...new Set(request.keyColumns)].sort((s, d) => s - d),
      outputSheetName: outputSheetName,
    }
  );
}
function Pz(workbook, r) {
  const t = new Set(workbook.sheets.map((sheet) => sheet.id)),
    n = (a) => {
      if (!(!a || typeof a != "object"))
        for (const [s, d] of Object.entries(a))
          /sheetId$/i.test(s) && typeof d == "string"
            ? t.add(d)
            : typeof d == "object" && n(d);
    };
  n(workbook);
  let o = 1;
  for (; t.has(`${r}-${o}`);) o++;
  return `${r}-${o}`;
}
function runDuplicateInspection(workbook, sheet, range, duplicateReviewScope) {
  const rangeAddress = parseRangeAddress(range),
    a = [];
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  )
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    )
      a.push(formatCellAddress(firstRow, firstCol));
  if (
    (a.some((v) => !sheet.cells[v] || sheet.cells[v].raw === null) &&
      failWithMessage(
        "Calculated comparison requires populated source cells; blank or possible spill-child cells are outside this subset. Use raw comparison or a populated selection.",
      ),
    !a.some((v) => isFormulaInputValue(sheet.cells[v])))
  )
    return new Map(a.map((v) => [v, sheet.cells[v].raw]));
  const s = {
    id: Pz(workbook, "duplicate-inspection"),
    name: "Duplicate inspection",
    rowCount: workbookLimits.maxRows,
    colCount: workbookLimits.maxCols,
    cells: {
      A1: { raw: `=SUM('${sheet.name.replace(/'/g, "''")}'!${range})` },
    },
  };
  let d = false;
  for (let v = 0; v < 4; v++) {
    s.name = `_dup_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`;
    const g = s.name.toLowerCase();
    if (
      !workbook.sheets.some(
        (sheet) =>
          sheet.name.toLowerCase() === g ||
          Object.values(sheet.cells).some(
            (cell) =>
              typeof cell.raw == "string" && cell.raw.toLowerCase().includes(g),
          ),
      ) &&
      !workbook.definedNames?.some((definedName) =>
        definedName.refersTo.toLowerCase().includes(g),
      )
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
        ...workbook,
        sheets: [
          s,
          ...workbook.sheets.map((sheet) =>
            sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
          ),
        ],
      },
      {
        sheetId: s.id,
        address: "A1",
        direction: "precedents",
        depth: structuralEditLimits.depth,
        includeHidden: duplicateReviewScope.includeHidden,
        allowedSheetIds: [...duplicateReviewScope.allowedSheetIds, s.id],
        limits: {
          maxNodes: structuralEditLimits.dependencyCells,
          maxFormulaCells: structuralEditLimits.formulaCells,
          maxReferences: 1e4,
        },
      },
    ),
    l = new Set(u.nodes.map((node) => node.key));
  (u.coverage.truncated ||
    u.coverage.depthLimited ||
    u.cycles.length ||
    u.issues.some(
      (issue) => issue.from === u.target.key || l.has(issue.from),
    )) &&
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
  for (const node of u.nodes) {
    const sheet = workbook.sheets.find((sheet) => sheet.id === node.sheetId),
      cell = sheet.cells[node.address];
    if (
      ((!cell || cell.raw === null) &&
        failWithMessage(
          "Calculated comparison requires populated static dependencies; blank or possible spill-child dependencies are unsupported.",
        ),
      isFormulaInputValue(cell))
    ) {
      const y = resolveDefinedNameValue(
        String(cell.raw),
        workbook,
        sheet.id,
      ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
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
      ...workbook,
      sheets: workbook.sheets
        .filter((sheet) =>
          duplicateReviewScope.allowedSheetIds.includes(sheet.id),
        )
        .map((v) => ({
          ...v,
          cells: Object.fromEntries(
            Object.entries(v.cells).filter(([g]) =>
              l.has(buildCompositeKey(v.id, g)),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(h),
    p = new Map();
  for (const node of u.nodes) {
    const g = workbookWithSettings(node.sheetId, node.address);
    (workbookWithSettings.getOperationalErrors().length &&
      failWithMessage(
        `An operational calculation error (${workbookWithSettings.getOperationalErrors().join(", ")}) prevents duplicate review, even when a formula handler conceals it.`,
      ),
      workbookWithSettings.getSpillRange(node.sheetId, node.address) &&
        failWithMessage(
          "Calculated comparison requires scalar cells; a dependency produced an array/spill result.",
        ),
      isFormulaInputValue(
        workbook.sheets.find((sheet) => sheet.id === node.sheetId).cells[
          node.address
        ],
      ) &&
        typeof g == "string" &&
        /^#(?:[A-Z0-9_/.]+[!?]?)/.test(g) &&
        failWithMessage(
          "Calculated comparison cannot group formula errors or error-like formula strings; correct them or use raw comparison.",
        ),
      typeof g == "number" &&
        !Number.isFinite(g) &&
        failWithMessage("Nonfinite values cannot be used in duplicate review."),
      node.sheetId === sheet.id &&
        a.includes(node.address) &&
        p.set(node.address, g));
  }
  return (
    (workbookWithSettings.getOperationalErrors().length ||
      p.size !== a.length) &&
      failWithMessage(
        "Calculated source coverage is incomplete or encountered an operational error.",
      ),
    p
  );
}
function createDuplicateReviewSheet(workbook, request, t = {}) {
  (t.readOnly &&
    failWithMessage(
      "This workbook is read-only. A derived worksheet cannot be created.",
    ),
    assertWorkbookStructureEditable(workbook));
  const duplicateReviewRequest = normalizeDuplicateReviewRequest(
      workbook,
      request,
    ),
    duplicateReviewScope = resolveDuplicateReviewScope(workbook, t),
    sheet = workbook.sheets.find(
      (sheet) => sheet.id === duplicateReviewRequest.sheetId,
    );
  sheet || failWithMessage("The source worksheet no longer exists.");
  const rangeAddress = parseRangeAddress(duplicateReviewRequest.range),
    d = rangeAddress.lastCol - rangeAddress.firstCol + 1;
  ((!duplicateReviewScope.allowedSheetIds.includes(sheet.id) ||
    (!duplicateReviewScope.includeHidden &&
      (sheet.hidden ||
        sheet.veryHidden ||
        cn(sheet, "row").some(
          (j) =>
            j.start <= rangeAddress.lastRow && j.end >= rangeAddress.firstRow,
        ) ||
        cn(sheet, "column").some(
          (j) =>
            j.start <= rangeAddress.lastCol && j.end >= rangeAddress.firstCol,
        )))) &&
    failWithMessage(
      "The source range is outside the visible, authorized host context. Include hidden only within permitted host scope.",
    ),
    (sheet.merges ?? []).some((j) => {
      const z = parseRangeAddress(j);
      return (
        z.firstRow <= rangeAddress.lastRow &&
        z.lastRow >= rangeAddress.firstRow &&
        z.firstCol <= rangeAddress.lastCol &&
        z.lastCol >= rangeAddress.firstCol
      );
    }) &&
      failWithMessage(
        "Unmerge the selected source before duplicate review so row boundaries are unambiguous.",
      ));
  const u =
      duplicateReviewRequest.comparison === "calculated"
        ? runDuplicateInspection(
            workbook,
            sheet,
            duplicateReviewRequest.range,
            duplicateReviewScope,
          )
        : null,
    l = (j, z) => {
      const cellAddress = formatCellAddress(j, z),
        cell = structuredClone(sheet.cells[cellAddress] ?? { raw: null }),
        B = u ? u.get(cellAddress) : cell.raw;
      return (
        typeof B == "number" &&
          !Number.isFinite(B) &&
          failWithMessage(
            "Nonfinite values cannot be used in duplicate review.",
          ),
        { cell: cell, value: B }
      );
    },
    headerCells = Array.from({ length: d }, (j, z) =>
      duplicateReviewRequest.headers
        ? l(rangeAddress.firstRow, rangeAddress.firstCol + z).cell
        : { raw: `Column ${columnIndexToLetters(rangeAddress.firstCol + z)}` },
    ),
    rows = [],
    f = new Map();
  for (
    let j = rangeAddress.firstRow + (duplicateReviewRequest.headers ? 1 : 0);
    j <= rangeAddress.lastRow;
    j++
  ) {
    const z = Array.from({ length: d }, (L, $) =>
        l(j, rangeAddress.firstCol + $),
      ),
      cells = z.map((L) => L.cell),
      values = z.map((L) => L.value),
      B = duplicateReviewRequest.keyColumns.map((keyColumn) => {
        const $ = keyColumn - rangeAddress.firstCol,
          U = cells[$];
        let G = values[$];
        return duplicateReviewRequest.comparison === "raw" &&
          isFormulaInputValue(U)
          ? ["formula", G]
          : (typeof G == "string" &&
              (duplicateReviewRequest.trimText && (G = G.trim()),
              duplicateReviewRequest.caseSensitive || (G = G.toLowerCase())),
            G === null || G === "" ? ["blank"] : [typeof G, G]);
      }),
      blankKey = B.some((L) => L[0] === "blank"),
      normalizedKey = JSON.stringify(B),
      T = {
        sourceRow: j + 1,
        sourceRange: `${formatCellAddress(j, rangeAddress.firstCol)}:${formatCellAddress(j, rangeAddress.lastCol)}`,
        cells: cells,
        values: values,
        normalizedKey: normalizedKey,
        blankKey: blankKey,
        disposition: "retained",
        groupId: "",
        firstSourceRow: j + 1,
        groupSize: 1,
        outputRow: 0,
      },
      N =
        blankKey && duplicateReviewRequest.blankPolicy === "keep"
          ? `blank-row:${j}`
          : normalizedKey,
      R = f.get(N) ?? [];
    (R.push(T), f.set(N, R), rows.push(T));
  }
  const groups = [];
  for (const j of f.values())
    if (j.length > 1) {
      const groupId = `G${groups.length + 1}`,
        sourceRow = j[0].sourceRow;
      groups.push({
        id: groupId,
        firstSourceRow: sourceRow,
        memberRows: j.map((D) => D.sourceRow),
        normalizedKey: j[0].normalizedKey,
      });
      for (let D = 0; D < j.length; D++)
        Object.assign(j[D], {
          groupId: groupId,
          firstSourceRow: sourceRow,
          groupSize: j.length,
          disposition: D === 0 ? "retained" : "excluded",
        });
    }
  const m = rows.filter((row) => row.disposition === "retained"),
    v = rows.filter((row) => row.disposition === "excluded"),
    counts = {
      input: rows.length,
      retained: m.length,
      excluded: v.length,
      duplicateGroups: groups.length,
      blankKeyRows: rows.filter((row) => row.blankKey).length,
    },
    b = {
      id: Pz(workbook, "duplicate-review"),
      name: duplicateReviewRequest.outputSheetName,
      rowCount: workbookLimits.maxRows,
      colCount: workbookLimits.maxCols,
      cells: {},
      frozenRows: 3,
    },
    y = (raw) => ({
      raw: raw,
      ...(typeof raw == "string" && raw.startsWith("=")
        ? { literal: true }
        : {}),
    }),
    w = (j, z, V) => {
      b.cells[formatCellAddress(j - 1, z)] = V;
    };
  (w(1, 0, y("Unique population — static source snapshot")),
    w(
      2,
      0,
      y(
        `${counts.input} source rows; ${counts.retained} retained; ${counts.excluded} excluded from population. ${duplicateReviewRequest.comparison} typed keys ${duplicateReviewRequest.keyColumns.map(columnIndexToLetters).join(", ")}; ${duplicateReviewRequest.headers ? "source header" : "no source header"}; ${duplicateReviewRequest.caseSensitive ? "case-sensitive" : "ignore case"}; ${duplicateReviewRequest.trimText ? "trim outer text" : "preserve whitespace"}; blanks ${duplicateReviewRequest.blankPolicy}. First source row wins. Source unchanged.`,
      ),
    ));
  const E = (j) => {
      (headerCells.forEach((cell, V) => {
        const D =
          duplicateReviewRequest.headers && u
            ? u.get(
                formatCellAddress(
                  rangeAddress.firstRow,
                  rangeAddress.firstCol + V,
                ),
              )
            : cell.raw;
        w(j, V, {
          ...structuredClone(cell),
          ...y(D),
          literal:
            typeof D == "string" && D.startsWith("=") ? true : cell.literal,
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
        j.cells.forEach((cell, B) => {
          const raw = j.values[B];
          w(z, B, {
            ...structuredClone(cell),
            raw: raw,
            ...(typeof raw == "string" && raw.startsWith("=")
              ? { literal: true }
              : {}),
          });
        }),
        [
          JSON.stringify({
            workbookId: workbook.id,
            revision: workbook.revision,
            sheetId: sheet.id,
            sheetName: sheet.name,
            range: j.sourceRange,
          }),
          j.groupId,
          j.firstSourceRow,
          j.groupSize,
        ].forEach((D, B) => w(z, d + B, y(D))));
    };
  (E(3), m.forEach((j, z) => A(j, 4 + z)));
  const populationRange = `A3:${columnIndexToLetters(d + 3)}${3 + m.length}`;
  let excludedRange = null;
  if (v.length) {
    const j = m.length + 6;
    (w(
      j,
      0,
      y("Excluded duplicate rows — retained here and unchanged on the source"),
    ),
      E(j + 1),
      v.forEach((z, V) => A(z, j + 2 + V)),
      (excludedRange = `A${j + 1}:${columnIndexToLetters(d + 3)}${j + 1 + v.length}`));
  }
  workbook.sheets.reduce((j, z) => j + Object.keys(z.cells).length, 0) +
    Object.keys(b.cells).length >
    workbookLimits.maxCells &&
    failWithMessage(
      "The derived worksheet would exceed the workbook's stored-cell budget.",
    );
  const next = structuredClone(workbook);
  return (
    next.sheets.push(b),
    new TextEncoder().encode(JSON.stringify(next)).byteLength >
      workbookLimits.maxNativeBytes &&
      failWithMessage(
        "The derived workbook would exceed the native file budget.",
      ),
    {
      snapshot: JSON.stringify(workbook),
      request: duplicateReviewRequest,
      context: duplicateReviewScope,
      headerCells: headerCells,
      rows: rows,
      groups: groups,
      counts: counts,
      outputSheetId: b.id,
      populationRange: populationRange,
      excludedRange: excludedRange,
      next: next,
    }
  );
}
function applyDuplicateReviewPreview(workbook, r, t = {}) {
  (t.readOnly && failWithMessage("This workbook is read-only."),
    JSON.stringify(workbook) !== r.snapshot &&
      failWithMessage(
        "The workbook changed after duplicate review. Preview again before Apply.",
      ));
  const duplicateReviewSheet = createDuplicateReviewSheet(
    workbook,
    r.request,
    t,
  );
  return (
    JSON.stringify(duplicateReviewSheet) !== JSON.stringify(r) &&
      failWithMessage(
        "The duplicate preview or host context changed. Preview and review again.",
      ),
    duplicateReviewSheet.next
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
      const rangeAddress = parseRangeAddress(e);
      return rangeAddress.lastCol - rangeAddress.firstCol + 1 <= 500
        ? Array.from(
            { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
            (t, n) => rangeAddress.firstCol + n,
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
