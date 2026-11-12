import i from "react/jsx-runtime";
import { X, ShieldCheck, Grid2x2, Paintbrush, CheckCheck } from "lucide-react";
import ne from "react";
import {
  CE,
  setTableColumnDataType,
  getPivotRefreshState,
  updateSheetProtection,
  previewColumnTypeChange,
  tableColumnDataTypes,
  formatOccupiedRangeAddress,
  validateRemoveDuplicatesRequest,
  applyColumnTypePreview,
  parseRangeAddress,
  renameTableHeaderField,
  mergeWorksheetCells,
  getPivotSourceFields,
  paintCellFormatting,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  applyStructuralEdit,
  DataValidationDialog,
} from "./27-validation-and-structural-edits.js";
import {
  DuplicateReviewDialog,
  StructureToolDialog,
} from "./28-duplicate-review-and-structure.js";

function DataToolDialogRouter(e) {
  return e.mode === "validation"
    ? i.jsx(DataValidationDialog, { ...e })
    : e.mode === "duplicates"
      ? i.jsx(DuplicateReviewDialog, { ...e })
      : e.mode === "insertRow" || e.mode === "insertColumn"
        ? i.jsx(StructureToolDialog, { ...e })
        : i.jsx(WorksheetToolDialog, { ...e });
}
function WorksheetToolDialog({
  mode,
  workbook,
  sheetId,
  range,
  readOnly,
  onClose,
  onWorkbookChange,
}) {
  const { text, language } = useWorkbookMessages(),
    l = workbook.sheets.find((sheet) => sheet.id === sheetId),
    c = range.split(":")[0],
    [h, f] = ne.useState(
      mode === "duplicates" && range === c
        ? formatOccupiedRangeAddress(l)
        : range,
    ),
    [p, m] = ne.useState(mode === "conditional" ? "greater" : "list"),
    [v, g] = ne.useState(""),
    [b, y] = ne.useState("100"),
    [fill, E] = ne.useState("#fff2cc"),
    [color, k] = ne.useState("#9c5700"),
    [operator, I] = ne.useState("between"),
    [O, j] = ne.useState(true),
    [z, V] = ne.useState("Enter a value that satisfies this rule."),
    [checked, B] = ne.useState(true),
    [_, C] = ne.useState(null),
    [T, N] = ne.useState(""),
    [count, L] = ne.useState(1),
    [remove, U] = ne.useState(false),
    [G, H] = ne.useState(""),
    [Q, X$1] = ne.useState(""),
    [J, q] = ne.useState(false),
    [W, K] = ne.useState(""),
    te = ne.useRef(null),
    ae = Object.entries(l.cells)
      .filter(([Re, cell]) => {
        try {
          const rangeAddress = parseRangeAddress(h),
            je = parseRangeAddress(Re);
          return (
            typeof cell.raw == "string" &&
            !cell.raw.startsWith("=") &&
            je.firstRow >= rangeAddress.firstRow &&
            je.lastRow <= rangeAddress.lastRow &&
            je.firstCol >= rangeAddress.firstCol &&
            je.lastCol <= rangeAddress.lastCol
          );
        } catch {
          return false;
        }
      })
      .slice(0, 500),
    [oe, de] = ne.useState(0),
    [raw, Se] = ne.useState(ae[0]?.[1].raw ?? "");
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      te.current?.querySelector("input,textarea,select,button")?.focus(),
      () => activeElement?.focus()
    );
  }, []);
  const le = {
      conditional: "Conditional Formatting",
      validation: "Data Validation",
      duplicates: "Remove Duplicates",
      protection: l.protection?.enabled ? "Unprotect Sheet" : "Protect Sheet",
      merge: "Merge & Center",
      paint: "Format Painter",
      spelling: "Spelling",
      insertRow: "Insert / Delete Rows",
      insertColumn: "Insert / Delete Columns",
    },
    ce = (Re, Me) => {
      (onWorkbookChange(Re, Me), onClose());
    };
  async function ge() {
    (H(""), q(true));
    try {
      if (readOnly && mode !== "spelling")
        throw Error("This workbook is read-only.");
      if (mode === "merge")
        ce(mergeWorksheetCells(workbook, sheetId, h, false), "Merge cells");
      else if (mode === "paint")
        ce(
          paintCellFormatting(workbook, sheetId, c, W),
          "Paint cell formatting",
        );
      else if (mode === "protection") {
        const ke = await updateSheetProtection(
          workbook,
          sheetId,
          T,
          !l.protection?.enabled,
        );
        ce(ke, l.protection?.enabled ? "Unprotect sheet" : "Protect sheet");
      } else if (mode === "conditional") {
        if (
          ["greater", "less", "between"].includes(p) &&
          (!v.trim() ||
            !Number.isFinite(Number(v)) ||
            (p === "between" && !Number.isFinite(Number(b))))
        )
          throw Error("Enter finite comparison values.");
        const ke = {
          id: crypto.randomUUID(),
          range: CE(h),
          kind: p,
          value: v,
          second: b,
          fill: fill,
          color: color,
        };
        ce(
          {
            ...workbook,
            sheets: workbook.sheets.map((sheet) =>
              sheet.id === sheetId
                ? {
                    ...sheet,
                    conditionalFormats: [
                      ...(sheet.conditionalFormats ?? []),
                      ke,
                    ],
                  }
                : sheet,
            ),
          },
          "Add conditional formatting",
        );
      } else if (mode === "validation") {
        const ke = {
          id: crypto.randomUUID(),
          range: CE(h),
          kind: p,
          operator: operator,
          min: Number(v || 0),
          max: Number(b),
          list: v
            .split(",")
            .map((je) => je.trim())
            .filter(Boolean),
          allowBlank: O,
          message: z,
        };
        ce(
          {
            ...workbook,
            sheets: workbook.sheets.map((sheet) =>
              sheet.id === sheetId
                ? { ...sheet, validations: [...(sheet.validations ?? []), ke] }
                : sheet,
            ),
          },
          "Add data validation",
        );
      } else if (mode === "duplicates") {
        const rangeAddress = parseRangeAddress(h),
          je =
            _ ??
            Array.from(
              { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
              (Je, ze) => rangeAddress.firstCol + ze,
            ),
          $e = validateRemoveDuplicatesRequest(
            workbook,
            sheetId,
            h,
            je,
            checked,
          );
        if (!$e.removed) {
          X$1("No duplicate rows found.");
          return;
        }
        ce($e.workbook, `Remove ${$e.removed} duplicate rows from selection`);
      } else if (mode === "insertRow" || mode === "insertColumn") {
        const rangeAddress = parseRangeAddress(h),
          je = {
            axis: mode === "insertRow" ? "row" : "column",
            index:
              mode === "insertRow"
                ? rangeAddress.firstRow
                : rangeAddress.firstCol,
            count: count,
            remove: remove,
          };
        ce(
          applyStructuralEdit(workbook, sheetId, je),
          `${remove ? "Delete" : "Insert"} ${count} ${je.axis}${count === 1 ? "" : "s"}`,
        );
      } else if (mode === "spelling") {
        if (readOnly) throw Error("This workbook is read-only.");
        const ke = ae[oe];
        if (!ke) throw Error("No text cells in the selected range.");
        const cells = { ...l.cells, [ke[0]]: { ...ke[1], raw: raw } };
        ce(
          renameTableHeaderField(workbook, {
            ...workbook,
            sheets: workbook.sheets.map((sheet) =>
              sheet.id === sheetId ? { ...sheet, cells: cells } : sheet,
            ),
          }),
          `Spelling correction at ${ke[0]}`,
        );
      }
    } catch (error) {
      H(
        error instanceof Error
          ? error.message
          : "The operation could not be completed.",
      );
    } finally {
      q(false);
    }
  }
  let se;
  try {
    se = parseRangeAddress(h);
  } catch {
    se = parseRangeAddress(range);
  }
  const we =
    se.lastCol - se.firstCol < 100
      ? Array.from(
          { length: se.lastCol - se.firstCol + 1 },
          (Re, Me) => se.firstCol + Me,
        )
      : [];
  return i.jsx("div", {
    className: "wt-backdrop",
    onPointerDown: (event) => {
      event.target === event.currentTarget && onClose();
    },
    children: i.jsxs("div", {
      className: "wt-dialog",
      ref: te,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": le[mode] ?? mode,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), onClose()),
          event.key === "Tab")
        ) {
          const ke = Array.from(
              te.current?.querySelectorAll(
                "button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled])",
              ) ?? [],
            ),
            je = ke[0],
            $e = ke.at(-1);
          event.shiftKey && document.activeElement === je
            ? (event.preventDefault(), $e == null || $e.focus())
            : !event.shiftKey &&
              document.activeElement === $e &&
              (event.preventDefault(), je == null || je.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("span", {
              children: [
                i.jsx(Grid2x2, { size: 19 }),
                i.jsx("h2", { children: le[mode] ?? mode }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close worksheet tool"),
              onClick: onClose,
              lang: language("Close worksheet tool"),
              "data-ui-label": "Close worksheet tool",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wt-body",
          children: [
            mode !== "protection" &&
              mode !== "paint" &&
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Range" }),
                  i.jsx("input", {
                    "aria-label": text("Tool range"),
                    value: h,
                    onChange: (event) => {
                      (f(event.target.value), C(null));
                    },
                    lang: language("Tool range"),
                    "data-ui-label": "Tool range",
                  }),
                ],
              }),
            mode === "paint" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("p", {
                    children: [
                      i.jsx(Paintbrush, { size: 16 }),
                      " ",
                      i.jsx(LocalizedText, {
                        children: "Copy formatting from",
                      }),
                      " ",
                      i.jsxs("strong", { children: [l.name, "!", c] }),
                      ".",
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Apply to range" }),
                      i.jsx("input", {
                        "aria-label": text("Format destination"),
                        value: W,
                        onChange: (event) => K(event.target.value),
                        placeholder: text("B2:D10"),
                        lang: language("Format destination"),
                        "data-ui-label": "Format destination",
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Values, formulas, notes and review marks stay in the destination.",
                    }),
                  }),
                ],
              }),
            mode === "merge" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "The selected rectangle becomes one displayed cell, centered on its first value. Covered values remain stored and reappear when unmerged.",
                    }),
                  }),
                  i.jsx("button", {
                    className: "wt-secondary",
                    disabled: readOnly,
                    onClick: () => {
                      try {
                        ce(
                          mergeWorksheetCells(workbook, sheetId, h, true),
                          "Unmerge cells",
                        );
                      } catch (error) {
                        H(String(error));
                      }
                    },
                    children: i.jsx(LocalizedText, {
                      children: "Unmerge selected cells",
                    }),
                  }),
                ],
              }),
            mode === "conditional" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Rule" }),
                      i.jsxs("select", {
                        "aria-label": text("Conditional rule"),
                        value: p,
                        onChange: (event) => m(event.target.value),
                        lang: language("Conditional rule"),
                        "data-ui-label": "Conditional rule",
                        children: [
                          i.jsx("option", {
                            value: "greater",
                            children: "Greater than",
                          }),
                          i.jsx("option", {
                            value: "less",
                            children: "Less than",
                          }),
                          i.jsx("option", {
                            value: "equal",
                            children: "Equal to",
                          }),
                          i.jsx("option", {
                            value: "between",
                            children: "Between",
                          }),
                          i.jsx("option", {
                            value: "contains",
                            children: "Text contains",
                          }),
                          i.jsx("option", {
                            value: "duplicate",
                            children: "Duplicate values",
                          }),
                          i.jsx("option", {
                            value: "colorScale",
                            children: "Green color scale",
                          }),
                        ],
                      }),
                    ],
                  }),
                  !["duplicate", "colorScale"].includes(p) &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Value" }),
                        i.jsx("input", {
                          "aria-label": text("Conditional value"),
                          value: v,
                          onChange: (event) => g(event.target.value),
                          lang: language("Conditional value"),
                          "data-ui-label": "Conditional value",
                        }),
                      ],
                    }),
                  p === "between" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Maximum" }),
                        i.jsx("input", {
                          "aria-label": text("Conditional maximum"),
                          value: b,
                          onChange: (event) => y(event.target.value),
                          lang: language("Conditional maximum"),
                          "data-ui-label": "Conditional maximum",
                        }),
                      ],
                    }),
                  i.jsxs("div", {
                    className: "wt-columns",
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Fill" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": text("Conditional fill"),
                            value: fill,
                            onChange: (event) => E(event.target.value),
                            lang: language("Conditional fill"),
                            "data-ui-label": "Conditional fill",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Text" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": text("Conditional color"),
                            value: color,
                            onChange: (event) => k(event.target.value),
                            lang: language("Conditional color"),
                            "data-ui-label": "Conditional color",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Rules update when calculated cell values change.",
                    }),
                  }),
                  l.conditionalFormats?.map((conditionalFormat) =>
                    i.jsxs(
                      "div",
                      {
                        className: "wt-existing",
                        children: [
                          i.jsxs("span", {
                            children: [
                              conditionalFormat.range,
                              " · ",
                              conditionalFormat.kind,
                            ],
                          }),
                          i.jsx("button", {
                            disabled: readOnly,
                            onClick: () =>
                              ce(
                                {
                                  ...workbook,
                                  sheets: workbook.sheets.map((sheet) =>
                                    sheet.id === sheetId
                                      ? {
                                          ...sheet,
                                          conditionalFormats:
                                            sheet.conditionalFormats?.filter(
                                              (je) =>
                                                je.id !== conditionalFormat.id,
                                            ),
                                        }
                                      : sheet,
                                  ),
                                },
                                "Remove conditional formatting",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Remove rule",
                            }),
                          }),
                        ],
                      },
                      conditionalFormat.id,
                    ),
                  ),
                ],
              }),
            mode === "validation" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Allow" }),
                      i.jsxs("select", {
                        "aria-label": text("Validation type"),
                        value: p,
                        onChange: (event) => m(event.target.value),
                        lang: language("Validation type"),
                        "data-ui-label": "Validation type",
                        children: [
                          i.jsx("option", { value: "list", children: "List" }),
                          i.jsx("option", {
                            value: "number",
                            children: "Decimal",
                          }),
                          i.jsx("option", {
                            value: "integer",
                            children: "Whole number",
                          }),
                          i.jsx("option", {
                            value: "date",
                            children: "Date (Excel serial)",
                          }),
                          i.jsx("option", {
                            value: "textLength",
                            children: "Text length",
                          }),
                        ],
                      }),
                    ],
                  }),
                  p !== "list" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Condition" }),
                        i.jsxs("select", {
                          "aria-label": text("Validation operator"),
                          value: operator,
                          onChange: (event) => I(event.target.value),
                          lang: language("Validation operator"),
                          "data-ui-label": "Validation operator",
                          children: [
                            i.jsx("option", {
                              value: "between",
                              children: "Between",
                            }),
                            i.jsx("option", {
                              value: "greater",
                              children: "Greater than",
                            }),
                            i.jsx("option", {
                              value: "less",
                              children: "Less than",
                            }),
                            i.jsx("option", {
                              value: "equal",
                              children: "Equal to",
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsxs("label", {
                    children: [
                      p === "list"
                        ? "Values, separated by commas"
                        : "Minimum / comparison value",
                      i.jsx("input", {
                        "aria-label": text("Validation values"),
                        value: v,
                        onChange: (event) => g(event.target.value),
                        placeholder:
                          p === "list" ? "Pending, Complete, Reviewed" : "0",
                        lang: language("Validation values"),
                        "data-ui-label": "Validation values",
                      }),
                    ],
                  }),
                  p !== "list" &&
                    operator === "between" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Maximum" }),
                        i.jsx("input", {
                          "aria-label": text("Validation maximum"),
                          value: b,
                          onChange: (event) => y(event.target.value),
                          lang: language("Validation maximum"),
                          "data-ui-label": "Validation maximum",
                        }),
                      ],
                    }),
                  i.jsxs("label", {
                    className: "wt-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: O,
                        onChange: (event) => j(event.target.checked),
                      }),
                      i.jsx(LocalizedText, { children: "Allow blank cells" }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Error message" }),
                      i.jsx("input", {
                        "aria-label": text("Validation message"),
                        value: z,
                        onChange: (event) => V(event.target.value),
                        lang: language("Validation message"),
                        "data-ui-label": "Validation message",
                      }),
                    ],
                  }),
                  l.validations?.map((validation) =>
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
                            ],
                          }),
                          i.jsx("button", {
                            disabled: readOnly,
                            onClick: () =>
                              ce(
                                {
                                  ...workbook,
                                  sheets: workbook.sheets.map((sheet) =>
                                    sheet.id === sheetId
                                      ? {
                                          ...sheet,
                                          validations:
                                            sheet.validations?.filter(
                                              (je) => je.id !== validation.id,
                                            ),
                                        }
                                      : sheet,
                                  ),
                                },
                                "Remove data validation",
                              ),
                            children: i.jsx(LocalizedText, {
                              children: "Remove rule",
                            }),
                          }),
                        ],
                      },
                      validation.id,
                    ),
                  ),
                ],
              }),
            mode === "duplicates" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    className: "wt-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: checked,
                        onChange: (event) => B(event.target.checked),
                      }),
                      i.jsx(LocalizedText, { children: "My data has headers" }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Compare these columns",
                        }),
                      }),
                      we.map((Re) =>
                        i.jsxs(
                          "label",
                          {
                            className: "wt-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: (_ ?? we).includes(Re),
                                onChange: (event) =>
                                  C(
                                    event.target.checked
                                      ? [...(_ ?? we), Re]
                                      : (_ ?? we).filter((je) => je !== Re),
                                  ),
                              }),
                              columnIndexToLetters(Re),
                              " ·",
                              " ",
                              String(
                                l.cells[
                                  `${columnIndexToLetters(Re)}${se.firstRow + 1}`
                                ]?.raw ?? "",
                              ),
                            ],
                          },
                          Re,
                        ),
                      ),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "The first occurrence is retained. Later duplicates are removed from the selected rectangle, remaining rows move up within it, and the vacated cells are cleared. Undo restores the full selection.",
                    }),
                  }),
                ],
              }),
            mode === "protection" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("p", {
                    children: [
                      i.jsx(ShieldCheck, { size: 17 }),
                      " ",
                      l.protection?.enabled
                        ? "Unlock editing on this worksheet."
                        : "Prevent changes to this worksheet until it is unprotected.",
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      l.protection?.hash ? "Password" : "Password (optional)",
                      i.jsx("input", {
                        type: "password",
                        "aria-label": text("Sheet password"),
                        value: T,
                        onChange: (event) => N(event.target.value),
                        autoComplete: "new-password",
                        lang: language("Sheet password"),
                        "data-ui-label": "Sheet password",
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "This prevents edits in the workbook editor. The embedding application controls access to files and data.",
                    }),
                  }),
                ],
              }),
            (mode === "insertRow" || mode === "insertColumn") &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Operation" }),
                      i.jsxs("select", {
                        "aria-label": text("Structure operation"),
                        value: remove ? "delete" : "insert",
                        onChange: (event) => U(event.target.value === "delete"),
                        lang: language("Structure operation"),
                        "data-ui-label": "Structure operation",
                        children: [
                          i.jsx("option", {
                            value: "insert",
                            children: "Insert before selection",
                          }),
                          i.jsx("option", {
                            value: "delete",
                            children: "Delete starting at selection",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Count" }),
                      i.jsx(LocalizedNumberField, {
                        min: "1",
                        max: "10000",
                        "aria-label": text("Structure count"),
                        value: count,
                        onValueChange: (Re) => L(Number(Re)),
                        lang: language("Structure count"),
                        "data-ui-label": "Structure count",
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Values, supported references, row/column sizes and supported range metadata move with the edit. Undo restores the previous workbook.",
                    }),
                  }),
                ],
              }),
            mode === "spelling" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("p", {
                    children: [
                      i.jsx(CheckCheck, { size: 17 }),
                      " ",
                      i.jsx(LocalizedText, {
                        children:
                          "Review text with your browser’s spelling dictionary. Right-click an underlined word to see available corrections.",
                      }),
                    ],
                  }),
                  ae.length
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, { children: "Cell" }),
                              i.jsx("select", {
                                "aria-label": text("Spelling cell"),
                                value: oe,
                                onChange: (event) => {
                                  const Me = Number(event.target.value);
                                  (de(Me), Se(String(ae[Me][1].raw)));
                                },
                                lang: language("Spelling cell"),
                                "data-ui-label": "Spelling cell",
                                children: ae.map(([Re], Me) =>
                                  i.jsx(
                                    "option",
                                    { value: Me, children: Re },
                                    Re,
                                  ),
                                ),
                              }),
                            ],
                          }),
                          i.jsx("textarea", {
                            "aria-label": text("Spelling text"),
                            spellCheck: true,
                            autoCorrect: "on",
                            value: raw,
                            onChange: (event) => Se(event.target.value),
                            readOnly,
                            lang: language("Spelling text"),
                            "data-ui-label": "Spelling text",
                          }),
                        ],
                      })
                    : i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "No text cells were found in this selection.",
                        }),
                      }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Spelling suggestions depend on the browser and its installed languages.",
                    }),
                  }),
                ],
              }),
            Q && i.jsx("p", { role: "status", children: Q }),
            G &&
              i.jsx("p", { className: "wt-error", role: "alert", children: G }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "wt-primary",
              disabled: J || readOnly,
              onClick: () => void ge(),
              children: J
                ? "Applying…"
                : mode === "spelling"
                  ? "Save correction"
                  : mode === "protection"
                    ? l.protection?.enabled
                      ? "Unprotect"
                      : "Protect"
                    : "Apply",
            }),
          ],
        }),
      ],
    }),
  });
}
function TableColumnTypesPanel({
  workbook,
  sheetId,
  tableId,
  context,
  readOnly,
  onWorkbookChange,
}) {
  const { text, language } = useWorkbookMessages(),
    u = workbook.sheets
      .find((sheet) => sheet.id === sheetId)
      .tables.find((table) => table.id === tableId),
    [columnId, c] = ne.useState(u.columns[0].id),
    [dataType, f] = ne.useState(u.columns[0].dataType ?? "any"),
    [mode, m] = ne.useState("convert"),
    [invalid, g] = ne.useState("reject"),
    [decimal, y] = ne.useState("."),
    [group, E] = ne.useState(","),
    [A, k] = ne.useState(null),
    [F, I] = ne.useState(""),
    [O, j] = ne.useState(0),
    request = {
      columnId: columnId,
      dataType: dataType,
      mode: mode,
      invalid: invalid,
      decimal: decimal,
      group: group,
    },
    V = JSON.stringify({ context, request: request }),
    D =
      !!A && (A.context !== V || A.plan.snapshot !== JSON.stringify(workbook)),
    B = ne.useMemo(() => {
      try {
        return setTableColumnDataType(workbook, sheetId, tableId, columnId);
      } catch {
        return null;
      }
    }, [workbook, sheetId, tableId, columnId]),
    _ = workbook.tableEvidence?.filter((T) => T.tableId === tableId) ?? [];
  return i.jsxs("details", {
    children: [
      i.jsx("summary", {
        children: i.jsx(LocalizedText, {
          children: "Column types and original source evidence",
        }),
      }),
      i.jsxs("fieldset", {
        className: "lg-table-list",
        children: [
          i.jsx("legend", {
            children: i.jsx(LocalizedText, { children: "Review column types" }),
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Typed column" }),
              i.jsx("select", {
                "aria-label": text("Typed column"),
                value: columnId,
                onChange: (event) => {
                  (c(event.target.value),
                    f(
                      u.columns.find(
                        (column) => column.id === event.target.value,
                      )?.dataType ?? "any",
                    ),
                    k(null));
                },
                lang: language("Typed column"),
                "data-ui-label": "Typed column",
                children: u.columns.map((column) =>
                  i.jsx(
                    "option",
                    { value: column.id, children: column.name },
                    column.id,
                  ),
                ),
              }),
            ],
          }),
          i.jsxs("p", {
            "aria-label": text("Current column type"),
            lang: language("Current column type"),
            "data-ui-label": "Current column type",
            children: [
              i.jsx(LocalizedText, { children: "Saved type: " }),
              B?.dataType ?? "unknown",
              "; ",
              B?.checked ?? 0,
              " ",
              i.jsx(LocalizedText, { children: "rows checked; " }),
              B?.exceptions.length ?? 0,
              " ",
              i.jsx(LocalizedText, { children: "exceptions" }),
              B?.truncated ? " — partial inspection (first 500 rows)" : "",
              ".",
            ],
          }),
          !!B?.exceptions.length &&
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(LocalizedText, {
                    children: "Current type exceptions",
                  }),
                }),
                B.exceptions
                  .slice(0, 50)
                  .map((T) =>
                    i.jsxs(
                      "p",
                      { children: [T.address, ": ", JSON.stringify(T.value)] },
                      T.address,
                    ),
                  ),
              ],
            }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Column type" }),
              i.jsx("select", {
                "aria-label": text("Column type"),
                value: dataType,
                onChange: (event) => f(event.target.value),
                lang: language("Column type"),
                "data-ui-label": "Column type",
                children: tableColumnDataTypes.map((tableColumnDataType) =>
                  i.jsx(
                    "option",
                    {
                      value: tableColumnDataType,
                      children:
                        tableColumnDataType === "date"
                          ? "Date (YYYY-MM-DD text)"
                          : tableColumnDataType,
                    },
                    tableColumnDataType,
                  ),
                ),
              }),
            ],
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Type operation" }),
              i.jsxs("select", {
                value: mode,
                onChange: (event) => m(event.target.value),
                children: [
                  i.jsx("option", {
                    value: "convert",
                    children: "Convert compatible values",
                  }),
                  i.jsx("option", {
                    value: "tag",
                    children: "Set type; keep every source value",
                  }),
                ],
              }),
            ],
          }),
          dataType === "number" &&
            mode === "convert" &&
            i.jsxs("div", {
              className: "lg-table-inline",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Decimal separator" }),
                    i.jsxs("select", {
                      value: decimal,
                      onChange: (event) => y(event.target.value),
                      children: [
                        i.jsx("option", { value: ".", children: "Period" }),
                        i.jsx("option", { value: ",", children: "Comma" }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Group separator" }),
                    i.jsxs("select", {
                      value: group,
                      onChange: (event) => E(event.target.value),
                      children: [
                        i.jsx("option", { value: ",", children: "Comma" }),
                        i.jsx("option", { value: ".", children: "Period" }),
                        i.jsx("option", { value: " ", children: "Space" }),
                        i.jsx("option", { value: "", children: "None" }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Unconvertible values" }),
              i.jsxs("select", {
                value: invalid,
                onChange: (event) => g(event.target.value),
                children: [
                  i.jsx("option", {
                    value: "reject",
                    children: "Block apply until resolved",
                  }),
                  i.jsx("option", {
                    value: "keep",
                    children: "Keep exceptions unchanged",
                  }),
                ],
              }),
            ],
          }),
          i.jsx("p", {
            className: "lg-table-help",
            children: i.jsx(LocalizedText, {
              children:
                "Preview up to 500 rows. Blanks and formulas are preserved. Types report current exceptions; ordinary edits are not coerced. Explicit row appends must match the saved types. Native JSON retains original headers and converted cells, including their notes and formats.",
            }),
          }),
          F && i.jsx("p", { role: "alert", children: F }),
          i.jsx("button", {
            type: "button",
            disabled: readOnly,
            onClick: () => {
              (I(""), j(0), k(null));
              try {
                k({
                  plan: previewColumnTypeChange(
                    workbook,
                    sheetId,
                    tableId,
                    request,
                    {
                      readOnly,
                    },
                  ),
                  context: V,
                });
              } catch (error) {
                I(error.message);
              }
            },
            children: i.jsx(LocalizedText, { children: "Preview column type" }),
          }),
          A &&
            i.jsxs("section", {
              "aria-label": text("Column type preview"),
              className: "lg-table-preview",
              lang: language("Column type preview"),
              "data-ui-label": "Column type preview",
              children: [
                i.jsx("h3", {
                  children: i.jsx(LocalizedText, {
                    children: "Review type changes",
                  }),
                }),
                i.jsxs("p", {
                  children: [
                    A.plan.beforeType,
                    " → ",
                    A.plan.request.dataType,
                    ";",
                    " ",
                    A.plan.changes.length,
                    " ",
                    i.jsx(LocalizedText, { children: "cell changes;" }),
                    " ",
                    A.plan.exceptions.length,
                    " ",
                    i.jsx(LocalizedText, { children: "exceptions." }),
                  ],
                }),
                D &&
                  i.jsx("p", {
                    role: "alert",
                    children: i.jsx(LocalizedText, {
                      children:
                        "The workbook, selection or settings changed. Preview again.",
                    }),
                  }),
                !A.plan.canApply &&
                  i.jsx("p", {
                    role: "alert",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Resolve the exceptions or explicitly choose Keep exceptions unchanged, then preview again.",
                    }),
                  }),
                i.jsx("ol", {
                  start: O * 20 + 1,
                  children: [
                    ...A.plan.changes.map((change) => ({
                      address: change.address,
                      text: `${JSON.stringify(change.before?.raw ?? null)} → ${JSON.stringify(change.after.raw)}`,
                    })),
                    ...A.plan.exceptions.map((exception) => ({
                      address: exception.address,
                      text: `Retained ${JSON.stringify(exception.value)}: ${exception.reason}`,
                    })),
                  ]
                    .slice(O * 20, O * 20 + 20)
                    .map((T) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            i.jsx("strong", { children: T.address }),
                            ": ",
                            T.text,
                          ],
                        },
                        T.address,
                      ),
                    ),
                }),
                i.jsxs("div", {
                  className: "lg-table-inline",
                  children: [
                    i.jsx("button", {
                      type: "button",
                      disabled: !O,
                      onClick: () => j((T) => T - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous type changes",
                      }),
                    }),
                    i.jsxs("span", {
                      children: [
                        i.jsx(LocalizedText, { children: "Page " }),
                        O + 1,
                      ],
                    }),
                    i.jsx("button", {
                      type: "button",
                      disabled:
                        (O + 1) * 20 >=
                        A.plan.changes.length + A.plan.exceptions.length,
                      onClick: () => j((T) => T + 1),
                      children: i.jsx(LocalizedText, {
                        children: "Next type changes",
                      }),
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: readOnly || D || !A.plan.canApply,
                  onClick: () => {
                    I("");
                    try {
                      if (D) throw Error("Preview again.");
                      const T = applyColumnTypePreview(workbook, A.plan, {
                        readOnly,
                      });
                      (T !== workbook &&
                        onWorkbookChange(T, "Apply reviewed table column type"),
                        k(null));
                    } catch (error) {
                      I(error.message);
                    }
                  },
                  children: i.jsx(LocalizedText, {
                    children: "Apply column type",
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  onClick: () => k(null),
                  children: i.jsx(LocalizedText, {
                    children: "Cancel type preview",
                  }),
                }),
              ],
            }),
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(LocalizedText, {
                    children: "Saved original source evidence (",
                  }),
                  _.length,
                  ")",
                ],
              }),
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "Historical addresses and source revisions. Save native JSON to retain these records; XLSX carries current values only.",
                }),
              }),
              _.map((T) =>
                i.jsxs(
                  "details",
                  {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          T.kind,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "· source revision ",
                          }),
                          T.sourceRevision,
                          " · ",
                          T.changes.length,
                          " ",
                          i.jsx(LocalizedText, { children: "cells" }),
                        ],
                      }),
                      i.jsx("pre", { children: JSON.stringify(T, null, 2) }),
                    ],
                  },
                  T.id,
                ),
              ),
            ],
          }),
        ],
      }),
    ],
  });
}
function PivotFieldOptionsPanel({
  fields,
  groups,
  dateGroups,
  setDateGroups,
  calculatedFields,
  setCalculatedFields,
  onAdd,
}) {
  const { text } = useWorkbookMessages();
  return i.jsxs(i.Fragment, {
    children: [
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(LocalizedText, { children: "Date grouping" }),
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Choose year, quarter or month for selected row/column fields. ISO requires YYYY-MM-DD text. Numeric dates require an explicit 1900 or 1904 date system.",
            }),
          }),
          [...new Set(groups)].map((field) => {
            const l = dateGroups.find((dateGroup) => dateGroup.field === field);
            return i.jsxs(
              "div",
              {
                className: "lg-table-inline",
                children: [
                  i.jsxs("label", {
                    children: [
                      field,
                      " ",
                      i.jsx(LocalizedText, { children: "grouping" }),
                      i.jsxs("select", {
                        "aria-label": `Date grouping ${field}`,
                        value: l?.unit ?? "",
                        onChange: (event) =>
                          setDateGroups([
                            ...dateGroups.filter(
                              (dateGroup) => dateGroup.field !== field,
                            ),
                            ...(event.target.value
                              ? [
                                  {
                                    field: field,
                                    unit: event.target.value,
                                    dateSystem: l?.dateSystem ?? "iso",
                                  },
                                ]
                              : []),
                          ]),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "No date grouping",
                          }),
                          i.jsx("option", { value: "year", children: "Year" }),
                          i.jsx("option", {
                            value: "quarter",
                            children: "Quarter",
                          }),
                          i.jsx("option", {
                            value: "month",
                            children: "Month",
                          }),
                        ],
                      }),
                    ],
                  }),
                  l &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Date system for " }),
                        field,
                        i.jsxs("select", {
                          value: l.dateSystem,
                          onChange: (event) =>
                            setDateGroups(
                              dateGroups.map((dateGroup) =>
                                dateGroup.field === field
                                  ? {
                                      ...dateGroup,
                                      dateSystem: event.target.value,
                                    }
                                  : dateGroup,
                              ),
                            ),
                          children: [
                            i.jsx("option", {
                              value: "iso",
                              children: "ISO date text",
                            }),
                            i.jsx("option", {
                              value: "1900",
                              children: "Excel 1900 serial",
                            }),
                            i.jsx("option", {
                              value: "1904",
                              children: "Excel 1904 serial",
                            }),
                          ],
                        }),
                      ],
                    }),
                ],
              },
              field,
            );
          }),
        ],
      }),
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(LocalizedText, {
              children: "Calculated pivot fields",
            }),
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Each calculation uses source-field sums within each group and total. These are aggregate calculations. Use a table calculated column for arithmetic on individual records. Choose sum in Value fields.",
            }),
          }),
          calculatedFields.map((calculatedField, l) => {
            const c = (h) =>
              setCalculatedFields(
                calculatedFields.map((f, p) => (l === p ? { ...f, ...h } : f)),
              );
            return i.jsxs(
              "fieldset",
              {
                className: "lg-table-list",
                children: [
                  i.jsxs("legend", {
                    children: [
                      i.jsx(LocalizedText, { children: "Calculation " }),
                      l + 1,
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Calculated field name ",
                      }),
                      l + 1,
                      i.jsx("input", {
                        value: calculatedField.name,
                        onChange: (event) => c({ name: event.target.value }),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Left source field " }),
                      l + 1,
                      i.jsx("select", {
                        value: calculatedField.field,
                        onChange: (event) => c({ field: event.target.value }),
                        children: fields.map((field) =>
                          i.jsx("option", { children: field }, field),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Calculation operation ",
                      }),
                      l + 1,
                      i.jsx("select", {
                        value: calculatedField.operation,
                        onChange: (event) =>
                          c({ operation: event.target.value }),
                        children: ["add", "subtract", "multiply", "divide"].map(
                          (h) => i.jsx("option", { children: h }, h),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Right operand " }),
                      l + 1,
                      i.jsxs("select", {
                        value: calculatedField.otherField ?? "",
                        onChange: (event) =>
                          setCalculatedFields(
                            calculatedFields.map((f, p) =>
                              p !== l
                                ? f
                                : {
                                    name: f.name,
                                    field: f.field,
                                    operation: f.operation,
                                    ...(event.target.value
                                      ? { otherField: event.target.value }
                                      : { constant: 1 }),
                                  },
                            ),
                          ),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Numeric constant",
                          }),
                          fields.map((field) =>
                            i.jsx("option", { children: field }, field),
                          ),
                        ],
                      }),
                    ],
                  }),
                  !calculatedField.otherField &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Calculation constant ",
                        }),
                        l + 1,
                        i.jsx(LocalizedNumberField, {
                          "aria-label": text("Calculation constant ") + (l + 1),
                          value: Number.isFinite(calculatedField.constant)
                            ? calculatedField.constant
                            : "",
                          onValueChange: (h) =>
                            c({ constant: h === "" ? NaN : Number(h) }),
                        }),
                      ],
                    }),
                  i.jsxs("button", {
                    type: "button",
                    onClick: () =>
                      setCalculatedFields(
                        calculatedFields.filter((h, f) => f !== l),
                      ),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove calculation " }),
                      l + 1,
                    ],
                  }),
                ],
              },
              l,
            );
          }),
          i.jsx("button", {
            type: "button",
            disabled: !fields.length || calculatedFields.length >= 100,
            onClick: () => {
              let u = 1;
              for (
                ;
                fields.includes(`Calculated${u}`) ||
                calculatedFields.some(
                  (calculatedField) =>
                    calculatedField.name === `Calculated${u}`,
                );
              )
                u++;
              const l = `Calculated${u}`;
              (setCalculatedFields([
                ...calculatedFields,
                {
                  name: l,
                  field: fields.at(-1),
                  operation: "multiply",
                  constant: 1,
                },
              ]),
                onAdd(l));
            },
            children: i.jsx(LocalizedText, {
              children: "Add calculated pivot field",
            }),
          }),
        ],
      }),
    ],
  });
}
function PivotStatusBadge({ workbook, pivotId, compact = false }) {
  const n = ne.useMemo(
      () => getPivotRefreshState(workbook, pivotId),
      [workbook, pivotId],
    ),
    definition = n.definition,
    refresh = definition.refresh;
  return i.jsxs("section", {
    "aria-label": `Pivot status ${definition.name}`,
    className: `lg-pivot-status ${n.state}`,
    children: [
      i.jsxs("strong", {
        children: [
          definition.name,
          ":",
          " ",
          n.state === "current"
            ? "Current"
            : n.state === "stale"
              ? "Stale — refresh required"
              : n.state === "unverified"
                ? "Unverified — refresh required"
                : "Source unavailable",
        ],
      }),
      i.jsxs("p", {
        children: [
          refresh?.sourceSheetName ?? definition.source.sheetId,
          "!",
          refresh?.sourceRange ?? definition.source.range,
          " ",
          i.jsx(LocalizedText, { children: "· Source revision " }),
          definition.sourceRevision,
          " ",
          i.jsx(LocalizedText, { children: "· Refreshed " }),
          refresh?.at ?? "not recorded",
        ],
      }),
      n.state !== "current" &&
        i.jsxs("p", {
          children: [
            n.reason,
            " ",
            i.jsx(LocalizedText, {
              children: "Existing totals are retained as historical output.",
            }),
          ],
        }),
      !compact &&
        refresh &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsxs("p", {
              children: [
                refresh.includedRows,
                " ",
                i.jsx(LocalizedText, { children: "included + " }),
                refresh.excludedRows,
                " ",
                i.jsx(LocalizedText, { children: "excluded =" }),
                " ",
                refresh.sourceRows,
                " ",
                i.jsx(LocalizedText, {
                  children:
                    "source records. Blank rows are omitted; hidden rows remain in the population unless explicitly filtered.",
                }),
              ],
            }),
            i.jsx("ul", {
              children: refresh.measures.map((measure, d) =>
                i.jsxs(
                  "li",
                  {
                    children: [
                      measure.name,
                      ": ",
                      JSON.stringify(measure.result),
                      " · ",
                      measure.numericCount,
                      " ",
                      i.jsx(LocalizedText, { children: "numeric / " }),
                      measure.nonblankCount,
                      " ",
                      i.jsx(LocalizedText, { children: "nonblank inputs" }),
                    ],
                  },
                  d,
                ),
              ),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Distinct counts exclude blanks, compare text without case and retain value types. Averages and calculated totals are recomputed from the included source records.",
              }),
            }),
          ],
        }),
    ],
  });
}
function PivotResultsStatusBar({ workbook, sheetId, onOpen, onExplore }) {
  const { text, language } = useWorkbookMessages(),
    s = workbook.sheets.find((sheet) => sheet.id === sheetId)?.pivots ?? [];
  return s.length
    ? i.jsxs("div", {
        className: "lg-pivot-status-bar",
        "aria-label": text("Pivot result freshness"),
        lang: language("Pivot result freshness"),
        "data-ui-label": "Pivot result freshness",
        children: [
          s
            .slice(0, 3)
            .map((u) =>
              i.jsx(
                PivotStatusBadge,
                { workbook, pivotId: u.id, compact: true },
                u.id,
              ),
            ),
          s.length > 3 &&
            i.jsxs("p", {
              children: [
                s.length - 3,
                " ",
                i.jsx(LocalizedText, {
                  children:
                    "additional results — review each in PivotTable tools.",
                }),
              ],
            }),
          i.jsx("button", {
            type: "button",
            onClick: onOpen,
            children: i.jsx(LocalizedText, {
              children: "Review / refresh PivotTable",
            }),
          }),
          onExplore &&
            i.jsx("button", {
              type: "button",
              onClick: onExplore,
              children: i.jsx(LocalizedText, {
                children: "Filter / explore PivotTable",
              }),
            }),
        ],
      })
    : null;
}
const toggleListValue = (e, r) =>
    e.includes(r) ? e.filter((t) => t !== r) : [...e, r],
  formatAggregationLabel = (e) =>
    e === "distinctCount" ? "Distinct count" : e[0].toUpperCase() + e.slice(1);
function tryComputePivotRows(workbook, sheetId, range) {
  try {
    return getPivotSourceFields(workbook, sheetId, range);
  } catch {
    return [];
  }
}
function nextDefaultTableName(workbook) {
  let r = 1;
  for (
    ;
    workbook.sheets.some((sheet) =>
      sheet.tables?.some((table) => table.name.toLowerCase() === `table${r}`),
    );
  )
    r++;
  return `Table${r}`;
}

export {
  PivotStatusBadge,
  formatAggregationLabel,
  nextDefaultTableName,
  PivotFieldOptionsPanel,
  PivotResultsStatusBar,
  toggleListValue,
  tryComputePivotRows,
  DataToolDialogRouter,
  TableColumnTypesPanel,
};
