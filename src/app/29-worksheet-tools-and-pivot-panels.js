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

function DataToolDialogRouter(props) {
  return props.mode === "validation"
    ? i.jsx(DataValidationDialog, { ...props })
    : props.mode === "duplicates"
      ? i.jsx(DuplicateReviewDialog, { ...props })
      : props.mode === "insertRow" || props.mode === "insertColumn"
        ? i.jsx(StructureToolDialog, { ...props })
        : i.jsx(WorksheetToolDialog, { ...props });
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
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    rangeStart = range.split(":")[0],
    [rangeText, setRangeText] = ne.useState(
      mode === "duplicates" && range === rangeStart
        ? formatOccupiedRangeAddress(sheet)
        : range,
    ),
    [ruleKind, setRuleKind] = ne.useState(
      mode === "conditional" ? "greater" : "list",
    ),
    [ruleValue, setRuleValue] = ne.useState(""),
    [ruleSecondValue, setRuleSecondValue] = ne.useState("100"),
    [fill, setFill] = ne.useState("#fff2cc"),
    [color, setColor] = ne.useState("#9c5700"),
    [operator, setOperator] = ne.useState("between"),
    [allowBlank, setAllowBlank] = ne.useState(true),
    [validationMessage, setValidationMessage] = ne.useState(
      "Enter a value that satisfies this rule.",
    ),
    [checked, setChecked] = ne.useState(true),
    [selectedColumns, setSelectedColumns] = ne.useState(null),
    [protectionPassword, setProtectionPassword] = ne.useState(""),
    [count, setCount] = ne.useState(1),
    [remove, setRemove] = ne.useState(false),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [statusMessage, X$1] = ne.useState(""),
    [isSubmitting, setIsSubmitting] = ne.useState(false),
    [W, K] = ne.useState(""),
    dialogRef = ne.useRef(null),
    textCellEntries = Object.entries(sheet.cells)
      .filter(([cellAddress, cell]) => {
        try {
          const rangeAddress = parseRangeAddress(rangeText),
            cellRange = parseRangeAddress(cellAddress);
          return (
            typeof cell.raw == "string" &&
            !cell.raw.startsWith("=") &&
            cellRange.firstRow >= rangeAddress.firstRow &&
            cellRange.lastRow <= rangeAddress.lastRow &&
            cellRange.firstCol >= rangeAddress.firstCol &&
            cellRange.lastCol <= rangeAddress.lastCol
          );
        } catch {
          return false;
        }
      })
      .slice(0, 500),
    [spellingIndex, setSpellingIndex] = ne.useState(0),
    [raw, setRawValue] = ne.useState(textCellEntries[0]?.[1].raw ?? "");
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("input,textarea,select,button")?.focus(),
      () => activeElement?.focus()
    );
  }, []);
  const modeLabels = {
      conditional: "Conditional Formatting",
      validation: "Data Validation",
      duplicates: "Remove Duplicates",
      protection: sheet.protection?.enabled
        ? "Unprotect Sheet"
        : "Protect Sheet",
      merge: "Merge & Center",
      paint: "Format Painter",
      spelling: "Spelling",
      insertRow: "Insert / Delete Rows",
      insertColumn: "Insert / Delete Columns",
    },
    applyWorkbookChange = (nextWorkbook, changeLabel) => {
      (onWorkbookChange(nextWorkbook, changeLabel), onClose());
    };
  async function runToolAction() {
    (setErrorMessage(""), setIsSubmitting(true));
    try {
      if (readOnly && mode !== "spelling")
        throw Error("This workbook is read-only.");
      if (mode === "merge")
        applyWorkbookChange(
          mergeWorksheetCells(workbook, sheetId, rangeText, false),
          "Merge cells",
        );
      else if (mode === "paint")
        applyWorkbookChange(
          paintCellFormatting(workbook, sheetId, rangeStart, W),
          "Paint cell formatting",
        );
      else if (mode === "protection") {
        const protectedWorkbook = await updateSheetProtection(
          workbook,
          sheetId,
          protectionPassword,
          !sheet.protection?.enabled,
        );
        applyWorkbookChange(
          protectedWorkbook,
          sheet.protection?.enabled ? "Unprotect sheet" : "Protect sheet",
        );
      } else if (mode === "conditional") {
        if (
          ["greater", "less", "between"].includes(ruleKind) &&
          (!ruleValue.trim() ||
            !Number.isFinite(Number(ruleValue)) ||
            (ruleKind === "between" &&
              !Number.isFinite(Number(ruleSecondValue))))
        )
          throw Error("Enter finite comparison values.");
        const newConditionalFormat = {
          id: crypto.randomUUID(),
          range: CE(rangeText),
          kind: ruleKind,
          value: ruleValue,
          second: ruleSecondValue,
          fill: fill,
          color: color,
        };
        applyWorkbookChange(
          {
            ...workbook,
            sheets: workbook.sheets.map((sheet) =>
              sheet.id === sheetId
                ? {
                    ...sheet,
                    conditionalFormats: [
                      ...(sheet.conditionalFormats ?? []),
                      newConditionalFormat,
                    ],
                  }
                : sheet,
            ),
          },
          "Add conditional formatting",
        );
      } else if (mode === "validation") {
        const newValidation = {
          id: crypto.randomUUID(),
          range: CE(rangeText),
          kind: ruleKind,
          operator: operator,
          min: Number(ruleValue || 0),
          max: Number(ruleSecondValue),
          list: ruleValue
            .split(",")
            .map((listEntry) => listEntry.trim())
            .filter(Boolean),
          allowBlank: allowBlank,
          message: validationMessage,
        };
        applyWorkbookChange(
          {
            ...workbook,
            sheets: workbook.sheets.map((sheet) =>
              sheet.id === sheetId
                ? {
                    ...sheet,
                    validations: [...(sheet.validations ?? []), newValidation],
                  }
                : sheet,
            ),
          },
          "Add data validation",
        );
      } else if (mode === "duplicates") {
        const rangeAddress = parseRangeAddress(rangeText),
          compareColumns =
            selectedColumns ??
            Array.from(
              { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
              (Je, offset) => rangeAddress.firstCol + offset,
            ),
          removalResult = validateRemoveDuplicatesRequest(
            workbook,
            sheetId,
            rangeText,
            compareColumns,
            checked,
          );
        if (!removalResult.removed) {
          X$1("No duplicate rows found.");
          return;
        }
        applyWorkbookChange(
          removalResult.workbook,
          `Remove ${removalResult.removed} duplicate rows from selection`,
        );
      } else if (mode === "insertRow" || mode === "insertColumn") {
        const rangeAddress = parseRangeAddress(rangeText),
          structuralEdit = {
            axis: mode === "insertRow" ? "row" : "column",
            index:
              mode === "insertRow"
                ? rangeAddress.firstRow
                : rangeAddress.firstCol,
            count: count,
            remove: remove,
          };
        applyWorkbookChange(
          applyStructuralEdit(workbook, sheetId, structuralEdit),
          `${remove ? "Delete" : "Insert"} ${count} ${structuralEdit.axis}${count === 1 ? "" : "s"}`,
        );
      } else if (mode === "spelling") {
        if (readOnly) throw Error("This workbook is read-only.");
        const spellingEntry = textCellEntries[spellingIndex];
        if (!spellingEntry) throw Error("No text cells in the selected range.");
        const cells = {
          ...sheet.cells,
          [spellingEntry[0]]: { ...spellingEntry[1], raw: raw },
        };
        applyWorkbookChange(
          renameTableHeaderField(workbook, {
            ...workbook,
            sheets: workbook.sheets.map((sheet) =>
              sheet.id === sheetId ? { ...sheet, cells: cells } : sheet,
            ),
          }),
          `Spelling correction at ${spellingEntry[0]}`,
        );
      }
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "The operation could not be completed.",
      );
    } finally {
      setIsSubmitting(false);
    }
  }
  let activeRange;
  try {
    activeRange = parseRangeAddress(rangeText);
  } catch {
    activeRange = parseRangeAddress(range);
  }
  const columnIndexes =
    activeRange.lastCol - activeRange.firstCol < 100
      ? Array.from(
          { length: activeRange.lastCol - activeRange.firstCol + 1 },
          (Re, columnOffset) => activeRange.firstCol + columnOffset,
        )
      : [];
  return i.jsx("div", {
    className: "wt-backdrop",
    onPointerDown: (event) => {
      event.target === event.currentTarget && onClose();
    },
    children: i.jsxs("div", {
      className: "wt-dialog",
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": modeLabels[mode] ?? mode,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                "button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled])",
              ) ?? [],
            ),
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey && document.activeElement === firstFocusable
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("span", {
              children: [
                i.jsx(Grid2x2, { size: 19 }),
                i.jsx("h2", { children: modeLabels[mode] ?? mode }),
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
                    value: rangeText,
                    onChange: (event) => {
                      (setRangeText(event.target.value),
                        setSelectedColumns(null));
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
                      i.jsxs("strong", {
                        children: [sheet.name, "!", rangeStart],
                      }),
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
                        applyWorkbookChange(
                          mergeWorksheetCells(
                            workbook,
                            sheetId,
                            rangeText,
                            true,
                          ),
                          "Unmerge cells",
                        );
                      } catch (error) {
                        setErrorMessage(String(error));
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
                        value: ruleKind,
                        onChange: (event) => setRuleKind(event.target.value),
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
                  !["duplicate", "colorScale"].includes(ruleKind) &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Value" }),
                        i.jsx("input", {
                          "aria-label": text("Conditional value"),
                          value: ruleValue,
                          onChange: (event) => setRuleValue(event.target.value),
                          lang: language("Conditional value"),
                          "data-ui-label": "Conditional value",
                        }),
                      ],
                    }),
                  ruleKind === "between" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Maximum" }),
                        i.jsx("input", {
                          "aria-label": text("Conditional maximum"),
                          value: ruleSecondValue,
                          onChange: (event) =>
                            setRuleSecondValue(event.target.value),
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
                            onChange: (event) => setFill(event.target.value),
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
                            onChange: (event) => setColor(event.target.value),
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
                  sheet.conditionalFormats?.map((conditionalFormat) =>
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
                              applyWorkbookChange(
                                {
                                  ...workbook,
                                  sheets: workbook.sheets.map((sheet) =>
                                    sheet.id === sheetId
                                      ? {
                                          ...sheet,
                                          conditionalFormats:
                                            sheet.conditionalFormats?.filter(
                                              (formatEntry) =>
                                                formatEntry.id !==
                                                conditionalFormat.id,
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
                        value: ruleKind,
                        onChange: (event) => setRuleKind(event.target.value),
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
                  ruleKind !== "list" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Condition" }),
                        i.jsxs("select", {
                          "aria-label": text("Validation operator"),
                          value: operator,
                          onChange: (event) => setOperator(event.target.value),
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
                      ruleKind === "list"
                        ? "Values, separated by commas"
                        : "Minimum / comparison value",
                      i.jsx("input", {
                        "aria-label": text("Validation values"),
                        value: ruleValue,
                        onChange: (event) => setRuleValue(event.target.value),
                        placeholder:
                          ruleKind === "list"
                            ? "Pending, Complete, Reviewed"
                            : "0",
                        lang: language("Validation values"),
                        "data-ui-label": "Validation values",
                      }),
                    ],
                  }),
                  ruleKind !== "list" &&
                    operator === "between" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Maximum" }),
                        i.jsx("input", {
                          "aria-label": text("Validation maximum"),
                          value: ruleSecondValue,
                          onChange: (event) =>
                            setRuleSecondValue(event.target.value),
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
                        checked: allowBlank,
                        onChange: (event) =>
                          setAllowBlank(event.target.checked),
                      }),
                      i.jsx(LocalizedText, { children: "Allow blank cells" }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Error message" }),
                      i.jsx("input", {
                        "aria-label": text("Validation message"),
                        value: validationMessage,
                        onChange: (event) =>
                          setValidationMessage(event.target.value),
                        lang: language("Validation message"),
                        "data-ui-label": "Validation message",
                      }),
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
                            ],
                          }),
                          i.jsx("button", {
                            disabled: readOnly,
                            onClick: () =>
                              applyWorkbookChange(
                                {
                                  ...workbook,
                                  sheets: workbook.sheets.map((sheet) =>
                                    sheet.id === sheetId
                                      ? {
                                          ...sheet,
                                          validations:
                                            sheet.validations?.filter(
                                              (validationEntry) =>
                                                validationEntry.id !==
                                                validation.id,
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
                        onChange: (event) => setChecked(event.target.checked),
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
                      columnIndexes.map((columnIndex) =>
                        i.jsxs(
                          "label",
                          {
                            className: "wt-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: (
                                  selectedColumns ?? columnIndexes
                                ).includes(columnIndex),
                                onChange: (event) =>
                                  setSelectedColumns(
                                    event.target.checked
                                      ? [
                                          ...(selectedColumns ?? columnIndexes),
                                          columnIndex,
                                        ]
                                      : (
                                          selectedColumns ?? columnIndexes
                                        ).filter(
                                          (selectedColumn) =>
                                            selectedColumn !== columnIndex,
                                        ),
                                  ),
                              }),
                              columnIndexToLetters(columnIndex),
                              " ·",
                              " ",
                              String(
                                sheet.cells[
                                  `${columnIndexToLetters(columnIndex)}${activeRange.firstRow + 1}`
                                ]?.raw ?? "",
                              ),
                            ],
                          },
                          columnIndex,
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
                      sheet.protection?.enabled
                        ? "Unlock editing on this worksheet."
                        : "Prevent changes to this worksheet until it is unprotected.",
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      sheet.protection?.hash
                        ? "Password"
                        : "Password (optional)",
                      i.jsx("input", {
                        type: "password",
                        "aria-label": text("Sheet password"),
                        value: protectionPassword,
                        onChange: (event) =>
                          setProtectionPassword(event.target.value),
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
                        onChange: (event) =>
                          setRemove(event.target.value === "delete"),
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
                        onValueChange: (countText) =>
                          setCount(Number(countText)),
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
                  textCellEntries.length
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, { children: "Cell" }),
                              i.jsx("select", {
                                "aria-label": text("Spelling cell"),
                                value: spellingIndex,
                                onChange: (event) => {
                                  const selectedIndex = Number(
                                    event.target.value,
                                  );
                                  (setSpellingIndex(selectedIndex),
                                    setRawValue(
                                      String(
                                        textCellEntries[selectedIndex][1].raw,
                                      ),
                                    ));
                                },
                                lang: language("Spelling cell"),
                                "data-ui-label": "Spelling cell",
                                children: textCellEntries.map(
                                  ([entryAddress], optionIndex) =>
                                    i.jsx(
                                      "option",
                                      {
                                        value: optionIndex,
                                        children: entryAddress,
                                      },
                                      entryAddress,
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
                            onChange: (event) =>
                              setRawValue(event.target.value),
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
            statusMessage &&
              i.jsx("p", { role: "status", children: statusMessage }),
            errorMessage &&
              i.jsx("p", {
                className: "wt-error",
                role: "alert",
                children: errorMessage,
              }),
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
              disabled: isSubmitting || readOnly,
              onClick: () => void runToolAction(),
              children: isSubmitting
                ? "Applying…"
                : mode === "spelling"
                  ? "Save correction"
                  : mode === "protection"
                    ? sheet.protection?.enabled
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
    table = workbook.sheets
      .find((sheet) => sheet.id === sheetId)
      .tables.find((table) => table.id === tableId),
    [columnId, setColumnId] = ne.useState(table.columns[0].id),
    [dataType, setDataType] = ne.useState(table.columns[0].dataType ?? "any"),
    [mode, setMode] = ne.useState("convert"),
    [invalid, setInvalid] = ne.useState("reject"),
    [decimal, setDecimal] = ne.useState("."),
    [group, setGroup] = ne.useState(","),
    [columnTypePreview, setColumnTypePreview] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [pageIndex, setPageIndex] = ne.useState(0),
    request = {
      columnId: columnId,
      dataType: dataType,
      mode: mode,
      invalid: invalid,
      decimal: decimal,
      group: group,
    },
    requestKey = JSON.stringify({ context, request: request }),
    isPreviewStale =
      !!columnTypePreview &&
      (columnTypePreview.context !== requestKey ||
        columnTypePreview.plan.snapshot !== JSON.stringify(workbook)),
    currentTypeReport = ne.useMemo(() => {
      try {
        return setTableColumnDataType(workbook, sheetId, tableId, columnId);
      } catch {
        return null;
      }
    }, [workbook, sheetId, tableId, columnId]),
    tableEvidence =
      workbook.tableEvidence?.filter(
        (evidenceRecord) => evidenceRecord.tableId === tableId,
      ) ?? [];
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
                  (setColumnId(event.target.value),
                    setDataType(
                      table.columns.find(
                        (column) => column.id === event.target.value,
                      )?.dataType ?? "any",
                    ),
                    setColumnTypePreview(null));
                },
                lang: language("Typed column"),
                "data-ui-label": "Typed column",
                children: table.columns.map((column) =>
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
              currentTypeReport?.dataType ?? "unknown",
              "; ",
              currentTypeReport?.checked ?? 0,
              " ",
              i.jsx(LocalizedText, { children: "rows checked; " }),
              currentTypeReport?.exceptions.length ?? 0,
              " ",
              i.jsx(LocalizedText, { children: "exceptions" }),
              currentTypeReport?.truncated
                ? " — partial inspection (first 500 rows)"
                : "",
              ".",
            ],
          }),
          !!currentTypeReport?.exceptions.length &&
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(LocalizedText, {
                    children: "Current type exceptions",
                  }),
                }),
                currentTypeReport.exceptions
                  .slice(0, 50)
                  .map((exceptionEntry) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          exceptionEntry.address,
                          ": ",
                          JSON.stringify(exceptionEntry.value),
                        ],
                      },
                      exceptionEntry.address,
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
                onChange: (event) => setDataType(event.target.value),
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
                onChange: (event) => setMode(event.target.value),
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
                      onChange: (event) => setDecimal(event.target.value),
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
                      onChange: (event) => setGroup(event.target.value),
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
                onChange: (event) => setInvalid(event.target.value),
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
          errorMessage && i.jsx("p", { role: "alert", children: errorMessage }),
          i.jsx("button", {
            type: "button",
            disabled: readOnly,
            onClick: () => {
              (setErrorMessage(""),
                setPageIndex(0),
                setColumnTypePreview(null));
              try {
                setColumnTypePreview({
                  plan: previewColumnTypeChange(
                    workbook,
                    sheetId,
                    tableId,
                    request,
                    {
                      readOnly,
                    },
                  ),
                  context: requestKey,
                });
              } catch (error) {
                setErrorMessage(error.message);
              }
            },
            children: i.jsx(LocalizedText, { children: "Preview column type" }),
          }),
          columnTypePreview &&
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
                    columnTypePreview.plan.beforeType,
                    " → ",
                    columnTypePreview.plan.request.dataType,
                    ";",
                    " ",
                    columnTypePreview.plan.changes.length,
                    " ",
                    i.jsx(LocalizedText, { children: "cell changes;" }),
                    " ",
                    columnTypePreview.plan.exceptions.length,
                    " ",
                    i.jsx(LocalizedText, { children: "exceptions." }),
                  ],
                }),
                isPreviewStale &&
                  i.jsx("p", {
                    role: "alert",
                    children: i.jsx(LocalizedText, {
                      children:
                        "The workbook, selection or settings changed. Preview again.",
                    }),
                  }),
                !columnTypePreview.plan.canApply &&
                  i.jsx("p", {
                    role: "alert",
                    children: i.jsx(LocalizedText, {
                      children:
                        "Resolve the exceptions or explicitly choose Keep exceptions unchanged, then preview again.",
                    }),
                  }),
                i.jsx("ol", {
                  start: pageIndex * 20 + 1,
                  children: [
                    ...columnTypePreview.plan.changes.map((change) => ({
                      address: change.address,
                      text: `${JSON.stringify(change.before?.raw ?? null)} → ${JSON.stringify(change.after.raw)}`,
                    })),
                    ...columnTypePreview.plan.exceptions.map((exception) => ({
                      address: exception.address,
                      text: `Retained ${JSON.stringify(exception.value)}: ${exception.reason}`,
                    })),
                  ]
                    .slice(pageIndex * 20, pageIndex * 20 + 20)
                    .map((changeItem) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            i.jsx("strong", { children: changeItem.address }),
                            ": ",
                            changeItem.text,
                          ],
                        },
                        changeItem.address,
                      ),
                    ),
                }),
                i.jsxs("div", {
                  className: "lg-table-inline",
                  children: [
                    i.jsx("button", {
                      type: "button",
                      disabled: !pageIndex,
                      onClick: () =>
                        setPageIndex((currentPage) => currentPage - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous type changes",
                      }),
                    }),
                    i.jsxs("span", {
                      children: [
                        i.jsx(LocalizedText, { children: "Page " }),
                        pageIndex + 1,
                      ],
                    }),
                    i.jsx("button", {
                      type: "button",
                      disabled:
                        (pageIndex + 1) * 20 >=
                        columnTypePreview.plan.changes.length +
                          columnTypePreview.plan.exceptions.length,
                      onClick: () =>
                        setPageIndex((pageNumber) => pageNumber + 1),
                      children: i.jsx(LocalizedText, {
                        children: "Next type changes",
                      }),
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  disabled:
                    readOnly ||
                    isPreviewStale ||
                    !columnTypePreview.plan.canApply,
                  onClick: () => {
                    setErrorMessage("");
                    try {
                      if (isPreviewStale) throw Error("Preview again.");
                      const appliedWorkbook = applyColumnTypePreview(
                        workbook,
                        columnTypePreview.plan,
                        {
                          readOnly,
                        },
                      );
                      (appliedWorkbook !== workbook &&
                        onWorkbookChange(
                          appliedWorkbook,
                          "Apply reviewed table column type",
                        ),
                        setColumnTypePreview(null));
                    } catch (error) {
                      setErrorMessage(error.message);
                    }
                  },
                  children: i.jsx(LocalizedText, {
                    children: "Apply column type",
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  onClick: () => setColumnTypePreview(null),
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
                  tableEvidence.length,
                  ")",
                ],
              }),
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "Historical addresses and source revisions. Save native JSON to retain these records; XLSX carries current values only.",
                }),
              }),
              tableEvidence.map((evidence) =>
                i.jsxs(
                  "details",
                  {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          evidence.kind,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "· source revision ",
                          }),
                          evidence.sourceRevision,
                          " · ",
                          evidence.changes.length,
                          " ",
                          i.jsx(LocalizedText, { children: "cells" }),
                        ],
                      }),
                      i.jsx("pre", {
                        children: JSON.stringify(evidence, null, 2),
                      }),
                    ],
                  },
                  evidence.id,
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
            const dateGroup = dateGroups.find(
              (dateGroup) => dateGroup.field === field,
            );
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
                        value: dateGroup?.unit ?? "",
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
                                    dateSystem: dateGroup?.dateSystem ?? "iso",
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
                  dateGroup &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Date system for " }),
                        field,
                        i.jsxs("select", {
                          value: dateGroup.dateSystem,
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
          calculatedFields.map((calculatedField, calculationIndex) => {
            const updateCalculation = (changes) =>
              setCalculatedFields(
                calculatedFields.map((calculatedField, itemIndex) =>
                  calculationIndex === itemIndex
                    ? { ...calculatedField, ...changes }
                    : calculatedField,
                ),
              );
            return i.jsxs(
              "fieldset",
              {
                className: "lg-table-list",
                children: [
                  i.jsxs("legend", {
                    children: [
                      i.jsx(LocalizedText, { children: "Calculation " }),
                      calculationIndex + 1,
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Calculated field name ",
                      }),
                      calculationIndex + 1,
                      i.jsx("input", {
                        value: calculatedField.name,
                        onChange: (event) =>
                          updateCalculation({ name: event.target.value }),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Left source field " }),
                      calculationIndex + 1,
                      i.jsx("select", {
                        value: calculatedField.field,
                        onChange: (event) =>
                          updateCalculation({ field: event.target.value }),
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
                      calculationIndex + 1,
                      i.jsx("select", {
                        value: calculatedField.operation,
                        onChange: (event) =>
                          updateCalculation({ operation: event.target.value }),
                        children: ["add", "subtract", "multiply", "divide"].map(
                          (operation) =>
                            i.jsx("option", { children: operation }, operation),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Right operand " }),
                      calculationIndex + 1,
                      i.jsxs("select", {
                        value: calculatedField.otherField ?? "",
                        onChange: (event) =>
                          setCalculatedFields(
                            calculatedFields.map(
                              (calculatedField, fieldIndex) =>
                                fieldIndex !== calculationIndex
                                  ? calculatedField
                                  : {
                                      name: calculatedField.name,
                                      field: calculatedField.field,
                                      operation: calculatedField.operation,
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
                        calculationIndex + 1,
                        i.jsx(LocalizedNumberField, {
                          "aria-label":
                            text("Calculation constant ") +
                            (calculationIndex + 1),
                          value: Number.isFinite(calculatedField.constant)
                            ? calculatedField.constant
                            : "",
                          onValueChange: (constantText) =>
                            updateCalculation({
                              constant:
                                constantText === ""
                                  ? NaN
                                  : Number(constantText),
                            }),
                        }),
                      ],
                    }),
                  i.jsxs("button", {
                    type: "button",
                    onClick: () =>
                      setCalculatedFields(
                        calculatedFields.filter(
                          (calculatedField, entryIndex) =>
                            entryIndex !== calculationIndex,
                        ),
                      ),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove calculation " }),
                      calculationIndex + 1,
                    ],
                  }),
                ],
              },
              calculationIndex,
            );
          }),
          i.jsx("button", {
            type: "button",
            disabled: !fields.length || calculatedFields.length >= 100,
            onClick: () => {
              let nameSuffix = 1;
              for (
                ;
                fields.includes(`Calculated${nameSuffix}`) ||
                calculatedFields.some(
                  (calculatedField) =>
                    calculatedField.name === `Calculated${nameSuffix}`,
                );
              )
                nameSuffix++;
              const newCalculatedName = `Calculated${nameSuffix}`;
              (setCalculatedFields([
                ...calculatedFields,
                {
                  name: newCalculatedName,
                  field: fields.at(-1),
                  operation: "multiply",
                  constant: 1,
                },
              ]),
                onAdd(newCalculatedName));
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
  const refreshState = ne.useMemo(
      () => getPivotRefreshState(workbook, pivotId),
      [workbook, pivotId],
    ),
    definition = refreshState.definition,
    refresh = definition.refresh;
  return i.jsxs("section", {
    "aria-label": `Pivot status ${definition.name}`,
    className: `lg-pivot-status ${refreshState.state}`,
    children: [
      i.jsxs("strong", {
        children: [
          definition.name,
          ":",
          " ",
          refreshState.state === "current"
            ? "Current"
            : refreshState.state === "stale"
              ? "Stale — refresh required"
              : refreshState.state === "unverified"
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
      refreshState.state !== "current" &&
        i.jsxs("p", {
          children: [
            refreshState.reason,
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
              children: refresh.measures.map((measure, measureIndex) =>
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
                  measureIndex,
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
    pivots =
      workbook.sheets.find((sheet) => sheet.id === sheetId)?.pivots ?? [];
  return pivots.length
    ? i.jsxs("div", {
        className: "lg-pivot-status-bar",
        "aria-label": text("Pivot result freshness"),
        lang: language("Pivot result freshness"),
        "data-ui-label": "Pivot result freshness",
        children: [
          pivots
            .slice(0, 3)
            .map((pivot) =>
              i.jsx(
                PivotStatusBadge,
                { workbook, pivotId: pivot.id, compact: true },
                pivot.id,
              ),
            ),
          pivots.length > 3 &&
            i.jsxs("p", {
              children: [
                pivots.length - 3,
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
const toggleListValue = (list, value) =>
    list.includes(value)
      ? list.filter((entry) => entry !== value)
      : [...list, value],
  formatAggregationLabel = (aggregation) =>
    aggregation === "distinctCount"
      ? "Distinct count"
      : aggregation[0].toUpperCase() + aggregation.slice(1);
function tryComputePivotRows(workbook, sheetId, range) {
  try {
    return getPivotSourceFields(workbook, sheetId, range);
  } catch {
    return [];
  }
}
function nextDefaultTableName(workbook) {
  let tableNumber = 1;
  for (
    ;
    workbook.sheets.some((sheet) =>
      sheet.tables?.some(
        (table) => table.name.toLowerCase() === `table${tableNumber}`,
      ),
    );
  )
    tableNumber++;
  return `Table${tableNumber}`;
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
