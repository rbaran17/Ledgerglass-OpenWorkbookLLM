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
    [editingRuleId, setEditingRuleId] = ne.useState(),
    [ruleRange, setRuleRange] = ne.useState(range),
    [ruleKind, setRuleKind] = ne.useState("list"),
    [operator, setOperator] = ne.useState("between"),
    [listText, setListText] = ne.useState("Pending,Complete,Reviewed"),
    [minValueText, onValueChange] = ne.useState("0"),
    [maxValueText, setMaxValueText] = ne.useState("100"),
    [formula, setFormula] = ne.useState(`=LEN(${range.split(":")[0]})>0`),
    [isBlankAllowed, setIsBlankAllowed] = ne.useState(true),
    [isDropdownShown, setIsDropdownShown] = ne.useState(true),
    [isInputMessageShown, setIsInputMessageShown] = ne.useState(true),
    [inputTitle, setInputTitle] = ne.useState(""),
    [inputMessage, setInputMessage] = ne.useState(""),
    [errorStyle, setErrorStyle] = ne.useState("stop"),
    [errorTitle, setErrorTitle] = ne.useState("Invalid value"),
    [errorMessage, setErrorMessage] = ne.useState(
      "Enter a value that satisfies this rule.",
    ),
    [X$1, setValidationError] = ne.useState(""),
    [inspection, setInspection] = ne.useState(),
    dialogRef = ne.useRef(null);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button,input")?.focus(),
      () => activeElement?.focus()
    );
  }, []);
  const disabled = readOnly || !!sheet.protection?.enabled,
    runAndReportError = (action) => {
      setValidationError("");
      try {
        action();
      } catch (error) {
        setValidationError(
          error instanceof Error
            ? error.message
            : "Validation could not be completed.",
        );
      }
    },
    onClick = () =>
      runAndReportError(() => {
        if (disabled) throw Error("This worksheet is read-only or protected.");
        if (
          (assertSheetUnprotected(sheet),
          !["list", "custom"].includes(ruleKind) &&
            (!minValueText.trim() || !maxValueText.trim()))
        )
          throw Error("Enter finite validation bounds.");
        const rule = {
            id: editingRuleId ?? crypto.randomUUID(),
            range: ruleRange,
            kind: ruleKind,
            operator: operator,
            min: Number(minValueText),
            max: Number(maxValueText),
            list:
              ruleKind === "list"
                ? listText.trim().startsWith("[")
                  ? JSON.parse(listText)
                  : listText.split(",").map((listItem) => listItem.trim())
                : [],
            allowBlank: isBlankAllowed,
            message: errorMessage,
            errorStyle: errorStyle,
            errorTitle: errorTitle,
            showDropdown: isDropdownShown,
            showInputMessage: isInputMessageShown,
            inputTitle: inputTitle,
            inputMessage: inputMessage,
            ...(ruleKind === "custom" ? { formula: formula } : {}),
          },
          dataValidationRules = normalizeDataValidationRules(
            editingRuleId
              ? (sheet.validations ?? []).map((existingRule) =>
                  existingRule.id === editingRuleId ? rule : existingRule,
                )
              : [...(sheet.validations ?? []), rule],
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
          editingRuleId ? "Update data validation" : "Add data validation",
        ),
          onClose());
      }),
    loadRuleIntoForm = (ruleToLoad) => {
      (setEditingRuleId(ruleToLoad.id),
        setRuleRange(ruleToLoad.range),
        setRuleKind(ruleToLoad.kind),
        setOperator(ruleToLoad.operator),
        setListText(JSON.stringify(ruleToLoad.list)),
        onValueChange(String(ruleToLoad.min)),
        setMaxValueText(String(ruleToLoad.max)),
        setFormula(ruleToLoad.formula ?? "=TRUE()"),
        setIsBlankAllowed(ruleToLoad.allowBlank),
        setIsDropdownShown(ruleToLoad.showDropdown !== false),
        setIsInputMessageShown(ruleToLoad.showInputMessage !== false),
        setInputTitle(ruleToLoad.inputTitle ?? ""),
        setInputMessage(ruleToLoad.inputMessage ?? ""),
        setErrorStyle(ruleToLoad.errorStyle ?? "stop"),
        setErrorTitle(ruleToLoad.errorTitle ?? "Invalid value"),
        setErrorMessage(ruleToLoad.message));
    };
  return i.jsx("div", {
    className: "wt-backdrop",
    onKeyDown: (event) => {
      if (
        (event.key === "Escape" && (event.stopPropagation(), onClose()),
        event.key === "Tab")
      ) {
        const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled)",
            ) ?? []),
          ],
          firstElement = focusableElements[0],
          lastElement = focusableElements.at(-1);
        event.shiftKey && document.activeElement === firstElement
          ? (event.preventDefault(), lastElement == null || lastElement.focus())
          : !event.shiftKey &&
            document.activeElement === lastElement &&
            (event.preventDefault(),
            firstElement == null || firstElement.focus());
      }
    },
    children: i.jsxs("div", {
      className: "wt-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": text("Data Validation"),
      ref: dialogRef,
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
                runAndReportError(() =>
                  setInspection({
                    value: summarizeValidationInspection(workbook, sheetId),
                    snapshot: JSON.stringify(workbook),
                  }),
                ),
              children: i.jsx(LocalizedText, {
                children: "Inspect existing values",
              }),
            }),
            inspection &&
              i.jsxs("section", {
                "aria-label": text("Existing validation findings"),
                lang: language("Existing validation findings"),
                "data-ui-label": "Existing validation findings",
                children: [
                  i.jsx("button", {
                    disabled: inspection.snapshot !== JSON.stringify(workbook),
                    onClick: () =>
                      runAndReportError(() => {
                        const validationInspection =
                            resolveValidationInspection(
                              workbook,
                              inspection.value,
                            ),
                          downloadUrl = URL.createObjectURL(
                            new Blob([validationInspection], {
                              type: "application/json",
                            }),
                          ),
                          downloadLink = document.createElement("a");
                        ((downloadLink.href = downloadUrl),
                          (downloadLink.download = "Validation-findings.json"),
                          downloadLink.click(),
                          setTimeout(
                            () => URL.revokeObjectURL(downloadUrl),
                            1e3,
                          ));
                      }),
                    children: i.jsx(LocalizedText, {
                      children: "Download validation findings",
                    }),
                  }),
                  i.jsx("p", {
                    role: "status",
                    children:
                      inspection.snapshot === JSON.stringify(workbook)
                        ? inspection.value.message
                        : "Workbook changed. Inspect existing values again.",
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Inspection recalculates locally, including in Manual mode; it does not publish current calculation results or change source data.",
                    }),
                  }),
                  inspection.snapshot === JSON.stringify(workbook) &&
                    i.jsx("div", {
                      style: { maxHeight: 200, overflow: "auto" },
                      children: inspection.value.issues.map(
                        (issue, issueIndex) =>
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
                                  children: [
                                    issue.sheetName,
                                    "!",
                                    issue.address,
                                  ],
                                }),
                                " ",
                                issue.style,
                                ": ",
                                issue.message,
                                " ",
                                issue.reason,
                              ],
                            },
                            issueIndex,
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
                          onClick: () => loadRuleIntoForm(validation),
                          children: [
                            i.jsx(LocalizedText, { children: "Edit " }),
                            validation.range,
                          ],
                        }),
                        i.jsxs("button", {
                          disabled: disabled,
                          onClick: () =>
                            runAndReportError(() => {
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
                                                (otherValidation) =>
                                                  otherValidation.id !==
                                                  validation.id,
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
                i.jsx("legend", {
                  children: editingRuleId ? "Edit rule" : "New rule",
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Apply to" }),
                    i.jsx("input", {
                      "aria-label": text("Validation range"),
                      value: ruleRange,
                      onChange: (event) => setRuleRange(event.target.value),
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
                          value: ruleKind,
                          onChange: (event) => setRuleKind(event.target.value),
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
                          ].map(([typeValue, typeLabel]) =>
                            i.jsx(
                              "option",
                              { value: typeValue, children: typeLabel },
                              typeValue,
                            ),
                          ),
                        }),
                      ],
                    }),
                    !["list", "custom"].includes(ruleKind) &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Condition" }),
                          i.jsx("select", {
                            "aria-label": text("Validation operator"),
                            value: operator,
                            onChange: (event) =>
                              setOperator(event.target.value),
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
                            ].map(([operatorValue, operatorLabel]) =>
                              i.jsx(
                                "option",
                                {
                                  value: operatorValue,
                                  children: operatorLabel,
                                },
                                operatorValue,
                              ),
                            ),
                          }),
                        ],
                      }),
                  ],
                }),
                ruleKind === "list"
                  ? i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children:
                            "Values, separated by commas (or a JSON array for exact text)",
                        }),
                        i.jsx("input", {
                          "aria-label": text("Validation values"),
                          value: listText,
                          onChange: (event) => setListText(event.target.value),
                          lang: language("Validation values"),
                          "data-ui-label": "Validation values",
                        }),
                      ],
                    })
                  : ruleKind === "custom"
                    ? i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Formula relative to the top-left cell",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Validation formula"),
                            value: formula,
                            onChange: (event) => setFormula(event.target.value),
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
                                value: minValueText,
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
                                  value: maxValueText,
                                  onValueChange: setMaxValueText,
                                  lang: language("Validation maximum"),
                                  "data-ui-label": "Validation maximum",
                                }),
                              ],
                            }),
                        ],
                      }),
                (ruleKind === "date" || ruleKind === "time") &&
                  i.jsx("p", {
                    children:
                      ruleKind === "date"
                        ? "Bounds use Gregorian/1900 date serials. Entered text follows the workbook's explicit MDY/DMY policy."
                        : "Bounds and numeric inputs use day fractions: 0.5 is noon. Time text such as 09:30 is supported.",
                  }),
                i.jsxs("label", {
                  className: "wt-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: isBlankAllowed,
                      onChange: (event) =>
                        setIsBlankAllowed(event.target.checked),
                    }),
                    i.jsx(LocalizedText, { children: "Allow blank cells" }),
                  ],
                }),
                ruleKind === "list" &&
                  i.jsxs("label", {
                    className: "wt-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: isDropdownShown,
                        onChange: (event) =>
                          setIsDropdownShown(event.target.checked),
                      }),
                      i.jsx(LocalizedText, { children: "Show dropdown" }),
                    ],
                  }),
                i.jsxs("label", {
                  className: "wt-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: isInputMessageShown,
                      onChange: (event) =>
                        setIsInputMessageShown(event.target.checked),
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
                      onChange: (event) => setInputTitle(event.target.value),
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
                      onChange: (event) => setInputMessage(event.target.value),
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
                      onChange: (event) => setErrorStyle(event.target.value),
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
                      onChange: (event) => setErrorTitle(event.target.value),
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
                      value: errorMessage,
                      onChange: (event) => setErrorMessage(event.target.value),
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
              children: editingRuleId ? "Save rule" : "Apply",
            }),
          ],
        }),
      ],
    }),
  });
}
function shiftRangeForAxisEdit(rangeText, axisEdit, rangeLabel, n = false) {
  const rangeAddress = parseRangeAddress(rangeText),
    axis = axisEdit.axis;
  return (
    resolveMovedRange(
      axis === "row" ? rangeAddress.firstRow : rangeAddress.firstCol,
      axis === "row" ? rangeAddress.lastRow : rangeAddress.lastCol,
      axisEdit,
      rangeLabel,
      n,
    ),
    assertRangeWithinSheetBounds(rangeText, axisEdit)
  );
}
function assertMoveKeepsTablesIntact(sourceWorkbook, sheetId, axisEdit) {
  const moveEdit = axisEdit,
    axis = axisEdit.axis,
    workbook = structuredClone(sourceWorkbook);
  for (const sheet of workbook.sheets) {
    for (const table of sheet.tables ?? []) {
      if (table.columns.some((column) => column.calculatedFormula))
        throw Error(
          "Clear saved calculated-column rules before moving worksheet structures, keeping their cell formulas; then reapply the rules.",
        );
      if (sheet.id !== sheetId) continue;
      const rangeAddress = parseRangeAddress(table.range),
        movedRangeText = shiftRangeForAxisEdit(
          table.range,
          axisEdit,
          `table ${table.name}`,
        ),
        movedRangeAddress = parseRangeAddress(movedRangeText);
      if (axis === "row") {
        if (
          mapIndexAfterMove(rangeAddress.firstRow, moveEdit) !==
            movedRangeAddress.firstRow ||
          (table.showTotals &&
            mapIndexAfterMove(rangeAddress.lastRow, moveEdit) !==
              movedRangeAddress.lastRow)
        )
          throw Error(
            "This move would displace a table header or totals row. Move the complete table or only its data rows.",
          );
      } else
        table.columns = table.columns
          .map((column, columnOffset) => ({
            column: column,
            index: mapIndexAfterMove(
              rangeAddress.firstCol + columnOffset,
              moveEdit,
            ),
          }))
          .sort((leftEntry, rightEntry) => leftEntry.index - rightEntry.index)
          .map((entry) => entry.column);
      table.range = movedRangeText;
    }
    for (const analysis of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
      if (
        (sheet.id === sheetId &&
          (analysis.outputRange = shiftRangeForAxisEdit(
            analysis.outputRange,
            axisEdit,
            "generated analysis output",
            true,
          )),
        analysis.source.sheetId === sheetId)
      ) {
        const rangeAddress = parseRangeAddress(analysis.source.range),
          movedSourceRange = shiftRangeForAxisEdit(
            analysis.source.range,
            axisEdit,
            "analysis source",
          );
        if (
          axis === "row" &&
          mapIndexAfterMove(rangeAddress.firstRow, moveEdit) !==
            parseRangeAddress(movedSourceRange).firstRow
        )
          throw Error(
            "This move would displace an analysis source header. Move the whole source or only its data rows.",
          );
        analysis.source.range = movedSourceRange;
      }
  }
  return workbook;
}
function collectAxisStructureConflicts(sheet, axisEdit) {
  const axis = axisEdit.axis,
    moveEdit = axisEdit,
    hiddenKey = axis === "row" ? "hiddenRows" : "hiddenCols",
    axisCount = axis === "row" ? sheet.rowCount : sheet.colCount,
    movedHiddenIntervals = (sheet[hiddenKey] ?? []).flatMap((hiddenInterval) =>
      mapRangeAfterMove(hiddenInterval.start, hiddenInterval.end, moveEdit),
    ),
    outlineGroups = sheet.outlineGroups?.map((outlineGroup) =>
      outlineGroup.axis !== axis
        ? outlineGroup
        : {
            ...outlineGroup,
            ...resolveMovedRange(
              outlineGroup.start,
              outlineGroup.end,
              moveEdit,
              "outline group",
              true,
            ),
          },
    );
  return {
    [hiddenKey]: normalizeHiddenIntervals(movedHiddenIntervals, axisCount),
    ...(outlineGroups ? { outlineGroups: outlineGroups } : {}),
  };
}
function shiftPrintLayoutForAxis(sheet, axisEdit) {
  if (!sheet.pageLayout) return;
  const nextPageLayout = { ...sheet.pageLayout },
    axis = axisEdit.axis,
    moveEdit = axisEdit,
    repeatKey = axis === "row" ? "repeatRows" : "repeatColumns",
    breakKey = axis === "row" ? "rowBreaks" : "columnBreaks",
    printTitles = parsePrintTitles(nextPageLayout[repeatKey], axis, sheet);
  if (printTitles) {
    const movedRange = resolveMovedRange(
      printTitles.start,
      printTitles.end,
      moveEdit,
      "print titles",
      true,
    );
    nextPageLayout[repeatKey] =
      axis === "row"
        ? `${movedRange.start + 1}:${movedRange.end + 1}`
        : `${columnIndexToLetters(movedRange.start)}:${columnIndexToLetters(movedRange.end)}`;
  }
  return (
    (nextPageLayout[breakKey] &&= [
      ...new Set(
        nextPageLayout[breakKey].map((breakIndex) =>
          mapIndexAfterMove(breakIndex, moveEdit),
        ),
      ),
    ].sort((leftIndex, rightIndex) => leftIndex - rightIndex)),
    nextPageLayout
  );
}
function remapAxisIndex(axisIndex, axisEdit) {
  return axisEdit.moveBefore !== undefined
    ? mapIndexAfterMove(axisIndex, axisEdit)
    : axisEdit.remove &&
        axisIndex >= axisEdit.index &&
        axisIndex < axisEdit.index + axisEdit.count
      ? null
      : axisIndex >= axisEdit.index + (axisEdit.remove ? axisEdit.count : 0)
        ? axisIndex + (axisEdit.remove ? -axisEdit.count : axisEdit.count)
        : axisIndex;
}
function assertInsertionWithinSheet(cellAddress, axisEdit) {
  const cellPosition = _t(cellAddress),
    newIndex = remapAxisIndex(
      axisEdit.axis === "row" ? cellPosition.row : cellPosition.col,
      axisEdit,
    );
  if (newIndex === null) return null;
  if (
    newIndex < 0 ||
    newIndex >=
      (axisEdit.axis === "row"
        ? workbookLimits.maxRows
        : workbookLimits.maxCols)
  )
    throw Error(
      "Insertion would move existing content beyond the Excel worksheet boundary.",
    );
  const addressMatch = /^(\$?)([A-Z]+)(\$?)(\d+)$/i.exec(cellAddress);
  return axisEdit.axis === "row"
    ? `${addressMatch[1]}${addressMatch[2]}${addressMatch[3]}${newIndex + 1}`
    : `${addressMatch[1]}${columnIndexToLetters(newIndex)}${addressMatch[3]}${addressMatch[4]}`;
}
function assertRangeWithinSheetBounds(rangeText, axisEdit) {
  const rangeAddress = parseRangeAddress(rangeText);
  let startIndex =
      axisEdit.axis === "row" ? rangeAddress.firstRow : rangeAddress.firstCol,
    endIndex =
      axisEdit.axis === "row" ? rangeAddress.lastRow : rangeAddress.lastCol;
  if (axisEdit.moveBefore !== undefined) {
    const movedRange = resolveMovedRange(
      startIndex,
      endIndex,
      axisEdit,
      `reference ${rangeText}`,
    );
    ((startIndex = movedRange.start), (endIndex = movedRange.end));
  } else if (axisEdit.remove) {
    const removedLastIndex = axisEdit.index + axisEdit.count - 1;
    if (startIndex >= axisEdit.index && endIndex <= removedLastIndex)
      return null;
    ((startIndex =
      startIndex >= axisEdit.index && startIndex <= removedLastIndex
        ? axisEdit.index
        : remapAxisIndex(startIndex, axisEdit)),
      (endIndex =
        endIndex >= axisEdit.index && endIndex <= removedLastIndex
          ? axisEdit.index - 1
          : remapAxisIndex(endIndex, axisEdit)));
  } else
    ((startIndex = remapAxisIndex(startIndex, axisEdit)),
      (endIndex = remapAxisIndex(endIndex, axisEdit)));
  if (
    endIndex >=
    (axisEdit.axis === "row" ? workbookLimits.maxRows : workbookLimits.maxCols)
  )
    throw Error("The range would exceed the Excel worksheet boundary.");
  const startAddress =
      axisEdit.axis === "row"
        ? `${columnIndexToLetters(rangeAddress.firstCol)}${startIndex + 1}`
        : `${columnIndexToLetters(startIndex)}${rangeAddress.firstRow + 1}`,
    endAddress =
      axisEdit.axis === "row"
        ? `${columnIndexToLetters(rangeAddress.lastCol)}${endIndex + 1}`
        : `${columnIndexToLetters(endIndex)}${rangeAddress.lastRow + 1}`;
  return startAddress === endAddress
    ? startAddress
    : `${startAddress}:${endAddress}`;
}
function rewriteFormulaForStructuralEdit(
  formulaText,
  currentSheetName,
  targetSheetName,
  structuralEdit,
) {
  if (!formulaText.startsWith("=")) return formulaText;
  const sheetNamePattern = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\\][A-Za-z0-9_.\\]*)`,
    cellReferencePattern = new RegExp(
      `^(${sheetNamePattern}!)?(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?::\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?[1-9]\\d*:\\$?[1-9]\\d*)`,
    ),
    threeDReferencePattern = new RegExp(
      `^${sheetNamePattern}:${sheetNamePattern}!`,
    );
  let rewrittenFormula = "",
    charIndex = 0;
  for (; charIndex < formulaText.length;) {
    const currentChar = formulaText[charIndex],
      remainingText = formulaText.slice(charIndex);
    if (currentChar === '"') {
      const quoteStart = charIndex++;
      for (; charIndex < formulaText.length;)
        if (formulaText[charIndex++] === '"')
          if (formulaText[charIndex] === '"') charIndex++;
          else break;
      rewrittenFormula += formulaText.slice(quoteStart, charIndex);
      continue;
    }
    const isReferenceStart =
      charIndex === 0 || !/[\w.$]/.test(formulaText[charIndex - 1]);
    if (isReferenceStart && threeDReferencePattern.test(remainingText))
      throw Error(
        "Structural edits cannot safely update 3-D worksheet references. Replace the 3-D reference with explicit worksheet references first.",
      );
    if (currentChar === "[") {
      const bracketStart = charIndex++;
      let bracketDepth = 1;
      for (; charIndex < formulaText.length && bracketDepth;)
        (formulaText[charIndex] === "[" && bracketDepth++,
          formulaText[charIndex] === "]" && bracketDepth--,
          charIndex++);
      if (isReferenceStart && /^[^+*/(),]*!/.test(formulaText.slice(charIndex)))
        throw Error(
          "Structural edits cannot safely update external workbook references.",
        );
      rewrittenFormula += formulaText.slice(bracketStart, charIndex);
      continue;
    }
    const referenceMatch = isReferenceStart
      ? cellReferencePattern.exec(remainingText)
      : null;
    if (
      referenceMatch &&
      !/[\w.!:]/.test(
        formulaText[charIndex + referenceMatch[0].length] ?? "",
      ) &&
      !/^\s*\(/.test(formulaText.slice(charIndex + referenceMatch[0].length))
    ) {
      const sheetPrefix = referenceMatch[1] ?? "",
        referencedSheetName = sheetPrefix
          ? sheetPrefix.slice(0, -1).replace(/^'|'$/g, "").replace(/''/g, "'")
          : currentSheetName;
      if (referencedSheetName.includes(":"))
        throw Error(
          "Structural edits cannot safely update 3-D worksheet references.",
        );
      if (referencedSheetName.includes("["))
        throw Error(
          "Structural edits cannot safely update external workbook references.",
        );
      let rewrittenReference = referenceMatch[0];
      if (referencedSheetName.toLowerCase() === targetSheetName.toLowerCase()) {
        const referenceAddress = referenceMatch[2],
          isWholeColumnRange = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$/.test(
            referenceAddress,
          ),
          isWholeRowRange = /^\$?[1-9]\d*:\$?[1-9]\d*$/.test(referenceAddress);
        try {
          if (isWholeColumnRange || isWholeRowRange) {
            if (
              (isWholeColumnRange && structuralEdit.axis === "column") ||
              (isWholeRowRange && structuralEdit.axis === "row")
            ) {
              const [startAxisText, endAxisText] = referenceAddress.split(":"),
                startAxisToken = startAxisText.replace(/\$/g, ""),
                endAxisToken = endAxisText.replace(/\$/g, ""),
                boundedAxisRange = assertRangeWithinSheetBounds(
                  isWholeColumnRange
                    ? `${startAxisToken}1:${endAxisToken}1`
                    : `A${startAxisToken}:A${endAxisToken}`,
                  structuralEdit,
                );
              if (!boundedAxisRange) rewrittenReference = sheetPrefix + "#REF!";
              else {
                const rangeAddress = parseRangeAddress(boundedAxisRange),
                  isReversedRange = isWholeColumnRange
                    ? _t(startAxisToken + "1").col > _t(endAxisToken + "1").col
                    : Number(startAxisToken) > Number(endAxisToken),
                  firstBoundText = isWholeColumnRange
                    ? columnIndexToLetters(rangeAddress.firstCol)
                    : String(rangeAddress.firstRow + 1),
                  lastBoundText = isWholeColumnRange
                    ? columnIndexToLetters(rangeAddress.lastCol)
                    : String(rangeAddress.lastRow + 1);
                rewrittenReference =
                  sheetPrefix +
                  (startAxisText.startsWith("$") ? "$" : "") +
                  (isReversedRange ? lastBoundText : firstBoundText) +
                  ":" +
                  (endAxisText.startsWith("$") ? "$" : "") +
                  (isReversedRange ? firstBoundText : lastBoundText);
              }
            }
          } else if (referenceAddress.includes(":")) {
            const [startCellText, endCellText] = referenceAddress.split(":"),
              boundedCellRange = assertRangeWithinSheetBounds(
                referenceAddress,
                structuralEdit,
              );
            if (!boundedCellRange) rewrittenReference = sheetPrefix + "#REF!";
            else {
              const rangeAddress = parseRangeAddress(boundedCellRange),
                startCellPosition = _t(startCellText),
                endCellPosition = _t(endCellText),
                startCellMatch = /^(\$?)[A-Z]+(\$?)\d+$/i.exec(startCellText),
                endCellMatch = /^(\$?)[A-Z]+(\$?)\d+$/i.exec(endCellText),
                formatEdgeCell = (isStartEndpoint, cellMatch) => {
                  const edgeColumnIndex =
                      startCellPosition.col <= endCellPosition.col ===
                      isStartEndpoint
                        ? rangeAddress.firstCol
                        : rangeAddress.lastCol,
                    edgeRowIndex =
                      startCellPosition.row <= endCellPosition.row ===
                      isStartEndpoint
                        ? rangeAddress.firstRow
                        : rangeAddress.lastRow;
                  return `${cellMatch[1]}${columnIndexToLetters(edgeColumnIndex)}${cellMatch[2]}${edgeRowIndex + 1}`;
                };
              rewrittenReference =
                sheetPrefix +
                formatEdgeCell(true, startCellMatch) +
                ":" +
                formatEdgeCell(false, endCellMatch);
            }
          } else
            rewrittenReference =
              sheetPrefix +
              (assertInsertionWithinSheet(referenceAddress, structuralEdit) ??
                "#REF!");
        } catch (error) {
          if (structuralEdit.moveBefore !== undefined) throw error;
          rewrittenReference = sheetPrefix + "#REF!";
        }
      }
      ((rewrittenFormula += rewrittenReference),
        (charIndex += referenceMatch[0].length));
      continue;
    }
    ((rewrittenFormula += currentChar), charIndex++);
  }
  return rewrittenFormula;
}
function applyStructuralEdit(workbook, sheetId, structuralEdit) {
  if (
    structuralEdit.permutation !== undefined &&
    structuralEdit.moveBefore === undefined
  )
    throw Error(
      "A structural permutation requires an explicit move operation.",
    );
  if (
    !["row", "column"].includes(structuralEdit.axis) ||
    !Number.isInteger(structuralEdit.index) ||
    structuralEdit.index < 0 ||
    !Number.isInteger(structuralEdit.count) ||
    structuralEdit.count < 1 ||
    structuralEdit.count > 1e4 ||
    structuralEdit.index + structuralEdit.count >
      (structuralEdit.axis === "row"
        ? workbookLimits.maxRows
        : workbookLimits.maxCols)
  )
    throw Error("Choose a valid position and 1–10,000 rows or columns.");
  if (structuralEdit.moveBefore !== undefined) {
    if (structuralEdit.remove)
      throw Error(
        "A structural operation cannot delete and move at the same time.",
      );
    validateRowColumnMove(
      structuralEdit,
      structuralEdit.axis === "row"
        ? workbookLimits.maxRows
        : workbookLimits.maxCols,
    );
  }
  const targetSheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!targetSheet) throw Error("The worksheet was not found.");
  assertSheetUnprotected(targetSheet);
  const checkedWorkbook =
    structuralEdit.moveBefore !== undefined
      ? assertMoveKeepsTablesIntact(workbook, sheetId, structuralEdit)
      : checkStructuralEditImpact(
          workbook,
          sheetId,
          structuralEdit.axis,
          structuralEdit.index,
          structuralEdit.count,
          structuralEdit.remove,
        );
  let verifiedCellCount = 0;
  const sheets = checkedWorkbook.sheets.map((sheet) => {
      const cells = {};
      for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
        const remappedAddress =
          sheet.id === sheetId
            ? assertInsertionWithinSheet(cellAddress, structuralEdit)
            : cellAddress;
        if (!remappedAddress) continue;
        const raw =
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          !cell.literal
            ? rewriteFormulaForStructuralEdit(
                cell.raw,
                sheet.name,
                targetSheet.name,
                structuralEdit,
              )
            : cell.raw;
        (raw !== cell.raw && assertSheetUnprotected(sheet),
          (cells[remappedAddress] = { ...cell, raw: raw }));
      }
      const rewriteRule = (rule) => {
          if (!["formula", "custom"].includes(rule.kind) || !rule.formula)
            return rule;
          const ruleLabel =
            rule.kind === "custom" ? "validation" : "conditional";
          if (
            [
              rule.formula,
              ...(workbook.definedNames ?? []).map(
                (nameEntry) => nameEntry.refersTo,
              ),
            ].some((formulaCandidate) =>
              /\b(?:INDIRECT|OFFSET)\s*\(/i.test(
                formulaCandidate.replace(/"(?:[^"]|"")*"/g, '""'),
              ),
            )
          )
            throw Error(
              `Structural edits cannot repair dynamic ${ruleLabel}-rule references. Remove the formula rule first.`,
            );
          if (sheet.id === sheetId && structuralEdit.remove) {
            const anchorCell = rule.range.split(":")[0];
            if (assertInsertionWithinSheet(anchorCell, structuralEdit) === null)
              throw Error(
                `This edit removes a ${ruleLabel} formula's anchor. Relocate or remove the rule first.`,
              );
          }
          const formula = rewriteFormulaForStructuralEdit(
              rule.formula,
              sheet.name,
              targetSheet.name,
              structuralEdit,
            ),
            rowShift = structuralEdit.axis === "row" ? 1 : 0,
            columnShift = structuralEdit.axis === "column" ? 1 : 0,
            rangeAddress = parseRangeAddress(rule.range);
          if (
            shiftFormulaReferences(rule.formula, rowShift, columnShift) !==
              rule.formula &&
            (structuralEdit.axis === "row"
              ? rangeAddress.lastRow > rangeAddress.firstRow
              : rangeAddress.lastCol > rangeAddress.firstCol)
          ) {
            const rangeCellCount =
              (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
              (rangeAddress.lastCol - rangeAddress.firstCol + 1);
            if (verifiedCellCount + rangeCellCount > 5e3)
              throw Error(
                `Structural edits can verify at most 5,000 applied cells for relative ${ruleLabel} formulas. Narrow or remove the rules first.`,
              );
            verifiedCellCount += rangeCellCount;
            const boundedRuleRange =
              sheet.id === sheetId
                ? assertRangeWithinSheetBounds(rule.range, structuralEdit)
                : rule.range;
            if (!boundedRuleRange)
              throw Error(
                `This edit removes a ${ruleLabel} formula range. Remove the rule first.`,
              );
            const boundedRuleAddress = parseRangeAddress(boundedRuleRange);
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
                const cellReference = `${columnIndexToLetters(firstCol)}${firstRow + 1}`,
                  remappedCellReference =
                    sheet.id === sheetId
                      ? assertInsertionWithinSheet(
                          cellReference,
                          structuralEdit,
                        )
                      : cellReference;
                if (!remappedCellReference) continue;
                const remappedCellPosition = _t(remappedCellReference),
                  rewrittenCellFormula = rewriteFormulaForStructuralEdit(
                    shiftFormulaReferences(
                      rule.formula,
                      firstRow - rangeAddress.firstRow,
                      firstCol - rangeAddress.firstCol,
                    ),
                    sheet.name,
                    targetSheet.name,
                    structuralEdit,
                  ),
                  expectedCellFormula = shiftFormulaReferences(
                    formula,
                    remappedCellPosition.row - boundedRuleAddress.firstRow,
                    remappedCellPosition.col - boundedRuleAddress.firstCol,
                  );
                if (rewrittenCellFormula !== expectedCellFormula)
                  throw Error(
                    `This structural edit requires different relative ${ruleLabel} formulas within one rule. Split, relocate, or remove the rule first.`,
                  );
              }
          }
          return (
            formula !== rule.formula && assertSheetUnprotected(sheet),
            { ...rule, formula: formula }
          );
        },
        conditionalFormats = sheet.conditionalFormats?.map(rewriteRule),
        validations = sheet.validations?.map(rewriteRule);
      if (sheet.id !== sheetId)
        return {
          ...sheet,
          cells: cells,
          ...(conditionalFormats
            ? { conditionalFormats: conditionalFormats }
            : {}),
          ...(validations ? { validations: validations } : {}),
        };
      const sizeMapName =
          structuralEdit.axis === "row" ? "rowHeights" : "columnWidths",
        remappedSizes = {};
      for (const [axisIndexKey, axisSize] of Object.entries(
        sheet[sizeMapName] ?? {},
      )) {
        const remappedAxisIndex = remapAxisIndex(
          Number(axisIndexKey),
          structuralEdit,
        );
        if (remappedAxisIndex !== null) {
          if (
            remappedAxisIndex >=
            (structuralEdit.axis === "row"
              ? workbookLimits.maxRows
              : workbookLimits.maxCols)
          )
            throw Error(
              "Insertion would discard explicit sizing at the worksheet boundary.",
            );
          remappedSizes[remappedAxisIndex] = axisSize;
        }
      }
      const merges = (sheet.merges ?? [])
          .map((mergeRange) =>
            structuralEdit.moveBefore !== undefined
              ? shiftRangeForAxisEdit(
                  mergeRange,
                  structuralEdit,
                  "merged range",
                  true,
                )
              : assertRangeWithinSheetBounds(mergeRange, structuralEdit),
          )
          .filter((keptMergeRange) => !!keptMergeRange),
        boundedValidations = (validations ?? []).flatMap((validation) => {
          const range = assertRangeWithinSheetBounds(
            validation.range,
            structuralEdit,
          );
          return range ? [{ ...validation, range: range }] : [];
        }),
        boundedConditionalFormats = (conditionalFormats ?? []).flatMap(
          (conditionalFormat) => {
            const range = assertRangeWithinSheetBounds(
              conditionalFormat.range,
              structuralEdit,
            );
            return range ? [{ ...conditionalFormat, range: range }] : [];
          },
        ),
        objects = sheet.objects?.map((object) => {
          const anchor = assertInsertionWithinSheet(
            object.anchor,
            structuralEdit,
          );
          if (!anchor)
            throw Error(
              `This edit would remove the anchor of ${object.label || object.kind}. Move or delete that drawing first.`,
            );
          if (object.kind !== "chart" || !object.range)
            return { ...object, anchor: anchor };
          const rangeAddress = parseRangeAddress(object.range),
            removedEndIndex = structuralEdit.index + structuralEdit.count - 1;
          if (
            structuralEdit.remove &&
            structuralEdit.axis === "row" &&
            object.headers &&
            structuralEdit.index <= rangeAddress.firstRow &&
            removedEndIndex >= rangeAddress.firstRow
          )
            throw Error(
              "This edit would remove a chart header row. Adjust or delete the chart first.",
            );
          if (
            structuralEdit.remove &&
            structuralEdit.axis === "column" &&
            rangeAddress.lastCol > rangeAddress.firstCol &&
            structuralEdit.index <= rangeAddress.firstCol &&
            removedEndIndex >= rangeAddress.firstCol
          )
            throw Error(
              "This edit would remove chart category labels. Adjust or delete the chart first.",
            );
          const range = assertRangeWithinSheetBounds(
            object.range,
            structuralEdit,
          );
          if (!range)
            throw Error(
              "This edit would remove the entire chart source. Adjust or delete the chart first.",
            );
          const chartSelection = measureChartSelection(range, sheet);
          if (
            structuralEdit.moveBefore !== undefined &&
            ((structuralEdit.axis === "row" &&
              object.headers &&
              remapAxisIndex(rangeAddress.firstRow, structuralEdit) !==
                chartSelection.firstRow) ||
              (structuralEdit.axis === "column" &&
                rangeAddress.lastCol > rangeAddress.firstCol &&
                remapAxisIndex(rangeAddress.firstCol, structuralEdit) !==
                  chartSelection.firstCol))
          )
            throw Error(
              "This move would displace chart headers or category labels. Move the complete chart source or only its data rows or value columns.",
            );
          if (
            (object.headers &&
              chartSelection.lastRow === chartSelection.firstRow) ||
            (rangeAddress.lastCol > rangeAddress.firstCol &&
              chartSelection.lastCol === chartSelection.firstCol)
          )
            throw Error(
              "This edit would remove all chart data. Adjust or delete the chart first.",
            );
          const chartOptions = object.chartOptions
            ? structuredClone(object.chartOptions)
            : undefined;
          if (chartOptions) {
            const remapChartIndex = (oneBasedIndex, selectionLabel) => {
              const remappedIndex = remapAxisIndex(
                oneBasedIndex - 1,
                structuralEdit,
              );
              if (remappedIndex === null)
                throw Error(
                  `This edit would remove a selected chart ${selectionLabel}. Reset the chart selection before deleting it.`,
                );
              return remappedIndex + 1;
            };
            (structuralEdit.axis === "row" &&
              chartOptions.includedRows &&
              (chartOptions.includedRows = chartOptions.includedRows.map(
                (includedRow) => remapChartIndex(includedRow, "filter row"),
              )),
              structuralEdit.axis === "column" &&
                chartOptions.series &&
                (chartOptions.series = chartOptions.series.map((sery) => ({
                  ...sery,
                  column: remapChartIndex(sery.column, "series"),
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
          structuralEdit.moveBefore !== undefined
            ? shiftPrintLayoutForAxis(sheet, structuralEdit)
            : sheet.pageLayout
              ? shiftPrintSettingsOnAxisEdit(
                  sheet.pageLayout,
                  sheet,
                  structuralEdit.axis,
                  structuralEdit.index,
                  structuralEdit.count,
                  structuralEdit.remove,
                )
              : undefined;
      if (pageLayout) {
        const boundedPrintArea = assertRangeWithinSheetBounds(
          pageLayout.printArea,
          structuralEdit,
        );
        if (!boundedPrintArea)
          throw Error(
            "This edit would remove the entire print area. Change the print area first.",
          );
        pageLayout.printArea = boundedPrintArea;
      }
      return {
        ...sheet,
        cells: cells,
        ...(structuralEdit.moveBefore !== undefined
          ? collectAxisStructureConflicts(sheet, structuralEdit)
          : applyAxisStructureEdit(
              { ...sheet, merges: merges },
              structuralEdit.axis,
              structuralEdit.index,
              structuralEdit.count,
              structuralEdit.remove,
            )),
        ...(structuralEdit.axis === "row" && sheet.rowOriginIds
          ? {
              rowOriginIds: Object.fromEntries(
                Object.entries(sheet.rowOriginIds).flatMap(
                  ([rowIndexKey, rowOriginId]) => {
                    const remappedRowIndex = remapAxisIndex(
                      Number(rowIndexKey),
                      structuralEdit,
                    );
                    if (remappedRowIndex === null) return [];
                    if (remappedRowIndex >= workbookLimits.maxRows)
                      throw Error(
                        "Insertion would discard source-row provenance at the worksheet boundary.",
                      );
                    return [[String(remappedRowIndex), rowOriginId]];
                  },
                ),
              ),
            }
          : {}),
        [sizeMapName]: remappedSizes,
        merges: merges,
        validations: boundedValidations,
        conditionalFormats: boundedConditionalFormats,
        ...(objects ? { objects: objects } : {}),
        ...(pageLayout ? { pageLayout: pageLayout } : {}),
      };
    }),
    definedNames = checkedWorkbook.definedNames?.map((definedName) => {
      const scopeSheetName = definedName.scopeSheetId
          ? (checkedWorkbook.sheets.find(
              (sheet) => sheet.id === definedName.scopeSheetId,
            )?.name ?? "")
          : "",
        refersTo = rewriteFormulaForStructuralEdit(
          definedName.refersTo,
          scopeSheetName,
          targetSheet.name,
          structuralEdit,
        );
      if (
        !definedName.scopeSheetId &&
        rewriteFormulaForStructuralEdit(
          definedName.refersTo,
          targetSheet.name,
          targetSheet.name,
          structuralEdit,
        ) !== refersTo
      )
        throw Error(
          `Workbook-scoped name ${definedName.name} has ambiguous unqualified worksheet references. Qualify those references before inserting or deleting rows or columns.`,
        );
      return { ...definedName, refersTo: refersTo };
    }),
    shiftedWorkbook = shiftReviewTargetsOnAxisEdit(
      shiftFormulaWatchesOnAxisEdit(
        {
          ...checkedWorkbook,
          sheets: sheets,
          ...(definedNames ? { definedNames: definedNames } : {}),
        },
        sheetId,
        structuralEdit,
      ),
      sheetId,
      structuralEdit,
    );
  for (const sheet of shiftedWorkbook.sheets) {
    const originalSheet = workbook.sheets.find(
      (workbookSheet) => workbookSheet.id === sheet.id,
    );
    sheet.id !== sheetId &&
      JSON.stringify(sheet.cells) !== JSON.stringify(originalSheet.cells) &&
      assertSheetUnprotected(originalSheet);
  }
  return normalizeWorkbookInput(shiftedWorkbook);
}
const structuralEditLimits = {
  sourceCells: 500,
  dependencyCells: 1e3,
  formulaCells: 5e3,
  depth: 8,
  pageSize: 20,
};

function resolveDuplicateReviewScope(workbook, hostContext) {
  return (
    hostContext.includeHidden !== undefined &&
      typeof hostContext.includeHidden != "boolean" &&
      failWithMessage("Include hidden must be boolean."),
    hostContext.allowedSheetIds !== undefined &&
      (!Array.isArray(hostContext.allowedSheetIds) ||
        hostContext.allowedSheetIds.some(
          (allowedSheetId) =>
            typeof allowedSheetId != "string" ||
            !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
        )) &&
      failWithMessage("The duplicate review host context is invalid."),
    {
      includeHidden: !!hostContext.includeHidden,
      allowedSheetIds: [
        ...new Set(
          hostContext.allowedSheetIds ??
            workbook.sheets.map((sheet) => sheet.id),
        ),
      ].sort(),
    }
  );
}
function Jde(workbook, formulaText, hostSheetId, address) {
  let rewrittenFormula = "",
    charIndex = 0;
  for (; charIndex < formulaText.length;) {
    if (formulaText[charIndex] === "'") {
      const quoteStart = charIndex++;
      for (; charIndex < formulaText.length;)
        if (formulaText[charIndex++] === "'")
          if (formulaText[charIndex] === "'") charIndex++;
          else break;
      rewrittenFormula += formulaText.slice(quoteStart, charIndex);
      continue;
    }
    const tokenStart = charIndex,
      structuredRefMatch = /^(?:[A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
        formulaText.slice(charIndex),
      );
    if (
      structuredRefMatch &&
      (charIndex === 0 || !/[A-Za-z0-9_.$]/.test(formulaText[charIndex - 1]))
    ) {
      charIndex += structuredRefMatch[0].length;
      let bracketDepth = 1;
      for (; charIndex < formulaText.length && bracketDepth;)
        (formulaText[charIndex] === "[" && bracketDepth++,
          formulaText[charIndex] === "]" && bracketDepth--,
          charIndex++);
      const structuredRef = formulaText.slice(tokenStart, charIndex),
        rewrittenStructuredRef =
          !bracketDepth &&
          !/^[A-Za-z0-9_!]/.test(formulaText.slice(charIndex)) &&
          rewriteStructuredReferences(
            workbook,
            hostSheetId,
            structuredRef,
            address,
          );
      rewrittenFormula += rewrittenStructuredRef ? "0" : structuredRef;
      continue;
    }
    rewrittenFormula += formulaText[charIndex++];
  }
  return rewrittenFormula;
}
function Qde(workbook, outputSheetName) {
  let formulaCharCount = 0,
    isOutputSheetReferenced = false;
  const inspectFormula = (formulaText, sheetId, address) => {
    if (!formulaText) return;
    ((formulaCharCount += formulaText.length),
      formulaCharCount > 1e6 &&
        failWithMessage(
          "Output-name safety inspection exceeds 1,000,000 formula characters. Use a smaller reviewed workbook.",
        ));
    const rewrittenFormula = Jde(
      workbook,
      formulaText.replace(/"(?:[^"]|"")*"/g, ""),
      sheetId,
      address,
    );
    for (const referenceMatch of rewrittenFormula.matchAll(
      /('(?:[^']|'')*'|[A-Za-z_.\d\u007F-\uFFFF]+)\s*!/g,
    )) {
      const sheetToken = referenceMatch[1];
      (sheetToken.startsWith("'")
        ? sheetToken.slice(1, -1).replace(/''/g, "'")
        : sheetToken
      ).toLowerCase() === outputSheetName.toLowerCase() &&
        (isOutputSheetReferenced = true);
    }
  };
  for (const sheet of workbook.sheets) {
    for (const [cellAddress, cell] of Object.entries(sheet.cells))
      isFormulaInputValue(cell) &&
        inspectFormula(String(cell.raw), sheet.id, cellAddress);
    for (const conditionalFormat of sheet.conditionalFormats ?? [])
      if (conditionalFormat.kind === "formula") {
        const rangeAddress = parseRangeAddress(conditionalFormat.range);
        inspectFormula(
          conditionalFormat.formula,
          sheet.id,
          formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol),
        );
      }
    for (const table of sheet.tables ?? []) {
      const rangeAddress = parseRangeAddress(table.range);
      table.columns.forEach((column, columnOffset) =>
        inspectFormula(
          column.calculatedFormula,
          sheet.id,
          formatCellAddress(
            Math.min(rangeAddress.firstRow + 1, rangeAddress.lastRow),
            rangeAddress.firstCol + columnOffset,
          ),
        ),
      );
    }
  }
  for (const definedName of workbook.definedNames ?? [])
    inspectFormula(
      definedName.refersTo,
      definedName.scopeSheetId ?? workbook.sheets[0].id,
    );
  return isOutputSheetReferenced;
}
function normalizeDuplicateReviewRequest(workbook, request) {
  ((!request ||
    typeof request != "object" ||
    Object.keys(request).some(
      (requestKey) =>
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
        ].includes(requestKey),
    )) &&
    failWithMessage("Invalid duplicate review request."),
    (typeof request.sheetId != "string" || typeof request.range != "string") &&
      failWithMessage("Choose a source worksheet and one rectangular range."));
  const rangeAddress = parseRangeAddress(request.range),
    columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1;
  (columnCount * rowCount > structuralEditLimits.sourceCells &&
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
      (flagValue) => typeof flagValue != "boolean",
    ) ||
      !["raw", "calculated"].includes(request.comparison) ||
      !["match", "keep"].includes(request.blankPolicy)) &&
      failWithMessage(
        "Choose valid comparison, header, case, trim and blank rules.",
      ),
    rowCount <= (request.headers ? 1 : 0) &&
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
      keyColumns: [...new Set(request.keyColumns)].sort(
        (leftColumn, rightColumn) => leftColumn - rightColumn,
      ),
      outputSheetName: outputSheetName,
    }
  );
}
function Pz(workbook, prefix) {
  const usedIds = new Set(workbook.sheets.map((sheet) => sheet.id)),
    collectSheetIds = (node) => {
      if (!(!node || typeof node != "object"))
        for (const [propertyName, propertyValue] of Object.entries(node))
          /sheetId$/i.test(propertyName) && typeof propertyValue == "string"
            ? usedIds.add(propertyValue)
            : typeof propertyValue == "object" &&
              collectSheetIds(propertyValue);
    };
  collectSheetIds(workbook);
  let suffix = 1;
  for (; usedIds.has(`${prefix}-${suffix}`);) suffix++;
  return `${prefix}-${suffix}`;
}
function runDuplicateInspection(workbook, sheet, range, duplicateReviewScope) {
  const rangeAddress = parseRangeAddress(range),
    cellAddresses = [];
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
      cellAddresses.push(formatCellAddress(firstRow, firstCol));
  if (
    (cellAddresses.some(
      (cellAddress) =>
        !sheet.cells[cellAddress] || sheet.cells[cellAddress].raw === null,
    ) &&
      failWithMessage(
        "Calculated comparison requires populated source cells; blank or possible spill-child cells are outside this subset. Use raw comparison or a populated selection.",
      ),
    !cellAddresses.some((cellAddress) =>
      isFormulaInputValue(sheet.cells[cellAddress]),
    ))
  )
    return new Map(
      cellAddresses.map((cellAddress) => [
        cellAddress,
        sheet.cells[cellAddress].raw,
      ]),
    );
  const inspectionSheet = {
    id: Pz(workbook, "duplicate-inspection"),
    name: "Duplicate inspection",
    rowCount: workbookLimits.maxRows,
    colCount: workbookLimits.maxCols,
    cells: {
      A1: { raw: `=SUM('${sheet.name.replace(/'/g, "''")}'!${range})` },
    },
  };
  let isIsolated = false;
  for (let attempt = 0; attempt < 4; attempt++) {
    inspectionSheet.name = `_dup_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`;
    const lowerCandidateName = inspectionSheet.name.toLowerCase();
    if (
      !workbook.sheets.some(
        (sheet) =>
          sheet.name.toLowerCase() === lowerCandidateName ||
          Object.values(sheet.cells).some(
            (cell) =>
              typeof cell.raw == "string" &&
              cell.raw.toLowerCase().includes(lowerCandidateName),
          ),
      ) &&
      !workbook.definedNames?.some((definedName) =>
        definedName.refersTo.toLowerCase().includes(lowerCandidateName),
      )
    ) {
      isIsolated = true;
      break;
    }
  }
  isIsolated ||
    failWithMessage(
      "A temporary duplicate inspection context could not be isolated.",
    );
  const dependencyResult = inspectFormulaDependencies(
      {
        ...workbook,
        sheets: [
          inspectionSheet,
          ...workbook.sheets.map((sheet) =>
            sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
          ),
        ],
      },
      {
        sheetId: inspectionSheet.id,
        address: "A1",
        direction: "precedents",
        depth: structuralEditLimits.depth,
        includeHidden: duplicateReviewScope.includeHidden,
        allowedSheetIds: [
          ...duplicateReviewScope.allowedSheetIds,
          inspectionSheet.id,
        ],
        limits: {
          maxNodes: structuralEditLimits.dependencyCells,
          maxFormulaCells: structuralEditLimits.formulaCells,
          maxReferences: 1e4,
        },
      },
    ),
    nodeKeys = new Set(dependencyResult.nodes.map((node) => node.key));
  (dependencyResult.coverage.truncated ||
    dependencyResult.coverage.depthLimited ||
    dependencyResult.cycles.length ||
    dependencyResult.issues.some(
      (issue) =>
        issue.from === dependencyResult.target.key || nodeKeys.has(issue.from),
    )) &&
    failWithMessage(
      "Calculated comparison requires a complete supported static dependency closure within 1,000 cells, 5,000 scanned formulas and depth 8; hidden, dynamic, unsupported or cyclic dependencies are excluded.",
    );
  const unsafeFunctionNames = new Set([
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
  for (const node of dependencyResult.nodes) {
    const sheet = workbook.sheets.find((sheet) => sheet.id === node.sheetId),
      cell = sheet.cells[node.address];
    if (
      ((!cell || cell.raw === null) &&
        failWithMessage(
          "Calculated comparison requires populated static dependencies; blank or possible spill-child dependencies are unsupported.",
        ),
      isFormulaInputValue(cell))
    ) {
      const resolvedFormula = resolveDefinedNameValue(
        String(cell.raw),
        workbook,
        sheet.id,
      ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
      (/[{}#]/.test(resolvedFormula) ||
        [...resolvedFormula.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)].some(
          (functionMatch) =>
            unsafeFunctionNames.has(
              functionMatch[1]
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
  const scopedWorkbook = {
      ...workbook,
      sheets: workbook.sheets
        .filter((sheet) =>
          duplicateReviewScope.allowedSheetIds.includes(sheet.id),
        )
        .map((scopedSheet) => ({
          ...scopedSheet,
          cells: Object.fromEntries(
            Object.entries(scopedSheet.cells).filter(([cellAddress]) =>
              nodeKeys.has(buildCompositeKey(scopedSheet.id, cellAddress)),
            ),
          ),
        })),
    },
    workbookWithSettings = calculateWorkbookWithSettings(scopedWorkbook),
    calculatedValueMap = new Map();
  for (const node of dependencyResult.nodes) {
    const calculatedValue = workbookWithSettings(node.sheetId, node.address);
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
        typeof calculatedValue == "string" &&
        /^#(?:[A-Z0-9_/.]+[!?]?)/.test(calculatedValue) &&
        failWithMessage(
          "Calculated comparison cannot group formula errors or error-like formula strings; correct them or use raw comparison.",
        ),
      typeof calculatedValue == "number" &&
        !Number.isFinite(calculatedValue) &&
        failWithMessage("Nonfinite values cannot be used in duplicate review."),
      node.sheetId === sheet.id &&
        cellAddresses.includes(node.address) &&
        calculatedValueMap.set(node.address, calculatedValue));
  }
  return (
    (workbookWithSettings.getOperationalErrors().length ||
      calculatedValueMap.size !== cellAddresses.length) &&
      failWithMessage(
        "Calculated source coverage is incomplete or encountered an operational error.",
      ),
    calculatedValueMap
  );
}
function createDuplicateReviewSheet(workbook, request, options = {}) {
  (options.readOnly &&
    failWithMessage(
      "This workbook is read-only. A derived worksheet cannot be created.",
    ),
    assertWorkbookStructureEditable(workbook));
  const duplicateReviewRequest = normalizeDuplicateReviewRequest(
      workbook,
      request,
    ),
    duplicateReviewScope = resolveDuplicateReviewScope(workbook, options),
    sheet = workbook.sheets.find(
      (sheet) => sheet.id === duplicateReviewRequest.sheetId,
    );
  sheet || failWithMessage("The source worksheet no longer exists.");
  const rangeAddress = parseRangeAddress(duplicateReviewRequest.range),
    columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1;
  ((!duplicateReviewScope.allowedSheetIds.includes(sheet.id) ||
    (!duplicateReviewScope.includeHidden &&
      (sheet.hidden ||
        sheet.veryHidden ||
        cn(sheet, "row").some(
          (hiddenRow) =>
            hiddenRow.start <= rangeAddress.lastRow &&
            hiddenRow.end >= rangeAddress.firstRow,
        ) ||
        cn(sheet, "column").some(
          (hiddenColumn) =>
            hiddenColumn.start <= rangeAddress.lastCol &&
            hiddenColumn.end >= rangeAddress.firstCol,
        )))) &&
    failWithMessage(
      "The source range is outside the visible, authorized host context. Include hidden only within permitted host scope.",
    ),
    (sheet.merges ?? []).some((mergeRange) => {
      const mergeAddress = parseRangeAddress(mergeRange);
      return (
        mergeAddress.firstRow <= rangeAddress.lastRow &&
        mergeAddress.lastRow >= rangeAddress.firstRow &&
        mergeAddress.firstCol <= rangeAddress.lastCol &&
        mergeAddress.lastCol >= rangeAddress.firstCol
      );
    }) &&
      failWithMessage(
        "Unmerge the selected source before duplicate review so row boundaries are unambiguous.",
      ));
  const calculatedValues =
      duplicateReviewRequest.comparison === "calculated"
        ? runDuplicateInspection(
            workbook,
            sheet,
            duplicateReviewRequest.range,
            duplicateReviewScope,
          )
        : null,
    readSourceCell = (rowIndex, columnIndex) => {
      const cellAddress = formatCellAddress(rowIndex, columnIndex),
        cell = structuredClone(sheet.cells[cellAddress] ?? { raw: null }),
        value = calculatedValues ? calculatedValues.get(cellAddress) : cell.raw;
      return (
        typeof value == "number" &&
          !Number.isFinite(value) &&
          failWithMessage(
            "Nonfinite values cannot be used in duplicate review.",
          ),
        { cell: cell, value: value }
      );
    },
    headerCells = Array.from({ length: columnCount }, (j, columnOffset) =>
      duplicateReviewRequest.headers
        ? readSourceCell(
            rangeAddress.firstRow,
            rangeAddress.firstCol + columnOffset,
          ).cell
        : {
            raw: `Column ${columnIndexToLetters(rangeAddress.firstCol + columnOffset)}`,
          },
    ),
    rows = [],
    rowsByKey = new Map();
  for (
    let sourceRowIndex =
      rangeAddress.firstRow + (duplicateReviewRequest.headers ? 1 : 0);
    sourceRowIndex <= rangeAddress.lastRow;
    sourceRowIndex++
  ) {
    const rowReads = Array.from({ length: columnCount }, (L, columnOffset) =>
        readSourceCell(sourceRowIndex, rangeAddress.firstCol + columnOffset),
      ),
      cells = rowReads.map((read) => read.cell),
      values = rowReads.map((read) => read.value),
      keyParts = duplicateReviewRequest.keyColumns.map((keyColumn) => {
        const keyOffset = keyColumn - rangeAddress.firstCol,
          keyCell = cells[keyOffset];
        let keyValue = values[keyOffset];
        return duplicateReviewRequest.comparison === "raw" &&
          isFormulaInputValue(keyCell)
          ? ["formula", keyValue]
          : (typeof keyValue == "string" &&
              (duplicateReviewRequest.trimText && (keyValue = keyValue.trim()),
              duplicateReviewRequest.caseSensitive ||
                (keyValue = keyValue.toLowerCase())),
            keyValue === null || keyValue === ""
              ? ["blank"]
              : [typeof keyValue, keyValue]);
      }),
      blankKey = keyParts.some((keyPart) => keyPart[0] === "blank"),
      normalizedKey = JSON.stringify(keyParts),
      rowRecord = {
        sourceRow: sourceRowIndex + 1,
        sourceRange: `${formatCellAddress(sourceRowIndex, rangeAddress.firstCol)}:${formatCellAddress(sourceRowIndex, rangeAddress.lastCol)}`,
        cells: cells,
        values: values,
        normalizedKey: normalizedKey,
        blankKey: blankKey,
        disposition: "retained",
        groupId: "",
        firstSourceRow: sourceRowIndex + 1,
        groupSize: 1,
        outputRow: 0,
      },
      groupingKey =
        blankKey && duplicateReviewRequest.blankPolicy === "keep"
          ? `blank-row:${sourceRowIndex}`
          : normalizedKey,
      matchingRows = rowsByKey.get(groupingKey) ?? [];
    (matchingRows.push(rowRecord),
      rowsByKey.set(groupingKey, matchingRows),
      rows.push(rowRecord));
  }
  const groups = [];
  for (const groupRows of rowsByKey.values())
    if (groupRows.length > 1) {
      const groupId = `G${groups.length + 1}`,
        sourceRow = groupRows[0].sourceRow;
      groups.push({
        id: groupId,
        firstSourceRow: sourceRow,
        memberRows: groupRows.map((memberRecord) => memberRecord.sourceRow),
        normalizedKey: groupRows[0].normalizedKey,
      });
      for (let memberIndex = 0; memberIndex < groupRows.length; memberIndex++)
        Object.assign(groupRows[memberIndex], {
          groupId: groupId,
          firstSourceRow: sourceRow,
          groupSize: groupRows.length,
          disposition: memberIndex === 0 ? "retained" : "excluded",
        });
    }
  const retainedRows = rows.filter((row) => row.disposition === "retained"),
    excludedRows = rows.filter((row) => row.disposition === "excluded"),
    counts = {
      input: rows.length,
      retained: retainedRows.length,
      excluded: excludedRows.length,
      duplicateGroups: groups.length,
      blankKeyRows: rows.filter((row) => row.blankKey).length,
    },
    outputSheet = {
      id: Pz(workbook, "duplicate-review"),
      name: duplicateReviewRequest.outputSheetName,
      rowCount: workbookLimits.maxRows,
      colCount: workbookLimits.maxCols,
      cells: {},
      frozenRows: 3,
    },
    toRawCell = (raw) => ({
      raw: raw,
      ...(typeof raw == "string" && raw.startsWith("=")
        ? { literal: true }
        : {}),
    }),
    setOutputCell = (rowNumber, columnIndex, outputCell) => {
      outputSheet.cells[formatCellAddress(rowNumber - 1, columnIndex)] =
        outputCell;
    };
  (setOutputCell(1, 0, toRawCell("Unique population — static source snapshot")),
    setOutputCell(
      2,
      0,
      toRawCell(
        `${counts.input} source rows; ${counts.retained} retained; ${counts.excluded} excluded from population. ${duplicateReviewRequest.comparison} typed keys ${duplicateReviewRequest.keyColumns.map(columnIndexToLetters).join(", ")}; ${duplicateReviewRequest.headers ? "source header" : "no source header"}; ${duplicateReviewRequest.caseSensitive ? "case-sensitive" : "ignore case"}; ${duplicateReviewRequest.trimText ? "trim outer text" : "preserve whitespace"}; blanks ${duplicateReviewRequest.blankPolicy}. First source row wins. Source unchanged.`,
      ),
    ));
  const writeHeaderRow = (headerRowNumber) => {
      (headerCells.forEach((cell, columnIndex) => {
        const headerValue =
          duplicateReviewRequest.headers && calculatedValues
            ? calculatedValues.get(
                formatCellAddress(
                  rangeAddress.firstRow,
                  rangeAddress.firstCol + columnIndex,
                ),
              )
            : cell.raw;
        setOutputCell(headerRowNumber, columnIndex, {
          ...structuredClone(cell),
          ...toRawCell(headerValue),
          literal:
            typeof headerValue == "string" && headerValue.startsWith("=")
              ? true
              : cell.literal,
        });
      }),
        [
          "Source locator",
          "Duplicate group",
          "First source row",
          "Group members",
        ].forEach((headerLabel, labelIndex) =>
          setOutputCell(
            headerRowNumber,
            columnCount + labelIndex,
            toRawCell(headerLabel),
          ),
        ));
    },
    writeRecordRow = (record, outputRow) => {
      ((record.outputRow = outputRow),
        record.cells.forEach((cell, columnIndex) => {
          const raw = record.values[columnIndex];
          setOutputCell(outputRow, columnIndex, {
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
            range: record.sourceRange,
          }),
          record.groupId,
          record.firstSourceRow,
          record.groupSize,
        ].forEach((cellValue, valueIndex) =>
          setOutputCell(
            outputRow,
            columnCount + valueIndex,
            toRawCell(cellValue),
          ),
        ));
    };
  (writeHeaderRow(3),
    retainedRows.forEach((retainedRow, retainedIndex) =>
      writeRecordRow(retainedRow, 4 + retainedIndex),
    ));
  const populationRange = `A3:${columnIndexToLetters(columnCount + 3)}${3 + retainedRows.length}`;
  let excludedRange = null;
  if (excludedRows.length) {
    const excludedStartRow = retainedRows.length + 6;
    (setOutputCell(
      excludedStartRow,
      0,
      toRawCell(
        "Excluded duplicate rows — retained here and unchanged on the source",
      ),
    ),
      writeHeaderRow(excludedStartRow + 1),
      excludedRows.forEach((excludedRow, excludedIndex) =>
        writeRecordRow(excludedRow, excludedStartRow + 2 + excludedIndex),
      ),
      (excludedRange = `A${excludedStartRow + 1}:${columnIndexToLetters(columnCount + 3)}${excludedStartRow + 1 + excludedRows.length}`));
  }
  workbook.sheets.reduce(
    (total, existingSheet) => total + Object.keys(existingSheet.cells).length,
    0,
  ) +
    Object.keys(outputSheet.cells).length >
    workbookLimits.maxCells &&
    failWithMessage(
      "The derived worksheet would exceed the workbook's stored-cell budget.",
    );
  const next = structuredClone(workbook);
  return (
    next.sheets.push(outputSheet),
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
      outputSheetId: outputSheet.id,
      populationRange: populationRange,
      excludedRange: excludedRange,
      next: next,
    }
  );
}
function applyDuplicateReviewPreview(workbook, preview, options = {}) {
  (options.readOnly && failWithMessage("This workbook is read-only."),
    JSON.stringify(workbook) !== preview.snapshot &&
      failWithMessage(
        "The workbook changed after duplicate review. Preview again before Apply.",
      ));
  const duplicateReviewSheet = createDuplicateReviewSheet(
    workbook,
    preview.request,
    options,
  );
  return (
    JSON.stringify(duplicateReviewSheet) !== JSON.stringify(preview) &&
      failWithMessage(
        "The duplicate preview or host context changed. Preview and review again.",
      ),
    duplicateReviewSheet.next
  );
}
const describeCellValueType = (value) =>
    value === null
      ? "Blank"
      : typeof value == "string"
        ? JSON.stringify(value)
        : `${typeof value}: ${String(value)}`,
  qk = (rangeText) => {
    try {
      const rangeAddress = parseRangeAddress(rangeText);
      return rangeAddress.lastCol - rangeAddress.firstCol + 1 <= 500
        ? Array.from(
            { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
            (t, columnOffset) => rangeAddress.firstCol + columnOffset,
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
