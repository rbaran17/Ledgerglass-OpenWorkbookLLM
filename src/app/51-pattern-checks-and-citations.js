import { parseNumericText } from "./37-statistics-and-linear-solver.js";
import { listRangeColumnIndexes } from "./21-aging-and-materiality-screens.js";

import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  buildCalculationClock,
  sha256HexDigest,
} from "./01-schemas-date-functions-and-validation.js";
import {
  findRowOrigin,
  _t,
  cn,
  parseRangeAddress,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  importNativeJsonFile,
  toFileProcessingError,
} from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  validateCitationRange,
  validateProposedCalculation,
  answerGroundingLimits,
  buildAssistantSelectionContext,
} from "./20-calculation-options-and-assistant.js";
import {
  formatScreenCountsSummary,
  exportScreenReportJson,
  nanToEmptyText,
  runPatternChecks,
  weekdayNames,
  checkReportCurrency,
  buildScreenRule,
  gf,
  normalizePatternCheckSettings,
  screenRuleLabels,
  parseScreenReportRequest,
} from "./50-materiality-screen-and-pattern-checks.js";

function PatternCheckDialog(props) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    dialogTitleId = ne.useId(),
    dialogRef = ne.useRef(null),
    requestSequence = ne.useRef(0),
    isMounted = ne.useRef(true),
    latestStateKey = ne.useRef(""),
    nextRuleNumber = ne.useRef(2),
    [configuration, setConfiguration] = ne.useState(() =>
      structuredClone(
        props.initialConfiguration ?? {
          sheetId: props.sheetId,
          range: props.range,
          headers: true,
          comparison: "calculated",
          excludeBlankRows: true,
          excludedRows: [],
          rules: [buildScreenRule("duplicates", "rule1")],
          combination: { mode: "all", ruleIds: [] },
        },
      ),
    ),
    [excludedRowsText, setExcludedRowsText] = ne.useState(() =>
      configuration.excludedRows.join(", "),
    ),
    [checked, setIncludeHidden] = ne.useState(false),
    [newRuleKind, setNewRuleKind] = ne.useState("duplicates"),
    [lastReport, setLastReport] = ne.useState(null),
    [reportSourceKey, setReportSourceKey] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [errorStateKey, setErrorStateKey] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [evidenceRuleId, setEvidenceRuleId] = ne.useState(""),
    [evidencePage, setEvidencePage] = ne.useState(0),
    settings = { ...configuration, sheetId: props.sheetId },
    checkOptions = {
      readOnly: props.readOnly,
      allowedSheetIds: props.allowedSheetIds,
      includeHidden: !!props.allowHidden && checked,
    },
    sourceKey = JSON.stringify([props.sheetId, props.range]),
    contextKey = JSON.stringify([
      props.workbook,
      props.allowedSheetIds,
      checkOptions,
      sourceKey,
    ]),
    stateKey = JSON.stringify([contextKey, settings, excludedRowsText]);
  latestStateKey.current = stateKey;
  const reportCurrency = lastReport
      ? reportSourceKey !== sourceKey
        ? {
            current: false,
            disclose: false,
            reason: "The host selection changed. Run checks again.",
          }
        : checkReportCurrency(
            props.workbook,
            lastReport,
            settings,
            checkOptions,
          )
      : null,
    disclosedReport = reportCurrency?.disclose ? lastReport : null,
    selectedResult =
      disclosedReport?.results.find(
        (result) => result.rule.id === evidenceRuleId,
      ) ?? disclosedReport?.results[0],
    activeResult =
      evidenceRuleId === "combined"
        ? disclosedReport?.combination
        : selectedResult,
    pageCount = Math.max(
      1,
      Math.ceil((activeResult?.rows.length ?? 0) / gf.pageSize),
    ),
    X$1 = Math.min(evidencePage, pageCount - 1),
    rangeColumnIndexes = listRangeColumnIndexes(settings.range),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    updateConfiguration = (patch) =>
      setConfiguration((previous) => ({ ...previous, ...patch })),
    updateRule = (ruleIndex, ruleChanges) =>
      setConfiguration((previous) => ({
        ...previous,
        rules: previous.rules.map((rule, rulePosition) =>
          rulePosition === ruleIndex ? { ...rule, ...ruleChanges } : rule,
        ),
      })),
    showCheckError = (error) => {
      (setErrorMessage(
        error instanceof Error
          ? error.message
          : "The check failed. Review source and settings.",
      ),
        setErrorStateKey(contextKey),
        setStatusMessage(""));
    };
  ne.useEffect(() => {
    isMounted.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Source range"]')
        ?.focus(),
      () => {
        ((isMounted.current = false),
          requestSequence.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []);
  const downloadJsonFile = (contents, fileName) => {
      const objectUrl = URL.createObjectURL(
          new Blob([contents], { type: "application/json" }),
        ),
        link = document.createElement("a");
      ((link.href = objectUrl),
        (link.download = fileName),
        link.click(),
        URL.revokeObjectURL(objectUrl));
    },
    onClick = () => {
      requestSequence.current++;
      try {
        const report = runPatternChecks(props.workbook, settings, checkOptions);
        (setLastReport(report),
          setReportSourceKey(sourceKey),
          setErrorMessage(""),
          setStatusMessage(""),
          setEvidenceRuleId(report.results[0].rule.id),
          setEvidencePage(0));
      } catch (error) {
        (setLastReport(null), showCheckError(error));
      }
    },
    loadConfigurationFile = async (file) => {
      if (!file) return;
      const requestId = ++requestSequence.current,
        startStateKey = stateKey;
      try {
        if (file.size > gf.reportBytes)
          throw Error("Choose a configuration no larger than 2,000,000 bytes.");
        const requestJson = await importNativeJsonFile(
          file,
          cancellableOperation.begin(),
        );
        if (
          !isMounted.current ||
          requestSequence.current !== requestId ||
          latestStateKey.current !== startStateKey
        )
          return;
        const screenReportRequest = parseScreenReportRequest(
          props.workbook,
          requestJson,
        );
        if (screenReportRequest.sheetId !== props.sheetId)
          throw Error("The saved configuration belongs to another worksheet.");
        (setConfiguration(screenReportRequest),
          setExcludedRowsText(screenReportRequest.excludedRows.join(", ")),
          setLastReport(null),
          setErrorMessage(""),
          setStatusMessage(
            "Configuration loaded. Run checks against the current source; saved findings are not trusted.",
          ));
      } catch (error) {
        isMounted.current &&
          requestSequence.current === requestId &&
          latestStateKey.current === startStateKey &&
          showCheckError(toFileProcessingError(error));
      }
    },
    columnLabel = (columnIndex) => {
      try {
        const rangeAddress = parseRangeAddress(settings.range),
          headerValue =
            sheet?.cells[
              `${columnIndexToLetters(columnIndex)}${rangeAddress.firstRow + 1}`
            ]?.raw;
        return `${columnIndexToLetters(columnIndex)}${settings.headers && headerValue !== null && headerValue !== undefined ? " — " + String(headerValue).slice(0, 60) : ""}`;
      } catch {
        return columnIndexToLetters(columnIndex);
      }
    },
    renderColumnSelect = (rule, ruleIndex) =>
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Field" }),
          i.jsxs("select", {
            "aria-label": `${rule.id} field`,
            value: "column" in rule ? (rule.column ?? "") : "",
            onChange: (event) =>
              updateRule(ruleIndex, {
                column:
                  event.target.value === "" ? null : Number(event.target.value),
              }),
            children: [
              i.jsx("option", { value: "", children: "Select a column" }),
              rangeColumnIndexes.map((rangeColumnIndex) =>
                i.jsx(
                  "option",
                  {
                    value: rangeColumnIndex,
                    children: columnLabel(rangeColumnIndex),
                  },
                  rangeColumnIndex,
                ),
              ),
            ],
          }),
        ],
      }),
    renderColumnCheckboxes = (rule, ruleIndex, fieldName, selectedColumns) =>
      i.jsxs("fieldset", {
        className: "pattern-checks",
        children: [
          i.jsx("legend", {
            children:
              fieldName === "keyColumns"
                ? "Matching key fields"
                : "Sequence grouping fields (optional)",
          }),
          rangeColumnIndexes.map((rangeColumnIndex) =>
            i.jsxs(
              "label",
              {
                children: [
                  i.jsx("input", {
                    "aria-label": `${rule.id} ${fieldName} ${columnIndexToLetters(rangeColumnIndex)}`,
                    type: "checkbox",
                    checked: selectedColumns.includes(rangeColumnIndex),
                    onChange: (event) =>
                      updateRule(ruleIndex, {
                        [fieldName]: event.target.checked
                          ? [...selectedColumns, rangeColumnIndex]
                          : selectedColumns.filter(
                              (selectedColumn) =>
                                selectedColumn !== rangeColumnIndex,
                            ),
                      }),
                  }),
                  columnLabel(rangeColumnIndex),
                ],
              },
              rangeColumnIndex,
            ),
          ),
        ],
      });
  return i.jsx("div", {
    className: "materiality-backdrop",
    children: i.jsxs("section", {
      className: "materiality-dialog pattern-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": dialogTitleId,
      ref: dialogRef,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? [],
            ).filter((element) => element.getClientRects().length),
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
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: dialogTitleId,
                  children: i.jsx(LocalizedText, {
                    children: "Pattern checks",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Duplicates, sequence gaps and unusual postings · read-only analysis",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close pattern checks"),
              onClick: props.onClose,
              lang: language("Close pattern checks"),
              "data-ui-label": "Close pattern checks",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Choose up to 500 source cells and map each rule explicitly. Results are review candidates, not proof of error or fraud. Every data-row position remains in the evidence ledger.",
          }),
        }),
        i.jsxs("div", {
          className: "materiality-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: settings.range,
                  onChange: (event) =>
                    updateConfiguration({ range: event.target.value }),
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Compare" }),
                i.jsxs("select", {
                  "aria-label": text("Comparison basis"),
                  value: settings.comparison,
                  onChange: (event) =>
                    updateConfiguration({ comparison: event.target.value }),
                  lang: language("Comparison basis"),
                  "data-ui-label": "Comparison basis",
                  children: [
                    i.jsx("option", {
                      value: "calculated",
                      children: "Calculated scalar values",
                    }),
                    i.jsx("option", {
                      value: "raw",
                      children: "Stored raw values / expressions",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Exclude row numbers (comma separated)",
                }),
                i.jsx("input", {
                  "aria-label": text("Excluded rows"),
                  value: excludedRowsText,
                  onChange: (event) => {
                    (setExcludedRowsText(event.target.value),
                      updateConfiguration({
                        excludedRows: event.target.value.trim()
                          ? event.target.value.split(",").map(parseNumericText)
                          : [],
                      }));
                  },
                  lang: language("Excluded rows"),
                  "data-ui-label": "Excluded rows",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "materiality-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: settings.headers,
                  onChange: (event) =>
                    updateConfiguration({ headers: event.target.checked }),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "materiality-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: settings.excludeBlankRows,
                  onChange: (event) =>
                    updateConfiguration({
                      excludeBlankRows: event.target.checked,
                    }),
                }),
                i.jsx(LocalizedText, {
                  children: "Exclude entirely blank rows",
                }),
              ],
            }),
            props.allowHidden &&
              i.jsxs("label", {
                className: "materiality-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => setIncludeHidden(event.target.checked),
                  }),
                  i.jsx(LocalizedText, {
                    children:
                      "Include authorized hidden / filtered rows and columns",
                  }),
                ],
              }),
          ],
        }),
        settings.rules.map((rule, ruleIndex) =>
          i.jsxs(
            "fieldset",
            {
              className: "pattern-rule",
              children: [
                i.jsxs("legend", {
                  children: [
                    ruleIndex + 1,
                    ". ",
                    screenRuleLabels[rule.kind],
                    " · ",
                    rule.id,
                  ],
                }),
                i.jsxs("div", {
                  className: "pattern-rule-heading",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Rule label" }),
                        i.jsx("input", {
                          "aria-label": `${rule.id} label`,
                          maxLength: 80,
                          value: rule.label,
                          onChange: (event) =>
                            updateRule(ruleIndex, {
                              label: event.target.value,
                            }),
                        }),
                      ],
                    }),
                    i.jsxs("button", {
                      disabled: settings.rules.length === 1,
                      onClick: () =>
                        updateConfiguration({
                          rules: settings.rules.filter(
                            (rule, position) => position !== ruleIndex,
                          ),
                          combination: {
                            ...settings.combination,
                            ruleIds: settings.combination.ruleIds.filter(
                              (ruleId) => ruleId !== rule.id,
                            ),
                          },
                        }),
                      children: [
                        i.jsx(LocalizedText, { children: "Remove " }),
                        rule.id,
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "materiality-fields",
                  children: [
                    rule.kind !== "duplicates" &&
                      renderColumnSelect(rule, ruleIndex),
                    (rule.kind === "duplicates" || rule.kind === "sequence") &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            className: "materiality-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: rule.caseSensitive,
                                onChange: (event) =>
                                  updateRule(ruleIndex, {
                                    caseSensitive: event.target.checked,
                                  }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Case sensitive keys",
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            className: "materiality-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: rule.trimText,
                                onChange: (event) =>
                                  updateRule(ruleIndex, {
                                    trimText: event.target.checked,
                                  }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Trim key whitespace",
                              }),
                            ],
                          }),
                        ],
                      }),
                    rule.kind === "duplicates" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Blank key policy",
                          }),
                          i.jsxs("select", {
                            "aria-label": `${rule.id} blank key policy`,
                            value: rule.blankKeys,
                            onChange: (event) =>
                              updateRule(ruleIndex, {
                                blankKeys: event.target.value,
                              }),
                            children: [
                              i.jsx("option", {
                                value: "fail",
                                children: "Fail the rule",
                              }),
                              i.jsx("option", {
                                value: "exclude",
                                children: "Exclude with reason",
                              }),
                            ],
                          }),
                        ],
                      }),
                    rule.kind === "sequence" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          ["prefix", "suffix"].map((fieldName) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  fieldName,
                                  i.jsx("input", {
                                    "aria-label": `${rule.id} ${fieldName}`,
                                    value: rule[fieldName],
                                    onChange: (event) =>
                                      updateRule(ruleIndex, {
                                        [fieldName]: event.target.value,
                                      }),
                                  }),
                                ],
                              },
                              fieldName,
                            ),
                          ),
                          ["step", "start", "end"].map((fieldName) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  fieldName,
                                  fieldName !== "step" ? " (optional)" : "",
                                  i.jsx(LocalizedNumberField, {
                                    allowEmpty: fieldName !== "step",
                                    "aria-label": `${rule.id} ${fieldName}`,
                                    min: fieldName === "step" ? 1 : 0,
                                    step: "1",
                                    value: nanToEmptyText(rule[fieldName]),
                                    onValueChange: (inputText) =>
                                      updateRule(ruleIndex, {
                                        [fieldName]:
                                          fieldName !== "step" &&
                                          inputText === ""
                                            ? null
                                            : parseNumericText(inputText),
                                      }),
                                  }),
                                ],
                              },
                              fieldName,
                            ),
                          ),
                        ],
                      }),
                    rule.kind === "round-amount" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          ["increment", "precision", "tolerance"].map(
                            (fieldName) =>
                              i.jsxs(
                                "label",
                                {
                                  children: [
                                    fieldName,
                                    i.jsx(LocalizedNumberField, {
                                      "aria-label": `${rule.id} ${fieldName}`,
                                      min: "0",
                                      step:
                                        fieldName === "precision" ? "1" : "any",
                                      value: nanToEmptyText(rule[fieldName]),
                                      onValueChange: (inputText) =>
                                        updateRule(ruleIndex, {
                                          [fieldName]:
                                            parseNumericText(inputText),
                                        }),
                                    }),
                                  ],
                                },
                                fieldName,
                              ),
                          ),
                          i.jsxs("label", {
                            className: "materiality-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: rule.includeZero,
                                onChange: (event) =>
                                  updateRule(ruleIndex, {
                                    includeZero: event.target.checked,
                                  }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Flag zero amounts",
                              }),
                            ],
                          }),
                        ],
                      }),
                    (rule.kind === "weekend" || rule.kind === "holiday") &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Excel date system",
                          }),
                          i.jsxs("select", {
                            "aria-label": `${rule.id} date system`,
                            value: rule.dateSystem,
                            onChange: (event) =>
                              updateRule(ruleIndex, {
                                dateSystem: event.target.value,
                              }),
                            children: [
                              i.jsx("option", {
                                value: "1900",
                                children: "1900",
                              }),
                              i.jsx("option", {
                                value: "1904",
                                children: "1904",
                              }),
                            ],
                          }),
                        ],
                      }),
                    rule.kind === "holiday" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calendar name",
                              }),
                              i.jsx("input", {
                                "aria-label": `${rule.id} calendar name`,
                                value: rule.calendarName,
                                onChange: (event) =>
                                  updateRule(ruleIndex, {
                                    calendarName: event.target.value,
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calendar version",
                              }),
                              i.jsx("input", {
                                "aria-label": `${rule.id} calendar version`,
                                value: rule.calendarVersion,
                                onChange: (event) =>
                                  updateRule(ruleIndex, {
                                    calendarVersion: event.target.value,
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children:
                                  "Holiday dates (YYYY-MM-DD, one per line)",
                              }),
                              i.jsx("textarea", {
                                "aria-label": `${rule.id} holiday dates`,
                                rows: 4,
                                value: rule.dates.join(`
`),
                                onChange: (event) =>
                                  updateRule(ruleIndex, {
                                    dates:
                                      event.target.value === ""
                                        ? []
                                        : event.target.value.split(`
`),
                                  }),
                              }),
                            ],
                          }),
                        ],
                      }),
                  ],
                }),
                rule.kind === "duplicates" &&
                  renderColumnCheckboxes(
                    rule,
                    ruleIndex,
                    "keyColumns",
                    rule.keyColumns,
                  ),
                rule.kind === "sequence" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      renderColumnCheckboxes(
                        rule,
                        ruleIndex,
                        "groupColumns",
                        rule.groupColumns,
                      ),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Without start/end bounds, only gaps between observed values can be inferred. Sequence gaps are intervals with boundary citations, not missing source rows.",
                        }),
                      }),
                    ],
                  }),
                rule.kind === "weekend" &&
                  i.jsxs("fieldset", {
                    className: "pattern-checks",
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Weekend weekdays",
                        }),
                      }),
                      weekdayNames.map((weekdayName, dayIndex) =>
                        i.jsxs(
                          "label",
                          {
                            children: [
                              i.jsx("input", {
                                "aria-label": `${rule.id} ${weekdayName}`,
                                type: "checkbox",
                                checked: rule.days.includes(dayIndex),
                                onChange: (event) =>
                                  updateRule(ruleIndex, {
                                    days: event.target.checked
                                      ? [...rule.days, dayIndex]
                                      : rule.days.filter(
                                          (day) => day !== dayIndex,
                                        ),
                                  }),
                              }),
                              weekdayName,
                            ],
                          },
                          weekdayName,
                        ),
                      ),
                    ],
                  }),
                rule.kind === "holiday" &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Use an approved, versioned calendar for the relevant jurisdiction. No official holiday dates are inferred.",
                    }),
                  }),
              ],
            },
            rule.id,
          ),
        ),
        i.jsxs("div", {
          className: "materiality-actions",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Add a rule" }),
                i.jsx("select", {
                  "aria-label": text("New rule type"),
                  value: newRuleKind,
                  onChange: (event) => setNewRuleKind(event.target.value),
                  lang: language("New rule type"),
                  "data-ui-label": "New rule type",
                  children: Object.entries(screenRuleLabels).map(
                    ([ruleKind, ruleLabel]) =>
                      i.jsx(
                        "option",
                        { value: ruleKind, children: ruleLabel },
                        ruleKind,
                      ),
                  ),
                }),
              ],
            }),
            i.jsx("button", {
              disabled: settings.rules.length >= gf.rules,
              onClick: () => {
                let ruleId = `rule${nextRuleNumber.current++}`;
                for (; settings.rules.some((rule) => rule.id === ruleId);)
                  ruleId = `rule${nextRuleNumber.current++}`;
                updateConfiguration({
                  rules: [
                    ...settings.rules,
                    buildScreenRule(newRuleKind, ruleId),
                  ],
                });
              },
              children: i.jsx(LocalizedText, { children: "Add rule" }),
            }),
            i.jsxs("span", {
              children: [
                settings.rules.length,
                " ",
                i.jsx(LocalizedText, { children: "/ 12 rules" }),
              ],
            }),
          ],
        }),
        i.jsxs("fieldset", {
          className: "pattern-rule",
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, {
                children: "Combine row rules (optional)",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Match" }),
                " ",
                i.jsxs("select", {
                  "aria-label": text("Combination mode"),
                  value: settings.combination.mode,
                  onChange: (event) =>
                    updateConfiguration({
                      combination: {
                        ...settings.combination,
                        mode: event.target.value,
                      },
                    }),
                  lang: language("Combination mode"),
                  "data-ui-label": "Combination mode",
                  children: [
                    i.jsx("option", {
                      value: "all",
                      children: "All selected rules",
                    }),
                    i.jsx("option", {
                      value: "any",
                      children: "Any selected rule",
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("div", {
              className: "pattern-checks",
              children: settings.rules
                .filter((rule) => rule.kind !== "sequence")
                .map((rule) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          "aria-label": `Combine ${rule.id}`,
                          checked: settings.combination.ruleIds.includes(
                            rule.id,
                          ),
                          onChange: (event) =>
                            updateConfiguration({
                              combination: {
                                ...settings.combination,
                                ruleIds: event.target.checked
                                  ? [...settings.combination.ruleIds, rule.id]
                                  : settings.combination.ruleIds.filter(
                                      (ruleId) => ruleId !== rule.id,
                                    ),
                              },
                            }),
                        }),
                        rule.label,
                        " (",
                        rule.id,
                        ")",
                      ],
                    },
                    rule.id,
                  ),
                ),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "A row must be tested by every selected rule. Failed or not-applicable rules block a combined conclusion.",
              }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "materiality-actions",
          children: [
            i.jsx("button", {
              className: "materiality-primary",
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Run pattern checks",
              }),
            }),
            i.jsx("button", {
              disabled: !reportCurrency?.current,
              onClick: () => {
                try {
                  if (!lastReport) throw Error("Run a current report first.");
                  (downloadJsonFile(
                    exportScreenReportJson(
                      props.workbook,
                      lastReport,
                      settings,
                      checkOptions,
                    ),
                    "pattern-checks-report.json",
                  ),
                    setStatusMessage(
                      "Report download requested with source evidence, settings, counts and findings.",
                    ));
                } catch (error) {
                  showCheckError(error);
                }
              },
              children: i.jsx(LocalizedText, { children: "Download report" }),
            }),
            i.jsx("button", {
              onClick: () => {
                try {
                  (downloadJsonFile(
                    JSON.stringify(
                      {
                        schemaVersion: 1,
                        kind: "pattern-audit",
                        request: normalizePatternCheckSettings(
                          props.workbook,
                          settings,
                        ),
                      },
                      null,
                      2,
                    ),
                    "pattern-checks-configuration.json",
                  ),
                    setStatusMessage("Configuration download requested."));
                } catch (error) {
                  showCheckError(error);
                }
              },
              children: i.jsx(LocalizedText, {
                children: "Save configuration",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Load configuration" }),
                i.jsx("input", {
                  "aria-label": text("Load pattern configuration"),
                  type: "file",
                  accept: ".json,application/json",
                  onChange: (event) => {
                    (loadConfigurationFile(event.target.files?.[0]),
                      (event.target.value = ""));
                  },
                  lang: language("Load pattern configuration"),
                  "data-ui-label": "Load pattern configuration",
                }),
              ],
            }),
          ],
        }),
        errorMessage &&
          errorStateKey === contextKey &&
          i.jsxs("p", {
            className: "materiality-error",
            role: "alert",
            children: [
              i.jsx(LocalizedText, { children: "Check failed: " }),
              errorMessage,
            ],
          }),
        statusMessage &&
          i.jsx("p", { role: "status", children: statusMessage }),
        reportCurrency &&
          !reportCurrency.current &&
          i.jsx("p", {
            role: "status",
            className: "materiality-stale",
            children: reportCurrency.reason,
          }),
        disclosedReport &&
          i.jsxs("section", {
            className: "materiality-result",
            "aria-label": text("Pattern report"),
            lang: language("Pattern report"),
            "data-ui-label": "Pattern report",
            children: [
              i.jsxs("h3", {
                children: [
                  reportCurrency?.current ? "Current" : "Previous",
                  " ",
                  i.jsx(LocalizedText, { children: "report —" }),
                  " ",
                  disclosedReport.status,
                ],
              }),
              i.jsxs("p", {
                children: [
                  disclosedReport.source.sheetName,
                  "!",
                  disclosedReport.source.range,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  disclosedReport.source.revision,
                  " · ",
                  disclosedReport.counts.population,
                  " ",
                  i.jsx(LocalizedText, { children: "population · " }),
                  disclosedReport.counts.excluded,
                  " ",
                  i.jsx(LocalizedText, { children: "globally excluded ·" }),
                  " ",
                  disclosedReport.counts.rulesComplete,
                  " ",
                  i.jsx(LocalizedText, { children: "complete /" }),
                  " ",
                  disclosedReport.counts.rulesFailed,
                  " ",
                  i.jsx(LocalizedText, { children: "failed /" }),
                  " ",
                  disclosedReport.counts.rulesNotApplicable,
                  " ",
                  i.jsx(LocalizedText, { children: "not applicable rules" }),
                ],
              }),
              disclosedReport.messages.map((message, messageIndex) =>
                i.jsx("p", { role: "alert", children: message }, messageIndex),
              ),
              disclosedReport.results.map((result) =>
                i.jsxs(
                  "article",
                  {
                    className: "pattern-rule-result",
                    children: [
                      i.jsxs("h4", {
                        children: [
                          result.rule.label,
                          " (",
                          result.rule.id,
                          ") — ",
                          result.status,
                        ],
                      }),
                      i.jsx("p", {
                        children: formatScreenCountsSummary(result.counts),
                      }),
                      result.messages.map((message, messageIndex) =>
                        i.jsx("p", { children: message }, messageIndex),
                      ),
                    ],
                  },
                  result.rule.id,
                ),
              ),
              i.jsxs("h4", {
                children: [
                  i.jsx(LocalizedText, { children: "Combined result — " }),
                  disclosedReport.combination.status,
                ],
              }),
              i.jsx("p", {
                children: formatScreenCountsSummary(
                  disclosedReport.combination.counts,
                ),
              }),
              disclosedReport.combination.messages.map(
                (message, messageIndex) =>
                  i.jsx("p", { children: message }, messageIndex),
              ),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Evidence ledger" }),
                  i.jsxs("select", {
                    "aria-label": text("Evidence rule"),
                    value: evidenceRuleId,
                    onChange: (event) => {
                      (setEvidenceRuleId(event.target.value),
                        setEvidencePage(0));
                    },
                    lang: language("Evidence rule"),
                    "data-ui-label": "Evidence rule",
                    children: [
                      disclosedReport.results.map((result) =>
                        i.jsxs(
                          "option",
                          {
                            value: result.rule.id,
                            children: [
                              result.rule.label,
                              " (",
                              result.rule.id,
                              ")",
                            ],
                          },
                          result.rule.id,
                        ),
                      ),
                      i.jsx("option", {
                        value: "combined",
                        children: "Combined result",
                      }),
                    ],
                  }),
                ],
              }),
              i.jsx("div", {
                className: "materiality-scroll",
                children: i.jsxs("table", {
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source row",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Disposition",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Reasons / finding IDs",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source values",
                            }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: activeResult?.rows
                        .slice(X$1 * 20, (X$1 + 1) * 20)
                        .map((resultRow) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: resultRow.sourceRow }),
                                i.jsx("td", {
                                  children: resultRow.disposition,
                                }),
                                i.jsx("td", {
                                  children:
                                    [
                                      ...resultRow.reasons,
                                      ...resultRow.findingIds,
                                    ].join(" · ") || "—",
                                }),
                                i.jsx("td", {
                                  children: JSON.stringify(
                                    disclosedReport.rows.find(
                                      (row) =>
                                        row.sourceRow === resultRow.sourceRow,
                                    )?.values,
                                  ),
                                }),
                              ],
                            },
                            resultRow.sourceRow,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Evidence pages"),
                lang: language("Evidence pages"),
                "data-ui-label": "Evidence pages",
                children: [
                  i.jsx("button", {
                    disabled: !X$1,
                    onClick: () => setEvidencePage(X$1 - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous rows",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      X$1 + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      pageCount,
                    ],
                  }),
                  i.jsx("button", {
                    disabled: X$1 >= pageCount - 1,
                    onClick: () => setEvidencePage(X$1 + 1),
                    children: i.jsx(LocalizedText, { children: "Next rows" }),
                  }),
                ],
              }),
              i.jsx("h4", {
                children: i.jsx(LocalizedText, {
                  children: "Review candidates",
                }),
              }),
              i.jsxs("p", {
                children: [
                  activeResult?.findings.length ?? 0,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "finding groups / intervals. Open a finding for complete rule settings, counts, matching group and raw/evaluated source-cell citations.",
                  }),
                ],
              }),
              i.jsx("div", {
                className: "pattern-findings",
                children: activeResult?.findings.map((finding) =>
                  i.jsxs(
                    "details",
                    {
                      children: [
                        i.jsxs("summary", {
                          children: [finding.title, " · ", finding.id],
                        }),
                        i.jsx("p", { children: finding.detail }),
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, { children: "Rule " }),
                            finding.rule.id,
                            " ",
                            i.jsx(LocalizedText, { children: "v" }),
                            finding.rule.version,
                            i.jsx(LocalizedText, { children: "; group " }),
                            finding.group.id,
                            i.jsx(LocalizedText, {
                              children: "; source rows ",
                            }),
                            finding.group.memberRows.join(", "),
                          ],
                        }),
                        finding.missing &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Missing values: ",
                              }),
                              finding.missing.count,
                              i.jsx(LocalizedText, { children: "; interval" }),
                              " ",
                              finding.missing.first,
                              "–",
                              finding.missing.last,
                              i.jsx(LocalizedText, { children: "; step " }),
                              finding.missing.step,
                              i.jsx(LocalizedText, { children: "; examples " }),
                              finding.missing.examples.join(", "),
                              finding.missing.examplesTruncated
                                ? " (first 20 only; complete interval above)"
                                : "",
                            ],
                          }),
                        i.jsx("pre", {
                          children: JSON.stringify(finding, null, 2),
                        }),
                      ],
                    },
                    finding.id,
                  ),
                ),
              }),
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Source ledger, settings and interpretation notes",
                    }),
                  }),
                  disclosedReport.notes.map((note) =>
                    i.jsx("p", { children: note }, note),
                  ),
                  i.jsx("pre", {
                    children: JSON.stringify(
                      {
                        request: disclosedReport.request,
                        source: disclosedReport.source,
                        rows: disclosedReport.rows,
                      },
                      null,
                      2,
                    ),
                  }),
                ],
              }),
            ],
          }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: props.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const hashCitationText = (text) =>
  sha256HexDigest(new TextEncoder().encode(text));
function buildCitationContext(assistantSelectionContext) {
  return hashCitationText(
    JSON.stringify({
      workbookId: assistantSelectionContext.workbookId,
      revision: assistantSelectionContext.revision,
      sheetId: assistantSelectionContext.sheetId,
      sheetName: assistantSelectionContext.sheetName,
      range: assistantSelectionContext.range,
      ...buildCalculationClock(assistantSelectionContext),
      cells: assistantSelectionContext.cells.map(({ address, raw, value }) => ({
        address,
        raw,
        value,
      })),
    }),
  );
}
const citationRangeCache = new WeakMap(),
  citationRangeKey = (citation) =>
    JSON.stringify([
      citation.sheetId,
      validateCitationRange(citation.range),
      citation.ranges && citation.ranges.length !== 1
        ? citation.ranges
        : undefined,
    ]);
function validateCitationVisibility(workbook, sheetId, cellAddresses) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet || sheet.hidden || sheet.veryHidden)
    throw Error("Citation worksheet is unavailable or hidden.");
  const hiddenRows = cn(sheet, "row"),
    hiddenColumns = cn(sheet, "column");
  for (const cellAddress of cellAddresses) {
    const cellPosition = _t(cellAddress);
    if (
      hiddenRows.some(
        (hiddenRange) =>
          cellPosition.row >= hiddenRange.start &&
          cellPosition.row <= hiddenRange.end,
      ) ||
      hiddenColumns.some(
        (hiddenRange) =>
          cellPosition.col >= hiddenRange.start &&
          cellPosition.col <= hiddenRange.end,
      )
    )
      throw Error(
        "Citation includes hidden cells. Make the cited evidence visible and request a fresh answer.",
      );
  }
  return sheet;
}
function buildAnswerCitations(
  activeWorkbook,
  assistantSelection,
  proposal,
  approvedContext,
  options = {},
) {
  const validatedProposal = validateProposedCalculation(
      activeWorkbook,
      assistantSelection,
      proposal,
    ),
    assistantSelectionContext = buildAssistantSelectionContext(
      activeWorkbook,
      assistantSelection,
    );
  if (
    JSON.stringify(assistantSelectionContext) !==
    JSON.stringify(approvedContext)
  )
    throw Error(
      "Answer source is stale or outside the original approved context. Request a fresh answer.",
    );
  if (assistantSelection.ranges && assistantSelection.ranges.length !== 1)
    throw Error("Citations require one approved rectangle.");
  const sheet = activeWorkbook.sheets.find(
      (sheet) => sheet.id === assistantSelectionContext.sheetId,
    ),
    citationContext = buildCitationContext(assistantSelectionContext),
    cellsByAddress = new Map(
      assistantSelectionContext.cells.map((cell) => [cell.address, cell]),
    ),
    citations = [],
    citationIdByRange = new Map(),
    citationIdForRange = (claimRange) => {
      const cellAddresses = validateCitationRange(claimRange);
      if (cellAddresses.some((address) => !cellsByAddress.has(address)))
        throw Error("Citation is outside approved context.");
      validateCitationVisibility(
        activeWorkbook,
        assistantSelectionContext.sheetId,
        cellAddresses,
      );
      const range =
        cellAddresses.length === 1
          ? cellAddresses[0]
          : `${cellAddresses[0]}:${cellAddresses.at(-1)}`;
      let citationId = citationIdByRange.get(range);
      return (
        citationId ||
        ((citationId = `cite-${citations.length + 1}`),
        citationIdByRange.set(range, citationId),
        citations.push({
          id: citationId,
          workbookId: activeWorkbook.id,
          revision: activeWorkbook.revision,
          sheetId: assistantSelectionContext.sheetId,
          sheetName: assistantSelectionContext.sheetName,
          range: range,
          contextSha256: citationContext,
          cells: cellAddresses.map((cellAddress) => {
            const cell = cellsByAddress.get(cellAddress),
              sheetCell = sheet.cells[cellAddress],
              literal = !!sheetCell?.literal;
            return {
              ...cell,
              literal: literal,
              formula:
                !literal &&
                typeof cell.raw == "string" &&
                cell.raw.startsWith("=")
                  ? cell.raw
                  : null,
              sourceOrigin:
                options.includeSourceOrigins === false
                  ? null
                  : findRowOrigin(
                      activeWorkbook,
                      sheet.id,
                      _t(cellAddress).row,
                    ),
            };
          }),
        }),
        citationId)
      );
    },
    claims = validatedProposal.grounding
      ? validatedProposal.grounding.claims.map((claim) => ({
          id: claim.id,
          kind: claim.kind,
          text: claim.text,
          citationIds:
            claim.kind === "workbook"
              ? claim.ranges.map(citationIdForRange)
              : [],
          sourceIds: claim.kind === "external" ? [...claim.sourceIds] : [],
        }))
      : [
          {
            id: "legacy",
            kind: "commentary",
            text: validatedProposal.explanation,
            citationIds: [
              ...new Set(validatedProposal.citations.map(citationIdForRange)),
            ],
            sourceIds: [],
          },
        ],
    evidence = {
      schemaVersion: 1,
      kind: "assistant-answer-evidence",
      capturedAt: new Date().toISOString(),
      proposalId: validatedProposal.id,
      mode: validatedProposal.grounding ? "linked-claims" : "legacy-unverified",
      contextSha256: citationContext,
      approvedSelection: {
        sheetId: assistantSelection.sheetId,
        range: assistantSelectionContext.range,
      },
      citations: citations,
      claims: claims,
      externalSources: (validatedProposal.grounding?.externalSources ?? []).map(
        (externalSource) => ({
          ...externalSource,
          status: "supplied-not-verified",
        }),
      ),
      sourceOriginsIncluded: options.includeSourceOrigins !== false,
      notes: [
        "Workbook citation identity, approved scope, captured values and revision are checked locally. These checks do not prove that the cited cells logically support every claim; the reviewer must inspect the evidence.",
        "Legacy narratives and commentary are unverified. A workbook citation cannot substantiate an external fact. External URLs are adapter-supplied references, not automatically fetched or verified.",
        "Source-row linkage is copied only from the current host workbook registry. It does not certify the original file bytes or identify an original column after transformations; missing linkage is explicitly unavailable.",
        "Only host-local evidence can include row-origin metadata. Provider context remains the selected raw/calculated cells. Downloaded evidence is mutable JSON, not an immutable audit service or a replayable authority.",
      ],
    };
  if (
    new TextEncoder().encode(JSON.stringify(evidence, null, 2)).length >
    answerGroundingLimits.receiptBytes
  )
    throw Error(
      "Complete answer evidence exceeds 2,000,000 bytes. Reduce the selection or answer; evidence is never silently truncated.",
    );
  return (
    citationRangeCache.set(evidence, {
      snapshot: JSON.stringify(activeWorkbook),
      content: JSON.stringify(evidence),
      selection: citationRangeKey(assistantSelection),
    }),
    evidence
  );
}
function evaluateAnswerEvidence(workbook, selection, evidence) {
  const cacheEntry = citationRangeCache.get(evidence),
    unavailableResult = (reason) => ({
      current: false,
      disclose: false,
      reason: reason,
    });
  if (!cacheEntry || JSON.stringify(evidence) !== cacheEntry.content)
    return unavailableResult(
      "Answer evidence is untrusted or altered. Request a fresh answer.",
    );
  try {
    if (citationRangeKey(selection) !== cacheEntry.selection)
      return unavailableResult(
        "Approved selection changed. Previous answer and citations are unavailable.",
      );
    if (JSON.stringify(workbook) !== cacheEntry.snapshot)
      return unavailableResult(
        "Workbook content or revision changed. Previous answer and citations are stale.",
      );
    for (const citation of evidence.citations)
      validateCitationVisibility(
        workbook,
        citation.sheetId,
        citation.cells.map((cell) => cell.address),
      );
    return buildCitationContext(
      buildAssistantSelectionContext(workbook, selection),
    ) !== evidence.contextSha256
      ? unavailableResult(
          "Calculated source values changed. Request a fresh answer.",
        )
      : {
          current: true,
          disclose: true,
          reason: "Current approved source and revision.",
        };
  } catch {
    return unavailableResult(
      "Citation source is unavailable or outside the current authorized selection.",
    );
  }
}
function getCitationStatus(workbook, selection, evidence, citationId) {
  const evidenceStatus = evaluateAnswerEvidence(workbook, selection, evidence);
  if (!evidenceStatus.current) throw Error(evidenceStatus.reason);
  const citation = evidence.citations.find(
    (citation) => citation.id === citationId,
  );
  if (!citation) throw Error("Unknown citation.");
  return structuredClone(citation);
}
function serializeCurrentCitation(workbook, selection, evidence) {
  const evidenceStatus = evaluateAnswerEvidence(workbook, selection, evidence);
  if (!evidenceStatus.current) throw Error(evidenceStatus.reason);
  return JSON.stringify(evidence, null, 2);
}
const formatCellValueForCitation = (value) =>
  value === null
    ? "(blank)"
    : typeof value == "string"
      ? JSON.stringify(value)
      : String(value);
function AssistantAnswerEvidence(props) {
  const { text, language } = useWorkbookMessages(),
    [selectedCitationId, setSelectedCitationId] = ne.useState(null),
    [actionError, setActionError] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [cellPageIndex, setCellPageIndex] = ne.useState(0),
    evidenceStatus = evaluateAnswerEvidence(
      props.workbook,
      props.selection,
      props.evidence,
    );
  if (!evidenceStatus.disclose)
    return i.jsx("p", {
      role: "status",
      className: "lg-evidence-warning",
      children: evidenceStatus.reason,
    });
  const evidence = props.evidence,
    citation = evidence.citations.find(
      (citation) => citation.id === selectedCitationId,
    ),
    pageCount = Math.max(1, Math.ceil((citation?.cells.length ?? 0) / 20)),
    currentPageIndex = Math.min(cellPageIndex, pageCount - 1),
    navigateToCitation = (targetCitationId, cellAddress) => {
      try {
        const citationStatus = getCitationStatus(
          props.workbook,
          props.selection,
          evidence,
          targetCitationId,
        );
        if (
          cellAddress &&
          !citationStatus.cells.some((cell) => cell.address === cellAddress)
        )
          throw Error("Cell is outside this citation.");
        (props.onNavigate(
          citationStatus.sheetId,
          cellAddress ?? citationStatus.cells[0].address,
        ),
          setSelectedCitationId(targetCitationId),
          cellAddress || setCellPageIndex(0),
          setActionError(""),
          setStatusMessage(""));
      } catch (error) {
        setActionError(
          error instanceof Error ? error.message : "Citation unavailable.",
        );
      }
    };
  return i.jsxs("section", {
    className: "lg-answer-evidence",
    "aria-label": text("Assistant answer evidence"),
    lang: language("Assistant answer evidence"),
    "data-ui-label": "Assistant answer evidence",
    children: [
      evidence.mode === "legacy-unverified" &&
        i.jsx("p", {
          className: "lg-evidence-warning",
          children: i.jsx(LocalizedText, {
            children:
              "Unverified narrative: this adapter did not link individual claims to evidence. The listed cells are references only.",
          }),
        }),
      i.jsx("p", {
        className: "lg-evidence-note",
        children: i.jsx(LocalizedText, {
          children:
            "Workbook links verify source identity, scope and revision. Review whether the evidence supports each statement.",
        }),
      }),
      evidence.claims.map((claim) =>
        i.jsxs(
          "article",
          {
            className: "lg-answer-claim",
            children: [
              i.jsx("strong", {
                children:
                  claim.kind === "workbook"
                    ? "Workbook claim"
                    : claim.kind === "external"
                      ? "External claim · source not verified"
                      : "Commentary · unverified",
              }),
              props.renderText
                ? i.jsx("div", {
                    className: "lg-answer-text",
                    children: props.renderText(claim.text),
                  })
                : i.jsx("p", { children: claim.text }),
              i.jsx("div", {
                className: "lg-citations",
                children: claim.citationIds.map((citationId) => {
                  const citation = evidence.citations.find(
                    (citation) => citation.id === citationId,
                  );
                  return i.jsxs(
                    "button",
                    {
                      type: "button",
                      "aria-label": `Inspect citation ${citation.sheetName}!${citation.range}`,
                      onClick: () => navigateToCitation(citationId),
                      children: [
                        citation.sheetName,
                        "!",
                        citation.range,
                        " ",
                        i.jsx(LocalizedText, { children: "· revision " }),
                        citation.revision,
                      ],
                    },
                    citationId,
                  );
                }),
              }),
              claim.sourceIds.map((sourceId) => {
                const externalSource = evidence.externalSources.find(
                  (externalSource) => externalSource.id === sourceId,
                );
                return i.jsxs(
                  "p",
                  {
                    className: "lg-external-source",
                    children: [
                      i.jsx("a", {
                        href: externalSource.url,
                        target: "_blank",
                        rel: "noopener noreferrer",
                        referrerPolicy: "no-referrer",
                        children: externalSource.title,
                      }),
                      i.jsx("br", {}),
                      externalSource.url,
                      i.jsx("br", {}),
                      i.jsx("small", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Adapter-supplied external source; not fetched or verified.",
                        }),
                      }),
                    ],
                  },
                  sourceId,
                );
              }),
            ],
          },
          claim.id,
        ),
      ),
      actionError &&
        i.jsx("p", {
          role: "alert",
          className: "lg-evidence-warning",
          children: actionError,
        }),
      statusMessage && i.jsx("p", { role: "status", children: statusMessage }),
      citation &&
        i.jsxs("section", {
          className: "lg-citation-inspector",
          "aria-label": text("Citation source details"),
          lang: language("Citation source details"),
          "data-ui-label": "Citation source details",
          children: [
            i.jsxs("h4", {
              children: [citation.sheetName, "!", citation.range],
            }),
            i.jsxs("dl", {
              children: [
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Workbook" }),
                }),
                i.jsx("dd", { children: citation.workbookId }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Revision" }),
                }),
                i.jsx("dd", { children: citation.revision }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Sheet ID" }),
                }),
                i.jsx("dd", { children: citation.sheetId }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Captured" }),
                }),
                i.jsx("dd", { children: evidence.capturedAt }),
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, {
                  children:
                    "The cited cell is focused in the grid. The approved shared selection remains ",
                }),
                evidence.approvedSelection.range,
                ".",
              ],
            }),
            citation.cells
              .slice(currentPageIndex * 20, (currentPageIndex + 1) * 20)
              .map((cell) =>
                i.jsxs(
                  "article",
                  {
                    className: "lg-citation-cell",
                    children: [
                      i.jsx("button", {
                        onClick: () =>
                          navigateToCitation(citation.id, cell.address),
                        "aria-label": `Focus cited cell ${cell.address}`,
                        children: cell.address,
                      }),
                      i.jsxs("dl", {
                        children: [
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Stored input",
                            }),
                          }),
                          i.jsx("dd", {
                            children: formatCellValueForCitation(cell.raw),
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Formula",
                            }),
                          }),
                          i.jsx("dd", {
                            children:
                              cell.formula ??
                              (cell.literal
                                ? "Literal text; not a formula"
                                : "No formula"),
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Captured value",
                            }),
                          }),
                          i.jsx("dd", {
                            children: formatCellValueForCitation(cell.value),
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Source linkage",
                            }),
                          }),
                          i.jsx("dd", {
                            children: cell.sourceOrigin
                              ? i.jsxs(i.Fragment, {
                                  children: [
                                    cell.sourceOrigin.sourceName,
                                    " ·",
                                    " ",
                                    cell.sourceOrigin.sourceSheet,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "· original row",
                                    }),
                                    " ",
                                    cell.sourceOrigin.originalRow,
                                    i.jsx("br", {}),
                                    i.jsx(LocalizedText, {
                                      children: "Source ID: ",
                                    }),
                                    cell.sourceOrigin.sourceId,
                                    i.jsx("br", {}),
                                    i.jsx(LocalizedText, {
                                      children: "Origin ID: ",
                                    }),
                                    cell.sourceOrigin.id,
                                    " (",
                                    cell.sourceOrigin.kind,
                                    ")",
                                    i.jsx("br", {}),
                                    i.jsx("small", {
                                      children: i.jsx(LocalizedText, {
                                        children:
                                          "Recorded row lineage; original file bytes and original column are not verified here.",
                                      }),
                                    }),
                                  ],
                                })
                              : "No original source-row linkage is available in this evidence.",
                          }),
                        ],
                      }),
                    ],
                  },
                  cell.address,
                ),
              ),
            i.jsxs("nav", {
              "aria-label": text("Citation evidence pages"),
              lang: language("Citation evidence pages"),
              "data-ui-label": "Citation evidence pages",
              children: [
                i.jsx("button", {
                  disabled: currentPageIndex === 0,
                  onClick: () => setCellPageIndex(currentPageIndex - 1),
                  children: i.jsx(LocalizedText, {
                    children: "Previous cells",
                  }),
                }),
                i.jsxs("span", {
                  children: [
                    i.jsx(LocalizedText, { children: "Page " }),
                    currentPageIndex + 1,
                    " ",
                    i.jsx(LocalizedText, { children: "of " }),
                    pageCount,
                  ],
                }),
                i.jsx("button", {
                  disabled: currentPageIndex >= pageCount - 1,
                  onClick: () => setCellPageIndex(currentPageIndex + 1),
                  children: i.jsx(LocalizedText, { children: "Next cells" }),
                }),
              ],
            }),
          ],
        }),
      i.jsx("button", {
        className: "lg-button full",
        onClick: () => {
          try {
            const evidenceJson = serializeCurrentCitation(
                props.workbook,
                props.selection,
                evidence,
              ),
              downloadUrl = URL.createObjectURL(
                new Blob([evidenceJson], { type: "application/json" }),
              ),
              downloadLink = document.createElement("a");
            ((downloadLink.href = downloadUrl),
              (downloadLink.download = "assistant-answer-evidence.json"),
              downloadLink.click(),
              URL.revokeObjectURL(downloadUrl),
              setStatusMessage(
                "Evidence download requested. Preserve it with the original workpaper.",
              ),
              setActionError(""));
          } catch (error) {
            setActionError(
              error instanceof Error ? error.message : "Evidence unavailable.",
            );
          }
        },
        children: i.jsx(LocalizedText, {
          children: "Download answer evidence",
        }),
      }),
    ],
  });
}
const patternCheckBudgets = {
    mappings: 20,
    selectedCells: 1e4,
    rowsPerRange: 1e3,
    columnsPerRange: 100,
    scannedCells: 2e4,
    formulaCharacters: 2e6,
    findings: 3e4,
    nameDepth: 16,
    keys: 8,
  },
  readFormulaText = (cell) => cell?.raw ?? null,
  isBlankFormulaCell = (cell) =>
    readFormulaText(cell) === null || readFormulaText(cell) === "",
  ww = (left, right) => {
    if (left === right) return true;
    if (!left || !right || typeof left != "object" || typeof right != "object")
      return false;
    if (Array.isArray(left) || Array.isArray(right))
      return (
        Array.isArray(left) &&
        Array.isArray(right) &&
        left.length === right.length &&
        left.every((leftItem, itemIndex) => ww(leftItem, right[itemIndex]))
      );
    const leftRecord = left,
      rightRecord = right,
      leftKeys = Object.keys(leftRecord)
        .filter((leftKey) => leftRecord[leftKey] !== undefined)
        .sort(),
      rightKeys = Object.keys(rightRecord)
        .filter((rightKey) => rightRecord[rightKey] !== undefined)
        .sort();
    return (
      leftKeys.length === rightKeys.length &&
      leftKeys.every(
        (key, keyIndex) =>
          key === rightKeys[keyIndex] && ww(leftRecord[key], rightRecord[key]),
      )
    );
  },
  unquotedSheetNamePattern = "[A-Za-z_.\\d\\u007F-\\uFFFF]+",
  I3 = `(?:'(?:[^']|'')*'|${unquotedSheetNamePattern})`,
  a1RangePattern =
    "(?:\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?:\\s*:\\s*\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)?|\\$?[A-Za-z]{1,3}\\s*:\\s*\\$?[A-Za-z]{1,3}|\\$?[1-9]\\d*\\s*:\\s*\\$?[1-9]\\d*)",
  referenceTokenPattern = new RegExp(
    `^(${I3}\\s*!\\s*)?(${a1RangePattern})(#)?`,
  ),
  definedNameReferencePattern = new RegExp(
    `^(${I3})\\s*!\\s*([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*)`,
  ),
  sheetSpanPattern = new RegExp(`^${I3}\\s*:\\s*${I3}\\s*!`),
  Bj = (rawSheetName) =>
    rawSheetName
      .trim()
      .replace(/!\s*$/, "")
      .trim()
      .replace(/^'|'$/g, "")
      .replace(/''/g, "'"),
  identifierPattern = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
  isNonReferenceToken = (token) => !/[A-Za-z0-9_.!:]/.test(token),
  o2e = (a1Text, originAddress) => {
    const origin = _t(originAddress);
    return a1Text
      .replace(/\s/g, "")
      .replace(
        /(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)|(?:(\$?)([A-Za-z]{1,3}))|(?:(\$?)([1-9]\d*))/g,
        (
          match,
          columnAbsolute,
          columnLetters,
          rowAbsolute,
          rowNumber,
          columnOnlyAbsolute,
          columnOnlyLetters,
          rowOnlyAbsolute,
          rowOnlyNumber,
        ) => {
          const columnIndexOf = (letters) =>
            [...letters.toUpperCase()].reduce(
              (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
              0,
            ) - 1;
          return columnLetters
            ? `${columnAbsolute ? "C$" + columnIndexOf(columnLetters) : "C[" + (columnIndexOf(columnLetters) - origin.col) + "]"}${rowAbsolute ? "R$" + rowNumber : "R[" + (Number(rowNumber) - 1 - origin.row) + "]"}`
            : columnOnlyLetters
              ? columnOnlyAbsolute
                ? "C$" + columnIndexOf(columnOnlyLetters)
                : "C[" + (columnIndexOf(columnOnlyLetters) - origin.col) + "]"
              : rowOnlyAbsolute
                ? "R$" + rowOnlyNumber
                : "R[" + (Number(rowOnlyNumber) - 1 - origin.row) + "]";
        },
      );
  };
function computeFormulaSignature(raw, originAddress) {
  if (!raw.startsWith("=") || raw.length > 4096)
    throw Error(
      "Pattern comparison requires a formula of at most 4,096 characters.",
    );
  _t(originAddress);
  let signature = "",
    position = 0;
  for (; position < raw.length;) {
    const remainingText = raw.slice(position);
    if (remainingText[0] === '"') {
      const textLiteralMatch = /^"(?:[^"]|"")*"/.exec(remainingText);
      if (!textLiteralMatch)
        throw Error(
          "An unterminated text literal has no reliable formula signature.",
        );
      ((signature += textLiteralMatch[0]),
        (position += textLiteralMatch[0].length));
      continue;
    }
    const structuredRefMatch = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(
      remainingText,
    );
    if (structuredRefMatch) {
      const referenceStart = position;
      position += structuredRefMatch[0].length;
      let bracketDepth = 1;
      for (; position < raw.length && bracketDepth;)
        (raw[position] === "[" && bracketDepth++,
          raw[position] === "]" && bracketDepth--,
          position++);
      if (bracketDepth)
        throw Error(
          "An incomplete structured reference has no reliable formula signature.",
        );
      signature += raw.slice(referenceStart, position);
      continue;
    }
    const referenceMatch = referenceTokenPattern.exec(remainingText);
    if (
      referenceMatch &&
      isNonReferenceToken(remainingText[referenceMatch[0].length] ?? "") &&
      !/^\s*\(/.test(remainingText.slice(referenceMatch[0].length))
    ) {
      ((signature +=
        (referenceMatch[1]?.toUpperCase() ?? "") +
        o2e(referenceMatch[2], originAddress) +
        (referenceMatch[3] ?? "")),
        (position += referenceMatch[0].length));
      continue;
    }
    const tokenMatch =
      /^(?:[A-Za-z_\\][A-Za-z0-9_.\\]*|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)/.exec(
        remainingText,
      );
    if (tokenMatch) {
      ((signature += tokenMatch[0].toUpperCase()),
        (position += tokenMatch[0].length));
      continue;
    }
    signature += raw[position++];
  }
  return signature;
}

export {
  patternCheckBudgets,
  Bj,
  isBlankFormulaCell,
  buildAnswerCitations,
  AssistantAnswerEvidence,
  PatternCheckDialog,
  computeFormulaSignature,
  evaluateAnswerEvidence,
  identifierPattern,
  sheetSpanPattern,
  isNonReferenceToken,
  definedNameReferencePattern,
  referenceTokenPattern,
  ww,
  readFormulaText,
};
