import { listRangeColumnIndexes } from "./21-aging-and-materiality-screens.js";
import { throwErrorMessage, isNullOrEmptyValue } from "./03-workbook-model.js";
import { cloneStructuredValue } from "./11-local-workbook-store-and-versions.js";

import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  string,
  number,
  boolean,
  array,
  enum as _enum,
  discriminatedUnion,
  literal,
  object,
} from "zod";
import { sha256HexDigest } from "./01-schemas-date-functions-and-validation.js";
import {
  findRowOrigin,
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
  materialityScreenLimits,
  loadMaterialityConfiguration,
  checkMaterialityReportFreshness,
  buildMaterialitySourceEvidence,
  exportMaterialityReportJson,
  runMaterialityScreen,
} from "./21-aging-and-materiality-screens.js";
import {
  comparatorSymbols,
  formatCellValueLabel,
} from "./49-review-tools-and-outlines.js";

function MaterialityRiskScreen(props) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    latestSettingsKey = ne.useRef(""),
    requestCounter = ne.useRef(0),
    isMountedRef = ne.useRef(true),
    rangeColumnIndexes = listRangeColumnIndexes(props.range),
    [range, setRange] = ne.useState(props.range),
    [hasHeaders, setHasHeaders] = ne.useState(true),
    [amountColumn, setAmountColumn] = ne.useState(rangeColumnIndexes[0] ?? 0),
    [currencyColumn, setCurrencyColumn] = ne.useState(
      rangeColumnIndexes[1] ?? 1,
    ),
    [dateColumn, setDateColumn] = ne.useState(rangeColumnIndexes[2] ?? 2),
    [currency, setCurrency] = ne.useState("USD"),
    [thresholdText, setThresholdText] = ne.useState("50000"),
    [amountMode, setAmountMode] = ne.useState("absolute"),
    [comparator, setComparator] = ne.useState("gte"),
    [effectiveFrom, setEffectiveFrom] = ne.useState("1900-01-01"),
    [effectiveTo, setEffectiveTo] = ne.useState("9999-12-31"),
    [checked, setIncludeHidden] = ne.useState(false),
    [screenReport, setScreenReport] = ne.useState(null),
    [reportSourceKey, setReportSourceKey] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [errorScopeKey, setErrorScopeKey] = ne.useState(""),
    [X$1, setStatusMessage] = ne.useState(""),
    [pageIndex, setPageIndex] = ne.useState(0),
    [dispositionFilter, setDispositionFilter] = ne.useState("all"),
    screenRequest = {
      sheetId: props.sheetId,
      range: range,
      headers: hasHeaders,
      amountColumn: amountColumn,
      currencyColumn: currencyColumn,
      dateColumn: dateColumn,
      currency: currency,
      threshold: thresholdText.trim() === "" ? NaN : Number(thresholdText),
      amountMode: amountMode,
      comparator: comparator,
      effectiveFrom: effectiveFrom,
      effectiveTo: effectiveTo,
    },
    screenOptions = {
      readOnly: props.readOnly,
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: props.allowedSheetIds,
    },
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    isSourceAvailable =
      !!sheet &&
      allowedSheetIds.includes(sheet.id) &&
      (screenOptions.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    sourceSelectionKey = JSON.stringify([props.sheetId, props.range]),
    workbookScopeKey = JSON.stringify([
      props.workbook,
      allowedSheetIds,
      screenOptions.includeHidden,
      sourceSelectionKey,
    ]),
    settingsKey = JSON.stringify([workbookScopeKey, screenRequest]);
  latestSettingsKey.current = settingsKey;
  const reportFreshness = screenReport
      ? reportSourceKey !== sourceSelectionKey
        ? {
            current: false,
            disclose: false,
            reason: "The source selection changed. Run the screen again.",
          }
        : checkMaterialityReportFreshness(
            props.workbook,
            screenReport,
            screenRequest,
            screenOptions,
          )
      : null,
    selectableColumnIndexes = listRangeColumnIndexes(range),
    visibleReport = reportFreshness?.disclose ? screenReport : null,
    filteredRows =
      visibleReport?.rows.filter(
        (row) =>
          dispositionFilter === "all" || row.disposition === dispositionFilter,
      ) ?? [],
    pageCount = Math.max(
      1,
      Math.ceil(filteredRows.length / materialityScreenLimits.pageSize),
    ),
    currentPageIndex = Math.min(pageIndex, pageCount - 1),
    reportScreenError = (error) => {
      (setErrorMessage(
        error instanceof Error
          ? error.message
          : "The screen could not complete. Review the settings and source.",
      ),
        setErrorScopeKey(workbookScopeKey),
        setStatusMessage(""));
    };
  ne.useEffect(() => {
    isMountedRef.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Source range"]')
        ?.focus(),
      () => {
        ((isMountedRef.current = false),
          requestCounter.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []);
  const onClick = () => {
      requestCounter.current++;
      try {
        const newReport = runMaterialityScreen(
          props.workbook,
          screenRequest,
          screenOptions,
        );
        (setScreenReport(newReport),
          setReportSourceKey(sourceSelectionKey),
          setErrorMessage(""),
          setStatusMessage(""),
          setPageIndex(0),
          setDispositionFilter("all"));
      } catch (error) {
        (setScreenReport(null), reportScreenError(error));
      }
    },
    downloadReport = () => {
      try {
        if (!screenReport || !reportFreshness?.current)
          throw Error("Run a current screen before downloading a report.");
        const reportJson = exportMaterialityReportJson(
            props.workbook,
            screenReport,
            screenRequest,
            screenOptions,
          ),
          objectUrl = URL.createObjectURL(
            new Blob([reportJson], { type: "application/json" }),
          ),
          downloadLink = document.createElement("a");
        ((downloadLink.href = objectUrl),
          (downloadLink.download = "materiality-screen.json"),
          downloadLink.click(),
          URL.revokeObjectURL(objectUrl),
          setStatusMessage(
            "Report downloaded with settings, source rows, exclusions and control totals.",
          ),
          setErrorMessage(""));
      } catch (error) {
        reportScreenError(error);
      }
    },
    loadConfiguration = async (configFile) => {
      if (!configFile) return;
      const startedSettingsKey = settingsKey,
        requestId = ++requestCounter.current,
        isCurrentRequest = () =>
          isMountedRef.current &&
          requestId === requestCounter.current &&
          latestSettingsKey.current === startedSettingsKey;
      try {
        if (configFile.size > 2e6)
          throw Error(
            "Choose a configuration/report no larger than 2,000,000 bytes.",
          );
        const configurationJson = await importNativeJsonFile(
          configFile,
          cancellableOperation.begin(),
        );
        if (!isCurrentRequest()) return;
        const materialityConfiguration = loadMaterialityConfiguration(
          props.workbook,
          configurationJson,
          screenOptions,
        );
        if (materialityConfiguration.sheetId !== props.sheetId)
          throw Error(
            "Open the saved configuration's source worksheet before loading it.",
          );
        (setRange(materialityConfiguration.range),
          setHasHeaders(materialityConfiguration.headers),
          setAmountColumn(materialityConfiguration.amountColumn),
          setCurrencyColumn(materialityConfiguration.currencyColumn),
          setDateColumn(materialityConfiguration.dateColumn),
          setCurrency(materialityConfiguration.currency),
          setThresholdText(String(materialityConfiguration.threshold)),
          setAmountMode(materialityConfiguration.amountMode),
          setComparator(materialityConfiguration.comparator),
          setEffectiveFrom(materialityConfiguration.effectiveFrom),
          setEffectiveTo(materialityConfiguration.effectiveTo),
          setScreenReport(null),
          setErrorMessage(""),
          setStatusMessage(
            "Configuration loaded. Run the screen against the current source; saved conclusions are not reused.",
          ));
      } catch (error) {
        isCurrentRequest() && reportScreenError(toFileProcessingError(error));
      }
    };
  return i.jsx("div", {
    className: "materiality-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      className: "materiality-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      ref: dialogRef,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
            dialogRef.current?.querySelectorAll(
              'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex="0"]',
            ) ?? [],
          ).filter(
            (element) =>
              !element.closest("details:not([open])") ||
              element.tagName === "SUMMARY",
          );
          event.shiftKey && document.activeElement === focusableElements[0]
            ? (event.preventDefault(),
              (lastFocusable = focusableElements.at(-1)) == null ||
                lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === focusableElements.at(-1) &&
              (event.preventDefault(),
              (firstFocusable = focusableElements[0]) == null ||
                firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${dialogId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Materiality and risk screen",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Review a user-defined quantitative threshold against a declared currency and period.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close materiality screen"),
              onClick: props.onClose,
              lang: language("Close materiality screen"),
              "data-ui-label": "Close materiality screen",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "This screen applies your settings; it does not recommend audit materiality. It reads the source without editing it. Protected and read-only workbooks support inspection.",
          }),
        }),
        i.jsxs("div", {
          className: "materiality-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source worksheet" }),
                i.jsx("strong", {
                  children: isSourceAvailable
                    ? sheet.name
                    : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: isSourceAvailable ? range : "",
                  onChange: (event) => setRange(event.target.value),
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "materiality-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: hasHeaders,
                  onChange: (event) => setHasHeaders(event.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            [
              ["Amount column", amountColumn, setAmountColumn],
              ["Currency column", currencyColumn, setCurrencyColumn],
              ["Date column", dateColumn, setDateColumn],
            ].map(
              ([columnLabel, selectedColumnIndex, setSelectedColumnIndex]) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      columnLabel,
                      i.jsx("select", {
                        "aria-label": columnLabel,
                        value: selectedColumnIndex,
                        onChange: (event) =>
                          setSelectedColumnIndex(Number(event.target.value)),
                        children: selectableColumnIndexes.map((columnIndex) =>
                          i.jsx(
                            "option",
                            {
                              value: columnIndex,
                              children: columnIndexToLetters(columnIndex),
                            },
                            columnIndex,
                          ),
                        ),
                      }),
                    ],
                  },
                  columnLabel,
                ),
            ),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Declared currency" }),
                i.jsx("input", {
                  "aria-label": text("Declared currency"),
                  value: currency,
                  maxLength: 3,
                  onChange: (event) => setCurrency(event.target.value),
                  lang: language("Declared currency"),
                  "data-ui-label": "Declared currency",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Threshold" }),
                i.jsx(LocalizedNumberField, {
                  "aria-label": text("Threshold"),
                  min: "0",
                  step: "any",
                  value: thresholdText,
                  onValueChange: (thresholdValue) =>
                    setThresholdText(thresholdValue),
                  lang: language("Threshold"),
                  "data-ui-label": "Threshold",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Amount comparison" }),
                i.jsxs("select", {
                  "aria-label": text("Amount comparison"),
                  value: amountMode,
                  onChange: (event) => setAmountMode(event.target.value),
                  lang: language("Amount comparison"),
                  "data-ui-label": "Amount comparison",
                  children: [
                    i.jsx("option", {
                      value: "absolute",
                      children: "Absolute amount",
                    }),
                    i.jsx("option", {
                      value: "signed",
                      children: "Signed amount",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Comparison rule" }),
                i.jsxs("select", {
                  "aria-label": text("Comparison rule"),
                  value: comparator,
                  onChange: (event) => setComparator(event.target.value),
                  lang: language("Comparison rule"),
                  "data-ui-label": "Comparison rule",
                  children: [
                    i.jsx("option", {
                      value: "gte",
                      children: "Greater than or equal to (≥)",
                    }),
                    i.jsx("option", {
                      value: "gt",
                      children: "Greater than (>)",
                    }),
                    i.jsx("option", {
                      value: "lte",
                      children: "Less than or equal to (≤)",
                    }),
                    i.jsx("option", { value: "lt", children: "Less than (<)" }),
                    i.jsx("option", {
                      value: "eq",
                      children: "Exactly equal to (=)",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Effective from (inclusive)",
                }),
                i.jsx("input", {
                  "aria-label": text("Effective from"),
                  type: "date",
                  value: effectiveFrom,
                  onChange: (event) => setEffectiveFrom(event.target.value),
                  lang: language("Effective from"),
                  "data-ui-label": "Effective from",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Effective to (inclusive)" }),
                i.jsx("input", {
                  "aria-label": text("Effective to"),
                  type: "date",
                  value: effectiveTo,
                  onChange: (event) => setEffectiveTo(event.target.value),
                  lang: language("Effective to"),
                  "data-ui-label": "Effective to",
                }),
              ],
            }),
          ],
        }),
        props.allowHidden &&
          i.jsxs("label", {
            className: "materiality-check",
            children: [
              i.jsx("input", {
                "aria-label": text("Include hidden content"),
                type: "checkbox",
                checked: checked,
                onChange: (event) => setIncludeHidden(event.target.checked),
                lang: language("Include hidden content"),
                "data-ui-label": "Include hidden content",
              }),
              i.jsx(LocalizedText, {
                children:
                  "Explicitly include hidden source and dependency cells within the permitted host context",
              }),
            ],
          }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Up to 500 selected cells. Numeric text is an exception. Dates must be ISO calendar days or integer Excel 1900-system serials. Currency identifiers are case-insensitive, with no trimming or conversion. Missing or mixed currency blocks the entire screen.",
          }),
        }),
        i.jsxs("div", {
          className: "materiality-actions",
          children: [
            i.jsx("button", {
              className: "materiality-primary",
              disabled: !isSourceAvailable,
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Run materiality screen",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Load saved configuration/report",
                }),
                i.jsx("input", {
                  "aria-label": text("Load saved configuration"),
                  type: "file",
                  accept: "application/json,.json",
                  onChange: (event) => {
                    const selectedFile = event.target.files?.[0];
                    ((event.target.value = ""),
                      loadConfiguration(selectedFile));
                  },
                  lang: language("Load saved configuration"),
                  "data-ui-label": "Load saved configuration",
                }),
              ],
            }),
            i.jsx("button", {
              onClick: downloadReport,
              disabled: !reportFreshness?.current,
              children: i.jsx(LocalizedText, {
                children: "Download current report",
              }),
            }),
          ],
        }),
        errorMessage &&
          errorScopeKey === workbookScopeKey &&
          i.jsx("p", {
            role: "alert",
            className: "materiality-error",
            children: errorMessage,
          }),
        X$1 && i.jsx("p", { role: "status", children: X$1 }),
        screenReport &&
          reportFreshness &&
          i.jsx("p", {
            role: "status",
            className: reportFreshness.current ? "" : "materiality-stale",
            children: reportFreshness.current
              ? reportFreshness.reason
              : `Stale result. ${reportFreshness.reason}`,
          }),
        visibleReport &&
          i.jsxs("section", {
            className: "materiality-result",
            "aria-label": text("Materiality screen results"),
            lang: language("Materiality screen results"),
            "data-ui-label": "Materiality screen results",
            children: [
              i.jsx("h3", {
                children:
                  visibleReport.status === "blocked"
                    ? "Blocked — resolve currency exceptions"
                    : visibleReport.status === "complete-with-exclusions"
                      ? "Complete with exclusions — review every excluded row"
                      : "Screen complete",
              }),
              i.jsxs("p", {
                children: [
                  visibleReport.request.amountMode === "absolute"
                    ? "Absolute amount"
                    : "Signed amount",
                  " ",
                  comparatorSymbols[visibleReport.request.comparator],
                  " ",
                  visibleReport.request.threshold,
                  " ",
                  visibleReport.request.currency,
                  " ",
                  "· ",
                  visibleReport.request.effectiveFrom,
                  " ",
                  i.jsx(LocalizedText, { children: "through" }),
                  " ",
                  visibleReport.request.effectiveTo,
                  i.jsx(LocalizedText, { children: ", inclusive." }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  visibleReport.counts.population,
                  " ",
                  i.jsx(LocalizedText, { children: "source rows ·" }),
                  " ",
                  visibleReport.counts.screened,
                  " ",
                  i.jsx(LocalizedText, { children: "screened ·" }),
                  " ",
                  visibleReport.counts.flagged,
                  " ",
                  i.jsx(LocalizedText, { children: "flagged ·" }),
                  " ",
                  visibleReport.counts.notFlagged,
                  " ",
                  i.jsx(LocalizedText, { children: "not flagged ·" }),
                  " ",
                  visibleReport.counts.excluded,
                  " ",
                  i.jsx(LocalizedText, { children: "excluded ·" }),
                  " ",
                  visibleReport.counts.notScreened,
                  " ",
                  i.jsx(LocalizedText, { children: "not screened" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  visibleReport.counts.currencyExceptions,
                  " ",
                  i.jsx(LocalizedText, { children: "currency exceptions ·" }),
                  " ",
                  visibleReport.counts.invalidAmounts,
                  " ",
                  i.jsx(LocalizedText, { children: "invalid amounts ·" }),
                  " ",
                  visibleReport.counts.invalidDates,
                  " ",
                  i.jsx(LocalizedText, { children: "invalid dates ·" }),
                  " ",
                  visibleReport.counts.outsidePeriod,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "outside period. Exception categories may overlap.",
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Signed " }),
                  visibleReport.totals.currency,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "control totals: known source ",
                  }),
                  visibleReport.totals.knownSourceAmount,
                  i.jsx(LocalizedText, { children: "; eligible" }),
                  " ",
                  visibleReport.totals.eligibleAmount,
                  i.jsx(LocalizedText, { children: "; excluded known amount" }),
                  " ",
                  visibleReport.totals.excludedKnownAmount,
                  i.jsx(LocalizedText, { children: "; flagged" }),
                  " ",
                  visibleReport.totals.flaggedAmount ?? "unavailable (blocked)",
                  ".",
                  " ",
                  visibleReport.totals.unknownOrOtherCurrencyRows,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "rows have unknown amounts or another/unknown currency and are not added to these totals.",
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Source: " }),
                  visibleReport.sheetName,
                  "!",
                  visibleReport.request.range,
                  " ",
                  i.jsx(LocalizedText, { children: "· workbook " }),
                  visibleReport.workbookId,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  visibleReport.revision,
                  " · ",
                  visibleReport.algorithm,
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Rows to review" }),
                  i.jsxs("select", {
                    "aria-label": text("Rows to review"),
                    value: dispositionFilter,
                    onChange: (event) => {
                      (setDispositionFilter(event.target.value),
                        setPageIndex(0));
                    },
                    lang: language("Rows to review"),
                    "data-ui-label": "Rows to review",
                    children: [
                      i.jsx("option", {
                        value: "all",
                        children: "All source rows",
                      }),
                      i.jsx("option", {
                        value: "flagged",
                        children: "Flagged",
                      }),
                      i.jsx("option", {
                        value: "not-flagged",
                        children: "Not flagged",
                      }),
                      i.jsx("option", {
                        value: "excluded",
                        children: "Excluded",
                      }),
                      i.jsx("option", {
                        value: "not-screened",
                        children: "Not screened",
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
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Source row / range",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Disposition",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Amount / compared",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Currency / date",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Reasons",
                            }),
                          }),
                          visibleReport.headers.map((header, headerIndex) =>
                            i.jsxs(
                              "th",
                              {
                                scope: "col",
                                children: [
                                  header,
                                  " (",
                                  columnIndexToLetters(
                                    parseRangeAddress(
                                      visibleReport.request.range,
                                    ).firstCol + headerIndex,
                                  ),
                                  ")",
                                ],
                              },
                              headerIndex,
                            ),
                          ),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: filteredRows
                        .slice(
                          currentPageIndex * 20,
                          currentPageIndex * 20 + 20,
                        )
                        .map((row) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsxs("th", {
                                  scope: "row",
                                  children: [
                                    row.sourceRow,
                                    " · ",
                                    row.sourceRange,
                                  ],
                                }),
                                i.jsx("td", { children: row.disposition }),
                                i.jsxs("td", {
                                  children: [
                                    row.amount ?? "Invalid",
                                    " /",
                                    " ",
                                    row.comparedAmount ?? "Unavailable",
                                  ],
                                }),
                                i.jsxs("td", {
                                  children: [
                                    row.currency ?? "Invalid",
                                    " / ",
                                    row.date ?? "Invalid",
                                  ],
                                }),
                                i.jsx("td", {
                                  children:
                                    row.reasons.join(" ") ||
                                    (row.disposition === "flagged"
                                      ? "Matches the configured comparison."
                                      : "Does not match the configured comparison."),
                                }),
                                row.cells.map((cell, cellIndex) =>
                                  i.jsx(
                                    "td",
                                    {
                                      children: i.jsxs("details", {
                                        children: [
                                          i.jsx("summary", {
                                            children: formatCellValueLabel(
                                              row.values[cellIndex],
                                            ),
                                          }),
                                          i.jsxs("p", {
                                            children: [
                                              i.jsx(LocalizedText, {
                                                children: "Stored: ",
                                              }),
                                              formatCellValueLabel(cell.raw),
                                            ],
                                          }),
                                          i.jsxs("p", {
                                            children: [
                                              i.jsx(LocalizedText, {
                                                children: "Literal: ",
                                              }),
                                              cell.literal ? "Yes" : "No",
                                            ],
                                          }),
                                          i.jsx("pre", {
                                            children: JSON.stringify(
                                              Object.fromEntries(
                                                Object.entries(cell).filter(
                                                  ([entryKey]) =>
                                                    ![
                                                      "raw",
                                                      "literal",
                                                    ].includes(entryKey),
                                                ),
                                              ),
                                              null,
                                              2,
                                            ),
                                          }),
                                        ],
                                      }),
                                    },
                                    cellIndex,
                                  ),
                                ),
                              ],
                            },
                            row.sourceRow,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Screen result pages"),
                lang: language("Screen result pages"),
                "data-ui-label": "Screen result pages",
                children: [
                  i.jsx("button", {
                    disabled: currentPageIndex === 0,
                    onClick: () => setPageIndex(currentPageIndex - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous rows",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      currentPageIndex + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      pageCount,
                      " · ",
                      filteredRows.length,
                      " ",
                      i.jsx(LocalizedText, { children: "rows in this view" }),
                    ],
                  }),
                  i.jsx("button", {
                    disabled: currentPageIndex + 1 >= pageCount,
                    onClick: () => setPageIndex(currentPageIndex + 1),
                    children: i.jsx(LocalizedText, { children: "Next rows" }),
                  }),
                ],
              }),
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children: "Screening rules and limitations",
                    }),
                  }),
                  visibleReport.notes.map((note) =>
                    i.jsx("p", { children: note }, note),
                  ),
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
const gf = Object.freeze({
    sourceCells: 500,
    rules: 12,
    keyColumns: 8,
    holidays: 1e3,
    reportBytes: 2e6,
    pageSize: 20,
    sequenceMaximum: 999999999999999,
  }),
  nullableColumnIndexSchema = number().int().min(0).max(16383).nullable(),
  keyColumnsSchema = array(number().int().min(0).max(16383)).max(gf.keyColumns),
  ruleIdentitySchema = {
    id: string().regex(/^[A-Za-z][A-Za-z0-9_-]{0,39}$/),
    label: string().trim().min(1).max(80),
  },
  textComparisonSchema = { caseSensitive: boolean(), trimText: boolean() },
  dateSystemSchema = _enum(["1900", "1904"]),
  sequenceValueSchema = number().int().min(0).max(gf.sequenceMaximum),
  screenRuleSchema = discriminatedUnion("kind", [
    object({
      ...ruleIdentitySchema,
      kind: literal("duplicates"),
      keyColumns: keyColumnsSchema,
      ...textComparisonSchema,
      blankKeys: _enum(["fail", "exclude"]),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("sequence"),
      column: nullableColumnIndexSchema,
      groupColumns: keyColumnsSchema,
      ...textComparisonSchema,
      prefix: string().max(40),
      suffix: string().max(40),
      step: sequenceValueSchema.min(1),
      start: sequenceValueSchema.nullable(),
      end: sequenceValueSchema.nullable(),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("round-amount"),
      column: nullableColumnIndexSchema,
      increment: number().finite().positive().max(1e15),
      precision: number().int().min(0).max(6),
      tolerance: number().finite().nonnegative(),
      includeZero: boolean(),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("weekend"),
      column: nullableColumnIndexSchema,
      dateSystem: dateSystemSchema,
      days: array(number().int().min(0).max(6)).min(1).max(7),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("holiday"),
      column: nullableColumnIndexSchema,
      dateSystem: dateSystemSchema,
      calendarName: string().trim().min(1).max(80),
      calendarVersion: string().trim().min(1).max(40),
      dates: array(string().max(10)).max(gf.holidays),
    }).strict(),
  ]),
  screenSettingsSchema = object({
    headers: boolean(),
    comparison: _enum(["raw", "calculated"]),
    excludeBlankRows: boolean(),
    excludedRows: array(number().int().min(1).max(1048576)).max(500),
    rules: array(screenRuleSchema).min(1).max(gf.rules),
    combination: object({
      mode: _enum(["any", "all"]),
      ruleIds: array(string().max(40)).max(gf.rules),
    }).strict(),
  }).strict(),
  sourceRangeSchema = screenSettingsSchema
    .extend({
      sheetId: string().min(1).max(100),
      range: string().regex(
        /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/,
      ),
    })
    .strict(),
  KW = new WeakMap(),
  rN = (rowNumber, columnIndex) =>
    `${columnIndexToLetters(columnIndex)}${rowNumber}`;
function parseIsoCalendarDate(dateText) {
  if (
    typeof dateText != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(dateText) ||
    dateText.startsWith("0000-")
  )
    return null;
  const timestamp = Date.parse(`${dateText}T00:00:00.000Z`);
  return Number.isFinite(timestamp) &&
    new Date(timestamp).toISOString().slice(0, 10) === dateText
    ? timestamp / 864e5
    : null;
}
function convertSerialToIsoDate(dateValue, dateSystem) {
  if (typeof dateValue == "string")
    return parseIsoCalendarDate(dateValue) === null ? null : dateValue;
  if (typeof dateValue != "number" || !Number.isFinite(dateValue)) return null;
  const serialDay = Math.floor(dateValue),
    maximumSerial = dateSystem === "1900" ? 2958465 : 2957003;
  return serialDay < (dateSystem === "1900" ? 1 : 0) ||
    serialDay > maximumSerial ||
    (dateSystem === "1900" && serialDay === 60)
    ? null
    : new Date(
        (dateSystem === "1900"
          ? Date.UTC(1899, 11, 31)
          : Date.UTC(1904, 0, 1)) +
          (dateSystem === "1900" && serialDay > 60
            ? serialDay - 1
            : serialDay) *
            864e5,
      )
        .toISOString()
        .slice(0, 10);
}
function toScaledIntegerAmount(amount, precision) {
  if (!Number.isFinite(amount) || Math.abs(amount) > 1e15) return null;
  const [mantissaText, exponentText = "0"] = Math.abs(amount)
      .toString()
      .split("e"),
    mantissaParts = mantissaText.split("."),
    scaledDigits = BigInt(mantissaParts.join("")),
    scaleShift =
      precision - (mantissaParts[1]?.length ?? 0) + Number(exponentText);
  if (scaleShift >= 0) return scaledDigits * 10n ** BigInt(scaleShift);
  const divisor = 10n ** BigInt(-scaleShift);
  return scaledDigits % divisor === 0n ? scaledDigits / divisor : null;
}
function ruleKeyColumns(rule) {
  return rule.kind === "duplicates"
    ? rule.keyColumns
    : rule.kind === "sequence"
      ? [...(rule.column === null ? [] : [rule.column]), ...rule.groupColumns]
      : rule.column === null
        ? []
        : [rule.column];
}
function normalizePatternCheckSettings(workbook, settings) {
  const parseResult = sourceRangeSchema.safeParse(settings);
  if (!parseResult.success)
    throw Error(
      `Invalid pattern-check settings: ${parseResult.error.issues
        .map((issue) => `${issue.path.join(".")}: ${issue.message}`)
        .join("; ")
        .slice(0, 1500)}`,
    );
  const data = parseResult.data,
    rangeAddress = parseRangeAddress(data.range),
    sheet = workbook.sheets.find((sheet) => sheet.id === data.sheetId);
  ((!sheet ||
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount) &&
    throwErrorMessage(
      "Pattern source is outside current worksheet dimensions.",
    ),
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      gf.sourceCells &&
      throwErrorMessage(
        "Pattern checks support at most 500 selected source cells.",
      ),
    new Set(data.rules.map((rule) => rule.id)).size !== data.rules.length &&
      throwErrorMessage("Every rule needs a unique ID."),
    (new Set(data.excludedRows).size !== data.excludedRows.length ||
      data.excludedRows.some(
        (excludedRow) =>
          excludedRow <= rangeAddress.firstRow + Number(data.headers) ||
          excludedRow > rangeAddress.lastRow + 1,
      )) &&
      throwErrorMessage(
        "Explicit exclusions must be distinct data-row numbers inside the selected range.",
      ));
  for (const rule of data.rules) {
    const keyColumns = ruleKeyColumns(rule);
    if (
      (new Set(keyColumns).size !== keyColumns.length &&
        throwErrorMessage(`Rule ${rule.id} has duplicate field mappings.`),
      rule.kind === "sequence" &&
        rule.start !== null &&
        rule.end !== null &&
        (rule.start > rule.end || (rule.end - rule.start) % rule.step !== 0) &&
        throwErrorMessage(
          "Sequence bounds must be chronological integers on the declared step.",
        ),
      rule.kind === "round-amount")
    ) {
      const scaledIntegerAmount = toScaledIntegerAmount(
          rule.increment,
          rule.precision,
        ),
        scaledTolerance = toScaledIntegerAmount(rule.tolerance, rule.precision);
      (scaledIntegerAmount === null ||
        scaledIntegerAmount === 0n ||
        scaledTolerance === null ||
        scaledTolerance * 2n >= scaledIntegerAmount) &&
        throwErrorMessage(
          "Round-amount increment/tolerance must fit the declared precision; tolerance must be less than half the increment.",
        );
    }
    (rule.kind === "weekend" &&
      new Set(rule.days).size !== rule.days.length &&
      throwErrorMessage(
        "Weekend weekdays must be distinct (0=Sunday through 6=Saturday).",
      ),
      rule.kind === "holiday" &&
        (new Set(rule.dates).size !== rule.dates.length ||
          rule.dates.some((date) => parseIsoCalendarDate(date) === null)) &&
        throwErrorMessage(
          "Holiday dates must be distinct valid ISO calendar dates.",
        ));
  }
  return (
    (new Set(data.combination.ruleIds).size !==
      data.combination.ruleIds.length ||
      data.combination.ruleIds.some(
        (ruleId) =>
          !data.rules.some(
            (rule) => rule.id === ruleId && rule.kind !== "sequence",
          ),
      )) &&
      throwErrorMessage(
        "Row combinations require distinct existing row rules. Sequence gaps are separate interval findings, not missing source rows.",
      ),
    {
      ...data,
      range: `${rN(rangeAddress.firstRow + 1, rangeAddress.firstCol)}:${rN(rangeAddress.lastRow + 1, rangeAddress.lastCol)}`,
    }
  );
}
const createEmptyCounts = (population) => ({
  population: population,
  excluded: 0,
  invalid: 0,
  tested: 0,
  matched: 0,
  notMatched: 0,
  notTested: 0,
  groups: 0,
  findings: 0,
  gapValues: 0,
});
function summarizeRuleOutcome(rows, findings) {
  const emptyCounts = createEmptyCounts(rows.length);
  for (const row of rows)
    row.disposition === "excluded"
      ? emptyCounts.excluded++
      : row.disposition === "invalid"
        ? emptyCounts.invalid++
        : row.disposition === "matched"
          ? emptyCounts.matched++
          : row.disposition === "not-matched"
            ? emptyCounts.notMatched++
            : emptyCounts.notTested++;
  ((emptyCounts.tested = emptyCounts.matched + emptyCounts.notMatched),
    (emptyCounts.groups = new Set(
      findings.map((finding) => finding.group.id),
    ).size),
    (emptyCounts.findings = findings.length));
  const missingTotal = findings.reduce(
    (total, finding) => total + BigInt(finding.missing?.count ?? 0),
    0n,
  );
  return (
    missingTotal > BigInt(Number.MAX_SAFE_INTEGER) &&
      throwErrorMessage(
        "The total missing-value count exceeds the exact integer reporting limit. Split sequence groups into smaller checks.",
      ),
    (emptyCounts.gapValues = Number(missingTotal)),
    emptyCounts
  );
}
function buildRowMatchKey(rowValues, rule) {
  const values = rowValues.map((rawValue) =>
    typeof rawValue == "string"
      ? rule.caseSensitive
        ? rule.trimText
          ? rawValue.trim()
          : rawValue
        : (rule.trimText ? rawValue.trim() : rawValue).toLowerCase()
      : rawValue,
  );
  return values.some(
    (value) =>
      isNullOrEmptyValue(value) ||
      (typeof value == "number" && !Number.isFinite(value)),
  )
    ? null
    : {
        key: JSON.stringify(values.map((value) => [typeof value, value])),
        values: values,
      };
}
function runPatternChecks(workbook, patternCheckRequest, evidenceOptions = {}) {
  const patternCheckSettings = normalizePatternCheckSettings(
      workbook,
      patternCheckRequest,
    ),
    rangeAddress = parseRangeAddress(patternCheckSettings.range),
    evidenceScope = {
      sheetId: patternCheckSettings.sheetId,
      range: patternCheckSettings.range,
      headers: patternCheckSettings.headers,
      columns: [],
    };
  let materialitySourceEvidence = buildMaterialitySourceEvidence(
      workbook,
      evidenceScope,
      evidenceOptions,
    ),
    calculationError = "";
  if (patternCheckSettings.comparison === "calculated") {
    const columns = [
      ...new Set(patternCheckSettings.rules.flatMap(ruleKeyColumns)),
    ].filter(
      (column) =>
        column >= rangeAddress.firstCol && column <= rangeAddress.lastCol,
    );
    try {
      materialitySourceEvidence = buildMaterialitySourceEvidence(
        workbook,
        { ...evidenceScope, columns: columns },
        evidenceOptions,
      );
    } catch (error) {
      calculationError =
        error instanceof Error
          ? error.message
          : "Selected fields could not be calculated.";
    }
  }
  const rows = materialitySourceEvidence.rows.map((row) => ({
      ...row,
      origin: findRowOrigin(
        workbook,
        materialitySourceEvidence.sheetId,
        row.sourceRow - 1,
      ),
      exclusionReason: patternCheckSettings.excludedRows.includes(row.sourceRow)
        ? "Explicitly excluded source row."
        : patternCheckSettings.excludeBlankRows &&
            row.values.every(isNullOrEmptyValue)
          ? "Entire selected row is blank; excluded under the visible blank-row policy."
          : null,
    })),
    source = {
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetId: materialitySourceEvidence.sheetId,
      sheetName: materialitySourceEvidence.sheetName,
      range: patternCheckSettings.range,
      sha256: sha256HexDigest(
        new TextEncoder().encode(
          JSON.stringify({
            headers: materialitySourceEvidence.headers,
            rows: rows,
          }),
        ),
      ),
    },
    rowsBySourceRow = new Map(rows.map((row) => [row.sourceRow, row])),
    sourceCellsFor = (sourceRows, columnIndexes) =>
      sourceRows.flatMap((sourceRow) =>
        columnIndexes.map((columnIndex) => {
          const evidenceRow = rowsBySourceRow.get(sourceRow);
          return {
            address: rN(sourceRow, columnIndex),
            raw: evidenceRow.cells[columnIndex - rangeAddress.firstCol].raw,
            value: evidenceRow.values[columnIndex - rangeAddress.firstCol],
            origin: evidenceRow.origin,
          };
        }),
      ),
    createRowDispositions = () =>
      rows.map((row) => ({
        sourceRow: row.sourceRow,
        disposition: row.exclusionReason ? "excluded" : "not-tested",
        reasons: row.exclusionReason ? [row.exclusionReason] : [],
        findingIds: [],
      })),
    results = patternCheckSettings.rules.map((rule) => {
      const ruleResult = {
          rule: cloneStructuredValue(rule),
          version: "1.0.0",
          status: "complete",
          messages: [],
          counts: createEmptyCounts(rows.length),
          rows: createRowDispositions(),
          findings: [],
        },
        dispositionsBySourceRow = new Map(
          ruleResult.rows.map((row) => [row.sourceRow, row]),
        ),
        keyColumns = ruleKeyColumns(rule),
        keyColumnsInRange = keyColumns.every(
          (keyColumn) =>
            keyColumn >= rangeAddress.firstCol &&
            keyColumn <= rangeAddress.lastCol,
        ),
        hasMissingMapping =
          rule.kind === "duplicates"
            ? !keyColumns.length
            : rule.column === null,
        finalizeRule = () => {
          if (ruleResult.status === "failed") {
            ruleResult.findings = [];
            for (const row of ruleResult.rows)
              row.disposition !== "excluded" &&
                row.disposition !== "invalid" &&
                ((row.disposition = "not-tested"),
                (row.findingIds = []),
                row.reasons.push(
                  "Rule failed; no clean or exception conclusion is available.",
                ));
          }
          ruleResult.counts = summarizeRuleOutcome(
            ruleResult.rows,
            ruleResult.findings,
          );
          for (const finding of ruleResult.findings)
            finding.counts = cloneStructuredValue(ruleResult.counts);
          return (
            ruleResult.status === "complete" &&
              !ruleResult.counts.tested &&
              ((ruleResult.status = "not-applicable"),
              ruleResult.messages.push(
                "No eligible records or sequence groups could be tested.",
              )),
            ruleResult
          );
        };
      if (hasMissingMapping || !keyColumnsInRange || calculationError)
        return (
          (ruleResult.status = "failed"),
          ruleResult.messages.push(
            calculationError ||
              "Missing field mapping: select all required columns inside the source range.",
          ),
          finalizeRule()
        );
      if (rule.kind === "holiday" && !rule.dates.length)
        return (
          (ruleResult.status = "not-applicable"),
          ruleResult.messages.push(
            "No holiday dates supplied. A calendar name alone is not a holiday test.",
          ),
          finalizeRule()
        );
      const markRowInvalid = (invalidSourceRow, invalidReason) => {
          const dispositionEntry =
            dispositionsBySourceRow.get(invalidSourceRow);
          ((dispositionEntry.disposition = "invalid"),
            dispositionEntry.reasons.push(invalidReason),
            (ruleResult.status = "failed"));
        },
        eligibleRows = rows.filter((row) => !row.exclusionReason),
        matchGroups = new Map(),
        addToMatchGroup = (groupKey, groupKeyValues, memberSourceRow) => {
          let existingGroup = matchGroups.get(groupKey);
          (existingGroup ||
            ((existingGroup = { key: groupKeyValues, members: [] }),
            matchGroups.set(groupKey, existingGroup)),
            existingGroup.members.push(memberSourceRow));
        },
        sequenceValueByRow = new Map(),
        holidayDateSet = new Set(rule.kind === "holiday" ? rule.dates : []);
      for (const eligibleRow of eligibleRows) {
        const rowDisposition = dispositionsBySourceRow.get(
            eligibleRow.sourceRow,
          ),
          valueInColumn = (sourceColumn) =>
            eligibleRow.values[sourceColumn - rangeAddress.firstCol];
        if (rule.kind === "duplicates") {
          const keyValues = rule.keyColumns.map(valueInColumn),
            rowMatchKey = buildRowMatchKey(keyValues, rule);
          if (!rowMatchKey) {
            rule.blankKeys === "exclude" &&
            keyValues.some(
              (keyValue) =>
                isNullOrEmptyValue(keyValue) ||
                (typeof keyValue == "string" &&
                  rule.trimText &&
                  !keyValue.trim()),
            )
              ? ((rowDisposition.disposition = "excluded"),
                rowDisposition.reasons.push(
                  "Blank duplicate key excluded under this rule policy.",
                ))
              : markRowInvalid(
                  eligibleRow.sourceRow,
                  "Duplicate key is blank or has an invalid value.",
                );
            continue;
          }
          ((rowDisposition.disposition = "not-matched"),
            addToMatchGroup(
              rowMatchKey.key,
              rowMatchKey.values,
              eligibleRow.sourceRow,
            ));
        } else if (rule.kind === "sequence") {
          const sequenceCellValue = valueInColumn(rule.column),
            groupMatchKey = rule.groupColumns.length
              ? buildRowMatchKey(rule.groupColumns.map(valueInColumn), rule)
              : { key: "[]", values: [] };
          let sequenceNumber = null;
          if (
            (typeof sequenceCellValue == "number" &&
              Number.isSafeInteger(sequenceCellValue) &&
              sequenceCellValue >= 0 &&
              sequenceCellValue <= gf.sequenceMaximum &&
              !rule.prefix &&
              !rule.suffix &&
              (sequenceNumber = sequenceCellValue),
            typeof sequenceCellValue == "string" &&
              sequenceCellValue.startsWith(rule.prefix) &&
              sequenceCellValue.endsWith(rule.suffix))
          ) {
            const digitText = sequenceCellValue.slice(
              rule.prefix.length,
              rule.suffix
                ? sequenceCellValue.length - rule.suffix.length
                : undefined,
            );
            /^\d{1,15}$/.test(digitText) &&
              (sequenceNumber = Number(digitText));
          }
          if (sequenceNumber === null || !groupMatchKey) {
            markRowInvalid(
              eligibleRow.sourceRow,
              "Sequence requires a nonnegative integer or matching prefix/digit/suffix text, and nonblank grouping keys.",
            );
            continue;
          }
          if (
            (rule.start !== null && sequenceNumber < rule.start) ||
            (rule.end !== null && sequenceNumber > rule.end)
          ) {
            markRowInvalid(
              eligibleRow.sourceRow,
              "Sequence value is outside the declared start/end bounds.",
            );
            continue;
          }
          (sequenceValueByRow.set(eligibleRow.sourceRow, sequenceNumber),
            addToMatchGroup(
              groupMatchKey.key,
              groupMatchKey.values,
              eligibleRow.sourceRow,
            ));
        } else if (rule.kind === "round-amount") {
          const amount = valueInColumn(rule.column),
            scaledAmount =
              typeof amount == "number"
                ? toScaledIntegerAmount(amount, rule.precision)
                : null;
          if (scaledAmount === null) {
            markRowInvalid(
              eligibleRow.sourceRow,
              "Amount must be a finite number of at most 1e15 magnitude, exactly representable at the declared decimal precision; numeric text is not coerced.",
            );
            continue;
          }
          if (
            ((rowDisposition.disposition = "not-matched"),
            scaledAmount === 0n && !rule.includeZero)
          )
            continue;
          const scaledIntegerAmount = toScaledIntegerAmount(
              rule.increment,
              rule.precision,
            ),
            scaledTolerance = toScaledIntegerAmount(
              rule.tolerance,
              rule.precision,
            ),
            remainder = scaledAmount % scaledIntegerAmount;
          (remainder < scaledIntegerAmount - remainder
            ? remainder
            : scaledIntegerAmount - remainder) <= scaledTolerance &&
            addToMatchGroup(
              JSON.stringify(amount),
              [amount],
              eligibleRow.sourceRow,
            );
        } else {
          const isoDate = convertSerialToIsoDate(
            valueInColumn(rule.column),
            rule.dateSystem,
          );
          if (!isoDate) {
            markRowInvalid(
              eligibleRow.sourceRow,
              "Date must be strict ISO YYYY-MM-DD or a valid Excel date/time serial in the chosen system; serial 60 in the 1900 system is invalid.",
            );
            continue;
          }
          rowDisposition.disposition = "not-matched";
          const dayOfWeek = new Date(
            parseIsoCalendarDate(isoDate) * 864e5,
          ).getUTCDay();
          (rule.kind === "weekend"
            ? rule.days.includes(dayOfWeek)
            : holidayDateSet.has(isoDate)) &&
            addToMatchGroup(isoDate, [isoDate], eligibleRow.sourceRow);
        }
      }
      if (ruleResult.status === "failed")
        return (
          ruleResult.messages.push(
            "Invalid field values prevent this rule from completing. Inspect the invalid-row ledger; other rows were not treated as a clean population.",
          ),
          finalizeRule()
        );
      const addReviewFinding = (
        findingGroupKey,
        memberRows,
        citedRows,
        title,
        detailText,
        missing,
      ) => {
        const findingId = `${rule.id}-G${ruleResult.findings.length + 1}`,
          reviewFinding = {
            id: findingId,
            classification: "review-candidate",
            title: title,
            detail:
              detailText +
              " Review candidate only; this is not proof of error or fraud.",
            rule: { id: rule.id, kind: rule.kind, version: "1.0.0" },
            parameters: cloneStructuredValue(rule),
            source: cloneStructuredValue(source),
            sourceCells: sourceCellsFor(citedRows, keyColumns),
            group: {
              id: findingId,
              key: findingGroupKey,
              memberRows: [...memberRows],
            },
            counts: createEmptyCounts(rows.length),
            ...(missing ? { missing: missing } : {}),
          };
        if ((ruleResult.findings.push(reviewFinding), rule.kind !== "sequence"))
          for (const memberRow of memberRows)
            ((dispositionsBySourceRow.get(memberRow).disposition = "matched"),
              dispositionsBySourceRow
                .get(memberRow)
                .findingIds.push(findingId));
        else
          for (const citedRow of citedRows)
            dispositionsBySourceRow.get(citedRow).findingIds.push(findingId);
      };
      if (rule.kind === "sequence") {
        let groupsWithoutBounds = 0;
        for (const matchGroup of matchGroups.values()) {
          const membersBySequenceNumber = new Map();
          for (const member of matchGroup.members) {
            const memberSequenceNumber = sequenceValueByRow.get(member);
            membersBySequenceNumber.set(memberSequenceNumber, [
              ...(membersBySequenceNumber.get(memberSequenceNumber) ?? []),
              member,
            ]);
          }
          const sortedSequenceValues = [...membersBySequenceNumber.keys()].sort(
              (leftValue, rightValue) => leftValue - rightValue,
            ),
            lowerBound = rule.start ?? sortedSequenceValues[0],
            upperBound = rule.end ?? sortedSequenceValues.at(-1);
          if (
            sortedSequenceValues.some(
              (sequenceValue) => (sequenceValue - lowerBound) % rule.step !== 0,
            ) ||
            (upperBound - lowerBound) % rule.step !== 0
          ) {
            for (const member of matchGroup.members)
              markRowInvalid(
                member,
                "Sequence values or bounds do not align with the declared step.",
              );
            continue;
          }
          if (
            sortedSequenceValues.length < 2 &&
            (rule.start === null || rule.end === null)
          ) {
            groupsWithoutBounds++;
            for (const member of matchGroup.members)
              dispositionsBySourceRow
                .get(member)
                .reasons.push(
                  "One distinct sequence value without both bounds cannot establish missing values.",
                );
            continue;
          }
          for (const member of matchGroup.members)
            dispositionsBySourceRow.get(member).disposition = "not-matched";
          let previousSequenceValue = null,
            expectedValue = lowerBound;
          const addMissingInterval = (
            first,
            last,
            lowerBoundValue,
            upperBoundValue,
          ) => {
            if (first > last) return;
            const count = (last - first) / rule.step + 1,
              boundaryMembers = [
                ...(lowerBoundValue === null
                  ? []
                  : (membersBySequenceNumber.get(lowerBoundValue) ?? [])),
                ...(upperBoundValue === null
                  ? []
                  : (membersBySequenceNumber.get(upperBoundValue) ?? [])),
              ],
              examples = Array.from(
                { length: Math.min(20, count) },
                (ae, exampleIndex) =>
                  rule.prefix +
                  String(first + exampleIndex * rule.step) +
                  rule.suffix,
              );
            addReviewFinding(
              matchGroup.key,
              matchGroup.members,
              boundaryMembers,
              `Missing sequence ${first === last ? first : `${first}–${last}`} (${count})`,
              "Missing interval under the declared bounds/step. Cited existing cells are boundary evidence, not erroneous rows.",
              {
                first: first,
                last: last,
                step: rule.step,
                count: count,
                examples: examples,
                examplesTruncated: count > 20,
              },
            );
          };
          for (const distinctValue of sortedSequenceValues)
            (addMissingInterval(
              expectedValue,
              distinctValue - rule.step,
              previousSequenceValue,
              distinctValue,
            ),
              (previousSequenceValue = distinctValue),
              (expectedValue = distinctValue + rule.step));
          addMissingInterval(
            expectedValue,
            upperBound,
            previousSequenceValue,
            null,
          );
        }
        (groupsWithoutBounds &&
          ruleResult.messages.push(
            `${groupsWithoutBounds} group(s) lack enough distinct values/bounds; their rows are not tested. No unseen groups or leading/trailing gaps are inferred without bounds.`,
          ),
          ruleResult.messages.push(
            "Gaps are missing intervals, never fabricated source rows. Duplicate sequence values are collapsed for gap detection; add a duplicate rule to review multiplicity.",
          ));
      } else
        for (const matchGroupEntry of matchGroups.values())
          (rule.kind === "duplicates" && matchGroupEntry.members.length < 2) ||
            addReviewFinding(
              matchGroupEntry.key,
              matchGroupEntry.members,
              matchGroupEntry.members,
              `${rule.label}: ${matchGroupEntry.members.length} row(s)`,
              rule.kind === "duplicates"
                ? "All members of the typed matching-key group are included."
                : rule.kind === "round-amount"
                  ? "Amount meets the declared increment, decimal precision, tolerance and zero policy."
                  : `Posting calendar date ${matchGroupEntry.key[0]} matches the supplied ${rule.kind === "holiday" ? "holiday calendar" : "weekend weekdays"}.`,
            );
      return finalizeRule();
    }),
    combination = {
      status: "complete",
      messages: [],
      counts: createEmptyCounts(rows.length),
      rows: createRowDispositions(),
      findings: [],
    },
    selectedResults = results.filter((result) =>
      patternCheckSettings.combination.ruleIds.includes(result.rule.id),
    );
  if (!selectedResults.length)
    ((combination.status = "not-applicable"),
      combination.messages.push(
        "No row combination configured. Sequence gaps remain independent interval findings.",
      ));
  else if (
    selectedResults.some(
      (selectedResult) => selectedResult.status !== "complete",
    )
  )
    ((combination.status = selectedResults.some(
      (selectedResult) => selectedResult.status === "failed",
    )
      ? "failed"
      : "not-applicable"),
      combination.messages.push(
        "Every selected rule must complete before a combined result is available.",
      ));
  else
    for (const row of combination.rows) {
      if (row.disposition === "excluded") continue;
      const ruleRowPairs = selectedResults.map((selectedResult) => ({
        result: selectedResult,
        row: selectedResult.rows.find(
          (resultRow) => resultRow.sourceRow === row.sourceRow,
        ),
      }));
      if (
        ruleRowPairs.some(
          (pair) => !["matched", "not-matched"].includes(pair.row.disposition),
        )
      ) {
        row.reasons.push(
          "Combination uses only rows tested by every selected rule (complete cases).",
        );
        continue;
      }
      const matchedPairs = ruleRowPairs.filter(
          (pair) => pair.row.disposition === "matched",
        ),
        isCombinationMatch =
          patternCheckSettings.combination.mode === "any"
            ? matchedPairs.length > 0
            : matchedPairs.length === selectedResults.length;
      if (
        ((row.disposition = isCombinationMatch ? "matched" : "not-matched"),
        !isCombinationMatch)
      )
        continue;
      const combinedFindingId = `combined-R${row.sourceRow}`;
      row.findingIds.push(combinedFindingId);
      const combinedKeyColumns = [
        ...new Set(
          selectedResults.flatMap((selectedResult) =>
            ruleKeyColumns(selectedResult.rule),
          ),
        ),
      ];
      combination.findings.push({
        id: combinedFindingId,
        classification: "review-candidate",
        title: `Combined review candidate — row ${row.sourceRow}`,
        detail: `${patternCheckSettings.combination.mode === "all" ? "All" : "At least one"} selected row rule matched. Review candidate only; not proof of error or fraud.`,
        rule: {
          id: "combined",
          kind: patternCheckSettings.combination.mode,
          version: "1.0.0",
        },
        parameters: cloneStructuredValue(patternCheckSettings.combination),
        source: cloneStructuredValue(source),
        sourceCells: sourceCellsFor([row.sourceRow], combinedKeyColumns),
        group: {
          id: combinedFindingId,
          key: [row.sourceRow],
          memberRows: [row.sourceRow],
        },
        counts: createEmptyCounts(rows.length),
        matchedRules: matchedPairs.map((matchedPair) => ({
          ruleId: matchedPair.result.rule.id,
          findingIds: [...matchedPair.row.findingIds],
        })),
      });
    }
  ((combination.counts = summarizeRuleOutcome(
    combination.rows,
    combination.findings,
  )),
    combination.status === "complete" &&
      !combination.counts.tested &&
      ((combination.status = "not-applicable"),
      combination.messages.push(
        "No row was eligible for every selected check.",
      )));
  for (const finding of combination.findings)
    finding.counts = cloneStructuredValue(combination.counts);
  const counts = {
    population: rows.length,
    excluded: rows.filter((row) => row.exclusionReason).length,
    eligible: rows.filter((row) => !row.exclusionReason).length,
    rulesComplete: results.filter((result) => result.status === "complete")
      .length,
    rulesFailed: results.filter((result) => result.status === "failed").length,
    rulesNotApplicable: results.filter(
      (result) => result.status === "not-applicable",
    ).length,
    findings: results.reduce(
      (total, ruleResult) => total + ruleResult.findings.length,
      0,
    ),
    gapValues: results.reduce(
      (total, ruleResult) => total + ruleResult.counts.gapValues,
      0,
    ),
  };
  Number.isSafeInteger(counts.gapValues) ||
    throwErrorMessage(
      "The aggregate missing-value count exceeds the exact integer reporting limit. Run fewer sequence rules.",
    );
  const report = {
    schemaVersion: 1,
    kind: "pattern-audit",
    algorithm: "ledgerglass-patterns-1",
    generatedAt: new Date().toISOString(),
    request: patternCheckSettings,
    context: materialitySourceEvidence.context,
    source: source,
    status: counts.rulesFailed
      ? "failed"
      : counts.rulesComplete
        ? "complete"
        : "not-applicable",
    messages: calculationError
      ? [
          "Calculation failed. Evidence contains raw fields; no rule conclusion is available.",
          calculationError,
        ]
      : [],
    headers: materialitySourceEvidence.headers,
    rows: rows,
    counts: counts,
    results: results,
    combination: combination,
    notes: [
      "Review candidates are not proof of error or fraud. No cells are edited, records removed or automatic totals excluded.",
      "Population includes every selected data-row position. Only the visible blank-row/explicit-row and rule blank-key policies exclude records. Hidden/filter content is included only when explicitly authorized; it is never silently skipped.",
      "Typed keys keep text and numbers distinct. Numeric text is not coerced for amount tests. Round amounts use exact decimal units of the stored number; no currency conversion, floating epsilon or silent rounding is applied.",
      "Dates use calendar days, independent of browser timezone. Fractional Excel serials retain their posting calendar day; ambiguous locale/timestamp strings must be converted explicitly. Holiday calendars are user supplied, not a current official jurisdiction calendar.",
      "Calculated mode validates mapped formula fields across the whole selected evidence, including explicitly excluded rows. Formula errors or unsupported dependency closure fail the run; select raw mode only when intentionally testing stored expressions. Sequence examples do not infer zero-padding width.",
      "Row combinations use the intersection of evaluable rows. Sequence gaps are independent intervals with observed boundary citations; examples may be abbreviated, but counts and full intervals are complete.",
      "Reports are current-snapshot evidence, not an immutable audit service. Store original source and the exported report under host access controls; native workbook JSON does not automatically retain a separately downloaded report.",
    ],
  };
  return (
    new TextEncoder().encode(JSON.stringify(report, null, 2)).length >
      gf.reportBytes &&
      throwErrorMessage(
        "The complete pattern report exceeds 2,000,000 bytes. Select fewer cells/rules; results are never silently truncated.",
      ),
    KW.set(report, {
      snapshot: JSON.stringify(workbook),
      request: JSON.stringify(patternCheckSettings),
      context: JSON.stringify(materialitySourceEvidence.context),
      content: JSON.stringify(report),
    }),
    report
  );
}
function checkReportCurrency(
  workbook,
  report,
  patternCheckRequest,
  evidenceOptions = {},
) {
  const storedEvidence = KW.get(report),
    notCurrentResult = (reason) => ({
      current: false,
      disclose: false,
      reason: reason,
    });
  if (!storedEvidence || storedEvidence.content !== JSON.stringify(report))
    return notCurrentResult(
      "This report is untrusted or changed. Run the pattern checks again.",
    );
  try {
    const patternCheckSettings = normalizePatternCheckSettings(
        workbook,
        patternCheckRequest,
      ),
      materialitySourceEvidence = buildMaterialitySourceEvidence(
        workbook,
        {
          sheetId: patternCheckSettings.sheetId,
          range: patternCheckSettings.range,
          headers: patternCheckSettings.headers,
          columns: [],
        },
        evidenceOptions,
      );
    return storedEvidence.snapshot !== JSON.stringify(workbook) ||
      storedEvidence.context !==
        JSON.stringify(materialitySourceEvidence.context) ||
      patternCheckSettings.sheetId !== report.request.sheetId ||
      patternCheckSettings.range !== report.request.range
      ? notCurrentResult(
          "Source or authorized scope changed. Previous evidence is hidden; run checks again.",
        )
      : storedEvidence.request !== JSON.stringify(patternCheckSettings)
        ? {
            current: false,
            disclose: true,
            reason:
              "Settings changed. This is a previous result; run checks again before export.",
          }
        : {
            current: true,
            disclose: true,
            reason: "Current source and rule configuration.",
          };
  } catch {
    return notCurrentResult(
      "Source, settings or authorization are no longer valid. Run checks again.",
    );
  }
}
function exportScreenReportJson(
  workbook,
  report,
  patternCheckRequest,
  evidenceOptions = {},
) {
  const reportCurrency = checkReportCurrency(
    workbook,
    report,
    patternCheckRequest,
    evidenceOptions,
  );
  return (
    reportCurrency.current || throwErrorMessage(reportCurrency.reason),
    JSON.stringify(report, null, 2)
  );
}
function parseScreenReportRequest(workbook, reportInput) {
  const screenReport = reportInput;
  return (
    (!screenReport ||
      screenReport.schemaVersion !== 1 ||
      screenReport.kind !== "pattern-audit") &&
      throwErrorMessage("Choose a pattern-audit configuration or report."),
    normalizePatternCheckSettings(workbook, screenReport.request)
  );
}
const screenRuleLabels = {
    duplicates: "Duplicate keys",
    sequence: "Sequence gaps",
    "round-amount": "Round amounts",
    weekend: "Weekend postings",
    holiday: "Holiday postings",
  },
  weekdayNames = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ],
  buildScreenRule = (kind, ruleId) => {
    const baseRule = { id: ruleId, label: screenRuleLabels[kind] };
    switch (kind) {
      case "duplicates":
        return {
          ...baseRule,
          kind: kind,
          keyColumns: [],
          caseSensitive: false,
          trimText: true,
          blankKeys: "fail",
        };
      case "sequence":
        return {
          ...baseRule,
          kind: kind,
          column: null,
          groupColumns: [],
          caseSensitive: false,
          trimText: true,
          prefix: "",
          suffix: "",
          step: 1,
          start: null,
          end: null,
        };
      case "round-amount":
        return {
          ...baseRule,
          kind: kind,
          column: null,
          increment: 100,
          precision: 2,
          tolerance: 0,
          includeZero: false,
        };
      case "weekend":
        return {
          ...baseRule,
          kind: kind,
          column: null,
          dateSystem: "1900",
          days: [0, 6],
        };
      case "holiday":
        return {
          ...baseRule,
          kind: kind,
          column: null,
          dateSystem: "1900",
          calendarName: "User-supplied calendar",
          calendarVersion: "1",
          dates: [],
        };
    }
  },
  formatScreenCountsSummary = (counts) =>
    `${counts.population} population · ${counts.excluded} excluded · ${counts.invalid} invalid · ${counts.tested} tested · ${counts.matched} matched · ${counts.notMatched} not matched · ${counts.notTested} not tested · ${counts.findings} findings · ${counts.gapValues} missing sequence values`,
  nanToEmptyText = (value) =>
    value === null || Number.isNaN(value) ? "" : value;

export {
  formatScreenCountsSummary,
  exportScreenReportJson,
  nanToEmptyText,
  MaterialityRiskScreen,
  runPatternChecks,
  weekdayNames,
  checkReportCurrency,
  buildScreenRule,
  gf,
  normalizePatternCheckSettings,
  screenRuleLabels,
  parseScreenReportRequest,
};
