import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { defineDataPropertyByKey } from "./01-schemas-date-functions-and-validation.js";
import { automationLimits, columnIndexToLetters } from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  localizeDecimalPoint,
} from "./04-cell-input-paste-and-fill.js";
import { importDelimitedFile } from "./07-file-import-and-worker-dispatch.js";
import {
  delimitedImportDefaults,
  exportWorksheetToCsv,
  getWorksheetUsedRange,
} from "./08-xlsx-import-export-and-preservation.js";
import { useOperationScope } from "./09-workbook-links.js";
import {
  buildWhatIfPreview,
  computeWhatIfResult,
  applyWhatIfScenario,
  saveWhatIfAnalysis,
  findSavedWhatIfAnalysis,
  buildWhatIfReportRows,
  isWhatIfAnalysisInScope,
  parseAssumptionNumbers,
} from "./39-report-outputs-and-what-if.js";

function WhatIfAnalysisDialog(props) {
  const { text, language, entryLocale } = useWorkbookMessages(),
    numericLocale = ne.useRef(entryLocale).current,
    isSemicolonLocale = numericLocale === "de-DE" || numericLocale === "fr-FR",
    formatAssumptionNumbers = (numbers) =>
      numbers
        .map((number) => localizeDecimalPoint(number, numericLocale))
        .join(isSemicolonLocale ? "; " : ", "),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    [analysisKind, setAnalysisKind] = ne.useState(
      props.initialKind ?? "scenarios",
    ),
    [analysisName, setAnalysisName] = ne.useState("Sensitivity analysis"),
    [sheetId, setSheetId] = ne.useState(props.sheetId),
    [inputCellsText, setInputCellsText] = ne.useState(""),
    [outputCellsText, setOutputCellsText] = ne.useState(props.address),
    [orientation, setOrientation] = ne.useState("column"),
    [primaryValuesText, setPrimaryValuesText] = ne.useState(""),
    [columnValuesText, setColumnValuesText] = ne.useState(""),
    [scenarios, setScenarios] = ne.useState([
      { name: "Downside", values: "" },
      { name: "Upside", values: "" },
    ]),
    [checked, setChecked] = ne.useState(false),
    [selectedSavedId, setSelectedSavedId] = ne.useState(""),
    [selectedScenarioIndex, setSelectedScenarioIndex] = ne.useState("0"),
    [preview, setPreview] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    scopeOptions = {
      readOnly: props.readOnly,
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: props.allowedSheetIds,
    },
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    scopedSheets = props.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        (scopeOptions.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    scopeKey = JSON.stringify([allowedSheetIds, scopeOptions.includeHidden]),
    previousScopeKey = ne.useRef(scopeKey),
    X$1 =
      !!preview &&
      preview.snapshot === JSON.stringify(props.workbook) &&
      preview.context.includeHidden === scopeOptions.includeHidden &&
      JSON.stringify(preview.context.allowedSheetIds) ===
        JSON.stringify(allowedSheetIds),
    savedAnalyses = (props.workbook.whatIfAnalyses ?? []).filter((analysis) =>
      isWhatIfAnalysisInScope(props.workbook, analysis, scopeOptions),
    );
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Analysis name"]')
        ?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      previousScopeKey.current !== scopeKey &&
        (setPreview(null),
        setErrorMessage(""),
        setSelectedSavedId(""),
        setInputCellsText(""),
        setOutputCellsText(""),
        setPrimaryValuesText(""),
        setColumnValuesText(""),
        setScenarios([{ name: "Scenario 1", values: "" }]),
        (previousScopeKey.current = scopeKey));
    }, [scopeKey]));
  const clearPreviewAfter = (applyChange) => {
      (applyChange(), setPreview(null), setErrorMessage(""));
    },
    runGuarded = (action) => {
      try {
        (action(), setErrorMessage(""));
      } catch (error) {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "What-if analysis could not complete.",
        );
      }
    },
    parseCellAddresses = (addressText) => {
      const addressParts =
        addressText.match(/'(?:[^']|'')*'![^,\n]+|[^,\n]+/g) ?? [];
      if (!addressParts.length || /[,\n]\s*[,\n]|[,\n]\s*$/.test(addressText))
        throw Error("Enter cell addresses separated by commas or new lines.");
      return addressParts.map((addressPart) => {
        const address = addressPart.trim(),
          separatorIndex = address.lastIndexOf("!");
        if (separatorIndex < 0) return { sheetId: sheetId, address: address };
        const sheetName = address
            .slice(0, separatorIndex)
            .replace(/^'(.*)'$/, "$1")
            .replaceAll("''", "'"),
          referencedSheet = scopedSheets.find(
            (sheet) => sheet.name.toLowerCase() === sheetName.toLowerCase(),
          );
        if (!referencedSheet)
          throw Error(
            "A referenced worksheet is missing or outside the current context.",
          );
        return {
          sheetId: referencedSheet.id,
          address: address.slice(separatorIndex + 1),
        };
      });
    },
    onClick = () =>
      runGuarded(() => {
        setPreview(null);
        const requestBase = {
            name: analysisName,
            inputs: parseCellAddresses(inputCellsText),
            outputs: parseCellAddresses(outputCellsText),
          },
          request =
            analysisKind === "scenarios"
              ? {
                  ...requestBase,
                  kind: analysisKind,
                  scenarios: scenarios.map((scenario) => ({
                    name: scenario.name,
                    values: parseAssumptionNumbers(
                      scenario.values,
                      numericLocale,
                    ),
                  })),
                }
              : analysisKind === "one-variable"
                ? {
                    ...requestBase,
                    kind: analysisKind,
                    orientation: orientation,
                    values: parseAssumptionNumbers(
                      primaryValuesText,
                      numericLocale,
                    ),
                  }
                : {
                    ...requestBase,
                    kind: analysisKind,
                    rowValues: parseAssumptionNumbers(
                      primaryValuesText,
                      numericLocale,
                    ),
                    columnValues: parseAssumptionNumbers(
                      columnValuesText,
                      numericLocale,
                    ),
                  };
        (setPreview(computeWhatIfResult(props.workbook, request, scopeOptions)),
          setSelectedScenarioIndex("0"));
      }),
    loadSavedAnalysis = (analysisId, isBaselineReview = false) =>
      runGuarded(() => {
        const savedAnalysis = savedAnalyses.find(
          (analysis) => analysis.id === analysisId,
        );
        if ((setPreview(null), setSelectedSavedId(analysisId), !savedAnalysis))
          return;
        if (
          savedAnalysis.assumptions.some(
            (assumption) =>
              !props.workbook.sheets.some(
                (sheet) => sheet.id === assumption.sheetId,
              ),
          )
        )
          throw Error(
            "A saved source worksheet was deleted. This definition cannot reconnect by name. Remove the definition or create a new analysis with current source cells; its historical report is retained.",
          );
        const request = savedAnalysis.request,
          formatCellReferences = (references) =>
            references
              .map(
                (reference) =>
                  `'${props.workbook.sheets.find((sheet) => sheet.id === reference.sheetId).name.replaceAll("'", "''")}'!${reference.address}`,
              )
              .join(", ");
        (setAnalysisName(request.name),
          setInputCellsText(formatCellReferences(request.inputs)),
          setOutputCellsText(formatCellReferences(request.outputs)),
          setAnalysisKind(isBaselineReview ? "scenarios" : request.kind),
          isBaselineReview
            ? setScenarios([
                {
                  name: "Restore saved baseline",
                  values: formatAssumptionNumbers(
                    savedAnalysis.baseline.inputs,
                  ),
                },
              ])
            : request.kind === "scenarios"
              ? setScenarios(
                  request.scenarios.map((scenario) => ({
                    name: scenario.name,
                    values: formatAssumptionNumbers(scenario.values),
                  })),
                )
              : request.kind === "one-variable"
                ? (setPrimaryValuesText(
                    formatAssumptionNumbers(request.values),
                  ),
                  setOrientation(request.orientation))
                : (setPrimaryValuesText(
                    formatAssumptionNumbers(request.rowValues),
                  ),
                  setColumnValuesText(
                    formatAssumptionNumbers(request.columnValues),
                  )));
      }),
    saveReport = () =>
      runGuarded(() => {
        preview &&
          (props.onWorkbookChange(
            saveWhatIfAnalysis(
              props.workbook,
              preview,
              scopeOptions,
              selectedSavedId || undefined,
            ),
            `Save what-if report: ${preview.request.name}`,
          ),
          setPreview(null));
      }),
    applyScenario = () =>
      runGuarded(() => {
        if (!preview) return;
        const updatedWorkbook = applyWhatIfScenario(
          props.workbook,
          preview,
          Number(selectedScenarioIndex),
          scopeOptions,
        );
        (updatedWorkbook !== props.workbook &&
          props.onWorkbookChange(
            updatedWorkbook,
            `Apply scenario: ${preview.trials[Number(selectedScenarioIndex)].name}`,
          ),
          setPreview(null));
      }),
    exportEvidence = () =>
      runGuarded(() => {
        if (!preview) return;
        const previewJson = JSON.stringify(
          buildWhatIfPreview(props.workbook, preview, scopeOptions),
          null,
          2,
        );
        if (props.onExport) props.onExport(previewJson);
        else {
          const blobUrl = URL.createObjectURL(
              new Blob([previewJson], { type: "application/json" }),
            ),
            downloadLink = document.createElement("a");
          ((downloadLink.href = blobUrl),
            (downloadLink.download = "what-if-analysis.json"),
            downloadLink.click(),
            setTimeout(() => URL.revokeObjectURL(blobUrl), 1e3));
        }
      }),
    reportRows = X$1 && preview ? buildWhatIfReportRows(preview) : [],
    sourceAssumptionsIndex = reportRows.findIndex(
      (row) => row[0] === "Source assumptions at preview",
    );
  return i.jsx("div", {
    className: "gs-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "gs-dialog wi-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        var lastElement, firstElement;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
            ) ?? []),
          ].filter(
            (element) =>
              !element.hidden &&
              (!element.closest("details:not([open])") ||
                element.tagName === "SUMMARY"),
          );
          event.shiftKey && document.activeElement === focusableElements[0]
            ? (event.preventDefault(),
              (lastElement = focusableElements.at(-1)) == null ||
                lastElement.focus())
            : !event.shiftKey &&
              document.activeElement === focusableElements.at(-1) &&
              (event.preventDefault(),
              (firstElement = focusableElements[0]) == null ||
                firstElement.focus());
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
                    children: "What-if Analysis",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Compare assumptions for estimates and valuation sensitivity.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close What-if Analysis"),
              onClick: props.onClose,
              lang: language("Close What-if Analysis"),
              "data-ui-label": "Close What-if Analysis",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "gs-fields",
          children: [
            i.jsx("p", {
              role: "note",
              children: text(
                isSemicolonLocale
                  ? "Separate numeric assumptions with semicolons or new lines. Use a comma for decimals."
                  : "Separate numeric assumptions with commas, semicolons or new lines.",
              ),
            }),
            entryLocale !== numericLocale &&
              i.jsx("p", {
                role: "status",
                children: text(
                  "The current edit keeps its original numeric locale until you apply or cancel it.",
                ),
              }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Analysis name" }),
                i.jsx("input", {
                  "aria-label": text("Analysis name"),
                  value: analysisName,
                  maxLength: 100,
                  onChange: (event) =>
                    clearPreviewAfter(() =>
                      setAnalysisName(event.target.value),
                    ),
                  lang: language("Analysis name"),
                  "data-ui-label": "Analysis name",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Analysis type" }),
                i.jsxs("select", {
                  "aria-label": text("Analysis type"),
                  value: analysisKind,
                  onChange: (event) =>
                    clearPreviewAfter(() => {
                      (setAnalysisKind(event.target.value),
                        setSelectedSavedId(""));
                    }),
                  lang: language("Analysis type"),
                  "data-ui-label": "Analysis type",
                  children: [
                    i.jsx("option", {
                      value: "scenarios",
                      children: "Named scenarios",
                    }),
                    i.jsx("option", {
                      value: "one-variable",
                      children: "One-variable data table",
                    }),
                    i.jsx("option", {
                      value: "two-variable",
                      children: "Two-variable data table",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Default worksheet" }),
                i.jsxs("select", {
                  "aria-label": text("Default worksheet"),
                  value: scopedSheets.some((sheet) => sheet.id === sheetId)
                    ? sheetId
                    : "",
                  onChange: (event) =>
                    clearPreviewAfter(() => setSheetId(event.target.value)),
                  lang: language("Default worksheet"),
                  "data-ui-label": "Default worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    scopedSheets.map((sheet) =>
                      i.jsx(
                        "option",
                        { value: sheet.id, children: sheet.name },
                        sheet.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Saved analyses" }),
                i.jsxs("select", {
                  "aria-label": text("Saved analyses"),
                  value: savedAnalyses.some(
                    (analysis) => analysis.id === selectedSavedId,
                  )
                    ? selectedSavedId
                    : "",
                  onChange: (event) => loadSavedAnalysis(event.target.value),
                  lang: language("Saved analyses"),
                  "data-ui-label": "Saved analyses",
                  children: [
                    i.jsx("option", { value: "", children: "New analysis" }),
                    savedAnalyses.map((analysis) =>
                      i.jsx(
                        "option",
                        { value: analysis.id, children: analysis.request.name },
                        analysis.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Changing cells" }),
                i.jsx("input", {
                  "aria-label": text("Changing cells"),
                  placeholder:
                    analysisKind === "two-variable"
                      ? "B2, B3 — across first, down second"
                      : "B2, B3",
                  value: inputCellsText,
                  onChange: (event) =>
                    clearPreviewAfter(() =>
                      setInputCellsText(event.target.value),
                    ),
                  lang: language("Changing cells"),
                  "data-ui-label": "Changing cells",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Result cells" }),
                i.jsx("input", {
                  "aria-label": text("Result cells"),
                  placeholder: text("B5, B6"),
                  value: outputCellsText,
                  onChange: (event) =>
                    clearPreviewAfter(() =>
                      setOutputCellsText(event.target.value),
                    ),
                  lang: language("Result cells"),
                  "data-ui-label": "Result cells",
                }),
              ],
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Use commas between cells. Cross-sheet references may use 'Sheet name'!B2. Result cells must contain formulas.",
          }),
        }),
        selectedSavedId &&
          i.jsxs("div", {
            className: "wi-actions",
            children: [
              i.jsx("button", {
                onClick: () => loadSavedAnalysis(selectedSavedId, true),
                children: i.jsx(LocalizedText, {
                  children: "Review saved baseline",
                }),
              }),
              i.jsx("button", {
                disabled: props.readOnly,
                onClick: () =>
                  runGuarded(() => {
                    (props.onWorkbookChange(
                      findSavedWhatIfAnalysis(
                        props.workbook,
                        selectedSavedId,
                        scopeOptions,
                      ),
                      "Remove saved what-if definition",
                    ),
                      setSelectedSavedId(""),
                      setPreview(null));
                  }),
                children: i.jsx(LocalizedText, {
                  children: "Remove saved definition",
                }),
              }),
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "Saved locations are historical. Check them after moving or deleting source rows; preview again before any action. Removing a definition keeps its report sheets.",
                }),
              }),
            ],
          }),
        analysisKind === "scenarios"
          ? i.jsxs("fieldset", {
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "Named scenarios",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Enter values in the same order as the changing cells. Up to 32 inputs and 100 named scenarios.",
                  }),
                }),
                scenarios.map((scenario, scenarioIndex) =>
                  i.jsxs(
                    "div",
                    {
                      className: "wi-scenario",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Scenario " }),
                            scenarioIndex + 1,
                            " ",
                            i.jsx(LocalizedText, { children: "name" }),
                            i.jsx("input", {
                              "aria-label": `Scenario ${scenarioIndex + 1} name`,
                              value: scenario.name,
                              maxLength: 100,
                              onChange: (event) =>
                                clearPreviewAfter(() =>
                                  setScenarios(
                                    scenarios.map((entry, entryIndex) =>
                                      entryIndex === scenarioIndex
                                        ? { ...entry, name: event.target.value }
                                        : entry,
                                    ),
                                  ),
                                ),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Scenario " }),
                            scenarioIndex + 1,
                            " ",
                            i.jsx(LocalizedText, { children: "values" }),
                            i.jsx("input", {
                              "aria-label": `Scenario ${scenarioIndex + 1} values`,
                              placeholder: isSemicolonLocale
                                ? "90000; 45000"
                                : "90000, 45000",
                              value: scenario.values,
                              onChange: (event) =>
                                clearPreviewAfter(() =>
                                  setScenarios(
                                    scenarios.map((entry, entryIndex) =>
                                      entryIndex === scenarioIndex
                                        ? {
                                            ...entry,
                                            values: event.target.value,
                                          }
                                        : entry,
                                    ),
                                  ),
                                ),
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: scenarios.length === 1,
                          "aria-label": `Remove scenario ${scenarioIndex + 1}`,
                          onClick: () =>
                            clearPreviewAfter(() =>
                              setScenarios(
                                scenarios.filter(
                                  (we, entryIndex) =>
                                    entryIndex !== scenarioIndex,
                                ),
                              ),
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "Remove",
                          }),
                        }),
                      ],
                    },
                    scenarioIndex,
                  ),
                ),
                i.jsx("button", {
                  disabled: scenarios.length >= 100,
                  onClick: () =>
                    clearPreviewAfter(() =>
                      setScenarios([
                        ...scenarios,
                        {
                          name: `Scenario ${scenarios.length + 1}`,
                          values: "",
                        },
                      ]),
                    ),
                  children: i.jsx(LocalizedText, { children: "Add scenario" }),
                }),
              ],
            })
          : i.jsxs("div", {
              className: "gs-fields",
              children: [
                i.jsxs("label", {
                  children: [
                    analysisKind === "two-variable"
                      ? "Across-row input values (first changing cell)"
                      : "Trial input values",
                    i.jsx("textarea", {
                      "aria-label":
                        analysisKind === "two-variable"
                          ? "Across-row input values"
                          : "Trial input values",
                      placeholder: isSemicolonLocale
                        ? "80000; 100000; 120000"
                        : "80000, 100000, 120000",
                      value: primaryValuesText,
                      onChange: (event) =>
                        clearPreviewAfter(() =>
                          setPrimaryValuesText(event.target.value),
                        ),
                    }),
                  ],
                }),
                analysisKind === "two-variable"
                  ? i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children:
                            "Down-column input values (second changing cell)",
                        }),
                        i.jsx("textarea", {
                          "aria-label": text("Down-column input values"),
                          placeholder: isSemicolonLocale
                            ? "40000; 50000; 60000"
                            : "40000, 50000, 60000",
                          value: columnValuesText,
                          onChange: (event) =>
                            clearPreviewAfter(() =>
                              setColumnValuesText(event.target.value),
                            ),
                          lang: language("Down-column input values"),
                          "data-ui-label": "Down-column input values",
                        }),
                      ],
                    })
                  : i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Values arranged" }),
                        i.jsxs("select", {
                          "aria-label": text("Values arranged"),
                          value: orientation,
                          onChange: (event) =>
                            clearPreviewAfter(() =>
                              setOrientation(event.target.value),
                            ),
                          lang: language("Values arranged"),
                          "data-ui-label": "Values arranged",
                          children: [
                            i.jsx("option", {
                              value: "column",
                              children: "Down a column",
                            }),
                            i.jsx("option", {
                              value: "row",
                              children: "Across a row",
                            }),
                          ],
                        }),
                      ],
                    }),
              ],
            }),
        props.allowHidden &&
          i.jsxs("label", {
            className: "wi-check",
            children: [
              i.jsx("input", {
                type: "checkbox",
                checked: checked,
                onChange: (event) =>
                  clearPreviewAfter(() => setChecked(event.target.checked)),
              }),
              i.jsx(LocalizedText, { children: "Include hidden sources" }),
            ],
          }),
        errorMessage &&
          i.jsx("p", {
            role: "alert",
            className: "gs-error",
            children: errorMessage,
          }),
        i.jsx("button", {
          className: "gs-primary",
          onClick: onClick,
          children: i.jsx(LocalizedText, { children: "Preview analysis" }),
        }),
        preview &&
          i.jsxs("section", {
            className: "gs-preview",
            "aria-label": text("What-if results"),
            lang: language("What-if results"),
            "data-ui-label": "What-if results",
            children: [
              X$1
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("h3", {
                        children: i.jsx(LocalizedText, {
                          children: "Review results",
                        }),
                      }),
                      i.jsxs("p", {
                        role: "status",
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Baseline inputs are unchanged.",
                          }),
                          " ",
                          preview.trials.filter((trial) => !trial.error).length,
                          " ",
                          i.jsx(LocalizedText, { children: "of" }),
                          " ",
                          preview.trials.length,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "trials succeeded. ",
                          }),
                          preview.dependencyCells,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "dependency cells; ",
                          }),
                          preview.workUnits,
                          " ",
                          i.jsx(LocalizedText, { children: "work units." }),
                        ],
                      }),
                      preview.baseline.error &&
                        i.jsxs("p", {
                          role: "alert",
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Baseline calculation failed: ",
                            }),
                            preview.baseline.error,
                          ],
                        }),
                      i.jsx("div", {
                        className: "wi-scroll",
                        children: i.jsx("table", {
                          "aria-label": text("Sensitivity results"),
                          lang: language("Sensitivity results"),
                          "data-ui-label": "Sensitivity results",
                          children: i.jsx("tbody", {
                            children: reportRows
                              .slice(4, sourceAssumptionsIndex - 1)
                              .map((row, rowIndex) =>
                                i.jsx(
                                  "tr",
                                  {
                                    children: row.map((cell, columnIndex) =>
                                      columnIndex === 0
                                        ? i.jsx(
                                            "th",
                                            {
                                              scope: "row",
                                              children: String(cell ?? ""),
                                            },
                                            columnIndex,
                                          )
                                        : i.jsx(
                                            "td",
                                            { children: String(cell ?? "") },
                                            columnIndex,
                                          ),
                                    ),
                                  },
                                  rowIndex,
                                ),
                              ),
                          }),
                        }),
                      }),
                      i.jsxs("details", {
                        children: [
                          i.jsx("summary", {
                            children: i.jsx(LocalizedText, {
                              children:
                                "Inspect source assumptions and formulas",
                            }),
                          }),
                          i.jsx("div", {
                            className: "wi-scroll",
                            children: i.jsxs("table", {
                              "aria-label": text("Source assumptions"),
                              lang: language("Source assumptions"),
                              "data-ui-label": "Source assumptions",
                              children: [
                                i.jsx("thead", {
                                  children: i.jsxs("tr", {
                                    children: [
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Source",
                                        }),
                                      }),
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Original value / formula",
                                        }),
                                      }),
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Literal text",
                                        }),
                                      }),
                                    ],
                                  }),
                                }),
                                i.jsx("tbody", {
                                  children: preview.assumptions.map(
                                    (assumption, assumptionIndex) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            i.jsxs("th", {
                                              scope: "row",
                                              children: [
                                                assumption.sheetName,
                                                "!",
                                                assumption.address,
                                              ],
                                            }),
                                            i.jsx("td", {
                                              children: String(assumption.raw),
                                            }),
                                            i.jsx("td", {
                                              children: assumption.literal
                                                ? "Yes"
                                                : "No",
                                            }),
                                          ],
                                        },
                                        assumptionIndex,
                                      ),
                                  ),
                                }),
                              ],
                            }),
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "wi-actions",
                        children: [
                          i.jsx("button", {
                            onClick: exportEvidence,
                            children: i.jsx(LocalizedText, {
                              children: "Export reviewed evidence",
                            }),
                          }),
                          i.jsx("button", {
                            disabled: props.readOnly,
                            onClick: saveReport,
                            children: selectedSavedId
                              ? "Update analysis and create report"
                              : "Save analysis and create report",
                          }),
                        ],
                      }),
                      preview.request.kind === "scenarios" &&
                        i.jsxs("div", {
                          className: "wi-actions",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Scenario to apply",
                                }),
                                i.jsx("select", {
                                  "aria-label": text("Scenario to apply"),
                                  value: selectedScenarioIndex,
                                  onChange: (event) =>
                                    setSelectedScenarioIndex(
                                      event.target.value,
                                    ),
                                  lang: language("Scenario to apply"),
                                  "data-ui-label": "Scenario to apply",
                                  children: preview.trials.map(
                                    (trial, trialIndex) =>
                                      i.jsxs(
                                        "option",
                                        {
                                          value: trialIndex,
                                          disabled: !!trial.error,
                                          children: [
                                            trial.name,
                                            trial.error ? " — failed" : "",
                                          ],
                                        },
                                        trialIndex,
                                      ),
                                  ),
                                }),
                              ],
                            }),
                            i.jsx("button", {
                              className: "gs-primary",
                              disabled:
                                props.readOnly ||
                                !!preview.trials[Number(selectedScenarioIndex)]
                                  ?.error,
                              onClick: applyScenario,
                              children: i.jsx(LocalizedText, {
                                children: "Apply selected scenario",
                              }),
                            }),
                          ],
                        }),
                    ],
                  })
                : i.jsx("p", {
                    role: "status",
                    children: i.jsx(LocalizedText, {
                      children:
                        "This preview is stale. Results and assumptions are withheld; preview again using the current workbook and permissions.",
                    }),
                  }),
              i.jsx("button", {
                onClick: () => setPreview(null),
                children: i.jsx(LocalizedText, { children: "Discard preview" }),
              }),
            ],
          }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Supported scope and limits",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Numeric inputs and deterministic scalar formulas. One-variable tables support up to 16 result formulas; two-variable tables support one. Up to 100 values per axis, 400 combinations and 50,000 dependency-cell work units. Invalid assumptions and failed calculations never change the source.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Reports are snapshots. Refresh by loading a saved definition and previewing again; previous reports remain intact. Native JSON retains definitions and evidence. XLSX retains report cells, without native Excel scenario or data-table definitions. Volatile, dynamic, array, blank and incomplete dependencies are excluded. Multivariable Solver is separate.",
              }),
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
function DelimitedImportExportDialog({
  source,
  blocked,
  onClose,
  onImport,
  onExport,
  onReport,
}) {
  const { text, language } = useWorkbookMessages(),
    isFileSource = "bytes" in source,
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    operationScope = useOperationScope(),
    requestCounterRef = ne.useRef(0),
    abortControllerRef = ne.useRef(null),
    [options, setOptions] = ne.useState(() => ({
      ...delimitedImportDefaults,
      columns: {},
      includeHeader: true,
      bom: false,
      quoteAll: false,
      dateOutput: "calendar",
      ...(isFileSource
        ? {}
        : {
            delimiter: ",",
            encoding: "utf-8",
            newline: `\r
`,
            range:
              source.population?.range ??
              getWorksheetUsedRange(source.snapshot, source.sheetId),
            ...(source.population
              ? {
                  population: {
                    definition: source.population,
                    scope: "full",
                  },
                }
              : {}),
            hasHeader: source.population
              ? true
              : (source.snapshot.sheets.find(
                  (sheet) => sheet.id === source.sheetId,
                )?.delimitedSource?.hasHeader ?? true),
          }),
    })),
    buildPreview = (exportOptions) => {
      try {
        return {
          result: isFileSource
            ? undefined
            : exportWorksheetToCsv(
                source.snapshot,
                source.sheetId,
                exportOptions,
              ),
          error: "",
        };
      } catch (error) {
        return {
          result: undefined,
          error: error instanceof Error ? error.message : "Preview failed.",
        };
      }
    },
    [preview, setPreview] = ne.useState(() =>
      isFileSource && source.initial
        ? { result: source.initial, error: "" }
        : buildPreview(options),
    ),
    [disabled, setDisabled] = ne.useState(isFileSource && !source.initial),
    [isPreviewStale, setIsPreviewStale] = ne.useState(false),
    [columnPage, setColumnPage] = ne.useState(0),
    [actionError, setActionError] = ne.useState(""),
    report = preview.result?.report,
    pageCount = Math.max(1, Math.ceil((report?.columns ?? 1) / 20));
  function updateOption(optionKey, optionValue) {
    (setOptions((previousOptions) => ({
      ...previousOptions,
      [optionKey]: optionValue,
    })),
      setIsPreviewStale(true),
      requestCounterRef.current++,
      abortControllerRef.current?.abort(),
      setDisabled(false),
      setActionError(""));
  }
  async function onClick() {
    const requestId = ++requestCounterRef.current;
    abortControllerRef.current?.abort();
    const abortController = new AbortController();
    if (((abortControllerRef.current = abortController), isFileSource)) {
      setDisabled(true);
      try {
        const importResult = await importDelimitedFile(
          source.file ?? new File([source.bytes], source.fileName),
          options,
          {
            ...operationScope,
            signal: operationScope.signal
              ? AbortSignal.any([abortController.signal, operationScope.signal])
              : abortController.signal,
          },
        );
        if (requestCounterRef.current !== requestId) return;
        setPreview({ result: importResult, error: "" });
      } catch (error) {
        if (requestCounterRef.current !== requestId) return;
        setPreview({
          error: error instanceof Error ? error.message : "Preview failed.",
        });
      } finally {
        requestCounterRef.current === requestId && setDisabled(false);
      }
    } else setPreview(buildPreview(options));
    (setIsPreviewStale(false), setColumnPage(0), setActionError(""));
  }
  (ne.useEffect(
    () => (
      isFileSource && !source.initial && onClick(),
      () => {
        (requestCounterRef.current++, abortControllerRef.current?.abort());
      }
    ),
    [source],
  ),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const renderSelect = (labelText, selectKey, choices) =>
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: labelText }),
          i.jsx("select", {
            value: String(options[selectKey]),
            onChange: (event) => updateOption(selectKey, event.target.value),
            children: choices.map(([choiceValue, choiceLabel]) =>
              i.jsx(
                "option",
                { value: choiceValue, children: text(choiceLabel) },
                choiceValue,
              ),
            ),
          }),
        ],
      }),
    renderCheckbox = (checkboxLabel, checkboxKey) =>
      i.jsxs("label", {
        className: "lg-delimited-check",
        children: [
          i.jsx("input", {
            type: "checkbox",
            checked: options[checkboxKey],
            onChange: (event) =>
              updateOption(checkboxKey, event.target.checked),
          }),
          checkboxLabel,
        ],
      });
  return i.jsx("div", {
    className: "lg-delimited-overlay",
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      ref: dialogRef,
      className: "lg-delimited-review",
      onKeyDown: (event) => {
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
              ) ?? []),
            ],
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey &&
          (document.activeElement === firstFocusable ||
            document.activeElement === dialogRef.current)
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
            i.jsx("h2", {
              id: titleId,
              children: i.jsx(LocalizedText, {
                children: isFileSource
                  ? "Review delimited import"
                  : "Review delimited export",
              }),
            }),
            i.jsxs("p", {
              children: [
                isFileSource ? source.fileName : source.snapshot.name,
                " ",
                i.jsx(LocalizedText, { children: "· CSV / TSV" }),
              ],
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-delimited-scroll",
          children: [
            i.jsx("p", {
              children: isFileSource
                ? "Opening replaces the workbook after review. Cancel keeps it unchanged; Undo restores it after importing."
                : `Export calculated values from revision ${source.snapshot.revision}. Choose full or visible scope when a column filter is active; otherwise the rectangle includes hidden and filtered rows. CSV/TSV cannot retain workbook formatting or formulas.`,
            }),
            !isFileSource &&
              source.population &&
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Export population" }),
                  i.jsxs("select", {
                    "aria-label": text("Export population"),
                    value: options.population?.scope ?? "full",
                    onChange: (event) =>
                      updateOption("population", {
                        definition: source.population,
                        scope: event.target.value,
                      }),
                    lang: language("Export population"),
                    "data-ui-label": "Export population",
                    children: [
                      i.jsx("option", {
                        value: "full",
                        children:
                          "Full filter range, including hidden and filtered rows",
                      }),
                      i.jsx("option", {
                        value: "visible",
                        children: "Visible population only",
                      }),
                    ],
                  }),
                  i.jsx("span", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Preview again after changing scope. Save the schema report to retain criteria and original row identities.",
                    }),
                  }),
                ],
              }),
            i.jsxs("fieldset", {
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "File and data controls",
                  }),
                }),
                i.jsxs("div", {
                  className: "lg-delimited-controls",
                  children: [
                    renderSelect("Delimiter", "delimiter", [
                      ...(isFileSource ? [["auto", "Automatic"]] : []),
                      [",", "Comma (CSV)"],
                      ["	", "Tab (TSV)"],
                      [";", "Semicolon"],
                      ["|", "Pipe"],
                    ]),
                    renderSelect("Encoding", "encoding", [
                      ...(isFileSource
                        ? [["auto", "Automatic: BOM or UTF-8"]]
                        : []),
                      ["utf-8", "UTF-8"],
                      ["utf-16le", "UTF-16 little-endian"],
                      ["utf-16be", "UTF-16 big-endian"],
                      ["windows-1252", "Windows-1252"],
                    ]),
                    renderSelect("Quote character", "quote", [
                      ['"', 'Double quote (")'],
                      ["'", "Single quote (')"],
                      ["none", "No quotes"],
                    ]),
                    renderSelect("Quote escape", "escape", [
                      ["double", "Double the quote"],
                      ["backslash", "Backslash"],
                    ]),
                    renderSelect("Record newline", "newline", [
                      ...(isFileSource ? [["auto", "Automatic"]] : []),
                      [
                        `\r
`,
                        "CRLF",
                      ],
                      [
                        `
`,
                        "LF",
                      ],
                      ["\r", "CR"],
                    ]),
                    renderSelect("Decimal separator", "decimal", [
                      [".", "Dot"],
                      [",", "Comma"],
                    ]),
                    renderSelect(
                      "Grouping separator (input text)",
                      "grouping",
                      [
                        ["none", "None"],
                        [",", "Comma"],
                        [".", "Dot"],
                        [" ", "Space"],
                      ],
                    ),
                    renderSelect("Date order", "dateOrder", [
                      ["ymd", "Year / month / day"],
                      ["dmy", "Day / month / year"],
                      ["mdy", "Month / day / year"],
                    ]),
                    renderSelect("Empty fields", "empty", [
                      ["blank", "Blank cell"],
                      ["text", "Empty text"],
                    ]),
                    renderCheckbox("First record is a header", "hasHeader"),
                    renderCheckbox("Trim surrounding whitespace", "trim"),
                    !isFileSource &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Export range",
                              }),
                              i.jsx("input", {
                                value: options.range ?? "",
                                onChange: (event) =>
                                  updateOption("range", event.target.value),
                              }),
                            ],
                          }),
                          renderSelect("Date output", "dateOutput", [
                            ["calendar", "Calendar date/time (nearest second)"],
                            ["serial", "Native serial number"],
                          ]),
                          renderCheckbox(
                            "Include header in output",
                            "includeHeader",
                          ),
                          renderCheckbox(
                            "Include byte order mark (BOM)",
                            "bom",
                          ),
                          renderCheckbox("Quote all fields", "quoteAll"),
                        ],
                      }),
                  ],
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Auto preserves leading-zero identifiers, long numbers, dates and Boolean text. Choose a column type to convert it. Dates need four-digit years. Empty records retain their source row positions.",
              }),
            }),
            report &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Column schema",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      report.records.toLocaleString(),
                      " ",
                      i.jsx(LocalizedText, { children: "records ·" }),
                      " ",
                      report.columns.toLocaleString(),
                      " ",
                      i.jsx(LocalizedText, { children: "columns ·" }),
                      " ",
                      report.storedCells.toLocaleString(),
                      " ",
                      i.jsx(LocalizedText, { children: "stored cells ·" }),
                      " ",
                      (report.fileBytes / 1024).toFixed(1),
                      " ",
                      i.jsx(LocalizedText, { children: "KiB" }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "lg-delimited-schema",
                    children: report.schema
                      .slice(columnPage * 20, columnPage * 20 + 20)
                      .map((schemaColumn) =>
                        i.jsxs(
                          "label",
                          {
                            children: [
                              columnIndexToLetters(schemaColumn.index),
                              " — ",
                              schemaColumn.label || "(blank header)",
                              i.jsx("select", {
                                "aria-label": `Type for column ${columnIndexToLetters(schemaColumn.index)}`,
                                value:
                                  options.columns[schemaColumn.index] ?? "auto",
                                onChange: (event) =>
                                  updateOption("columns", {
                                    ...options.columns,
                                    [schemaColumn.index]: event.target.value,
                                  }),
                                children: [
                                  "auto",
                                  "text",
                                  "number",
                                  "integer",
                                  "boolean",
                                  "date",
                                ].map((columnType) =>
                                  i.jsx(
                                    "option",
                                    { value: columnType, children: columnType },
                                    columnType,
                                  ),
                                ),
                              }),
                              i.jsxs("small", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Observed: ",
                                  }),
                                  schemaColumn.observed.join(", ") ||
                                    "header only",
                                ],
                              }),
                            ],
                          },
                          schemaColumn.index,
                        ),
                      ),
                  }),
                  pageCount > 1 &&
                    i.jsxs("div", {
                      children: [
                        i.jsx("button", {
                          disabled: columnPage === 0,
                          onClick: () =>
                            setColumnPage((previousPage) => previousPage - 1),
                          children: i.jsx(LocalizedText, {
                            children: "Previous columns",
                          }),
                        }),
                        " ",
                        i.jsx(LocalizedText, { children: "Page " }),
                        columnPage + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "of " }),
                        pageCount,
                        " ",
                        i.jsx("button", {
                          disabled: columnPage + 1 >= pageCount,
                          onClick: () =>
                            setColumnPage((currentPage) => currentPage + 1),
                          children: i.jsx(LocalizedText, {
                            children: "Next columns",
                          }),
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("button", {
              onClick: onClick,
              disabled: disabled,
              children: i.jsx(LocalizedText, { children: "Refresh preview" }),
            }),
            disabled &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "Processing file in an isolated worker…",
                }),
              }),
            isPreviewStale &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Settings changed. Refresh preview before continuing.",
                }),
              }),
            preview.error &&
              i.jsx("p", { role: "alert", children: preview.error }),
            report &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Data preview",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, {
                        children:
                          "First 25 logical records and 200 columns. Validation checks the entire file or export range. ",
                      }),
                      report.issueCount,
                      " ",
                      i.jsx(LocalizedText, { children: "issues;" }),
                      " ",
                      report.changeCount,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "conversions or text changes, including",
                      }),
                      " ",
                      report.safeTextChanges,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "safe-export changes.",
                      }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "lg-delimited-table",
                    children: i.jsxs("table", {
                      "aria-label": text("Delimited data preview"),
                      lang: language("Delimited data preview"),
                      "data-ui-label": "Delimited data preview",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Record",
                                }),
                              }),
                              report.schema.slice(0, 200).map((headerColumn) =>
                                i.jsxs(
                                  "th",
                                  {
                                    children: [
                                      columnIndexToLetters(headerColumn.index),
                                      " — ",
                                      headerColumn.label,
                                    ],
                                  },
                                  headerColumn.index,
                                ),
                              ),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: report.preview.map((cell) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", { children: cell.record }),
                                  cell.values.map((value, valueIndex) =>
                                    i.jsx(
                                      "td",
                                      {
                                        title: cell.raw[valueIndex],
                                        children:
                                          value === null
                                            ? "(blank)"
                                            : value === ""
                                              ? "(empty text)"
                                              : String(value),
                                      },
                                      valueIndex,
                                    ),
                                  ),
                                ],
                              },
                              cell.record,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  !!report.issueCount &&
                    i.jsxs("div", {
                      role: "alert",
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(LocalizedText, {
                            children: "Resolve issues before continuing",
                          }),
                        }),
                        i.jsx("ul", {
                          children: report.issues.map((issue, issueIndex) =>
                            i.jsxs(
                              "li",
                              {
                                children: [
                                  i.jsx(LocalizedText, { children: "Record " }),
                                  issue.record,
                                  issue.column
                                    ? `, column ${issue.column}`
                                    : "",
                                  ": ",
                                  issue.message,
                                  issue.value !== undefined
                                    ? ` Value: ${issue.value}`
                                    : "",
                                ],
                              },
                              issueIndex,
                            ),
                          ),
                        }),
                        report.issueCount > report.issues.length &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Showing the first ",
                              }),
                              report.issues.length,
                              " ",
                              i.jsx(LocalizedText, { children: "of" }),
                              " ",
                              report.issueCount,
                              " ",
                              i.jsx(LocalizedText, { children: "issues." }),
                            ],
                          }),
                      ],
                    }),
                  i.jsxs("details", {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Conversions and changed text (",
                          }),
                          report.changeCount,
                          ")",
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "The saved report contains every changed value; the dialog shows the first 100.",
                        }),
                      }),
                      i.jsxs("table", {
                        "aria-label": text("Delimited changes"),
                        lang: language("Delimited changes"),
                        "data-ui-label": "Delimited changes",
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Cell",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Before",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "After",
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
                            children: report.changes
                              .slice(0, 100)
                              .map((change, changeIndex) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("td", { children: change.address }),
                                      i.jsx("td", { children: change.before }),
                                      i.jsx("td", { children: change.after }),
                                      i.jsx("td", { children: change.reason }),
                                    ],
                                  },
                                  changeIndex,
                                ),
                              ),
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Exchange policy",
                    }),
                  }),
                  i.jsx("ul", {
                    children: report.warnings.map((warning, warningIndex) =>
                      i.jsx("li", { children: warning }, warningIndex),
                    ),
                  }),
                ],
              }),
          ],
        }),
        (blocked || actionError) &&
          i.jsx("p", { role: "alert", children: blocked || actionError }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: !report || isPreviewStale || disabled,
              onClick: () => report && onReport(report),
              children: i.jsx(LocalizedText, {
                children: "Save schema and changes",
              }),
            }),
            i.jsx("button", {
              disabled:
                isPreviewStale || disabled || !!blocked || !report?.valid,
              onClick: () => {
                var importPreview, exportPreview;
                if (!(isPreviewStale || disabled || blocked || !report?.valid))
                  try {
                    isFileSource
                      ? (importPreview = preview.result) != null &&
                        importPreview.workbook &&
                        onImport(preview.result.workbook, report)
                      : (exportPreview = preview.result) != null &&
                        exportPreview.bytes &&
                        onExport(preview.result.bytes, report);
                  } catch (error) {
                    setActionError(
                      error instanceof Error
                        ? error.message
                        : "Exchange failed.",
                    );
                  }
              },
              children: i.jsx(LocalizedText, {
                children: isFileSource
                  ? "Import reviewed data"
                  : "Download reviewed data",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
function tokenizeVbaSource(sourceText) {
  const tokens = [];
  let position = 0,
    line = 1;
  const fail = (message) => {
    throw { line: line, message: message };
  };
  for (; position < sourceText.length;) {
    const character = sourceText[position];
    if (/\s/.test(character)) {
      (character ===
        `
` && line++,
        position++);
      continue;
    }
    if (sourceText.slice(position, position + 2) === "//") {
      for (
        ;
        position < sourceText.length &&
        sourceText[position] !==
          `
`;
      )
        position++;
      continue;
    }
    if (sourceText.slice(position, position + 2) === "/*") {
      for (
        position += 2;
        position < sourceText.length &&
        sourceText.slice(position, position + 2) !== "*/";
      )
        sourceText[position++] ===
          `
` && line++;
      (position === sourceText.length && fail("Unclosed block comment."),
        (position += 2));
      continue;
    }
    const tokenLine = line;
    if (character === '"' || character === "'") {
      let stringValue = "";
      const quoteChar = character;
      for (
        position++;
        position < sourceText.length && sourceText[position] !== quoteChar;
      ) {
        let stringChar = sourceText[position++];
        if (
          ((stringChar ===
            `
` ||
            stringChar === "\r") &&
            fail("Use an escaped newline inside a string literal."),
          stringChar === "\\")
        ) {
          const escapeChar = sourceText[position++],
            escapeMap = {
              n: `
`,
              r: "\r",
              t: "	",
              b: "\b",
              f: "\f",
              v: "\v",
              "\\": "\\",
              '"': '"',
              "'": "'",
            };
          if (escapeChar === "u" || escapeChar === "x") {
            const hexDigitCount = escapeChar === "u" ? 4 : 2,
              hexDigits = sourceText.slice(position, position + hexDigitCount);
            (new RegExp(`^[a-f0-9]{${hexDigitCount}}$`, "i").test(hexDigits) ||
              fail("Invalid string escape."),
              (stringChar = String.fromCharCode(parseInt(hexDigits, 16))),
              (position += hexDigitCount));
          } else
            Object.hasOwn(escapeMap, escapeChar)
              ? (stringChar = escapeMap[escapeChar])
              : fail("Unsupported string escape.");
        }
        stringValue += stringChar;
      }
      (sourceText[position++] !== quoteChar && fail("Unclosed string literal."),
        tokens.push({ type: "str", value: stringValue, line: tokenLine }));
    } else {
      const remainingText = sourceText.slice(position),
        numberMatch = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/.exec(
          remainingText,
        ),
        identifierMatch = /^[A-Za-z_$][\w$]*/.exec(remainingText);
      numberMatch
        ? (tokens.push({
            type: "number",
            value: numberMatch[0],
            line: tokenLine,
          }),
          (position += numberMatch[0].length))
        : identifierMatch
          ? (tokens.push({
              type: "id",
              value: identifierMatch[0],
              line: tokenLine,
            }),
            (position += identifierMatch[0].length))
          : "().,;:=[]{}+-".includes(character)
            ? (tokens.push({
                type: "symbol",
                value: character,
                line: tokenLine,
              }),
              position++)
            : fail(
                `Unsupported source token ${character}. Templates, expressions and arbitrary JavaScript are outside this adapter.`,
              );
    }
    tokens.length > automationLimits.maxTokens &&
      fail("Source exceeds the 20,000-token migration limit.");
  }
  return tokens;
}
class VbaAssignmentParser {
  constructor(tokens) {
    defineDataPropertyByKey(this, "pos", 0);
    defineDataPropertyByKey(this, "bindings", new Map());
    defineDataPropertyByKey(this, "commands", []);
    defineDataPropertyByKey(this, "issues", []);
    this.ts = tokens;
  }
  peek(expectedValue) {
    return expectedValue === undefined
      ? this.ts[this.pos]
      : this.ts[this.pos]?.value === expectedValue;
  }
  error(message) {
    throw {
      line: this.ts[this.pos]?.line ?? this.ts.at(-1)?.line ?? 1,
      message: message,
    };
  }
  take(expectedToken) {
    const token = this.ts[this.pos++];
    return (
      (!token ||
        (expectedToken !== undefined &&
          (token.value !== expectedToken ||
            token.type !==
              ("().,;:=[]{}+-".includes(expectedToken) ? "symbol" : "id")))) &&
        (this.pos--,
        this.error(
          `Expected ${expectedToken ?? "source token"}; source is outside the literal-range grammar.`,
        )),
      token
    );
  }
  id() {
    const identifierToken = this.take();
    return (
      identifierToken.type !== "id" &&
        this.error("Expected a plain identifier."),
      identifierToken.value
    );
  }
  literal() {
    if (this.peek("[")) {
      this.take("[");
      const rows = [];
      for (; !this.peek("]");) {
        this.take("[");
        const row = [];
        for (; !this.peek("]");) {
          const cellValue = this.scalar();
          if (
            (row.push(cellValue),
            row.length > automationLimits.maxCells &&
              this.error("Matrix exceeds 500 cells."),
            this.peek(","))
          )
            this.take(",");
          else break;
        }
        if (
          (this.take("]"),
          rows.push(row),
          (rows.length > automationLimits.maxCells ||
            rows.reduce((total, matrixRow) => total + matrixRow.length, 0) >
              500) &&
            this.error("Matrix exceeds 500 cells."),
          this.peek(","))
        )
          this.take(",");
        else break;
      }
      return (this.take("]"), rows);
    }
    return this.scalar();
  }
  scalar() {
    let sign = 1;
    const hasUnarySign = this.peek("-") || this.peek("+");
    this.peek("-")
      ? ((sign = -1), this.take("-"))
      : this.peek("+") && this.take("+");
    const literalToken = this.take();
    if (literalToken.type === "number") {
      const numberValue = sign * Number(literalToken.value);
      return (
        Number.isFinite(numberValue) ||
          this.error("Numeric literal must be finite."),
        numberValue
      );
    }
    if (
      (hasUnarySign &&
        this.error("Only numeric literals may have a unary sign."),
      literalToken.type === "str")
    )
      return literalToken.value;
    if (
      literalToken.type === "id" &&
      ["true", "false", "null"].includes(literalToken.value)
    )
      return literalToken.value === "null"
        ? null
        : literalToken.value === "true";
    this.error(
      "Only literal text, numbers, booleans and rectangular literal matrices are supported.",
    );
  }
  argument() {
    const nextToken = this.peek();
    if (
      (nextToken || this.error("Missing argument."),
      nextToken.type === "id" &&
        !["true", "false", "null"].includes(nextToken.value))
    ) {
      const binding = this.bindings.get(this.id());
      return (binding || this.error("Unknown or dynamic binding."), binding);
    }
    return { kind: "value", value: this.literal() };
  }
  expression() {
    let receiver = this.argument();
    for (; this.peek(".");) {
      this.take(".");
      const methodName = this.id(),
        line = this.ts[this.pos - 1].line;
      if (
        (this.take("("),
        receiver.kind === "workbook" && methodName === "getActiveWorksheet")
      )
        (this.take(")"), (receiver = { kind: "sheet" }));
      else if (receiver.kind === "workbook" && methodName === "getWorksheet") {
        const sheetArgument = this.argument();
        (this.take(")"),
          (sheetArgument.kind !== "value" ||
            typeof sheetArgument.value != "string") &&
            this.error("Worksheet name must be literal text."),
          (receiver = { kind: "sheet", sheet: sheetArgument.value }));
      } else if (receiver.kind === "sheet" && methodName === "getRange") {
        const rangeArgument = this.argument();
        (this.take(")"),
          (rangeArgument.kind !== "value" ||
            typeof rangeArgument.value != "string") &&
            this.error("Range address must be literal text."),
          (receiver = {
            kind: "range",
            sheet: receiver.sheet,
            range: rangeArgument.value,
          }));
      } else if (
        receiver.kind === "range" &&
        ["setValue", "setValues", "setFormula", "setFormulas"].includes(
          methodName,
        )
      ) {
        const valuesArgument = this.argument();
        (this.take(")"),
          valuesArgument.kind !== "value" &&
            this.error("Range write arguments must be literals."));
        const isMatrix = Array.isArray(valuesArgument.value);
        (isMatrix !== ["setValues", "setFormulas"].includes(methodName) &&
          this.error(
            `${methodName} requires ${isMatrix ? "a scalar" : "a matrix"}.`,
          ),
          (receiver = {
            kind: "command",
            command: {
              line: line,
              kind: methodName.includes("Formula")
                ? "set-formulas"
                : "set-values",
              sheet: receiver.sheet,
              range: receiver.range,
              values: valuesArgument.value,
            },
          }));
      } else if (receiver.kind === "range" && methodName === "clear") {
        for (const expectedPart of [
          "ExcelScript",
          ".",
          "ClearApplyTo",
          ".",
          "contents",
          ")",
        ])
          this.take(expectedPart);
        receiver = {
          kind: "command",
          command: {
            line: line,
            kind: "clear-contents",
            sheet: receiver.sheet,
            range: receiver.range,
          },
        };
      } else
        this.error(
          `Unsupported Office Scripts API: ${methodName}. No part of this script can run until every issue is resolved.`,
        );
    }
    return receiver;
  }
  parse() {
    try {
      (this.take("function"), this.take("main"), this.take("("));
      const workbookName = this.id();
      for (
        this.take(":"),
          this.take("ExcelScript"),
          this.take("."),
          this.take("Workbook"),
          this.take(")"),
          this.take("{"),
          this.bindings.set(workbookName, { kind: "workbook" });
        this.pos < this.ts.length && !this.peek("}");
      )
        try {
          if (this.peek("const") || this.peek("let")) {
            this.take();
            const bindingName = this.id();
            (this.bindings.has(bindingName) &&
              this.error("Binding redeclaration is not supported."),
              this.take("="));
            const assignedValue = this.expression();
            (assignedValue.kind === "command" &&
              this.error("A command cannot be assigned to a binding."),
              this.take(";"),
              this.bindings.set(bindingName, assignedValue));
          } else {
            const expressionResult = this.expression();
            (expressionResult.kind !== "command" &&
              this.error("Only literal range commands may be executed."),
              this.take(";"),
              this.commands.push(expressionResult.command));
          }
        } catch (error) {
          for (
            this.issues.push(createVbaSyntaxError(error));
            this.pos < this.ts.length && !this.peek(";") && !this.peek("}");
          )
            this.pos++;
          this.peek(";") && this.pos++;
        }
      (this.take("}"),
        this.pos !== this.ts.length &&
          this.error(
            "Additional procedures or top-level executable statements are unsupported.",
          ));
    } catch (error) {
      this.issues.push(createVbaSyntaxError(error));
    }
  }
}
function createVbaSyntaxError(error) {
  const failure = error;
  return {
    line: failure.line ?? 1,
    message: failure.message ?? "Unsupported source syntax.",
  };
}
const vbaStringLiteralPattern = '"(?:[^"]|"")*"';
function parseVbaAssignmentValue(rawValue) {
  if (new RegExp("^" + vbaStringLiteralPattern + "$").test(rawValue))
    return rawValue.slice(1, -1).replaceAll('""', '"');
  if (/^(true|false)$/i.test(rawValue)) return /^true$/i.test(rawValue);
  if (
    /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(rawValue) &&
    Number.isFinite(Number(rawValue))
  )
    return Number(rawValue);
  throw Error(
    "Only literal text, finite numbers or booleans are supported in VBA assignments.",
  );
}
function splitVbaStatements(sourceText) {
  let statement = "",
    isInsideString = false;
  for (let charIndex = 0; charIndex < sourceText.length; charIndex++) {
    const character = sourceText[charIndex];
    if (character === '"') {
      if (
        ((statement += character),
        isInsideString && sourceText[charIndex + 1] === '"')
      ) {
        statement += sourceText[++charIndex];
        continue;
      }
      isInsideString = !isInsideString;
    } else {
      if (character === "'" && !isInsideString) break;
      if (character === ":" && !isInsideString)
        throw Error("Multiple colon-separated VBA statements are unsupported.");
      statement += character;
    }
  }
  if (isInsideString) throw Error("Unclosed VBA string.");
  return statement.trim();
}

export {
  WhatIfAnalysisDialog,
  splitVbaStatements,
  vbaStringLiteralPattern,
  tokenizeVbaSource,
  VbaAssignmentParser,
  createVbaSyntaxError,
  parseVbaAssignmentValue,
  DelimitedImportExportDialog,
};
