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

function TextCleanupDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    isApplied = ne.useRef(false),
    [sourceRange, setSourceRange] = ne.useState(props.range),
    [outputName, setOutputName] = ne.useState("Cleaned text"),
    [checked, setChecked] = ne.useState(false),
    [steps, setSteps] = ne.useState([
      getTextFunctionSpec("CLEAN"),
      getTextFunctionSpec("TRIM"),
    ]),
    [selectedFunctionName, setSelectedFunctionName] = ne.useState("SUBSTITUTE"),
    [previewState, setPreviewState] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [rowPage, setRowPage] = ne.useState(0),
    [selectedCleanupId, setSelectedCleanupId] = ne.useState(
      props.workbook.textCleanups?.at(-1)?.id ?? "",
    ),
    [evidencePage, setEvidencePage] = ne.useState(0),
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    cleanupOptions = {
      readOnly: props.readOnly,
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: allowedSheetIds,
      allowNewSheet: props.allowNewSheet ?? true,
    },
    requestKey = JSON.stringify({
      sourceRange: sourceRange,
      outputName: outputName,
      steps: steps,
      options: cleanupOptions,
      selection: [props.sheetId, props.range],
    }),
    workbookSnapshot = JSON.stringify(props.workbook),
    currentPlan =
      previewState?.key === requestKey &&
      previewState.plan.snapshot === workbookSnapshot
        ? previewState.plan
        : null,
    latestInputsRef = ne.useRef({
      p: props,
      key: requestKey,
      options: cleanupOptions,
    });
  latestInputsRef.current = {
    p: props,
    key: requestKey,
    options: cleanupOptions,
  };
  const clearPreview = () => {
    (setPreviewState(null),
      setErrorMessage(""),
      (isApplied.current = false),
      setRowPage(0));
  };
  (ne.useEffect(() => {
    (setSourceRange(props.range),
      setPreviewState(null),
      setErrorMessage(""),
      (isApplied.current = false));
  }, [props.sheetId, props.range]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current
          ?.querySelector("input:not(:disabled),button")
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const reportError = (failure) =>
      setErrorMessage(
        failure instanceof Error
          ? failure.message
          : "Cleanup failed. Review the source and steps.",
      ),
    onClick = () => {
      clearPreview();
      try {
        const cleanupSteps = steps.map((step) => ({
          function: step.name,
          args: step.fields
            .filter((field) => !(field.optional && field.value === ""))
            .map((field) =>
              field.kind === "number"
                ? field.value.trim()
                  ? Number(field.value)
                  : NaN
                : field.kind === "boolean"
                  ? field.value === "TRUE"
                  : field.value,
            ),
        }));
        setPreviewState({
          key: requestKey,
          plan: previewTextCleanup(
            props.workbook,
            {
              sourceSheetId: props.sheetId,
              sourceRange: sourceRange,
              outputName: outputName,
              steps: cleanupSteps,
            },
            cleanupOptions,
          ),
        });
      } catch (error) {
        reportError(error);
      }
    },
    applyCleanup = () => {
      if (!(!currentPlan || isApplied.current))
        try {
          if (latestInputsRef.current.key !== previewState?.key)
            throw Error("The cleanup settings changed. Preview again.");
          const textCleanupOutput = createTextCleanupOutput(
            latestInputsRef.current.p.workbook,
            currentPlan,
            latestInputsRef.current.options,
          );
          ((isApplied.current = true),
            latestInputsRef.current.p.onWorkbookChange(
              textCleanupOutput,
              `Text cleanup: ${currentPlan.sourceSheetName}!${currentPlan.request.sourceRange} → ${currentPlan.request.outputName}`,
            ),
            setPreviewState(null),
            setSelectedCleanupId(textCleanupOutput.textCleanups.at(-1).id),
            setEvidencePage(0));
        } catch (error) {
          reportError(error);
        }
    },
    visibleCleanups = (props.workbook.textCleanups ?? []).filter((cleanup) => {
      const sheet = props.workbook.sheets.find(
        (sheet) => sheet.id === cleanup.sourceSheetId,
      );
      return (
        (props.allowedSheetIds === undefined ||
          allowedSheetIds.includes(cleanup.sourceSheetId)) &&
        (cleanupOptions.includeHidden ||
          !sheet ||
          (!sheet.hidden &&
            !sheet.veryHidden &&
            !cn(sheet, "row").length &&
            !cn(sheet, "column").length))
      );
    }),
    selectedCleanup =
      visibleCleanups.find((cleanup) => cleanup.id === selectedCleanupId) ??
      visibleCleanups.at(-1);
  return i.jsx("div", {
    className: "ttc-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      className: "ttc-dialog cleanup-dialog",
      ref: dialogRef,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button,input,select,summary",
            ) ?? []),
          ].filter(
            (element) =>
              !element.matches(":disabled") &&
              !element.closest("details:not([open]) > :not(summary)"),
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
            i.jsx("h2", {
              id: `${dialogId}-title`,
              children: i.jsx(LocalizedText, { children: "Clean up text" }),
            }),
            i.jsx("button", {
              "aria-label": text("Close text cleanup"),
              onClick: props.onClose,
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
          disabled: props.readOnly || props.allowNewSheet === false,
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
                    (setSourceRange(event.target.value), clearPreview());
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
                    (setOutputName(event.target.value), clearPreview());
                  },
                }),
              ],
            }),
            props.allowHidden &&
              i.jsxs("label", {
                className: "ttc-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (setChecked(event.target.checked), clearPreview());
                    },
                  }),
                  i.jsx(LocalizedText, {
                    children: "Include hidden source cells",
                  }),
                ],
              }),
          ],
        }),
        props.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "This workbook is read-only. Saved cleanup evidence can still be reviewed.",
            }),
          }),
        i.jsxs("fieldset", {
          disabled: props.readOnly || props.allowNewSheet === false,
          className: "cleanup-steps",
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, { children: "Steps in order" }),
            }),
            steps.map((step, stepIndex) =>
              i.jsxs(
                "div",
                {
                  className: "cleanup-step",
                  children: [
                    i.jsxs("strong", {
                      children: [stepIndex + 1, ". ", step.name],
                    }),
                    step.fields.map((field, fieldIndex) =>
                      i.jsxs(
                        "label",
                        {
                          children: [
                            field.label,
                            field.kind === "boolean"
                              ? i.jsxs("select", {
                                  "aria-label": `Step ${stepIndex + 1} ${field.label}`,
                                  value: field.value,
                                  onChange: (event) => {
                                    (setSteps(
                                      steps.map((step, candidateStepIndex) =>
                                        candidateStepIndex === stepIndex
                                          ? {
                                              ...step,
                                              fields: step.fields.map(
                                                (field, candidateFieldIndex) =>
                                                  candidateFieldIndex ===
                                                  fieldIndex
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
                                      clearPreview());
                                  },
                                  children: [
                                    i.jsx("option", { children: "TRUE" }),
                                    i.jsx("option", { children: "FALSE" }),
                                  ],
                                })
                              : i.jsx("input", {
                                  "aria-label": `Step ${stepIndex + 1} ${field.label}`,
                                  value: field.value,
                                  onChange: (event) => {
                                    (setSteps(
                                      steps.map((step, candidateStepIndex) =>
                                        candidateStepIndex === stepIndex
                                          ? {
                                              ...step,
                                              fields: step.fields.map(
                                                (field, candidateFieldIndex) =>
                                                  candidateFieldIndex ===
                                                  fieldIndex
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
                                      clearPreview());
                                  },
                                }),
                          ],
                        },
                        fieldIndex,
                      ),
                    ),
                    i.jsx("button", {
                      "aria-label": `Move step ${stepIndex + 1} up`,
                      disabled: stepIndex === 0,
                      onClick: () => {
                        const reorderedSteps = [...steps];
                        (([
                          reorderedSteps[stepIndex - 1],
                          reorderedSteps[stepIndex],
                        ] = [
                          reorderedSteps[stepIndex],
                          reorderedSteps[stepIndex - 1],
                        ]),
                          setSteps(reorderedSteps),
                          clearPreview());
                      },
                      children: "↑",
                    }),
                    i.jsx("button", {
                      "aria-label": `Remove step ${stepIndex + 1}`,
                      onClick: () => {
                        (setSteps(
                          steps.filter(
                            (step, candidateStepIndex) =>
                              stepIndex !== candidateStepIndex,
                          ),
                        ),
                          clearPreview());
                      },
                      children: i.jsx(LocalizedText, { children: "Remove" }),
                    }),
                  ],
                },
                stepIndex,
              ),
            ),
            i.jsxs("div", {
              className: "cleanup-add",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Add function" }),
                    i.jsx("select", {
                      value: selectedFunctionName,
                      onChange: (event) =>
                        setSelectedFunctionName(event.target.value),
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
                    (setSteps([
                      ...steps,
                      getTextFunctionSpec(selectedFunctionName),
                    ]),
                      clearPreview());
                  },
                  children: i.jsx(LocalizedText, { children: "Add step" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (setSteps([
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
                      clearPreview());
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
          disabled: props.readOnly || props.allowNewSheet === false,
          children: i.jsx(LocalizedText, { children: "Preview cleanup" }),
        }),
        errorMessage &&
          i.jsx("p", {
            className: "ttc-error",
            role: "alert",
            children: errorMessage,
          }),
        previewState &&
          !currentPlan &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook or settings changed. Preview again before applying.",
            }),
          }),
        currentPlan &&
          i.jsxs("section", {
            className: "ttc-preview",
            "aria-label": text("Text cleanup preview"),
            lang: language("Text cleanup preview"),
            "data-ui-label": "Text cleanup preview",
            children: [
              i.jsxs("h3", {
                children: [
                  currentPlan.rows.length,
                  " ",
                  i.jsx(LocalizedText, { children: "results · " }),
                  currentPlan.errors.length,
                  " ",
                  i.jsx(LocalizedText, { children: "errors" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Source revision " }),
                  props.workbook.revision,
                  i.jsx(LocalizedText, {
                    children:
                      ". Quoted values expose spaces; nonbreaking spaces appear as \\u00a0.",
                  }),
                ],
              }),
              i.jsx("code", { children: currentPlan.formulas[0] }),
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
                      children: currentPlan.rows
                        .slice(rowPage * 20, rowPage * 20 + 20)
                        .map((previewRow) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", {
                                  children: previewRow.sourceAddress,
                                }),
                                i.jsx("td", {
                                  children: i.jsx("code", {
                                    children: stringifyEscapedJson(
                                      previewRow.original?.raw ?? null,
                                    ),
                                  }),
                                }),
                                i.jsx("td", {
                                  children: i.jsx("code", {
                                    children: stringifyEscapedJson(
                                      previewRow.result,
                                    ),
                                  }),
                                }),
                                i.jsx("td", {
                                  children:
                                    currentPlan.errors.find(
                                      (error) =>
                                        error.sourceAddress ===
                                        previewRow.sourceAddress,
                                    )?.error ?? typeof previewRow.result,
                                }),
                              ],
                            },
                            previewRow.sourceAddress,
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
                    disabled: rowPage === 0,
                    onClick: () => setRowPage(rowPage - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous results",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      rowPage + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      Math.ceil(currentPlan.rows.length / 20),
                    ],
                  }),
                  i.jsx("button", {
                    disabled: (rowPage + 1) * 20 >= currentPlan.rows.length,
                    onClick: () => setRowPage(rowPage + 1),
                    children: i.jsx(LocalizedText, {
                      children: "Next results",
                    }),
                  }),
                ],
              }),
              currentPlan.errors.length > 0 &&
                i.jsx("p", {
                  role: "alert",
                  children: i.jsx(LocalizedText, {
                    children:
                      "Resolve all errors before applying. No partial output will be created.",
                  }),
                }),
            ],
          }),
        visibleCleanups.length > 0 &&
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(LocalizedText, {
                    children: "Saved text cleanup evidence (",
                  }),
                  visibleCleanups.length,
                  ")",
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Cleanup record" }),
                  i.jsx("select", {
                    value: selectedCleanup?.id ?? "",
                    onChange: (event) => {
                      (setSelectedCleanupId(event.target.value),
                        setEvidencePage(0));
                    },
                    children: visibleCleanups.map((cleanup) =>
                      i.jsxs(
                        "option",
                        {
                          value: cleanup.id,
                          children: [
                            cleanup.outputSheetName,
                            " · source revision ",
                            cleanup.sourceRevision,
                            " ·",
                            " ",
                            cleanup.createdAt,
                          ],
                        },
                        cleanup.id,
                      ),
                    ),
                  }),
                ],
              }),
              selectedCleanup &&
                i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("p", {
                      children: [
                        i.jsx(LocalizedText, { children: "Source: " }),
                        selectedCleanup.sourceSheetName,
                        " ",
                        i.jsx(LocalizedText, { children: "· workbook" }),
                        " ",
                        selectedCleanup.workbookId,
                        " ",
                        i.jsx(LocalizedText, { children: "· revision " }),
                        selectedCleanup.sourceRevision,
                        i.jsx(LocalizedText, {
                          children:
                            ". Addresses and values are snapshots from the cleanup, even if cells later move or are deleted.",
                        }),
                      ],
                    }),
                    i.jsx("code", {
                      children: stringifyEscapedJson(selectedCleanup.steps),
                    }),
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
                            children: selectedCleanup.rows
                              .slice(evidencePage * 20, evidencePage * 20 + 20)
                              .map((evidenceRow) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("td", {
                                        children: evidenceRow.sourceAddress,
                                      }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: stringifyEscapedJson(
                                            evidenceRow.original ?? {
                                              raw: null,
                                            },
                                          ),
                                        }),
                                      }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: stringifyEscapedJson(
                                            evidenceRow.result,
                                          ),
                                        }),
                                      }),
                                      i.jsx("td", {
                                        children: evidenceRow.origin
                                          ? `${evidenceRow.origin.sourceName} · ${evidenceRow.origin.sourceSheet} · row ${evidenceRow.origin.originalRow}`
                                          : "Editor value snapshot",
                                      }),
                                    ],
                                  },
                                  evidenceRow.sourceAddress,
                                ),
                              ),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("button", {
                      disabled: evidencePage === 0,
                      onClick: () => setEvidencePage(evidencePage - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous evidence",
                      }),
                    }),
                    i.jsx("button", {
                      disabled:
                        (evidencePage + 1) * 20 >= selectedCleanup.rows.length,
                      onClick: () => setEvidencePage(evidencePage + 1),
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
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
            i.jsx("button", {
              className: "ttc-primary",
              disabled:
                !currentPlan ||
                !!currentPlan.errors.length ||
                props.readOnly ||
                isApplied.current,
              onClick: applyCleanup,
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
  parseFieldPositions = (positionsText, canBeEmpty = false) => {
    if (!positionsText.trim() && canBeEmpty) return [];
    if (!/^\s*\d+(?:\s*,\s*\d+)*\s*$/.test(positionsText))
      throw Error("Enter field positions separated by commas, such as 1, 3.");
    const positions = positionsText
      .split(",")
      .map((positionText) => Number(positionText) - 1);
    if (
      positions.some((position) => position < 0 || position > 49) ||
      new Set(positions).size !== positions.length
    )
      throw Error("Choose distinct field positions from 1 to 50.");
    return positions;
  };
function describeSplitFieldStep(step) {
  if (step.kind === "pattern")
    return `Infer “${step.name}”: ${describeInferredTextRule(step.pattern)} (${step.examples.length} examples)`;
  if (step.kind === "split")
    return `Split field ${step.column + 1}: ${step.mode === "fixed" ? "breaks " + step.breaks.join(", ") : "delimiters " + JSON.stringify(step.delimiters) + (step.collapse ? ", collapse consecutive" : "")}`;
  const fieldList = step.columns.map((column) => column + 1).join(", ");
  return `${textCleanupStepLabels.find((textCleanupStepLabel) => textCleanupStepLabel[0] === step.kind)?.[1]} · fields ${fieldList}${step.kind === "convert" ? ` → ${step.to} (decimal ${step.decimal}, groups ${step.group || "none"}, dates ${step.dateOrder})` : step.kind === "fill" ? ` → ${step.direction}${step.groupColumns.length ? " within fields " + step.groupColumns.map((groupColumn) => groupColumn + 1).join(", ") : ""}` : step.kind === "concat" ? ` → ${step.name}; separator ${JSON.stringify(step.separator)}` : step.kind === "case" || step.kind === "trim" ? ` → ${step.mode}` : step.kind === "replace" ? ` ${JSON.stringify(step.find)} → ${JSON.stringify(step.replacement)}` : ""}`;
}

export {
  textCleanupStepLabels,
  describeSplitFieldStep,
  parseFieldPositions,
  TextCleanupDialog,
};
