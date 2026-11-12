import { stringifyJsonValue } from "./42-combine-populations-and-text-columns.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  normalizeDatePolicy,
  dateSystemMaxSerial,
  parseDateText,
} from "./01-schemas-date-functions-and-validation.js";
import {
  findSheetOrThrow,
  calculateWorkbookWithSettings,
  _t,
  parseRangeAddress,
  assertSheetUnprotected,
  matrixResultFunctionNames,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { assertSelectionInAuthorizedSheet } from "./15-data-model-dialog.js";
import {
  commitApprovedTransformation,
  verifyTransformationPreview,
  selectExceptionFreeRows,
  buildTransformationPreview,
} from "./16-transformation-recipe-engine.js";
import {
  textCleanupStepLabels,
  describeSplitFieldStep,
  parseFieldPositions,
} from "./43-text-cleanup.js";

function CleanAndReshapeDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    hasAppliedPlan = ne.useRef(false),
    previousSource = ne.useRef([props.sheetId, props.range]),
    [sourceSheetId, setSourceSheetId] = ne.useState(
      props.initialRequest?.sourceSheetId ?? props.sheetId,
    ),
    [sourceRange, setSourceRange] = ne.useState(
      props.initialRequest?.sourceRange ?? props.range,
    ),
    [hasHeader, setHasHeader] = ne.useState(
      props.initialRequest?.header ?? true,
    ),
    [outputName, setOutputName] = ne.useState(
      props.initialRequest?.outputName ?? "Reshaped fields",
    ),
    [checked, setChecked] = ne.useState(false),
    [steps, setSteps] = ne.useState(() =>
      structuredClone(props.initialRequest?.steps ?? []),
    ),
    [editing, setEditing] = ne.useState(null),
    [operation, setOperation] = ne.useState("trim"),
    [columns, setColumns] = ne.useState("1"),
    [trimMode, setTrimMode] = ne.useState("collapse"),
    [caseMode, setCaseMode] = ne.useState("upper"),
    [find, setFind] = ne.useState(""),
    [replacement, setReplacement] = ne.useState(""),
    [splitMode, setSplitMode] = ne.useState("delimited"),
    [delimiters, setDelimiters] = ne.useState(",;"),
    [shouldCollapse, setShouldCollapse] = ne.useState(false),
    [breaks, setBreaks] = ne.useState("5, 10"),
    [separator, X$1] = ne.useState(" "),
    [fieldName, setFieldName] = ne.useState("Combined"),
    [shouldSkipBlank, setShouldSkipBlank] = ne.useState(false),
    [convertTo, setConvertTo] = ne.useState("number"),
    [decimal, setDecimal] = ne.useState("."),
    [group, setGroup] = ne.useState(","),
    [dateOrder, setDateOrder] = ne.useState("MDY"),
    [direction, setDirection] = ne.useState("down"),
    [fillValue, setFillValue] = ne.useState(""),
    [fillType, setFillType] = ne.useState("text"),
    [groupColumns, setGroupColumns] = ne.useState(""),
    [exampleRows, setExampleRows] = ne.useState("2, 3"),
    [exampleValues, setExampleValues] = ne.useState(""),
    [candidates, setCandidates] = ne.useState([]),
    [candidate, setCandidate] = ne.useState(""),
    [candidateStatus, setCandidateStatus] = ne.useState(""),
    [planState, setPlanState] = ne.useState(null),
    [previewErrorState, setPreviewErrorState] = ne.useState(null),
    [outputPage, setOutputPage] = ne.useState(0),
    [exceptionPage, setExceptionPage] = ne.useState(0),
    [shouldAcceptExceptions, setShouldAcceptExceptions] = ne.useState(false),
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    transformOptions = {
      readOnly: props.readOnly,
      allowedSheetIds: allowedSheetIds,
      includeHidden: !!props.allowHidden && checked,
      allowNewSheet: props.allowNewSheet ?? true,
    },
    selectableSheets = props.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        (transformOptions.includeHidden ||
          (!sheet.hidden && !sheet.veryHidden)),
    ),
    request = {
      sourceSheetId: sourceSheetId,
      sourceRange: sourceRange,
      header: hasHeader,
      outputName: outputName,
      steps: steps,
    },
    requestKey = JSON.stringify({
      request: request,
      options: transformOptions,
      selection: [props.sheetId, props.range],
      editing: editing,
      operation: operation,
      columns: columns,
      trimMode: trimMode,
      caseMode: caseMode,
      find: find,
      replacement: replacement,
      splitMode: splitMode,
      delimiters: delimiters,
      collapse: shouldCollapse,
      breaks: breaks,
      separator: separator,
      fieldName: fieldName,
      skipBlank: shouldSkipBlank,
      convertTo: convertTo,
      decimal: decimal,
      group: group,
      dateOrder: dateOrder,
      direction: direction,
      fillValue: fillValue,
      fillType: fillType,
      groupColumns: groupColumns,
      exampleRows: exampleRows,
      exampleValues: exampleValues,
      candidate: candidate,
    }),
    currentPlan =
      planState?.key === requestKey &&
      planState.plan.snapshot === JSON.stringify(props.workbook)
        ? planState.plan
        : null,
    previewErrorMessage =
      previewErrorState?.key === requestKey &&
      previewErrorState.snapshot === JSON.stringify(props.workbook)
        ? previewErrorState.message
        : null,
    candidateSearchKey = JSON.stringify({
      sourceSheetId: sourceSheetId,
      sourceRange: sourceRange,
      header: hasHeader,
      outputName: outputName,
      steps: steps,
      editing: editing,
      columns: columns,
      exampleRows: exampleRows,
      exampleValues: exampleValues,
      options: transformOptions,
      selection: [props.sheetId, props.range],
      snapshot: props.workbook,
    }),
    [candidatesKey, setCandidatesKey] = ne.useState(""),
    visibleCandidates = candidatesKey === candidateSearchKey ? candidates : [];
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Transform source range"]')
        ?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      (previousSource.current[0] === props.sheetId &&
        previousSource.current[1] === props.range) ||
        ((previousSource.current = [props.sheetId, props.range]),
        setSourceSheetId(props.sheetId),
        setSourceRange(props.range),
        setPlanState(null),
        setCandidates([]),
        setCandidate(""),
        setPreviewErrorState(null),
        setShouldAcceptExceptions(false));
    }, [props.sheetId, props.range]));
  const recordPreviewError = (error) =>
      setPreviewErrorState({
        message:
          error instanceof Error
            ? error.message
            : "Review the transformation settings.",
        key: requestKey,
        snapshot: JSON.stringify(props.workbook),
      }),
    clearPreview = () => {
      (setPlanState(null),
        setPreviewErrorState(null),
        setShouldAcceptExceptions(false),
        (hasAppliedPlan.current = false));
    },
    parseExamples = () => {
      if (!/^\s*\d+(?:\s*,\s*\d+)+\s*$/.test(exampleRows))
        throw Error(
          "Enter at least two worksheet row numbers, separated by commas.",
        );
      const rowNumbers = exampleRows
          .split(",")
          .map((rowText) => Number(rowText)),
        valueLines = exampleValues.split(`
`);
      if (rowNumbers.length !== valueLines.length)
        throw Error("Provide one example value per row, on separate lines.");
      return rowNumbers.map((row, rowIndex) => ({
        row: row,
        value: valueLines[rowIndex],
      }));
    },
    onClick = () => {
      clearPreview();
      try {
        const exceptionFreeRows = selectExceptionFreeRows(
          props.workbook,
          {
            ...request,
            steps: editing === null ? steps : steps.slice(0, editing),
          },
          parseExamples(),
          parseFieldPositions(columns),
          transformOptions,
        );
        (setCandidates(exceptionFreeRows.candidates),
          setCandidate(""),
          setCandidatesKey(candidateSearchKey),
          setCandidateStatus(
            `${exceptionFreeRows.candidates.length} matching patterns from ${exceptionFreeRows.searched} candidates. ${exceptionFreeRows.truncated ? "Results are capped. " : ""}Search is bounded; choose a pattern, add the step, and review every resulting row.`,
          ));
      } catch (error) {
        recordPreviewError(error);
      }
    },
    buildStep = () => {
      const fieldPositions = parseFieldPositions(columns);
      if (operation === "trim")
        return { kind: operation, columns: fieldPositions, mode: trimMode };
      if (operation === "clean" || operation === "select")
        return { kind: operation, columns: fieldPositions };
      if (operation === "case")
        return { kind: operation, columns: fieldPositions, mode: caseMode };
      if (operation === "replace")
        return {
          kind: operation,
          columns: fieldPositions,
          find: find,
          replacement: replacement,
        };
      if (operation === "split") {
        if (fieldPositions.length !== 1)
          throw Error("Split one field at a time.");
        if (splitMode === "fixed" && !/^\s*\d+(?:\s*,\s*\d+)*\s*$/.test(breaks))
          throw Error(
            "Enter increasing fixed-width break positions separated by commas.",
          );
        const breakPositions =
          splitMode === "fixed" ? breaks.split(",").map(Number) : [];
        return {
          kind: operation,
          column: fieldPositions[0],
          mode: splitMode,
          delimiters: delimiters,
          collapse: shouldCollapse,
          breaks: breakPositions,
        };
      }
      if (operation === "concat")
        return {
          kind: operation,
          columns: fieldPositions,
          separator: separator,
          name: fieldName,
          skipBlank: shouldSkipBlank,
        };
      if (operation === "convert")
        return {
          kind: operation,
          columns: fieldPositions,
          to: convertTo,
          decimal: decimal,
          group: group,
          dateOrder: dateOrder,
        };
      if (operation === "fill") {
        let typedFillValue = fillValue;
        if (direction === "value" && fillType === "number") {
          if (!fillValue.trim() || !Number.isFinite(Number(fillValue)))
            throw Error("Enter a finite fill number.");
          typedFillValue = Number(fillValue);
        }
        if (direction === "value" && fillType === "boolean") {
          if (!/^(TRUE|FALSE)$/i.test(fillValue))
            throw Error("Enter TRUE or FALSE for Boolean fill.");
          typedFillValue = fillValue.toUpperCase() === "TRUE";
        }
        return {
          kind: operation,
          columns: fieldPositions,
          direction: direction,
          value: typedFillValue,
          groupColumns: parseFieldPositions(groupColumns, true),
        };
      }
      const selectedCandidate = visibleCandidates.find(
        (suggestion) => suggestion.id === candidate,
      );
      if (!selectedCandidate)
        throw Error(
          "Suggest patterns and explicitly choose one before adding the inferred step.",
        );
      return {
        kind: "pattern",
        name: fieldName,
        pattern: selectedCandidate.pattern,
        examples: parseExamples(),
        inferenceColumns: fieldPositions,
      };
    },
    saveStep = () => {
      clearPreview();
      try {
        const nextStep = buildStep(),
          nextSteps = [...steps];
        (editing === null
          ? nextSteps.push(nextStep)
          : (nextSteps[editing] = nextStep),
          buildTransformationPreview(
            props.workbook,
            { ...request, steps: nextSteps },
            transformOptions,
          ),
          setSteps(nextSteps),
          setEditing(null),
          setCandidates([]),
          setCandidate(""));
      } catch (error) {
        recordPreviewError(error);
      }
    },
    editStep = (stepIndex) => {
      clearPreview();
      const step = steps[stepIndex];
      (setEditing(stepIndex),
        setOperation(step.kind),
        setColumns(
          ("columns" in step
            ? step.columns
            : "column" in step
              ? [step.column]
              : step.inferenceColumns
          )
            .map((columnIndex) => columnIndex + 1)
            .join(", "),
        ),
        step.kind === "trim" && setTrimMode(step.mode),
        step.kind === "case" && setCaseMode(step.mode),
        step.kind === "replace" &&
          (setFind(step.find), setReplacement(step.replacement)),
        step.kind === "split" &&
          (setSplitMode(step.mode),
          setDelimiters(step.delimiters),
          setShouldCollapse(step.collapse),
          setBreaks(step.breaks.join(", "))),
        step.kind === "concat" &&
          (setFieldName(step.name),
          X$1(step.separator),
          setShouldSkipBlank(step.skipBlank)),
        step.kind === "convert" &&
          (setConvertTo(step.to),
          setDecimal(step.decimal),
          setGroup(step.group),
          setDateOrder(step.dateOrder)),
        step.kind === "fill" &&
          (setDirection(step.direction),
          setFillValue(step.value === null ? "" : String(step.value)),
          setFillType(
            typeof step.value == "number"
              ? "number"
              : typeof step.value == "boolean"
                ? "boolean"
                : "text",
          ),
          setGroupColumns(
            step.groupColumns.map((groupColumn) => groupColumn + 1).join(", "),
          )),
        step.kind === "pattern" &&
          (setFieldName(step.name),
          setExampleRows(
            step.examples.map((example) => example.row).join(", "),
          ),
          setExampleValues(
            step.examples.map((example) => example.value).join(`
`),
          ),
          setCandidates([]),
          setCandidate("")));
    },
    runPreview = () => {
      clearPreview();
      try {
        const transformationPreview = buildTransformationPreview(
          props.workbook,
          request,
          transformOptions,
        );
        (setOutputPage(0),
          setExceptionPage(0),
          setPlanState({ plan: transformationPreview, key: requestKey }));
      } catch (error) {
        recordPreviewError(error);
      }
    },
    commitPlan = () => {
      if (!(!currentPlan || props.readOnly || hasAppliedPlan.current))
        try {
          const nextWorkbook = commitApprovedTransformation(
            props.workbook,
            currentPlan,
            {
              ...transformOptions,
              approved: true,
              acceptExceptions: shouldAcceptExceptions,
            },
          );
          ((hasAppliedPlan.current = true),
            props.onWorkbookChange(
              nextWorkbook,
              `Clean & Reshape: ${currentPlan.sourceSheetName}!${currentPlan.request.sourceRange} → ${currentPlan.request.outputName} (${currentPlan.exceptions.length} retained parse exceptions)`,
            ),
            setPlanState(null));
        } catch (error) {
          recordPreviewError(error);
        }
    },
    downloadJson = (jsonText, fileName) => {
      const objectUrl = URL.createObjectURL(
          new Blob([jsonText], { type: "application/json" }),
        ),
        anchor = document.createElement("a");
      ((anchor.href = objectUrl),
        (anchor.download = fileName),
        anchor.click(),
        URL.revokeObjectURL(objectUrl));
    },
    accessibleTransforms = (props.workbook.fieldTransforms ?? []).filter(
      (transform) => {
        try {
          const sheet = props.workbook.sheets.find(
            (sheet) => sheet.id === transform.sourceSheetId,
          );
          return sheet
            ? (assertSelectionInAuthorizedSheet(
                sheet,
                transform.request.sourceRange,
                {
                  allowedSheetIds: allowedSheetIds,
                  includeHidden: transformOptions.includeHidden,
                  allowNewSheet: transformOptions.allowNewSheet,
                },
              ),
              true)
            : false;
        } catch {
          return false;
        }
      },
    );
  return i.jsx("div", {
    className: "fr-backdrop",
    children: i.jsxs("section", {
      className: "fr-dialog",
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button,input,select,textarea,summary",
            ) ?? []),
          ].filter(
            (element) =>
              !element.matches(":disabled") &&
              !element.hidden &&
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
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${dialogId}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Clean & Reshape",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Review typed results and exceptions before creating a separate worksheet.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Clean & Reshape"),
              onClick: props.onClose,
              lang: language("Close Clean & Reshape"),
              "data-ui-label": "Close Clean & Reshape",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fr-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Original cells stay intact. Steps run in order using the field positions from the preceding step. Inference never applies automatically.",
              }),
            }),
            props.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Read-only: saved evidence is available; recipe preview and Apply require write access.",
                }),
              }),
            i.jsxs("fieldset", {
              disabled: props.readOnly,
              className: "fr-fields",
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "Source and output",
                  }),
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Source worksheet" }),
                    i.jsx("select", {
                      "aria-label": text("Transform source worksheet"),
                      value: sourceSheetId,
                      onChange: (event) => setSourceSheetId(event.target.value),
                      lang: language("Transform source worksheet"),
                      "data-ui-label": "Transform source worksheet",
                      children: selectableSheets.map((sheet) =>
                        i.jsx(
                          "option",
                          { value: sheet.id, children: sheet.name },
                          sheet.id,
                        ),
                      ),
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Source range" }),
                    i.jsx("input", {
                      "aria-label": text("Transform source range"),
                      value: sourceRange,
                      onChange: (event) => setSourceRange(event.target.value),
                      lang: language("Transform source range"),
                      "data-ui-label": "Transform source range",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Output worksheet" }),
                    i.jsx("input", {
                      "aria-label": text("Transform output worksheet"),
                      value: outputName,
                      onChange: (event) => setOutputName(event.target.value),
                      lang: language("Transform output worksheet"),
                      "data-ui-label": "Transform output worksheet",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  className: "fr-check",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: hasHeader,
                      onChange: (event) => setHasHeader(event.target.checked),
                    }),
                    i.jsx(LocalizedText, {
                      children: "First source row contains headings",
                    }),
                  ],
                }),
                props.allowHidden &&
                  i.jsxs("label", {
                    className: "fr-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: checked,
                        onChange: (event) => setChecked(event.target.checked),
                      }),
                      i.jsx(LocalizedText, {
                        children: "Include hidden source cells",
                      }),
                    ],
                  }),
              ],
            }),
            i.jsxs("section", {
              "aria-label": text("Transformation recipe"),
              lang: language("Transformation recipe"),
              "data-ui-label": "Transformation recipe",
              children: [
                i.jsxs("h3", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Transformation steps (",
                    }),
                    steps.length,
                    "/12)",
                  ],
                }),
                !steps.length &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Add one or more steps, then preview the complete result.",
                    }),
                  }),
                i.jsx("ol", {
                  className: "fr-steps",
                  children: steps.map((step, stepIndex) =>
                    i.jsxs(
                      "li",
                      {
                        children: [
                          i.jsx("span", {
                            children: describeSplitFieldStep(step),
                          }),
                          i.jsxs("div", {
                            children: [
                              i.jsxs("button", {
                                disabled: props.readOnly,
                                onClick: () => editStep(stepIndex),
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Edit step ",
                                  }),
                                  stepIndex + 1,
                                ],
                              }),
                              i.jsx("button", {
                                disabled: props.readOnly || stepIndex === 0,
                                "aria-label": `Move step ${stepIndex + 1} up`,
                                onClick: () => {
                                  clearPreview();
                                  const reorderedSteps = [...steps];
                                  (([
                                    reorderedSteps[stepIndex - 1],
                                    reorderedSteps[stepIndex],
                                  ] = [
                                    reorderedSteps[stepIndex],
                                    reorderedSteps[stepIndex - 1],
                                  ]),
                                    setSteps(reorderedSteps),
                                    setEditing(null));
                                },
                                children: "↑",
                              }),
                              i.jsx("button", {
                                disabled:
                                  props.readOnly ||
                                  stepIndex === steps.length - 1,
                                "aria-label": `Move step ${stepIndex + 1} down`,
                                onClick: () => {
                                  clearPreview();
                                  const reorderedSteps = [...steps];
                                  (([
                                    reorderedSteps[stepIndex + 1],
                                    reorderedSteps[stepIndex],
                                  ] = [
                                    reorderedSteps[stepIndex],
                                    reorderedSteps[stepIndex + 1],
                                  ]),
                                    setSteps(reorderedSteps),
                                    setEditing(null));
                                },
                                children: "↓",
                              }),
                              i.jsxs("button", {
                                disabled: props.readOnly,
                                onClick: () => {
                                  (clearPreview(),
                                    setSteps(
                                      steps.filter(
                                        (step, index) => index !== stepIndex,
                                      ),
                                    ),
                                    setEditing(null));
                                },
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Remove step ",
                                  }),
                                  stepIndex + 1,
                                ],
                              }),
                            ],
                          }),
                        ],
                      },
                      stepIndex,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("fieldset", {
              disabled: props.readOnly,
              className: "fr-fields",
              children: [
                i.jsx("legend", {
                  children:
                    editing === null
                      ? "Add a step"
                      : `Edit step ${editing + 1}`,
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Operation" }),
                    i.jsx("select", {
                      "aria-label": text("Field operation"),
                      value: operation,
                      onChange: (event) => setOperation(event.target.value),
                      lang: language("Field operation"),
                      "data-ui-label": "Field operation",
                      children: textCleanupStepLabels.map(([kind, label]) =>
                        i.jsx("option", { value: kind, children: label }, kind),
                      ),
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Field positions" }),
                    i.jsx("input", {
                      "aria-label": text("Field positions"),
                      value: columns,
                      onChange: (event) => setColumns(event.target.value),
                      placeholder: "1, 3",
                      lang: language("Field positions"),
                      "data-ui-label": "Field positions",
                    }),
                  ],
                }),
                operation === "trim" &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Whitespace" }),
                      i.jsxs("select", {
                        "aria-label": text("Whitespace handling"),
                        value: trimMode,
                        onChange: (event) => setTrimMode(event.target.value),
                        lang: language("Whitespace handling"),
                        "data-ui-label": "Whitespace handling",
                        children: [
                          i.jsx("option", {
                            value: "collapse",
                            children: "Trim edges and collapse spaces",
                          }),
                          i.jsx("option", {
                            value: "edges",
                            children: "Trim edges only",
                          }),
                        ],
                      }),
                    ],
                  }),
                operation === "case" &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Case" }),
                      i.jsxs("select", {
                        "aria-label": text("Case conversion"),
                        value: caseMode,
                        onChange: (event) => setCaseMode(event.target.value),
                        lang: language("Case conversion"),
                        "data-ui-label": "Case conversion",
                        children: [
                          i.jsx("option", {
                            value: "upper",
                            children: "UPPER",
                          }),
                          i.jsx("option", {
                            value: "lower",
                            children: "lower",
                          }),
                          i.jsx("option", {
                            value: "proper",
                            children: "Proper",
                          }),
                          i.jsx("option", {
                            value: "none",
                            children: "Keep case",
                          }),
                        ],
                      }),
                    ],
                  }),
                operation === "replace" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Find literal text",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Find literal text"),
                            value: find,
                            onChange: (event) => setFind(event.target.value),
                            lang: language("Find literal text"),
                            "data-ui-label": "Find literal text",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Replace with" }),
                          i.jsx("input", {
                            "aria-label": text("Replacement text"),
                            value: replacement,
                            onChange: (event) =>
                              setReplacement(event.target.value),
                            lang: language("Replacement text"),
                            "data-ui-label": "Replacement text",
                          }),
                        ],
                      }),
                    ],
                  }),
                operation === "split" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Split mode" }),
                          i.jsxs("select", {
                            "aria-label": text("Field split mode"),
                            value: splitMode,
                            onChange: (event) =>
                              setSplitMode(event.target.value),
                            lang: language("Field split mode"),
                            "data-ui-label": "Field split mode",
                            children: [
                              i.jsx("option", {
                                value: "delimited",
                                children: "Delimited",
                              }),
                              i.jsx("option", {
                                value: "fixed",
                                children: "Fixed width",
                              }),
                            ],
                          }),
                        ],
                      }),
                      splitMode === "delimited"
                        ? i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Delimiter characters",
                                  }),
                                  i.jsx("input", {
                                    "aria-label": text("Delimiter characters"),
                                    value: delimiters,
                                    onChange: (event) =>
                                      setDelimiters(event.target.value),
                                    lang: language("Delimiter characters"),
                                    "data-ui-label": "Delimiter characters",
                                  }),
                                ],
                              }),
                              i.jsx("button", {
                                onClick: () => setDelimiters(delimiters + "	"),
                                children: i.jsx(LocalizedText, {
                                  children: "Add tab delimiter",
                                }),
                              }),
                              i.jsxs("label", {
                                className: "fr-check",
                                children: [
                                  i.jsx("input", {
                                    type: "checkbox",
                                    checked: shouldCollapse,
                                    onChange: (event) =>
                                      setShouldCollapse(event.target.checked),
                                  }),
                                  i.jsx(LocalizedText, {
                                    children: "Collapse consecutive delimiters",
                                  }),
                                ],
                              }),
                            ],
                          })
                        : i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Break after character positions",
                              }),
                              i.jsx("input", {
                                "aria-label": text("Fixed-width breaks"),
                                value: breaks,
                                onChange: (event) =>
                                  setBreaks(event.target.value),
                                lang: language("Fixed-width breaks"),
                                "data-ui-label": "Fixed-width breaks",
                              }),
                            ],
                          }),
                    ],
                  }),
                (operation === "concat" || operation === "pattern") &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "New field name" }),
                      i.jsx("input", {
                        "aria-label": text("New field name"),
                        value: fieldName,
                        onChange: (event) => setFieldName(event.target.value),
                        lang: language("New field name"),
                        "data-ui-label": "New field name",
                      }),
                    ],
                  }),
                operation === "concat" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Separator" }),
                          i.jsx("input", {
                            "aria-label": text("Concatenation separator"),
                            value: separator,
                            onChange: (event) => X$1(event.target.value),
                            lang: language("Concatenation separator"),
                            "data-ui-label": "Concatenation separator",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        className: "fr-check",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: shouldSkipBlank,
                            onChange: (event) =>
                              setShouldSkipBlank(event.target.checked),
                          }),
                          i.jsx(LocalizedText, {
                            children: "Skip blank fields",
                          }),
                        ],
                      }),
                    ],
                  }),
                operation === "convert" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Result type" }),
                          i.jsx("select", {
                            "aria-label": text("Conversion type"),
                            value: convertTo,
                            onChange: (event) =>
                              setConvertTo(event.target.value),
                            lang: language("Conversion type"),
                            "data-ui-label": "Conversion type",
                            children: [
                              "text",
                              "number",
                              "integer",
                              "boolean",
                              "date",
                            ].map((resultType) =>
                              i.jsx(
                                "option",
                                { value: resultType, children: resultType },
                                resultType,
                              ),
                            ),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Decimal separator",
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Transform decimal separator"),
                            value: decimal,
                            onChange: (event) => setDecimal(event.target.value),
                            lang: language("Transform decimal separator"),
                            "data-ui-label": "Transform decimal separator",
                            children: [
                              i.jsx("option", { value: ".", children: "Dot" }),
                              i.jsx("option", {
                                value: ",",
                                children: "Comma",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Grouping separator",
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Transform grouping separator"),
                            value: group,
                            onChange: (event) => setGroup(event.target.value),
                            lang: language("Transform grouping separator"),
                            "data-ui-label": "Transform grouping separator",
                            children: [
                              i.jsx("option", { value: "", children: "None" }),
                              i.jsx("option", {
                                value: ",",
                                children: "Comma",
                              }),
                              i.jsx("option", { value: ".", children: "Dot" }),
                              i.jsx("option", {
                                value: " ",
                                children: "Space",
                              }),
                              i.jsx("option", {
                                value: "'",
                                children: "Apostrophe",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Slash date order",
                          }),
                          i.jsxs("select", {
                            "aria-label": text("Transform date order"),
                            value: dateOrder,
                            onChange: (event) =>
                              setDateOrder(event.target.value),
                            lang: language("Transform date order"),
                            "data-ui-label": "Transform date order",
                            children: [
                              i.jsx("option", {
                                value: "MDY",
                                children: "Month/day/year",
                              }),
                              i.jsx("option", {
                                value: "DMY",
                                children: "Day/month/year",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                operation === "fill" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Fill direction" }),
                          i.jsxs("select", {
                            "aria-label": text("Fill direction"),
                            value: direction,
                            onChange: (event) =>
                              setDirection(event.target.value),
                            lang: language("Fill direction"),
                            "data-ui-label": "Fill direction",
                            children: [
                              i.jsx("option", {
                                value: "down",
                                children: "Down from previous value",
                              }),
                              i.jsx("option", {
                                value: "up",
                                children: "Up from next value",
                              }),
                              i.jsx("option", {
                                value: "value",
                                children: "Constant value",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Group field positions (optional)",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Fill group fields"),
                            value: groupColumns,
                            onChange: (event) =>
                              setGroupColumns(event.target.value),
                            lang: language("Fill group fields"),
                            "data-ui-label": "Fill group fields",
                          }),
                        ],
                      }),
                      direction === "value" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Fill type" }),
                                i.jsxs("select", {
                                  "aria-label": text("Fill value type"),
                                  value: fillType,
                                  onChange: (event) =>
                                    setFillType(event.target.value),
                                  lang: language("Fill value type"),
                                  "data-ui-label": "Fill value type",
                                  children: [
                                    i.jsx("option", {
                                      value: "text",
                                      children: "Text",
                                    }),
                                    i.jsx("option", {
                                      value: "number",
                                      children: "Number",
                                    }),
                                    i.jsx("option", {
                                      value: "boolean",
                                      children: "Boolean",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Fill value",
                                }),
                                i.jsx("input", {
                                  "aria-label": text("Fill value"),
                                  value: fillValue,
                                  onChange: (event) =>
                                    setFillValue(event.target.value),
                                  lang: language("Fill value"),
                                  "data-ui-label": "Fill value",
                                }),
                              ],
                            }),
                          ],
                        }),
                    ],
                  }),
                operation === "pattern" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Example worksheet rows",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Inference example rows"),
                            value: exampleRows,
                            onChange: (event) =>
                              setExampleRows(event.target.value),
                            lang: language("Inference example rows"),
                            "data-ui-label": "Inference example rows",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Expected text, one example per line",
                          }),
                          i.jsx("textarea", {
                            "aria-label": text("Inference example values"),
                            value: exampleValues,
                            onChange: (event) =>
                              setExampleValues(event.target.value),
                            rows: 3,
                            lang: language("Inference example values"),
                            "data-ui-label": "Inference example values",
                          }),
                        ],
                      }),
                      i.jsx("button", {
                        onClick: onClick,
                        children: i.jsx(LocalizedText, {
                          children: "Suggest patterns",
                        }),
                      }),
                      candidatesKey === candidateSearchKey &&
                        i.jsx("p", {
                          role: "status",
                          className: "fr-wide",
                          children: candidateStatus,
                        }),
                      visibleCandidates.length > 0 &&
                        i.jsxs("label", {
                          className: "fr-wide",
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Choose a matching pattern",
                            }),
                            i.jsxs("select", {
                              "aria-label": text("Inferred pattern"),
                              value: candidate,
                              onChange: (event) =>
                                setCandidate(event.target.value),
                              lang: language("Inferred pattern"),
                              "data-ui-label": "Inferred pattern",
                              children: [
                                i.jsx("option", {
                                  value: "",
                                  children: "Choose explicitly…",
                                }),
                                visibleCandidates.map((suggestion) =>
                                  i.jsx(
                                    "option",
                                    {
                                      value: suggestion.id,
                                      children: suggestion.description,
                                    },
                                    suggestion.id,
                                  ),
                                ),
                              ],
                            }),
                          ],
                        }),
                    ],
                  }),
                i.jsxs("div", {
                  className: "fr-wide fr-actions",
                  children: [
                    i.jsx("button", {
                      onClick: saveStep,
                      disabled: editing === null && steps.length >= 12,
                      children:
                        editing === null
                          ? "Add transformation step"
                          : "Save transformation step",
                    }),
                    editing !== null &&
                      i.jsx("button", {
                        onClick: () => {
                          (clearPreview(), setEditing(null));
                        },
                        children: i.jsx(LocalizedText, {
                          children: "Cancel step edit",
                        }),
                      }),
                  ],
                }),
              ],
            }),
            previewErrorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "fr-error",
                children: previewErrorMessage,
              }),
            planState &&
              !currentPlan &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Settings, source or permissions changed. Preview again before applying.",
                }),
              }),
            i.jsxs("div", {
              className: "fr-actions",
              children: [
                i.jsx("button", {
                  disabled: props.readOnly || !steps.length || editing !== null,
                  onClick: runPreview,
                  children: i.jsx(LocalizedText, {
                    children: "Preview transformation",
                  }),
                }),
                currentPlan &&
                  i.jsx("button", {
                    onClick: () => {
                      try {
                        downloadJson(
                          verifyTransformationPreview(
                            props.workbook,
                            currentPlan,
                            transformOptions,
                          ),
                          "Field-transformation-preview.json",
                        );
                      } catch (error) {
                        recordPreviewError(error);
                      }
                    },
                    children: i.jsx(LocalizedText, {
                      children: "Download transformation preview",
                    }),
                  }),
              ],
            }),
            currentPlan &&
              i.jsxs("section", {
                className: "fr-preview",
                "aria-label": text("Transformation preview"),
                lang: language("Transformation preview"),
                "data-ui-label": "Transformation preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Review before Apply",
                    }),
                  }),
                  i.jsxs("p", {
                    role: "status",
                    children: [
                      currentPlan.counts.sourceRows,
                      " ",
                      i.jsx(LocalizedText, { children: "source rows · " }),
                      currentPlan.counts.outputCells,
                      " ",
                      i.jsx(LocalizedText, { children: "output cells · " }),
                      currentPlan.counts.changedCells,
                      " ",
                      i.jsx(LocalizedText, { children: "changed values ·" }),
                      " ",
                      currentPlan.counts.exceptions,
                      " ",
                      i.jsx(LocalizedText, { children: "parse exceptions in" }),
                      " ",
                      currentPlan.counts.exceptionRows,
                      " ",
                      i.jsx(LocalizedText, { children: "rows." }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Original: " }),
                      currentPlan.sourceSheetName,
                      "!",
                      currentPlan.request.sourceRange,
                      i.jsx(LocalizedText, { children: ", revision " }),
                      props.workbook.revision,
                      i.jsx(LocalizedText, { children: ". Output:" }),
                      " ",
                      currentPlan.request.outputName,
                      "!",
                      currentPlan.targetRange,
                      ".",
                    ],
                  }),
                  i.jsxs("details", {
                    children: [
                      i.jsx("summary", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Step results and resulting field positions",
                        }),
                      }),
                      currentPlan.stages.map((stage) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsx(LocalizedText, { children: "Step " }),
                              stage.step,
                              " (",
                              stage.kind,
                              "):",
                              " ",
                              stage.fields
                                .map(
                                  (field, fieldIndex) =>
                                    `${fieldIndex + 1}: ${field}`,
                                )
                                .join(" · "),
                              ".",
                              " ",
                              stage.changedCells,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "affected cells; ",
                              }),
                              stage.exceptions,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "new exceptions.",
                              }),
                            ],
                          },
                          stage.step,
                        ),
                      ),
                    ],
                  }),
                  i.jsx("div", {
                    className: "fr-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("caption", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Before/after values and original source locators",
                          }),
                        }),
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Output",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Field",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Original source / raw value",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "After",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Type",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Review",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: currentPlan.outputs
                            .slice(outputPage * 20, (outputPage + 1) * 20)
                            .map((output) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("td", { children: output.address }),
                                    i.jsx("td", { children: output.field }),
                                    i.jsx("td", {
                                      children: i.jsxs("details", {
                                        children: [
                                          i.jsx("summary", {
                                            children: output.sources
                                              .map(
                                                (source) =>
                                                  `${currentPlan.sourceSheetName}!${source}`,
                                              )
                                              .join(", "),
                                          }),
                                          output.sources.map((source) =>
                                            i.jsxs(
                                              "pre",
                                              {
                                                children: [
                                                  source,
                                                  ":",
                                                  " ",
                                                  stringifyJsonValue(
                                                    currentPlan.sources.find(
                                                      (sourceEntry) =>
                                                        sourceEntry.address ===
                                                        source,
                                                    )?.original?.raw ?? null,
                                                  ),
                                                ],
                                              },
                                              source,
                                            ),
                                          ),
                                        ],
                                      }),
                                    }),
                                    i.jsx("td", {
                                      children: i.jsx("code", {
                                        children: stringifyJsonValue(
                                          output.value,
                                        ),
                                      }),
                                    }),
                                    i.jsx("td", { children: output.type }),
                                    i.jsx("td", {
                                      children: output.exception
                                        ? "Exception retained"
                                        : "Ready",
                                    }),
                                  ],
                                },
                                output.address,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "fr-actions",
                    children: [
                      i.jsx("button", {
                        disabled: outputPage === 0,
                        onClick: () => setOutputPage(outputPage - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous results",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Results " }),
                          outputPage * 20 + 1,
                          "–",
                          Math.min(
                            (outputPage + 1) * 20,
                            currentPlan.outputs.length,
                          ),
                          " ",
                          i.jsx(LocalizedText, { children: "of" }),
                          " ",
                          currentPlan.outputs.length,
                        ],
                      }),
                      i.jsx("button", {
                        disabled:
                          (outputPage + 1) * 20 >= currentPlan.outputs.length,
                        onClick: () => setOutputPage(outputPage + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next results",
                        }),
                      }),
                    ],
                  }),
                  currentPlan.exceptions.length > 0 &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(LocalizedText, {
                            children: "Retained parse exceptions",
                          }),
                        }),
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Raw values remain in the derived output and an exceptions appendix. Later steps do not convert a failed field. Source cells remain unchanged.",
                          }),
                        }),
                        i.jsx("div", {
                          className: "fr-scroll",
                          children: i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Row / field",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Step",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Raw value",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(LocalizedText, {
                                        children: "Source",
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
                                children: currentPlan.exceptions
                                  .slice(
                                    exceptionPage * 20,
                                    (exceptionPage + 1) * 20,
                                  )
                                  .map((cell, exceptionIndex) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              cell.row,
                                              " / ",
                                              cell.field,
                                            ],
                                          }),
                                          i.jsx("td", { children: cell.step }),
                                          i.jsx("td", {
                                            children: i.jsx("code", {
                                              children: stringifyJsonValue(
                                                cell.raw,
                                              ),
                                            }),
                                          }),
                                          i.jsx("td", {
                                            children: cell.sources.join(", "),
                                          }),
                                          i.jsx("td", {
                                            children: cell.reason,
                                          }),
                                        ],
                                      },
                                      exceptionIndex,
                                    ),
                                  ),
                              }),
                            ],
                          }),
                        }),
                        i.jsxs("div", {
                          className: "fr-actions",
                          children: [
                            i.jsx("button", {
                              disabled: !exceptionPage,
                              onClick: () =>
                                setExceptionPage(exceptionPage - 1),
                              children: i.jsx(LocalizedText, {
                                children: "Previous exceptions",
                              }),
                            }),
                            i.jsxs("span", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Exceptions ",
                                }),
                                exceptionPage * 20 + 1,
                                "–",
                                Math.min(
                                  (exceptionPage + 1) * 20,
                                  currentPlan.exceptions.length,
                                ),
                                " ",
                                i.jsx(LocalizedText, { children: "of " }),
                                currentPlan.exceptions.length,
                              ],
                            }),
                            i.jsx("button", {
                              disabled:
                                (exceptionPage + 1) * 20 >=
                                currentPlan.exceptions.length,
                              onClick: () =>
                                setExceptionPage(exceptionPage + 1),
                              children: i.jsx(LocalizedText, {
                                children: "Next exceptions",
                              }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          className: "fr-check",
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: shouldAcceptExceptions,
                              onChange: (event) =>
                                setShouldAcceptExceptions(event.target.checked),
                            }),
                            i.jsx(LocalizedText, {
                              children:
                                "Keep raw values and the exceptions appendix",
                            }),
                          ],
                        }),
                      ],
                    }),
                ],
              }),
            accessibleTransforms.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Saved transformation evidence (",
                      }),
                      accessibleTransforms.length,
                      ")",
                    ],
                  }),
                  accessibleTransforms.map((savedTransform) =>
                    i.jsxs(
                      "article",
                      {
                        className: "fr-record",
                        children: [
                          i.jsx("h3", {
                            children: savedTransform.outputSheetName,
                          }),
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, { children: "Source " }),
                              savedTransform.sourceSheetName,
                              "!",
                              savedTransform.request.sourceRange,
                              " ",
                              i.jsx(LocalizedText, {
                                children: "· original revision ",
                              }),
                              savedTransform.sourceRevision,
                              " · ",
                              savedTransform.createdAt,
                              " · ",
                              savedTransform.exceptions.length,
                              " ",
                              i.jsx(LocalizedText, {
                                children:
                                  "exceptions. These are historical snapshot locators.",
                              }),
                            ],
                          }),
                          i.jsx("ol", {
                            children: savedTransform.request.steps.map(
                              (step, stepIndex) =>
                                i.jsx(
                                  "li",
                                  { children: describeSplitFieldStep(step) },
                                  stepIndex,
                                ),
                            ),
                          }),
                          i.jsxs("button", {
                            onClick: () =>
                              downloadJson(
                                JSON.stringify(savedTransform, null, 2),
                                "Field-transformation-record.json",
                              ),
                            children: [
                              i.jsx(LocalizedText, { children: "Download " }),
                              savedTransform.outputSheetName,
                              " ",
                              i.jsx(LocalizedText, { children: "evidence" }),
                            ],
                          }),
                          i.jsxs("button", {
                            disabled: props.readOnly,
                            onClick: () => {
                              (clearPreview(),
                                setSourceSheetId(
                                  savedTransform.request.sourceSheetId,
                                ),
                                setSourceRange(
                                  savedTransform.request.sourceRange,
                                ),
                                setHasHeader(savedTransform.request.header),
                                setSteps(
                                  structuredClone(savedTransform.request.steps),
                                ),
                                setOutputName(
                                  `${savedTransform.outputSheetName.slice(0, 26)} copy`,
                                ),
                                setEditing(null),
                                setCandidates([]),
                                setCandidate(""));
                            },
                            children: [
                              i.jsx(LocalizedText, { children: "Reuse " }),
                              savedTransform.outputSheetName,
                              " ",
                              i.jsx(LocalizedText, { children: "recipe" }),
                            ],
                          }),
                        ],
                      },
                      savedTransform.id,
                    ),
                  ),
                ],
              }),
            i.jsx("p", {
              className: "fr-boundary",
              children: i.jsx(LocalizedText, {
                children:
                  "Up to 500 source cells, 12 steps and 5,000 output cells including exceptions. Sources must be reviewed values. Patterns are deterministic and bounded. Native JSON retains recipes and original evidence; XLSX retains current cells and notes. Local history is editable outside this workflow.",
              }),
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "fr-primary",
              disabled:
                !currentPlan ||
                props.readOnly ||
                (currentPlan.exceptions.length > 0 &&
                  !shouldAcceptExceptions) ||
                editing !== null,
              onClick: commitPlan,
              children: i.jsx(LocalizedText, {
                children: "Apply reviewed transformation",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const sortEngineLimits = Object.freeze({
    rows: 1e4,
    cells: 1e5,
    keys: 64,
    customItems: 100,
    formulaCells: 5e3,
    formulaCharacters: 1e6,
    rangeRewriteWork: 5e6,
    pageSize: 25,
  }),
  stableSortPolicy = Object.freeze({
    ties: "Stable ties retain their physical source-row order at preview time, in both directions.",
    records:
      "Whole worksheet records move across every populated column in the chosen row interval. A partial selection requires expansion; cancelling preserves the source.",
    types:
      "Automatic uses numbers, text, booleans, errors, then blanks. Descending reverses nonblank order; blanks stay last. Text uses en-US collation with case sensitivity selectable. Numeric-looking text stays text unless Number is explicitly selected. Date accepts numeric serials or the workbook's strict ISO/MDY/DMY policy.",
    references:
      "Direct references follow record identities, including absolute markers. Contiguous range membership is preserved; a range that would split is refused. Order-sensitive calculations may change after sorting.",
    scope:
      "All source rows in the range participate, including empty-key rows. Hidden rows require explicit inclusion. Temporary view filters are separate. Color/icon and left-to-right sorts are outside this contract.",
  }),
  MW = new WeakMap(),
  formatSortRangeAddress = (range) => {
    const rangeAddress = parseRangeAddress(range),
      cellAddress = formatCellAddress(
        rangeAddress.firstRow,
        rangeAddress.firstCol,
      ),
      lastCellAddress = formatCellAddress(
        rangeAddress.lastRow,
        rangeAddress.lastCol,
      );
    return cellAddress === lastCellAddress
      ? cellAddress
      : `${cellAddress}:${lastCellAddress}`;
  },
  isNonBlankCellEntry = (cell) =>
    cell.raw !== null ||
    Object.keys(cell).some((propertyName) => propertyName !== "raw"),
  N3 = (cell) =>
    !cell.literal && typeof cell.raw == "string" && cell.raw.startsWith("="),
  buildAllowedSheetKey = (workbook, accessPolicy) =>
    JSON.stringify(
      [
        ...(accessPolicy.allowedSheetIds ??
          workbook.sheets.map((sheet) => sheet.id)),
      ].sort(),
    );
function assertSortWorkbookAuthorized(workbook, sheetId, accessPolicy) {
  if (accessPolicy.readOnly)
    throw Error(
      "This workbook is read-only. Sorting cannot change its records.",
    );
  if (
    accessPolicy.allowedSheetIds &&
    !accessPolicy.allowedSheetIds.includes(sheetId)
  )
    throw Error("The sort source is outside the host-approved worksheets.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  if (sheetOrThrow.hidden || sheetOrThrow.veryHidden)
    throw Error("Unhide this worksheet before sorting its records.");
  return (assertSheetUnprotected(sheetOrThrow), sheetOrThrow);
}
function resolveSortSourceRange(workbook, sheetId, range, headers = false) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(range);
  if (
    rangeAddress.lastRow >= sheetOrThrow.rowCount ||
    rangeAddress.lastCol >= sheetOrThrow.colCount
  )
    throw Error("The sort selection is outside the worksheet.");
  let firstRow = rangeAddress.firstRow,
    lastRow = rangeAddress.lastRow,
    firstCol = rangeAddress.firstCol,
    lastCol = rangeAddress.lastCol;
  const populatedCells = Object.entries(sheetOrThrow.cells)
    .filter(([, cell]) => isNonBlankCellEntry(cell))
    .map(([cellAddress]) => _t(cellAddress));
  if (firstRow === lastRow && firstCol === lastCol) {
    const containingTable = sheetOrThrow.tables
      ?.map((table) => ({
        bounds: parseRangeAddress(table.range),
        totals: table.showTotals,
      }))
      .find(
        (tableEntry) =>
          firstRow >= tableEntry.bounds.firstRow &&
          firstRow <= tableEntry.bounds.lastRow &&
          firstCol >= tableEntry.bounds.firstCol &&
          firstCol <= tableEntry.bounds.lastCol,
      );
    if (containingTable)
      ((firstRow = containingTable.bounds.firstRow),
        (lastRow =
          containingTable.bounds.lastRow - (containingTable.totals ? 1 : 0)),
        (firstCol = containingTable.bounds.firstCol),
        (lastCol = containingTable.bounds.lastCol));
    else {
      const populatedRows = new Set(
        populatedCells.map((cellPosition) => cellPosition.row),
      );
      if (!populatedRows.has(firstRow))
        throw Error("Select a populated record or an explicit range to sort.");
      for (; firstRow > 0 && populatedRows.has(firstRow - 1);) firstRow--;
      for (
        ;
        lastRow + 1 < sheetOrThrow.rowCount && populatedRows.has(lastRow + 1);
      )
        lastRow++;
    }
  }
  for (const populatedCell of populatedCells)
    populatedCell.row >= firstRow &&
      populatedCell.row <= lastRow &&
      ((firstCol = Math.min(firstCol, populatedCell.col)),
      (lastCol = Math.max(lastCol, populatedCell.col)));
  const sortRangeAddress = formatSortRangeAddress(
      `${formatCellAddress(firstRow, firstCol)}:${formatCellAddress(lastRow, lastCol)}`,
    ),
    selectedRange = formatSortRangeAddress(range);
  return {
    selectedRange: selectedRange,
    expandedRange: sortRangeAddress,
    expansionRequired: selectedRange !== sortRangeAddress,
    firstDataRow: firstRow + (headers ? 1 : 0),
    lastRow: lastRow,
  };
}
function normalizeSortSpecification(workbook, specification) {
  if (
    !specification ||
    typeof specification.headers != "boolean" ||
    !Array.isArray(specification.keys) ||
    !specification.keys.length ||
    specification.keys.length > sortEngineLimits.keys ||
    [
      specification.caseSensitive,
      specification.expandSelection,
      specification.includeHidden,
    ].some(
      (optionValue) =>
        optionValue !== undefined && typeof optionValue != "boolean",
    )
  )
    throw Error("Choose headers and 1–64 sort levels with valid options.");
  const normalizedSpecification = structuredClone(specification);
  normalizedSpecification.range = formatSortRangeAddress(specification.range);
  const seenColumns = new Set();
  for (const key of normalizedSpecification.keys) {
    if (
      !Number.isInteger(key.column) ||
      key.column < 0 ||
      key.column >=
        findSheetOrThrow(workbook, specification.sheetId).colCount ||
      seenColumns.has(key.column) ||
      !["ascending", "descending"].includes(key.direction) ||
      !["auto", "number", "date", "text"].includes(key.type)
    )
      throw Error(
        "Choose distinct valid sort columns, directions and data types.",
      );
    if ((seenColumns.add(key.column), key.customOrder !== undefined)) {
      const normalizeLabel = (label) =>
        normalizedSpecification.caseSensitive
          ? label.normalize("NFC")
          : label.normalize("NFC").toLocaleLowerCase("en-US");
      if (
        !Array.isArray(key.customOrder) ||
        !key.customOrder.length ||
        key.customOrder.length > sortEngineLimits.customItems ||
        key.type !== "text" ||
        key.customOrder.some(
          (customLabel) =>
            typeof customLabel != "string" ||
            !customLabel.length ||
            customLabel.length > 255,
        ) ||
        new Set(key.customOrder.map(normalizeLabel)).size !==
          key.customOrder.length
      )
        throw Error(
          "Custom orders require Text, 1–100 distinct nonempty labels of at most 255 characters.",
        );
    }
  }
  return normalizedSpecification;
}
function buildSortPlan(workbook, specification, sourceRange) {
  const rangeAddress = parseRangeAddress(sourceRange.expandedRange),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    datePolicy = normalizeDatePolicy(workbook.datePolicy);
  return Array.from(
    { length: sourceRange.lastRow - sourceRange.firstDataRow + 1 },
    (s, rowOffset) => {
      const row = sourceRange.firstDataRow + rowOffset,
        keys = specification.keys.map((key) => {
          if (
            key.column < rangeAddress.firstCol ||
            key.column > rangeAddress.lastCol
          )
            throw Error(
              "Every sort key must be inside the complete record range.",
            );
          const cellAddress = formatCellAddress(row, key.column),
            cell = findSheetOrThrow(workbook, specification.sheetId).cells[
              cellAddress
            ],
            calculatedValue = workbookWithSettings(
              specification.sheetId,
              cellAddress,
            );
          if (
            cell &&
            N3(cell) &&
            typeof calculatedValue == "string" &&
            ["#PENDING!", "#TIMEOUT!", "#LIMIT!", "#BLOCKED!"].includes(
              calculatedValue,
            )
          )
            throw Error(
              `Calculate ${cellAddress} successfully before sorting. Its current result is ${calculatedValue}.`,
            );
          const isErrorValue =
            !!cell &&
            N3(cell) &&
            workbookWithSettings.getTypedValue(
              specification.sheetId,
              cellAddress,
            ).type === "error";
          let compare = calculatedValue;
          if (
            calculatedValue !== null &&
            calculatedValue !== "" &&
            !isErrorValue
          ) {
            if (
              (key.type === "text" && (compare = String(calculatedValue)),
              key.type === "number")
            ) {
              if (typeof calculatedValue == "number") compare = calculatedValue;
              else if (
                typeof calculatedValue == "string" &&
                /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(
                  calculatedValue.trim(),
                )
              )
                compare = Number(calculatedValue);
              else
                throw Error(
                  `${cellAddress} cannot be interpreted as a Number. Choose Automatic or correct the source.`,
                );
              if (!Number.isFinite(compare))
                throw Error(
                  `${cellAddress} is outside the finite numeric range.`,
                );
            }
            if (
              key.type === "date" &&
              ((compare =
                typeof calculatedValue == "number"
                  ? calculatedValue
                  : typeof calculatedValue == "string"
                    ? (parseDateText(calculatedValue, datePolicy) ?? null)
                    : null),
              compare === null ||
                typeof compare != "number" ||
                compare < 0 ||
                compare >= dateSystemMaxSerial(datePolicy.dateSystem) + 1)
            )
              throw Error(
                `${cellAddress} is not a valid date under the workbook's date policy.`,
              );
          }
          return {
            value: calculatedValue,
            compare: compare,
            error: isErrorValue,
          };
        });
      return { row: row, keys: keys };
    },
  );
}
function assertSortSafeFormulas(workbook, dataRowCount, accessPolicy) {
  let formulaCellCount = 0,
    formulaCharacterCount = 0,
    rangeReferenceCount = 0;
  const reviewFormula = (formulaText) => {
    (formulaCellCount++, (formulaCharacterCount += formulaText.length));
    const formulaWithoutLiterals = formulaText.replace(
      /"(?:[^"]|"")*"|'(?:[^']|'')*'/g,
      "",
    );
    if (/\b(?:INDIRECT|OFFSET)\s*\(/i.test(formulaWithoutLiterals))
      throw Error(
        "Sorting cannot safely rewrite dynamic INDIRECT/OFFSET references. Replace them with explicit references first.",
      );
    if (
      /[{}#]/.test(formulaWithoutLiterals) ||
      matrixResultFunctionNames.some((matrixResultFunctionName) =>
        new RegExp(
          `\\b${matrixResultFunctionName.replaceAll(".", "\\.")}\\s*\\(`,
          "i",
        ).test(formulaWithoutLiterals),
      )
    )
      throw Error(
        "Sorting cannot safely relocate array/spill formulas or existing reference-error tokens. Resolve or separate those formulas first.",
      );
    rangeReferenceCount += (formulaWithoutLiterals.match(/:/g) ?? []).length;
  };
  for (const sheet of workbook.sheets) {
    for (const cell of Object.values(sheet.cells))
      N3(cell) && reviewFormula(cell.raw);
    for (const sheetRule of [
      ...(sheet.validations ?? []),
      ...(sheet.conditionalFormats ?? []),
    ])
      sheetRule.formula && reviewFormula(sheetRule.formula);
  }
  for (const definedName of workbook.definedNames ?? [])
    reviewFormula(definedName.refersTo);
  if (
    formulaCellCount > sortEngineLimits.formulaCells ||
    formulaCharacterCount > sortEngineLimits.formulaCharacters ||
    rangeReferenceCount * dataRowCount > sortEngineLimits.rangeRewriteWork
  )
    throw Error(
      "The sort exceeds the bounded formula-reference review budget. Reduce or separate the workbook before sorting.",
    );
  if (
    accessPolicy.allowedSheetIds &&
    workbook.sheets.some(
      (sheet) => !accessPolicy.allowedSheetIds.includes(sheet.id),
    )
  )
    throw Error(
      "A structural sort requires the host to authorize all worksheets so cross-sheet reference changes can be reviewed.",
    );
}

export {
  assertSortWorkbookAuthorized,
  buildAllowedSheetKey,
  buildSortPlan,
  MW,
  N3,
  stableSortPolicy,
  resolveSortSourceRange,
  assertSortSafeFormulas,
  isNonBlankCellEntry,
  CleanAndReshapeDialog,
  normalizeSortSpecification,
  sortEngineLimits,
};
export { formatSortRangeAddress };
