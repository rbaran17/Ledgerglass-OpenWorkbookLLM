import { LabeledField } from "./14-dax-data-model-engine.js";

import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import c_ from "react-dom/server";
import {
  deepFreezeValue,
  fingerprintCanonicalValue,
} from "./01-schemas-date-functions-and-validation.js";
import {
  assertWorkbookStructureEditable,
  formatNumberWithPattern,
  validateObjectReportRequest,
  reportOutputFormatVersion,
  parseReportRange,
  inspectPictureDataUrl,
  formatOccupiedRangeAddress,
  reportOutputLimits,
  hashRecordExcludingHash,
  parseReportSnapshot,
  inspectChartSource,
  _t,
  formatDiagnosticsMessage,
  cn,
  parseRangeAddress,
  applyFilterConditions,
  normalizeWorkbookInput,
  columnIndexToLetters,
  resolveEffectiveCellStyle,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { readSourceRectangles } from "./12-consolidation-and-dependency-inspection.js";
import { ChartReportSvg } from "./31-pivot-explorer-and-charts.js";
import { normalizeAnalysisPermissions } from "./36-watch-window-and-goal-seek.js";
import {
  readSavedAnalysisReport,
  commitAnalysisResult,
  runAnalysisRequest,
  parseNumericText,
} from "./37-statistics-and-linear-solver.js";

function DataAnalysisDialog(props) {
  const { text } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    analysisReport = sheet.analysisReport,
    [analysisKind, setAnalysisKind] = ne.useState(
      analysisReport?.request.kind ?? props.initialKind ?? "statistics",
    ),
    [activeSection, setActiveSection] = ne.useState("Setup"),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [preview, setPreview] = ne.useState(null),
    [checked, setChecked] = ne.useState(analysisReport?.includeHidden ?? false),
    [resultPage, setResultPage] = ne.useState(0),
    defaultOutputName = () => {
      let analysisNumber = 1;
      for (
        ;
        props.workbook.sheets.some(
          (sheet) => sheet.name.toLowerCase() === `analysis ${analysisNumber}`,
        );
      )
        analysisNumber++;
      return `Analysis ${analysisNumber}`;
    },
    [outputName, setOutputName] = ne.useState(defaultOutputName),
    [solverRequest, setSolverRequest] = ne.useState(() =>
      analysisReport?.request.kind === "solver"
        ? analysisReport.request
        : {
            kind: "solver",
            outputName: outputName,
            objective: {
              sheetId: props.sheetId,
              address:
                Object.entries(sheet.cells).find(
                  ([, cell]) =>
                    !cell.literal &&
                    typeof cell.raw == "string" &&
                    cell.raw.startsWith("="),
                )?.[0] ?? "C2",
            },
            direction: "maximize",
            variables: Object.entries(sheet.cells)
              .filter(([, cell]) => typeof cell.raw == "number")
              .slice(0, 2)
              .map(([address]) => ({
                cell: { sheetId: props.sheetId, address: address },
                lower: 0,
                upper: 100,
                domain: "continuous",
              })),
            constraints: [],
            tolerance: 1e-8,
            maxCandidates: 5e4,
          },
    ),
    [statisticsRequest, setStatisticsRequest] = ne.useState(() =>
      analysisReport?.request.kind === "statistics"
        ? analysisReport.request
        : {
            kind: "statistics",
            outputName: outputName,
            sheetId: props.sheetId,
            range: props.range.includes(":")
              ? props.range
              : formatOccupiedRangeAddress(sheet),
            method: "descriptive",
            missing: "reject",
          },
    ),
    [binsText, setBinsText] = ne.useState(
      statisticsRequest.bins?.join(", ") ?? "10, 20, 30",
    ),
    [inputRevision, setInputRevision] = ne.useState(0),
    [previewRevision, setPreviewRevision] = ne.useState(-1),
    [previewWorkbook, setPreviewWorkbook] = ne.useState(null),
    [previewPermissionKey, setPreviewPermissionKey] = ne.useState(""),
    permissionKey = JSON.stringify([
      props.readOnly,
      props.allowedSheetIds,
      !!props.allowHidden && checked,
    ]),
    isPreviewCurrent =
      !!preview &&
      previewRevision === inputRevision &&
      previewWorkbook === props.workbook &&
      previewPermissionKey === permissionKey,
    analysisOptions = {
      readOnly: props.readOnly,
      allowedSheetIds: props.allowedSheetIds,
      includeHidden: !!props.allowHidden && checked,
    },
    invalidatePreview = () => {
      (setInputRevision((previousRevision) => previousRevision + 1),
        setPreview(null),
        setErrorMessage(""),
        setResultPage(0));
    },
    updateSolverRequest = (solverPatch) => {
      (setSolverRequest((previousSolverRequest) => ({
        ...previousSolverRequest,
        ...solverPatch,
      })),
        invalidatePreview());
    },
    updateStatisticsRequest = (statisticsPatch) => {
      (setStatisticsRequest((previousStatisticsRequest) => ({
        ...previousStatisticsRequest,
        ...statisticsPatch,
      })),
        invalidatePreview());
    },
    selectableSheets = props.workbook.sheets.filter(
      (sheet) =>
        !sheet.analysisReport &&
        !sheet.dataModelReport &&
        (!props.allowedSheetIds || props.allowedSheetIds.includes(sheet.id)) &&
        (analysisOptions.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    X$1 = (selectedSheetId, onSheetChange, selectLabel) =>
      i.jsx("select", {
        "aria-label": text(selectLabel),
        value: selectedSheetId,
        onChange: (event) => onSheetChange(event.target.value),
        children: selectableSheets.map((worksheet) =>
          i.jsx(
            "option",
            { value: worksheet.id, children: worksheet.name },
            worksheet.id,
          ),
        ),
      }),
    renderNumberField = (fieldLabel, fieldValue, onFieldChange) =>
      i.jsx(LocalizedNumberField, {
        "aria-label": text(fieldLabel),
        value:
          fieldValue === undefined || !Number.isFinite(fieldValue)
            ? ""
            : fieldValue,
        onValueChange: (numericText) =>
          onFieldChange(parseNumericText(numericText)),
      });
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      const modelBody = dialogRef.current?.querySelector(".lg-model-body");
      modelBody && (modelBody.scrollTop = 0);
    }, [activeSection, preview]));
  const onClick = () => {
      (setPreview(null), setErrorMessage(""));
      try {
        const request =
            analysisKind === "solver"
              ? { ...solverRequest, outputName: outputName }
              : {
                  ...statisticsRequest,
                  outputName: outputName,
                  ...(statisticsRequest.method === "histogram"
                    ? { bins: binsText.split(",").map(parseNumericText) }
                    : {}),
                },
          nextPreview = runAnalysisRequest(
            props.workbook,
            request,
            analysisOptions,
          );
        (setPreview(nextPreview),
          setPreviewRevision(inputRevision),
          setPreviewWorkbook(props.workbook),
          setPreviewPermissionKey(permissionKey),
          setActiveSection("Results"),
          setResultPage(0));
      } catch (error) {
        setErrorMessage(
          error instanceof Error ? error.message : "Analysis failed.",
        );
      }
    },
    saveAnalysis = (commitMode) => {
      try {
        if (!isPreviewCurrent || !preview)
          throw Error("Preview the current inputs and permissions first.");
        const nextWorkbook = commitAnalysisResult(
          props.workbook,
          preview,
          commitMode,
          analysisOptions,
        );
        (props.onWorkbookChange(
          nextWorkbook,
          commitMode === "report"
            ? "Save analysis report"
            : "Apply Solver solution and report",
        ),
          setPreview(null),
          props.onClose());
      } catch (error) {
        setErrorMessage(
          error instanceof Error
            ? error.message
            : "The analysis was not saved.",
        );
      }
    };
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "lg-data-model lg-analysis",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...dialogRef.current.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary,[tabindex="0"]',
              ),
            ].filter(
              (element) =>
                !element.hidden &&
                (element.getClientRects().length ||
                  navigator.userAgent.includes("jsdom")),
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
            i.jsxs("div", {
              children: [
                i.jsx("p", {
                  className: "eyebrow",
                  children: i.jsx(LocalizedText, {
                    children: "Data · Analysis",
                  }),
                }),
                i.jsx("h2", {
                  id: titleId,
                  children: i.jsx(LocalizedText, {
                    children:
                      analysisKind === "solver" ? "Solver" : "Data Analysis",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      analysisKind === "solver"
                        ? "Define a bounded linear model, review the result, then choose whether to apply it."
                        : "Choose a statistical method and review the sample before saving a report.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close analysis"),
              onClick: props.onClose,
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("nav", {
          "aria-label": text("Analysis sections"),
          children: [
            ["Setup", "Results"].map((section) =>
              i.jsx(
                "button",
                {
                  "aria-pressed": activeSection === section,
                  onClick: () => setActiveSection(section),
                  children: i.jsx(LocalizedText, { children: section }),
                },
                section,
              ),
            ),
            i.jsx("span", {
              className: "lg-analysis-mode",
              children: LabeledField(
                "Tool",
                i.jsxs("select", {
                  value: analysisKind,
                  onChange: (event) => {
                    (setAnalysisKind(event.target.value),
                      invalidatePreview(),
                      setActiveSection("Setup"));
                  },
                  children: [
                    i.jsx("option", { value: "solver", children: "Solver" }),
                    i.jsx("option", {
                      value: "statistics",
                      children: "Data Analysis",
                    }),
                  ],
                }),
              ),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-model-body",
          children: [
            analysisReport &&
              i.jsx("p", {
                className: "lg-analysis-hint",
                children: i.jsx(LocalizedText, {
                  children:
                    "Loaded the saved definition. A new preview and a new report name are required to refresh it.",
                }),
              }),
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "lg-analysis-error",
                children: errorMessage,
              }),
            activeSection === "Setup"
              ? i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("div", {
                      className: "lg-model-controls",
                      children: [
                        LabeledField(
                          "Output worksheet",
                          i.jsx("input", {
                            value: outputName,
                            onChange: (event) => {
                              (setOutputName(event.target.value),
                                invalidatePreview());
                            },
                          }),
                        ),
                        props.allowHidden &&
                          LabeledField(
                            "Include hidden source cells",
                            i.jsx("input", {
                              type: "checkbox",
                              checked: checked,
                              onChange: (event) => {
                                (setChecked(event.target.checked),
                                  invalidatePreview());
                              },
                            }),
                          ),
                      ],
                    }),
                    analysisKind === "solver"
                      ? i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Objective",
                                  }),
                                }),
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    LabeledField(
                                      "Objective worksheet",
                                      X$1(
                                        solverRequest.objective.sheetId,
                                        (sheetId) =>
                                          updateSolverRequest({
                                            objective: {
                                              ...solverRequest.objective,
                                              sheetId: sheetId,
                                            },
                                          }),
                                        "Objective worksheet",
                                      ),
                                    ),
                                    LabeledField(
                                      "Objective formula cell",
                                      i.jsx("input", {
                                        value: solverRequest.objective.address,
                                        onChange: (event) =>
                                          updateSolverRequest({
                                            objective: {
                                              ...solverRequest.objective,
                                              address: event.target.value
                                                .replace(/\$/g, "")
                                                .toUpperCase(),
                                            },
                                          }),
                                      }),
                                    ),
                                    LabeledField(
                                      "Direction",
                                      i.jsxs("select", {
                                        value: solverRequest.direction,
                                        onChange: (event) =>
                                          updateSolverRequest({
                                            direction: event.target.value,
                                          }),
                                        children: [
                                          i.jsx("option", {
                                            value: "maximize",
                                            children: "Maximize",
                                          }),
                                          i.jsx("option", {
                                            value: "minimize",
                                            children: "Minimize",
                                          }),
                                          i.jsx("option", {
                                            value: "target",
                                            children: "Target value",
                                          }),
                                        ],
                                      }),
                                    ),
                                    solverRequest.direction === "target" &&
                                      LabeledField(
                                        "Target value",
                                        renderNumberField(
                                          "Target value",
                                          solverRequest.target,
                                          (target) =>
                                            updateSolverRequest({
                                              target: target,
                                            }),
                                        ),
                                      ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Changing cells",
                                  }),
                                }),
                                i.jsx("p", {
                                  className: "lg-analysis-hint",
                                  children: i.jsx(LocalizedText, {
                                    children:
                                      "Up to six numeric inputs. Give every input finite lower and upper bounds. Use continuous inputs together, or integer and binary inputs together.",
                                  }),
                                }),
                                solverRequest.variables.map(
                                  (variable, variableIndex) =>
                                    i.jsxs(
                                      "div",
                                      {
                                        className:
                                          "lg-model-controls lg-analysis-row",
                                        children: [
                                          LabeledField(
                                            `Variable ${variableIndex + 1} worksheet`,
                                            X$1(
                                              variable.cell.sheetId,
                                              (sheetId) =>
                                                updateSolverRequest({
                                                  variables:
                                                    solverRequest.variables.map(
                                                      (
                                                        variable,
                                                        variableSheetIndex,
                                                      ) =>
                                                        variableSheetIndex ===
                                                        variableIndex
                                                          ? {
                                                              ...variable,
                                                              cell: {
                                                                ...variable.cell,
                                                                sheetId:
                                                                  sheetId,
                                                              },
                                                            }
                                                          : variable,
                                                    ),
                                                }),
                                              `Variable ${variableIndex + 1} worksheet`,
                                            ),
                                          ),
                                          LabeledField(
                                            `Variable ${variableIndex + 1} cell`,
                                            i.jsx("input", {
                                              value: variable.cell.address,
                                              onChange: (event) =>
                                                updateSolverRequest({
                                                  variables:
                                                    solverRequest.variables.map(
                                                      (
                                                        variable,
                                                        variableCellIndex,
                                                      ) =>
                                                        variableCellIndex ===
                                                        variableIndex
                                                          ? {
                                                              ...variable,
                                                              cell: {
                                                                ...variable.cell,
                                                                address:
                                                                  event.target.value
                                                                    .replace(
                                                                      /\$/g,
                                                                      "",
                                                                    )
                                                                    .toUpperCase(),
                                                              },
                                                            }
                                                          : variable,
                                                    ),
                                                }),
                                            }),
                                          ),
                                          ["lower", "upper"].map((boundKey) =>
                                            LabeledField(
                                              `Variable ${variableIndex + 1} ${boundKey}`,
                                              renderNumberField(
                                                `Variable ${variableIndex + 1} ${boundKey}`,
                                                variable[boundKey],
                                                (boundValue) =>
                                                  updateSolverRequest({
                                                    variables:
                                                      solverRequest.variables.map(
                                                        (
                                                          variable,
                                                          variableBoundIndex,
                                                        ) =>
                                                          variableBoundIndex ===
                                                          variableIndex
                                                            ? {
                                                                ...variable,
                                                                [boundKey]:
                                                                  boundValue,
                                                              }
                                                            : variable,
                                                      ),
                                                  }),
                                              ),
                                            ),
                                          ),
                                          LabeledField(
                                            `Variable ${variableIndex + 1} domain`,
                                            i.jsxs("select", {
                                              value: variable.domain,
                                              onChange: (event) =>
                                                updateSolverRequest({
                                                  variables:
                                                    solverRequest.variables.map(
                                                      (
                                                        variable,
                                                        variableDomainIndex,
                                                      ) =>
                                                        variableDomainIndex ===
                                                        variableIndex
                                                          ? {
                                                              ...variable,
                                                              domain:
                                                                event.target
                                                                  .value,
                                                              ...(event.target
                                                                .value ===
                                                              "binary"
                                                                ? {
                                                                    lower: 0,
                                                                    upper: 1,
                                                                  }
                                                                : {}),
                                                            }
                                                          : variable,
                                                    ),
                                                }),
                                              children: [
                                                i.jsx("option", {
                                                  value: "continuous",
                                                  children: "Continuous",
                                                }),
                                                i.jsx("option", {
                                                  value: "integer",
                                                  children: "Integer",
                                                }),
                                                i.jsx("option", {
                                                  value: "binary",
                                                  children: "Binary",
                                                }),
                                              ],
                                            }),
                                          ),
                                          i.jsx("button", {
                                            "aria-label": text(
                                              `Remove variable ${variableIndex + 1}`,
                                            ),
                                            onClick: () =>
                                              updateSolverRequest({
                                                variables:
                                                  solverRequest.variables.filter(
                                                    (
                                                      variable,
                                                      removeVariableIndex,
                                                    ) =>
                                                      removeVariableIndex !==
                                                      variableIndex,
                                                  ),
                                              }),
                                            children: "×",
                                          }),
                                        ],
                                      },
                                      variableIndex,
                                    ),
                                ),
                                i.jsx("button", {
                                  disabled: solverRequest.variables.length >= 6,
                                  onClick: () =>
                                    updateSolverRequest({
                                      variables: [
                                        ...solverRequest.variables,
                                        {
                                          cell: {
                                            sheetId:
                                              selectableSheets[0]?.id ??
                                              props.sheetId,
                                            address: "A2",
                                          },
                                          lower: 0,
                                          upper: 100,
                                          domain: "continuous",
                                        },
                                      ],
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Add changing cell",
                                  }),
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Constraints",
                                  }),
                                }),
                                solverRequest.constraints.map(
                                  (constraint, constraintIndex) =>
                                    i.jsxs(
                                      "div",
                                      {
                                        className:
                                          "lg-model-controls lg-analysis-row",
                                        children: [
                                          LabeledField(
                                            `Constraint ${constraintIndex + 1} worksheet`,
                                            X$1(
                                              constraint.left.sheetId,
                                              (sheetId) =>
                                                updateSolverRequest({
                                                  constraints:
                                                    solverRequest.constraints.map(
                                                      (
                                                        constraint,
                                                        constraintSheetIndex,
                                                      ) =>
                                                        constraintSheetIndex ===
                                                        constraintIndex
                                                          ? {
                                                              ...constraint,
                                                              left: {
                                                                ...constraint.left,
                                                                sheetId:
                                                                  sheetId,
                                                              },
                                                            }
                                                          : constraint,
                                                    ),
                                                }),
                                              `Constraint ${constraintIndex + 1} worksheet`,
                                            ),
                                          ),
                                          LabeledField(
                                            `Constraint ${constraintIndex + 1} cell`,
                                            i.jsx("input", {
                                              value: constraint.left.address,
                                              onChange: (event) =>
                                                updateSolverRequest({
                                                  constraints:
                                                    solverRequest.constraints.map(
                                                      (
                                                        constraint,
                                                        constraintCellIndex,
                                                      ) =>
                                                        constraintCellIndex ===
                                                        constraintIndex
                                                          ? {
                                                              ...constraint,
                                                              left: {
                                                                ...constraint.left,
                                                                address:
                                                                  event.target.value
                                                                    .replace(
                                                                      /\$/g,
                                                                      "",
                                                                    )
                                                                    .toUpperCase(),
                                                              },
                                                            }
                                                          : constraint,
                                                    ),
                                                }),
                                            }),
                                          ),
                                          LabeledField(
                                            `Constraint ${constraintIndex + 1} relation`,
                                            i.jsx("select", {
                                              value: constraint.operator,
                                              onChange: (event) =>
                                                updateSolverRequest({
                                                  constraints:
                                                    solverRequest.constraints.map(
                                                      (
                                                        constraint,
                                                        constraintRelationIndex,
                                                      ) =>
                                                        constraintRelationIndex ===
                                                        constraintIndex
                                                          ? {
                                                              ...constraint,
                                                              operator:
                                                                event.target
                                                                  .value,
                                                            }
                                                          : constraint,
                                                    ),
                                                }),
                                              children: ["<=", ">=", "="].map(
                                                (relation) =>
                                                  i.jsx(
                                                    "option",
                                                    { children: relation },
                                                    relation,
                                                  ),
                                              ),
                                            }),
                                          ),
                                          LabeledField(
                                            `Constraint ${constraintIndex + 1} right type`,
                                            i.jsxs("select", {
                                              value:
                                                typeof constraint.right ==
                                                "number"
                                                  ? "number"
                                                  : "cell",
                                              onChange: (event) =>
                                                updateSolverRequest({
                                                  constraints:
                                                    solverRequest.constraints.map(
                                                      (
                                                        rightTypeConstraint,
                                                        rightTypeIndex,
                                                      ) =>
                                                        rightTypeIndex ===
                                                        constraintIndex
                                                          ? {
                                                              ...rightTypeConstraint,
                                                              right:
                                                                event.target
                                                                  .value ===
                                                                "number"
                                                                  ? 0
                                                                  : {
                                                                      sheetId:
                                                                        constraint
                                                                          .left
                                                                          .sheetId,
                                                                      address:
                                                                        "A2",
                                                                    },
                                                            }
                                                          : rightTypeConstraint,
                                                    ),
                                                }),
                                              children: [
                                                i.jsx("option", {
                                                  value: "number",
                                                  children: "Number",
                                                }),
                                                i.jsx("option", {
                                                  value: "cell",
                                                  children: "Cell",
                                                }),
                                              ],
                                            }),
                                          ),
                                          typeof constraint.right == "number"
                                            ? LabeledField(
                                                `Constraint ${constraintIndex + 1} value`,
                                                renderNumberField(
                                                  `Constraint ${constraintIndex + 1} value`,
                                                  constraint.right,
                                                  (right) =>
                                                    updateSolverRequest({
                                                      constraints:
                                                        solverRequest.constraints.map(
                                                          (
                                                            constraint,
                                                            rightValueIndex,
                                                          ) =>
                                                            rightValueIndex ===
                                                            constraintIndex
                                                              ? {
                                                                  ...constraint,
                                                                  right: right,
                                                                }
                                                              : constraint,
                                                        ),
                                                    }),
                                                ),
                                              )
                                            : i.jsxs(i.Fragment, {
                                                children: [
                                                  LabeledField(
                                                    `Constraint ${constraintIndex + 1} right worksheet`,
                                                    X$1(
                                                      constraint.right.sheetId,
                                                      (sheetId) =>
                                                        updateSolverRequest({
                                                          constraints:
                                                            solverRequest.constraints.map(
                                                              (
                                                                rightSheetConstraint,
                                                                rightSheetIndex,
                                                              ) =>
                                                                rightSheetIndex ===
                                                                constraintIndex
                                                                  ? {
                                                                      ...rightSheetConstraint,
                                                                      right: {
                                                                        ...constraint.right,
                                                                        sheetId:
                                                                          sheetId,
                                                                      },
                                                                    }
                                                                  : rightSheetConstraint,
                                                            ),
                                                        }),
                                                      `Constraint ${constraintIndex + 1} right worksheet`,
                                                    ),
                                                  ),
                                                  LabeledField(
                                                    `Constraint ${constraintIndex + 1} right cell`,
                                                    i.jsx("input", {
                                                      value:
                                                        constraint.right
                                                          .address,
                                                      onChange: (event) =>
                                                        updateSolverRequest({
                                                          constraints:
                                                            solverRequest.constraints.map(
                                                              (
                                                                rightCellConstraint,
                                                                rightCellIndex,
                                                              ) =>
                                                                rightCellIndex ===
                                                                constraintIndex
                                                                  ? {
                                                                      ...rightCellConstraint,
                                                                      right: {
                                                                        ...constraint.right,
                                                                        address:
                                                                          event.target.value
                                                                            .replace(
                                                                              /\$/g,
                                                                              "",
                                                                            )
                                                                            .toUpperCase(),
                                                                      },
                                                                    }
                                                                  : rightCellConstraint,
                                                            ),
                                                        }),
                                                    }),
                                                  ),
                                                ],
                                              }),
                                          i.jsx("button", {
                                            "aria-label": text(
                                              `Remove constraint ${constraintIndex + 1}`,
                                            ),
                                            onClick: () =>
                                              updateSolverRequest({
                                                constraints:
                                                  solverRequest.constraints.filter(
                                                    (
                                                      constraint,
                                                      removeConstraintIndex,
                                                    ) =>
                                                      removeConstraintIndex !==
                                                      constraintIndex,
                                                  ),
                                              }),
                                            children: "×",
                                          }),
                                        ],
                                      },
                                      constraintIndex,
                                    ),
                                ),
                                i.jsx("button", {
                                  disabled:
                                    solverRequest.constraints.length >= 16,
                                  onClick: () =>
                                    updateSolverRequest({
                                      constraints: [
                                        ...solverRequest.constraints,
                                        {
                                          left: {
                                            sheetId:
                                              solverRequest.objective.sheetId,
                                            address:
                                              solverRequest.objective.address,
                                          },
                                          operator: "<=",
                                          right: 100,
                                        },
                                      ],
                                    }),
                                  children: i.jsx(LocalizedText, {
                                    children: "Add constraint",
                                  }),
                                }),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "lg-model-controls",
                              children: [
                                LabeledField(
                                  "Feasibility tolerance",
                                  renderNumberField(
                                    "Feasibility tolerance",
                                    solverRequest.tolerance,
                                    (tolerance) =>
                                      updateSolverRequest({
                                        tolerance: tolerance,
                                      }),
                                  ),
                                ),
                                LabeledField(
                                  "Maximum candidates",
                                  renderNumberField(
                                    "Maximum candidates",
                                    solverRequest.maxCandidates,
                                    (maxCandidates) =>
                                      updateSolverRequest({
                                        maxCandidates: maxCandidates,
                                      }),
                                  ),
                                ),
                              ],
                            }),
                            i.jsx("p", {
                              className: "lg-analysis-hint",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "Supported formulas: numeric cell references, defined names, +, −, ×, ÷, SUM, AVERAGE and linear SUMPRODUCT. Nonlinear and mixed integer models require another engine.",
                              }),
                            }),
                          ],
                        })
                      : i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("fieldset", {
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Source and method",
                                  }),
                                }),
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    LabeledField(
                                      "Source worksheet",
                                      X$1(
                                        statisticsRequest.sheetId,
                                        (sheetId) =>
                                          updateStatisticsRequest({
                                            sheetId: sheetId,
                                          }),
                                        "Source worksheet",
                                      ),
                                    ),
                                    LabeledField(
                                      "Headed source range",
                                      i.jsx("input", {
                                        value: statisticsRequest.range,
                                        onChange: (event) =>
                                          updateStatisticsRequest({
                                            range: event.target.value
                                              .replace(/\$/g, "")
                                              .toUpperCase(),
                                          }),
                                      }),
                                    ),
                                    LabeledField(
                                      "Analysis method",
                                      i.jsx("select", {
                                        value: statisticsRequest.method,
                                        onChange: (event) =>
                                          updateStatisticsRequest({
                                            method: event.target.value,
                                          }),
                                        children: [
                                          [
                                            "descriptive",
                                            "Descriptive statistics",
                                          ],
                                          ["correlation", "Correlation"],
                                          ["covariance", "Covariance"],
                                          ["regression", "Regression"],
                                          ["histogram", "Histogram"],
                                          [
                                            "distribution",
                                            "Probability distributions",
                                          ],
                                          ["forecast", "Forecasting"],
                                        ].map(([methodValue, methodLabel]) =>
                                          i.jsx(
                                            "option",
                                            {
                                              value: methodValue,
                                              children: methodLabel,
                                            },
                                            methodValue,
                                          ),
                                        ),
                                      }),
                                    ),
                                    LabeledField(
                                      "Incomplete rows",
                                      i.jsxs("select", {
                                        value: statisticsRequest.missing,
                                        onChange: (event) =>
                                          updateStatisticsRequest({
                                            missing: event.target.value,
                                          }),
                                        children: [
                                          i.jsx("option", {
                                            value: "reject",
                                            children: "Reject analysis",
                                          }),
                                          i.jsx("option", {
                                            value: "exclude",
                                            children:
                                              "Exclude and list affected rows",
                                          }),
                                        ],
                                      }),
                                    ),
                                  ],
                                }),
                                i.jsx("p", {
                                  className: "lg-analysis-hint",
                                  children: i.jsx(LocalizedText, {
                                    children:
                                      "Include text column headers and numeric data. Up to 500 data rows, six columns and 2,000 cells. Formula inputs are limited to 50 deterministic scalar formulas with bounded dependencies.",
                                  }),
                                }),
                              ],
                            }),
                            statisticsRequest.method === "regression" &&
                              i.jsx("p", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Put the response Y in the first column and predictors X in the remaining columns. The model includes an intercept, coefficient confidence intervals and residuals.",
                                }),
                              }),
                            statisticsRequest.method === "histogram" &&
                              LabeledField(
                                "Bin upper bounds",
                                i.jsx("input", {
                                  value: binsText,
                                  onChange: (event) => {
                                    (setBinsText(event.target.value),
                                      invalidatePreview());
                                  },
                                }),
                              ),
                            statisticsRequest.method === "distribution" &&
                              i.jsxs("fieldset", {
                                children: [
                                  i.jsx("legend", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Distribution parameters",
                                    }),
                                  }),
                                  i.jsxs("div", {
                                    className: "lg-model-controls",
                                    children: [
                                      LabeledField(
                                        "Distribution",
                                        i.jsxs("select", {
                                          value:
                                            statisticsRequest.distribution ??
                                            "normal",
                                          onChange: (event) =>
                                            updateStatisticsRequest({
                                              distribution: event.target.value,
                                            }),
                                          children: [
                                            i.jsx("option", {
                                              value: "normal",
                                              children: "Normal",
                                            }),
                                            i.jsx("option", {
                                              value: "binomial",
                                              children: "Binomial",
                                            }),
                                            i.jsx("option", {
                                              value: "poisson",
                                              children: "Poisson",
                                            }),
                                          ],
                                        }),
                                      ),
                                      !statisticsRequest.distribution ||
                                      statisticsRequest.distribution ===
                                        "normal"
                                        ? i.jsxs(i.Fragment, {
                                            children: [
                                              LabeledField(
                                                "Mean",
                                                renderNumberField(
                                                  "Mean",
                                                  statisticsRequest.mean,
                                                  (mean) =>
                                                    updateStatisticsRequest({
                                                      mean: mean,
                                                    }),
                                                ),
                                              ),
                                              LabeledField(
                                                "Standard deviation",
                                                renderNumberField(
                                                  "Standard deviation",
                                                  statisticsRequest.standardDeviation,
                                                  (standardDeviation) =>
                                                    updateStatisticsRequest({
                                                      standardDeviation:
                                                        standardDeviation,
                                                    }),
                                                ),
                                              ),
                                            ],
                                          })
                                        : statisticsRequest.distribution ===
                                            "binomial"
                                          ? i.jsxs(i.Fragment, {
                                              children: [
                                                LabeledField(
                                                  "Trials",
                                                  renderNumberField(
                                                    "Trials",
                                                    statisticsRequest.trials,
                                                    (trials) =>
                                                      updateStatisticsRequest({
                                                        trials: trials,
                                                      }),
                                                  ),
                                                ),
                                                LabeledField(
                                                  "Success probability",
                                                  renderNumberField(
                                                    "Success probability",
                                                    statisticsRequest.probability,
                                                    (probability) =>
                                                      updateStatisticsRequest({
                                                        probability:
                                                          probability,
                                                      }),
                                                  ),
                                                ),
                                              ],
                                            })
                                          : LabeledField(
                                              "Rate (lambda)",
                                              renderNumberField(
                                                "Rate (lambda)",
                                                statisticsRequest.lambda,
                                                (lambda) =>
                                                  updateStatisticsRequest({
                                                    lambda: lambda,
                                                  }),
                                              ),
                                            ),
                                    ],
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Select one column of evaluation points. Parameters are supplied explicitly; they are not estimated from that column.",
                                    }),
                                  }),
                                ],
                              }),
                            statisticsRequest.method === "forecast" &&
                              i.jsxs("fieldset", {
                                children: [
                                  i.jsx("legend", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Forecast settings",
                                    }),
                                  }),
                                  i.jsxs("div", {
                                    className: "lg-model-controls",
                                    children: [
                                      LabeledField(
                                        "Forecast method",
                                        i.jsxs("select", {
                                          value:
                                            statisticsRequest.forecast ??
                                            "linear",
                                          onChange: (event) =>
                                            updateStatisticsRequest({
                                              forecast: event.target.value,
                                            }),
                                          children: [
                                            i.jsx("option", {
                                              value: "linear",
                                              children: "Linear trend",
                                            }),
                                            i.jsx("option", {
                                              value: "moving-average",
                                              children: "Moving average",
                                            }),
                                            i.jsx("option", {
                                              value: "exponential",
                                              children: "Exponential smoothing",
                                            }),
                                          ],
                                        }),
                                      ),
                                      LabeledField(
                                        "Forecast periods",
                                        renderNumberField(
                                          "Forecast periods",
                                          statisticsRequest.horizon ?? 6,
                                          (horizon) =>
                                            updateStatisticsRequest({
                                              horizon: horizon,
                                            }),
                                        ),
                                      ),
                                      statisticsRequest.forecast ===
                                        "moving-average" &&
                                        LabeledField(
                                          "Moving average window",
                                          renderNumberField(
                                            "Moving average window",
                                            statisticsRequest.window ?? 3,
                                            (windowSize) =>
                                              updateStatisticsRequest({
                                                window: windowSize,
                                              }),
                                          ),
                                        ),
                                      statisticsRequest.forecast ===
                                        "exponential" &&
                                        LabeledField(
                                          "Smoothing alpha",
                                          renderNumberField(
                                            "Smoothing alpha",
                                            statisticsRequest.alpha ?? 0.3,
                                            (alpha) =>
                                              updateStatisticsRequest({
                                                alpha: alpha,
                                              }),
                                          ),
                                        ),
                                    ],
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Select numeric period then value. Periods must be increasing and equally spaced. No seasonal model or prediction interval is claimed.",
                                    }),
                                  }),
                                ],
                              }),
                          ],
                        }),
                  ],
                })
              : i.jsx(i.Fragment, {
                  children: preview
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("div", {
                            className: `lg-analysis-summary ${preview.result.success ? "" : "failed"}`,
                            role: "status",
                            children: [
                              i.jsx("strong", {
                                children: preview.result.status.toUpperCase(),
                              }),
                              i.jsx("span", {
                                children: preview.result.message,
                              }),
                              i.jsx("span", {
                                children: preview.result.algorithm,
                              }),
                              i.jsxs("span", {
                                children: [
                                  "Included ",
                                  preview.result.included,
                                  " / ",
                                  preview.result.inspected,
                                  " ",
                                  "· Excluded ",
                                  preview.result.exclusions.length,
                                  " · Candidates",
                                  " ",
                                  preview.result.iterations,
                                ],
                              }),
                              preview.result.objective !== undefined &&
                                i.jsxs("b", {
                                  children: [
                                    "Objective:",
                                    " ",
                                    preview.result.objective.toLocaleString(
                                      undefined,
                                      {
                                        maximumFractionDigits: 10,
                                      },
                                    ),
                                  ],
                                }),
                            ],
                          }),
                          !isPreviewCurrent &&
                            i.jsx("p", {
                              role: "alert",
                              children: i.jsx(LocalizedText, {
                                children:
                                  "The workbook or permissions changed. Preview again.",
                              }),
                            }),
                          preview.result.rows.length > 0 &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsx("div", {
                                  className: "lg-analysis-results",
                                  tabIndex: 0,
                                  "aria-label": text("Analysis results"),
                                  children: i.jsx("table", {
                                    children: i.jsx("tbody", {
                                      children: preview.result.rows
                                        .slice(
                                          resultPage * 20,
                                          resultPage * 20 + 20,
                                        )
                                        .map((resultRow, rowIndex) =>
                                          i.jsx(
                                            "tr",
                                            {
                                              children: resultRow.map(
                                                (cellValue, columnIndex) =>
                                                  i.jsx(
                                                    "td",
                                                    {
                                                      children:
                                                        cellValue === null
                                                          ? "Undefined"
                                                          : typeof cellValue ==
                                                              "number"
                                                            ? cellValue !== 0 &&
                                                              (Math.abs(
                                                                cellValue,
                                                              ) < 1e-6 ||
                                                                Math.abs(
                                                                  cellValue,
                                                                ) >= 1e12)
                                                              ? cellValue.toExponential(
                                                                  6,
                                                                )
                                                              : cellValue.toLocaleString(
                                                                  undefined,
                                                                  {
                                                                    maximumSignificantDigits: 12,
                                                                  },
                                                                )
                                                            : String(cellValue),
                                                    },
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
                                i.jsxs("div", {
                                  className: "lg-model-controls",
                                  children: [
                                    i.jsx("button", {
                                      disabled: resultPage === 0,
                                      onClick: () =>
                                        setResultPage(
                                          (currentPage) => currentPage - 1,
                                        ),
                                      children: i.jsx(LocalizedText, {
                                        children: "Previous results",
                                      }),
                                    }),
                                    i.jsxs("span", {
                                      children: [
                                        "Page ",
                                        resultPage + 1,
                                        " /",
                                        " ",
                                        Math.max(
                                          1,
                                          Math.ceil(
                                            preview.result.rows.length / 20,
                                          ),
                                        ),
                                      ],
                                    }),
                                    i.jsx("button", {
                                      disabled:
                                        (resultPage + 1) * 20 >=
                                        preview.result.rows.length,
                                      onClick: () =>
                                        setResultPage(
                                          (pageCursor) => pageCursor + 1,
                                        ),
                                      children: i.jsx(LocalizedText, {
                                        children: "Next results",
                                      }),
                                    }),
                                  ],
                                }),
                              ],
                            }),
                          i.jsxs("details", {
                            open: true,
                            children: [
                              i.jsx("summary", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Method, assumptions and reference checks",
                                }),
                              }),
                              i.jsx("ul", {
                                children: preview.result.assumptions.map(
                                  (assumption, assumptionIndex) =>
                                    i.jsx(
                                      "li",
                                      { children: assumption },
                                      assumptionIndex,
                                    ),
                                ),
                              }),
                            ],
                          }),
                          i.jsxs("details", {
                            children: [
                              i.jsxs("summary", {
                                children: [
                                  "Sample exclusions (",
                                  preview.result.exclusions.length,
                                  ")",
                                ],
                              }),
                              preview.result.exclusions.map((exclusion) =>
                                i.jsxs(
                                  "p",
                                  {
                                    children: [
                                      exclusion.locator,
                                      " — ",
                                      exclusion.reason,
                                    ],
                                  },
                                  exclusion.row,
                                ),
                              ),
                            ],
                          }),
                        ],
                      })
                    : i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Choose the inputs in Setup, then preview the analysis.",
                        }),
                      }),
                }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              className: "lg-analysis-license",
              children: i.jsx(LocalizedText, {
                children:
                  "Local computation · MIT engines · Preview leaves source cells unchanged",
              }),
            }),
            i.jsx("button", {
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Preview analysis" }),
            }),
            i.jsx("button", {
              disabled:
                props.readOnly || !isPreviewCurrent || !preview?.result.success,
              onClick: () => saveAnalysis("report"),
              children: i.jsx(LocalizedText, { children: "Save report" }),
            }),
            analysisKind === "solver" &&
              i.jsx("button", {
                className: "primary",
                disabled:
                  props.readOnly ||
                  !isPreviewCurrent ||
                  !preview?.result.success,
                onClick: () => saveAnalysis("apply-solution"),
                children: i.jsx(LocalizedText, {
                  children: "Apply solution and report",
                }),
              }),
          ],
        }),
      ],
    }),
  });
}
function AnalysisSnapshotNotice({ workbook, sheetId, onReview }) {
  const analysisReport = workbook.sheets.find(
    (sheet) => sheet.id === sheetId,
  )?.analysisReport;
  if (!analysisReport) return null;
  const savedAnalysisReport = readSavedAnalysisReport(workbook, sheetId, {
    includeHidden: analysisReport.includeHidden,
  });
  return i.jsxs("div", {
    className: "lg-notice",
    role: "status",
    children: [
      i.jsxs("strong", {
        children: ["Analysis snapshot · ", savedAnalysisReport.status],
      }),
      i.jsx("span", { children: savedAnalysisReport.message }),
      i.jsx("button", {
        onClick: onReview,
        children: i.jsx(LocalizedText, { children: "Review analysis" }),
      }),
    ],
  });
}
function captureAnalysisWorkbookState(workbookInput) {
  return fingerprintCanonicalValue({
    id: workbookInput.id,
    name: workbookInput.name,
    sheets: workbookInput.sheets,
    datePolicy: workbookInput.datePolicy ?? null,
    calculation: workbookInput.calculation ?? null,
    definedNames: workbookInput.definedNames ?? null,
    styleCatalog: workbookInput.styleCatalog ?? null,
  });
}
const isCellHiddenInSheet = (sheet, rowIndex, columnIndex) =>
  !!(
    sheet.hidden ||
    sheet.veryHidden ||
    cn(sheet, "row").some(
      (hiddenRowRange) =>
        hiddenRowRange.start <= rowIndex && hiddenRowRange.end >= rowIndex,
    ) ||
    cn(sheet, "column").some(
      (hiddenColumnRange) =>
        hiddenColumnRange.start <= columnIndex &&
        hiddenColumnRange.end >= columnIndex,
    )
  );
function buildReportSourceSheet(reportSnapshot) {
  return {
    id: reportSnapshot.request.sheetId,
    name: reportSnapshot.sheetName,
    rowCount: 1048576,
    colCount: 16384,
    cells: Object.fromEntries(
      reportSnapshot.rows.flatMap((row) =>
        row.cells.map((cell) => [
          cell.address,
          { raw: cell.value, literal: true, style: cell.style },
        ]),
      ),
    ),
  };
}
function buildReportSnapshot(workbook, requestInput, hostContext = {}) {
  const request = validateObjectReportRequest(requestInput),
    analysisPermissions = normalizeAnalysisPermissions(workbook, hostContext),
    workbookInput = normalizeWorkbookInput(workbook),
    sheet = workbookInput.sheets.find((sheet) => sheet.id === request.sheetId);
  if (!sheet || !analysisPermissions.allowedSheetIds.includes(sheet.id))
    throw Error("The report source is outside the authorized host context.");
  if (!analysisPermissions.includeHidden && (sheet.hidden || sheet.veryHidden))
    throw Error("The report source worksheet is hidden.");
  if (
    workbookInput.calculation?.mode === "manual" ||
    workbookInput.calculation?.iterative
  )
    throw Error(
      "Report capture requires automatic, non-iterative calculation. Recalculate and retry.",
    );
  const object =
    request.kind === "object"
      ? structuredClone(
          sheet.objects?.find((object) => object.id === request.objectId),
        )
      : undefined;
  if (request.kind === "object" && !object)
    throw Error("The source drawing no longer exists.");
  if (
    object &&
    parseReportRange(object.kind === "chart" ? object.range : object.anchor) !==
      request.range
  )
    throw Error("The report range must match the drawing source.");
  if (object && object.kind !== "chart" && request.filter)
    throw Error("Pictures and shapes do not use population filters.");
  const rangeAddress = parseRangeAddress(request.range),
    columns = Array.from(
      { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
      (j, columnOffset) => rangeAddress.firstCol + columnOffset,
    );
  let gridValues,
    dependencyCells = [];
  if (object && object.kind !== "chart") {
    if (
      !analysisPermissions.includeHidden &&
      isCellHiddenInSheet(sheet, rangeAddress.firstRow, rangeAddress.firstCol)
    )
      throw Error(
        "The drawing anchor is hidden. Explicit authorization is required.",
      );
    ((gridValues = [[null]]),
      (dependencyCells = [{ sheetId: sheet.id, address: object.anchor }]));
  } else {
    const sourceRectangles = readSourceRectangles(
      workbookInput,
      [{ sheetId: sheet.id, range: request.range }],
      hostContext,
    );
    ((gridValues = sourceRectangles.grids[0].values),
      (dependencyCells = sourceRectangles.dependencyCells));
  }
  const valuesByAddress = new Map();
  gridValues.forEach((rowValues, rowOffset) =>
    rowValues.forEach((cellValue, columnOffset) =>
      valuesByAddress.set(
        `${columnIndexToLetters(rangeAddress.firstCol + columnOffset)}${rangeAddress.firstRow + rowOffset + 1}`,
        cellValue,
      ),
    ),
  );
  const getCapturedValue = (sheetId, address) => {
      if (sheetId !== sheet.id || !valuesByAddress.has(address))
        throw Error("Report calculation is outside the captured range.");
      return valuesByAddress.get(address);
    },
    filterResult = request.filter
      ? applyFilterConditions(
          workbookInput,
          sheet.id,
          request.filter,
          getCapturedValue,
        )
      : undefined;
  filterResult && (request.filter = filterResult.context.definition);
  const rowNumbers = gridValues.map(
      (j, rowIndex) => rangeAddress.firstRow + rowIndex + 1,
    ),
    populationRows =
      request.population === "visible"
        ? rowNumbers.filter(
            (rowNumber) =>
              !cn(sheet, "row").some(
                (rowRange) =>
                  rowRange.start <= rowNumber - 1 &&
                  rowRange.end >= rowNumber - 1,
              ) &&
              (!filterResult ||
                rowNumber === rangeAddress.firstRow + 1 ||
                filterResult.context.matchingSourceRows.includes(rowNumber)),
          )
        : rowNumbers;
  let retainedRows = populationRows;
  if (object?.kind === "chart") {
    const includedRows = object.chartOptions?.includedRows,
      selectedRows = populationRows.filter(
        (rowNumber) =>
          (object.headers !== false &&
            rowNumber === rangeAddress.firstRow + 1) ||
          !includedRows ||
          includedRows.includes(rowNumber),
      );
    object.chartOptions = {
      ...object.chartOptions,
      includedRows: selectedRows.filter(
        (rowNumber) =>
          !(
            object.headers !== false && rowNumber === rangeAddress.firstRow + 1
          ),
      ),
    };
    const chartDiagnostics = inspectChartSource(
      sheet,
      object,
      getCapturedValue,
    );
    (formatDiagnosticsMessage(chartDiagnostics), (retainedRows = selectedRows));
  }
  if (!retainedRows.length)
    throw Error("No rows remain in the report population.");
  const warnings = [
    "Fixed values and formatting captured locally; links refresh only after host authorization and explicit review. No network access or remote refresh is performed.",
    "Table images use readable, wrapped cell text with portable fonts. Merges are refused; conditional formatting, layout fidelity and native Office OLE links are not represented.",
  ];
  ((sheet.analysisReport ||
    sheet.dataModelReport ||
    sheet.pivots?.length ||
    sheet.queries?.length) &&
    warnings.push(
      "This worksheet contains generated report output. Captured cells are historical source values; verify the underlying analysis, model, pivot or query freshness separately.",
    ),
    request.filter &&
      warnings.push(
        request.population === "full"
          ? "Filter criteria are recorded, but full population was selected: filtered rows are included."
          : "Only matching visible rows are retained; excluded rows have locators but no values in this snapshot.",
      ),
    object?.kind === "chart" &&
      warnings.push(
        "Chart row selections are preserved. Explicitly excluded chart rows are omitted from retained cells.",
      ));
  const snapshot = {
    engine: reportOutputFormatVersion,
    id: crypto.randomUUID(),
    request: request,
    workbookId: workbookInput.id,
    workbookName: workbookInput.name,
    revision: workbookInput.revision,
    sheetName: sheet.name,
    calculatedAt: new Date().toISOString(),
    sourceHash: captureAnalysisWorkbookState(workbookInput),
    includeHidden: analysisPermissions.includeHidden,
    authorizedSheetIds: [
      ...new Set([
        sheet.id,
        ...dependencyCells.map((dependencyCell) => dependencyCell.sheetId),
      ]),
    ].sort(),
    sourceCells: dependencyCells.map((dependencyCell) => ({
      sheetId: dependencyCell.sheetId,
      address: dependencyCell.address,
    })),
    rows: retainedRows.map((row) => ({
      row: row,
      cells: columns.map((column, columnIndex) => {
        const address = `${columnIndexToLetters(column)}${row}`,
          cellValue = gridValues[row - rangeAddress.firstRow - 1][columnIndex],
          cell = sheet.cells[address];
        return {
          address: address,
          value: cellValue,
          text: formatNumberWithPattern(
            cellValue,
            cell,
            workbookInput.datePolicy?.dateSystem,
          ),
          ...(cell?.style
            ? { style: resolveEffectiveCellStyle(cell.style) }
            : {}),
        };
      }),
    })),
    columns: columns,
    widths: columns.map((column) => sheet.columnWidths?.[String(column)] ?? 84),
    excludedRows: rowNumbers.filter(
      (rowNumber) => !retainedRows.includes(rowNumber),
    ),
    ...(object ? { object: object } : {}),
    warnings: warnings,
    hash: "",
  };
  return (
    (snapshot.hash = hashRecordExcludingHash(snapshot)),
    deepFreezeValue(parseReportSnapshot(snapshot))
  );
}
const qT = new WeakMap();
function refreshReportOutput(workbook, request, hostContext = {}, refreshId) {
  if (refreshId) {
    const existingOutput = workbook.reportOutputs?.find(
      (reportOutput) => reportOutput.id === refreshId,
    );
    if (!existingOutput || existingOutput.snapshot.request.mode !== "linked")
      throw Error("Only an existing linked output can be refreshed.");
    if (
      fingerprintCanonicalValue(validateObjectReportRequest(request)) !==
      fingerprintCanonicalValue(existingOutput.snapshot.request)
    )
      throw Error(
        "Refresh uses the saved definition. Publish a new output to change it.",
      );
    if (existingOutput.snapshot.workbookId !== workbook.id)
      throw Error(
        "This link belongs to a different workbook. Publish a new output after reviewing the source.",
      );
  }
  const reportSnapshot = buildReportSnapshot(workbook, request, hostContext),
    preview = {
      request: reportSnapshot.request,
      snapshot: reportSnapshot,
      workbookHash: fingerprintCanonicalValue(workbook),
      contextHash: fingerprintCanonicalValue(
        normalizeAnalysisPermissions(workbook, hostContext),
      ),
      ...(refreshId ? { refreshId: refreshId } : {}),
    };
  return (qT.set(preview, fingerprintCanonicalValue(preview)), preview);
}
function publishReportOutput(workbook, preview, hostContext = {}) {
  if (hostContext.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    qT.get(preview) !== fingerprintCanonicalValue(preview) ||
      preview.workbookHash !== fingerprintCanonicalValue(workbook) ||
      preview.contextHash !==
        fingerprintCanonicalValue(
          normalizeAnalysisPermissions(workbook, hostContext),
        ))
  )
    throw Error(
      "The preview, workbook or permissions changed. Preview the report again.",
    );
  const workbookInput = normalizeWorkbookInput(workbook),
    reportOutputs = [...(workbookInput.reportOutputs ?? [])];
  if (preview.refreshId) {
    const outputIndex = reportOutputs.findIndex(
      (reportOutput) => reportOutput.id === preview.refreshId,
    );
    if (outputIndex < 0) throw Error("The linked output was removed.");
    reportOutputs[outputIndex] = {
      ...reportOutputs[outputIndex],
      version: reportOutputs[outputIndex].version + 1,
      snapshot: structuredClone(preview.snapshot),
    };
  } else {
    if (
      reportOutputs.some(
        (reportOutput) =>
          reportOutput.snapshot.request.name.toLowerCase() ===
          preview.request.name.toLowerCase(),
      )
    )
      throw Error("A report with this name already exists.");
    reportOutputs.push({
      id: crypto.randomUUID(),
      version: 1,
      snapshot: structuredClone(preview.snapshot),
    });
  }
  const nextWorkbook = normalizeWorkbookInput({
    ...workbookInput,
    reportOutputs: reportOutputs,
  });
  return (qT.delete(preview), nextWorkbook);
}
function removeReportOutput(workbook, outputId, hostContext = {}) {
  if (hostContext.readOnly) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(workbook);
  const existingOutput = workbook.reportOutputs?.find(
    (reportOutput) => reportOutput.id === outputId,
  );
  if (!existingOutput) throw Error("Report output not found.");
  return (
    validateReportSourceScope(workbook, existingOutput.snapshot, hostContext),
    normalizeWorkbookInput({
      ...workbook,
      reportOutputs: workbook.reportOutputs.filter(
        (reportOutput) => reportOutput.id !== outputId,
      ),
    })
  );
}
function validateReportSourceScope(workbook, snapshot, hostContext = {}) {
  const analysisPermissions = normalizeAnalysisPermissions(
    workbook,
    hostContext,
  );
  if (
    snapshot.workbookId !== workbook.id ||
    snapshot.authorizedSheetIds.some(
      (authorizedSheetId) =>
        !analysisPermissions.allowedSheetIds.includes(authorizedSheetId),
    ) ||
    (snapshot.includeHidden && !analysisPermissions.includeHidden)
  )
    throw Error("The report is outside the current authorized source context.");
  for (const sourceCell of snapshot.sourceCells) {
    const sheet = workbook.sheets.find(
        (sheet) => sheet.id === sourceCell.sheetId,
      ),
      cellPosition = _t(sourceCell.address);
    if (
      !sheet ||
      !analysisPermissions.allowedSheetIds.includes(sourceCell.sheetId) ||
      (!analysisPermissions.includeHidden &&
        isCellHiddenInSheet(sheet, cellPosition.row, cellPosition.col))
    )
      throw Error(
        "A report source cell is missing, hidden or outside the authorized context.",
      );
  }
  for (const authorizedSheetId of snapshot.authorizedSheetIds) {
    const sheet = workbook.sheets.find(
      (sheet) => sheet.id === authorizedSheetId,
    );
    if (
      !sheet ||
      (!analysisPermissions.includeHidden && (sheet.hidden || sheet.veryHidden))
    )
      throw Error("The report source is missing or hidden.");
  }
}
function describeReportOutputStatus(workbook, outputId, hostContext = {}) {
  try {
    const existingOutput = workbook.reportOutputs?.find(
      (reportOutput) => reportOutput.id === outputId,
    );
    if (!existingOutput) throw Error("Report output not found.");
    validateReportSourceScope(workbook, existingOutput.snapshot, hostContext);
    const reportSnapshot = parseReportSnapshot(existingOutput.snapshot);
    return reportSnapshot.request.mode === "snapshot"
      ? {
          status: "snapshot",
          message: `Fixed snapshot of revision ${reportSnapshot.revision}. Later edits do not change it.`,
        }
      : captureAnalysisWorkbookState(workbook) !== reportSnapshot.sourceHash
        ? {
            status: "stale",
            message:
              "Source content or settings changed. Preview and review a refresh; the saved snapshot remains unchanged.",
          }
        : {
            status: "current",
            message:
              "Source fingerprint matches. The rendered content remains the saved snapshot; use a trusted read for current values.",
          };
  } catch (error) {
    return {
      status: "unavailable",
      message: error instanceof Error ? error.message : String(error),
    };
  }
}
const escapeHtmlText = (text) =>
    String(text).replace(
      /[&<>"']/g,
      (character) =>
        ({
          "&": "&amp;",
          "<": "&lt;",
          ">": "&gt;",
          '"': "&quot;",
          "'": "&apos;",
        })[character],
    ),
  toSafeFileName = (rawName) =>
    rawName.replace(/[\x00-\x1f\\/:*?"<>|]/g, "_").slice(0, 100) || "Report",
  formatReportCaption = (snapshot) =>
    `${snapshot.workbookName} · revision ${snapshot.revision} · ${snapshot.sheetName}!${snapshot.request.range} · calculated ${snapshot.calculatedAt}`;
function wrapTextByCharacters(text, charsPerLine) {
  return text
    .replace(
      /\r\n?/g,
      `
`,
    )
    .split(
      `
`,
    )
    .flatMap((line) => {
      const characters = Array.from(line),
        wrappedLines = [];
      for (
        let startIndex = 0;
        startIndex < characters.length;
        startIndex += charsPerLine
      )
        wrappedLines.push(
          characters.slice(startIndex, startIndex + charsPerLine).join(""),
        );
      return wrappedLines.length ? wrappedLines : [""];
    });
}
function renderReportSvg(rawSnapshot) {
  const reportSnapshot = parseReportSnapshot(rawSnapshot),
    object = reportSnapshot.object;
  if (
    [
      reportSnapshot.request.name,
      reportSnapshot.workbookName,
      reportSnapshot.sheetName,
      object?.label ?? "",
      ...reportSnapshot.rows.flatMap((row) =>
        row.cells.map((cell) => cell.text),
      ),
    ].some((candidateText) =>
      new RegExp(
        "[\\u0000-\\u0008\\u000B\\u000C\\u000E-\\u001F\\uFFFE\\uFFFF]|\\p{Surrogate}",
        "u",
      ).test(candidateText),
    )
  )
    throw Error(
      "Source text contains characters SVG cannot represent. Clean the text or export JSON/HTML.",
    );
  let width = 640,
    height = 0,
    svgBody = "";
  if (object?.kind === "chart") {
    const reportSourceSheet = buildReportSourceSheet(reportSnapshot),
      report = inspectChartSource(
        reportSourceSheet,
        object,
        (m, address) => reportSourceSheet.cells[address]?.raw ?? null,
      );
    formatDiagnosticsMessage(report);
    const chartMarkup = c_.renderToStaticMarkup(
      i.jsx(ChartReportSvg, { report: report }),
    );
    ((width = 640),
      (height = object.chart === "sparkline" ? 128 : 380),
      (svgBody = `<svg x="18" y="112" width="${width}" height="${height}" viewBox="${object.chart === "sparkline" ? "0 0 200 40" : "0 0 640 380"}">${chartMarkup}</svg>`));
  } else if (object?.kind === "picture")
    (inspectPictureDataUrl(object.dataURL),
      (width = Math.max(640, object.width)),
      (height = object.height),
      (svgBody = `<image x="18" y="112" width="${object.width}" height="${height}" href="${escapeHtmlText(object.dataURL)}"/>`));
  else if (object?.kind === "shape") {
    ((width = Math.max(640, object.width)),
      (height = Math.max(80, object.height)));
    const shapeAttributes = `fill="${object.fill}" stroke="${object.stroke}" stroke-width="2"`,
      shapeMarkup =
        object.shape === "ellipse"
          ? `<ellipse cx="${object.width / 2}" cy="${height / 2}" rx="${object.width / 2 - 2}" ry="${height / 2 - 2}" ${shapeAttributes}/>`
          : object.shape === "arrow"
            ? `<path d="M2 20 H${object.width - 30} V2 L${object.width - 2} ${height / 2} L${object.width - 30} ${height - 2} V${height - 20} H2 Z" ${shapeAttributes}/>`
            : object.shape === "text"
              ? ""
              : `<rect x="1" y="1" width="${object.width - 2}" height="${height - 2}" ${shapeAttributes}/>`,
      labelLines = wrapTextByCharacters(
        object.label,
        Math.max(4, Math.floor((object.width - 16) / 8)),
      );
    if (labelLines.length * 18 + 16 > height)
      throw Error(
        "The shape label does not fit the report image. Enlarge the source shape before publishing.",
      );
    svgBody = `<g transform="translate(18 112)">${shapeMarkup}${labelLines.map((line, lineIndex) => `<text x="8" y="${24 + lineIndex * 18}" font-size="13" fill="#202334">${escapeHtmlText(line)}</text>`).join("")}</g>`;
  } else {
    const columnWidths = reportSnapshot.columns.map((column, columnIndex) =>
      Math.min(
        600,
        Math.max(
          84,
          reportSnapshot.widths[columnIndex],
          ...reportSnapshot.rows.map((row) =>
            typeof row.cells[columnIndex].value == "number"
              ? Array.from(row.cells[columnIndex].text).length * 8 + 20
              : 84,
          ),
        ),
      ),
    );
    width = Math.max(
      640,
      columnWidths.reduce((total, columnWidth) => total + columnWidth, 0),
    );
    let rowTop = 112;
    for (const row of reportSnapshot.rows) {
      const cellLines = row.cells.map((cell, cellIndex) =>
          wrapTextByCharacters(
            cell.text,
            Math.max(1, Math.floor((columnWidths[cellIndex] - 16) / 8)),
          ),
        ),
        rowHeight = Math.max(
          28,
          ...cellLines.map((lines) => lines.length * 18 + 12),
        );
      let cellLeft = 18;
      (row.cells.forEach((cell, cellIndex) => {
        const cellStyle = cell.style ?? {},
          cellWidth = columnWidths[cellIndex];
        ((svgBody += `<rect x="${cellLeft}" y="${rowTop}" width="${cellWidth}" height="${rowHeight}" fill="${escapeHtmlText(cellStyle.fill ?? "#ffffff")}" stroke="#ccd6fa"/><text font-family="Courier New,monospace" font-size="13" fill="${escapeHtmlText(cellStyle.color ?? "#202334")}"${cellStyle.bold ? ' font-weight="bold"' : ""}${cellStyle.italic ? ' font-style="italic"' : ""}${cellStyle.underline ? ' text-decoration="underline"' : ""}>${cellLines[cellIndex].map((line, lineIndex) => `<tspan x="${cellLeft + 8}" y="${rowTop + 20 + lineIndex * 18}">${escapeHtmlText(line)}</tspan>`).join("")}</text>`),
          (cellLeft += cellWidth));
      }),
        (rowTop += rowHeight));
    }
    height = rowTop - 112;
  }
  if (
    ((width += 36),
    (height += 162),
    width > reportOutputLimits.imageWidth ||
      height > reportOutputLimits.imageHeight ||
      width * height > reportOutputLimits.imagePixels)
  )
    throw Error(
      "Report image exceeds 4,096 × 16,384 / 16 megapixels. Narrow the range or export JSON/HTML instead.",
    );
  const populationSummary = `${reportSnapshot.request.population} population · ${reportSnapshot.rows.length} retained / ${reportSnapshot.rows.length + reportSnapshot.excludedRows.length} range rows · ${reportSnapshot.request.filter ? "filter criteria recorded" : "no population filter"}`,
    captionLines = wrapTextByCharacters(
      formatReportCaption(reportSnapshot),
      Math.floor((width - 36) / 7.4),
    ),
    headerLines = [
      ...wrapTextByCharacters(
        reportSnapshot.request.name,
        Math.floor((width - 36) / 9),
      ),
      ...captionLines,
      ...wrapTextByCharacters(
        populationSummary,
        Math.floor((width - 36) / 7.4),
      ),
    ],
    headerOffset = Math.max(0, headerLines.length * 18 + 24 - 112);
  if (
    ((height += headerOffset),
    (svgBody = `<g transform="translate(0 ${headerOffset})">${svgBody}</g>`),
    height > reportOutputLimits.imageHeight ||
      width * height > reportOutputLimits.imagePixels)
  )
    throw Error(
      "Report image exceeds its pixel budget. Narrow the range or export JSON/HTML.",
    );
  const content = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}" role="img"><title>${escapeHtmlText(reportSnapshot.request.name)}</title><desc>${escapeHtmlText(formatReportCaption(reportSnapshot))}. ${escapeHtmlText(populationSummary)}. Fixed snapshot; no automatic refresh.</desc><metadata>${escapeHtmlText(JSON.stringify({ engine: reportSnapshot.engine, hash: reportSnapshot.hash, request: reportSnapshot.request, workbookId: reportSnapshot.workbookId, revision: reportSnapshot.revision, calculatedAt: reportSnapshot.calculatedAt }))}</metadata><rect width="100%" height="100%" fill="white"/><g font-family="Arial,sans-serif">${headerLines.map((headerLine, headerIndex) => `<text x="18" y="${25 + headerIndex * 18}" font-size="${headerIndex === 0 ? 16 : 12}" fill="#362c63">${escapeHtmlText(headerLine)}</text>`).join("")}${svgBody}<text x="18" y="${height - 17}" font-size="11" fill="#545969">Fixed snapshot · ${escapeHtmlText(reportSnapshot.id)} · full values and context in the JSON export</text></g></svg>`;
  return {
    filename: toSafeFileName(reportSnapshot.request.name) + ".svg",
    mime: "image/svg+xml",
    content: content,
    width: width,
    height: height,
  };
}
function buildReportSvgFile(rawSnapshot) {
  const reportSnapshot = parseReportSnapshot(rawSnapshot);
  return {
    filename: toSafeFileName(reportSnapshot.request.name) + ".json",
    mime: "application/json",
    content: JSON.stringify(reportSnapshot, null, 2),
  };
}
function ReportOutputView({ snapshot, showData = true }) {
  const reportSnapshot = parseReportSnapshot(snapshot);
  let renderedSvg,
    renderError = "";
  try {
    renderedSvg = renderReportSvg(reportSnapshot);
  } catch (error) {
    renderError = error.message;
  }
  return i.jsxs("figure", {
    className: "lg-report-output-view",
    style: { margin: 0, color: "#202334", fontFamily: "Arial,sans-serif" },
    children: [
      i.jsxs("figcaption", {
        children: [
          i.jsx("strong", { children: reportSnapshot.request.name }),
          i.jsx("p", { children: formatReportCaption(reportSnapshot) }),
          i.jsxs("p", {
            children: [
              "Fixed snapshot · ",
              reportSnapshot.request.population,
              " population · ",
              reportSnapshot.rows.length,
              " ",
              "retained / ",
              reportSnapshot.rows.length + reportSnapshot.excludedRows.length,
              " range rows.",
              " ",
              reportSnapshot.request.filter
                ? "Filter criteria retained in context."
                : "No population filter.",
            ],
          }),
        ],
      }),
      renderedSvg
        ? i.jsx("img", {
            alt:
              reportSnapshot.request.name +
              " — " +
              formatReportCaption(reportSnapshot),
            src:
              "data:image/svg+xml;charset=utf-8," +
              encodeURIComponent(renderedSvg.content),
            width: renderedSvg.width,
            height: renderedSvg.height,
            style: { maxWidth: "100%", height: "auto", display: "block" },
          })
        : i.jsx("p", { role: "status", children: renderError }),
      showData &&
        i.jsxs("details", {
          open: !renderedSvg,
          children: [
            i.jsx("summary", {
              children: "Accessible values and report context",
            }),
            i.jsxs("table", {
              style: { borderCollapse: "collapse" },
              children: [
                i.jsxs("caption", {
                  children: [
                    reportSnapshot.sheetName,
                    "!",
                    reportSnapshot.request.range,
                  ],
                }),
                i.jsx("thead", {
                  children: i.jsxs("tr", {
                    children: [
                      i.jsx("th", { children: "Source row" }),
                      reportSnapshot.columns.map((column) =>
                        i.jsxs(
                          "th",
                          {
                            children: [
                              columnIndexToLetters(column),
                              " · column ",
                              column + 1,
                            ],
                          },
                          column,
                        ),
                      ),
                    ],
                  }),
                }),
                i.jsx("tbody", {
                  children: reportSnapshot.rows.map((row) =>
                    i.jsxs(
                      "tr",
                      {
                        children: [
                          i.jsx("th", { scope: "row", children: row.row }),
                          row.cells.map((cell) =>
                            i.jsx(
                              "td",
                              {
                                style: {
                                  border: "1px solid #ccd6fa",
                                  padding: 8,
                                  whiteSpace: "pre-wrap",
                                },
                                children: cell.text,
                              },
                              cell.address,
                            ),
                          ),
                        ],
                      },
                      row.row,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("p", {
              children: [
                "Excluded source rows: ",
                reportSnapshot.excludedRows.join(", ") || "none",
                ".",
              ],
            }),
            reportSnapshot.request.filter &&
              i.jsx("pre", {
                style: { whiteSpace: "pre-wrap" },
                children: JSON.stringify(
                  reportSnapshot.request.filter,
                  null,
                  2,
                ),
              }),
            reportSnapshot.warnings.map((warning, warningIndex) =>
              i.jsx("p", { children: warning }, warningIndex),
            ),
          ],
        }),
    ],
  });
}
function buildReportHtmlFile(rawSnapshot) {
  const reportSnapshot = parseReportSnapshot(rawSnapshot),
    bodyMarkup = c_.renderToStaticMarkup(
      i.jsx(ReportOutputView, { snapshot: reportSnapshot }),
    );
  return {
    filename: toSafeFileName(reportSnapshot.request.name) + ".html",
    mime: "text/html",
    content: `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"><title>${escapeHtmlText(reportSnapshot.request.name)}</title></head><body style="padding:18px;overflow:auto">${bodyMarkup}</body></html>`,
  };
}
async function renderReportPngBlob(rawSnapshot) {
  const svgResult = renderReportSvg(rawSnapshot),
    svgUrl = URL.createObjectURL(
      new Blob([svgResult.content], { type: "image/svg+xml" }),
    );
  try {
    const svgImage = new Image();
    await new Promise((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        ((svgImage.src = ""),
          reject(Error("Report image rendering timed out.")));
      }, 15e3);
      ((svgImage.onload = () => {
        (clearTimeout(timeoutId), resolve());
      }),
        (svgImage.onerror = () => {
          (clearTimeout(timeoutId),
            reject(
              Error(
                "The browser could not render this report image. Export SVG or HTML.",
              ),
            ));
        }),
        (svgImage.src = svgUrl));
    });
    const canvas = document.createElement("canvas");
    ((canvas.width = svgResult.width), (canvas.height = svgResult.height));
    const context = canvas.getContext("2d");
    if (!context) throw Error("PNG rendering is unavailable in this browser.");
    context.drawImage(svgImage, 0, 0);
    const blob = await new Promise((resolve, reject) => {
      const timeoutId = setTimeout(
        () => reject(Error("PNG encoding timed out.")),
        15e3,
      );
      canvas.toBlob((encodedBlob) => {
        (clearTimeout(timeoutId),
          encodedBlob
            ? resolve(encodedBlob)
            : reject(Error("PNG encoding failed.")));
      }, "image/png");
    });
    if (blob.size > 16 * 1024 * 1024)
      throw Error("PNG exceeds 16 MB. Export SVG or narrow the report.");
    return {
      filename: toSafeFileName(rawSnapshot.request.name) + ".png",
      blob: blob,
    };
  } finally {
    URL.revokeObjectURL(svgUrl);
  }
}

export {
  describeReportOutputStatus,
  publishReportOutput,
  buildReportHtmlFile,
  buildReportSvgFile,
  refreshReportOutput,
  AnalysisSnapshotNotice,
  ReportOutputView,
  removeReportOutput,
  renderReportSvg,
  renderReportPngBlob,
  DataAnalysisDialog,
  validateReportSourceScope,
};
