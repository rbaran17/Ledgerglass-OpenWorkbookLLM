import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  transformationRecipeLimits,
  fingerprintCanonicalValue,
  recipeStoreFormatId,
} from "./01-schemas-date-functions-and-validation.js";
import { formatOccupiedRangeAddress } from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  RecipeSourceSelector,
  runTransformationRecipe,
  saveTransformationRecipe,
  deleteTransformationRecipe,
  recipeStepKindLabels,
  publishRecipeOutput,
  createRecipeStep,
  formatCellValueForDisplay,
  normalizeRecipeRunRequest,
  RecipeStepEditor,
} from "./17-recipe-step-editors.js";

function TransformationRecipesDialog(props) {
  const { text, language } = useWorkbookMessages(),
    dialogRef = ne.useRef(null),
    scrollTargetRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [definition, setDefinition] = ne.useState(() => ({
      engine: recipeStoreFormatId,
      source: {
        sheetId: props.sheetId,
        range: props.range.includes(":")
          ? props.range
          : formatOccupiedRangeAddress(
              props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
            ),
      },
      steps: [],
    })),
    [recipeName, setRecipeName] = ne.useState("Transformation recipe"),
    [recipeId, setRecipeId] = ne.useState(""),
    [recipeVersion, setRecipeVersion] = ne.useState(0),
    [selectedRecipeId, setSelectedRecipeId] = ne.useState(""),
    [selectedVersion, setSelectedVersion] = ne.useState(1),
    [selectedStepIndex, setSelectedStepIndex] = ne.useState(-1),
    [newStepKind, setNewStepKind] = ne.useState("types"),
    [outputName, setOutputName] = ne.useState("Recipe result"),
    [checked, setChecked] = ne.useState(false),
    [preview, setPreview] = ne.useState(null),
    [actionError, setActionError] = ne.useState(null),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [resultTab, setResultTab] = ne.useState("output"),
    [recordPage, setRecordPage] = ne.useState(0),
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    runOptions = {
      readOnly: props.readOnly,
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: allowedSheetIds,
    },
    sheets = props.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        (runOptions.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    savedRecipes = props.workbook.transformationRecipes ?? [],
    selectedRecipe = savedRecipes.find(
      (recipe) => recipe.id === selectedRecipeId,
    ),
    loadedRecipe = savedRecipes.find((recipe) => recipe.id === recipeId),
    X$1 = loadedRecipe?.versions.find(
      (version) => version.version === recipeVersion,
    ),
    isDraftEdited =
      !X$1 || X$1.definitionHash !== fingerprintCanonicalValue(definition),
    stamp = JSON.stringify([
      props.workbook,
      props.sheetId,
      props.range,
      definition,
      recipeId,
      recipeVersion,
      outputName,
      runOptions,
    ]),
    isPreviewCurrent = preview?.stamp === stamp,
    currentRun = isPreviewCurrent ? preview.plan.run : undefined;
  function clearPreview() {
    (setPreview(null), setActionError(null), setRecordPage(0));
  }
  function replaceDefinition(nextDefinition) {
    (setDefinition(nextDefinition), clearPreview(), setStatusMessage(""));
  }
  function previewUpToStep(stepCount) {
    try {
      const partialRun = runTransformationRecipe(
        props.workbook,
        { ...definition, steps: definition.steps.slice(0, stepCount) },
        runOptions,
      );
      return partialRun.status === "complete"
        ? { fields: partialRun.output.fields, error: "" }
        : {
            fields: partialRun.failure.inputSchema.map((field) => field.name),
            error: partialRun.failure.message,
          };
    } catch (error) {
      return { fields: [], error: error.message };
    }
  }
  const stepPreview = ne.useMemo(
      () => previewUpToStep(Math.max(0, selectedStepIndex)),
      [
        props.workbook,
        definition,
        selectedStepIndex,
        JSON.stringify(runOptions),
      ],
    ),
    step = definition.steps[selectedStepIndex],
    rightFields = ne.useMemo(() => {
      if (!step || !["merge", "append"].includes(step.kind)) return [];
      try {
        const rightRun = runTransformationRecipe(
          props.workbook,
          { engine: recipeStoreFormatId, source: step.source, steps: [] },
          runOptions,
        );
        return rightRun.status === "complete" ? rightRun.output.fields : [];
      } catch {
        return [];
      }
    }, [props.workbook, step, JSON.stringify(runOptions)]);
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("input,select,button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      var scrollTarget;
      isPreviewCurrent &&
        ((scrollTarget = scrollTargetRef.current) == null
          ? undefined
          : scrollTarget.scrollIntoView
        )?.call(scrollTarget, {
          block: "start",
        });
    }, [preview, isPreviewCurrent]));
  const runAction = (action) => {
    try {
      action();
    } catch (error) {
      setActionError({ text: error.message, stamp: stamp });
    }
  };
  function onClick() {
    runAction(() => {
      const savedRecipe = savedRecipes.find(
          (recipe) => recipe.id === selectedRecipeId,
        ),
        savedVersion = savedRecipe?.versions.find(
          (version) => version.version === selectedVersion,
        );
      if (!savedRecipe || !savedVersion)
        throw Error("Choose a saved recipe version.");
      (setRecipeId(savedRecipe.id),
        setRecipeVersion(savedVersion.version),
        setRecipeName(savedRecipe.name),
        setDefinition(structuredClone(savedVersion.definition)),
        setSelectedStepIndex(savedVersion.definition.steps.length ? 0 : -1),
        clearPreview(),
        setStatusMessage(
          "Saved coordinates are unchanged by structural edits. Review the source and steps, then preview current data.",
        ));
    });
  }
  function saveRecipe() {
    runAction(() => {
      const nextWorkbook = saveTransformationRecipe(
          props.workbook,
          recipeName,
          definition,
          recipeId || undefined,
          runOptions,
        ),
        transformationRecipe = nextWorkbook.transformationRecipes.find(
          (transformationRecipe) =>
            recipeId
              ? transformationRecipe.id === recipeId
              : transformationRecipe.name === recipeName.trim(),
        );
      (props.onWorkbookChange(
        nextWorkbook,
        `Saved recipe ${transformationRecipe.name} version ${transformationRecipe.versions.length}`,
      ),
        setDefinition(transformationRecipe.versions.at(-1).definition),
        setRecipeId(transformationRecipe.id),
        setRecipeVersion(transformationRecipe.versions.length),
        setSelectedRecipeId(transformationRecipe.id),
        setSelectedVersion(transformationRecipe.versions.length),
        clearPreview(),
        setStatusMessage(
          `Saved version ${transformationRecipe.versions.length}. Preview this version before publishing.`,
        ));
    });
  }
  function previewRecipe() {
    runAction(() => {
      const recipeRunRequest = normalizeRecipeRunRequest(
        props.workbook,
        {
          definition: definition,
          ...(!isDraftEdited && recipeId
            ? { recipeId: recipeId, version: recipeVersion }
            : {}),
          outputName: outputName,
        },
        runOptions,
      );
      (setPreview({ plan: recipeRunRequest, stamp: stamp }),
        setRecordPage(0),
        setResultTab(
          recipeRunRequest.run.status === "failed" ? "steps" : "output",
        ),
        setActionError(null));
    });
  }
  function publishOutput() {
    runAction(() => {
      if (!isPreviewCurrent || !preview)
        throw Error("Preview the current recipe before publishing.");
      const nextWorkbook = publishRecipeOutput(
        props.workbook,
        preview.plan,
        runOptions,
      );
      (clearPreview(),
        props.onWorkbookChange(
          nextWorkbook,
          `Published recipe ${recipeName} version ${recipeVersion}`,
        ),
        props.onClose());
    });
  }
  function downloadReview() {
    if (!isPreviewCurrent || !preview) return;
    const { snapshot, ...reviewPlan } = preview.plan,
      objectUrl = URL.createObjectURL(
        new Blob([JSON.stringify(reviewPlan, null, 2)], {
          type: "application/json",
        }),
      ),
      downloadLink = document.createElement("a");
    ((downloadLink.href = objectUrl),
      (downloadLink.download = "Recipe-execution-review.json"),
      downloadLink.click(),
      setTimeout(() => URL.revokeObjectURL(objectUrl), 5e3));
  }
  const recipeRunSheets = props.workbook.sheets.filter(
    (sheet) => sheet.transformationRun?.recipeId === recipeId,
  );
  return i.jsx("div", {
    className: "recipe-backdrop",
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": "recipe-title",
      className: "recipe-dialog",
      ref: dialogRef,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusables = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary,[tabindex="0"]',
              ) ?? [],
            ).filter(
              (element) =>
                !element.closest("details:not([open])") ||
                element.tagName === "SUMMARY",
            ),
            firstFocusable = focusables[0],
            lastFocusable = focusables.at(-1);
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
                  id: "recipe-title",
                  children: i.jsx(LocalizedText, {
                    children: "Transformation recipes",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Save ordered steps. Review each run before publishing a separate result.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close transformation recipes"),
              onClick: props.onClose,
              lang: language("Close transformation recipes"),
              "data-ui-label": "Close transformation recipes",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "recipe-body",
          children: [
            i.jsxs("section", {
              className: "recipe-saved",
              "aria-label": text("Saved transformation recipes"),
              lang: language("Saved transformation recipes"),
              "data-ui-label": "Saved transformation recipes",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Saved recipe" }),
                    i.jsxs("select", {
                      "aria-label": text("Saved recipe"),
                      value: selectedRecipeId,
                      onChange: (event) => {
                        (setSelectedRecipeId(event.target.value),
                          setSelectedVersion(
                            savedRecipes.find(
                              (recipe) => recipe.id === event.target.value,
                            )?.versions.length ?? 1,
                          ));
                      },
                      lang: language("Saved recipe"),
                      "data-ui-label": "Saved recipe",
                      children: [
                        i.jsx("option", {
                          value: "",
                          children: "Choose a recipe",
                        }),
                        savedRecipes.map((recipe) =>
                          i.jsx(
                            "option",
                            { value: recipe.id, children: recipe.name },
                            recipe.id,
                          ),
                        ),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Version" }),
                    i.jsx("select", {
                      "aria-label": text("Saved recipe version"),
                      value: selectedVersion,
                      onChange: (event) =>
                        setSelectedVersion(Number(event.target.value)),
                      lang: language("Saved recipe version"),
                      "data-ui-label": "Saved recipe version",
                      children: (selectedRecipe?.versions ?? []).map(
                        (version) =>
                          i.jsx(
                            "option",
                            {
                              value: version.version,
                              children: version.version,
                            },
                            version.version,
                          ),
                      ),
                    }),
                  ],
                }),
                i.jsx("button", {
                  disabled: !selectedRecipe,
                  onClick: onClick,
                  children: i.jsx(LocalizedText, { children: "Load version" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (setRecipeId(""),
                      setRecipeVersion(0),
                      setRecipeName("Transformation recipe"),
                      setDefinition({
                        engine: recipeStoreFormatId,
                        source: definition.source,
                        steps: [],
                      }),
                      setSelectedStepIndex(-1),
                      clearPreview(),
                      setStatusMessage("New recipe draft."));
                  },
                  children: i.jsx(LocalizedText, { children: "New recipe" }),
                }),
                i.jsx("button", {
                  disabled: !selectedRecipe || props.readOnly,
                  onClick: () =>
                    runAction(() => {
                      const nextWorkbook = deleteTransformationRecipe(
                        props.workbook,
                        selectedRecipeId,
                        runOptions,
                      );
                      (props.onWorkbookChange(
                        nextWorkbook,
                        "Deleted saved recipe definition; outputs retained",
                      ),
                        recipeId === selectedRecipeId &&
                          (setRecipeId(""), setRecipeVersion(0)),
                        setSelectedRecipeId(""),
                        clearPreview(),
                        setStatusMessage(
                          "Saved definition deleted. Historical outputs and their run evidence are retained.",
                        ));
                    }),
                  children: i.jsx(LocalizedText, {
                    children: "Delete saved recipe",
                  }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "recipe-line",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Recipe name" }),
                    i.jsx("input", {
                      "aria-label": text("Recipe name"),
                      value: recipeName,
                      onChange: (event) => setRecipeName(event.target.value),
                      maxLength: 128,
                      lang: language("Recipe name"),
                      "data-ui-label": "Recipe name",
                    }),
                  ],
                }),
                i.jsx("p", {
                  className: "recipe-version",
                  children: recipeId
                    ? `Version ${recipeVersion}${isDraftEdited ? " · edited draft" : " · saved"}`
                    : "Unsaved draft",
                }),
                i.jsx("button", {
                  disabled: props.readOnly,
                  onClick: saveRecipe,
                  children: recipeId ? "Save next version" : "Save new recipe",
                }),
              ],
            }),
            i.jsx(RecipeSourceSelector, {
              label: text("Primary source"),
              source: definition.source,
              sheets: sheets,
              onChange: (source) =>
                replaceDefinition({ ...definition, source: source }),
            }),
            props.allowHidden &&
              i.jsxs("label", {
                className: "recipe-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (setChecked(event.target.checked), clearPreview());
                    },
                  }),
                  i.jsx(LocalizedText, {
                    children:
                      "Include hidden source cells and formula dependencies",
                  }),
                ],
              }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Sources include every record in the declared ranges, regardless of current filter views. Tables use their current headers and data body. Source cells stay unchanged.",
              }),
            }),
            i.jsxs("div", {
              className: "recipe-layout",
              children: [
                i.jsxs("aside", {
                  "aria-label": text("Ordered recipe steps"),
                  lang: language("Ordered recipe steps"),
                  "data-ui-label": "Ordered recipe steps",
                  children: [
                    i.jsx("h3", {
                      children: i.jsx(LocalizedText, {
                        children: "Applied steps",
                      }),
                    }),
                    i.jsx("ol", {
                      children: definition.steps.map((step, stepIndex) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsxs("button", {
                                "aria-pressed": selectedStepIndex === stepIndex,
                                onClick: () => setSelectedStepIndex(stepIndex),
                                children: [
                                  stepIndex + 1,
                                  ". ",
                                  recipeStepKindLabels[step.kind],
                                ],
                              }),
                              i.jsxs("span", {
                                children: [
                                  i.jsx("button", {
                                    "aria-label": `Move step ${stepIndex + 1} up`,
                                    disabled: stepIndex === 0,
                                    onClick: () => {
                                      const steps = [...definition.steps];
                                      (([
                                        steps[stepIndex - 1],
                                        steps[stepIndex],
                                      ] = [
                                        steps[stepIndex],
                                        steps[stepIndex - 1],
                                      ]),
                                        replaceDefinition({
                                          ...definition,
                                          steps: steps,
                                        }),
                                        setSelectedStepIndex(stepIndex - 1));
                                    },
                                    children: "↑",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move step ${stepIndex + 1} down`,
                                    disabled:
                                      stepIndex === definition.steps.length - 1,
                                    onClick: () => {
                                      const steps = [...definition.steps];
                                      (([
                                        steps[stepIndex],
                                        steps[stepIndex + 1],
                                      ] = [
                                        steps[stepIndex + 1],
                                        steps[stepIndex],
                                      ]),
                                        replaceDefinition({
                                          ...definition,
                                          steps: steps,
                                        }),
                                        setSelectedStepIndex(stepIndex + 1));
                                    },
                                    children: "↓",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Remove step ${stepIndex + 1}`,
                                    onClick: () => {
                                      (replaceDefinition({
                                        ...definition,
                                        steps: definition.steps.filter(
                                          (step, filterIndex) =>
                                            filterIndex !== stepIndex,
                                        ),
                                      }),
                                        setSelectedStepIndex(
                                          Math.min(
                                            stepIndex,
                                            definition.steps.length - 2,
                                          ),
                                        ));
                                    },
                                    children: "×",
                                  }),
                                ],
                              }),
                            ],
                          },
                          step.id,
                        ),
                      ),
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Add transformation",
                        }),
                        i.jsx("select", {
                          "aria-label": text("New transformation type"),
                          value: newStepKind,
                          onChange: (event) =>
                            setNewStepKind(event.target.value),
                          lang: language("New transformation type"),
                          "data-ui-label": "New transformation type",
                          children: Object.entries(recipeStepKindLabels).map(
                            ([kind, label]) =>
                              i.jsx(
                                "option",
                                { value: kind, children: label },
                                kind,
                              ),
                          ),
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      disabled:
                        definition.steps.length >=
                        transformationRecipeLimits.steps,
                      onClick: () => {
                        const fields = previewUpToStep(
                          definition.steps.length,
                        ).fields;
                        (replaceDefinition({
                          ...definition,
                          steps: [
                            ...definition.steps,
                            createRecipeStep(
                              newStepKind,
                              fields,
                              definition.source,
                            ),
                          ],
                        }),
                          setSelectedStepIndex(definition.steps.length));
                      },
                      children: i.jsx(LocalizedText, { children: "Add step" }),
                    }),
                    i.jsxs("p", {
                      children: [
                        definition.steps.length,
                        " / ",
                        transformationRecipeLimits.steps,
                        " ",
                        i.jsx(LocalizedText, { children: "steps" }),
                      ],
                    }),
                  ],
                }),
                i.jsx("section", {
                  className: "recipe-editor",
                  "aria-label": text("Recipe step editor"),
                  lang: language("Recipe step editor"),
                  "data-ui-label": "Recipe step editor",
                  children: step
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("h3", {
                            children: [
                              selectedStepIndex + 1,
                              ". ",
                              recipeStepKindLabels[step.kind],
                            ],
                          }),
                          stepPreview.error &&
                            i.jsxs("p", {
                              role: "status",
                              children: [
                                i.jsx(LocalizedText, {
                                  children:
                                    "Earlier step or source needs attention: ",
                                }),
                                stepPreview.error,
                              ],
                            }),
                          i.jsx(RecipeStepEditor, {
                            step: step,
                            fields: stepPreview.fields,
                            rightFields: rightFields,
                            sheets: sheets,
                            onChange: (nextStep) =>
                              replaceDefinition({
                                ...definition,
                                steps: definition.steps.map(
                                  (step, stepPosition) =>
                                    stepPosition === selectedStepIndex
                                      ? nextStep
                                      : step,
                                ),
                              }),
                          }),
                        ],
                      })
                    : i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Add a transformation to define the sequence. Each step reads the previous step's complete result.",
                        }),
                      }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "recipe-line",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "New output worksheet" }),
                    i.jsx("input", {
                      "aria-label": text("Recipe output worksheet"),
                      value: outputName,
                      onChange: (event) => {
                        (setOutputName(event.target.value), clearPreview());
                      },
                      maxLength: 31,
                      lang: language("Recipe output worksheet"),
                      "data-ui-label": "Recipe output worksheet",
                    }),
                  ],
                }),
                i.jsx("button", {
                  className: "recipe-primary",
                  onClick: previewRecipe,
                  children: i.jsx(LocalizedText, {
                    children: "Preview recipe",
                  }),
                }),
              ],
            }),
            statusMessage &&
              i.jsx("p", { role: "status", children: statusMessage }),
            actionError?.stamp === stamp &&
              i.jsx("p", {
                role: "alert",
                className: "recipe-error",
                children: actionError.text,
              }),
            preview &&
              !isPreviewCurrent &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Workbook, selection or permissions changed. Preview again; previous results are withheld.",
                }),
              }),
            recipeRunSheets.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Previous saved outputs (",
                      }),
                      recipeRunSheets.length,
                      ")",
                    ],
                  }),
                  recipeRunSheets.map((sheet) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          sheet.name,
                          " ",
                          i.jsx(LocalizedText, { children: "· version " }),
                          sheet.transformationRun.version,
                          " ·",
                          " ",
                          sheet.transformationRun.outputRows,
                          " ",
                          i.jsx(LocalizedText, { children: "rows ·" }),
                          " ",
                          sheet.transformationRun.at,
                          i.jsx(LocalizedText, {
                            children:
                              ". Historical result; use a fresh preview to compare.",
                          }),
                        ],
                      },
                      sheet.id,
                    ),
                  ),
                ],
              }),
            currentRun &&
              i.jsxs("section", {
                className: "recipe-review",
                "aria-label": text("Recipe execution preview"),
                ref: scrollTargetRef,
                lang: language("Recipe execution preview"),
                "data-ui-label": "Recipe execution preview",
                children: [
                  i.jsx("h3", {
                    children:
                      currentRun.status === "complete"
                        ? `Complete preview · ${currentRun.output.rows.length} records`
                        : `Stopped at ${currentRun.failure.index === 0 ? "source" : `step ${currentRun.failure.index}`}`,
                  }),
                  currentRun.status === "failed" &&
                    i.jsxs("p", {
                      role: "alert",
                      className: "recipe-error",
                      children: [
                        currentRun.failure.message,
                        " ",
                        i.jsx(LocalizedText, {
                          children:
                            "No output will be published. Downstream steps were skipped.",
                        }),
                      ],
                    }),
                  isDraftEdited &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Draft preview. Save a version and preview it before publishing.",
                      }),
                    }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Definition fingerprint ",
                      }),
                      i.jsx("code", { children: currentRun.definitionHash }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Source fingerprint " }),
                      i.jsx("code", { children: currentRun.sourceHash }),
                    ],
                  }),
                  i.jsx("nav", {
                    "aria-label": text("Recipe review views"),
                    lang: language("Recipe review views"),
                    "data-ui-label": "Recipe review views",
                    children: [
                      ["output", "Output records"],
                      ["steps", "Step schemas & results"],
                      ["sources", "Sources & control totals"],
                    ].map(([view, label]) =>
                      i.jsx(
                        "button",
                        {
                          "aria-pressed": resultTab === view,
                          disabled:
                            view === "output" &&
                            currentRun.status !== "complete",
                          onClick: () => setResultTab(view),
                          children: label,
                        },
                        view,
                      ),
                    ),
                  }),
                  resultTab === "output" &&
                    currentRun.status === "complete" &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("div", {
                          className: "recipe-table",
                          tabIndex: 0,
                          "aria-label": text("Complete recipe output"),
                          lang: language("Complete recipe output"),
                          "data-ui-label": "Complete recipe output",
                          children: [
                            i.jsxs("table", {
                              children: [
                                i.jsx("thead", {
                                  children: i.jsxs("tr", {
                                    children: [
                                      currentRun.output.fields.map((field) =>
                                        i.jsx("th", { children: field }, field),
                                      ),
                                      i.jsx("th", {
                                        children: i.jsx(LocalizedText, {
                                          children: "Original source rows",
                                        }),
                                      }),
                                    ],
                                  }),
                                }),
                                i.jsx("tbody", {
                                  children: currentRun.output.rows
                                    .slice(
                                      recordPage * 25,
                                      (recordPage + 1) * 25,
                                    )
                                    .map((row, rowIndex) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            row.values.map(
                                              (value, valueIndex) =>
                                                i.jsx(
                                                  "td",
                                                  {
                                                    children:
                                                      formatCellValueForDisplay(
                                                        value,
                                                      ),
                                                  },
                                                  valueIndex,
                                                ),
                                            ),
                                            i.jsx("td", {
                                              children: i.jsxs("details", {
                                                children: [
                                                  i.jsxs("summary", {
                                                    children: [
                                                      row.sources.length,
                                                      " ",
                                                      i.jsx(LocalizedText, {
                                                        children: "source row",
                                                      }),
                                                      row.sources.length === 1
                                                        ? ""
                                                        : "s",
                                                    ],
                                                  }),
                                                  row.sources.map((source) =>
                                                    i.jsx(
                                                      "p",
                                                      { children: source },
                                                      source,
                                                    ),
                                                  ),
                                                ],
                                              }),
                                            }),
                                          ],
                                        },
                                        rowIndex,
                                      ),
                                    ),
                                }),
                              ],
                            }),
                            !currentRun.output.rows.length &&
                              i.jsx("p", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "No records remain. Source records are retained.",
                                }),
                              }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "recipe-line",
                          children: [
                            i.jsx("button", {
                              disabled: recordPage === 0,
                              onClick: () => setRecordPage((page) => page - 1),
                              children: i.jsx(LocalizedText, {
                                children: "Previous recipe records",
                              }),
                            }),
                            i.jsxs("span", {
                              children: [
                                i.jsx(LocalizedText, { children: "Page " }),
                                recordPage + 1,
                                " ",
                                i.jsx(LocalizedText, { children: "of" }),
                                " ",
                                Math.max(
                                  1,
                                  Math.ceil(currentRun.output.rows.length / 25),
                                ),
                                " ·",
                                " ",
                                currentRun.output.rows.length,
                                " ",
                                i.jsx(LocalizedText, { children: "records" }),
                              ],
                            }),
                            i.jsx("button", {
                              disabled:
                                (recordPage + 1) * 25 >=
                                currentRun.output.rows.length,
                              onClick: () => setRecordPage((page) => page + 1),
                              children: i.jsx(LocalizedText, {
                                children: "Next recipe records",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  resultTab === "steps" &&
                    i.jsx("div", {
                      "aria-label": text("Recipe step trace"),
                      lang: language("Recipe step trace"),
                      "data-ui-label": "Recipe step trace",
                      children: currentRun.steps.map((step) =>
                        i.jsxs(
                          "details",
                          {
                            open: step.status === "failed",
                            children: [
                              i.jsxs("summary", {
                                children: [
                                  step.index,
                                  ".",
                                  " ",
                                  recipeStepKindLabels[step.kind] ?? step.kind,
                                  " ",
                                  "· ",
                                  step.status,
                                  " · ",
                                  step.inputRows,
                                  " → ",
                                  step.outputRows ?? "—",
                                  " ",
                                  i.jsx(LocalizedText, { children: "rows" }),
                                ],
                              }),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Input schema:",
                                  }),
                                  " ",
                                  step.inputSchema
                                    .map(
                                      (field) =>
                                        `${field.name} (${field.types.join(" / ") || "empty"}${field.declaredType ? `; declared ${field.declaredType}` : ""})`,
                                    )
                                    .join(", ") || "not executed",
                                ],
                              }),
                              step.outputSchema &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Output schema:",
                                    }),
                                    " ",
                                    step.outputSchema
                                      .map(
                                        (field) =>
                                          `${field.name} (${field.types.join(" / ") || "empty"})`,
                                      )
                                      .join(", "),
                                  ],
                                }),
                              step.error &&
                                i.jsx("p", {
                                  className: "recipe-error",
                                  children: step.error,
                                }),
                              step.controls &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Control totals:",
                                    }),
                                    " ",
                                    step.controls
                                      .map(
                                        (control) =>
                                          `${control.field}: ${control.sum} across ${control.numericCount} numeric values; ${control.blankCount} blanks`,
                                      )
                                      .join(" · "),
                                  ],
                                }),
                              step.details?.matchPairs !== undefined &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Merge: ",
                                    }),
                                    String(step.details.matchPairs),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "pairs · matched",
                                    }),
                                    " ",
                                    String(step.details.matchedLeft),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "left /",
                                    }),
                                    " ",
                                    String(step.details.matchedRight),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "right · unmatched",
                                    }),
                                    " ",
                                    step.details.unmatchedLeft.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "left /",
                                    }),
                                    " ",
                                    step.details.unmatchedRight.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "right · ",
                                    }),
                                    String(step.details.manyToMany),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "many-to-many groups ·",
                                    }),
                                    " ",
                                    String(step.details.extraPairs),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children:
                                        "extra pairs. The saved review contains every match explanation and exception.",
                                    }),
                                  ],
                                }),
                              step.outputHash &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Output fingerprint ",
                                    }),
                                    i.jsx("code", {
                                      children: step.outputHash,
                                    }),
                                  ],
                                }),
                            ],
                          },
                          step.id,
                        ),
                      ),
                    }),
                  resultTab === "sources" &&
                    i.jsxs("div", {
                      "aria-label": text("Recipe sources and totals"),
                      lang: language("Recipe sources and totals"),
                      "data-ui-label": "Recipe sources and totals",
                      children: [
                        currentRun.sources.map((source, sourceIndex) =>
                          i.jsxs(
                            "details",
                            {
                              open: true,
                              children: [
                                i.jsxs("summary", {
                                  children: [
                                    source.sheetName,
                                    "!",
                                    source.resolvedRange,
                                    " · ",
                                    source.rows,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "source records",
                                    }),
                                  ],
                                }),
                                i.jsx("p", {
                                  children: source.schema
                                    .map(
                                      (field) =>
                                        `${field.name}: ${field.types.join(" / ") || "empty"}`,
                                    )
                                    .join(", "),
                                }),
                                i.jsx("p", {
                                  children: source.controls
                                    .map(
                                      (control) =>
                                        `${control.field}: ${control.sum} (${control.numericCount} numbers, ${control.blankCount} blanks)`,
                                    )
                                    .join(" · "),
                                }),
                                i.jsx("code", { children: source.hash }),
                              ],
                            },
                            sourceIndex,
                          ),
                        ),
                        currentRun.status === "complete" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("h4", {
                                children: i.jsx(LocalizedText, {
                                  children: "Final controls",
                                }),
                              }),
                              currentRun.controls.map((control) =>
                                i.jsxs(
                                  "p",
                                  {
                                    children: [
                                      control.field,
                                      i.jsx(LocalizedText, {
                                        children: ": sum ",
                                      }),
                                      control.sum,
                                      " · ",
                                      control.numericCount,
                                      " ",
                                      i.jsx(LocalizedText, {
                                        children: "numeric values · ",
                                      }),
                                      control.blankCount,
                                      " ",
                                      i.jsx(LocalizedText, {
                                        children: "blanks",
                                      }),
                                    ],
                                  },
                                  control.field,
                                ),
                              ),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Output fingerprint ",
                                  }),
                                  i.jsx("code", {
                                    children: currentRun.outputHash,
                                  }),
                                ],
                              }),
                            ],
                          }),
                      ],
                    }),
                  i.jsx("button", {
                    onClick: downloadReview,
                    children: i.jsx(LocalizedText, {
                      children: "Save recipe review",
                    }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Output is a separate static snapshot. Publishing preserves prior outputs and records one Undo action. Saved steps, fingerprints and lineage are retained in native JSON; XLSX preserves the supported values and notes.",
                    }),
                  }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: props.readOnly
                ? "Read-only host"
                : isDraftEdited
                  ? "Save a version before publishing"
                  : "Review the current version before publishing",
            }),
            i.jsx("button", {
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "recipe-primary",
              disabled:
                props.readOnly ||
                !isPreviewCurrent ||
                isDraftEdited ||
                currentRun?.status !== "complete",
              onClick: publishOutput,
              children: i.jsx(LocalizedText, {
                children: "Publish recipe output",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const createDefaultFilterCriterion = (column) => ({
    column: column,
    type: "text",
    operator: "contains",
    value: "",
  }),
  filterOperatorLabels = {
    equals: "Equals",
    notEquals: "Does not equal",
    contains: "Contains",
    notContains: "Does not contain",
    startsWith: "Begins with",
    endsWith: "Ends with",
    wildcard: "Matches wildcard (* ? ~)",
    gt: "Greater than / after",
    gte: "At least / on or after",
    lt: "Less than / before",
    lte: "At most / on or before",
    blank: "Is blank",
    notBlank: "Is not blank",
  };

export {
  createDefaultFilterCriterion,
  TransformationRecipesDialog,
  filterOperatorLabels,
};
