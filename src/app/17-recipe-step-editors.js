import i from "react/jsx-runtime";
import {
  parseTransformationRecipe,
  validateSavedRecipeMetadata,
  fingerprintCanonicalValue,
  canonicalJsonStringify,
  recipeStoreFormatId,
  normalizeDatePolicy,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  createWorksheetRecord,
  formatOccupiedRangeAddress,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { inspectOutputNameSafety } from "./12-consolidation-and-dependency-inspection.js";
import {
  executeRecipeStep,
  resolveRecipeSourceTable,
  validateRecipeOutputTable,
  profileFieldTypes,
  buildFieldControlTotals,
} from "./16-transformation-recipe-engine.js";

function runTransformationRecipe(workbook, definition, options = {}) {
  const transformationRecipe = parseTransformationRecipe(definition),
    sources = [],
    steps = [],
    sourceCache = new Map(),
    readSourceTable = (source) => {
      const sourceKey = canonicalJsonStringify(source);
      let sourceData = sourceCache.get(sourceKey);
      if (!sourceData) {
        const recipeSourceTable = resolveRecipeSourceTable(
          workbook,
          source,
          options,
        );
        ((sourceData = recipeSourceTable.data),
          sources.push(recipeSourceTable.evidence),
          sourceCache.set(sourceKey, sourceData));
      }
      return structuredClone(sourceData);
    },
    buildRunSummary = () => ({
      engine: recipeStoreFormatId,
      definitionHash: fingerprintCanonicalValue(transformationRecipe),
      sourceHash: fingerprintCanonicalValue({
        datePolicy: normalizeDatePolicy(workbook.datePolicy),
        sources: sources,
      }),
      sources: sources,
      steps: steps,
    });
  let output;
  try {
    output = readSourceTable(transformationRecipe.source);
  } catch (error) {
    return {
      ...buildRunSummary(),
      status: "failed",
      failure: {
        index: 0,
        kind: "source",
        message: error.message,
        inputSchema: [],
        inputRows: 0,
      },
    };
  }
  for (const [stepIndex, step] of transformationRecipe.steps.entries()) {
    const stepReport = {
      index: stepIndex + 1,
      id: step.id,
      kind: step.kind,
      status: "failed",
      inputSchema: profileFieldTypes(output),
      inputRows: output.rows.length,
      inputHash: fingerprintCanonicalValue(output),
    };
    try {
      const stepResult = executeRecipeStep(
        workbook,
        output,
        step,
        readSourceTable,
      );
      output = validateRecipeOutputTable(stepResult.data);
      const fieldControlTotals = buildFieldControlTotals(output);
      (Object.assign(stepReport, {
        status: "complete",
        outputSchema: profileFieldTypes(output),
        outputRows: output.rows.length,
        outputHash: fingerprintCanonicalValue(output),
        controls: fieldControlTotals,
        ...(stepResult.details ? { details: stepResult.details } : {}),
      }),
        steps.push(stepReport));
    } catch (error) {
      return (
        (stepReport.error = error.message),
        steps.push(
          stepReport,
          ...transformationRecipe.steps
            .slice(stepIndex + 1)
            .map((remainingStep, offset) => ({
              index: stepIndex + offset + 2,
              id: remainingStep.id,
              kind: remainingStep.kind,
              status: "skipped",
              inputSchema: [],
              inputRows: 0,
            })),
        ),
        {
          ...buildRunSummary(),
          status: "failed",
          failure: {
            index: stepIndex + 1,
            kind: step.kind,
            message: stepReport.error,
            inputSchema: stepReport.inputSchema,
            inputRows: stepReport.inputRows,
          },
        }
      );
    }
  }
  return {
    ...buildRunSummary(),
    status: "complete",
    output: output,
    outputHash: fingerprintCanonicalValue(output),
    controls: buildFieldControlTotals(output),
  };
}
function kI(workbook, readOnly = false) {
  if (readOnly) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(workbook);
}
function saveTransformationRecipe(
  workbook,
  recipeName,
  definition,
  recipeId,
  options = {},
) {
  kI(workbook, !!options.readOnly);
  const transformationRecipe = parseTransformationRecipe(definition);
  if (
    ((recipeName = recipeName.trim()), !recipeName || recipeName.length > 128)
  )
    throw Error("Use a recipe name of 1–128 characters.");
  const transformationRecipes = structuredClone(
      workbook.transformationRecipes ?? [],
    ),
    existingRecipe = recipeId
      ? transformationRecipes.find(
          (transformationRecipe) => transformationRecipe.id === recipeId,
        )
      : undefined;
  if (recipeId && !existingRecipe)
    throw Error("Saved recipe no longer exists.");
  if (
    transformationRecipes.some(
      (transformationRecipe) =>
        transformationRecipe.id !== recipeId &&
        transformationRecipe.name.toLowerCase() === recipeName.toLowerCase(),
    )
  )
    throw Error("A recipe with this name already exists.");
  const recipeVersion = {
    version: (existingRecipe?.versions.length ?? 0) + 1,
    definition: transformationRecipe,
    definitionHash: fingerprintCanonicalValue(transformationRecipe),
    createdAt: new Date().toISOString(),
  };
  existingRecipe
    ? ((existingRecipe.name = recipeName),
      existingRecipe.versions.push(recipeVersion))
    : transformationRecipes.push({
        id: crypto.randomUUID(),
        name: recipeName,
        versions: [recipeVersion],
      });
  const nextWorkbook = {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => ({ ...sheet })),
    transformationRecipes: transformationRecipes,
  };
  return (
    validateSavedRecipeMetadata(nextWorkbook, nextWorkbook),
    nextWorkbook
  );
}
function deleteTransformationRecipe(workbook, recipeId, options = {}) {
  if (
    (kI(workbook, !!options.readOnly),
    !workbook.transformationRecipes?.some(
      (transformationRecipe) => transformationRecipe.id === recipeId,
    ))
  )
    throw Error("Saved recipe no longer exists.");
  return {
    ...workbook,
    transformationRecipes: workbook.transformationRecipes.filter(
      (transformationRecipe) => transformationRecipe.id !== recipeId,
    ),
  };
}
function normalizeRecipeRunRequest(workbook, request, options = {}) {
  const transformationRecipe = parseTransformationRecipe(request.definition);
  if (
    Object.keys(request).some(
      (fieldName) =>
        !["definition", "recipeId", "version", "outputName"].includes(
          fieldName,
        ),
    )
  )
    throw Error("Unknown recipe request field.");
  if ((request.recipeId === undefined) != (request.version === undefined))
    throw Error("Choose a saved recipe ID and version together.");
  if (request.recipeId) {
    const savedVersion = workbook.transformationRecipes
      ?.find(
        (transformationRecipe) => transformationRecipe.id === request.recipeId,
      )
      ?.versions.find((version) => version.version === request.version);
    if (
      !savedVersion ||
      savedVersion.definitionHash !==
        fingerprintCanonicalValue(transformationRecipe)
    )
      throw Error(
        "The recipe differs from the saved version. Save a new version or reload the selected version.",
      );
  }
  if (
    typeof request.outputName != "string" ||
    !request.outputName.trim() ||
    request.outputName !== request.outputName.trim() ||
    request.outputName.length > 31 ||
    /[\\/\[\]*?:\u0000-\u001f]/.test(request.outputName) ||
    /^'|'$/.test(request.outputName)
  )
    throw Error("Choose an unpadded output worksheet name of 1–31 characters.");
  const context = {
    includeHidden: !!options.includeHidden,
    allowedSheetIds: [
      ...new Set(
        options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
  };
  return {
    snapshot: JSON.stringify(workbook),
    request: { ...request, definition: transformationRecipe },
    context: context,
    run: runTransformationRecipe(workbook, transformationRecipe, options),
  };
}
function publishRecipeOutput(workbook, plan, options = {}) {
  if (
    (kI(workbook, !!options.readOnly),
    plan.snapshot !== JSON.stringify(workbook))
  )
    throw Error("Workbook changed after preview. Preview the recipe again.");
  if (!plan.request.recipeId || !plan.request.version)
    throw Error("Save a recipe version before publishing its output.");
  const recipeRunRequest = normalizeRecipeRunRequest(
    workbook,
    plan.request,
    options,
  );
  if (canonicalJsonStringify(recipeRunRequest) !== canonicalJsonStringify(plan))
    throw Error("Recipe preview or authorization changed. Preview again.");
  if (recipeRunRequest.run.status !== "complete")
    throw Error(
      "A failed recipe cannot publish output. Repair the failed step and preview again.",
    );
  const outputName = recipeRunRequest.request.outputName;
  if (
    workbook.sheets.some(
      (sheet) => sheet.name.toLowerCase() === outputName.toLowerCase(),
    )
  )
    throw Error(
      "Output worksheet already exists. Choose a new name; previous successful output is preserved.",
    );
  inspectOutputNameSafety(workbook, outputName);
  const output = recipeRunRequest.run.output;
  if (
    workbook.sheets.reduce(
      (total, existingSheet) => total + Object.keys(existingSheet.cells).length,
      0,
    ) +
      (output.rows.length + 1) * output.fields.length >
    workbookLimits.maxCells
  )
    throw Error("Recipe output exceeds workbook cell capacity.");
  const worksheetRecord = createWorksheetRecord(outputName);
  ((worksheetRecord.id = `recipe-${fingerprintCanonicalValue([workbook.id, workbook.revision, outputName, recipeRunRequest.run.sourceHash, recipeRunRequest.run.definitionHash]).slice(0, 24)}`),
    output.fields.forEach(
      (field, columnIndex) =>
        (worksheetRecord.cells[`${columnIndexToLetters(columnIndex)}1`] = {
          raw: field,
          ...(field.startsWith("=") ? { literal: true } : {}),
          style: { bold: true, fill: "#eee9fa" },
        }),
    ));
  for (const [rowIndex, row] of output.rows.entries()) {
    const note = `Recipe ${recipeRunRequest.request.recipeId} version ${recipeRunRequest.request.version}
Original source rows: ${row.sources.join("; ")}`;
    if (note.length > 32767)
      throw Error("Recipe source notes exceed the cell note budget.");
    for (const [valueIndex, raw] of row.values.entries())
      worksheetRecord.cells[
        `${columnIndexToLetters(valueIndex)}${rowIndex + 2}`
      ] = {
        raw: raw,
        ...(typeof raw == "string" && raw.startsWith("=")
          ? { literal: true }
          : {}),
        note: note,
        ...(output.declaredTypes[output.fields[valueIndex]] === "date"
          ? { format: "date" }
          : {}),
      };
  }
  const transformationRecipe = workbook.transformationRecipes.find(
    (transformationRecipe) =>
      transformationRecipe.id === recipeRunRequest.request.recipeId,
  );
  worksheetRecord.transformationRun = {
    engine: recipeStoreFormatId,
    recipeId: transformationRecipe.id,
    recipeName: transformationRecipe.name,
    version: recipeRunRequest.request.version,
    definition: recipeRunRequest.request.definition,
    definitionHash: recipeRunRequest.run.definitionHash,
    sourceWorkbookId: workbook.id,
    sourceRevision: workbook.revision,
    sourceHash: recipeRunRequest.run.sourceHash,
    outputHash: recipeRunRequest.run.outputHash,
    outputRows: output.rows.length,
    outputRange: `A1:${columnIndexToLetters(output.fields.length - 1)}${output.rows.length + 1}`,
    at: new Date().toISOString(),
    steps: recipeRunRequest.run.steps,
    controls: recipeRunRequest.run.controls,
    lineage: Object.fromEntries(
      output.rows.map((row, lineageRowIndex) => [
        lineageRowIndex + 2,
        row.sources,
      ]),
    ),
  };
  const nextWorkbook = {
    ...workbook,
    sheets: [...workbook.sheets, worksheetRecord],
  };
  return normalizeWorkbookInput(nextWorkbook);
}
const recipeStepKindLabels = {
  select: "Select columns",
  rename: "Rename columns",
  filter: "Filter rows",
  types: "Change types",
  replace: "Replace values",
  derive: "Custom column",
  sort: "Sort rows",
  unpivot: "Unpivot columns",
  group: "Group rows",
  pivot: "Pivot column",
  append: "Append rows",
  merge: "Merge populations",
};
function RecipeFieldChecklist({ label, fields, selected, onChange }) {
  return i.jsxs("fieldset", {
    className: "recipe-fields",
    children: [
      i.jsx("legend", { children: label }),
      fields.length
        ? fields.map((field) =>
            i.jsxs(
              "label",
              {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    "aria-label": `${label}: ${field}`,
                    checked: selected.includes(field),
                    onChange: (event) =>
                      onChange(
                        event.target.checked
                          ? [...selected, field]
                          : selected.filter(
                              (selectedField) => selectedField !== field,
                            ),
                      ),
                  }),
                  field,
                ],
              },
              field,
            ),
          )
        : i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children: "Preview or repair the source to obtain fields.",
            }),
          }),
    ],
  });
}
function RecipeSourceSelector({ label, source, sheets, onChange }) {
  const sheet = sheets.find((sheet) => sheet.id === source.sheetId);
  return i.jsxs("fieldset", {
    className: "recipe-source",
    children: [
      i.jsx("legend", { children: label }),
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Worksheet" }),
          i.jsxs("select", {
            "aria-label": `${label} worksheet`,
            value: source.sheetId,
            onChange: (event) => {
              const sheet = sheets.find(
                (sheet) => sheet.id === event.target.value,
              );
              onChange({
                sheetId: event.target.value,
                range: sheet ? formatOccupiedRangeAddress(sheet) : "A1:B10",
              });
            },
            children: [
              i.jsx("option", { value: "", children: "Choose worksheet" }),
              sheets.map((sheet) =>
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
          i.jsx(LocalizedText, { children: "Source type" }),
          i.jsxs("select", {
            "aria-label": `${label} table`,
            value: source.tableId ?? "",
            onChange: (event) =>
              onChange({ ...source, tableId: event.target.value || undefined }),
            children: [
              i.jsx("option", { value: "", children: "Explicit headed range" }),
              sheet?.tables?.map((table) =>
                i.jsxs(
                  "option",
                  {
                    value: table.id,
                    children: [table.name, " · current table body"],
                  },
                  table.id,
                ),
              ),
            ],
          }),
        ],
      }),
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Headed range" }),
          i.jsx("input", {
            "aria-label": `${label} range`,
            value: source.range,
            disabled: !!source.tableId,
            onChange: (event) =>
              onChange({ ...source, range: event.target.value }),
          }),
        ],
      }),
    ],
  });
}
function RecipeValueInput({ label, value, onChange }) {
  const valueType =
    value === null
      ? "blank"
      : typeof value == "number"
        ? "number"
        : typeof value == "boolean"
          ? "boolean"
          : "text";
  return i.jsxs("span", {
    className: "recipe-typed",
    children: [
      i.jsxs("label", {
        children: [
          label,
          " ",
          i.jsx(LocalizedText, { children: "type" }),
          i.jsxs("select", {
            "aria-label": `${label} type`,
            value: valueType,
            onChange: (event) =>
              onChange(
                event.target.value === "blank"
                  ? null
                  : event.target.value === "number"
                    ? 0
                    : event.target.value === "boolean"
                      ? true
                      : "",
              ),
            children: [
              i.jsx("option", { value: "text", children: "Text" }),
              i.jsx("option", { value: "number", children: "Number" }),
              i.jsx("option", { value: "boolean", children: "Boolean" }),
              i.jsx("option", { value: "blank", children: "Blank" }),
            ],
          }),
        ],
      }),
      valueType !== "blank" &&
        (valueType === "boolean"
          ? i.jsxs("label", {
              children: [
                label,
                i.jsxs("select", {
                  "aria-label": label,
                  value: String(value),
                  onChange: (event) => onChange(event.target.value === "true"),
                  children: [
                    i.jsx("option", { value: "true", children: "TRUE" }),
                    i.jsx("option", { value: "false", children: "FALSE" }),
                  ],
                }),
              ],
            })
          : i.jsxs("label", {
              children: [
                label,
                i.jsx("input", {
                  "aria-label": label,
                  type: valueType === "number" ? "number" : "text",
                  value:
                    typeof value == "number" && !Number.isFinite(value)
                      ? ""
                      : String(value ?? ""),
                  onChange: (event) =>
                    onChange(
                      valueType === "number"
                        ? event.target.value === ""
                          ? NaN
                          : Number(event.target.value)
                        : event.target.value,
                    ),
                }),
              ],
            })),
    ],
  });
}
const recipeAggregateOperations = [
  "sum",
  "count",
  "average",
  "min",
  "max",
  "distinctCount",
];
function RecipeStepEditor({ step, fields, rightFields, sheets, onChange }) {
  const { text, language } = useWorkbookMessages(),
    updateStep = (patch) => onChange({ ...step, ...patch }),
    renderFieldSelect = (
      selectLabel,
      selectedValue,
      fieldOptions,
      onFieldChange,
    ) =>
      i.jsxs("label", {
        children: [
          selectLabel,
          i.jsxs("select", {
            "aria-label": selectLabel,
            value: selectedValue,
            onChange: (event) => onFieldChange(event.target.value),
            children: [
              i.jsx("option", { value: "", children: "Choose field" }),
              fieldOptions.map((fieldOption) =>
                i.jsx(
                  "option",
                  { value: fieldOption, children: fieldOption },
                  fieldOption,
                ),
              ),
            ],
          }),
        ],
      }),
    renderTextInput = (inputLabel, inputValue, onInputChange) =>
      i.jsxs("label", {
        children: [
          inputLabel,
          i.jsx("input", {
            "aria-label": inputLabel,
            value: inputValue,
            onChange: (event) => onInputChange(event.target.value),
          }),
        ],
      }),
    renderFieldChecklist = (checklistLabel, selected, onChange) =>
      i.jsx(RecipeFieldChecklist, {
        label: checklistLabel,
        fields,
        selected: selected,
        onChange: onChange,
      });
  switch (step.kind) {
    case "select":
      return renderFieldChecklist("Keep columns", step.fields, (fields) =>
        updateStep({ fields: fields }),
      );
    case "rename":
      return i.jsxs(i.Fragment, {
        children: [
          step.mappings.map((mapping, mappingIndex) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  renderFieldSelect(
                    `Rename source ${mappingIndex + 1}`,
                    mapping.from,
                    fields,
                    (from) =>
                      updateStep({
                        mappings: step.mappings.map(
                          (mapping, mappingEntryIndex) =>
                            mappingIndex === mappingEntryIndex
                              ? { ...mapping, from: from }
                              : mapping,
                        ),
                      }),
                  ),
                  renderTextInput(
                    `Rename to ${mappingIndex + 1}`,
                    mapping.to,
                    (renamedTo) =>
                      updateStep({
                        mappings: step.mappings.map(
                          (mapping, toMappingIndex) =>
                            mappingIndex === toMappingIndex
                              ? { ...mapping, to: renamedTo }
                              : mapping,
                        ),
                      }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      updateStep({
                        mappings: step.mappings.filter(
                          (mapping, removeIndex) =>
                            mappingIndex !== removeIndex,
                        ),
                      }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove rename " }),
                      mappingIndex + 1,
                    ],
                  }),
                ],
              },
              mappingIndex,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              updateStep({
                mappings: [
                  ...step.mappings,
                  { from: fields[0] ?? "", to: "Renamed" },
                ],
              }),
            children: i.jsx(LocalizedText, { children: "Add rename" }),
          }),
        ],
      });
    case "filter":
      return i.jsxs(i.Fragment, {
        children: [
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Combine conditions" }),
              i.jsxs("select", {
                "aria-label": text("Filter condition logic"),
                value: step.join,
                onChange: (event) => updateStep({ join: event.target.value }),
                lang: language("Filter condition logic"),
                "data-ui-label": "Filter condition logic",
                children: [
                  i.jsx("option", {
                    value: "and",
                    children: "All conditions (AND)",
                  }),
                  i.jsx("option", {
                    value: "or",
                    children: "Any condition (OR)",
                  }),
                ],
              }),
            ],
          }),
          step.conditions.map((condition, conditionIndex) =>
            i.jsxs(
              "fieldset",
              {
                children: [
                  i.jsxs("legend", {
                    children: [
                      i.jsx(LocalizedText, { children: "Condition " }),
                      conditionIndex + 1,
                    ],
                  }),
                  renderFieldSelect(
                    `Filter field ${conditionIndex + 1}`,
                    condition.field,
                    fields,
                    (field) =>
                      updateStep({
                        conditions: step.conditions.map(
                          (condition, fieldEntryIndex) =>
                            conditionIndex === fieldEntryIndex
                              ? { ...condition, field: field }
                              : condition,
                        ),
                      }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Comparison" }),
                      i.jsx("select", {
                        "aria-label": `Filter operator ${conditionIndex + 1}`,
                        value: condition.operator,
                        onChange: (event) =>
                          updateStep({
                            conditions: step.conditions.map(
                              (condition, operatorEntryIndex) =>
                                conditionIndex === operatorEntryIndex
                                  ? {
                                      ...condition,
                                      operator: event.target.value,
                                    }
                                  : condition,
                            ),
                          }),
                        children: [
                          "equals",
                          "notEquals",
                          "contains",
                          "startsWith",
                          "endsWith",
                          "gt",
                          "gte",
                          "lt",
                          "lte",
                          "blank",
                          "notBlank",
                        ].map((operatorOption) =>
                          i.jsx(
                            "option",
                            { children: operatorOption },
                            operatorOption,
                          ),
                        ),
                      }),
                    ],
                  }),
                  !["blank", "notBlank"].includes(condition.operator) &&
                    i.jsx(RecipeValueInput, {
                      label: `Filter value ${conditionIndex + 1}`,
                      value: condition.value,
                      onChange: (filterValue) =>
                        updateStep({
                          conditions: step.conditions.map(
                            (condition, valueEntryIndex) =>
                              conditionIndex === valueEntryIndex
                                ? { ...condition, value: filterValue }
                                : condition,
                          ),
                        }),
                    }),
                  i.jsxs("button", {
                    onClick: () =>
                      updateStep({
                        conditions: step.conditions.filter(
                          (condition, removeConditionIndex) =>
                            removeConditionIndex !== conditionIndex,
                        ),
                      }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove condition " }),
                      conditionIndex + 1,
                    ],
                  }),
                ],
              },
              conditionIndex,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              updateStep({
                conditions: [
                  ...step.conditions,
                  { field: fields[0] ?? "", operator: "equals", value: "" },
                ],
              }),
            children: i.jsx(LocalizedText, { children: "Add condition" }),
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Comparisons preserve types. Text is case-sensitive; numeric comparisons require numbers.",
            }),
          }),
        ],
      });
    case "types":
      return i.jsxs(i.Fragment, {
        children: [
          renderFieldChecklist("Convert columns", step.fields, (fields) =>
            updateStep({ fields: fields }),
          ),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Output type" }),
              i.jsx("select", {
                "aria-label": text("Conversion type"),
                value: step.to,
                onChange: (event) => updateStep({ to: event.target.value }),
                lang: language("Conversion type"),
                "data-ui-label": "Conversion type",
                children: ["text", "number", "integer", "boolean", "date"].map(
                  (typeOption) =>
                    i.jsx("option", { children: typeOption }, typeOption),
                ),
              }),
            ],
          }),
          i.jsxs("div", {
            className: "recipe-line",
            children: [
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Decimal separator" }),
                  i.jsxs("select", {
                    "aria-label": text("Recipe decimal separator"),
                    value: step.decimal,
                    onChange: (event) =>
                      updateStep({ decimal: event.target.value }),
                    lang: language("Recipe decimal separator"),
                    "data-ui-label": "Recipe decimal separator",
                    children: [
                      i.jsx("option", { children: "." }),
                      i.jsx("option", { children: "," }),
                    ],
                  }),
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Grouping separator" }),
                  i.jsxs("select", {
                    "aria-label": text("Recipe grouping separator"),
                    value: step.group,
                    onChange: (event) =>
                      updateStep({ group: event.target.value }),
                    lang: language("Recipe grouping separator"),
                    "data-ui-label": "Recipe grouping separator",
                    children: [
                      i.jsx("option", { value: "", children: "None" }),
                      [".", ",", " ", "'"].map((separatorOption) =>
                        i.jsx(
                          "option",
                          {
                            value: separatorOption,
                            children:
                              separatorOption === " "
                                ? "Space"
                                : separatorOption,
                          },
                          separatorOption,
                        ),
                      ),
                    ],
                  }),
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Date order" }),
                  i.jsxs("select", {
                    "aria-label": text("Recipe date order"),
                    value: step.dateOrder,
                    onChange: (event) =>
                      updateStep({ dateOrder: event.target.value }),
                    lang: language("Recipe date order"),
                    "data-ui-label": "Recipe date order",
                    children: [
                      i.jsx("option", { children: "MDY" }),
                      i.jsx("option", { children: "DMY" }),
                    ],
                  }),
                ],
              }),
            ],
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Invalid values stop the recipe. Date conversion uses the workbook date system.",
            }),
          }),
        ],
      });
    case "replace":
      return i.jsxs(i.Fragment, {
        children: [
          renderFieldChecklist("Replace in columns", step.fields, (fields) =>
            updateStep({ fields: fields }),
          ),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Replacement mode" }),
              i.jsxs("select", {
                "aria-label": text("Replacement mode"),
                value: step.mode,
                onChange: (event) => updateStep({ mode: event.target.value }),
                lang: language("Replacement mode"),
                "data-ui-label": "Replacement mode",
                children: [
                  i.jsx("option", {
                    value: "value",
                    children: "Exact typed value",
                  }),
                  i.jsx("option", {
                    value: "text",
                    children: "All literal text occurrences",
                  }),
                ],
              }),
            ],
          }),
          i.jsx(RecipeValueInput, {
            label: text("Find value"),
            value: step.find,
            onChange: (find) => updateStep({ find: find }),
          }),
          i.jsx(RecipeValueInput, {
            label: text("Replacement value"),
            value: step.replacement,
            onChange: (replacement) => updateStep({ replacement: replacement }),
          }),
        ],
      });
    case "derive":
      return i.jsxs(i.Fragment, {
        children: [
          renderTextInput("Derived column name", step.name, (derivedName) =>
            updateStep({ name: derivedName }),
          ),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Row expression" }),
              i.jsx("textarea", {
                "aria-label": text("Row expression"),
                rows: 4,
                value: step.expression,
                onChange: (event) =>
                  updateStep({ expression: event.target.value }),
                lang: language("Row expression"),
                "data-ui-label": "Row expression",
              }),
            ],
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Use [Amount] * 2 or IF([Quantity] = 0, 0, [Amount] / [Quantity]). Functions: ABS, ROUND, UPPER, LOWER, TRIM, LEN, CONCAT, COALESCE, IF, AND, OR, NOT. Text literals use double quotes.",
            }),
          }),
        ],
      });
    case "sort":
      return i.jsxs(i.Fragment, {
        children: [
          step.fields.map((field, sortIndex) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  renderFieldSelect(
                    `Sort field ${sortIndex + 1}`,
                    field.field,
                    fields,
                    (field) =>
                      updateStep({
                        fields: step.fields.map((sortField, sortEntryIndex) =>
                          sortEntryIndex === sortIndex
                            ? { ...sortField, field: field }
                            : sortField,
                        ),
                      }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Direction" }),
                      i.jsxs("select", {
                        "aria-label": `Sort direction ${sortIndex + 1}`,
                        value: field.direction,
                        onChange: (event) =>
                          updateStep({
                            fields: step.fields.map(
                              (field, sortDirectionIndex) =>
                                sortDirectionIndex === sortIndex
                                  ? { ...field, direction: event.target.value }
                                  : field,
                            ),
                          }),
                        children: [
                          i.jsx("option", {
                            value: "ascending",
                            children: "Ascending",
                          }),
                          i.jsx("option", {
                            value: "descending",
                            children: "Descending",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("button", {
                    onClick: () =>
                      updateStep({
                        fields: step.fields.filter(
                          (field, removeSortIndex) =>
                            removeSortIndex !== sortIndex,
                        ),
                      }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove sort " }),
                      sortIndex + 1,
                    ],
                  }),
                ],
              },
              sortIndex,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              updateStep({
                fields: [
                  ...step.fields,
                  { field: fields[0] ?? "", direction: "ascending" },
                ],
              }),
            children: i.jsx(LocalizedText, { children: "Add sort field" }),
          }),
        ],
      });
    case "unpivot":
      return i.jsxs(i.Fragment, {
        children: [
          renderFieldChecklist("Unpivot columns", step.fields, (fields) =>
            updateStep({ fields: fields }),
          ),
          renderTextInput(
            "Attribute column name",
            step.attribute,
            (attribute) => updateStep({ attribute: attribute }),
          ),
          renderTextInput("Value column name", step.value, (valueColumnName) =>
            updateStep({ value: valueColumnName }),
          ),
          i.jsxs("label", {
            className: "recipe-check",
            children: [
              i.jsx("input", {
                type: "checkbox",
                checked: step.keepNulls,
                onChange: (event) =>
                  updateStep({ keepNulls: event.target.checked }),
              }),
              i.jsx(LocalizedText, { children: "Retain blank values" }),
            ],
          }),
        ],
      });
    case "group":
      return i.jsxs(i.Fragment, {
        children: [
          renderFieldChecklist("Group keys", step.keys, (keys) =>
            updateStep({ keys: keys }),
          ),
          step.values.map((value, aggregateIndex) =>
            i.jsxs(
              "fieldset",
              {
                children: [
                  i.jsxs("legend", {
                    children: [
                      i.jsx(LocalizedText, { children: "Aggregate " }),
                      aggregateIndex + 1,
                    ],
                  }),
                  renderFieldSelect(
                    `Aggregate field ${aggregateIndex + 1}`,
                    value.field,
                    fields,
                    (field) =>
                      updateStep({
                        values: step.values.map(
                          (value, aggregateFieldEntryIndex) =>
                            aggregateFieldEntryIndex === aggregateIndex
                              ? { ...value, field: field }
                              : value,
                        ),
                      }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Operation" }),
                      i.jsx("select", {
                        "aria-label": `Aggregate operation ${aggregateIndex + 1}`,
                        value: value.operation,
                        onChange: (event) =>
                          updateStep({
                            values: step.values.map(
                              (value, aggregateOperationEntryIndex) =>
                                aggregateOperationEntryIndex === aggregateIndex
                                  ? { ...value, operation: event.target.value }
                                  : value,
                            ),
                          }),
                        children: recipeAggregateOperations.map(
                          (recipeAggregateOperation) =>
                            i.jsx(
                              "option",
                              { children: recipeAggregateOperation },
                              recipeAggregateOperation,
                            ),
                        ),
                      }),
                    ],
                  }),
                  renderTextInput(
                    `Aggregate name ${aggregateIndex + 1}`,
                    value.name,
                    (aggregateName) =>
                      updateStep({
                        values: step.values.map(
                          (value, aggregateNameEntryIndex) =>
                            aggregateNameEntryIndex === aggregateIndex
                              ? { ...value, name: aggregateName }
                              : value,
                        ),
                      }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      updateStep({
                        values: step.values.filter(
                          (value, removeAggregateIndex) =>
                            removeAggregateIndex !== aggregateIndex,
                        ),
                      }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove aggregate " }),
                      aggregateIndex + 1,
                    ],
                  }),
                ],
              },
              aggregateIndex,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              updateStep({
                values: [
                  ...step.values,
                  {
                    field: fields.at(-1) ?? "",
                    operation: "sum",
                    name: `Total ${step.values.length + 1}`,
                  },
                ],
              }),
            children: i.jsx(LocalizedText, { children: "Add aggregate" }),
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Count counts nonblank values; numeric aggregates require numeric fields. No keys groups all records.",
            }),
          }),
        ],
      });
    case "pivot":
      return i.jsxs(i.Fragment, {
        children: [
          renderFieldChecklist("Pivot grouping keys", step.keys, (keys) =>
            updateStep({ keys: keys }),
          ),
          renderFieldSelect(
            "Pivot attribute field",
            step.column,
            fields,
            (column) => updateStep({ column: column }),
          ),
          renderFieldSelect(
            "Pivot value field",
            step.value,
            fields,
            (pivotValueField) => updateStep({ value: pivotValueField }),
          ),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Aggregation" }),
              i.jsx("select", {
                "aria-label": text("Pivot aggregation"),
                value: step.operation,
                onChange: (event) =>
                  updateStep({ operation: event.target.value }),
                lang: language("Pivot aggregation"),
                "data-ui-label": "Pivot aggregation",
                children: recipeAggregateOperations.map(
                  (recipeAggregateOperation) =>
                    i.jsx(
                      "option",
                      { children: recipeAggregateOperation },
                      recipeAggregateOperation,
                    ),
                ),
              }),
            ],
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "New headers include each attribute's type and value, in first-appearance order.",
            }),
          }),
        ],
      });
    case "append":
      return i.jsxs(i.Fragment, {
        children: [
          i.jsx(RecipeSourceSelector, {
            label: text("Append source"),
            source: step.source,
            sheets,
            onChange: (source) => updateStep({ source: source }),
          }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Headers must match exactly. Column order is aligned; no duplicate rows are removed.",
            }),
          }),
        ],
      });
    case "merge":
      return i.jsxs(i.Fragment, {
        children: [
          i.jsx(RecipeSourceSelector, {
            label: text("Merge source"),
            source: step.source,
            sheets,
            onChange: (source) => updateStep({ source: source }),
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Join mode" }),
              i.jsx("select", {
                "aria-label": text("Recipe join mode"),
                value: step.mode,
                onChange: (event) => updateStep({ mode: event.target.value }),
                lang: language("Recipe join mode"),
                "data-ui-label": "Recipe join mode",
                children: [
                  "inner",
                  "left",
                  "right",
                  "full",
                  "leftAnti",
                  "rightAnti",
                ].map((joinModeOption) =>
                  i.jsx("option", { children: joinModeOption }, joinModeOption),
                ),
              }),
            ],
          }),
          renderTextInput("Right column prefix", step.prefix, (prefix) =>
            updateStep({ prefix: prefix }),
          ),
          step.keys.map((key, keyIndex) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  renderFieldSelect(
                    `Recipe left key ${keyIndex + 1}`,
                    key.left,
                    fields,
                    (left) =>
                      updateStep({
                        keys: step.keys.map((key, keyEntryIndex) =>
                          keyIndex === keyEntryIndex
                            ? { ...key, left: left }
                            : key,
                        ),
                      }),
                  ),
                  renderFieldSelect(
                    `Recipe right key ${keyIndex + 1}`,
                    key.right,
                    rightFields,
                    (right) =>
                      updateStep({
                        keys: step.keys.map((key, rightKeyEntryIndex) =>
                          keyIndex === rightKeyEntryIndex
                            ? { ...key, right: right }
                            : key,
                        ),
                      }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      updateStep({
                        keys: step.keys.filter(
                          (key, removeKeyIndex) => removeKeyIndex !== keyIndex,
                        ),
                      }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove key " }),
                      keyIndex + 1,
                    ],
                  }),
                ],
              },
              keyIndex,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              updateStep({
                keys: [
                  ...step.keys,
                  { left: fields[0] ?? "", right: rightFields[0] ?? "" },
                ],
              }),
            children: i.jsx(LocalizedText, { children: "Add merge key" }),
          }),
          [
            ["trim", "Trim keys"],
            ["collapseWhitespace", "Collapse key whitespace"],
            ["ignoreCase", "Ignore key case"],
          ].map(([normalizationKey, normalizationLabel]) =>
            i.jsxs(
              "label",
              {
                className: "recipe-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: !!step.normalization?.[normalizationKey],
                    onChange: (event) =>
                      updateStep({
                        normalization: {
                          ...step.normalization,
                          [normalizationKey]: event.target.checked,
                        },
                      }),
                  }),
                  normalizationLabel,
                ],
              },
              normalizationKey,
            ),
          ),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Unicode normalization" }),
              i.jsxs("select", {
                "aria-label": text("Recipe Unicode normalization"),
                value: step.normalization?.unicode ?? "",
                onChange: (event) =>
                  updateStep({
                    normalization: {
                      ...step.normalization,
                      unicode: event.target.value || undefined,
                    },
                  }),
                lang: language("Recipe Unicode normalization"),
                "data-ui-label": "Recipe Unicode normalization",
                children: [
                  i.jsx("option", { value: "", children: "None" }),
                  i.jsx("option", { children: "NFC" }),
                  i.jsx("option", { children: "NFKC" }),
                ],
              }),
            ],
          }),
          i.jsxs("label", {
            className: "recipe-check",
            children: [
              i.jsx("input", {
                type: "checkbox",
                checked: !!step.fuzzy,
                onChange: (event) =>
                  updateStep({
                    fuzzy: event.target.checked
                      ? { key: 0, threshold: 0.8 }
                      : undefined,
                  }),
              }),
              i.jsx(LocalizedText, { children: "Fuzzy text key" }),
            ],
          }),
          step.fuzzy &&
            i.jsxs("div", {
              className: "recipe-line",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Key pair" }),
                    i.jsx("select", {
                      "aria-label": text("Recipe fuzzy key"),
                      value: step.fuzzy.key,
                      onChange: (event) =>
                        updateStep({
                          fuzzy: {
                            ...step.fuzzy,
                            key: Number(event.target.value),
                          },
                        }),
                      lang: language("Recipe fuzzy key"),
                      "data-ui-label": "Recipe fuzzy key",
                      children: step.keys.map((key, fuzzyKeyIndex) =>
                        i.jsx(
                          "option",
                          { value: fuzzyKeyIndex, children: fuzzyKeyIndex + 1 },
                          fuzzyKeyIndex,
                        ),
                      ),
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Minimum similarity" }),
                    i.jsx(LocalizedNumberField, {
                      "aria-label": text("Recipe fuzzy threshold"),
                      min: "0",
                      max: "1",
                      step: "0.01",
                      value: step.fuzzy.threshold,
                      onValueChange: (thresholdText) =>
                        updateStep({
                          fuzzy: {
                            ...step.fuzzy,
                            threshold:
                              thresholdText === ""
                                ? NaN
                                : Number(thresholdText),
                          },
                        }),
                      lang: language("Recipe fuzzy threshold"),
                      "data-ui-label": "Recipe fuzzy threshold",
                    }),
                  ],
                }),
              ],
            }),
          i.jsx("p", {
            children: i.jsx(LocalizedText, {
              children:
                "Every qualifying match is retained. Review unmatched records and many-to-many expansion in the step report.",
            }),
          }),
        ],
      });
  }
}
function createRecipeStep(stepKind, fields, source) {
  const stepId = crypto.randomUUID(),
    firstField = fields[0] ?? "",
    field = fields.at(-1) ?? "";
  switch (stepKind) {
    case "select":
      return { id: stepId, kind: stepKind, fields: [...fields] };
    case "rename":
      return {
        id: stepId,
        kind: stepKind,
        mappings: [{ from: firstField, to: "Renamed" }],
      };
    case "filter":
      return {
        id: stepId,
        kind: stepKind,
        join: "and",
        conditions: [{ field: firstField, operator: "notBlank" }],
      };
    case "types":
      return {
        id: stepId,
        kind: stepKind,
        fields: [field],
        to: "number",
        decimal: ".",
        group: ",",
        dateOrder: "MDY",
      };
    case "replace":
      return {
        id: stepId,
        kind: stepKind,
        fields: [firstField],
        mode: "value",
        find: "",
        replacement: "",
      };
    case "derive":
      return {
        id: stepId,
        kind: stepKind,
        name: "Calculated",
        expression: `[${field.replaceAll("]", "]]")}] * 2`,
      };
    case "sort":
      return {
        id: stepId,
        kind: stepKind,
        fields: [{ field: firstField, direction: "ascending" }],
      };
    case "unpivot":
      return {
        id: stepId,
        kind: stepKind,
        fields: [field],
        attribute: "Attribute",
        value: "Value",
        keepNulls: true,
      };
    case "group":
      return {
        id: stepId,
        kind: stepKind,
        keys: firstField === field ? [] : [firstField],
        values: [{ field: field, operation: "sum", name: "Total" }],
      };
    case "pivot":
      return {
        id: stepId,
        kind: stepKind,
        keys: [],
        column: firstField,
        value: field,
        operation: "sum",
      };
    case "append":
      return { id: stepId, kind: stepKind, source: { ...source } };
    case "merge":
      return {
        id: stepId,
        kind: stepKind,
        source: { ...source },
        mode: "left",
        keys: [{ left: firstField, right: firstField }],
        prefix: "Ref.",
      };
  }
}
const formatCellValueForDisplay = (cellValue) =>
  cellValue === null
    ? "(blank)"
    : typeof cellValue == "boolean"
      ? String(cellValue).toUpperCase()
      : String(cellValue);

export {
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
};
