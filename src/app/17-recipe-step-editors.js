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

function runTransformationRecipe(workbook, definition, t = {}) {
  const transformationRecipe = parseTransformationRecipe(definition),
    sources = [],
    steps = [],
    s = new Map(),
    d = (source) => {
      const h = canonicalJsonStringify(source);
      let f = s.get(h);
      if (!f) {
        const recipeSourceTable = resolveRecipeSourceTable(workbook, source, t);
        ((f = recipeSourceTable.data),
          sources.push(recipeSourceTable.evidence),
          s.set(h, f));
      }
      return structuredClone(f);
    },
    u = () => ({
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
    output = d(transformationRecipe.source);
  } catch (error) {
    return {
      ...u(),
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
  for (const [c, h] of transformationRecipe.steps.entries()) {
    const f = {
      index: c + 1,
      id: h.id,
      kind: h.kind,
      status: "failed",
      inputSchema: profileFieldTypes(output),
      inputRows: output.rows.length,
      inputHash: fingerprintCanonicalValue(output),
    };
    try {
      const p = executeRecipeStep(workbook, output, h, d);
      output = validateRecipeOutputTable(p.data);
      const fieldControlTotals = buildFieldControlTotals(output);
      (Object.assign(f, {
        status: "complete",
        outputSchema: profileFieldTypes(output),
        outputRows: output.rows.length,
        outputHash: fingerprintCanonicalValue(output),
        controls: fieldControlTotals,
        ...(p.details ? { details: p.details } : {}),
      }),
        steps.push(f));
    } catch (error) {
      return (
        (f.error = error.message),
        steps.push(
          f,
          ...transformationRecipe.steps.slice(c + 1).map((m, v) => ({
            index: c + v + 2,
            id: m.id,
            kind: m.kind,
            status: "skipped",
            inputSchema: [],
            inputRows: 0,
          })),
        ),
        {
          ...u(),
          status: "failed",
          failure: {
            index: c + 1,
            kind: h.kind,
            message: f.error,
            inputSchema: f.inputSchema,
            inputRows: f.inputRows,
          },
        }
      );
    }
  }
  return {
    ...u(),
    status: "complete",
    output: output,
    outputHash: fingerprintCanonicalValue(output),
    controls: buildFieldControlTotals(output),
  };
}
function kI(workbook, r = false) {
  if (r) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(workbook);
}
function saveTransformationRecipe(workbook, r, definition, n, o = {}) {
  kI(workbook, !!o.readOnly);
  const transformationRecipe = parseTransformationRecipe(definition);
  if (((r = r.trim()), !r || r.length > 128))
    throw Error("Use a recipe name of 1–128 characters.");
  const transformationRecipes = structuredClone(
      workbook.transformationRecipes ?? [],
    ),
    d = n
      ? transformationRecipes.find(
          (transformationRecipe) => transformationRecipe.id === n,
        )
      : undefined;
  if (n && !d) throw Error("Saved recipe no longer exists.");
  if (
    transformationRecipes.some(
      (transformationRecipe) =>
        transformationRecipe.id !== n &&
        transformationRecipe.name.toLowerCase() === r.toLowerCase(),
    )
  )
    throw Error("A recipe with this name already exists.");
  const u = {
    version: (d?.versions.length ?? 0) + 1,
    definition: transformationRecipe,
    definitionHash: fingerprintCanonicalValue(transformationRecipe),
    createdAt: new Date().toISOString(),
  };
  d
    ? ((d.name = r), d.versions.push(u))
    : transformationRecipes.push({
        id: crypto.randomUUID(),
        name: r,
        versions: [u],
      });
  const l = {
    ...workbook,
    sheets: workbook.sheets.map((sheet) => ({ ...sheet })),
    transformationRecipes: transformationRecipes,
  };
  return (validateSavedRecipeMetadata(l, l), l);
}
function deleteTransformationRecipe(workbook, r, t = {}) {
  if (
    (kI(workbook, !!t.readOnly),
    !workbook.transformationRecipes?.some(
      (transformationRecipe) => transformationRecipe.id === r,
    ))
  )
    throw Error("Saved recipe no longer exists.");
  return {
    ...workbook,
    transformationRecipes: workbook.transformationRecipes.filter(
      (transformationRecipe) => transformationRecipe.id !== r,
    ),
  };
}
function normalizeRecipeRunRequest(workbook, request, t = {}) {
  const transformationRecipe = parseTransformationRecipe(request.definition);
  if (
    Object.keys(request).some(
      (d) => !["definition", "recipeId", "version", "outputName"].includes(d),
    )
  )
    throw Error("Unknown recipe request field.");
  if ((request.recipeId === undefined) != (request.version === undefined))
    throw Error("Choose a saved recipe ID and version together.");
  if (request.recipeId) {
    const d = workbook.transformationRecipes
      ?.find(
        (transformationRecipe) => transformationRecipe.id === request.recipeId,
      )
      ?.versions.find((version) => version.version === request.version);
    if (
      !d ||
      d.definitionHash !== fingerprintCanonicalValue(transformationRecipe)
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
    includeHidden: !!t.includeHidden,
    allowedSheetIds: [
      ...new Set(t.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id)),
    ].sort(),
  };
  return {
    snapshot: JSON.stringify(workbook),
    request: { ...request, definition: transformationRecipe },
    context: context,
    run: runTransformationRecipe(workbook, transformationRecipe, t),
  };
}
function publishRecipeOutput(workbook, plan, t = {}) {
  if ((kI(workbook, !!t.readOnly), plan.snapshot !== JSON.stringify(workbook)))
    throw Error("Workbook changed after preview. Preview the recipe again.");
  if (!plan.request.recipeId || !plan.request.version)
    throw Error("Save a recipe version before publishing its output.");
  const recipeRunRequest = normalizeRecipeRunRequest(workbook, plan.request, t);
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
    workbook.sheets.reduce((c, h) => c + Object.keys(h.cells).length, 0) +
      (output.rows.length + 1) * output.fields.length >
    workbookLimits.maxCells
  )
    throw Error("Recipe output exceeds workbook cell capacity.");
  const worksheetRecord = createWorksheetRecord(outputName);
  ((worksheetRecord.id = `recipe-${fingerprintCanonicalValue([workbook.id, workbook.revision, outputName, recipeRunRequest.run.sourceHash, recipeRunRequest.run.definitionHash]).slice(0, 24)}`),
    output.fields.forEach(
      (field, h) =>
        (worksheetRecord.cells[`${columnIndexToLetters(h)}1`] = {
          raw: field,
          ...(field.startsWith("=") ? { literal: true } : {}),
          style: { bold: true, fill: "#eee9fa" },
        }),
    ));
  for (const [c, h] of output.rows.entries()) {
    const note = `Recipe ${recipeRunRequest.request.recipeId} version ${recipeRunRequest.request.version}
Original source rows: ${h.sources.join("; ")}`;
    if (note.length > 32767)
      throw Error("Recipe source notes exceed the cell note budget.");
    for (const [p, raw] of h.values.entries())
      worksheetRecord.cells[`${columnIndexToLetters(p)}${c + 2}`] = {
        raw: raw,
        ...(typeof raw == "string" && raw.startsWith("=")
          ? { literal: true }
          : {}),
        note: note,
        ...(output.declaredTypes[output.fields[p]] === "date"
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
      output.rows.map((row, h) => [h + 2, row.sources]),
    ),
  };
  const l = { ...workbook, sheets: [...workbook.sheets, worksheetRecord] };
  return normalizeWorkbookInput(l);
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
                          : selected.filter((s) => s !== field),
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
  const n =
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
            value: n,
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
      n !== "blank" &&
        (n === "boolean"
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
                  type: n === "number" ? "number" : "text",
                  value:
                    typeof value == "number" && !Number.isFinite(value)
                      ? ""
                      : String(value ?? ""),
                  onChange: (event) =>
                    onChange(
                      n === "number"
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
    d = (f) => onChange({ ...step, ...f }),
    u = (f, p, m, v) =>
      i.jsxs("label", {
        children: [
          f,
          i.jsxs("select", {
            "aria-label": f,
            value: p,
            onChange: (event) => v(event.target.value),
            children: [
              i.jsx("option", { value: "", children: "Choose field" }),
              m.map((g) => i.jsx("option", { value: g, children: g }, g)),
            ],
          }),
        ],
      }),
    l = (f, p, m) =>
      i.jsxs("label", {
        children: [
          f,
          i.jsx("input", {
            "aria-label": f,
            value: p,
            onChange: (event) => m(event.target.value),
          }),
        ],
      }),
    c = (f, selected, onChange) =>
      i.jsx(RecipeFieldChecklist, {
        label: f,
        fields,
        selected: selected,
        onChange: onChange,
      });
  switch (step.kind) {
    case "select":
      return c("Keep columns", step.fields, (fields) => d({ fields: fields }));
    case "rename":
      return i.jsxs(i.Fragment, {
        children: [
          step.mappings.map((mapping, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Rename source ${p + 1}`, mapping.from, fields, (from) =>
                    d({
                      mappings: step.mappings.map((mapping, g) =>
                        p === g ? { ...mapping, from: from } : mapping,
                      ),
                    }),
                  ),
                  l(`Rename to ${p + 1}`, mapping.to, (m) =>
                    d({
                      mappings: step.mappings.map((mapping, g) =>
                        p === g ? { ...mapping, to: m } : mapping,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({
                        mappings: step.mappings.filter((mapping, v) => p !== v),
                      }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove rename " }),
                      p + 1,
                    ],
                  }),
                ],
              },
              p,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              d({
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
                onChange: (event) => d({ join: event.target.value }),
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
          step.conditions.map((condition, p) =>
            i.jsxs(
              "fieldset",
              {
                children: [
                  i.jsxs("legend", {
                    children: [
                      i.jsx(LocalizedText, { children: "Condition " }),
                      p + 1,
                    ],
                  }),
                  u(`Filter field ${p + 1}`, condition.field, fields, (field) =>
                    d({
                      conditions: step.conditions.map((condition, g) =>
                        p === g ? { ...condition, field: field } : condition,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Comparison" }),
                      i.jsx("select", {
                        "aria-label": `Filter operator ${p + 1}`,
                        value: condition.operator,
                        onChange: (event) =>
                          d({
                            conditions: step.conditions.map((condition, g) =>
                              p === g
                                ? { ...condition, operator: event.target.value }
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
                        ].map((m) => i.jsx("option", { children: m }, m)),
                      }),
                    ],
                  }),
                  !["blank", "notBlank"].includes(condition.operator) &&
                    i.jsx(RecipeValueInput, {
                      label: `Filter value ${p + 1}`,
                      value: condition.value,
                      onChange: (m) =>
                        d({
                          conditions: step.conditions.map((condition, g) =>
                            p === g ? { ...condition, value: m } : condition,
                          ),
                        }),
                    }),
                  i.jsxs("button", {
                    onClick: () =>
                      d({
                        conditions: step.conditions.filter(
                          (condition, v) => v !== p,
                        ),
                      }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove condition " }),
                      p + 1,
                    ],
                  }),
                ],
              },
              p,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              d({
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
          c("Convert columns", step.fields, (fields) => d({ fields: fields })),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Output type" }),
              i.jsx("select", {
                "aria-label": text("Conversion type"),
                value: step.to,
                onChange: (event) => d({ to: event.target.value }),
                lang: language("Conversion type"),
                "data-ui-label": "Conversion type",
                children: ["text", "number", "integer", "boolean", "date"].map(
                  (f) => i.jsx("option", { children: f }, f),
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
                    onChange: (event) => d({ decimal: event.target.value }),
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
                    onChange: (event) => d({ group: event.target.value }),
                    lang: language("Recipe grouping separator"),
                    "data-ui-label": "Recipe grouping separator",
                    children: [
                      i.jsx("option", { value: "", children: "None" }),
                      [".", ",", " ", "'"].map((f) =>
                        i.jsx(
                          "option",
                          { value: f, children: f === " " ? "Space" : f },
                          f,
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
                    onChange: (event) => d({ dateOrder: event.target.value }),
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
          c("Replace in columns", step.fields, (fields) =>
            d({ fields: fields }),
          ),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Replacement mode" }),
              i.jsxs("select", {
                "aria-label": text("Replacement mode"),
                value: step.mode,
                onChange: (event) => d({ mode: event.target.value }),
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
            onChange: (find) => d({ find: find }),
          }),
          i.jsx(RecipeValueInput, {
            label: text("Replacement value"),
            value: step.replacement,
            onChange: (replacement) => d({ replacement: replacement }),
          }),
        ],
      });
    case "derive":
      return i.jsxs(i.Fragment, {
        children: [
          l("Derived column name", step.name, (f) => d({ name: f })),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Row expression" }),
              i.jsx("textarea", {
                "aria-label": text("Row expression"),
                rows: 4,
                value: step.expression,
                onChange: (event) => d({ expression: event.target.value }),
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
          step.fields.map((field, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Sort field ${p + 1}`, field.field, fields, (field) =>
                    d({
                      fields: step.fields.map((v, g) =>
                        g === p ? { ...v, field: field } : v,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Direction" }),
                      i.jsxs("select", {
                        "aria-label": `Sort direction ${p + 1}`,
                        value: field.direction,
                        onChange: (event) =>
                          d({
                            fields: step.fields.map((field, g) =>
                              g === p
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
                      d({ fields: step.fields.filter((field, v) => v !== p) }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove sort " }),
                      p + 1,
                    ],
                  }),
                ],
              },
              p,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              d({
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
          c("Unpivot columns", step.fields, (fields) => d({ fields: fields })),
          l("Attribute column name", step.attribute, (attribute) =>
            d({ attribute: attribute }),
          ),
          l("Value column name", step.value, (f) => d({ value: f })),
          i.jsxs("label", {
            className: "recipe-check",
            children: [
              i.jsx("input", {
                type: "checkbox",
                checked: step.keepNulls,
                onChange: (event) => d({ keepNulls: event.target.checked }),
              }),
              i.jsx(LocalizedText, { children: "Retain blank values" }),
            ],
          }),
        ],
      });
    case "group":
      return i.jsxs(i.Fragment, {
        children: [
          c("Group keys", step.keys, (keys) => d({ keys: keys })),
          step.values.map((value, p) =>
            i.jsxs(
              "fieldset",
              {
                children: [
                  i.jsxs("legend", {
                    children: [
                      i.jsx(LocalizedText, { children: "Aggregate " }),
                      p + 1,
                    ],
                  }),
                  u(`Aggregate field ${p + 1}`, value.field, fields, (field) =>
                    d({
                      values: step.values.map((value, g) =>
                        g === p ? { ...value, field: field } : value,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Operation" }),
                      i.jsx("select", {
                        "aria-label": `Aggregate operation ${p + 1}`,
                        value: value.operation,
                        onChange: (event) =>
                          d({
                            values: step.values.map((value, g) =>
                              g === p
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
                  l(`Aggregate name ${p + 1}`, value.name, (m) =>
                    d({
                      values: step.values.map((value, g) =>
                        g === p ? { ...value, name: m } : value,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ values: step.values.filter((value, v) => v !== p) }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove aggregate " }),
                      p + 1,
                    ],
                  }),
                ],
              },
              p,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              d({
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
          c("Pivot grouping keys", step.keys, (keys) => d({ keys: keys })),
          u("Pivot attribute field", step.column, fields, (column) =>
            d({ column: column }),
          ),
          u("Pivot value field", step.value, fields, (f) => d({ value: f })),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Aggregation" }),
              i.jsx("select", {
                "aria-label": text("Pivot aggregation"),
                value: step.operation,
                onChange: (event) => d({ operation: event.target.value }),
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
            onChange: (source) => d({ source: source }),
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
            onChange: (source) => d({ source: source }),
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Join mode" }),
              i.jsx("select", {
                "aria-label": text("Recipe join mode"),
                value: step.mode,
                onChange: (event) => d({ mode: event.target.value }),
                lang: language("Recipe join mode"),
                "data-ui-label": "Recipe join mode",
                children: [
                  "inner",
                  "left",
                  "right",
                  "full",
                  "leftAnti",
                  "rightAnti",
                ].map((f) => i.jsx("option", { children: f }, f)),
              }),
            ],
          }),
          l("Right column prefix", step.prefix, (prefix) =>
            d({ prefix: prefix }),
          ),
          step.keys.map((key, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Recipe left key ${p + 1}`, key.left, fields, (left) =>
                    d({
                      keys: step.keys.map((key, g) =>
                        p === g ? { ...key, left: left } : key,
                      ),
                    }),
                  ),
                  u(
                    `Recipe right key ${p + 1}`,
                    key.right,
                    rightFields,
                    (right) =>
                      d({
                        keys: step.keys.map((key, g) =>
                          p === g ? { ...key, right: right } : key,
                        ),
                      }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ keys: step.keys.filter((key, v) => v !== p) }),
                    children: [
                      i.jsx(LocalizedText, { children: "Remove key " }),
                      p + 1,
                    ],
                  }),
                ],
              },
              p,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              d({
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
          ].map(([f, p]) =>
            i.jsxs(
              "label",
              {
                className: "recipe-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: !!step.normalization?.[f],
                    onChange: (event) =>
                      d({
                        normalization: {
                          ...step.normalization,
                          [f]: event.target.checked,
                        },
                      }),
                  }),
                  p,
                ],
              },
              f,
            ),
          ),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Unicode normalization" }),
              i.jsxs("select", {
                "aria-label": text("Recipe Unicode normalization"),
                value: step.normalization?.unicode ?? "",
                onChange: (event) =>
                  d({
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
                  d({
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
                        d({
                          fuzzy: {
                            ...step.fuzzy,
                            key: Number(event.target.value),
                          },
                        }),
                      lang: language("Recipe fuzzy key"),
                      "data-ui-label": "Recipe fuzzy key",
                      children: step.keys.map((key, p) =>
                        i.jsx("option", { value: p, children: p + 1 }, p),
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
                      onValueChange: (f) =>
                        d({
                          fuzzy: {
                            ...step.fuzzy,
                            threshold: f === "" ? NaN : Number(f),
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
function createRecipeStep(e, fields, source) {
  const n = crypto.randomUUID(),
    o = fields[0] ?? "",
    field = fields.at(-1) ?? "";
  switch (e) {
    case "select":
      return { id: n, kind: e, fields: [...fields] };
    case "rename":
      return { id: n, kind: e, mappings: [{ from: o, to: "Renamed" }] };
    case "filter":
      return {
        id: n,
        kind: e,
        join: "and",
        conditions: [{ field: o, operator: "notBlank" }],
      };
    case "types":
      return {
        id: n,
        kind: e,
        fields: [field],
        to: "number",
        decimal: ".",
        group: ",",
        dateOrder: "MDY",
      };
    case "replace":
      return {
        id: n,
        kind: e,
        fields: [o],
        mode: "value",
        find: "",
        replacement: "",
      };
    case "derive":
      return {
        id: n,
        kind: e,
        name: "Calculated",
        expression: `[${field.replaceAll("]", "]]")}] * 2`,
      };
    case "sort":
      return { id: n, kind: e, fields: [{ field: o, direction: "ascending" }] };
    case "unpivot":
      return {
        id: n,
        kind: e,
        fields: [field],
        attribute: "Attribute",
        value: "Value",
        keepNulls: true,
      };
    case "group":
      return {
        id: n,
        kind: e,
        keys: o === field ? [] : [o],
        values: [{ field: field, operation: "sum", name: "Total" }],
      };
    case "pivot":
      return {
        id: n,
        kind: e,
        keys: [],
        column: o,
        value: field,
        operation: "sum",
      };
    case "append":
      return { id: n, kind: e, source: { ...source } };
    case "merge":
      return {
        id: n,
        kind: e,
        source: { ...source },
        mode: "left",
        keys: [{ left: o, right: o }],
        prefix: "Ref.",
      };
  }
}
const formatCellValueForDisplay = (e) =>
  e === null
    ? "(blank)"
    : typeof e == "boolean"
      ? String(e).toUpperCase()
      : String(e);

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
