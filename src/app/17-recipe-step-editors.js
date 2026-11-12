import i from "react/jsx-runtime";
import {
  parseTransformationRecipe,
  validateSavedRecipeMetadata,
  fingerprintCanonicalValue,
  canonicalJsonStringify,
  recipeStoreFormatId,
  normalizeDatePolicy,
} from "./01-core-schemas-and-hashing.js";
import { workbookLimits } from "./02-financial-functions.js";
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
import { inspectOutputNameSafety } from "./12-formula-dependency-inspector.js";
import {
  executeRecipeStep,
  resolveRecipeSourceTable,
  validateRecipeOutputTable,
  profileFieldTypes,
  buildFieldControlTotals,
} from "./16-text-split-tools.js";

function runTransformationRecipe(e, r, t = {}) {
  const n = parseTransformationRecipe(r),
    o = [],
    a = [],
    s = new Map(),
    d = (c) => {
      const h = canonicalJsonStringify(c);
      let f = s.get(h);
      if (!f) {
        const p = resolveRecipeSourceTable(e, c, t);
        ((f = p.data), o.push(p.evidence), s.set(h, f));
      }
      return structuredClone(f);
    },
    u = () => ({
      engine: recipeStoreFormatId,
      definitionHash: fingerprintCanonicalValue(n),
      sourceHash: fingerprintCanonicalValue({
        datePolicy: normalizeDatePolicy(e.datePolicy),
        sources: o,
      }),
      sources: o,
      steps: a,
    });
  let l;
  try {
    l = d(n.source);
  } catch (c) {
    return {
      ...u(),
      status: "failed",
      failure: {
        index: 0,
        kind: "source",
        message: c.message,
        inputSchema: [],
        inputRows: 0,
      },
    };
  }
  for (const [c, h] of n.steps.entries()) {
    const f = {
      index: c + 1,
      id: h.id,
      kind: h.kind,
      status: "failed",
      inputSchema: profileFieldTypes(l),
      inputRows: l.rows.length,
      inputHash: fingerprintCanonicalValue(l),
    };
    try {
      const p = executeRecipeStep(e, l, h, d);
      l = validateRecipeOutputTable(p.data);
      const m = buildFieldControlTotals(l);
      (Object.assign(f, {
        status: "complete",
        outputSchema: profileFieldTypes(l),
        outputRows: l.rows.length,
        outputHash: fingerprintCanonicalValue(l),
        controls: m,
        ...(p.details ? { details: p.details } : {}),
      }),
        a.push(f));
    } catch (p) {
      return (
        (f.error = p.message),
        a.push(
          f,
          ...n.steps.slice(c + 1).map((m, v) => ({
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
    output: l,
    outputHash: fingerprintCanonicalValue(l),
    controls: buildFieldControlTotals(l),
  };
}
function kI(e, r = false) {
  if (r) throw Error("This workbook is read-only.");
  assertWorkbookStructureEditable(e);
}
function saveTransformationRecipe(e, r, t, n, o = {}) {
  kI(e, !!o.readOnly);
  const a = parseTransformationRecipe(t);
  if (((r = r.trim()), !r || r.length > 128))
    throw Error("Use a recipe name of 1–128 characters.");
  const s = structuredClone(e.transformationRecipes ?? []),
    d = n ? s.find((c) => c.id === n) : undefined;
  if (n && !d) throw Error("Saved recipe no longer exists.");
  if (s.some((c) => c.id !== n && c.name.toLowerCase() === r.toLowerCase()))
    throw Error("A recipe with this name already exists.");
  const u = {
    version: (d?.versions.length ?? 0) + 1,
    definition: a,
    definitionHash: fingerprintCanonicalValue(a),
    createdAt: new Date().toISOString(),
  };
  d
    ? ((d.name = r), d.versions.push(u))
    : s.push({ id: crypto.randomUUID(), name: r, versions: [u] });
  const l = {
    ...e,
    sheets: e.sheets.map((c) => ({ ...c })),
    transformationRecipes: s,
  };
  return (validateSavedRecipeMetadata(l, l), l);
}
function deleteTransformationRecipe(e, r, t = {}) {
  var n;
  if (
    (kI(e, !!t.readOnly),
    !((n = e.transformationRecipes) != null && n.some((o) => o.id === r)))
  )
    throw Error("Saved recipe no longer exists.");
  return {
    ...e,
    transformationRecipes: e.transformationRecipes.filter((o) => o.id !== r),
  };
}
function normalizeRecipeRunRequest(e, r, t = {}) {
  const n = parseTransformationRecipe(r.definition);
  if (
    Object.keys(r).some(
      (d) => !["definition", "recipeId", "version", "outputName"].includes(d),
    )
  )
    throw Error("Unknown recipe request field.");
  if ((r.recipeId === undefined) != (r.version === undefined))
    throw Error("Choose a saved recipe ID and version together.");
  if (r.recipeId) {
    const d = e.transformationRecipes
      ?.find((u) => u.id === r.recipeId)
      ?.versions.find((u) => u.version === r.version);
    if (!d || d.definitionHash !== fingerprintCanonicalValue(n))
      throw Error(
        "The recipe differs from the saved version. Save a new version or reload the selected version.",
      );
  }
  if (
    typeof r.outputName != "string" ||
    !r.outputName.trim() ||
    r.outputName !== r.outputName.trim() ||
    r.outputName.length > 31 ||
    /[\\/\[\]*?:\u0000-\u001f]/.test(r.outputName) ||
    /^'|'$/.test(r.outputName)
  )
    throw Error("Choose an unpadded output worksheet name of 1–31 characters.");
  const o = {
    includeHidden: !!t.includeHidden,
    allowedSheetIds: [
      ...new Set(t.allowedSheetIds ?? e.sheets.map((d) => d.id)),
    ].sort(),
  };
  return {
    snapshot: JSON.stringify(e),
    request: { ...r, definition: n },
    context: o,
    run: runTransformationRecipe(e, n, t),
  };
}
function publishRecipeOutput(e, r, t = {}) {
  if ((kI(e, !!t.readOnly), r.snapshot !== JSON.stringify(e)))
    throw Error("Workbook changed after preview. Preview the recipe again.");
  if (!r.request.recipeId || !r.request.version)
    throw Error("Save a recipe version before publishing its output.");
  const n = normalizeRecipeRunRequest(e, r.request, t);
  if (canonicalJsonStringify(n) !== canonicalJsonStringify(r))
    throw Error("Recipe preview or authorization changed. Preview again.");
  if (n.run.status !== "complete")
    throw Error(
      "A failed recipe cannot publish output. Repair the failed step and preview again.",
    );
  const o = n.request.outputName;
  if (e.sheets.some((c) => c.name.toLowerCase() === o.toLowerCase()))
    throw Error(
      "Output worksheet already exists. Choose a new name; previous successful output is preserved.",
    );
  inspectOutputNameSafety(e, o);
  const a = n.run.output;
  if (
    e.sheets.reduce((c, h) => c + Object.keys(h.cells).length, 0) +
      (a.rows.length + 1) * a.fields.length >
    workbookLimits.maxCells
  )
    throw Error("Recipe output exceeds workbook cell capacity.");
  const d = createWorksheetRecord(o);
  ((d.id = `recipe-${fingerprintCanonicalValue([e.id, e.revision, o, n.run.sourceHash, n.run.definitionHash]).slice(0, 24)}`),
    a.fields.forEach(
      (c, h) =>
        (d.cells[`${columnIndexToLetters(h)}1`] = {
          raw: c,
          ...(c.startsWith("=") ? { literal: true } : {}),
          style: { bold: true, fill: "#eee9fa" },
        }),
    ));
  for (const [c, h] of a.rows.entries()) {
    const f = `Recipe ${n.request.recipeId} version ${n.request.version}
Original source rows: ${h.sources.join("; ")}`;
    if (f.length > 32767)
      throw Error("Recipe source notes exceed the cell note budget.");
    for (const [p, m] of h.values.entries())
      d.cells[`${columnIndexToLetters(p)}${c + 2}`] = {
        raw: m,
        ...(typeof m == "string" && m.startsWith("=") ? { literal: true } : {}),
        note: f,
        ...(a.declaredTypes[a.fields[p]] === "date" ? { format: "date" } : {}),
      };
  }
  const u = e.transformationRecipes.find((c) => c.id === n.request.recipeId);
  d.transformationRun = {
    engine: recipeStoreFormatId,
    recipeId: u.id,
    recipeName: u.name,
    version: n.request.version,
    definition: n.request.definition,
    definitionHash: n.run.definitionHash,
    sourceWorkbookId: e.id,
    sourceRevision: e.revision,
    sourceHash: n.run.sourceHash,
    outputHash: n.run.outputHash,
    outputRows: a.rows.length,
    outputRange: `A1:${columnIndexToLetters(a.fields.length - 1)}${a.rows.length + 1}`,
    at: new Date().toISOString(),
    steps: n.run.steps,
    controls: n.run.controls,
    lineage: Object.fromEntries(a.rows.map((c, h) => [h + 2, c.sources])),
  };
  const l = { ...e, sheets: [...e.sheets, d] };
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
        ? fields.map((o) =>
            i.jsxs(
              "label",
              {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    "aria-label": `${label}: ${o}`,
                    checked: selected.includes(o),
                    onChange: (a) =>
                      onChange(
                        a.target.checked
                          ? [...selected, o]
                          : selected.filter((s) => s !== o),
                      ),
                  }),
                  o,
                ],
              },
              o,
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
  const o = sheets.find((s) => s.id === source.sheetId);
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
            onChange: (s) => {
              const d = sheets.find((u) => u.id === s.target.value);
              onChange({
                sheetId: s.target.value,
                range: d ? formatOccupiedRangeAddress(d) : "A1:B10",
              });
            },
            children: [
              i.jsx("option", { value: "", children: "Choose worksheet" }),
              sheets.map((s) =>
                i.jsx("option", { value: s.id, children: s.name }, s.id),
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
            onChange: (s) =>
              onChange({ ...source, tableId: s.target.value || undefined }),
            children: [
              i.jsx("option", { value: "", children: "Explicit headed range" }),
              o?.tables?.map((s) =>
                i.jsxs(
                  "option",
                  {
                    value: s.id,
                    children: [s.name, " · current table body"],
                  },
                  s.id,
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
            onChange: (s) => onChange({ ...source, range: s.target.value }),
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
            onChange: (o) =>
              onChange(
                o.target.value === "blank"
                  ? null
                  : o.target.value === "number"
                    ? 0
                    : o.target.value === "boolean"
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
                  onChange: (o) => onChange(o.target.value === "true"),
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
                  onChange: (o) =>
                    onChange(
                      n === "number"
                        ? o.target.value === ""
                          ? NaN
                          : Number(o.target.value)
                        : o.target.value,
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
            onChange: (g) => v(g.target.value),
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
            onChange: (v) => m(v.target.value),
          }),
        ],
      }),
    c = (f, p, m) =>
      i.jsx(RecipeFieldChecklist, {
        label: f,
        fields,
        selected: p,
        onChange: m,
      });
  switch (step.kind) {
    case "select":
      return c("Keep columns", step.fields, (f) => d({ fields: f }));
    case "rename":
      return i.jsxs(i.Fragment, {
        children: [
          step.mappings.map((f, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Rename source ${p + 1}`, f.from, fields, (m) =>
                    d({
                      mappings: step.mappings.map((v, g) =>
                        p === g ? { ...v, from: m } : v,
                      ),
                    }),
                  ),
                  l(`Rename to ${p + 1}`, f.to, (m) =>
                    d({
                      mappings: step.mappings.map((v, g) =>
                        p === g ? { ...v, to: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ mappings: step.mappings.filter((m, v) => p !== v) }),
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
                onChange: (f) => d({ join: f.target.value }),
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
          step.conditions.map((f, p) =>
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
                  u(`Filter field ${p + 1}`, f.field, fields, (m) =>
                    d({
                      conditions: step.conditions.map((v, g) =>
                        p === g ? { ...v, field: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Comparison" }),
                      i.jsx("select", {
                        "aria-label": `Filter operator ${p + 1}`,
                        value: f.operator,
                        onChange: (m) =>
                          d({
                            conditions: step.conditions.map((v, g) =>
                              p === g ? { ...v, operator: m.target.value } : v,
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
                  !["blank", "notBlank"].includes(f.operator) &&
                    i.jsx(RecipeValueInput, {
                      label: `Filter value ${p + 1}`,
                      value: f.value,
                      onChange: (m) =>
                        d({
                          conditions: step.conditions.map((v, g) =>
                            p === g ? { ...v, value: m } : v,
                          ),
                        }),
                    }),
                  i.jsxs("button", {
                    onClick: () =>
                      d({
                        conditions: step.conditions.filter((m, v) => v !== p),
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
          c("Convert columns", step.fields, (f) => d({ fields: f })),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Output type" }),
              i.jsx("select", {
                "aria-label": text("Conversion type"),
                value: step.to,
                onChange: (f) => d({ to: f.target.value }),
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
                    onChange: (f) => d({ decimal: f.target.value }),
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
                    onChange: (f) => d({ group: f.target.value }),
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
                    onChange: (f) => d({ dateOrder: f.target.value }),
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
          c("Replace in columns", step.fields, (f) => d({ fields: f })),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Replacement mode" }),
              i.jsxs("select", {
                "aria-label": text("Replacement mode"),
                value: step.mode,
                onChange: (f) => d({ mode: f.target.value }),
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
            onChange: (f) => d({ find: f }),
          }),
          i.jsx(RecipeValueInput, {
            label: text("Replacement value"),
            value: step.replacement,
            onChange: (f) => d({ replacement: f }),
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
                onChange: (f) => d({ expression: f.target.value }),
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
          step.fields.map((f, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Sort field ${p + 1}`, f.field, fields, (m) =>
                    d({
                      fields: step.fields.map((v, g) =>
                        g === p ? { ...v, field: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Direction" }),
                      i.jsxs("select", {
                        "aria-label": `Sort direction ${p + 1}`,
                        value: f.direction,
                        onChange: (m) =>
                          d({
                            fields: step.fields.map((v, g) =>
                              g === p ? { ...v, direction: m.target.value } : v,
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
                      d({ fields: step.fields.filter((m, v) => v !== p) }),
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
          c("Unpivot columns", step.fields, (f) => d({ fields: f })),
          l("Attribute column name", step.attribute, (f) =>
            d({ attribute: f }),
          ),
          l("Value column name", step.value, (f) => d({ value: f })),
          i.jsxs("label", {
            className: "recipe-check",
            children: [
              i.jsx("input", {
                type: "checkbox",
                checked: step.keepNulls,
                onChange: (f) => d({ keepNulls: f.target.checked }),
              }),
              i.jsx(LocalizedText, { children: "Retain blank values" }),
            ],
          }),
        ],
      });
    case "group":
      return i.jsxs(i.Fragment, {
        children: [
          c("Group keys", step.keys, (f) => d({ keys: f })),
          step.values.map((f, p) =>
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
                  u(`Aggregate field ${p + 1}`, f.field, fields, (m) =>
                    d({
                      values: step.values.map((v, g) =>
                        g === p ? { ...v, field: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Operation" }),
                      i.jsx("select", {
                        "aria-label": `Aggregate operation ${p + 1}`,
                        value: f.operation,
                        onChange: (m) =>
                          d({
                            values: step.values.map((v, g) =>
                              g === p ? { ...v, operation: m.target.value } : v,
                            ),
                          }),
                        children: recipeAggregateOperations.map((m) =>
                          i.jsx("option", { children: m }, m),
                        ),
                      }),
                    ],
                  }),
                  l(`Aggregate name ${p + 1}`, f.name, (m) =>
                    d({
                      values: step.values.map((v, g) =>
                        g === p ? { ...v, name: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ values: step.values.filter((m, v) => v !== p) }),
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
          c("Pivot grouping keys", step.keys, (f) => d({ keys: f })),
          u("Pivot attribute field", step.column, fields, (f) =>
            d({ column: f }),
          ),
          u("Pivot value field", step.value, fields, (f) => d({ value: f })),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Aggregation" }),
              i.jsx("select", {
                "aria-label": text("Pivot aggregation"),
                value: step.operation,
                onChange: (f) => d({ operation: f.target.value }),
                lang: language("Pivot aggregation"),
                "data-ui-label": "Pivot aggregation",
                children: recipeAggregateOperations.map((f) =>
                  i.jsx("option", { children: f }, f),
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
            onChange: (f) => d({ source: f }),
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
            onChange: (f) => d({ source: f }),
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Join mode" }),
              i.jsx("select", {
                "aria-label": text("Recipe join mode"),
                value: step.mode,
                onChange: (f) => d({ mode: f.target.value }),
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
          l("Right column prefix", step.prefix, (f) => d({ prefix: f })),
          step.keys.map((f, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Recipe left key ${p + 1}`, f.left, fields, (m) =>
                    d({
                      keys: step.keys.map((v, g) =>
                        p === g ? { ...v, left: m } : v,
                      ),
                    }),
                  ),
                  u(`Recipe right key ${p + 1}`, f.right, rightFields, (m) =>
                    d({
                      keys: step.keys.map((v, g) =>
                        p === g ? { ...v, right: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ keys: step.keys.filter((m, v) => v !== p) }),
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
          ].map(([f, p]) => {
            var m;
            return i.jsxs(
              "label",
              {
                className: "recipe-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: !!((m = step.normalization) != null && m[f]),
                    onChange: (v) =>
                      d({
                        normalization: {
                          ...step.normalization,
                          [f]: v.target.checked,
                        },
                      }),
                  }),
                  p,
                ],
              },
              f,
            );
          }),
          i.jsxs("label", {
            children: [
              i.jsx(LocalizedText, { children: "Unicode normalization" }),
              i.jsxs("select", {
                "aria-label": text("Recipe Unicode normalization"),
                value: step.normalization?.unicode ?? "",
                onChange: (f) =>
                  d({
                    normalization: {
                      ...step.normalization,
                      unicode: f.target.value || undefined,
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
                onChange: (f) =>
                  d({
                    fuzzy: f.target.checked
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
                      onChange: (f) =>
                        d({
                          fuzzy: { ...step.fuzzy, key: Number(f.target.value) },
                        }),
                      lang: language("Recipe fuzzy key"),
                      "data-ui-label": "Recipe fuzzy key",
                      children: step.keys.map((f, p) =>
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
function createRecipeStep(e, r, t) {
  const n = crypto.randomUUID(),
    o = r[0] ?? "",
    a = r.at(-1) ?? "";
  switch (e) {
    case "select":
      return { id: n, kind: e, fields: [...r] };
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
        fields: [a],
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
        expression: `[${a.replaceAll("]", "]]")}] * 2`,
      };
    case "sort":
      return { id: n, kind: e, fields: [{ field: o, direction: "ascending" }] };
    case "unpivot":
      return {
        id: n,
        kind: e,
        fields: [a],
        attribute: "Attribute",
        value: "Value",
        keepNulls: true,
      };
    case "group":
      return {
        id: n,
        kind: e,
        keys: o === a ? [] : [o],
        values: [{ field: a, operation: "sum", name: "Total" }],
      };
    case "pivot":
      return {
        id: n,
        kind: e,
        keys: [],
        column: o,
        value: a,
        operation: "sum",
      };
    case "append":
      return { id: n, kind: e, source: { ...t } };
    case "merge":
      return {
        id: n,
        kind: e,
        source: { ...t },
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
