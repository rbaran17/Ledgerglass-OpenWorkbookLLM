import i from "react/jsx-runtime";
import { CN, GF, hr, mE, pg, ts } from "./part-01.js";
import { Mt } from "./part-02.js";
import { $a, Ku, Ta, rn, ut } from "./part-03.js";
import { Dr, M, ci } from "./part-04.js";
import { q1 } from "./part-12.js";
import { Bce, Fce, bz, wT, xT } from "./part-16.js";

function ET(e, r, t = {}) {
  const n = CN(r),
    o = [],
    a = [],
    s = new Map(),
    d = (c) => {
      const h = mE(c);
      let f = s.get(h);
      if (!f) {
        const p = Fce(e, c, t);
        ((f = p.data), o.push(p.evidence), s.set(h, f));
      }
      return structuredClone(f);
    },
    u = () => ({
      engine: pg,
      definitionHash: hr(n),
      sourceHash: hr({ datePolicy: ts(e.datePolicy), sources: o }),
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
      inputSchema: wT(l),
      inputRows: l.rows.length,
      inputHash: hr(l),
    };
    try {
      const p = Bce(e, l, h, d);
      l = bz(p.data);
      const m = xT(l);
      (Object.assign(f, {
        status: "complete",
        outputSchema: wT(l),
        outputRows: l.rows.length,
        outputHash: hr(l),
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
    outputHash: hr(l),
    controls: xT(l),
  };
}
function kI(e, r = false) {
  if (r) throw Error("This workbook is read-only.");
  $a(e);
}
function Lce(e, r, t, n, o = {}) {
  kI(e, !!o.readOnly);
  const a = CN(t);
  if (((r = r.trim()), !r || r.length > 128))
    throw Error("Use a recipe name of 1–128 characters.");
  const s = structuredClone(e.transformationRecipes ?? []),
    d = n ? s.find((c) => c.id === n) : void 0;
  if (n && !d) throw Error("Saved recipe no longer exists.");
  if (s.some((c) => c.id !== n && c.name.toLowerCase() === r.toLowerCase()))
    throw Error("A recipe with this name already exists.");
  const u = {
    version: ((d == null ? void 0 : d.versions.length) ?? 0) + 1,
    definition: a,
    definitionHash: hr(a),
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
  return (GF(l, l), l);
}
function Pce(e, r, t = {}) {
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
function yz(e, r, t = {}) {
  var a, s;
  const n = CN(r.definition);
  if (
    Object.keys(r).some(
      (d) => !["definition", "recipeId", "version", "outputName"].includes(d),
    )
  )
    throw Error("Unknown recipe request field.");
  if ((r.recipeId === void 0) != (r.version === void 0))
    throw Error("Choose a saved recipe ID and version together.");
  if (r.recipeId) {
    const d =
      (s =
        (a = e.transformationRecipes) == null
          ? void 0
          : a.find((u) => u.id === r.recipeId)) == null
        ? void 0
        : s.versions.find((u) => u.version === r.version);
    if (!d || d.definitionHash !== hr(n))
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
    run: ET(e, n, t),
  };
}
function Uce(e, r, t = {}) {
  if ((kI(e, !!t.readOnly), r.snapshot !== JSON.stringify(e)))
    throw Error("Workbook changed after preview. Preview the recipe again.");
  if (!r.request.recipeId || !r.request.version)
    throw Error("Save a recipe version before publishing its output.");
  const n = yz(e, r.request, t);
  if (mE(n) !== mE(r))
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
  q1(e, o);
  const a = n.run.output;
  if (
    e.sheets.reduce((c, h) => c + Object.keys(h.cells).length, 0) +
      (a.rows.length + 1) * a.fields.length >
    Mt.maxCells
  )
    throw Error("Recipe output exceeds workbook cell capacity.");
  const d = Ku(o);
  ((d.id = `recipe-${hr([e.id, e.revision, o, n.run.sourceHash, n.run.definitionHash]).slice(0, 24)}`),
    a.fields.forEach(
      (c, h) =>
        (d.cells[`${ut(h)}1`] = {
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
      d.cells[`${ut(p)}${c + 2}`] = {
        raw: m,
        ...(typeof m == "string" && m.startsWith("=") ? { literal: true } : {}),
        note: f,
        ...(a.declaredTypes[a.fields[p]] === "date" ? { format: "date" } : {}),
      };
  }
  const u = e.transformationRecipes.find((c) => c.id === n.request.recipeId);
  d.transformationRun = {
    engine: pg,
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
    outputRange: `A1:${ut(a.fields.length - 1)}${a.rows.length + 1}`,
    at: new Date().toISOString(),
    steps: n.run.steps,
    controls: n.run.controls,
    lineage: Object.fromEntries(a.rows.map((c, h) => [h + 2, c.sources])),
  };
  const l = { ...e, sheets: [...e.sheets, d] };
  return rn(l);
}
const Q4 = {
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
function $ce({ label: e, fields: r, selected: t, onChange: n }) {
  return i.jsxs("fieldset", {
    className: "recipe-fields",
    children: [
      i.jsx("legend", { children: e }),
      r.length
        ? r.map((o) =>
            i.jsxs(
              "label",
              {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    "aria-label": `${e}: ${o}`,
                    checked: t.includes(o),
                    onChange: (a) =>
                      n(
                        a.target.checked ? [...t, o] : t.filter((s) => s !== o),
                      ),
                  }),
                  o,
                ],
              },
              o,
            ),
          )
        : i.jsx("p", {
            children: i.jsx(M, {
              children: "Preview or repair the source to obtain fields.",
            }),
          }),
    ],
  });
}
function CT({ label: e, source: r, sheets: t, onChange: n }) {
  var a;
  const o = t.find((s) => s.id === r.sheetId);
  return i.jsxs("fieldset", {
    className: "recipe-source",
    children: [
      i.jsx("legend", { children: e }),
      i.jsxs("label", {
        children: [
          i.jsx(M, { children: "Worksheet" }),
          i.jsxs("select", {
            "aria-label": `${e} worksheet`,
            value: r.sheetId,
            onChange: (s) => {
              const d = t.find((u) => u.id === s.target.value);
              n({ sheetId: s.target.value, range: d ? Ta(d) : "A1:B10" });
            },
            children: [
              i.jsx("option", { value: "", children: "Choose worksheet" }),
              t.map((s) =>
                i.jsx("option", { value: s.id, children: s.name }, s.id),
              ),
            ],
          }),
        ],
      }),
      i.jsxs("label", {
        children: [
          i.jsx(M, { children: "Source type" }),
          i.jsxs("select", {
            "aria-label": `${e} table`,
            value: r.tableId ?? "",
            onChange: (s) => n({ ...r, tableId: s.target.value || void 0 }),
            children: [
              i.jsx("option", { value: "", children: "Explicit headed range" }),
              (a = o == null ? void 0 : o.tables) == null
                ? void 0
                : a.map((s) =>
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
          i.jsx(M, { children: "Headed range" }),
          i.jsx("input", {
            "aria-label": `${e} range`,
            value: r.range,
            disabled: !!r.tableId,
            onChange: (s) => n({ ...r, range: s.target.value }),
          }),
        ],
      }),
    ],
  });
}
function Fk({ label: e, value: r, onChange: t }) {
  const n =
    r === null
      ? "blank"
      : typeof r == "number"
        ? "number"
        : typeof r == "boolean"
          ? "boolean"
          : "text";
  return i.jsxs("span", {
    className: "recipe-typed",
    children: [
      i.jsxs("label", {
        children: [
          e,
          " ",
          i.jsx(M, { children: "type" }),
          i.jsxs("select", {
            "aria-label": `${e} type`,
            value: n,
            onChange: (o) =>
              t(
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
                e,
                i.jsxs("select", {
                  "aria-label": e,
                  value: String(r),
                  onChange: (o) => t(o.target.value === "true"),
                  children: [
                    i.jsx("option", { value: "true", children: "TRUE" }),
                    i.jsx("option", { value: "false", children: "FALSE" }),
                  ],
                }),
              ],
            })
          : i.jsxs("label", {
              children: [
                e,
                i.jsx("input", {
                  "aria-label": e,
                  type: n === "number" ? "number" : "text",
                  value:
                    typeof r == "number" && !Number.isFinite(r)
                      ? ""
                      : String(r ?? ""),
                  onChange: (o) =>
                    t(
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
const B7 = ["sum", "count", "average", "min", "max", "distinctCount"];
function zce({ step: e, fields: r, rightFields: t, sheets: n, onChange: o }) {
  var h;
  const { text: a, language: s } = Dr(),
    d = (f) => o({ ...e, ...f }),
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
      i.jsx($ce, { label: f, fields: r, selected: p, onChange: m });
  switch (e.kind) {
    case "select":
      return c("Keep columns", e.fields, (f) => d({ fields: f }));
    case "rename":
      return i.jsxs(i.Fragment, {
        children: [
          e.mappings.map((f, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Rename source ${p + 1}`, f.from, r, (m) =>
                    d({
                      mappings: e.mappings.map((v, g) =>
                        p === g ? { ...v, from: m } : v,
                      ),
                    }),
                  ),
                  l(`Rename to ${p + 1}`, f.to, (m) =>
                    d({
                      mappings: e.mappings.map((v, g) =>
                        p === g ? { ...v, to: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ mappings: e.mappings.filter((m, v) => p !== v) }),
                    children: [i.jsx(M, { children: "Remove rename " }), p + 1],
                  }),
                ],
              },
              p,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              d({
                mappings: [...e.mappings, { from: r[0] ?? "", to: "Renamed" }],
              }),
            children: i.jsx(M, { children: "Add rename" }),
          }),
        ],
      });
    case "filter":
      return i.jsxs(i.Fragment, {
        children: [
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Combine conditions" }),
              i.jsxs("select", {
                "aria-label": a("Filter condition logic"),
                value: e.join,
                onChange: (f) => d({ join: f.target.value }),
                lang: s("Filter condition logic"),
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
          e.conditions.map((f, p) =>
            i.jsxs(
              "fieldset",
              {
                children: [
                  i.jsxs("legend", {
                    children: [i.jsx(M, { children: "Condition " }), p + 1],
                  }),
                  u(`Filter field ${p + 1}`, f.field, r, (m) =>
                    d({
                      conditions: e.conditions.map((v, g) =>
                        p === g ? { ...v, field: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Comparison" }),
                      i.jsx("select", {
                        "aria-label": `Filter operator ${p + 1}`,
                        value: f.operator,
                        onChange: (m) =>
                          d({
                            conditions: e.conditions.map((v, g) =>
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
                    i.jsx(Fk, {
                      label: `Filter value ${p + 1}`,
                      value: f.value,
                      onChange: (m) =>
                        d({
                          conditions: e.conditions.map((v, g) =>
                            p === g ? { ...v, value: m } : v,
                          ),
                        }),
                    }),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ conditions: e.conditions.filter((m, v) => v !== p) }),
                    children: [
                      i.jsx(M, { children: "Remove condition " }),
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
                  ...e.conditions,
                  { field: r[0] ?? "", operator: "equals", value: "" },
                ],
              }),
            children: i.jsx(M, { children: "Add condition" }),
          }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Comparisons preserve types. Text is case-sensitive; numeric comparisons require numbers.",
            }),
          }),
        ],
      });
    case "types":
      return i.jsxs(i.Fragment, {
        children: [
          c("Convert columns", e.fields, (f) => d({ fields: f })),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Output type" }),
              i.jsx("select", {
                "aria-label": a("Conversion type"),
                value: e.to,
                onChange: (f) => d({ to: f.target.value }),
                lang: s("Conversion type"),
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
                  i.jsx(M, { children: "Decimal separator" }),
                  i.jsxs("select", {
                    "aria-label": a("Recipe decimal separator"),
                    value: e.decimal,
                    onChange: (f) => d({ decimal: f.target.value }),
                    lang: s("Recipe decimal separator"),
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
                  i.jsx(M, { children: "Grouping separator" }),
                  i.jsxs("select", {
                    "aria-label": a("Recipe grouping separator"),
                    value: e.group,
                    onChange: (f) => d({ group: f.target.value }),
                    lang: s("Recipe grouping separator"),
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
                  i.jsx(M, { children: "Date order" }),
                  i.jsxs("select", {
                    "aria-label": a("Recipe date order"),
                    value: e.dateOrder,
                    onChange: (f) => d({ dateOrder: f.target.value }),
                    lang: s("Recipe date order"),
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
            children: i.jsx(M, {
              children:
                "Invalid values stop the recipe. Date conversion uses the workbook date system.",
            }),
          }),
        ],
      });
    case "replace":
      return i.jsxs(i.Fragment, {
        children: [
          c("Replace in columns", e.fields, (f) => d({ fields: f })),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Replacement mode" }),
              i.jsxs("select", {
                "aria-label": a("Replacement mode"),
                value: e.mode,
                onChange: (f) => d({ mode: f.target.value }),
                lang: s("Replacement mode"),
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
          i.jsx(Fk, {
            label: a("Find value"),
            value: e.find,
            onChange: (f) => d({ find: f }),
          }),
          i.jsx(Fk, {
            label: a("Replacement value"),
            value: e.replacement,
            onChange: (f) => d({ replacement: f }),
          }),
        ],
      });
    case "derive":
      return i.jsxs(i.Fragment, {
        children: [
          l("Derived column name", e.name, (f) => d({ name: f })),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Row expression" }),
              i.jsx("textarea", {
                "aria-label": a("Row expression"),
                rows: 4,
                value: e.expression,
                onChange: (f) => d({ expression: f.target.value }),
                lang: s("Row expression"),
                "data-ui-label": "Row expression",
              }),
            ],
          }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Use [Amount] * 2 or IF([Quantity] = 0, 0, [Amount] / [Quantity]). Functions: ABS, ROUND, UPPER, LOWER, TRIM, LEN, CONCAT, COALESCE, IF, AND, OR, NOT. Text literals use double quotes.",
            }),
          }),
        ],
      });
    case "sort":
      return i.jsxs(i.Fragment, {
        children: [
          e.fields.map((f, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Sort field ${p + 1}`, f.field, r, (m) =>
                    d({
                      fields: e.fields.map((v, g) =>
                        g === p ? { ...v, field: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Direction" }),
                      i.jsxs("select", {
                        "aria-label": `Sort direction ${p + 1}`,
                        value: f.direction,
                        onChange: (m) =>
                          d({
                            fields: e.fields.map((v, g) =>
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
                      d({ fields: e.fields.filter((m, v) => v !== p) }),
                    children: [i.jsx(M, { children: "Remove sort " }), p + 1],
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
                  ...e.fields,
                  { field: r[0] ?? "", direction: "ascending" },
                ],
              }),
            children: i.jsx(M, { children: "Add sort field" }),
          }),
        ],
      });
    case "unpivot":
      return i.jsxs(i.Fragment, {
        children: [
          c("Unpivot columns", e.fields, (f) => d({ fields: f })),
          l("Attribute column name", e.attribute, (f) => d({ attribute: f })),
          l("Value column name", e.value, (f) => d({ value: f })),
          i.jsxs("label", {
            className: "recipe-check",
            children: [
              i.jsx("input", {
                type: "checkbox",
                checked: e.keepNulls,
                onChange: (f) => d({ keepNulls: f.target.checked }),
              }),
              i.jsx(M, { children: "Retain blank values" }),
            ],
          }),
        ],
      });
    case "group":
      return i.jsxs(i.Fragment, {
        children: [
          c("Group keys", e.keys, (f) => d({ keys: f })),
          e.values.map((f, p) =>
            i.jsxs(
              "fieldset",
              {
                children: [
                  i.jsxs("legend", {
                    children: [i.jsx(M, { children: "Aggregate " }), p + 1],
                  }),
                  u(`Aggregate field ${p + 1}`, f.field, r, (m) =>
                    d({
                      values: e.values.map((v, g) =>
                        g === p ? { ...v, field: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Operation" }),
                      i.jsx("select", {
                        "aria-label": `Aggregate operation ${p + 1}`,
                        value: f.operation,
                        onChange: (m) =>
                          d({
                            values: e.values.map((v, g) =>
                              g === p ? { ...v, operation: m.target.value } : v,
                            ),
                          }),
                        children: B7.map((m) =>
                          i.jsx("option", { children: m }, m),
                        ),
                      }),
                    ],
                  }),
                  l(`Aggregate name ${p + 1}`, f.name, (m) =>
                    d({
                      values: e.values.map((v, g) =>
                        g === p ? { ...v, name: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ values: e.values.filter((m, v) => v !== p) }),
                    children: [
                      i.jsx(M, { children: "Remove aggregate " }),
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
                  ...e.values,
                  {
                    field: r.at(-1) ?? "",
                    operation: "sum",
                    name: `Total ${e.values.length + 1}`,
                  },
                ],
              }),
            children: i.jsx(M, { children: "Add aggregate" }),
          }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Count counts nonblank values; numeric aggregates require numeric fields. No keys groups all records.",
            }),
          }),
        ],
      });
    case "pivot":
      return i.jsxs(i.Fragment, {
        children: [
          c("Pivot grouping keys", e.keys, (f) => d({ keys: f })),
          u("Pivot attribute field", e.column, r, (f) => d({ column: f })),
          u("Pivot value field", e.value, r, (f) => d({ value: f })),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Aggregation" }),
              i.jsx("select", {
                "aria-label": a("Pivot aggregation"),
                value: e.operation,
                onChange: (f) => d({ operation: f.target.value }),
                lang: s("Pivot aggregation"),
                "data-ui-label": "Pivot aggregation",
                children: B7.map((f) => i.jsx("option", { children: f }, f)),
              }),
            ],
          }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "New headers include each attribute's type and value, in first-appearance order.",
            }),
          }),
        ],
      });
    case "append":
      return i.jsxs(i.Fragment, {
        children: [
          i.jsx(CT, {
            label: a("Append source"),
            source: e.source,
            sheets: n,
            onChange: (f) => d({ source: f }),
          }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Headers must match exactly. Column order is aligned; no duplicate rows are removed.",
            }),
          }),
        ],
      });
    case "merge":
      return i.jsxs(i.Fragment, {
        children: [
          i.jsx(CT, {
            label: a("Merge source"),
            source: e.source,
            sheets: n,
            onChange: (f) => d({ source: f }),
          }),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Join mode" }),
              i.jsx("select", {
                "aria-label": a("Recipe join mode"),
                value: e.mode,
                onChange: (f) => d({ mode: f.target.value }),
                lang: s("Recipe join mode"),
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
          l("Right column prefix", e.prefix, (f) => d({ prefix: f })),
          e.keys.map((f, p) =>
            i.jsxs(
              "div",
              {
                className: "recipe-line",
                children: [
                  u(`Recipe left key ${p + 1}`, f.left, r, (m) =>
                    d({
                      keys: e.keys.map((v, g) =>
                        p === g ? { ...v, left: m } : v,
                      ),
                    }),
                  ),
                  u(`Recipe right key ${p + 1}`, f.right, t, (m) =>
                    d({
                      keys: e.keys.map((v, g) =>
                        p === g ? { ...v, right: m } : v,
                      ),
                    }),
                  ),
                  i.jsxs("button", {
                    onClick: () =>
                      d({ keys: e.keys.filter((m, v) => v !== p) }),
                    children: [i.jsx(M, { children: "Remove key " }), p + 1],
                  }),
                ],
              },
              p,
            ),
          ),
          i.jsx("button", {
            onClick: () =>
              d({ keys: [...e.keys, { left: r[0] ?? "", right: t[0] ?? "" }] }),
            children: i.jsx(M, { children: "Add merge key" }),
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
                    checked: !!((m = e.normalization) != null && m[f]),
                    onChange: (v) =>
                      d({
                        normalization: {
                          ...e.normalization,
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
              i.jsx(M, { children: "Unicode normalization" }),
              i.jsxs("select", {
                "aria-label": a("Recipe Unicode normalization"),
                value:
                  ((h = e.normalization) == null ? void 0 : h.unicode) ?? "",
                onChange: (f) =>
                  d({
                    normalization: {
                      ...e.normalization,
                      unicode: f.target.value || void 0,
                    },
                  }),
                lang: s("Recipe Unicode normalization"),
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
                checked: !!e.fuzzy,
                onChange: (f) =>
                  d({
                    fuzzy: f.target.checked
                      ? { key: 0, threshold: 0.8 }
                      : void 0,
                  }),
              }),
              i.jsx(M, { children: "Fuzzy text key" }),
            ],
          }),
          e.fuzzy &&
            i.jsxs("div", {
              className: "recipe-line",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Key pair" }),
                    i.jsx("select", {
                      "aria-label": a("Recipe fuzzy key"),
                      value: e.fuzzy.key,
                      onChange: (f) =>
                        d({
                          fuzzy: { ...e.fuzzy, key: Number(f.target.value) },
                        }),
                      lang: s("Recipe fuzzy key"),
                      "data-ui-label": "Recipe fuzzy key",
                      children: e.keys.map((f, p) =>
                        i.jsx("option", { value: p, children: p + 1 }, p),
                      ),
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Minimum similarity" }),
                    i.jsx(ci, {
                      "aria-label": a("Recipe fuzzy threshold"),
                      min: "0",
                      max: "1",
                      step: "0.01",
                      value: e.fuzzy.threshold,
                      onValueChange: (f) =>
                        d({
                          fuzzy: {
                            ...e.fuzzy,
                            threshold: f === "" ? NaN : Number(f),
                          },
                        }),
                      lang: s("Recipe fuzzy threshold"),
                      "data-ui-label": "Recipe fuzzy threshold",
                    }),
                  ],
                }),
              ],
            }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Every qualifying match is retained. Review unmatched records and many-to-many expansion in the step report.",
            }),
          }),
        ],
      });
  }
}
function Vce(e, r, t) {
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
const Wce = (e) =>
  e === null
    ? "(blank)"
    : typeof e == "boolean"
      ? String(e).toUpperCase()
      : String(e);

export { CT, ET, Lce, Pce, Q4, Uce, Vce, Wce, yz, zce };
