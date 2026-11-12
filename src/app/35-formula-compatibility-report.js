import { formatCellValueForDisplay } from "./17-recipe-step-editors.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  parseFormulaTree,
  walkFormulaNodes,
  normalizeFunctionName,
  tokenizeFormulaText,
  normalizeBindingName,
} from "./01-core-schemas-and-hashing.js";
import {
  rewriteFormulaReferences,
  frozenConversionUnits,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { inspectFormulaDependencies } from "./12-formula-dependency-inspector.js";
import {
  nameScanLimits,
  buildFunctionCatalogEntry,
  functionBrowserCategories,
  buildFunctionCatalog,
  buildFunctionBrowserGroups,
} from "./34-function-browser-and-names.js";

function buildFormulaCompatibilityReport(e, r = {}) {
  for (const v of ["includeHidden", "resolveNames", "includeMetadata"])
    if (r[v] !== undefined && typeof r[v] != "boolean")
      throw Error(`${v} must be boolean.`);
  if (
    r.allowedSheetIds !== undefined &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (v) => typeof v != "string" || !e.sheets.some((g) => g.id === v),
      ))
  )
    throw Error("Invalid authorized worksheet scope.");
  const t = { ...nameScanLimits, ...r.limits };
  for (const v of ["maxLocations", "maxCharacters", "maxFindings"])
    if (!Number.isInteger(t[v]) || t[v] < 1 || t[v] > nameScanLimits[v])
      throw Error(`Invalid ${v}; limits can only be narrowed.`);
  const n = r.includeHidden ?? true,
    o = r.resolveNames ?? true,
    a = r.includeMetadata ?? true,
    s = new Set(r.allowedSheetIds ?? e.sheets.map((v) => v.id)),
    d = e.sheets.filter((v) => s.has(v.id) && (n || !v.hidden)),
    u = new Set(d.map((v) => v.id)),
    l = {
      workbookId: e.id,
      revision: e.revision,
      scope: {
        sheetIds: [...u],
        includeHidden: n,
        resolvesNames: o,
        includesMetadata: a,
      },
      coverage: {
        locationsScanned: 0,
        charactersScanned: 0,
        findings: 0,
        complete: true,
        notes: [
          "Static function inventory only: does not certify argument accuracy, compute downstream cell dependencies, execute formulas or contact services.",
          "Unsupported function inventory is representative; unknown custom names are also reported.",
        ],
      },
      findings: [],
      families: buildFunctionBrowserGroups(),
      connectorPolicy: {
        enabled: false,
        required:
          "CUBE, network and linked-data functions stay disabled. An implemented connector and explicit trusted-host approval are required; workbook data and LLM proposals cannot authorize access.",
      },
      excelVerification: { status: "not-run", version: null },
    };
  if (
    (o ||
      l.coverage.notes.push(
        "Named definitions are outside this report scope; unresolved names require host review.",
      ),
    n ||
      l.coverage.notes.push(
        "Hidden worksheets, rows and columns are excluded by the chosen visibility scope.",
      ),
    !u.size)
  )
    return (
      l.coverage.notes.push(
        "No worksheets are in the authorized visible scope.",
      ),
      l
    );
  const c = (v) => {
    ((l.coverage.complete = false),
      l.coverage.notes.includes(v) || l.coverage.notes.push(v));
  };
  let h = false;
  const f = o ? (e.definedNames ?? []) : [],
    p = (v, g, b = []) => {
      if (
        v.length > 4096 ||
        l.coverage.charactersScanned + v.length > t.maxCharacters
      )
        return (
          c("Formula character budget reached; uninspected content remains."),
          [{ kind: "limit", message: "Formula character budget reached." }]
        );
      l.coverage.charactersScanned += v.length;
      const y = [];
      rewriteFormulaReferences(v, e, g ?? d[0]?.id ?? "", "A1") &&
        y.push({
          kind: "external-reference",
          message:
            "A reference form is external or unsupported. Inspect the original formula in the authorized workbook.",
        });
      try {
        walkFormulaNodes(tokenizeFormulaText(v), (A, k, F, I) => {
          if (I || (!A.qualifier && k.has(normalizeBindingName(A.name))))
            return;
          const O = normalizeFunctionName(A.name);
          if (!F && ["TRUE", "FALSE"].includes(O)) return;
          const j = buildFunctionCatalogEntry(O);
          if (F && !A.qualifier && j.status === "local") return;
          const z = A.qualifier
            ? e.sheets.find(
                (B) => B.name.toLowerCase() === A.qualifier.toLowerCase(),
              )
            : undefined;
          if (A.qualifier && (!z || !u.has(z.id))) {
            y.push({
              kind: "excluded",
              message:
                "A qualified name refers to a missing or excluded worksheet. Its definition was not inspected.",
            });
            return;
          }
          const V = z?.id ?? g,
            D =
              f.find(
                (B) =>
                  B.scopeSheetId === V &&
                  B.name.toLowerCase() === A.name.toLowerCase(),
              ) ??
              (A.qualifier
                ? undefined
                : f.find(
                    (B) =>
                      !B.scopeSheetId &&
                      B.name.toLowerCase() === A.name.toLowerCase(),
                  ));
          if (D) {
            const B = JSON.stringify([
              D.scopeSheetId ?? V,
              D.name.toLowerCase(),
            ]);
            if (b.includes(B)) {
              parseFormulaTree(D.refersTo) ||
                y.push({
                  kind: "name-cycle",
                  viaName: D.name,
                  message:
                    "Circular named definition; inspection stopped at the cycle.",
                });
              return;
            }
            if (b.length >= nameScanLimits.maxNameDepth) {
              (c(
                "Named-definition depth reached; uninspected definitions remain.",
              ),
                y.push({
                  kind: "limit",
                  message:
                    "Named-definition inspection is limited to 32 levels.",
                }));
              return;
            }
            const _ = p(D.refersTo, D.scopeSheetId ?? V, [...b, B]);
            y.push(..._.map((C) => ({ ...C, viaName: D.name })));
            return;
          }
          if (!F) {
            y.push({
              kind: "unresolved-name",
              message:
                "A name is unresolved or its definition is outside this report scope.",
            });
            return;
          }
          y.push({
            kind:
              j.status === "connector-required"
                ? "connector-required"
                : "unsupported-function",
            functionName: O,
            family: j.family,
            message:
              j.status === "connector-required"
                ? "Disabled: an approved, implemented host connector is required. No data was requested."
                : "This function has no local implementation. Review its source or add-in contract.",
          });
        });
      } catch {
        (y.push({
          kind: "syntax",
          message:
            "Formula syntax or lexical bindings could not be inspected. Review the original formula.",
        }),
          (l.coverage.complete = false));
      }
      return [...new Map(y.map((A) => [JSON.stringify(A), A])).values()];
    },
    m = (v, g, b, y) => {
      if (h) return;
      if (++l.coverage.locationsScanned > t.maxLocations) {
        (l.coverage.locationsScanned--,
          (h = true),
          c(
            "Location budget reached; uninspected cells or definitions remain.",
          ));
        return;
      }
      const w = p(v, y?.id);
      for (const E of w) {
        if (l.findings.length >= t.maxFindings) {
          ((h = true),
            c(
              "Finding budget reached; additional findings were not included.",
            ));
          break;
        }
        l.findings.push({
          ...E,
          source: g,
          location: b,
          ...(y ? { sheetId: y.id, sheetName: y.name } : {}),
          message:
            (g === "literal"
              ? "Stored as literal text; will not execute. "
              : "") + E.message,
        });
      }
      l.coverage.charactersScanned >= t.maxCharacters && (h = true);
    };
  for (const v of d) {
    const g = cn(v, "row"),
      b = cn(v, "column");
    for (const [y, w] of Object.entries(v.cells)) {
      if (h) break;
      if (!(typeof w.raw != "string" || !w.raw.startsWith("="))) {
        if (!n) {
          const E = _t(y);
          if (
            g.some(({ start, end }) => E.row >= start && E.row <= end) ||
            b.some(({ start, end }) => E.col >= start && E.col <= end)
          )
            continue;
        }
        m(w.raw, w.literal ? "literal" : "cell", y, v);
      }
    }
    if (a && n) {
      for (const y of v.conditionalFormats ?? [])
        y.formula && m(y.formula, "conditional-format", y.range, v);
      for (const y of v.tables ?? [])
        for (const w of y.columns)
          w.calculatedFormula &&
            m(w.calculatedFormula, "table-template", `${y.name}[${w.name}]`, v);
    }
  }
  if (a && o)
    for (const v of f)
      (v.scopeSheetId && !u.has(v.scopeSheetId)) ||
        m(
          v.refersTo,
          "name",
          v.name,
          d.find((g) => g.id === v.scopeSheetId),
        );
  return (
    a &&
      !n &&
      l.coverage.notes.push(
        "Conditional and table templates are excluded when hidden content is excluded; named definitions remain in scope.",
      ),
    (l.coverage.findings = l.findings.length),
    l
  );
}
function FormulaCompatibilityDialog(e) {
  const { text, language } = useWorkbookMessages(),
    [n, o] = ne.useState("report"),
    [a, s] = ne.useState(true),
    [d, u] = ne.useState(""),
    [l, c] = ne.useState(""),
    [h, f] = ne.useState("DSUM"),
    p = ne.useMemo(() => {
      try {
        return {
          report: buildFormulaCompatibilityReport(e.workbook, {
            allowedSheetIds: e.allowedSheetIds,
            includeHidden: a,
          }),
        };
      } catch (w) {
        return {
          error:
            w instanceof Error ? w.message : "Could not inspect this workbook.",
        };
      }
    }, [e.workbook, e.allowedSheetIds, a]),
    v = ne
      .useMemo(buildFunctionCatalog, [])
      .filter(
        (w) =>
          (!l || w.family === l) &&
          `${w.name} ${w.outputs}`.toLowerCase().includes(d.toLowerCase()),
      ),
    g = v.find((w) => w.name === h) ?? v[0],
    b = p.report,
    y = () => {
      if (!b) return;
      if (e.onExport) return e.onExport(b);
      const w = new Blob([JSON.stringify(b, null, 2)], {
          type: "application/json",
        }),
        E = URL.createObjectURL(w),
        A = document.createElement("a");
      ((A.href = E),
        (A.download = "Formula-compatibility.json"),
        A.click(),
        setTimeout(() => URL.revokeObjectURL(E), 1e3));
    };
  return i.jsx("div", {
    className: "formula-tool-backdrop",
    onMouseDown: (w) => {
      w.target === w.currentTarget && e.onClose();
    },
    onKeyDown: (w) => {
      w.key === "Escape" && e.onClose();
    },
    children: i.jsxs("section", {
      className: "formula-tools function-compatibility",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": text("Formula compatibility"),
      lang: language("Formula compatibility"),
      "data-ui-label": "Formula compatibility",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children: i.jsx(LocalizedText, {
                children: "Formula compatibility",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close compatibility"),
              onClick: e.onClose,
              lang: language("Close compatibility"),
              "data-ui-label": "Close compatibility",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Review support limits and affected locations before relying on calculations. Excel version comparison has not been run.",
          }),
        }),
        i.jsxs("div", {
          role: "tablist",
          "aria-label": text("Compatibility views"),
          lang: language("Compatibility views"),
          "data-ui-label": "Compatibility views",
          children: [
            i.jsx("button", {
              autoFocus: true,
              role: "tab",
              "aria-selected": n === "report",
              onClick: () => o("report"),
              children: i.jsx(LocalizedText, {
                children: "Compatibility report",
              }),
            }),
            i.jsx("button", {
              role: "tab",
              "aria-selected": n === "registry",
              onClick: () => o("registry"),
              children: i.jsx(LocalizedText, { children: "Function registry" }),
            }),
          ],
        }),
        n === "report"
          ? i.jsxs("div", {
              role: "tabpanel",
              "aria-label": text("Compatibility report"),
              lang: language("Compatibility report"),
              "data-ui-label": "Compatibility report",
              children: [
                i.jsxs("label", {
                  className: "compatibility-checkbox",
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: a,
                      onChange: (w) => s(w.target.checked),
                    }),
                    i.jsx(LocalizedText, {
                      children: "Include hidden worksheets, rows and columns",
                    }),
                  ],
                }),
                p.error && i.jsx("p", { role: "alert", children: p.error }),
                b &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("p", {
                        role: "status",
                        children: [
                          i.jsx(LocalizedText, { children: "Revision " }),
                          b.revision,
                          " · ",
                          b.coverage.locationsScanned,
                          " ",
                          i.jsx(LocalizedText, {
                            children: "locations inspected ·",
                          }),
                          " ",
                          b.findings.length,
                          " ",
                          i.jsx(LocalizedText, { children: "findings ·" }),
                          " ",
                          b.coverage.complete
                            ? "Scan complete within chosen scope"
                            : "Incomplete scan — review limits below",
                        ],
                      }),
                      i.jsx("div", {
                        className: "compatibility-findings",
                        children: i.jsxs("table", {
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Location",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Function / source",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Finding",
                                    }),
                                  }),
                                  i.jsx("th", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Action",
                                    }),
                                  }),
                                ],
                              }),
                            }),
                            i.jsxs("tbody", {
                              children: [
                                b.findings.slice(0, 200).map((w, E) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsxs("td", {
                                          children: [
                                            w.sheetName
                                              ? `${w.sheetName}!`
                                              : "",
                                            w.location,
                                            i.jsx("small", {
                                              children: w.source,
                                            }),
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            w.functionName ??
                                              "Reference or name",
                                            w.viaName &&
                                              i.jsxs("small", {
                                                children: [
                                                  i.jsx(LocalizedText, {
                                                    children: "Via ",
                                                  }),
                                                  w.viaName,
                                                ],
                                              }),
                                          ],
                                        }),
                                        i.jsx("td", { children: w.message }),
                                        i.jsx("td", {
                                          children:
                                            (w.source === "cell" ||
                                              w.source === "literal") &&
                                            w.sheetId &&
                                            i.jsx("button", {
                                              onClick: () => {
                                                (e.onNavigate(
                                                  w.sheetId,
                                                  w.location,
                                                ),
                                                  e.onClose());
                                              },
                                              children: i.jsx(LocalizedText, {
                                                children: "Go to cell",
                                              }),
                                            }),
                                        }),
                                      ],
                                    },
                                    E,
                                  ),
                                ),
                                !b.findings.length &&
                                  i.jsx("tr", {
                                    children: i.jsx("td", {
                                      colSpan: 4,
                                      children: i.jsx(LocalizedText, {
                                        children:
                                          "No unsupported function calls found in the inspected scope. This does not establish calculation accuracy.",
                                      }),
                                    }),
                                  }),
                              ],
                            }),
                          ],
                        }),
                      }),
                      b.findings.length > 200 &&
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children:
                                "Showing 200 findings. Export the report for all",
                            }),
                            " ",
                            b.findings.length,
                            ".",
                          ],
                        }),
                      i.jsxs("details", {
                        children: [
                          i.jsx("summary", {
                            children: i.jsx(LocalizedText, {
                              children: "Family inventory and engagement uses",
                            }),
                          }),
                          b.families.map((w) =>
                            i.jsxs(
                              "p",
                              {
                                children: [
                                  i.jsx("strong", { children: w.title }),
                                  " · ",
                                  w.adopted.length,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "adopted ·",
                                  }),
                                  " ",
                                  w.unavailable.length,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "inventoried as unavailable",
                                  }),
                                  i.jsx("br", {}),
                                  w.demand,
                                ],
                              },
                              w.id,
                            ),
                          ),
                        ],
                      }),
                      i.jsx("p", { children: b.connectorPolicy.required }),
                      i.jsx("ul", {
                        children: b.coverage.notes.map((w) =>
                          i.jsx("li", { children: w }, w),
                        ),
                      }),
                    ],
                  }),
                i.jsxs("footer", {
                  children: [
                    i.jsx("button", {
                      onClick: e.onClose,
                      children: i.jsx(LocalizedText, { children: "Close" }),
                    }),
                    i.jsx("button", {
                      disabled: !b,
                      onClick: y,
                      children: i.jsx(LocalizedText, {
                        children: "Export compatibility report",
                      }),
                    }),
                  ],
                }),
              ],
            })
          : i.jsxs("div", {
              role: "tabpanel",
              "aria-label": text("Function registry"),
              lang: language("Function registry"),
              "data-ui-label": "Function registry",
              children: [
                i.jsxs("div", {
                  className: "compatibility-filters",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Search function registry",
                        }),
                        i.jsx("input", {
                          value: d,
                          onChange: (w) => u(w.target.value),
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Function family" }),
                        i.jsxs("select", {
                          value: l,
                          onChange: (w) => c(w.target.value),
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "All families",
                            }),
                            functionBrowserCategories.map((w) =>
                              i.jsx(
                                "option",
                                { value: w.id, children: w.title },
                                w.id,
                              ),
                            ),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
                g
                  ? i.jsxs("div", {
                      className: "formula-function-layout",
                      children: [
                        i.jsx("select", {
                          "aria-label": text("Function contracts"),
                          size: 12,
                          value: g.name,
                          onChange: (w) => f(w.target.value),
                          lang: language("Function contracts"),
                          "data-ui-label": "Function contracts",
                          children: v.map((w) =>
                            i.jsxs(
                              "option",
                              {
                                value: w.name,
                                children: [
                                  w.name,
                                  " ·",
                                  " ",
                                  w.status === "local"
                                    ? "Local"
                                    : w.status === "connector-required"
                                      ? "Connector required"
                                      : "Unavailable",
                                ],
                              },
                              w.name,
                            ),
                          ),
                        }),
                        i.jsxs("div", {
                          "aria-label": text("Selected function contract"),
                          lang: language("Selected function contract"),
                          "data-ui-label": "Selected function contract",
                          children: [
                            i.jsx("h3", { children: g.name }),
                            i.jsxs("p", {
                              className: "formula-signature",
                              children: [g.name, "(", g.inputs, ")"],
                            }),
                            i.jsxs("dl", {
                              children: [
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Output",
                                  }),
                                }),
                                i.jsx("dd", { children: g.outputs }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Errors",
                                  }),
                                }),
                                i.jsx("dd", { children: g.errors.join(", ") }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Numeric tolerance",
                                  }),
                                }),
                                i.jsx("dd", {
                                  children: g.numericTolerance
                                    ? `Absolute ${g.numericTolerance.absolute}; relative ${g.numericTolerance.relative}. ${g.numericTolerance.appliesTo}`
                                    : "Not applicable — no local implementation.",
                                }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Verified Excel version",
                                  }),
                                }),
                                i.jsx("dd", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Not run; no verified version.",
                                  }),
                                }),
                                i.jsx("dt", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Limits",
                                  }),
                                }),
                                i.jsx("dd", { children: g.limits }),
                              ],
                            }),
                            g.name === "CONVERT" &&
                              i.jsxs("details", {
                                children: [
                                  i.jsx("summary", {
                                    children: i.jsx(LocalizedText, {
                                      children: "Supported base unit symbols",
                                    }),
                                  }),
                                  i.jsx("p", {
                                    children: frozenConversionUnits
                                      .map((w) => w.symbol)
                                      .join(", "),
                                  }),
                                  i.jsx("p", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Decimal prefixes apply only to metric entries. Binary prefixes apply only to bit and byte. Unknown or incompatible units return #N/A.",
                                    }),
                                  }),
                                ],
                              }),
                          ],
                        }),
                      ],
                    })
                  : i.jsx("p", {
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children: "No matching functions.",
                      }),
                    }),
              ],
            }),
      ],
    }),
  });
}
function describeEvaluatedValue(e) {
  if (e instanceof Error && "error" in e)
    return { kind: "error", text: String(e.error) };
  if (Array.isArray(e)) {
    const a = e.every(Array.isArray) ? e : [e],
      s = [];
    for (const d of a) {
      for (const u of d) {
        if (s.length === 8) break;
        s.push(describeEvaluatedValue(u));
      }
      if (s.length === 8) break;
    }
    return {
      kind: "array",
      text: `${a.length} × ${a[0]?.length ?? 0} array`,
      rows: a.length,
      columns: a[0]?.length ?? 0,
      preview: s,
      truncated: a.length * (a[0]?.length ?? 0) > 8,
    };
  }
  if (e && typeof e == "object") {
    if ("kind" in e && e.kind === "formula-closure")
      return { kind: "function", text: "LAMBDA function" };
    if ("ref" in e && e.ref && typeof e.ref == "object") {
      const a = e.ref,
        s = (d) =>
          Number.isInteger(d.row) && Number.isInteger(d.col) && d.col > 0
            ? `${columnIndexToLetters(d.col - 1)}${d.row}`
            : "axis reference";
      return {
        kind: "reference",
        text: `${a.sheet ? `${a.sheet}!` : ""}${a.from && a.to ? `${s(a.from)}:${s(a.to)}` : s(a)}`,
      };
    }
    return { kind: "unavailable", text: "Internal value" };
  }
  const r =
    e == null
      ? "(blank)"
      : typeof e == "boolean"
        ? String(e).toUpperCase()
        : String(e);
  return {
    kind: "value",
    text: r.slice(0, 256),
    ...(r.length <= 256 ? { value: e ?? null } : { truncated: true }),
  };
}
const validateInspectionLimit = (e, r) => {
  if (e === undefined) return r;
  if (!Number.isInteger(e) || e < 1 || e > r)
    throw Error(`Inspection limits must be integers between 1 and ${r}.`);
  return e;
};
function inspectFormulaCalculation(e, r) {
  const t = validateInspectionLimit(r.limits?.maxSteps, 500),
    n = validateInspectionLimit(r.limits?.maxReads, 2e3);
  if (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
    throw Error("includeHidden must be boolean.");
  if (r.includeNames !== undefined && typeof r.includeNames != "boolean")
    throw Error("includeNames must be boolean.");
  if (
    r.allowedSheetIds !== undefined &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some((C) => !e.sheets.some((T) => T.id === C)))
  )
    throw Error("Invalid authorized worksheet context.");
  const o = new Set(r.allowedSheetIds ?? e.sheets.map((C) => C.id)),
    a = _t(r.address),
    s = `${columnIndexToLetters(a.col)}${a.row + 1}`,
    d = r.allowedRanges?.map((C) => {
      if (!o.has(C.sheetId))
        throw Error(
          "Inspection ranges must remain within authorized worksheets.",
        );
      return { ...C, bounds: parseRangeAddress(C.range) };
    }),
    u = new Map(
      e.sheets.map((C) => [
        C.id,
        { sheet: C, rows: cn(C, "row"), cols: cn(C, "column") },
      ]),
    ),
    l = (C, T) => {
      const N = u.get(C);
      if (!N || !o.has(C)) return false;
      const { row, col } = _t(T);
      return row >= N.sheet.rowCount ||
        col >= N.sheet.colCount ||
        (d &&
          !d.some(
            ($) =>
              $.sheetId === C &&
              row >= $.bounds.firstRow &&
              row <= $.bounds.lastRow &&
              col >= $.bounds.firstCol &&
              col <= $.bounds.lastCol,
          ))
        ? false
        : !!r.includeHidden ||
            (!N.sheet.hidden &&
              !N.rows.some(($) => row >= $.start && row <= $.end) &&
              !N.cols.some(($) => col >= $.start && col <= $.end));
    };
  if (!l(r.sheetId, s))
    throw Error(
      "The selected cell is outside the visible, authorized inspection context.",
    );
  const c = [],
    h = [];
  let f = 0,
    p = 0,
    m = false,
    v = false;
  const g = (C, T) => {
      !h.some((N) => N.kind === C && N.message === T) &&
        h.length < 50 &&
        h.push({ kind: C, message: T });
    },
    b = r.includeNames === false ? { ...e, definedNames: [] } : e,
    y = calculateWorkbookWithSettings(
      b,
      {},
      {
        allowName(C) {
          var T;
          return !o.has(C) ||
            (!r.includeHidden && (T = u.get(C)) != null && T.sheet.hidden)
            ? ((m = true),
              g(
                "excluded",
                "A defined name is outside the visible, authorized context. Its definition was not evaluated and the final result is withheld.",
              ),
              false)
            : true;
        },
        allowRead(C, T) {
          if (++f > n)
            return (
              (v = true),
              g(
                "limit",
                "The calculation read budget was reached; the final result is withheld.",
              ),
              false
            );
          const N = u.get(C)?.sheet;
          if (!N) return true;
          const R = _t(T);
          return R.row >= N.rowCount || R.col >= N.colCount || l(C, T)
            ? true
            : ((m = true),
              g(
                "excluded",
                "A calculation attempted to read outside the visible, authorized context. The source was not read and the final result is withheld.",
              ),
              false);
        },
        step(C) {
          l(C.sheetId, C.address) &&
            (p++,
            !(c.length >= t) &&
              c.push({
                index: c.length + 1,
                kind: C.kind,
                sheetId: C.sheetId,
                address: C.address,
                operation: C.operation.slice(0, 1024),
                inputs: (C.inputs ?? [])
                  .slice(0, 10)
                  .map(describeEvaluatedValue),
                result: describeEvaluatedValue(C.result),
              }));
        },
      },
    ),
    w = e.sheets.find((C) => C.id === r.sheetId).cells[s],
    E = w?.raw ?? null,
    A = typeof E == "string" && E.startsWith("=") && !(w != null && w.literal),
    k = y(r.sheetId, s),
    { range, diagnostic } = y.peekSpill(r.sheetId, s);
  for (const C of y.getReferenceDiagnostics()) g("reference", C.message);
  for (const C of y.getNamedFormulaDiagnostics()) g("name", C.message);
  for (const C of y.getFinancialDiagnostics())
    (C.status === "error" ||
      (C.solver &&
        (C.solver.candidates.length > 1 || (C.signChanges ?? 0) > 1))) &&
      g("calculation", `${C.functionName}: ${C.message}`);
  const O = y.getOperationalErrors();
  O.length &&
    g("calculation", `Local calculation diagnostics: ${O.join(", ")}.`);
  const j = p > t;
  j &&
    g(
      "limit",
      `Only the first ${t} steps were retained; calculation continued within its own limits.`,
    );
  const z = m || v;
  return {
    workbookId: e.id,
    dateContext: y.getDateContext(),
    ...(!z && y.getFinancialDiagnostics().length
      ? { financial: y.getFinancialDiagnostics() }
      : {}),
    revision: e.revision,
    sheetId: r.sheetId,
    address: s,
    raw: E,
    result: z ? null : k,
    resultStatus: z ? "withheld" : A ? "calculated" : "literal",
    steps: c,
    firstErrorStep: c.find((C) => C.result.kind === "error")?.index,
    ...(!z && range ? { spillRange: range } : {}),
    ...(!z && diagnostic ? { spillDiagnostic: diagnostic } : {}),
    issues: h,
    coverage: {
      complete: !j && !z && !O.length,
      truncated: j,
      reads: f,
      stepsObserved: p,
      maxSteps: t,
      maxReads: n,
    },
    notes: [
      "Steps record one actual local calculation, in execution order. Function internals are a single step; references are shown as references without a second read.",
      "Restart and Previous replay the captured result. Reopen or change the workbook revision to recalculate; volatile values may then change.",
      "Array previews contain at most 8 values; text previews contain at most 256 characters and function inputs at most 10 arguments. Full Excel step-order and native tracer-arrow parity are unverified.",
      "Static precedents include potential branches; this log shows only work performed by the local engine, including spill discovery and cache reads. Host authorization remains the embedding application's responsibility.",
    ],
  };
}
const calculationErrorHints = {
  "#DIV/0!": "Review the denominator and blank inputs.",
  "#REF!":
    "A reference is missing, invalid, or no longer points to an available spill.",
  "#NAME?":
    "Review function spelling, defined names, and the supported function catalog.",
  "#VALUE!": "Review argument types and array dimensions.",
  "#N/A": "No matching value was found or a requested value is unavailable.",
  "#NUM!": "Review numeric domains, convergence, and overflow.",
  "#CYCLE!":
    "A calculation depends on itself. Review the recorded cell reads and static precedents.",
  "#SPILL!":
    "The array cannot occupy its output range. Review the obstruction below.",
  "#BLOCKED!":
    "External execution or a source outside inspection scope was refused.",
  "#LIMIT!": "The bounded local calculator reached a work or size limit.",
  "#CALC!": "Review empty arrays and uninvoked reusable functions.",
};
function EvaluateFormulaDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    [a, s] = ne.useState(false),
    [d, u] = ne.useState(0),
    [l, c] = ne.useState(""),
    h = ne.useMemo(() => {
      try {
        return {
          report: inspectFormulaCalculation(e.workbook, {
            sheetId: e.sheetId,
            address: e.address,
            allowedSheetIds: e.allowedSheetIds,
            includeHidden: !!e.allowHidden && a,
          }),
        };
      } catch (b) {
        return {
          error:
            b instanceof Error
              ? b.message
              : "Calculation could not be inspected.",
        };
      }
    }, [e.workbook, e.sheetId, e.address, e.allowedSheetIds, e.allowHidden, a]);
  (ne.useEffect(() => {
    (u(0), c(""));
  }, [h]),
    ne.useEffect(() => {
      var w;
      const b = document.activeElement;
      return (
        (w = o.current?.querySelector("button")) == null || w.focus(),
        () => {
          b != null && b.isConnected && b.focus();
        }
      );
    }, []));
  const f = h.report,
    p = f ? Math.min(d, f.steps.length - 1) : -1,
    m = f?.steps[p],
    v = (b, y) => {
      return `${e.workbook.sheets.find((E) => E.id === b)?.name ?? "Unavailable sheet"}!${y}`;
    },
    g = () => {
      if (f)
        try {
          if (e.onExport) e.onExport(structuredClone(f));
          else {
            const b = URL.createObjectURL(
                new Blob([JSON.stringify(f, null, 2)], {
                  type: "application/json",
                }),
              ),
              y = document.createElement("a");
            ((y.href = b),
              (y.download = `Calculation-${f.address}-revision-${f.revision}.json`),
              y.click(),
              setTimeout(() => URL.revokeObjectURL(b), 1e3));
          }
          c("");
        } catch (b) {
          c(b instanceof Error ? b.message : "Export failed.");
        }
    };
  return i.jsx("div", {
    className: "fi-backdrop",
    onMouseDown: (b) => {
      b.target === b.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "fi-dialog fe-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      ref: o,
      onKeyDown: (b) => {
        if (
          (b.stopPropagation(),
          b.key === "Escape" && (b.preventDefault(), e.onClose()),
          b.key === "Tab")
        ) {
          const y = [
              ...o.current.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),summary",
              ),
            ],
            w = y[0],
            E = y.at(-1);
          b.shiftKey && document.activeElement === w
            ? (b.preventDefault(), E == null || E.focus())
            : !b.shiftKey &&
              document.activeElement === E &&
              (b.preventDefault(), w == null || w.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Evaluate Formula",
                  }),
                }),
                i.jsxs("p", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Read-only calculation replay · revision ",
                    }),
                    e.workbook.revision,
                  ],
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close formula evaluation"),
              onClick: e.onClose,
              lang: language("Close formula evaluation"),
              "data-ui-label": "Close formula evaluation",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fi-controls",
          children: [
            e.allowHidden &&
              i.jsxs("label", {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: a,
                    onChange: (b) => s(b.target.checked),
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden cells" }),
                ],
              }),
            i.jsx("button", {
              disabled: !f,
              onClick: g,
              children: i.jsx(LocalizedText, {
                children: "Export calculation steps",
              }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fe-body",
          children: [
            h.error && i.jsx("p", { role: "alert", children: h.error }),
            l && i.jsx("p", { role: "alert", children: l }),
            f &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("section", {
                    className: "fi-target",
                    "aria-label": text("Evaluation target"),
                    lang: language("Evaluation target"),
                    "data-ui-label": "Evaluation target",
                    children: [
                      i.jsx("h3", { children: v(f.sheetId, f.address) }),
                      i.jsx("code", { children: String(f.raw ?? "(blank)") }),
                      i.jsxs("p", {
                        children: [
                          i.jsx("strong", {
                            children: i.jsx(LocalizedText, {
                              children: "Final result: ",
                            }),
                          }),
                          f.resultStatus === "withheld"
                            ? "Withheld"
                            : String(f.result ?? "(blank)"),
                          f.resultStatus === "literal"
                            ? " · literal input"
                            : "",
                        ],
                      }),
                      f.spillRange &&
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Spill output: ",
                            }),
                            f.spillRange,
                          ],
                        }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: `fi-coverage ${f.coverage.complete ? "" : "fi-warning"}`,
                    role: "status",
                    children: [
                      f.coverage.complete
                        ? "Calculation captured"
                        : "Partial calculation capture",
                      " ",
                      "· ",
                      f.coverage.reads,
                      " ",
                      i.jsx(LocalizedText, { children: "reads · " }),
                      f.steps.length,
                      " ",
                      i.jsx(LocalizedText, { children: "retained steps" }),
                    ],
                  }),
                  f.issues.length > 0 &&
                    i.jsx("ul", {
                      "aria-label": text("Evaluation diagnostics"),
                      lang: language("Evaluation diagnostics"),
                      "data-ui-label": "Evaluation diagnostics",
                      children: f.issues.map((b, y) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsxs("strong", { children: [b.kind, ": "] }),
                              b.message,
                            ],
                          },
                          y,
                        ),
                      ),
                    }),
                  f.spillDiagnostic &&
                    i.jsx("p", {
                      className: "fi-warning",
                      children: f.spillDiagnostic.message,
                    }),
                  i.jsxs("div", {
                    className: "fi-controls fe-playback",
                    children: [
                      i.jsx("button", {
                        onClick: () => u(0),
                        disabled: p <= 0,
                        children: i.jsx(LocalizedText, { children: "Restart" }),
                      }),
                      i.jsx("button", {
                        onClick: () => u((b) => b - 1),
                        disabled: p <= 0,
                        children: i.jsx(LocalizedText, {
                          children: "Previous step",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () => u((b) => b + 1),
                        disabled: !f.steps.length || p >= f.steps.length - 1,
                        children: i.jsx(LocalizedText, {
                          children: "Next step",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () => u(f.steps.length - 1),
                        disabled: !f.steps.length || p >= f.steps.length - 1,
                        children: i.jsx(LocalizedText, {
                          children: "Last step",
                        }),
                      }),
                      i.jsx("button", {
                        onClick: () => u(f.firstErrorStep - 1),
                        disabled: !f.firstErrorStep,
                        children: i.jsx(LocalizedText, {
                          children: "First error",
                        }),
                      }),
                    ],
                  }),
                  m &&
                    i.jsxs("section", {
                      className: "fe-step",
                      "aria-label": text("Current calculation step"),
                      "aria-live": "polite",
                      lang: language("Current calculation step"),
                      "data-ui-label": "Current calculation step",
                      children: [
                        i.jsxs("h3", {
                          children: [
                            i.jsx(LocalizedText, { children: "Step " }),
                            m.index,
                            " ",
                            i.jsx(LocalizedText, { children: "of " }),
                            f.steps.length,
                            " · ",
                            m.kind,
                          ],
                        }),
                        i.jsx("p", {
                          children: i.jsx("button", {
                            onClick: () => {
                              try {
                                (e.onNavigate(m.sheetId, m.address), c(""));
                              } catch (b) {
                                c(
                                  b instanceof Error
                                    ? b.message
                                    : "Navigation failed.",
                                );
                              }
                            },
                            children: v(m.sheetId, m.address),
                          }),
                        }),
                        i.jsx("code", { children: m.operation }),
                        m.inputs.length > 0 &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, { children: "Inputs: " }),
                              m.inputs.map((b) => b.text).join("; "),
                            ],
                          }),
                        i.jsxs("p", {
                          children: [
                            i.jsx("strong", {
                              children: i.jsx(LocalizedText, {
                                children: "Result: ",
                              }),
                            }),
                            i.jsx("code", { children: m.result.text }),
                            m.result.truncated && " (preview truncated)",
                          ],
                        }),
                        m.result.preview &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Array preview:",
                              }),
                              " ",
                              m.result.preview.map((b) => b.text).join(", "),
                            ],
                          }),
                        m.result.kind === "error" &&
                          i.jsx("p", {
                            className: "fi-warning",
                            children:
                              calculationErrorHints[m.result.text] ??
                              "Review the source formula and arguments.",
                          }),
                      ],
                    }),
                  i.jsxs("details", {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Captured steps (",
                          }),
                          f.steps.length,
                          ")",
                        ],
                      }),
                      i.jsx("ol", {
                        className: "fe-log",
                        children: f.steps.map((b) =>
                          i.jsx(
                            "li",
                            {
                              children: i.jsxs("button", {
                                "aria-current":
                                  b.index === m?.index ? "step" : undefined,
                                onClick: () => u(b.index - 1),
                                children: [
                                  b.index,
                                  ". ",
                                  v(b.sheetId, b.address),
                                  " · ",
                                  b.operation,
                                  " ",
                                  "→ ",
                                  b.result.text,
                                ],
                              }),
                            },
                            b.index,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("details", {
                    children: [
                      i.jsx("summary", {
                        children: i.jsx(LocalizedText, {
                          children: "Scope and calculation limits",
                        }),
                      }),
                      i.jsx("ul", {
                        children: f.notes.map((b) =>
                          i.jsx("li", { children: b }, b),
                        ),
                      }),
                    ],
                  }),
                ],
              }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx(LocalizedText, {
            children:
              "Inspection and replay do not edit formulas, history, or calculation settings.",
          }),
        }),
      ],
    }),
  });
}

function CalculationInspectorDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const [s, d] = ne.useState(e.initialDirection ?? "precedents"),
    [u, l] = ne.useState(1),
    [c, h] = ne.useState(false),
    [f, p] = ne.useState(0),
    [m, v] = ne.useState(""),
    g = ne.useMemo(() => {
      try {
        return {
          result: inspectFormulaDependencies(e.workbook, {
            sheetId: e.sheetId,
            address: e.address,
            direction: s,
            depth: u,
            includeHidden: !!e.allowHidden && c,
            allowedSheetIds: e.allowedSheetIds,
          }),
        };
      } catch (k) {
        return {
          error:
            k instanceof Error
              ? k.message
              : "The selected cell could not be inspected.",
        };
      }
    }, [
      e.workbook,
      e.sheetId,
      e.address,
      e.allowedSheetIds,
      e.allowHidden,
      s,
      u,
      c,
    ]);
  (ne.useEffect(() => {
    (p(0), v(""));
  }, [g]),
    ne.useEffect(() => {
      var I;
      const k = document.activeElement;
      return (
        (I = o.current?.querySelector('[role="tab"][aria-selected="true"]')) ==
          null || I.focus(),
        () => {
          k != null && k.isConnected && k.focus();
        }
      );
    }, []));
  const b = g.result,
    y = (k) => {
      const [F, I] = JSON.parse(k);
      return `${e.workbook.sheets.find((j) => j.id === F)?.name ?? "Unavailable sheet"}!${I}`;
    },
    w = Math.max(1, Math.ceil((b?.nodes.length ?? 0) / 50)),
    E = Math.min(f, w - 1),
    A = (k) => {
      try {
        (inspectFormulaDependencies(e.workbook, {
          sheetId: k.sheetId,
          address: k.address,
          direction: "precedents",
          allowedSheetIds: e.allowedSheetIds,
          includeHidden: !!e.allowHidden && c,
          limits: { maxFormulaCells: 1, maxNodes: 1, maxReferences: 1 },
        }),
          e.onNavigate(k.sheetId, k.address),
          v(""));
      } catch (F) {
        v(
          F instanceof Error
            ? F.message
            : "The host could not navigate to this cell.",
        );
      }
    };
  return i.jsx("div", {
    className: "fi-backdrop",
    onMouseDown: (k) => {
      k.target === k.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: o,
      className: "fi-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (k) => {
        if (
          (k.stopPropagation(),
          k.key === "Escape" && (k.preventDefault(), a.current()),
          k.key === "Tab")
        ) {
          const I = [
              ...(o.current?.querySelectorAll(
                'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
              ) ?? []),
            ].filter((z) => !z.hidden),
            O = I[0],
            j = I.at(-1);
          k.shiftKey && document.activeElement === O
            ? (k.preventDefault(), j == null || j.focus())
            : !k.shiftKey &&
              document.activeElement === j &&
              (k.preventDefault(), O == null || O.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Inspect calculations",
                  }),
                }),
                i.jsxs("p", {
                  children: [
                    i.jsx(LocalizedText, {
                      children: "Read-only formula relationships · revision ",
                    }),
                    e.workbook.revision,
                  ],
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close calculation inspection"),
              onClick: e.onClose,
              lang: language("Close calculation inspection"),
              "data-ui-label": "Close calculation inspection",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fi-controls",
          children: [
            i.jsx("div", {
              role: "tablist",
              "aria-label": text("Formula relationships"),
              lang: language("Formula relationships"),
              "data-ui-label": "Formula relationships",
              children: ["precedents", "dependents"].map((k) =>
                i.jsx(
                  "button",
                  {
                    role: "tab",
                    id: `${n}-${k}`,
                    "aria-controls": `${n}-panel`,
                    "aria-selected": s === k,
                    tabIndex: s === k ? 0 : -1,
                    onClick: () => d(k),
                    onKeyDown: (F) => {
                      var I;
                      if (
                        ["ArrowLeft", "ArrowRight", "Home", "End"].includes(
                          F.key,
                        )
                      ) {
                        F.preventDefault();
                        const O =
                          F.key === "Home"
                            ? "precedents"
                            : F.key === "End" || k === "precedents"
                              ? "dependents"
                              : "precedents";
                        (d(O),
                          (I = document.getElementById(`${n}-${O}`)) == null ||
                            I.focus());
                      }
                    },
                    children: k === "precedents" ? "Precedents" : "Dependents",
                  },
                  k,
                ),
              ),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Traversal depth" }),
                i.jsx("select", {
                  "aria-label": text("Traversal depth"),
                  value: u,
                  onChange: (k) => l(Number(k.target.value)),
                  lang: language("Traversal depth"),
                  "data-ui-label": "Traversal depth",
                  children: Array.from({ length: 8 }, (k, F) =>
                    i.jsx(
                      "option",
                      {
                        value: F + 1,
                        children:
                          F === 0 ? "1 — Direct" : `${F + 1} — Transitive`,
                      },
                      F,
                    ),
                  ),
                }),
              ],
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "fi-hidden",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: c,
                    onChange: (k) => h(k.target.checked),
                  }),
                  i.jsx(LocalizedText, {
                    children: "Include hidden cells in the allowed context",
                  }),
                ],
              }),
          ],
        }),
        g.error &&
          i.jsx("p", {
            role: "alert",
            className: "fi-warning",
            children: g.error,
          }),
        m &&
          i.jsx("p", { role: "alert", className: "fi-warning", children: m }),
        b &&
          i.jsxs("div", {
            role: "tabpanel",
            id: `${n}-panel`,
            "aria-labelledby": `${n}-${s}`,
            children: [
              i.jsxs("section", {
                className: "fi-target",
                "aria-label": text("Selected cell"),
                lang: language("Selected cell"),
                "data-ui-label": "Selected cell",
                children: [
                  i.jsxs("h3", {
                    children: [b.target.sheetName, "!", b.target.address],
                  }),
                  i.jsxs("dl", {
                    children: [
                      i.jsx("dt", {
                        children: i.jsx(LocalizedText, {
                          children: "Stored input",
                        }),
                      }),
                      i.jsx("dd", {
                        children: i.jsx("code", {
                          children: formatCellValueForDisplay(b.target.raw),
                        }),
                      }),
                      i.jsx("dt", {
                        children: i.jsx(LocalizedText, {
                          children: "Current local result",
                        }),
                      }),
                      i.jsxs("dd", {
                        children: [
                          b.target.resultStatus === "unavailable"
                            ? "Unavailable"
                            : formatCellValueForDisplay(b.target.result),
                          b.target.resultNote &&
                            i.jsx("small", { children: b.target.resultNote }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              i.jsxs("div", {
                role: "status",
                className: b.coverage.complete
                  ? "fi-coverage"
                  : "fi-coverage fi-warning",
                children: [
                  i.jsxs("strong", {
                    children: [
                      b.nodes.length,
                      " ",
                      s,
                      " ",
                      i.jsx(LocalizedText, { children: "listed ·" }),
                      " ",
                      b.coverage.complete
                        ? "Complete within the selected static scope"
                        : "Partial static inspection",
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      b.coverage.formulaCellsScanned,
                      " ",
                      i.jsx(LocalizedText, { children: "of" }),
                      " ",
                      b.coverage.formulaCellsAvailable,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "eligible formula cells scanned; ",
                      }),
                      b.coverage.referencesScanned,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "references inspected. ",
                      }),
                      b.coverage.excludedCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children:
                          "populated cells excluded by visibility or host context.",
                      }),
                    ],
                  }),
                  b.coverage.notes.map((k) => i.jsx("p", { children: k }, k)),
                ],
              }),
              i.jsxs("p", {
                className: "fi-explanation",
                children: [
                  s === "precedents"
                    ? "These cells are referenced by the selected cell, directly or through other formulas."
                    : "These formula cells contain a potential reference path to the selected cell.",
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "Results are freshly evaluated by the local calculator; volatile functions can change between inspections.",
                  }),
                ],
              }),
              b.nodes.length
                ? i.jsx("div", {
                    className: "fi-table-wrap",
                    children: i.jsxs("table", {
                      "aria-label": `${s} cells`,
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
                                  children: "Depth",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Stored input",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(LocalizedText, {
                                  children: "Local result",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: b.nodes
                            .slice(E * 50, (E + 1) * 50)
                            .map((k) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("td", {
                                      children: i.jsxs("button", {
                                        onClick: () => A(k),
                                        "aria-label": `Go to ${k.sheetName}!${k.address}`,
                                        children: [k.sheetName, "!", k.address],
                                      }),
                                    }),
                                    i.jsx("td", { children: k.depth }),
                                    i.jsx("td", {
                                      children: i.jsx("code", {
                                        children: formatCellValueForDisplay(
                                          k.raw,
                                        ),
                                      }),
                                    }),
                                    i.jsx("td", {
                                      children:
                                        k.resultStatus === "unavailable"
                                          ? i.jsx("span", {
                                              title: k.resultNote,
                                              children: i.jsx(LocalizedText, {
                                                children: "Unavailable",
                                              }),
                                            })
                                          : formatCellValueForDisplay(k.result),
                                    }),
                                  ],
                                },
                                k.key,
                              ),
                            ),
                        }),
                      ],
                    }),
                  })
                : i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "No " }),
                      s,
                      " ",
                      i.jsx(LocalizedText, {
                        children:
                          "were found in the inspected subset. Review coverage and unresolved issues before concluding there are none.",
                      }),
                    ],
                  }),
              w > 1 &&
                i.jsxs("nav", {
                  className: "fi-pages",
                  "aria-label": text("Inspection result pages"),
                  lang: language("Inspection result pages"),
                  "data-ui-label": "Inspection result pages",
                  children: [
                    i.jsx("button", {
                      disabled: E === 0,
                      onClick: () => p(E - 1),
                      children: i.jsx(LocalizedText, {
                        children: "Previous results",
                      }),
                    }),
                    i.jsxs("span", {
                      children: [
                        i.jsx(LocalizedText, { children: "Page " }),
                        E + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "of " }),
                        w,
                      ],
                    }),
                    i.jsx("button", {
                      disabled: E + 1 >= w,
                      onClick: () => p(E + 1),
                      children: i.jsx(LocalizedText, {
                        children: "Next results",
                      }),
                    }),
                  ],
                }),
              b.cycles.length > 0 &&
                i.jsxs("section", {
                  className: "fi-warning",
                  "aria-label": text("Detected cycles"),
                  lang: language("Detected cycles"),
                  "data-ui-label": "Detected cycles",
                  children: [
                    i.jsx("h3", {
                      children: i.jsx(LocalizedText, {
                        children: "Detected static cycles",
                      }),
                    }),
                    i.jsx("ul", {
                      children: b.cycles.map((k, F) =>
                        i.jsx("li", { children: k.map(y).join(" → ") }, F),
                      ),
                    }),
                  ],
                }),
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, { children: "Reference ranges (" }),
                      b.references.length,
                      ")",
                    ],
                  }),
                  i.jsx("ul", {
                    children: b.references.map((k, F) => {
                      return i.jsxs(
                        "li",
                        {
                          children: [
                            i.jsx("code", { children: y(k.from) }),
                            " →",
                            " ",
                            e.workbook.sheets.find((O) => O.id === k.sheetId)
                              ?.name,
                            "!",
                            k.range,
                            " · ",
                            k.cellCount.toLocaleString(),
                            " ",
                            i.jsx(LocalizedText, { children: "cells ·" }),
                            " ",
                            k.kind,
                            " ·",
                            " ",
                            i.jsx("code", { children: k.token }),
                          ],
                        },
                        F,
                      );
                    }),
                  }),
                ],
              }),
              i.jsxs("details", {
                open: b.issues.length > 0,
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Unresolved or excluded dependencies (",
                      }),
                      b.issues.length,
                      ")",
                    ],
                  }),
                  b.issues.length
                    ? i.jsx("ul", {
                        children: b.issues.map((k, F) =>
                          i.jsxs(
                            "li",
                            {
                              children: [
                                i.jsx("strong", { children: k.kind }),
                                " · ",
                                y(k.from),
                                k.token &&
                                  i.jsxs(i.Fragment, {
                                    children: [
                                      " ",
                                      "· ",
                                      i.jsx("code", { children: k.token }),
                                    ],
                                  }),
                                " ",
                                "— ",
                                k.message,
                              ],
                            },
                            F,
                          ),
                        ),
                      })
                    : i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "No unresolved references were identified by the static scanner.",
                        }),
                      }),
                ],
              }),
            ],
          }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: i.jsx(LocalizedText, {
                children:
                  "Inspection does not change cells, history, or calculation settings.",
              }),
            }),
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function normalizeFormulaWatchContext(e, r) {
  if (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  if (
    r.allowedSheetIds !== undefined &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (t) => typeof t != "string" || !e.sheets.some((n) => n.id === t),
      ))
  )
    throw Error(
      "The formula watch host context contains an unknown worksheet.",
    );
  return {
    includeHidden: !!r.includeHidden,
    allowedSheetIds: [
      ...new Set(r.allowedSheetIds ?? e.sheets.map((t) => t.id)),
    ].sort(),
  };
}
function classifyWatchedCell(e, r, t) {
  const n = e.sheets.find((o) => o.id === r.anchor.sheetId);
  if (!n) return "orphaned";
  if (!t.allowedSheetIds.includes(n.id)) return "excluded";
  if (!t.includeHidden) {
    const o = _t(r.anchor.address);
    if (
      n.hidden ||
      n.veryHidden ||
      cn(n, "row").some((a) => a.start <= o.row && a.end >= o.row) ||
      cn(n, "column").some((a) => a.start <= o.col && a.end >= o.col)
    )
      return "hidden";
  }
  return r.orphaned ? "orphaned" : "ready";
}

export {
  normalizeFormulaWatchContext,
  classifyWatchedCell,
  FormulaCompatibilityDialog,
  CalculationInspectorDialog,
  EvaluateFormulaDialog,
};
