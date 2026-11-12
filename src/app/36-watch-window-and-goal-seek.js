import { isFiniteNumber } from "./03-workbook-model.js";
import { buildCompositeKey } from "./11-workbook-formats-dialog.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { fingerprintCanonicalValue } from "./01-core-schemas-and-hashing.js";
import { workbookLimits } from "./02-financial-functions.js";
import {
  B2,
  enumerateSelectionCells,
  validateFormulaWatches,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  validateDataValidationRule,
  resolveDefinedNameValue,
  columnIndexToLetters,
  matrixResultFunctionNames,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  formatCellEntryForEdit,
  validateLocalizedNumberEntry,
} from "./04-cell-input-paste-and-fill.js";
import { inspectFormulaDependencies } from "./12-formula-dependency-inspector.js";
import {
  normalizeFormulaWatchContext,
  classifyWatchedCell,
} from "./35-formula-compatibility-report.js";

function applyFormulaWatchAction(e, r, t = {}) {
  if (t.readOnly)
    throw Error(
      "This workbook is read-only. Watch inspection remains available.",
    );
  const n = normalizeFormulaWatchContext(e, t),
    o = validateFormulaWatches(e.formulaWatches, e);
  let a = o,
    s = [],
    d = [];
  if (r.kind === "add") {
    const l = e.sheets.find((f) => f.id === r.sheetId);
    if (!l)
      throw Error("The worksheet selected for watching no longer exists.");
    if (!Array.isArray(r.ranges))
      throw Error("Provide a list of selected ranges to watch.");
    const c = enumerateSelectionCells(r.ranges, B2),
      h = new Set(o.map((f) => f.id));
    for (const f of c) {
      if (
        o.some(
          (v) =>
            !v.orphaned && v.anchor.sheetId === l.id && v.anchor.address === f,
        )
      )
        continue;
      let p = 1;
      for (; h.has(`watch-${p}`);) p++;
      const m = {
        id: `watch-${p}`,
        anchor: { sheetId: l.id, sheetName: l.name, address: f },
      };
      if (classifyWatchedCell(e, m, n) !== "ready")
        throw Error(
          "A selected watch source is outside the visible, authorized context. Review visibility or host access before adding it.",
        );
      (h.add(m.id), s.push(m));
    }
    if (o.length + s.length > B2)
      throw Error(
        "A workbook may contain at most 100 formula watches, including orphaned entries. Remove a watch before adding another.",
      );
    a = [...o, ...s];
  } else if (r.kind === "remove") {
    if (
      !Array.isArray(r.ids) ||
      r.ids.length > B2 ||
      r.ids.some((c) => typeof c != "string" || !o.some((h) => h.id === c))
    )
      throw Error(
        "A selected watch no longer exists. Review the current watch list.",
      );
    const l = new Set(r.ids);
    ((d = o.filter((c) => l.has(c.id))), (a = o.filter((c) => !l.has(c.id))));
  } else if (r.kind === "clear") ((d = o), (a = []));
  else throw Error("Unknown formula watch action.");
  const u =
    s.length || d.length ? { ...e, formulaWatches: structuredClone(a) } : e;
  return {
    snapshot: JSON.stringify(e),
    request: structuredClone(r),
    context: n,
    added: s,
    removed: d,
    next: u,
  };
}
function assertWatchPreviewCurrent(e, r, t = {}) {
  if (JSON.stringify(e) !== r.snapshot)
    throw Error(
      "The workbook changed after the watch preview. Review the action again.",
    );
  const n = applyFormulaWatchAction(e, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error(
      "The watch preview or host context changed. Review the action again.",
    );
  return n.next;
}
function inspectFormulaWatches(e, r = {}) {
  const t = normalizeFormulaWatchContext(e, r),
    n = validateFormulaWatches(e.formulaWatches, e),
    o = [],
    a = n.map((w) => {
      const E = classifyWatchedCell(e, w, t);
      return E === "ready"
        ? (o.push(w), { id: w.id, status: "unavailable", anchor: w.anchor })
        : E === "orphaned"
          ? r.allowedSheetIds !== undefined &&
            !t.allowedSheetIds.includes(w.anchor.sheetId)
            ? {
                id: w.id,
                status: "excluded",
                note: "This watch is outside the current host context.",
              }
            : {
                id: w.id,
                status: "orphaned",
                anchor: w.anchor,
                note: `Source unavailable (${w.orphaned ?? "missing-sheet"}). This watch will not reconnect automatically.`,
              }
          : {
              id: w.id,
              status: E,
              note:
                E === "hidden"
                  ? "This source is hidden. Use Include hidden only within the host-authorized context."
                  : "This watch is outside the current host context.",
            };
    }),
    s = {
      workbookId: e.id,
      revision: e.revision,
      rows: a,
      formulaCellsScanned: 0,
      inspectionPasses: 0,
      notes: [
        "Watch results are current local calculations, not Excel step-by-step evaluation. Unresolved, hidden or out-of-context dependencies withhold the result.",
        "The embedding host must supply authorized workbook data. Include hidden is a local visibility choice.",
      ],
    };
  if (!o.length) return s;
  let d = "",
    u = "";
  for (let w = 0; w < 4; w++) {
    ((d = `_watch_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`),
      (u = `watch-inspection-${crypto.randomUUID()}`));
    const E = d.toLowerCase();
    if (
      !e.sheets.some(
        (A) =>
          A.id === u ||
          A.name.toLowerCase() === E ||
          Object.values(A.cells).some(
            (k) => typeof k.raw == "string" && k.raw.toLowerCase().includes(E),
          ),
      ) &&
      !e.definedNames?.some((A) => A.refersTo.toLowerCase().includes(E))
    )
      break;
    if (w === 3)
      throw Error(
        "A temporary watch inspection context could not be isolated.",
      );
  }
  const l = {
    id: u,
    name: d,
    rowCount: workbookLimits.maxRows,
    colCount: workbookLimits.maxCols,
    cells: Object.fromEntries(
      o.map((w, E) => [
        `A${E + 1}`,
        {
          raw: `='${w.anchor.sheetName.replace(/'/g, "''")}'!${w.anchor.address}`,
        },
      ]),
    ),
  };
  l.cells.B1 = { raw: `=SUM(A1:A${o.length})` };
  const c = inspectFormulaDependencies(
    {
      ...e,
      sheets: [
        l,
        ...e.sheets.map((w) => (w.veryHidden ? { ...w, hidden: true } : w)),
      ],
    },
    {
      sheetId: l.id,
      address: "B1",
      direction: "precedents",
      depth: 8,
      includeHidden: t.includeHidden,
      allowedSheetIds: [...t.allowedSheetIds, l.id],
      limits: {
        maxFormulaCells: 5e3 + o.length + 1,
        maxNodes: 1e3,
        maxReferences: 1e4,
      },
    },
  );
  ((s.inspectionPasses = 1),
    (s.formulaCellsScanned = Math.max(
      0,
      c.coverage.formulaCellsScanned - o.length - 1,
    )));
  const h = new Map(
      c.nodes.map((w) => [JSON.stringify([w.sheetId, w.address]), w]),
    ),
    f = new Map(
      e.sheets.map((w) => [
        w.id,
        Object.entries(w.cells)
          .filter(
            ([, E]) =>
              !E.literal && typeof E.raw == "string" && E.raw.startsWith("="),
          )
          .map(([E]) => _t(E)),
      ]),
    ),
    p = new Set();
  let m = 0,
    v = false;
  for (const [w, E] of h) {
    if (E.raw !== null) continue;
    const A = _t(E.address);
    for (const k of f.get(E.sheetId) ?? []) {
      if (++m > 1e5) {
        v = true;
        break;
      }
      if (
        k.row <= A.row &&
        k.col <= A.col &&
        (A.row - k.row + 1) * (A.col - k.col + 1) <= 1e4
      ) {
        p.add(w);
        break;
      }
    }
    if (v) break;
  }
  const g = new Map();
  for (const w of c.edges) g.set(w.from, [...(g.get(w.from) ?? []), w.to]);
  const b = (w, E = new Set()) =>
    p.has(w)
      ? true
      : E.has(w)
        ? false
        : (E.add(w), (g.get(w) ?? []).some((A) => b(A, E)));
  for (const w of s.rows) {
    if (w.status !== "unavailable" || !w.anchor) continue;
    const E = h.get(JSON.stringify([w.anchor.sheetId, w.anchor.address]));
    if (!E) {
      w.note =
        "Result withheld: this source is beyond the safe inspection budget.";
      continue;
    }
    if (
      ((w.raw = E.raw),
      E.resultStatus !== "unavailable" &&
        (b(JSON.stringify([w.anchor.sheetId, w.anchor.address])) ||
          ((v || c.coverage.truncated || c.coverage.depthLimited) &&
            (E.raw === null ||
              (typeof E.raw == "string" &&
                E.raw.startsWith("=") &&
                E.resultStatus === "calculated")))))
    ) {
      ((w.result = null),
        (w.resultStatus = "unavailable"),
        (w.status = "unavailable"),
        (w.note =
          "Result withheld: possible runtime spill output or an incomplete dependency inspection. Watch the stored source formula or narrow the workbook context."));
      continue;
    }
    ((w.result = E.result),
      (w.resultStatus = E.resultStatus),
      (w.status =
        E.resultStatus === "unavailable" ? "unavailable" : "available"),
      E.resultNote && (w.note = E.resultNote));
  }
  return (
    (p.size || v) &&
      s.notes.push(
        "Blank cells inside a possible formula output footprint are conservatively withheld, as are formulas depending on them. Runtime spill-child identification remains outside the supported watch scope.",
      ),
    c.coverage.truncated &&
      s.notes.push(
        "The shared dependency scan reached a limit. Some results may be withheld; no incomplete result is presented as complete.",
      ),
    s
  );
}
function navigateToFormulaWatch(e, r, t = {}) {
  const n = validateFormulaWatches(e.formulaWatches, e).find((o) => o.id === r);
  if (
    !n ||
    classifyWatchedCell(e, n, normalizeFormulaWatchContext(e, t)) !== "ready"
  )
    throw Error(
      "This watch cannot be navigated in the current visible, authorized context.",
    );
  return structuredClone(n.anchor);
}
const rj = (e) =>
  e === undefined
    ? "Unavailable"
    : e === null
      ? "(blank)"
      : typeof e == "boolean"
        ? String(e).toUpperCase()
        : String(e);
function FormulaWatchWindow(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const [s, d] = ne.useState(false),
    [u, l] = ne.useState([]),
    [c, h] = ne.useState(null),
    [f, p] = ne.useState(""),
    [m, v] = ne.useState(""),
    g = {
      includeHidden: !!e.allowHidden && s,
      allowedSheetIds: e.allowedSheetIds,
    },
    b = JSON.stringify([e.sheetId, e.ranges ?? [e.range]]),
    y = ne.useMemo(() => {
      try {
        return {
          result: inspectFormulaWatches(e.workbook, {
            includeHidden: !!e.allowHidden && s,
            allowedSheetIds: e.allowedSheetIds,
          }),
        };
      } catch (z) {
        return {
          error:
            z instanceof Error
              ? z.message
              : "The watch list could not be read.",
        };
      }
    }, [e.workbook, e.allowHidden, e.allowedSheetIds, s]);
  ne.useEffect(() => {
    const z = document.activeElement;
    return (
      o.current?.querySelector('[data-ui-label="Close Watch Window"]')?.focus(),
      () => {
        z?.isConnected && z.focus();
      }
    );
  }, []);
  const w = (z) =>
      p(
        z instanceof Error
          ? z.message
          : "The watch action could not be completed.",
      ),
    E = (z) => {
      try {
        const V = applyFormulaWatchAction(e.workbook, z, {
          ...g,
          readOnly: e.readOnly,
        });
        (h(V), v(b), p(""));
      } catch (V) {
        (h(null), w(V));
      }
    },
    A = () => {
      if (c)
        try {
          if (c.request.kind === "add" && m !== b)
            throw Error(
              "The worksheet selection changed after the watch preview. Review the selection again.",
            );
          const z = assertWatchPreviewCurrent(e.workbook, c, {
            ...g,
            readOnly: e.readOnly,
          });
          (z !== e.workbook &&
            e.onWorkbookChange(
              z,
              c.request.kind === "add"
                ? `Add ${c.added.length} formula watch${c.added.length === 1 ? "" : "es"}`
                : c.request.kind === "clear"
                  ? "Clear formula watches"
                  : `Remove ${c.removed.length} formula watch${c.removed.length === 1 ? "" : "es"}`,
            ),
            h(null),
            l([]),
            p(""));
        } catch (z) {
          w(z);
        }
    },
    k = (z) => {
      try {
        const V = navigateToFormulaWatch(e.workbook, z, g);
        (e.onNavigate(V.sheetId, V.address), p(""));
      } catch (V) {
        w(V);
      }
    },
    F = y.result?.rows ?? [],
    I =
      !!c &&
      c.snapshot === JSON.stringify(e.workbook) &&
      c.context.includeHidden === g.includeHidden &&
      JSON.stringify(c.context.allowedSheetIds) ===
        JSON.stringify(
          [
            ...new Set(g.allowedSheetIds ?? e.workbook.sheets.map((z) => z.id)),
          ].sort(),
        ) &&
      (c.request.kind !== "add" || m === b);
  return i.jsx("div", {
    className: "fw-backdrop",
    onMouseDown: (z) => {
      z.target === z.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: o,
      className: "fw-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (z) => {
        var D, B;
        if (
          (z.stopPropagation(),
          z.key === "Escape" && (z.preventDefault(), a.current()),
          z.key === "Tab")
        ) {
          const _ = [
            ...(o.current?.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
            ) ?? []),
          ].filter((C) => !C.hidden);
          z.shiftKey && document.activeElement === _[0]
            ? (z.preventDefault(), (D = _.at(-1)) == null || D.focus())
            : !z.shiftKey &&
              document.activeElement === _.at(-1) &&
              (z.preventDefault(), (B = _[0]) == null || B.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(LocalizedText, { children: "Watch Window" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Keep key cells and calculations in view across worksheets.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Watch Window"),
              onClick: e.onClose,
              lang: language("Close Watch Window"),
              "data-ui-label": "Close Watch Window",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fw-toolbar",
          children: [
            i.jsx("button", {
              disabled:
                e.readOnly ||
                !e.workbook.sheets.some((z) => z.id === e.sheetId),
              onClick: () =>
                E({
                  kind: "add",
                  sheetId: e.sheetId,
                  ranges: e.ranges ?? [e.range],
                }),
              children: i.jsx(LocalizedText, { children: "Add selection" }),
            }),
            i.jsx("button", {
              disabled: e.readOnly || !u.length,
              onClick: () => E({ kind: "remove", ids: u }),
              children: i.jsx(LocalizedText, { children: "Remove selected" }),
            }),
            i.jsx("button", {
              disabled: e.readOnly || !F.length,
              onClick: () => E({ kind: "clear" }),
              children: i.jsx(LocalizedText, { children: "Clear watches" }),
            }),
            e.allowHidden &&
              i.jsxs("label", {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: s,
                    onChange: (z) => {
                      (d(z.target.checked), h(null), p(""));
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden" }),
                ],
              }),
          ],
        }),
        i.jsxs("p", {
          className: "fw-context",
          children: [
            e.workbook.name,
            " ",
            i.jsx(LocalizedText, { children: "· Revision " }),
            e.workbook.revision,
            " · ",
            F.length,
            i.jsx(LocalizedText, { children: "/100 watches" }),
            e.readOnly ? " · Read-only inspection" : "",
          ],
        }),
        (y.error || f) &&
          i.jsx("p", {
            role: "alert",
            className: "fw-error",
            children: f || y.error,
          }),
        i.jsxs("div", {
          className: "fw-table-wrap",
          children: [
            i.jsxs("table", {
              children: [
                i.jsx("thead", {
                  children: i.jsxs("tr", {
                    children: [
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, { children: "Select" }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, {
                          children: "Worksheet / Cell",
                        }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, {
                          children: "Formula / Raw value",
                        }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, {
                          children: "Current result",
                        }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, { children: "Status" }),
                      }),
                    ],
                  }),
                }),
                i.jsx("tbody", {
                  children: F.map((z, V) =>
                    i.jsxs(
                      "tr",
                      {
                        children: [
                          i.jsx("td", {
                            children: i.jsx("input", {
                              type: "checkbox",
                              "aria-label": `Select watch ${V + 1}`,
                              checked: u.includes(z.id),
                              disabled: e.readOnly,
                              onChange: (D) =>
                                l((B) =>
                                  D.target.checked
                                    ? [...B, z.id]
                                    : B.filter((_) => _ !== z.id),
                                ),
                            }),
                          }),
                          i.jsx("td", {
                            children: z.anchor
                              ? i.jsxs("button", {
                                  className: "fw-link",
                                  disabled: z.status === "orphaned",
                                  onClick: () => k(z.id),
                                  "aria-label": `Go to ${z.anchor.sheetName}!${z.anchor.address}`,
                                  children: [
                                    z.anchor.sheetName,
                                    "!",
                                    z.anchor.address,
                                  ],
                                })
                              : i.jsxs("span", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Watch ",
                                    }),
                                    V + 1,
                                  ],
                                }),
                          }),
                          i.jsx("td", {
                            children: i.jsx("code", { children: rj(z.raw) }),
                          }),
                          i.jsx("td", {
                            children:
                              z.status === "available"
                                ? rj(z.result)
                                : "Unavailable",
                          }),
                          i.jsxs("td", {
                            children: [
                              i.jsx("strong", {
                                children:
                                  z.status === "available"
                                    ? z.resultStatus === "literal"
                                      ? "Value"
                                      : "Calculated"
                                    : z.status === "unavailable"
                                      ? "Result withheld"
                                      : z.status[0].toUpperCase() +
                                        z.status.slice(1),
                              }),
                              z.note && i.jsx("small", { children: z.note }),
                            ],
                          }),
                        ],
                      },
                      z.id,
                    ),
                  ),
                }),
              ],
            }),
            !F.length &&
              !y.error &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "No watched cells. Select cells in the worksheet and choose Add selection.",
                }),
              }),
          ],
        }),
        c &&
          i.jsxs("section", {
            className: "fw-preview",
            "aria-label": text("Watch action preview"),
            lang: language("Watch action preview"),
            "data-ui-label": "Watch action preview",
            children: [
              i.jsx("h3", {
                children:
                  c.request.kind === "add"
                    ? "Add watches"
                    : c.request.kind === "clear"
                      ? "Clear watch list"
                      : "Remove watches",
              }),
              i.jsxs("p", {
                children: [
                  c.added.length,
                  " ",
                  i.jsx(LocalizedText, { children: "added · " }),
                  c.removed.length,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "removed. Cell content and calculation formulas remain unchanged.",
                  }),
                ],
              }),
              !I &&
                i.jsx("p", {
                  role: "status",
                  children: i.jsx(LocalizedText, {
                    children:
                      "This preview is stale. Locations are withheld until you review the current selection and access context again.",
                  }),
                }),
              i.jsx("ul", {
                children: [
                  ...c.added.map((z) => ({ watch: z, action: "Add" })),
                  ...c.removed.map((z) => ({ watch: z, action: "Remove" })),
                ].map(({ watch, action }) => {
                  const D = F.find((_) => _.id === watch.id),
                    B =
                      I && (action === "Add" || D?.anchor)
                        ? `${watch.anchor.sheetName}!${watch.anchor.address}`
                        : "Watch outside current context";
                  return i.jsxs(
                    "li",
                    {
                      children: [
                        action,
                        ": ",
                        B,
                        watch.orphaned ? ` (${watch.orphaned})` : "",
                      ],
                    },
                    watch.id,
                  );
                }),
              }),
              !c.added.length &&
                !c.removed.length &&
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "The selected cells are already watched. No change is needed.",
                  }),
                }),
              i.jsx("button", {
                onClick: () => h(null),
                children: i.jsx(LocalizedText, {
                  children: "Cancel watch action",
                }),
              }),
              i.jsx("button", {
                className: "fw-primary",
                disabled: e.readOnly,
                onClick: A,
                children: i.jsx(LocalizedText, {
                  children: "Apply watch action",
                }),
              }),
            ],
          }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Calculation scope and limits",
              }),
            }),
            y.result?.notes.map((z) => i.jsx("p", { children: z }, z)),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Results share one bounded inspection pass across the watch list. Unsupported references or sources beyond the safe dependency budget display Result withheld. Orphaned entries retain their former source and never reconnect automatically.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: e.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const goalSeekLimits = {
    maxIterations: 100,
    maxEvaluations: 103,
    maxDependencyCells: 500,
    maxWorkUnits: 5e4,
    maxDepth: 8,
    maxFormulaCells: 5e3,
  },
  isCellInRange = (e, r) => {
    const t = parseRangeAddress(e),
      n = _t(r);
    return (
      n.row >= t.firstRow &&
      n.row <= t.lastRow &&
      n.col >= t.firstCol &&
      n.col <= t.lastCol
    );
  };
function parseGoalSeekAddress(e) {
  if (typeof e != "string" || !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}$/.test(e))
    throw Error(
      "Goal Seek needs one valid Excel cell address for each source.",
    );
  const r = _t(e.replace(/\$/g, "").toUpperCase());
  if (r.row >= workbookLimits.maxRows || r.col >= workbookLimits.maxCols)
    throw Error("A Goal Seek cell is outside Excel's address bounds.");
  return `${columnIndexToLetters(r.col)}${r.row + 1}`;
}
function normalizeGoalSeekRequest(e) {
  if (
    !e ||
    typeof e != "object" ||
    Object.keys(e).some(
      (t) =>
        ![
          "formulaSheetId",
          "formulaAddress",
          "variableSheetId",
          "variableAddress",
          "target",
          "lowerBound",
          "upperBound",
          "tolerance",
          "maxIterations",
        ].includes(t),
    )
  )
    throw Error("Invalid Goal Seek request.");
  if (
    typeof e.formulaSheetId != "string" ||
    typeof e.variableSheetId != "string"
  )
    throw Error("Choose a formula worksheet and variable worksheet.");
  const r = {
    ...e,
    formulaAddress: parseGoalSeekAddress(e.formulaAddress),
    variableAddress: parseGoalSeekAddress(e.variableAddress),
    tolerance: e.tolerance ?? 1e-6,
    maxIterations: e.maxIterations ?? 64,
  };
  if (
    !isFiniteNumber(r.target) ||
    !isFiniteNumber(r.lowerBound) ||
    !isFiniteNumber(r.upperBound) ||
    r.lowerBound >= r.upperBound
  )
    throw Error(
      "Target and bounds must be finite numbers, with lower bound smaller than upper bound.",
    );
  if (!isFiniteNumber(r.tolerance) || r.tolerance <= 0)
    throw Error("Tolerance must be a finite positive absolute error.");
  if (
    !Number.isInteger(r.maxIterations) ||
    r.maxIterations < 1 ||
    r.maxIterations > goalSeekLimits.maxIterations
  )
    throw Error("Goal Seek supports from 1 to 100 iterations.");
  return r;
}
function solveGoalSeek(e, r, t = {}) {
  if (t.readOnly)
    throw Error(
      "This workbook is read-only. Goal Seek cannot propose a writable input.",
    );
  const n = normalizeGoalSeekRequest(r),
    o = e.sheets.find((z) => z.id === n.formulaSheetId),
    a = e.sheets.find((z) => z.id === n.variableSheetId);
  if (!o || !a) throw Error("A selected Goal Seek worksheet no longer exists.");
  if (t.includeHidden !== undefined && typeof t.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  if (
    t.allowedSheetIds !== undefined &&
    (!Array.isArray(t.allowedSheetIds) ||
      t.allowedSheetIds.some(
        (z) => typeof z != "string" || !e.sheets.some((V) => V.id === z),
      ))
  )
    throw Error("Goal Seek host scope contains an unknown worksheet.");
  const s = {
    includeHidden: !!t.includeHidden,
    allowedSheetIds: [
      ...new Set(t.allowedSheetIds ?? e.sheets.map((z) => z.id)),
    ].sort(),
  };
  for (const [z, V] of [
    [o, n.formulaAddress],
    [a, n.variableAddress],
  ]) {
    const D = _t(V);
    if (
      !s.allowedSheetIds.includes(z.id) ||
      (!s.includeHidden &&
        (z.hidden ||
          z.veryHidden ||
          cn(z, "row").some((B) => B.start <= D.row && B.end >= D.row) ||
          cn(z, "column").some((B) => B.start <= D.col && B.end >= D.col)))
    )
      throw Error(
        "A Goal Seek source is outside the visible, authorized host context.",
      );
  }
  const d = o.cells[n.formulaAddress],
    u = a.cells[n.variableAddress];
  if (!d || d.literal || typeof d.raw != "string" || !d.raw.startsWith("="))
    throw Error("Set cell must contain an actual calculation formula.");
  if (!u || !isFiniteNumber(u.raw))
    throw Error(
      "Changing cell must contain a finite numeric input, not a formula, text, blank or spill output.",
    );
  if (a.protection?.enabled)
    throw Error(
      "The changing worksheet is protected. Unprotect it before Goal Seek.",
    );
  if (a.merges?.some((z) => isCellInRange(z, n.variableAddress)))
    throw Error("A merged cell cannot be the Goal Seek variable.");
  for (const z of a.tables ?? []) {
    if (!isCellInRange(z.range, n.variableAddress)) continue;
    const V = parseRangeAddress(z.range),
      D = _t(n.variableAddress),
      B = z.columns[D.col - V.firstCol],
      _ = B?.calculatedFormula;
    if (
      D.row === V.firstRow ||
      (z.showTotals && D.row === V.lastRow) ||
      _ ||
      Object.entries(a.cells).some(([C, T]) => {
        const N = _t(C);
        return (
          N.col === D.col &&
          N.row > V.firstRow &&
          N.row <= V.lastRow - (z.showTotals ? 1 : 0) &&
          !T.literal &&
          typeof T.raw == "string" &&
          T.raw.startsWith("=")
        );
      })
    )
      throw Error(
        "A table header, totals row or calculated column cannot be the Goal Seek variable.",
      );
  }
  if (
    [...(a.pivots ?? []), ...(a.queries ?? [])].some((z) =>
      isCellInRange(z.outputRange, n.variableAddress),
    )
  )
    throw Error(
      "A generated pivot/query output cannot be the Goal Seek variable.",
    );
  if (u.raw < n.lowerBound || u.raw > n.upperBound)
    throw Error("The starting input must be inside the supplied bounds.");
  const l = (a.validations ?? []).filter((z) =>
    isCellInRange(z.range, n.variableAddress),
  );
  if (l.some((z) => !["number", "date"].includes(z.kind)))
    throw Error(
      "Goal Seek supports continuous numeric/date validation only; discrete list, integer or text-length constraints need a different input.",
    );
  if (
    l.some(
      (z) =>
        ![u.raw, n.lowerBound, n.upperBound].every((V) =>
          validateDataValidationRule(z, V),
        ),
    )
  )
    throw Error(
      "The starting input and both bounds must satisfy the variable's data validation rules.",
    );
  const c = inspectFormulaDependencies(
      {
        ...e,
        sheets: e.sheets.map((z) =>
          z.veryHidden ? { ...z, hidden: true } : z,
        ),
      },
      {
        sheetId: o.id,
        address: n.formulaAddress,
        direction: "precedents",
        depth: goalSeekLimits.maxDepth,
        includeHidden: s.includeHidden,
        allowedSheetIds: s.allowedSheetIds,
        limits: {
          maxFormulaCells: goalSeekLimits.maxFormulaCells,
          maxNodes: goalSeekLimits.maxDependencyCells,
          maxReferences: 1e4,
        },
      },
    ),
    h = [c.target, ...c.nodes],
    f = new Set(h.map((z) => z.key));
  if (
    h.length > goalSeekLimits.maxDependencyCells ||
    c.coverage.truncated ||
    c.coverage.depthLimited
  )
    throw Error(
      "Goal Seek requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
    );
  const p = c.issues.find((z) => f.has(z.from));
  if (p || c.cycles.length)
    throw Error(
      `Goal Seek requires supported static dependencies: ${p?.message ?? "the dependency graph contains a cycle."}`,
    );
  if (!f.has(buildCompositeKey(a.id, n.variableAddress)))
    throw Error(
      "The Set cell formula does not depend on the selected changing cell.",
    );
  const m = new Set([
    ...matrixResultFunctionNames,
    "ISFORMULA",
    "INDEX",
    "CHOOSE",
    "INDIRECT",
    "OFFSET",
    "LET",
    "LAMBDA",
    "RAND",
    "RANDBETWEEN",
    "RANDARRAY",
    "NOW",
    "TODAY",
    "CELL",
    "INFO",
    "RTD",
    "WEBSERVICE",
    "STOCKHISTORY",
    "EVALUATE",
    "GET.CELL",
  ]);
  for (const z of h) {
    const V = e.sheets.find((B) => B.id === z.sheetId),
      D = V.cells[z.address];
    if (D?.raw === null || !D)
      throw Error(
        "Goal Seek requires populated static dependency cells; blank or possible spill-child dependencies are not supported.",
      );
    if (!D.literal && typeof D.raw == "string" && D.raw.startsWith("=")) {
      const B = resolveDefinedNameValue(D.raw, e, V.id).replace(
        /"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g,
        "",
      );
      if (
        [...B.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)]
          .map((C) =>
            C[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, ""),
          )
          .some((C) => m.has(C)) ||
        /[{}#]/.test(B)
      )
        throw Error(
          "Goal Seek does not run volatile, dynamic-reference, array/spill or unsupported reusable-function dependencies.",
        );
    }
  }
  const v = {
    ...e,
    sheets: e.sheets
      .filter((z) => s.allowedSheetIds.includes(z.id))
      .map((z) => ({
        ...z,
        cells: Object.fromEntries(
          Object.entries(z.cells).filter(([V]) =>
            f.has(buildCompositeKey(z.id, V)),
          ),
        ),
      })),
  };
  let g = 0,
    b = 0,
    y = 0;
  const w = { input: u.raw, output: null };
  let E = { input: u.raw, value: null, residual: null };
  const A = new Map();
  class k extends Error {
    constructor(V, D) {
      (super(D), (this.status = V));
    }
  }
  const F = (z) => {
      const V = A.get(z);
      if (V) return V;
      if (
        g >= goalSeekLimits.maxEvaluations ||
        y + h.length > goalSeekLimits.maxWorkUnits
      )
        throw new k(
          "work-limit",
          "The Goal Seek evaluation/work budget was reached.",
        );
      (g++, (y += h.length));
      const D = {
          ...v,
          sheets: v.sheets.map((N) =>
            N.id === a.id
              ? {
                  ...N,
                  cells: { ...N.cells, [n.variableAddress]: { ...u, raw: z } },
                }
              : N,
          ),
        },
        B = calculateWorkbookWithSettings(D),
        _ = B(o.id, n.formulaAddress);
      if (B.getOperationalErrors().length)
        throw new k(
          "calculation-error",
          `A trial encountered an operational calculation error (${B.getOperationalErrors().join(", ")}), even if a formula error handler concealed it. The baseline is unchanged.`,
        );
      if (!isFiniteNumber(_))
        throw new k(
          "calculation-error",
          `The formula returned ${String(_)} at input ${z}; a finite numeric result is required.`,
        );
      for (const N of h)
        if (B.getSpillRange(N.sheetId, N.address))
          throw new k(
            "calculation-error",
            "A trial produced an array/spill result. Goal Seek requires scalar dependencies.",
          );
      if (B.getOperationalErrors().length)
        throw new k(
          "calculation-error",
          "Dependency inspection encountered an operational calculation error; the baseline is unchanged.",
        );
      const C = _ - n.target;
      if (!isFiniteNumber(C))
        throw new k(
          "calculation-error",
          "The residual exceeded finite numeric precision.",
        );
      const T = { input: z, value: _, residual: C };
      return (
        A.set(z, T),
        (E.residual === null || Math.abs(C) < Math.abs(E.residual)) && (E = T),
        T
      );
    },
    I = (z, V) => {
      const D = z === "converged" || z === "already-satisfied";
      if (D && l.some((C) => !validateDataValidationRule(C, E.input)))
        return I(
          "validation-error",
          "The candidate violates the variable's data validation rule.",
        );
      const B =
          D && E.input !== u.raw
            ? [
                {
                  sheetId: a.id,
                  address: n.variableAddress,
                  cell: { ...u, raw: E.input },
                },
              ]
            : [],
        _ = B.length
          ? {
              ...e,
              sheets: e.sheets.map((C) =>
                C.id === a.id
                  ? {
                      ...C,
                      cells: { ...C.cells, [n.variableAddress]: B[0].cell },
                    }
                  : C,
              ),
            }
          : e;
      return {
        snapshot: JSON.stringify(e),
        request: n,
        context: s,
        baseline: w,
        dependencyCells: h.length,
        result: {
          status: z,
          converged: D,
          input: E.input,
          achievedValue: E.value,
          residual: E.residual,
          absoluteResidual: E.residual === null ? null : Math.abs(E.residual),
          iterations: b,
          evaluations: g,
          workUnits: y,
          algorithm: "Bracketed bisection",
          message: V,
        },
        patches: B,
        next: _,
      };
    };
  try {
    const z = F(u.raw);
    if (((w.output = z.value), Math.abs(z.residual) <= n.tolerance))
      return I(
        "already-satisfied",
        "The baseline already meets the requested tolerance. No input change is needed.",
      );
    let V = F(n.lowerBound);
    if (Math.abs(V.residual) <= n.tolerance)
      return I("converged", "The lower bound meets the requested tolerance.");
    let D = F(n.upperBound);
    if (Math.abs(D.residual) <= n.tolerance)
      return I("converged", "The upper bound meets the requested tolerance.");
    if (Math.sign(V.residual) === Math.sign(D.residual))
      return I(
        "not-bracketed",
        "The supplied bounds do not bracket a target crossing. A solution may exist elsewhere or inside a non-monotonic interval; choose a narrower valid bracket.",
      );
    for (b = 1; b <= n.maxIterations; b++) {
      const B = V.input / 2 + D.input / 2;
      if (B === V.input || B === D.input)
        return I(
          "precision-limit",
          "The bracket cannot be refined further at numeric precision; the residual still exceeds tolerance.",
        );
      const _ = F(B);
      if (Math.abs(_.residual) <= n.tolerance)
        return I(
          "converged",
          "The candidate meets the requested absolute residual tolerance.",
        );
      Math.sign(_.residual) === Math.sign(V.residual) ? (V = _) : (D = _);
    }
    return (
      (b = n.maxIterations),
      I(
        "iteration-limit",
        "The iteration limit was reached without meeting the residual tolerance. The baseline is unchanged.",
      )
    );
  } catch (z) {
    return I(
      z instanceof k ? z.status : "calculation-error",
      z instanceof Error
        ? z.message
        : "The trial calculation failed; the baseline is unchanged.",
    );
  }
}
function applyGoalSeekResult(e, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (JSON.stringify(e) !== r.snapshot)
    throw Error(
      "The workbook changed after the Goal Seek preview. Solve and review again.",
    );
  if (!r.result.converged)
    throw Error(
      "Only a candidate that meets the residual tolerance can be applied.",
    );
  const n = solveGoalSeek(e, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error(
      "The Goal Seek preview or host context changed. Solve and review again.",
    );
  return n.next;
}
function GoalSeekDialog(e) {
  const r = useWorkbookMessages(),
    [t] = ne.useState(r.entryLocale),
    n = (W, K) => validateLocalizedNumberEntry(W, K, t),
    { text, language } = useWorkbookMessages(),
    s = ne.useId(),
    d = ne.useRef(null),
    u = ne.useRef(e.onClose);
  u.current = e.onClose;
  const [l, c] = ne.useState(e.sheetId),
    [h, f] = ne.useState(e.address),
    [p, m] = ne.useState(e.sheetId),
    [v, g] = ne.useState(""),
    [b, y] = ne.useState(""),
    [w, E] = ne.useState("0"),
    [A, k] = ne.useState("100"),
    [F, I] = ne.useState(() => formatCellEntryForEdit({ raw: 1e-6 }, t)),
    [O, j] = ne.useState("64"),
    [z, V] = ne.useState(false),
    [D, B] = ne.useState(null),
    [_, C] = ne.useState(""),
    T = {
      includeHidden: !!e.allowHidden && z,
      allowedSheetIds: e.allowedSheetIds,
    },
    N = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((W) => W.id)),
    ].sort(),
    R = e.workbook.sheets.filter(
      (W) =>
        N.includes(W.id) && (T.includeHidden || (!W.hidden && !W.veryHidden)),
    ),
    L = R.some((W) => W.id === l),
    $ = R.some((W) => W.id === p),
    U =
      !!D &&
      D.snapshot === JSON.stringify(e.workbook) &&
      D.context.includeHidden === T.includeHidden &&
      JSON.stringify(D.context.allowedSheetIds) === JSON.stringify(N);
  ne.useEffect(() => {
    const W = document.activeElement;
    return (
      d.current?.querySelector('[data-ui-label="Set cell"]')?.focus(),
      () => {
        W?.isConnected && W.focus();
      }
    );
  }, []);
  const G = (W, K) => {
      (W(K), B(null), C(""));
    },
    H = (W) =>
      C(
        W instanceof Error
          ? W.message
          : "Goal Seek could not complete the operation.",
      ),
    Q = () => {
      try {
        const W = {
          formulaSheetId: l,
          formulaAddress: h,
          variableSheetId: p,
          variableAddress: v,
          target: n(b, "Target value"),
          lowerBound: n(w, "Lower bound"),
          upperBound: n(A, "Upper bound"),
          tolerance: n(F, "Tolerance"),
          maxIterations: n(O, "Maximum iterations"),
        };
        (B(solveGoalSeek(e.workbook, W, { ...T, readOnly: e.readOnly })),
          C(""));
      } catch (W) {
        (B(null), H(W));
      }
    },
    X$1 = () => {
      if (D)
        try {
          const W = applyGoalSeekResult(e.workbook, D, {
            ...T,
            readOnly: e.readOnly,
          });
          (W !== e.workbook &&
            e.onWorkbookChange(
              W,
              `Goal Seek: ${D.request.variableAddress} → ${D.result.input}`,
            ),
            B(null),
            C(""));
        } catch (W) {
          H(W);
        }
    };
  return i.jsx("div", {
    className: "gs-backdrop",
    onMouseDown: (W) => {
      W.target === W.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: d,
      className: "gs-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${s}-title`,
      onKeyDown: (W) => {
        var te, ae;
        if (
          (W.stopPropagation(),
          W.key === "Escape" && (W.preventDefault(), u.current()),
          W.key === "Tab")
        ) {
          const oe = [
            ...(d.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),summary",
            ) ?? []),
          ].filter((de) => !de.hidden);
          W.shiftKey && document.activeElement === oe[0]
            ? (W.preventDefault(), (te = oe.at(-1)) == null || te.focus())
            : !W.shiftKey &&
              document.activeElement === oe.at(-1) &&
              (W.preventDefault(), (ae = oe[0]) == null || ae.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${s}-title`,
                  children: i.jsx(LocalizedText, { children: "Goal Seek" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Find one numeric input that brings a formula to a target value.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Goal Seek"),
              onClick: e.onClose,
              lang: language("Close Goal Seek"),
              "data-ui-label": "Close Goal Seek",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Enter a lower and upper bound around the starting input. Solve previews a candidate; Apply changes the workbook only after review.",
          }),
        }),
        i.jsxs("div", {
          className: "gs-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Formula worksheet" }),
                i.jsxs("select", {
                  "aria-label": text("Formula worksheet"),
                  value: L ? l : "",
                  disabled: e.readOnly,
                  onChange: (W) => G(c, W.target.value),
                  lang: language("Formula worksheet"),
                  "data-ui-label": "Formula worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    R.map((W) =>
                      i.jsx("option", { value: W.id, children: W.name }, W.id),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Set cell" }),
                i.jsx("input", {
                  "aria-label": text("Set cell"),
                  value: L ? h : "",
                  placeholder: text("B1"),
                  readOnly: e.readOnly,
                  onChange: (W) => G(f, W.target.value),
                  lang: language("Set cell"),
                  "data-ui-label": "Set cell",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Changing worksheet" }),
                i.jsxs("select", {
                  "aria-label": text("Changing worksheet"),
                  value: $ ? p : "",
                  disabled: e.readOnly,
                  onChange: (W) => G(m, W.target.value),
                  lang: language("Changing worksheet"),
                  "data-ui-label": "Changing worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    R.map((W) =>
                      i.jsx("option", { value: W.id, children: W.name }, W.id),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "By changing cell" }),
                i.jsx("input", {
                  "aria-label": text("By changing cell"),
                  value: $ ? v : "",
                  placeholder: text("A1"),
                  readOnly: e.readOnly,
                  onChange: (W) => G(g, W.target.value),
                  lang: language("By changing cell"),
                  "data-ui-label": "By changing cell",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "To value" }),
                i.jsx("input", {
                  "aria-label": text("Target value"),
                  value: b,
                  inputMode: "decimal",
                  readOnly: e.readOnly,
                  onChange: (W) => G(y, W.target.value),
                  lang: language("Target value"),
                  "data-ui-label": "Target value",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Absolute tolerance" }),
                i.jsx("input", {
                  "aria-label": text("Absolute tolerance"),
                  value: F,
                  inputMode: "decimal",
                  readOnly: e.readOnly,
                  onChange: (W) => G(I, W.target.value),
                  lang: language("Absolute tolerance"),
                  "data-ui-label": "Absolute tolerance",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Lower bound" }),
                i.jsx("input", {
                  "aria-label": text("Lower bound"),
                  value: w,
                  inputMode: "decimal",
                  readOnly: e.readOnly,
                  onChange: (W) => G(E, W.target.value),
                  lang: language("Lower bound"),
                  "data-ui-label": "Lower bound",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Upper bound" }),
                i.jsx("input", {
                  "aria-label": text("Upper bound"),
                  value: A,
                  inputMode: "decimal",
                  readOnly: e.readOnly,
                  onChange: (W) => G(k, W.target.value),
                  lang: language("Upper bound"),
                  "data-ui-label": "Upper bound",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Maximum iterations" }),
                i.jsx("input", {
                  "aria-label": text("Maximum iterations"),
                  value: O,
                  inputMode: "numeric",
                  readOnly: e.readOnly,
                  onChange: (W) => G(j, W.target.value),
                  lang: language("Maximum iterations"),
                  "data-ui-label": "Maximum iterations",
                }),
              ],
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "gs-checkbox",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: z,
                    onChange: (W) => {
                      (V(W.target.checked), B(null), C(""));
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden sources" }),
                ],
              }),
          ],
        }),
        _ && i.jsx("p", { role: "alert", className: "gs-error", children: _ }),
        i.jsx("button", {
          className: "gs-primary",
          disabled: e.readOnly || !L || !$,
          onClick: Q,
          children: i.jsx(LocalizedText, { children: "Solve" }),
        }),
        D &&
          i.jsxs("section", {
            className: "gs-preview",
            "aria-label": text("Goal Seek result"),
            lang: language("Goal Seek result"),
            "data-ui-label": "Goal Seek result",
            children: [
              U
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("h3", {
                        children: D.result.converged
                          ? D.result.status === "already-satisfied"
                            ? "Target already satisfied"
                            : "Candidate found"
                          : "No applicable solution",
                      }),
                      i.jsx("p", {
                        role: "status",
                        children: D.result.message,
                      }),
                      i.jsxs("dl", {
                        children: [
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Formula",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              e.workbook.sheets.find(
                                (W) => W.id === D.request.formulaSheetId,
                              )?.name,
                              "!",
                              D.request.formulaAddress,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Changing input",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              e.workbook.sheets.find(
                                (W) => W.id === D.request.variableSheetId,
                              )?.name,
                              "!",
                              D.request.variableAddress,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Baseline input",
                            }),
                          }),
                          i.jsx("dd", { children: D.baseline.input }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Baseline result",
                            }),
                          }),
                          i.jsx("dd", {
                            children: D.baseline.output ?? "Unavailable",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Target",
                            }),
                          }),
                          i.jsx("dd", { children: D.request.target }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Candidate input",
                            }),
                          }),
                          i.jsx("dd", { children: D.result.input }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Achieved value",
                            }),
                          }),
                          i.jsx("dd", {
                            children: D.result.achievedValue ?? "Unavailable",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Residual (result − target)",
                            }),
                          }),
                          i.jsx("dd", {
                            children: D.result.residual ?? "Unavailable",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Absolute tolerance",
                            }),
                          }),
                          i.jsx("dd", { children: D.request.tolerance }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Iterations / Evaluations",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              D.result.iterations,
                              " / ",
                              D.result.evaluations,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Method",
                            }),
                          }),
                          i.jsx("dd", { children: D.result.algorithm }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Search bounds",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              D.request.lowerBound,
                              " ",
                              i.jsx(LocalizedText, { children: "to " }),
                              D.request.upperBound,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Dependency cells / Work units",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              D.dependencyCells,
                              " / ",
                              D.result.workUnits,
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "The baseline workbook has not changed.",
                          }),
                          !D.result.converged &&
                            " Choose different bounds or settings and solve again; a failed candidate cannot be applied.",
                        ],
                      }),
                    ],
                  })
                : i.jsx("p", {
                    role: "status",
                    children: i.jsx(LocalizedText, {
                      children:
                        "This preview is stale. Its source values and locations are withheld. Solve again using the current workbook and access context.",
                    }),
                  }),
              i.jsx("button", {
                onClick: () => B(null),
                children: i.jsx(LocalizedText, { children: "Discard result" }),
              }),
              i.jsx("button", {
                className: "gs-primary",
                disabled: e.readOnly || !D.result.converged,
                onClick: X$1,
                children: i.jsx(LocalizedText, { children: "Apply result" }),
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
                  "One continuous numeric input and one scalar formula, using deterministic supported dependencies. Volatile/dynamic references, array/spill outputs, blank dependency cells, discrete validation constraints and calculated table inputs are excluded. A valid bracket is required; this search does not find every possible root or prove a unique solution.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "At most 100 iterations, 103 trial evaluations, 500 dependency cells through depth 8, and 50,000 dependency-cell work units. Absolute residual tolerance determines success; a narrow bracket alone does not.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Include hidden is a local visibility choice within the host-authorized workbook. The embedding host owns approval, permissions, Undo and persistence.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: e.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
function normalizeAnalysisPermissions(e, r = {}) {
  if (
    (r.readOnly !== undefined && typeof r.readOnly != "boolean") ||
    (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
  )
    throw Error("Analysis permissions must be boolean.");
  if (
    r.allowedSheetIds !== undefined &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (t) => typeof t != "string" || !e.sheets.some((n) => n.id === t),
      ))
  )
    throw Error("The host scope contains an unknown worksheet.");
  return {
    readOnly: !!r.readOnly,
    includeHidden: !!r.includeHidden,
    allowedSheetIds: [
      ...new Set(r.allowedSheetIds ?? e.sheets.map((t) => t.id)),
    ].sort(),
  };
}
function validateAnalysisSource(e, r, t) {
  const n = normalizeAnalysisPermissions(e, t),
    o = e.sheets.find((d) => d.id === r.sheetId);
  if (!o || !n.allowedSheetIds.includes(o.id))
    throw Error("A source is outside the authorized host context.");
  if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(r.address))
    throw Error("Use a canonical cell address.");
  const a = _t(r.address);
  if (a.row >= 1048576 || a.col >= 16384)
    throw Error("The cell address exceeds Excel dimensions.");
  if (
    !n.includeHidden &&
    (o.hidden ||
      o.veryHidden ||
      cn(o, "row").some((d) => d.start <= a.row && a.row <= d.end) ||
      cn(o, "column").some((d) => d.start <= a.col && a.col <= d.end))
  )
    throw Error(
      "A source is hidden. Explicitly include hidden cells within the authorized scope.",
    );
  if (o.analysisReport || o.dataModelReport)
    throw Error(
      "Analysis and model report snapshots cannot be analysis inputs. Use the original source cells.",
    );
  if (o.merges?.some((d) => parseRangeBounds(d, r.address)))
    throw Error("Merged analysis sources are ambiguous.");
  return o;
}
function parseRangeBounds(e, r) {
  const t = parseRangeAddress(e),
    n = _t(r);
  return (
    n.row >= t.firstRow &&
    n.row <= t.lastRow &&
    n.col >= t.firstCol &&
    n.col <= t.lastCol
  );
}
function p_(e) {
  return fingerprintCanonicalValue({
    sheets: e.sheets.filter((r) => !r.analysisReport).map((r) => ({ ...r })),
    calculation: e.calculation ?? null,
    datePolicy: e.datePolicy ?? null,
    definedNames: e.definedNames ?? null,
  });
}
function assertAutomaticCalculation(e) {
  if (e.calculation?.mode === "manual" || e.calculation?.iterative)
    throw Error(
      "Analysis formulas require automatic, non-iterative calculation.",
    );
}
const rejectRangeAsScalar = (e) => {
    if (Array.isArray(e))
      throw Error("Use SUM, AVERAGE or SUMPRODUCT to reduce a range.");
    return e;
  },
  assertSolverMagnitude = (e) => {
    if (!Number.isFinite(e) || Math.abs(e) > 1e100)
      throw Error("Solver arithmetic exceeds the supported finite magnitude.");
    return e;
  };

export {
  assertSolverMagnitude,
  parseRangeBounds,
  FormulaWatchWindow,
  assertAutomaticCalculation,
  normalizeAnalysisPermissions,
  rejectRangeAsScalar,
  GoalSeekDialog,
  p_,
  validateAnalysisSource,
};
export { isCellInRange };
