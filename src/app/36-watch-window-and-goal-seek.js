import { isFiniteNumber } from "./03-workbook-model.js";
import { buildCompositeKey } from "./11-local-workbook-store-and-versions.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
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
import { inspectFormulaDependencies } from "./12-consolidation-and-dependency-inspection.js";
import {
  normalizeFormulaWatchContext,
  classifyWatchedCell,
} from "./35-formula-inspection-and-compatibility.js";

function applyFormulaWatchAction(workbook, request, t = {}) {
  if (t.readOnly)
    throw Error(
      "This workbook is read-only. Watch inspection remains available.",
    );
  const formulaWatchContext = normalizeFormulaWatchContext(workbook, t),
    o = validateFormulaWatches(workbook.formulaWatches, workbook);
  let a = o,
    added = [],
    removed = [];
  if (request.kind === "add") {
    const sheet = workbook.sheets.find((sheet) => sheet.id === request.sheetId);
    if (!sheet)
      throw Error("The worksheet selected for watching no longer exists.");
    if (!Array.isArray(request.ranges))
      throw Error("Provide a list of selected ranges to watch.");
    const c = enumerateSelectionCells(request.ranges, B2),
      h = new Set(o.map((f) => f.id));
    for (const address of c) {
      if (
        o.some(
          (v) =>
            !v.orphaned &&
            v.anchor.sheetId === sheet.id &&
            v.anchor.address === address,
        )
      )
        continue;
      let p = 1;
      for (; h.has(`watch-${p}`);) p++;
      const m = {
        id: `watch-${p}`,
        anchor: { sheetId: sheet.id, sheetName: sheet.name, address: address },
      };
      if (classifyWatchedCell(workbook, m, formulaWatchContext) !== "ready")
        throw Error(
          "A selected watch source is outside the visible, authorized context. Review visibility or host access before adding it.",
        );
      (h.add(m.id), added.push(m));
    }
    if (o.length + added.length > B2)
      throw Error(
        "A workbook may contain at most 100 formula watches, including orphaned entries. Remove a watch before adding another.",
      );
    a = [...o, ...added];
  } else if (request.kind === "remove") {
    if (
      !Array.isArray(request.ids) ||
      request.ids.length > B2 ||
      request.ids.some(
        (c) => typeof c != "string" || !o.some((h) => h.id === c),
      )
    )
      throw Error(
        "A selected watch no longer exists. Review the current watch list.",
      );
    const l = new Set(request.ids);
    ((removed = o.filter((c) => l.has(c.id))),
      (a = o.filter((c) => !l.has(c.id))));
  } else if (request.kind === "clear") ((removed = o), (a = []));
  else throw Error("Unknown formula watch action.");
  const next =
    added.length || removed.length
      ? { ...workbook, formulaWatches: structuredClone(a) }
      : workbook;
  return {
    snapshot: JSON.stringify(workbook),
    request: structuredClone(request),
    context: formulaWatchContext,
    added: added,
    removed: removed,
    next: next,
  };
}
function assertWatchPreviewCurrent(workbook, r, t = {}) {
  if (JSON.stringify(workbook) !== r.snapshot)
    throw Error(
      "The workbook changed after the watch preview. Review the action again.",
    );
  const n = applyFormulaWatchAction(workbook, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error(
      "The watch preview or host context changed. Review the action again.",
    );
  return n.next;
}
function inspectFormulaWatches(workbook, r = {}) {
  const formulaWatchContext = normalizeFormulaWatchContext(workbook, r),
    n = validateFormulaWatches(workbook.formulaWatches, workbook),
    o = [],
    rows = n.map((w) => {
      const E = classifyWatchedCell(workbook, w, formulaWatchContext);
      return E === "ready"
        ? (o.push(w), { id: w.id, status: "unavailable", anchor: w.anchor })
        : E === "orphaned"
          ? r.allowedSheetIds !== undefined &&
            !formulaWatchContext.allowedSheetIds.includes(w.anchor.sheetId)
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
      workbookId: workbook.id,
      revision: workbook.revision,
      rows: rows,
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
      !workbook.sheets.some(
        (sheet) =>
          sheet.id === u ||
          sheet.name.toLowerCase() === E ||
          Object.values(sheet.cells).some(
            (cell) =>
              typeof cell.raw == "string" && cell.raw.toLowerCase().includes(E),
          ),
      ) &&
      !workbook.definedNames?.some((definedName) =>
        definedName.refersTo.toLowerCase().includes(E),
      )
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
      ...workbook,
      sheets: [
        l,
        ...workbook.sheets.map((sheet) =>
          sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
        ),
      ],
    },
    {
      sheetId: l.id,
      address: "B1",
      direction: "precedents",
      depth: 8,
      includeHidden: formulaWatchContext.includeHidden,
      allowedSheetIds: [...formulaWatchContext.allowedSheetIds, l.id],
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
      c.nodes.map((node) => [
        JSON.stringify([node.sheetId, node.address]),
        node,
      ]),
    ),
    f = new Map(
      workbook.sheets.map((sheet) => [
        sheet.id,
        Object.entries(sheet.cells)
          .filter(
            ([, cell]) =>
              !cell.literal &&
              typeof cell.raw == "string" &&
              cell.raw.startsWith("="),
          )
          .map(([E]) => _t(E)),
      ]),
    ),
    p = new Set();
  let m = 0,
    v = false;
  for (const [w, cell] of h) {
    if (cell.raw !== null) continue;
    const A = _t(cell.address);
    for (const k of f.get(cell.sheetId) ?? []) {
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
  for (const edge of c.edges)
    g.set(edge.from, [...(g.get(edge.from) ?? []), edge.to]);
  const b = (w, E = new Set()) =>
    p.has(w)
      ? true
      : E.has(w)
        ? false
        : (E.add(w), (g.get(w) ?? []).some((A) => b(A, E)));
  for (const row of s.rows) {
    if (row.status !== "unavailable" || !row.anchor) continue;
    const cell = h.get(
      JSON.stringify([row.anchor.sheetId, row.anchor.address]),
    );
    if (!cell) {
      row.note =
        "Result withheld: this source is beyond the safe inspection budget.";
      continue;
    }
    if (
      ((row.raw = cell.raw),
      cell.resultStatus !== "unavailable" &&
        (b(JSON.stringify([row.anchor.sheetId, row.anchor.address])) ||
          ((v || c.coverage.truncated || c.coverage.depthLimited) &&
            (cell.raw === null ||
              (typeof cell.raw == "string" &&
                cell.raw.startsWith("=") &&
                cell.resultStatus === "calculated")))))
    ) {
      ((row.result = null),
        (row.resultStatus = "unavailable"),
        (row.status = "unavailable"),
        (row.note =
          "Result withheld: possible runtime spill output or an incomplete dependency inspection. Watch the stored source formula or narrow the workbook context."));
      continue;
    }
    ((row.result = cell.result),
      (row.resultStatus = cell.resultStatus),
      (row.status =
        cell.resultStatus === "unavailable" ? "unavailable" : "available"),
      cell.resultNote && (row.note = cell.resultNote));
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
function navigateToFormulaWatch(workbook, r, t = {}) {
  const n = validateFormulaWatches(workbook.formulaWatches, workbook).find(
    (o) => o.id === r,
  );
  if (
    !n ||
    classifyWatchedCell(
      workbook,
      n,
      normalizeFormulaWatchContext(workbook, t),
    ) !== "ready"
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
  const [checked, d] = ne.useState(false),
    [ids, l] = ne.useState([]),
    [c, h] = ne.useState(null),
    [f, p] = ne.useState(""),
    [m, v] = ne.useState(""),
    g = {
      includeHidden: !!e.allowHidden && checked,
      allowedSheetIds: e.allowedSheetIds,
    },
    b = JSON.stringify([e.sheetId, e.ranges ?? [e.range]]),
    y = ne.useMemo(() => {
      try {
        return {
          result: inspectFormulaWatches(e.workbook, {
            includeHidden: !!e.allowHidden && checked,
            allowedSheetIds: e.allowedSheetIds,
          }),
        };
      } catch (error) {
        return {
          error:
            error instanceof Error
              ? error.message
              : "The watch list could not be read.",
        };
      }
    }, [e.workbook, e.allowHidden, e.allowedSheetIds, checked]);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.querySelector('[data-ui-label="Close Watch Window"]')?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const w = (z) =>
      p(
        z instanceof Error
          ? z.message
          : "The watch action could not be completed.",
      ),
    E = (request) => {
      try {
        const V = applyFormulaWatchAction(e.workbook, request, {
          ...g,
          readOnly: e.readOnly,
        });
        (h(V), v(b), p(""));
      } catch (error) {
        (h(null), w(error));
      }
    },
    onClick = () => {
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
        } catch (error) {
          w(error);
        }
    },
    k = (z) => {
      try {
        const V = navigateToFormulaWatch(e.workbook, z, g);
        (e.onNavigate(V.sheetId, V.address), p(""));
      } catch (error) {
        w(error);
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
            ...new Set(
              g.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
            ),
          ].sort(),
        ) &&
      (c.request.kind !== "add" || m === b);
  return i.jsx("div", {
    className: "fw-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: o,
      className: "fw-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (event) => {
        var D, B;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), a.current()),
          event.key === "Tab")
        ) {
          const _ = [
            ...(o.current?.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
            ) ?? []),
          ].filter((C) => !C.hidden);
          event.shiftKey && document.activeElement === _[0]
            ? (event.preventDefault(), (D = _.at(-1)) == null || D.focus())
            : !event.shiftKey &&
              document.activeElement === _.at(-1) &&
              (event.preventDefault(), (B = _[0]) == null || B.focus());
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
                !e.workbook.sheets.some((sheet) => sheet.id === e.sheetId),
              onClick: () =>
                E({
                  kind: "add",
                  sheetId: e.sheetId,
                  ranges: e.ranges ?? [e.range],
                }),
              children: i.jsx(LocalizedText, { children: "Add selection" }),
            }),
            i.jsx("button", {
              disabled: e.readOnly || !ids.length,
              onClick: () => E({ kind: "remove", ids: ids }),
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
                    checked: checked,
                    onChange: (event) => {
                      (d(event.target.checked), h(null), p(""));
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
                  children: F.map((cell, V) =>
                    i.jsxs(
                      "tr",
                      {
                        children: [
                          i.jsx("td", {
                            children: i.jsx("input", {
                              type: "checkbox",
                              "aria-label": `Select watch ${V + 1}`,
                              checked: ids.includes(cell.id),
                              disabled: e.readOnly,
                              onChange: (event) =>
                                l((B) =>
                                  event.target.checked
                                    ? [...B, cell.id]
                                    : B.filter((_) => _ !== cell.id),
                                ),
                            }),
                          }),
                          i.jsx("td", {
                            children: cell.anchor
                              ? i.jsxs("button", {
                                  className: "fw-link",
                                  disabled: cell.status === "orphaned",
                                  onClick: () => k(cell.id),
                                  "aria-label": `Go to ${cell.anchor.sheetName}!${cell.anchor.address}`,
                                  children: [
                                    cell.anchor.sheetName,
                                    "!",
                                    cell.anchor.address,
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
                            children: i.jsx("code", { children: rj(cell.raw) }),
                          }),
                          i.jsx("td", {
                            children:
                              cell.status === "available"
                                ? rj(cell.result)
                                : "Unavailable",
                          }),
                          i.jsxs("td", {
                            children: [
                              i.jsx("strong", {
                                children:
                                  cell.status === "available"
                                    ? cell.resultStatus === "literal"
                                      ? "Value"
                                      : "Calculated"
                                    : cell.status === "unavailable"
                                      ? "Result withheld"
                                      : cell.status[0].toUpperCase() +
                                        cell.status.slice(1),
                              }),
                              cell.note &&
                                i.jsx("small", { children: cell.note }),
                            ],
                          }),
                        ],
                      },
                      cell.id,
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
                  ...c.added.map((watch) => ({ watch: watch, action: "Add" })),
                  ...c.removed.map((watch) => ({
                    watch: watch,
                    action: "Remove",
                  })),
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
                onClick: onClick,
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
            y.result?.notes.map((note) => i.jsx("p", { children: note }, note)),
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
    const rangeAddress = parseRangeAddress(e),
      n = _t(r);
    return (
      n.row >= rangeAddress.firstRow &&
      n.row <= rangeAddress.lastRow &&
      n.col >= rangeAddress.firstCol &&
      n.col <= rangeAddress.lastCol
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
function normalizeGoalSeekRequest(request) {
  if (
    !request ||
    typeof request != "object" ||
    Object.keys(request).some(
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
    typeof request.formulaSheetId != "string" ||
    typeof request.variableSheetId != "string"
  )
    throw Error("Choose a formula worksheet and variable worksheet.");
  const r = {
    ...request,
    formulaAddress: parseGoalSeekAddress(request.formulaAddress),
    variableAddress: parseGoalSeekAddress(request.variableAddress),
    tolerance: request.tolerance ?? 1e-6,
    maxIterations: request.maxIterations ?? 64,
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
function solveGoalSeek(workbook, request, t = {}) {
  if (t.readOnly)
    throw Error(
      "This workbook is read-only. Goal Seek cannot propose a writable input.",
    );
  const goalSeekRequest = normalizeGoalSeekRequest(request),
    sheet = workbook.sheets.find(
      (sheet) => sheet.id === goalSeekRequest.formulaSheetId,
    ),
    a = workbook.sheets.find(
      (sheet) => sheet.id === goalSeekRequest.variableSheetId,
    );
  if (!sheet || !a)
    throw Error("A selected Goal Seek worksheet no longer exists.");
  if (t.includeHidden !== undefined && typeof t.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  if (
    t.allowedSheetIds !== undefined &&
    (!Array.isArray(t.allowedSheetIds) ||
      t.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("Goal Seek host scope contains an unknown worksheet.");
  const context = {
    includeHidden: !!t.includeHidden,
    allowedSheetIds: [
      ...new Set(t.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id)),
    ].sort(),
  };
  for (const [z, V] of [
    [sheet, goalSeekRequest.formulaAddress],
    [a, goalSeekRequest.variableAddress],
  ]) {
    const D = _t(V);
    if (
      !context.allowedSheetIds.includes(z.id) ||
      (!context.includeHidden &&
        (z.hidden ||
          z.veryHidden ||
          cn(z, "row").some((B) => B.start <= D.row && B.end >= D.row) ||
          cn(z, "column").some((B) => B.start <= D.col && B.end >= D.col)))
    )
      throw Error(
        "A Goal Seek source is outside the visible, authorized host context.",
      );
  }
  const cell = sheet.cells[goalSeekRequest.formulaAddress],
    u = a.cells[goalSeekRequest.variableAddress];
  if (
    !cell ||
    cell.literal ||
    typeof cell.raw != "string" ||
    !cell.raw.startsWith("=")
  )
    throw Error("Set cell must contain an actual calculation formula.");
  if (!u || !isFiniteNumber(u.raw))
    throw Error(
      "Changing cell must contain a finite numeric input, not a formula, text, blank or spill output.",
    );
  if (a.protection?.enabled)
    throw Error(
      "The changing worksheet is protected. Unprotect it before Goal Seek.",
    );
  if (
    a.merges?.some((merge) =>
      isCellInRange(merge, goalSeekRequest.variableAddress),
    )
  )
    throw Error("A merged cell cannot be the Goal Seek variable.");
  for (const z of a.tables ?? []) {
    if (!isCellInRange(z.range, goalSeekRequest.variableAddress)) continue;
    const rangeAddress = parseRangeAddress(z.range),
      D = _t(goalSeekRequest.variableAddress),
      B = z.columns[D.col - rangeAddress.firstCol],
      _ = B?.calculatedFormula;
    if (
      D.row === rangeAddress.firstRow ||
      (z.showTotals && D.row === rangeAddress.lastRow) ||
      _ ||
      Object.entries(a.cells).some(([C, cell]) => {
        const N = _t(C);
        return (
          N.col === D.col &&
          N.row > rangeAddress.firstRow &&
          N.row <= rangeAddress.lastRow - (z.showTotals ? 1 : 0) &&
          !cell.literal &&
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=")
        );
      })
    )
      throw Error(
        "A table header, totals row or calculated column cannot be the Goal Seek variable.",
      );
  }
  if (
    [...(a.pivots ?? []), ...(a.queries ?? [])].some((z) =>
      isCellInRange(z.outputRange, goalSeekRequest.variableAddress),
    )
  )
    throw Error(
      "A generated pivot/query output cannot be the Goal Seek variable.",
    );
  if (u.raw < goalSeekRequest.lowerBound || u.raw > goalSeekRequest.upperBound)
    throw Error("The starting input must be inside the supplied bounds.");
  const l = (a.validations ?? []).filter((z) =>
    isCellInRange(z.range, goalSeekRequest.variableAddress),
  );
  if (l.some((z) => !["number", "date"].includes(z.kind)))
    throw Error(
      "Goal Seek supports continuous numeric/date validation only; discrete list, integer or text-length constraints need a different input.",
    );
  if (
    l.some(
      (z) =>
        ![u.raw, goalSeekRequest.lowerBound, goalSeekRequest.upperBound].every(
          (V) => validateDataValidationRule(z, V),
        ),
    )
  )
    throw Error(
      "The starting input and both bounds must satisfy the variable's data validation rules.",
    );
  const c = inspectFormulaDependencies(
      {
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
        ),
      },
      {
        sheetId: sheet.id,
        address: goalSeekRequest.formulaAddress,
        direction: "precedents",
        depth: goalSeekLimits.maxDepth,
        includeHidden: context.includeHidden,
        allowedSheetIds: context.allowedSheetIds,
        limits: {
          maxFormulaCells: goalSeekLimits.maxFormulaCells,
          maxNodes: goalSeekLimits.maxDependencyCells,
          maxReferences: 1e4,
        },
      },
    ),
    h = [c.target, ...c.nodes],
    f = new Set(h.map((event) => event.key));
  if (
    h.length > goalSeekLimits.maxDependencyCells ||
    c.coverage.truncated ||
    c.coverage.depthLimited
  )
    throw Error(
      "Goal Seek requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
    );
  const issue = c.issues.find((issue) => f.has(issue.from));
  if (issue || c.cycles.length)
    throw Error(
      `Goal Seek requires supported static dependencies: ${issue?.message ?? "the dependency graph contains a cycle."}`,
    );
  if (!f.has(buildCompositeKey(a.id, goalSeekRequest.variableAddress)))
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
    const sheet = workbook.sheets.find((sheet) => sheet.id === z.sheetId),
      cell = sheet.cells[z.address];
    if (cell?.raw === null || !cell)
      throw Error(
        "Goal Seek requires populated static dependency cells; blank or possible spill-child dependencies are not supported.",
      );
    if (
      !cell.literal &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=")
    ) {
      const B = resolveDefinedNameValue(cell.raw, workbook, sheet.id).replace(
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
    ...workbook,
    sheets: workbook.sheets
      .filter((sheet) => context.allowedSheetIds.includes(sheet.id))
      .map((z) => ({
        ...z,
        cells: Object.fromEntries(
          Object.entries(z.cells).filter(([V]) =>
            f.has(buildCompositeKey(z.id, V)),
          ),
        ),
      })),
  };
  let evaluations = 0,
    iterations = 0,
    workUnits = 0;
  const baseline = { input: u.raw, output: null };
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
        evaluations >= goalSeekLimits.maxEvaluations ||
        workUnits + h.length > goalSeekLimits.maxWorkUnits
      )
        throw new k(
          "work-limit",
          "The Goal Seek evaluation/work budget was reached.",
        );
      (evaluations++, (workUnits += h.length));
      const workbook = {
          ...v,
          sheets: v.sheets.map((sheet) =>
            sheet.id === a.id
              ? {
                  ...sheet,
                  cells: {
                    ...sheet.cells,
                    [goalSeekRequest.variableAddress]: { ...u, raw: z },
                  },
                }
              : sheet,
          ),
        },
        workbookWithSettings = calculateWorkbookWithSettings(workbook),
        _ = workbookWithSettings(sheet.id, goalSeekRequest.formulaAddress);
      if (workbookWithSettings.getOperationalErrors().length)
        throw new k(
          "calculation-error",
          `A trial encountered an operational calculation error (${workbookWithSettings.getOperationalErrors().join(", ")}), even if a formula error handler concealed it. The baseline is unchanged.`,
        );
      if (!isFiniteNumber(_))
        throw new k(
          "calculation-error",
          `The formula returned ${String(_)} at input ${z}; a finite numeric result is required.`,
        );
      for (const N of h)
        if (workbookWithSettings.getSpillRange(N.sheetId, N.address))
          throw new k(
            "calculation-error",
            "A trial produced an array/spill result. Goal Seek requires scalar dependencies.",
          );
      if (workbookWithSettings.getOperationalErrors().length)
        throw new k(
          "calculation-error",
          "Dependency inspection encountered an operational calculation error; the baseline is unchanged.",
        );
      const residual = _ - goalSeekRequest.target;
      if (!isFiniteNumber(residual))
        throw new k(
          "calculation-error",
          "The residual exceeded finite numeric precision.",
        );
      const T = { input: z, value: _, residual: residual };
      return (
        A.set(z, T),
        (E.residual === null || Math.abs(residual) < Math.abs(E.residual)) &&
          (E = T),
        T
      );
    },
    I = (z, V) => {
      const converged = z === "converged" || z === "already-satisfied";
      if (converged && l.some((C) => !validateDataValidationRule(C, E.input)))
        return I(
          "validation-error",
          "The candidate violates the variable's data validation rule.",
        );
      const patches =
          converged && E.input !== u.raw
            ? [
                {
                  sheetId: a.id,
                  address: goalSeekRequest.variableAddress,
                  cell: { ...u, raw: E.input },
                },
              ]
            : [],
        next = patches.length
          ? {
              ...workbook,
              sheets: workbook.sheets.map((sheet) =>
                sheet.id === a.id
                  ? {
                      ...sheet,
                      cells: {
                        ...sheet.cells,
                        [goalSeekRequest.variableAddress]: patches[0].cell,
                      },
                    }
                  : sheet,
              ),
            }
          : workbook;
      return {
        snapshot: JSON.stringify(workbook),
        request: goalSeekRequest,
        context: context,
        baseline: baseline,
        dependencyCells: h.length,
        result: {
          status: z,
          converged: converged,
          input: E.input,
          achievedValue: E.value,
          residual: E.residual,
          absoluteResidual: E.residual === null ? null : Math.abs(E.residual),
          iterations: iterations,
          evaluations: evaluations,
          workUnits: workUnits,
          algorithm: "Bracketed bisection",
          message: V,
        },
        patches: patches,
        next: next,
      };
    };
  try {
    const z = F(u.raw);
    if (
      ((baseline.output = z.value),
      Math.abs(z.residual) <= goalSeekRequest.tolerance)
    )
      return I(
        "already-satisfied",
        "The baseline already meets the requested tolerance. No input change is needed.",
      );
    let V = F(goalSeekRequest.lowerBound);
    if (Math.abs(V.residual) <= goalSeekRequest.tolerance)
      return I("converged", "The lower bound meets the requested tolerance.");
    let D = F(goalSeekRequest.upperBound);
    if (Math.abs(D.residual) <= goalSeekRequest.tolerance)
      return I("converged", "The upper bound meets the requested tolerance.");
    if (Math.sign(V.residual) === Math.sign(D.residual))
      return I(
        "not-bracketed",
        "The supplied bounds do not bracket a target crossing. A solution may exist elsewhere or inside a non-monotonic interval; choose a narrower valid bracket.",
      );
    for (
      iterations = 1;
      iterations <= goalSeekRequest.maxIterations;
      iterations++
    ) {
      const B = V.input / 2 + D.input / 2;
      if (B === V.input || B === D.input)
        return I(
          "precision-limit",
          "The bracket cannot be refined further at numeric precision; the residual still exceeds tolerance.",
        );
      const _ = F(B);
      if (Math.abs(_.residual) <= goalSeekRequest.tolerance)
        return I(
          "converged",
          "The candidate meets the requested absolute residual tolerance.",
        );
      Math.sign(_.residual) === Math.sign(V.residual) ? (V = _) : (D = _);
    }
    return (
      (iterations = goalSeekRequest.maxIterations),
      I(
        "iteration-limit",
        "The iteration limit was reached without meeting the residual tolerance. The baseline is unchanged.",
      )
    );
  } catch (error) {
    return I(
      error instanceof k ? error.status : "calculation-error",
      error instanceof Error
        ? error.message
        : "The trial calculation failed; the baseline is unchanged.",
    );
  }
}
function applyGoalSeekResult(workbook, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (JSON.stringify(workbook) !== r.snapshot)
    throw Error(
      "The workbook changed after the Goal Seek preview. Solve and review again.",
    );
  if (!r.result.converged)
    throw Error(
      "Only a candidate that meets the residual tolerance can be applied.",
    );
  const n = solveGoalSeek(workbook, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error(
      "The Goal Seek preview or host context changed. Solve and review again.",
    );
  return n.next;
}
function GoalSeekDialog(e) {
  const workbookMessages = useWorkbookMessages(),
    [entryLocale] = ne.useState(workbookMessages.entryLocale),
    n = (W, K) => validateLocalizedNumberEntry(W, K, entryLocale),
    { text, language } = useWorkbookMessages(),
    s = ne.useId(),
    d = ne.useRef(null),
    u = ne.useRef(e.onClose);
  u.current = e.onClose;
  const [formulaSheetId, c] = ne.useState(e.sheetId),
    [formulaAddress, f] = ne.useState(e.address),
    [variableSheetId, m] = ne.useState(e.sheetId),
    [variableAddress, g] = ne.useState(""),
    [b, y] = ne.useState(""),
    [w, E] = ne.useState("0"),
    [A, k] = ne.useState("100"),
    [F, I] = ne.useState(() =>
      formatCellEntryForEdit({ raw: 1e-6 }, entryLocale),
    ),
    [O, j] = ne.useState("64"),
    [checked, V] = ne.useState(false),
    [D, B] = ne.useState(null),
    [_, C] = ne.useState(""),
    T = {
      includeHidden: !!e.allowHidden && checked,
      allowedSheetIds: e.allowedSheetIds,
    },
    N = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    R = e.workbook.sheets.filter(
      (sheet) =>
        N.includes(sheet.id) &&
        (T.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    L = R.some((W) => W.id === formulaSheetId),
    $ = R.some((W) => W.id === variableSheetId),
    U =
      !!D &&
      D.snapshot === JSON.stringify(e.workbook) &&
      D.context.includeHidden === T.includeHidden &&
      JSON.stringify(D.context.allowedSheetIds) === JSON.stringify(N);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      d.current?.querySelector('[data-ui-label="Set cell"]')?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
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
    onClick = () => {
      try {
        const request = {
          formulaSheetId: formulaSheetId,
          formulaAddress: formulaAddress,
          variableSheetId: variableSheetId,
          variableAddress: variableAddress,
          target: n(b, "Target value"),
          lowerBound: n(w, "Lower bound"),
          upperBound: n(A, "Upper bound"),
          tolerance: n(F, "Tolerance"),
          maxIterations: n(O, "Maximum iterations"),
        };
        (B(solveGoalSeek(e.workbook, request, { ...T, readOnly: e.readOnly })),
          C(""));
      } catch (error) {
        (B(null), H(error));
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
        } catch (error) {
          H(error);
        }
    };
  return i.jsx("div", {
    className: "gs-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: d,
      className: "gs-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${s}-title`,
      onKeyDown: (event) => {
        var te, ae;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), u.current()),
          event.key === "Tab")
        ) {
          const oe = [
            ...(d.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),summary",
            ) ?? []),
          ].filter((de) => !de.hidden);
          event.shiftKey && document.activeElement === oe[0]
            ? (event.preventDefault(), (te = oe.at(-1)) == null || te.focus())
            : !event.shiftKey &&
              document.activeElement === oe.at(-1) &&
              (event.preventDefault(), (ae = oe[0]) == null || ae.focus());
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
                  value: L ? formulaSheetId : "",
                  disabled: e.readOnly,
                  onChange: (event) => G(c, event.target.value),
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
                  value: L ? formulaAddress : "",
                  placeholder: text("B1"),
                  readOnly: e.readOnly,
                  onChange: (event) => G(f, event.target.value),
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
                  value: $ ? variableSheetId : "",
                  disabled: e.readOnly,
                  onChange: (event) => G(m, event.target.value),
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
                  value: $ ? variableAddress : "",
                  placeholder: text("A1"),
                  readOnly: e.readOnly,
                  onChange: (event) => G(g, event.target.value),
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
                  onChange: (event) => G(y, event.target.value),
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
                  onChange: (event) => G(I, event.target.value),
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
                  onChange: (event) => G(E, event.target.value),
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
                  onChange: (event) => G(k, event.target.value),
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
                  onChange: (event) => G(j, event.target.value),
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
                    checked: checked,
                    onChange: (event) => {
                      (V(event.target.checked), B(null), C(""));
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
          onClick: onClick,
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
                                (sheet) =>
                                  sheet.id === D.request.formulaSheetId,
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
                                (sheet) =>
                                  sheet.id === D.request.variableSheetId,
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
function normalizeAnalysisPermissions(workbook, r = {}) {
  if (
    (r.readOnly !== undefined && typeof r.readOnly != "boolean") ||
    (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
  )
    throw Error("Analysis permissions must be boolean.");
  if (
    r.allowedSheetIds !== undefined &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("The host scope contains an unknown worksheet.");
  return {
    readOnly: !!r.readOnly,
    includeHidden: !!r.includeHidden,
    allowedSheetIds: [
      ...new Set(r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id)),
    ].sort(),
  };
}
function validateAnalysisSource(workbook, r, t) {
  const analysisPermissions = normalizeAnalysisPermissions(workbook, t),
    sheet = workbook.sheets.find((sheet) => sheet.id === r.sheetId);
  if (!sheet || !analysisPermissions.allowedSheetIds.includes(sheet.id))
    throw Error("A source is outside the authorized host context.");
  if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(r.address))
    throw Error("Use a canonical cell address.");
  const a = _t(r.address);
  if (a.row >= 1048576 || a.col >= 16384)
    throw Error("The cell address exceeds Excel dimensions.");
  if (
    !analysisPermissions.includeHidden &&
    (sheet.hidden ||
      sheet.veryHidden ||
      cn(sheet, "row").some((d) => d.start <= a.row && a.row <= d.end) ||
      cn(sheet, "column").some((d) => d.start <= a.col && a.col <= d.end))
  )
    throw Error(
      "A source is hidden. Explicitly include hidden cells within the authorized scope.",
    );
  if (sheet.analysisReport || sheet.dataModelReport)
    throw Error(
      "Analysis and model report snapshots cannot be analysis inputs. Use the original source cells.",
    );
  if (sheet.merges?.some((merge) => parseRangeBounds(merge, r.address)))
    throw Error("Merged analysis sources are ambiguous.");
  return sheet;
}
function parseRangeBounds(e, address) {
  const rangeAddress = parseRangeAddress(e),
    n = _t(address);
  return (
    n.row >= rangeAddress.firstRow &&
    n.row <= rangeAddress.lastRow &&
    n.col >= rangeAddress.firstCol &&
    n.col <= rangeAddress.lastCol
  );
}
function p_(workbook) {
  return fingerprintCanonicalValue({
    sheets: workbook.sheets
      .filter((sheet) => !sheet.analysisReport)
      .map((r) => ({ ...r })),
    calculation: workbook.calculation ?? null,
    datePolicy: workbook.datePolicy ?? null,
    definedNames: workbook.definedNames ?? null,
  });
}
function assertAutomaticCalculation(workbook) {
  if (
    workbook.calculation?.mode === "manual" ||
    workbook.calculation?.iterative
  )
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
