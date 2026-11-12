import { clampNumber } from "./53-find-replace-go-to.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  string,
  strictObject,
  enum as _enum,
  discriminatedUnion,
  literal,
} from "zod";
import { fingerprintCanonicalValue } from "./01-core-schemas-and-hashing.js";
import { workbookLimits } from "./02-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  isSameContentHash,
  createWorksheetRecord,
  isExternalSourceError,
  externalSourceFormatVersion,
  ExternalSourceError,
  validateExternalSourcePayload,
  parseRangeAddress,
  externalSourceLimits,
  parseActiveSourceApproval,
  buildCanonicalWorkbookSnapshot,
  normalizeWorkbookInput,
  columnIndexToLetters,
  computeSourceControlTotals,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  createLocalWorkbookStore,
  WorkbookRevisionConflictError,
  compareWorkbookVersions,
} from "./11-workbook-formats-dialog.js";
import { inspectOutputNameSafety } from "./12-formula-dependency-inspector.js";
import {
  resolveAxisScale,
  physicalToLogicalPixels,
} from "./53-find-replace-go-to.js";

function computeVirtualizedRange(e, r, t, n, o = 5) {
  const a = resolveAxisScale(e.total, t, n),
    s = clampNumber(r, 0, a.physicalMax),
    d = physicalToLogicalPixels(s, e.total, t, n),
    u = e.count ? Math.max(0, e.indexAt(d) - o) : 0,
    l = e.count ? Math.min(e.count, e.indexAt(d + a.viewport / n) + o + 1) : 0,
    c = Math.max(0, s / n + e.offset(u) - d),
    h = Math.max(0, a.physicalPixels / n - c - (e.offset(l) - e.offset(u)));
  return {
    first: u,
    last: l,
    leading: c,
    trailing: h,
    logicalOffset: d,
    total: a.physicalPixels / n,
  };
}
function workbookLockName(e) {
  return `ledgerglass:workbook:${e}`;
}
function getBrowserLockManager() {
  try {
    return typeof navigator < "u" ? (navigator.locks ?? null) : null;
  } catch {
    return null;
  }
}
function cloneSaveInputs(e) {
  try {
    return structuredClone(e);
  } catch {
    throw new WorkbookRevisionConflictError(
      "invalid",
      "The workbook save inputs could not be captured. Keep your unsaved work open and correct the invalid input before retrying.",
    );
  }
}
function saveWorkbookWithLock(e, r, t = {}) {
  const n = createLocalWorkbookStore(e, r, {
      maxBytes: t.maxBytes,
      maxCheckpoints: t.maxCheckpoints,
    }),
    o = t.locks === undefined ? getBrowserLockManager() : t.locks,
    a = t.lockTimeoutMs ?? 5e3,
    s = workbookLockName(r);
  if (!Number.isInteger(a) || a < 1 || a > 6e4)
    throw new WorkbookRevisionConflictError(
      "invalid",
      "The browser save lock timeout must be between 1 and 60,000 milliseconds.",
    );
  function d(u, l) {
    return new Promise((c, h) => {
      const f = (j) => new WorkbookRevisionConflictError("unavailable", j),
        p = new AbortController(),
        m = performance.now() + a;
      let v = true,
        g = false,
        b = false,
        y,
        w,
        E;
      const A = () => {
          (E !== undefined && clearTimeout(E),
            l == null || l.removeEventListener("abort", F));
        },
        k = (j) => {
          !v || g || ((v = false), (w = j), A(), p.abort(j), h(j));
        },
        F = () =>
          k(
            f(
              "The queued browser save was cancelled. No changes were written; your unsaved work remains in memory.",
            ),
          ),
        I = () =>
          f(
            "Timed out waiting for another workbook window to finish saving. No changes were written; retry or export your unsaved work.",
          ),
        O = (j) => {
          ((v = false),
            A(),
            h(
              b
                ? y
                : j instanceof WorkbookRevisionConflictError
                  ? j
                  : f(
                      "Browser save locking is unavailable. No changes were written; keep the workbook open and export your unsaved work.",
                    ),
            ));
        };
      if (l != null && l.aborted) {
        F();
        return;
      }
      try {
        if (!o || typeof o.request != "function") {
          O(
            f(
              "Browser save locking is unavailable. No changes were written; keep the workbook open and export your unsaved work.",
            ),
          );
          return;
        }
        (l == null || l.addEventListener("abort", F, { once: true }),
          (E = setTimeout(() => k(I()), a)));
        const j = o.request(s, { mode: "exclusive", signal: p.signal }, (z) => {
          if (!v || p.signal.aborted)
            throw w ?? f("The queued browser save was cancelled.");
          if (performance.now() >= m) {
            const V = I();
            throw (k(V), V);
          }
          if (!z || z.name !== s || z.mode !== "exclusive")
            throw f(
              "The browser did not grant the required exclusive save lock.",
            );
          ((g = true), A());
          try {
            return u();
          } catch (V) {
            throw ((b = true), (y = V), V);
          }
        });
        Promise.resolve(j).then((z) => {
          ((v = false), A(), c(z));
        }, O);
      } catch (j) {
        O(j);
      }
    });
  }
  return {
    coordination: "web-locks",
    load: n.load,
    listCheckpoints: n.listCheckpoints,
    compareCheckpoint: n.compareCheckpoint,
    async save(u, l, c, h = {}) {
      const f = cloneSaveInputs({
          workbook: u,
          view: l,
          expected: c,
          options: { allowIdentityChange: h.allowIdentityChange },
        }),
        p = h.signal;
      return d(() => n.save(f.workbook, f.view, f.expected, f.options), p);
    },
    async checkpoint(u, l, c = {}) {
      const h = cloneSaveInputs({ label: u, expected: l }),
        f = c.signal;
      return d(() => n.checkpoint(h.label, h.expected), f);
    },
    async restoreCheckpoint(u, l, c, h, f = {}) {
      const p = cloneSaveInputs({ id: u, workbook: l, expected: c, actor: h }),
        m = f.signal;
      return d(
        () => n.restoreCheckpoint(p.id, p.workbook, p.expected, p.actor),
        m,
      );
    },
  };
}
const workbookSectionLabels = {
    name: "Name",
    metadata: "Workbook properties",
    createdAt: "Created",
    modifiedAt: "Last modified",
    engagement: "Engagement",
    period: "Period",
    template: "Template",
    sourceRevision: "Original revision",
    sourceWorkbookId: "Original workbook",
    definedNames: "Named ranges",
    formulaWatches: "Watched formulas",
    styleCatalog: "Styles and theme",
    reviewThreads: "Review comments",
    structureProtection: "Workbook protection",
    protection: "Protection",
    enabled: "Enabled",
    hidden: "Hidden",
    veryHidden: "Hidden from worksheet menu",
    tabColor: "Tab color",
    rowCount: "Worksheet rows",
    colCount: "Worksheet columns",
    columnWidths: "Column widths",
    rowHeights: "Row heights",
    frozenRows: "Frozen rows",
    frozenCols: "Frozen columns",
    hiddenRows: "Hidden rows",
    hiddenCols: "Hidden columns",
    merges: "Merged cells",
    drawings: "Pictures, shapes and charts",
    validations: "Cell validation rules",
    conditionalFormats: "Conditional formatting",
    pageLayout: "Print layout",
    tables: "Tables",
    pivots: "Pivot tables",
    queries: "Queries",
    literal: "Entered as text",
    numFmt: "Number format",
    format: "Number format category",
    style: "Formatting",
    note: "Note",
    tickmark: "Review mark",
    bold: "Bold",
    italic: "Italic",
    underline: "Underline",
    fontFamily: "Font",
    fontSize: "Font size",
    wrap: "Wrap text",
    border: "Borders",
    fill: "Fill color",
    color: "Text color",
    align: "Alignment",
  },
  humanizeFieldName = (e) =>
    workbookSectionLabels[e] ??
    e
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .replace(/[_-]/g, " ")
      .replace(/^./, (r) => r.toUpperCase()),
  wp = (e) => e !== null && typeof e == "object" && !Array.isArray(e),
  stableSerialize = (e) =>
    Array.isArray(e)
      ? `[${e.map(stableSerialize).join(",")}]`
      : wp(e)
        ? `{${Object.keys(e)
            .sort()
            .map((r) => `${JSON.stringify(r)}:${stableSerialize(e[r])}`)
            .join(",")}}`
        : (JSON.stringify(e) ?? "undefined"),
  truncateDisplayText = (e, r = 220) =>
    e.length > r ? `${e.slice(0, r)}…` : e;
function summarizeConflictValue(e, r = 0) {
  if (e === undefined) return "Not set";
  if (e === null || e === "") return "Blank";
  if (typeof e == "boolean") return e ? "Yes" : "No";
  if (typeof e == "string")
    return e.startsWith("data:image/")
      ? "Embedded image"
      : /^\d{4}-\d{2}-\d{2}T/.test(e) && Number.isFinite(Date.parse(e))
        ? new Date(e).toLocaleString(undefined, {
            year: "numeric",
            month: "short",
            day: "numeric",
            hour: "numeric",
            minute: "2-digit",
            second: "2-digit",
            fractionalSecondDigits: 3,
            timeZoneName: "short",
          })
        : truncateDisplayText(e);
  if (Array.isArray(e))
    return e.length
      ? r >= 2
        ? `${e.length} items`
        : truncateDisplayText(
            e
              .slice(0, 4)
              .map((t) => summarizeConflictValue(t, r + 1))
              .join("; ") +
              (e.length > 4 ? `; ${e.length - 4} more items` : ""),
          )
      : "None";
  if (wp(e)) {
    const t = Object.entries(e);
    return t.length
      ? r >= 2
        ? `${t.length} settings`
        : truncateDisplayText(
            t
              .slice(0, 4)
              .map(
                ([n, o]) =>
                  `${humanizeFieldName(n)}: ${summarizeConflictValue(o, r + 1)}`,
              )
              .join(" · ") +
              (t.length > 4 ? ` · ${t.length - 4} more settings` : ""),
          )
      : "No settings";
  }
  return String(e);
}
function describeCellContent(e) {
  return e
    ? e.raw === null || e.raw === ""
      ? "Blank"
      : typeof e.raw == "boolean"
        ? e.raw
          ? "TRUE"
          : "FALSE"
        : typeof e.raw == "string"
          ? (e.literal ? "Text: " : e.raw.startsWith("=") ? "Formula: " : "") +
            truncateDisplayText(e.raw)
          : String(e.raw)
    : "Empty cell";
}
function collectConflictChanges(e) {
  const r = [],
    t = (n, o, a, s) => {
      if (stableSerialize(o) !== stableSerialize(a)) {
        if (
          s < 3 &&
          (wp(o) || wp(a)) &&
          (o === undefined || wp(o)) &&
          (a === undefined || wp(a))
        ) {
          const d = wp(o) ? o : {},
            u = wp(a) ? a : {},
            l = new Set([...Object.keys(d), ...Object.keys(u)]);
          if (l.size) {
            for (const c of l) t([...n, c], d[c], u[c], s + 1);
            return;
          }
        }
        r.push({ path: n, before: o, after: a });
      }
    };
  return (t(e.path.split("."), e.before, e.after, 0), r);
}
function WorkbookSaveConflictDialog(e) {
  const { text, language } = useWorkbookMessages(),
    [n, o] = ne.useState(null),
    [a, s] = ne.useState(null),
    [d, u] = ne.useState(""),
    [l, c] = ne.useState(false),
    [h, f] = ne.useState(false),
    p = ne.useRef(null),
    m = () => {
      (c(false), f(false));
      try {
        const w = e.loadSaved();
        (o(w),
          s(
            w?.workbook.id === e.workbook.id
              ? compareWorkbookVersions(w.workbook, e.workbook)
              : null,
          ),
          u(
            w
              ? ""
              : "The saved workbook is no longer available. Export your local changes before continuing.",
          ));
      } catch (w) {
        (o(null),
          s(null),
          u(
            w instanceof Error
              ? w.message
              : "The saved workbook cannot be read.",
          ));
      }
    };
  (ne.useEffect(m, [e.workbook]),
    ne.useEffect(() => {
      var A;
      const w = document.activeElement;
      return (
        (A = p.current?.querySelector("button")) == null || A.focus(),
        () => {
          w != null && w.isConnected && w.focus();
        }
      );
    }, []));
  const v = (w) => {
      return (
        e.workbook.sheets.find((k) => k.id === w)?.name ??
        n?.workbook.sheets.find((k) => k.id === w)?.name ??
        "Removed worksheet"
      );
    },
    g = (w) =>
      w[0] === "sheets" && w[1] === "order"
        ? "Worksheet order"
        : w[0] === "sheets"
          ? [v(w[1]), ...w.slice(2).map(humanizeFieldName)].join(" · ")
          : w.map(humanizeFieldName).join(" · "),
    b = a
      ? [...a.metadataChanges, ...a.structuralChanges].flatMap(
          collectConflictChanges,
        )
      : [],
    y = (w, E) =>
      w[0] === "sheets" && w[1] === "order" && Array.isArray(E)
        ? E.map((A) => v(String(A))).join(" → ") || "No worksheets"
        : summarizeConflictValue(E);
  return i.jsx("div", {
    className: "wm-backdrop",
    children: i.jsxs("div", {
      ref: p,
      className: "wm-dialog pc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": text("Workbook save conflict"),
      onKeyDown: (w) => {
        if (
          (w.key === "Escape" &&
            (w.preventDefault(), w.stopPropagation(), e.onClose()),
          w.key === "Tab")
        ) {
          const A = [
              ...(p.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
              ) ?? []),
            ],
            k = A[0],
            F = A.at(-1);
          w.shiftKey && document.activeElement === k
            ? (w.preventDefault(), F == null || F.focus())
            : !w.shiftKey &&
              document.activeElement === F &&
              (w.preventDefault(), k == null || k.focus());
        }
      },
      lang: language("Workbook save conflict"),
      "data-ui-label": "Workbook save conflict",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children: i.jsx(LocalizedText, {
                children: "Workbook save conflict",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close save conflict"),
              onClick: e.onClose,
              lang: language("Close save conflict"),
              "data-ui-label": "Close save conflict",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wm-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Your edits are still in this editor. Export them to keep a portable copy, or compare them with the saved version before reloading.",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsx("strong", {
                  children: i.jsx(LocalizedText, {
                    children: "My local workbook:",
                  }),
                }),
                " ",
                e.workbook.name,
                " ",
                i.jsx(LocalizedText, { children: "· revision" }),
                " ",
                e.workbook.revision,
              ],
            }),
            n &&
              i.jsxs("p", {
                children: [
                  i.jsx("strong", {
                    children: i.jsx(LocalizedText, {
                      children: "Current saved workbook:",
                    }),
                  }),
                  " ",
                  n.workbook.name,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision " }),
                  n.workbook.revision,
                ],
              }),
            i.jsxs("div", {
              className: "pc-actions",
              children: [
                i.jsx("button", {
                  onClick: e.onExport,
                  children: i.jsx(LocalizedText, {
                    children: "Export my changes",
                  }),
                }),
                i.jsx("button", {
                  onClick: m,
                  children: i.jsx(LocalizedText, {
                    children: "Refresh comparison",
                  }),
                }),
              ],
            }),
            n &&
              n.workbook.id !== e.workbook.id &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "The saved location now contains a different workbook. Export your local workbook before deciding whether to load it.",
                }),
              }),
            a &&
              i.jsxs("section", {
                "aria-label": text("Conflicting workbook changes"),
                lang: language("Conflicting workbook changes"),
                "data-ui-label": "Conflicting workbook changes",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Changes from saved to my local workbook",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      a.cellChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "cell changes ·" }),
                      " ",
                      a.metadataChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "workbook changes ·" }),
                      " ",
                      a.structuralChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "worksheet changes" }),
                    ],
                  }),
                  a.cellChanges.length > 0 &&
                    i.jsxs("table", {
                      className: "pc-table",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Cell",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Saved",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "My local value",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: a.cellChanges.slice(0, 100).map((w) => {
                            const E = [
                              ...new Set([
                                ...Object.keys(w.before ?? {}),
                                ...Object.keys(w.after ?? {}),
                              ]),
                            ].filter((A) => {
                              return (
                                A !== "raw" &&
                                stableSerialize(w.before?.[A]) !==
                                  stableSerialize(w.after?.[A])
                              );
                            });
                            return i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("th", {
                                    scope: "row",
                                    children: [v(w.sheetId), "!", w.address],
                                  }),
                                  [w.before, w.after].map((A, k) =>
                                    i.jsxs(
                                      "td",
                                      {
                                        children: [
                                          i.jsx("span", {
                                            className: "pc-cell-value",
                                            children: describeCellContent(A),
                                          }),
                                          E.length > 0 &&
                                            i.jsx("ul", {
                                              className: "pc-cell-fields",
                                              children: E.map((F) =>
                                                i.jsxs(
                                                  "li",
                                                  {
                                                    children: [
                                                      i.jsxs("strong", {
                                                        children: [
                                                          humanizeFieldName(F),
                                                          ":",
                                                        ],
                                                      }),
                                                      " ",
                                                      summarizeConflictValue(
                                                        A?.[F],
                                                      ),
                                                    ],
                                                  },
                                                  F,
                                                ),
                                              ),
                                            }),
                                        ],
                                      },
                                      k,
                                    ),
                                  ),
                                ],
                              },
                              `${w.sheetId}:${w.address}`,
                            );
                          }),
                        }),
                      ],
                    }),
                  a.cellChanges.length > 100 &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Showing the first 100 changed cells. Full comparison details include every change.",
                      }),
                    }),
                  b.length > 0 &&
                    i.jsxs("table", {
                      className: "pc-table",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Property",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Saved",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "My local setting",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: b.slice(0, 100).map((w) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", {
                                    scope: "row",
                                    children: g(w.path),
                                  }),
                                  i.jsx("td", {
                                    children: y(w.path, w.before),
                                  }),
                                  i.jsx("td", {
                                    children: y(w.path, w.after),
                                  }),
                                ],
                              },
                              JSON.stringify(w.path),
                            ),
                          ),
                        }),
                      ],
                    }),
                  b.length > 100 &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Showing the first 100 changed properties. Full comparison details include every change.",
                      }),
                    }),
                  (a.cellChanges.length > 0 || b.length > 0) &&
                    i.jsxs("details", {
                      className: "pc-details",
                      open: h,
                      onToggle: (w) => f(w.currentTarget.open),
                      children: [
                        i.jsx("summary", {
                          children: i.jsx(LocalizedText, {
                            children: "Full comparison details",
                          }),
                        }),
                        h &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("p", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Complete recorded values, formulas, notes, formatting and properties, including long values and additional items summarized above.",
                                }),
                              }),
                              i.jsx("pre", {
                                tabIndex: 0,
                                children: JSON.stringify(a, null, 2),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "pc-footer",
          children: [
            d && i.jsx("p", { role: "alert", children: d }),
            i.jsxs("label", {
              className: "wm-confirm",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: l,
                  disabled: !n,
                  onChange: (w) => c(w.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children:
                    "Discard my local changes and load the saved workbook.",
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Reloading replaces this editor's data and clears its local Undo/Redo history. Other open workbooks remain unchanged.",
              }),
            }),
            i.jsxs("div", {
              className: "pc-footer-actions",
              children: [
                i.jsx("button", {
                  disabled: !n || !l,
                  onClick: () => {
                    try {
                      e.onReload(n);
                    } catch (w) {
                      (c(false),
                        u(
                          w instanceof Error
                            ? w.message
                            : "The saved workbook could not be loaded.",
                        ));
                    }
                  },
                  children: i.jsx(LocalizedText, {
                    children: "Reload saved workbook",
                  }),
                }),
                i.jsx("button", {
                  onClick: e.onClose,
                  children: i.jsx(LocalizedText, {
                    children: "Keep editing my local copy",
                  }),
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}
const diagnosticGroupDefinitions = {
    persistence: ["load", "save", "checkpoint", "list", "compare", "restore"],
    calculation: ["run"],
    file: [
      "workbook",
      "delimited",
      "automation",
      "json",
      "picture",
      "formats",
      "source",
    ],
    provider: ["request"],
    source: ["catalog", "authorize", "refresh"],
    host: ["change", "selection", "loaded", "save-state", "proposal"],
  },
  wH = new Set([
    "unavailable",
    "timeout",
    "cancelled",
    "conflict",
    "corrupt",
    "quota",
    "limit",
    "invalid",
    "refused",
    "unknown",
  ]),
  workbookErrorCodes = {
    expired: "refused",
    schema: "invalid",
    stale: "conflict",
    FILE_TIMEOUT: "timeout",
    FILE_CANCELLED: "cancelled",
    FILE_LIMIT: "limit",
    FILE_BUSY: "limit",
    FILE_INVALID: "invalid",
    FILE_ACTIVE: "refused",
    FILE_TYPE: "invalid",
    FILE_READ: "unavailable",
    FILE_WORKER: "unavailable",
    malformed_response: "invalid",
    invalid_request: "invalid",
    "#TIMEOUT!": "timeout",
    "#LIMIT!": "limit",
    "#BLOCKED!": "unavailable",
    "#PENDING!": "cancelled",
  };
function isAbortOrQuotaError(e) {
  try {
    if (e && typeof e == "object") {
      const { code, name } = e;
      if (typeof code == "string") {
        if (wH.has(code)) return code;
        if (Object.hasOwn(workbookErrorCodes, code))
          return workbookErrorCodes[code];
      }
      if (name === "AbortError") return "cancelled";
      if (name === "TimeoutError") return "timeout";
      if (name === "QuotaExceededError") return "quota";
    }
  } catch {}
  return "unknown";
}
function createWorkbookDiagnosticsMonitor(e = {}) {
  const r = e.maxEvents ?? 200;
  if (!Number.isInteger(r) || r < 2 || r > 1e3)
    throw new Error(
      "Diagnostic history must retain between 2 and 1000 events.",
    );
  const t = new Set(),
    n = new Map();
  let o = 0,
    a = 0;
  const s = Object.fromEntries(
    Object.keys(diagnosticGroupDefinitions).map((m) => [
      m,
      Object.freeze({
        active: 0,
        completed: 0,
        failures: 0,
        cancellations: 0,
        maxElapsedMs: 0,
        totalElapsedMs: 0,
        issues: Object.freeze({}),
      }),
    ]),
  );
  let d = Object.freeze({
      events: Object.freeze([]),
      health: Object.freeze(s),
      droppedEvents: 0,
      sinkFailures: 0,
      droppedDeliveries: 0,
    }),
    u = false;
  const l = () => {
      u ||
        ((u = true),
        queueMicrotask(() => {
          u = false;
          for (const m of t)
            try {
              m();
            } catch {}
        }));
    },
    c = () => {
      ((d = Object.freeze({ ...d, sinkFailures: d.sinkFailures + 1 })), l());
    };
  let h = 0;
  const f = (m, v) => {
      const g = [...d.events, Object.freeze(m)],
        b = Math.max(0, g.length - r);
      if (
        ((d = Object.freeze({
          ...d,
          events: Object.freeze(g.slice(-r)),
          health: Object.freeze({ ...d.health, [m.adapter]: Object.freeze(v) }),
          droppedEvents: d.droppedEvents + b,
        })),
        l(),
        h >= 16)
      ) {
        d = Object.freeze({ ...d, droppedDeliveries: d.droppedDeliveries + 1 });
        return;
      }
      try {
        const w = e.onEvent?.call(e, m);
        w &&
          (h++,
          Promise.resolve(w).then(
            () => {
              h--;
            },
            () => {
              (h--, c());
            },
          ));
      } catch {
        c();
      }
    },
    p = (m, v) => {
      if (
        !Object.hasOwn(diagnosticGroupDefinitions, m) ||
        !diagnosticGroupDefinitions[m].includes(v)
      )
        throw new Error("Unknown diagnostic adapter or operation.");
      const g = ++a,
        b = performance.now(),
        y = `${m}:${v}`;
      let w = false;
      f(
        {
          version: 1,
          sequence: ++o,
          attempt: g,
          adapter: m,
          operation: v,
          phase: "started",
        },
        { ...d.health[m], active: d.health[m].active + 1 },
      );
      const E = (A = "succeeded", k) => {
        if (w) return;
        ((w = true),
          ["succeeded", "failed", "cancelled", "recovered"].includes(A) ||
            (A = "failed"));
        const F = k && wH.has(k) ? k : "unknown",
          I = Math.max(0, Math.round((performance.now() - b) * 10) / 10),
          O = d.health[m],
          j = { ...O.issues };
        (g >= (n.get(y) ?? 0) &&
          A !== "cancelled" &&
          (n.set(y, g), A === "failed" ? (j[v] = F) : delete j[v]),
          f(
            {
              version: 1,
              sequence: ++o,
              attempt: g,
              adapter: m,
              operation: v,
              phase: "completed",
              outcome: A,
              ...(A === "failed"
                ? { code: F }
                : A === "cancelled"
                  ? { code: "cancelled" }
                  : {}),
              elapsedMs: I,
            },
            {
              active: Math.max(0, O.active - 1),
              completed: O.completed + 1,
              failures: O.failures + +(A === "failed"),
              cancellations: O.cancellations + +(A === "cancelled"),
              maxElapsedMs: Math.max(O.maxElapsedMs, I),
              totalElapsedMs: O.totalElapsedMs + I,
              issues: Object.freeze(j),
            },
          ));
      };
      return {
        finish: E,
        fail(A) {
          const k = isAbortOrQuotaError(A);
          E(k === "cancelled" ? "cancelled" : "failed", k);
        },
      };
    };
  return {
    start: p,
    run(m, v, g) {
      const b = p(m, v);
      try {
        const y = g();
        return y && typeof y.then == "function"
          ? Promise.resolve(y).then(
              (w) => (b.finish(), w),
              (w) => {
                throw (b.fail(w), w);
              },
            )
          : (b.finish(), y);
      } catch (y) {
        throw (b.fail(y), y);
      }
    },
    getSnapshot: () => d,
    subscribe(m) {
      return (
        t.add(m),
        () => {
          t.delete(m);
        }
      );
    },
    clearEvents() {
      ((d = Object.freeze({ ...d, events: Object.freeze([]) })), l());
    },
  };
}
const invalidStorageAckError = () =>
    new WorkbookRevisionConflictError(
      "invalid",
      "The storage adapter returned an invalid acknowledgement. Your local changes are retained; check the saved copy before retrying.",
    ),
  workbookSnapshotsEqual = (e, r) =>
    e === r || JSON.stringify(e) === JSON.stringify(r);
function captureWorkbookView(e, r) {
  const t = e.sheets.find((s) => !s.hidden && s.id === r.activeSheetId);
  if (!t || r.selection.sheetId !== t.id) throw invalidStorageAckError();
  const n = parseRangeAddress(r.selection.range);
  if (n.lastRow >= t.rowCount || n.lastCol >= t.colCount)
    throw invalidStorageAckError();
  const o = `${columnIndexToLetters(n.firstCol)}${n.firstRow + 1}`,
    a = `${columnIndexToLetters(n.lastCol)}${n.lastRow + 1}`;
  return {
    activeSheetId: t.id,
    selection: { sheetId: t.id, range: o === a ? o : `${o}:${a}` },
  };
}
function normalizeCheckpointRecord(e, r) {
  try {
    const t = buildCanonicalWorkbookSnapshot(e.workbook),
      n = captureWorkbookView(t, e.view),
      o = e.guard;
    if (
      !o ||
      typeof o.namespace != "string" ||
      !o.namespace ||
      !Number.isSafeInteger(o.generation) ||
      o.generation < 0 ||
      typeof o.digest != "string" ||
      !o.digest ||
      o.workbookId !== t.id ||
      o.revision !== t.revision ||
      typeof e.recovered != "boolean" ||
      !Array.isArray(e.checkpoints) ||
      e.checkpoints.length > 20 ||
      e.checkpoints.some(
        (a) =>
          !a ||
          typeof a.id != "string" ||
          typeof a.label != "string" ||
          typeof a.workbookId != "string" ||
          !Number.isSafeInteger(a.revision) ||
          a.revision < 0 ||
          typeof a.createdAt != "string" ||
          !Number.isFinite(Date.parse(a.createdAt)),
      ) ||
      (r &&
        (!workbookSnapshotsEqual(
          t,
          buildCanonicalWorkbookSnapshot(r.workbook),
        ) ||
          !workbookSnapshotsEqual(n, captureWorkbookView(r.workbook, r.view))))
    )
      throw invalidStorageAckError();
    return {
      ...e,
      workbook: t,
      view: n,
      guard: { ...o },
      checkpoints: e.checkpoints.map((a) => ({ ...a })),
    };
  } catch {
    throw invalidStorageAckError();
  }
}
function instrumentWorkbookPersistence(e, r) {
  const t = (n, o) =>
    n && typeof n.then == "function"
      ? Promise.resolve(n).then((a) => normalizeCheckpointRecord(a, o))
      : normalizeCheckpointRecord(n, o);
  return {
    ...("coordination" in e ? { coordination: e.coordination } : {}),
    load: () => {
      const n = r.start("persistence", "load");
      try {
        const o = e.load(),
          a = o ? normalizeCheckpointRecord(o) : null;
        return (
          n.finish(a != null && a.recovered ? "recovered" : "succeeded"),
          a
        );
      } catch (o) {
        throw (n.fail(o), o);
      }
    },
    save: (n, o, a, s) =>
      r.run("persistence", "save", () => {
        const d = {
          workbook: buildCanonicalWorkbookSnapshot(n),
          view: captureWorkbookView(n, o),
        };
        return t(
          e.save(d.workbook, structuredClone(d.view), a ? { ...a } : null, s),
          d,
        );
      }),
    checkpoint: (n, o, a) =>
      r.run("persistence", "checkpoint", () => t(e.checkpoint(n, o, a))),
    listCheckpoints: () =>
      r.run("persistence", "list", () => e.listCheckpoints()),
    compareCheckpoint: (n, o) =>
      r.run("persistence", "compare", () => e.compareCheckpoint(n, o)),
    restoreCheckpoint: (n, o, a, s, d) =>
      r.run("persistence", "restore", () =>
        t(e.restoreCheckpoint(n, o, a, s, d)),
      ),
  };
}
function findRecoverableCheckpoint(e, r, t) {
  const n = e.load();
  if (!n) return null;
  const o = normalizeCheckpointRecord(n);
  return o.recovered ||
    !workbookSnapshotsEqual(o.workbook, buildCanonicalWorkbookSnapshot(r)) ||
    t
    ? null
    : o;
}
const diagnosticGroupLabels = {
  persistence: "Saving and recovery",
  calculation: "Calculation",
  file: "File processing",
  provider: "AI provider",
  source: "External sources",
  host: "Application callbacks",
};
function WorkbookHealthDialog({ monitor, onClose }) {
  const t = ne.useSyncExternalStore(
      monitor.subscribe,
      monitor.getSnapshot,
      monitor.getSnapshot,
    ),
    n = ne.useId(),
    o = ne.useRef(null),
    a = useWorkbookMessages();
  ne.useEffect(() => {
    var u;
    const d = document.activeElement;
    return (
      (u = o.current) == null || u.focus(),
      () => {
        d != null && d.isConnected && d.focus();
      }
    );
  }, []);
  const s = () => {
    const d = URL.createObjectURL(
        new Blob(
          [
            JSON.stringify(
              {
                format: "ledgerglass-reliability-v1",
                ...monitor.getSnapshot(),
              },
              null,
              2,
            ),
          ],
          { type: "application/json" },
        ),
      ),
      u = document.createElement("a");
    ((u.href = d),
      (u.download = "Workbook-diagnostics.json"),
      u.click(),
      setTimeout(() => URL.revokeObjectURL(d), 1e3));
  };
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      className: "lg-reliability",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      onKeyDown: (d) => {
        var u, l;
        if (
          (d.key === "Escape" && (d.stopPropagation(), onClose()),
          d.key === "Tab")
        ) {
          const c = [
            ...d.currentTarget.querySelectorAll("button:not(:disabled)"),
          ];
          d.shiftKey && document.activeElement === c[0]
            ? (d.preventDefault(), (u = c.at(-1)) == null || u.focus())
            : !d.shiftKey &&
              document.activeElement === c.at(-1) &&
              (d.preventDefault(), (l = c[0]) == null || l.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: n,
              children: i.jsx(LocalizedText, { children: "Workbook health" }),
            }),
            i.jsx("button", {
              ref: o,
              type: "button",
              "aria-label": a.text("Close workbook health"),
              onClick: onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Diagnostics contain operation categories, outcomes and timings. Workbook contents, names, prompts and error details are excluded.",
          }),
        }),
        i.jsx("div", {
          className: "lg-reliability-table",
          children: i.jsxs("table", {
            children: [
              i.jsx("thead", {
                children: i.jsx("tr", {
                  children: [
                    "Service",
                    "State",
                    "Completed",
                    "Failures",
                    "Longest operation",
                  ].map((d) =>
                    i.jsx(
                      "th",
                      {
                        scope: "col",
                        children: i.jsx(LocalizedText, { children: d }),
                      },
                      d,
                    ),
                  ),
                }),
              }),
              i.jsx("tbody", {
                children: Object.keys(diagnosticGroupLabels).map((d) => {
                  const u = t.health[d],
                    l = Object.keys(u.issues).length
                      ? "Needs attention"
                      : u.active
                        ? "Working…"
                        : u.completed
                          ? "Ready"
                          : "Not used";
                  return i.jsxs(
                    "tr",
                    {
                      children: [
                        i.jsx("th", {
                          scope: "row",
                          children: i.jsx(LocalizedText, {
                            children: diagnosticGroupLabels[d],
                          }),
                        }),
                        i.jsx("td", {
                          children: i.jsx(LocalizedText, { children: l }),
                        }),
                        i.jsx("td", { children: u.completed }),
                        i.jsx("td", { children: u.failures }),
                        i.jsxs("td", {
                          children: [
                            a.locale
                              ? u.maxElapsedMs.toLocaleString(a.locale)
                              : u.maxElapsedMs,
                            " ",
                            "ms",
                          ],
                        }),
                      ],
                    },
                    d,
                  );
                }),
              }),
            ],
          }),
        }),
        i.jsxs("p", {
          role: "status",
          children: [
            i.jsx(LocalizedText, { children: "Retained events" }),
            ": ",
            t.events.length,
            " ",
            "· ",
            i.jsx(LocalizedText, { children: "Discarded older events" }),
            ":",
            " ",
            t.droppedEvents,
            " ·",
            " ",
            i.jsx(LocalizedText, { children: "Diagnostic delivery failures" }),
            ":",
            " ",
            t.sinkFailures,
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Save failures remain visible in the editor. Export your local changes before reloading a saved copy. A pending save must finish before another retry.",
          }),
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              type: "button",
              onClick: s,
              children: i.jsx(LocalizedText, {
                children: "Download diagnostics",
              }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: () => monitor.clearEvents(),
              children: i.jsx(LocalizedText, {
                children: "Clear diagnostic history",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const cellSnapshotCache = new WeakMap();
function snapshotSheetContent(e) {
  return fingerprintCanonicalValue({
    cells: Object.fromEntries(
      Object.entries(e.cells).map(([r, t]) => [
        r,
        { ...t, style: undefined, numFmt: undefined, format: undefined },
      ]),
    ),
    tables: e.tables,
    merges: e.merges,
    rowOrigins: e.rowOriginIds,
  });
}
function checkSourceSheetAccess(e, r, t) {
  var o;
  if (t.readOnly) throw new ExternalSourceError("refused");
  if (!r) {
    assertWorkbookStructureEditable(e);
    return;
  }
  const n = e.sheets.find((a) => a.id === r);
  if (
    !n ||
    n.hidden ||
    n.veryHidden ||
    ((o = n.protection) != null && o.enabled) ||
    (t.allowedSheetIds && !t.allowedSheetIds.includes(r))
  )
    throw new ExternalSourceError("refused");
  if (!n.externalConnection) throw new ExternalSourceError("invalid");
  if (snapshotSheetContent(n) !== n.externalConnection.outputHash)
    throw new ExternalSourceError("stale");
}
function EH(e, r, t, n, o = {}) {
  return evaluateExternalSourceQuery(e, r, t, n, o);
}
function evaluateExternalSourceQuery(e, r, t, n, o = {}, a = []) {
  checkSourceSheetAccess(e, n.sheetId, o);
  const s = parseActiveSourceApproval(t),
    d = validateExternalSourcePayload(r, s.source),
    u = n.sheetId
      ? e.sheets.find((h) => h.id === n.sheetId).externalConnection
      : undefined;
  if (u && !isSameContentHash(u.source, d.source))
    throw new ExternalSourceError("schema");
  if (
    !n.sheetId &&
    (e.sheets.filter((h) => h.externalConnection).length >=
      externalSourceLimits.sources ||
      !n.outputName.trim() ||
      n.outputName !== n.outputName.trim() ||
      n.outputName.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(n.outputName) ||
      /^'|'$/.test(n.outputName) ||
      e.sheets.some((h) => h.name.toLowerCase() === n.outputName.toLowerCase()))
  )
    throw new ExternalSourceError("invalid");
  if (
    (n.sheetId || inspectOutputNameSafety(e, n.outputName),
    e.sheets
      .filter((h) => h.id !== n.sheetId && !a.includes(h.id))
      .reduce((h, f) => h + Object.keys(f.cells).length, 0) +
      (d.rows.length + 1) * d.source.columns.length >
      workbookLimits.maxCells)
  )
    throw new ExternalSourceError("limit");
  const c = {
    workbookHash: fingerprintCanonicalValue(e),
    outputName: n.outputName,
    ...(n.sheetId ? { sheetId: n.sheetId } : {}),
    snapshot: d,
    grant: s,
    ...(u
      ? {
          previous: {
            rows: u.lastSuccess.rows,
            controls: structuredClone(u.lastSuccess.controls),
            at: u.lastSuccess.at,
          },
        }
      : {}),
    controls: computeSourceControlTotals(d),
  };
  if (
    c.controls.some((h) => {
      return !Number.isFinite(
        h.sum -
          (c.previous?.controls.find((m) => m.field === h.field)?.sum ?? 0),
      );
    })
  )
    throw new ExternalSourceError("limit");
  return (cellSnapshotCache.set(c, fingerprintCanonicalValue(c)), c);
}
function grantSheetSnapshotSource(e, r, t, n = {}) {
  if (
    !cellSnapshotCache.has(r) ||
    cellSnapshotCache.get(r) !== fingerprintCanonicalValue(r) ||
    fingerprintCanonicalValue(e) !== r.workbookHash
  )
    throw new ExternalSourceError("stale");
  checkSourceSheetAccess(e, r.sheetId, n);
  const o = parseActiveSourceApproval(t);
  if (
    (parseActiveSourceApproval(r.grant),
    !isSameContentHash(o.source, r.snapshot.source))
  )
    throw new ExternalSourceError("schema");
  EH(e, r.snapshot, o, r, n);
  const a = buildSourceSheetSnapshot(e, r),
    s = normalizeWorkbookInput({
      ...e,
      sheets: r.sheetId
        ? e.sheets.map((d) => (d.id === a.id ? a : d))
        : [...e.sheets, a],
    });
  return (cellSnapshotCache.delete(r), s);
}
function buildSourceSheetSnapshot(e, r) {
  const t = e.sheets.find((d) => d.id === r.sheetId),
    n = t ? { ...t } : createWorksheetRecord(r.outputName),
    o = {};
  ([r.snapshot.source.columns.map((d) => d.name), ...r.snapshot.rows].forEach(
    (d, u) =>
      d.forEach((l, c) => {
        const h = `${columnIndexToLetters(c)}${u + 1}`,
          f = t?.cells[h];
        o[h] = {
          raw: l,
          ...(typeof l == "string" && /^[=+@-]/.test(l)
            ? { literal: true }
            : {}),
          ...(f != null && f.style
            ? { style: structuredClone(f.style) }
            : u === 0
              ? { style: { bold: true, fill: "#eef2ff" } }
              : {}),
          ...(f != null && f.numFmt ? { numFmt: f.numFmt } : {}),
          ...(f != null && f.format ? { format: f.format } : {}),
          note: `Source ${r.snapshot.source.id} · version ${r.snapshot.version} · ${u === 0 ? "header" : `record ${u}`} · as of ${r.snapshot.sourceTimestamp}`,
        };
      }),
  ),
    (n.cells = o));
  const s = new Date().toISOString();
  return (
    (n.externalConnection = {
      protocol: externalSourceFormatVersion,
      source: structuredClone(r.snapshot.source),
      lastSuccess: {
        at: s,
        sourceTimestamp: r.snapshot.sourceTimestamp,
        version: r.snapshot.version,
        rows: r.snapshot.rows.length,
        controls: structuredClone(r.controls),
        dataHash: fingerprintCanonicalValue(r.snapshot),
      },
      lastAttempt: { at: s, outcome: "succeeded" },
      outputHash: snapshotSheetContent(n),
    }),
    n
  );
}
const sourceGrantCache = new WeakMap();
function assertSourceGrantsValid(e, r, t = {}) {
  if (
    !r.length ||
    r.length > externalSourceLimits.sources ||
    new Set(r).size !== r.length
  )
    throw new ExternalSourceError("invalid");
  for (const n of r) checkSourceSheetAccess(e, n, t);
}
function planSourceGrants(e, r, t = {}) {
  const n = r.map((d) => d.sheetId);
  if (
    (assertSourceGrantsValid(e, n, t),
    new TextEncoder().encode(JSON.stringify(r)).length >
      externalSourceLimits.responseBytes)
  )
    throw new ExternalSourceError("limit");
  const o = r.map((d) =>
    evaluateExternalSourceQuery(
      e,
      d.snapshot,
      d.grant,
      {
        sheetId: d.sheetId,
        outputName: e.sheets.find((u) => u.id === d.sheetId).name,
      },
      t,
      n,
    ),
  );
  if (
    e.sheets
      .filter((d) => !n.includes(d.id))
      .reduce((d, u) => d + Object.keys(u.cells).length, 0) +
      o.reduce(
        (d, u) =>
          d + (u.snapshot.rows.length + 1) * u.snapshot.source.columns.length,
        0,
      ) >
    workbookLimits.maxCells
  )
    throw new ExternalSourceError("limit");
  const s = { workbookHash: fingerprintCanonicalValue(e), plans: o };
  return (sourceGrantCache.set(s, fingerprintCanonicalValue(s)), s);
}
function applySourceGrantPlans(e, r, t, n = {}) {
  if (
    !sourceGrantCache.has(r) ||
    sourceGrantCache.get(r) !== fingerprintCanonicalValue(r) ||
    r.workbookHash !== fingerprintCanonicalValue(e)
  )
    throw new ExternalSourceError("stale");
  const o = t.map((u) => parseActiveSourceApproval(u)),
    a = new Set(r.plans.map((u) => u.snapshot.source.id));
  if (
    o.length !== a.size ||
    new Set(o.map((u) => u.source.id)).size !== o.length ||
    o.some((u) => !a.has(u.source.id))
  )
    throw new ExternalSourceError("refused");
  for (const u of r.plans) {
    parseActiveSourceApproval(u.grant);
    const l = o.find((c) => c.source.id === u.snapshot.source.id);
    if (!l || !isSameContentHash(l.source, u.snapshot.source))
      throw new ExternalSourceError("schema");
  }
  planSourceGrants(
    e,
    r.plans.map((u) => ({
      sheetId: u.sheetId,
      snapshot: u.snapshot,
      grant: o.find((l) => l.source.id === u.snapshot.source.id),
    })),
    n,
  );
  const s = new Map(
      r.plans.map((u) => [u.sheetId, buildSourceSheetSnapshot(e, u)]),
    ),
    d = normalizeWorkbookInput({
      ...e,
      sheets: e.sheets.map((u) => s.get(u.id) ?? u),
    });
  sourceGrantCache.delete(r);
  for (const u of r.plans) cellSnapshotCache.delete(u);
  return d;
}
function normalizeExternalSourceState(e, r, t, n, o = {}) {
  var u;
  if (o.readOnly) return e;
  const a = e.sheets.find((l) => l.id === r);
  if (
    !(a != null && a.externalConnection) ||
    a.externalConnection.source.id !== t ||
    a.hidden ||
    a.veryHidden ||
    ((u = a.protection) != null && u.enabled) ||
    (o.allowedSheetIds && !o.allowedSheetIds.includes(r))
  )
    return e;
  const s = a.externalConnection,
    d = new Date(
      Math.max(Date.now(), Date.parse(s.lastSuccess.at)),
    ).toISOString();
  return normalizeWorkbookInput({
    ...e,
    sheets: e.sheets.map((l) =>
      l === a
        ? {
            ...l,
            externalConnection: {
              ...s,
              lastAttempt: {
                at: d,
                outcome: n === "cancelled" ? "cancelled" : "failed",
                code: n,
              },
            },
          }
        : l,
    ),
  });
}
async function runSourceWithTimeout(e, r, t = {}) {
  var l, c;
  const n = t.timeoutMs ?? externalSourceLimits.timeoutMs;
  if (!Number.isInteger(n) || n < 1 || n > externalSourceLimits.timeoutMs)
    throw new ExternalSourceError("invalid");
  const o = t.reliability?.start("source", e),
    a = new AbortController();
  let s,
    d = () => {};
  try {
    if ((l = t.signal) != null && l.aborted)
      throw new ExternalSourceError("cancelled");
    const h = new Promise((p, m) => {
        var v;
        ((d = () => {
          (a.abort(), m(new ExternalSourceError("cancelled")));
        }),
          (v = t.signal) == null ||
            v.addEventListener("abort", d, { once: true }),
          (s = setTimeout(() => {
            (a.abort(), m(new ExternalSourceError("timeout")));
          }, n)));
      }),
      f = await Promise.race([Promise.resolve().then(() => r(a.signal)), h]);
    if (a.signal.aborted) throw new ExternalSourceError("cancelled");
    return (o == null || o.finish(), f);
  } catch (h) {
    const f = isExternalSourceError(h);
    throw (o == null || o.fail(f), f);
  } finally {
    (clearTimeout(s),
      (c = t.signal) == null || c.removeEventListener("abort", d));
  }
}
async function fetchSourceCatalog(e, r = {}) {
  return runSourceWithTimeout(
    "catalog",
    async (t) => {
      const n = await e.catalog({ signal: t });
      if (!Array.isArray(n) || n.length > externalSourceLimits.sources)
        throw new ExternalSourceError("invalid");
      const o = n.map((a) => parseActiveSourceApproval(a));
      if (new Set(o.map((a) => a.source.id)).size !== o.length)
        throw new ExternalSourceError("invalid");
      return o;
    },
    r,
  );
}
async function authorizeSourceAccess(e, r, t = {}) {
  return runSourceWithTimeout(
    "authorize",
    async (n) => {
      const o = parseActiveSourceApproval(await e.authorize(r, { signal: n }));
      if (o.source.id !== r) throw new ExternalSourceError("refused");
      return o;
    },
    t,
  );
}
async function refreshSourceSnapshot(e, r, t = {}) {
  const n = await authorizeSourceAccess(e, r, t),
    o = await runSourceWithTimeout(
      "refresh",
      async (a) => {
        const s = validateExternalSourcePayload(
          await e.refresh(r, { signal: a }),
          n.source,
        );
        return (parseActiveSourceApproval(n), s);
      },
      t,
    );
  return { grant: n, snapshot: o };
}
discriminatedUnion("operation", [
  strictObject({ operation: literal("catalog") }),
  strictObject({
    operation: _enum(["authorize", "refresh"]),
    sourceId: string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  }),
]);
async function runSourceRefreshBatch(e, r, t, n = {}) {
  assertSourceGrantsValid(e, t, n);
  const o = [],
    a = new Map();
  let s;
  try {
    await runSourceWithTimeout(
      "refresh",
      async (d) => {
        var u;
        for (const l of t) {
          if (d.aborted) throw new ExternalSourceError("cancelled");
          const c = e.sheets.find((f) => f.id === l).externalConnection.source
            .id;
          let h = a.get(c);
          if (!h) {
            try {
              h = await refreshSourceSnapshot(r, c, { ...n, signal: d });
            } catch (f) {
              h = isExternalSourceError(f).code;
            }
            if (d.aborted) throw new ExternalSourceError("cancelled");
            a.set(c, h);
          }
          if (
            (o.push(
              typeof h == "string"
                ? { sheetId: l, sourceId: c, error: h }
                : { sheetId: l, sourceId: c, ...h },
            ),
            new TextEncoder().encode(JSON.stringify(o)).length >
              externalSourceLimits.responseBytes)
          )
            throw new ExternalSourceError("limit");
          (u = n.onProgress) == null || u.call(n, o.length, t.length);
        }
      },
      n,
    );
  } catch (d) {
    s = isExternalSourceError(d).code;
  }
  if (s) {
    for (const d of t)
      o.some((u) => u.sheetId === d) ||
        o.push({
          sheetId: d,
          sourceId: e.sheets.find((u) => u.id === d).externalConnection.source
            .id,
          error: s,
        });
    return s === "limit"
      ? {
          entries: o.map(({ sheetId, sourceId }) => ({
            sheetId,
            sourceId,
            error: s,
          })),
          error: s,
        }
      : { entries: o, error: s };
  }
  if (o.some((d) => d.error)) return { entries: o };
  try {
    const d = planSourceGrants(
      e,
      o.map((u) => ({
        sheetId: u.sheetId,
        snapshot: u.snapshot,
        grant: u.grant,
      })),
      n,
    );
    return { entries: o, plan: d };
  } catch (d) {
    return { entries: o, error: isExternalSourceError(d).code };
  }
}
async function G2e(e, r, t = {}) {
  return runSourceWithTimeout(
    "authorize",
    async (n) => {
      const o = [];
      for (const a of new Set(r.plans.map((s) => s.snapshot.source.id))) {
        if (n.aborted) throw new ExternalSourceError("cancelled");
        o.push(await authorizeSourceAccess(e, a, { ...t, signal: n }));
      }
      return o.map((a) => parseActiveSourceApproval(a));
    },
    t,
  );
}

export {
  instrumentWorkbookPersistence,
  computeVirtualizedRange,
  snapshotSheetContent,
  EH,
  G2e,
  applySourceGrantPlans,
  WorkbookSaveConflictDialog,
  authorizeSourceAccess,
  refreshSourceSnapshot,
  WorkbookHealthDialog,
  grantSheetSnapshotSource,
  runSourceRefreshBatch,
  normalizeExternalSourceState,
  workbookLockName,
  fetchSourceCatalog,
  createWorkbookDiagnosticsMonitor,
  saveWorkbookWithLock,
  findRecoverableCheckpoint,
};
