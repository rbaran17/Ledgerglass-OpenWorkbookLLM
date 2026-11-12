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
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
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
} from "./11-local-workbook-store-and-versions.js";
import { inspectOutputNameSafety } from "./12-consolidation-and-dependency-inspection.js";
import {
  resolveAxisScale,
  physicalToLogicalPixels,
} from "./53-find-replace-go-to.js";

function computeVirtualizedRange(e, r, t, zoomScale, o = 5) {
  const axisScale = resolveAxisScale(e.total, t, zoomScale),
    s = clampNumber(r, 0, axisScale.physicalMax),
    logicalOffset = physicalToLogicalPixels(s, e.total, t, zoomScale),
    first = e.count ? Math.max(0, e.indexAt(logicalOffset) - o) : 0,
    last = e.count
      ? Math.min(
          e.count,
          e.indexAt(logicalOffset + axisScale.viewport / zoomScale) + o + 1,
        )
      : 0,
    leading = Math.max(0, s / zoomScale + e.offset(first) - logicalOffset),
    trailing = Math.max(
      0,
      axisScale.physicalPixels / zoomScale -
        leading -
        (e.offset(last) - e.offset(first)),
    );
  return {
    first: first,
    last: last,
    leading: leading,
    trailing: trailing,
    logicalOffset: logicalOffset,
    total: axisScale.physicalPixels / zoomScale,
  };
}
function workbookLockName(embeddedDemoKey) {
  return `ledgerglass:workbook:${embeddedDemoKey}`;
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
  const localWorkbookStore = createLocalWorkbookStore(e, r, {
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
            l?.removeEventListener("abort", F));
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
      if (l?.aborted) {
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
        (l?.addEventListener("abort", F, { once: true }),
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
          } catch (error) {
            throw ((b = true), (y = error), error);
          }
        });
        Promise.resolve(j).then((z) => {
          ((v = false), A(), c(z));
        }, O);
      } catch (error) {
        O(error);
      }
    });
  }
  return {
    coordination: "web-locks",
    load: localWorkbookStore.load,
    listCheckpoints: localWorkbookStore.listCheckpoints,
    compareCheckpoint: localWorkbookStore.compareCheckpoint,
    async save(workbook, view, expected, h = {}) {
      const f = cloneSaveInputs({
          workbook: workbook,
          view: view,
          expected: expected,
          options: { allowIdentityChange: h.allowIdentityChange },
        }),
        signal = h.signal;
      return d(
        () =>
          localWorkbookStore.save(f.workbook, f.view, f.expected, f.options),
        signal,
      );
    },
    async checkpoint(u, expected, c = {}) {
      const h = cloneSaveInputs({ label: u, expected: expected }),
        signal = c.signal;
      return d(
        () => localWorkbookStore.checkpoint(h.label, h.expected),
        signal,
      );
    },
    async restoreCheckpoint(u, workbook, expected, actor, f = {}) {
      const p = cloneSaveInputs({
          id: u,
          workbook: workbook,
          expected: expected,
          actor: actor,
        }),
        signal = f.signal;
      return d(
        () =>
          localWorkbookStore.restoreCheckpoint(
            p.id,
            p.workbook,
            p.expected,
            p.actor,
          ),
        signal,
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
function describeCellContent(cell) {
  return cell
    ? cell.raw === null || cell.raw === ""
      ? "Blank"
      : typeof cell.raw == "boolean"
        ? cell.raw
          ? "TRUE"
          : "FALSE"
        : typeof cell.raw == "string"
          ? (cell.literal
              ? "Text: "
              : cell.raw.startsWith("=")
                ? "Formula: "
                : "") + truncateDisplayText(cell.raw)
          : String(cell.raw)
    : "Empty cell";
}
function collectConflictChanges(e) {
  const r = [],
    t = (path, before, after, s) => {
      if (stableSerialize(before) !== stableSerialize(after)) {
        if (
          s < 3 &&
          (wp(before) || wp(after)) &&
          (before === undefined || wp(before)) &&
          (after === undefined || wp(after))
        ) {
          const d = wp(before) ? before : {},
            u = wp(after) ? after : {},
            l = new Set([...Object.keys(d), ...Object.keys(u)]);
          if (l.size) {
            for (const c of l) t([...path, c], d[c], u[c], s + 1);
            return;
          }
        }
        r.push({ path: path, before: before, after: after });
      }
    };
  return (t(e.path.split("."), e.before, e.after, 0), r);
}
function WorkbookSaveConflictDialog(e) {
  const { text, language } = useWorkbookMessages(),
    [n, o] = ne.useState(null),
    [a, s] = ne.useState(null),
    [d, u] = ne.useState(""),
    [checked, c] = ne.useState(false),
    [open, f] = ne.useState(false),
    p = ne.useRef(null),
    onClick = () => {
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
      } catch (error) {
        (o(null),
          s(null),
          u(
            error instanceof Error
              ? error.message
              : "The saved workbook cannot be read.",
          ));
      }
    };
  (ne.useEffect(onClick, [e.workbook]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        p.current?.querySelector("button")?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const v = (w) =>
      e.workbook.sheets.find((sheet) => sheet.id === w)?.name ??
      n?.workbook.sheets.find((sheet) => sheet.id === w)?.name ??
      "Removed worksheet",
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
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(), event.stopPropagation(), e.onClose()),
          event.key === "Tab")
        ) {
          const A = [
              ...(p.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
              ) ?? []),
            ],
            k = A[0],
            F = A.at(-1);
          event.shiftKey && document.activeElement === k
            ? (event.preventDefault(), F == null || F.focus())
            : !event.shiftKey &&
              document.activeElement === F &&
              (event.preventDefault(), k == null || k.focus());
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
                  onClick: onClick,
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
                            ].filter(
                              (A) =>
                                A !== "raw" &&
                                stableSerialize(w.before?.[A]) !==
                                  stableSerialize(w.after?.[A]),
                            );
                            return i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("th", {
                                    scope: "row",
                                    children: [v(w.sheetId), "!", w.address],
                                  }),
                                  [w.before, w.after].map((cell, k) =>
                                    i.jsxs(
                                      "td",
                                      {
                                        children: [
                                          i.jsx("span", {
                                            className: "pc-cell-value",
                                            children: describeCellContent(cell),
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
                                                        cell?.[F],
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
                      open: open,
                      onToggle: (event) => f(event.currentTarget.open),
                      children: [
                        i.jsx("summary", {
                          children: i.jsx(LocalizedText, {
                            children: "Full comparison details",
                          }),
                        }),
                        open &&
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
                  checked: checked,
                  disabled: !n,
                  onChange: (event) => c(event.target.checked),
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
                  disabled: !n || !checked,
                  onClick: () => {
                    try {
                      e.onReload(n);
                    } catch (error) {
                      (c(false),
                        u(
                          error instanceof Error
                            ? error.message
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
    Object.keys(diagnosticGroupDefinitions).map((diagnosticGroupDefinition) => [
      diagnosticGroupDefinition,
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
        const w = e.onEvent?.(m);
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
    start = (adapter, operation) => {
      if (
        !Object.hasOwn(diagnosticGroupDefinitions, adapter) ||
        !diagnosticGroupDefinitions[adapter].includes(operation)
      )
        throw new Error("Unknown diagnostic adapter or operation.");
      const attempt = ++a,
        b = performance.now(),
        y = `${adapter}:${operation}`;
      let w = false;
      f(
        {
          version: 1,
          sequence: ++o,
          attempt: attempt,
          adapter: adapter,
          operation: operation,
          phase: "started",
        },
        { ...d.health[adapter], active: d.health[adapter].active + 1 },
      );
      const finish = (outcome = "succeeded", k) => {
        if (w) return;
        ((w = true),
          ["succeeded", "failed", "cancelled", "recovered"].includes(outcome) ||
            (outcome = "failed"));
        const code = k && wH.has(k) ? k : "unknown",
          elapsedMs = Math.max(
            0,
            Math.round((performance.now() - b) * 10) / 10,
          ),
          O = d.health[adapter],
          j = { ...O.issues };
        (attempt >= (n.get(y) ?? 0) &&
          outcome !== "cancelled" &&
          (n.set(y, attempt),
          outcome === "failed" ? (j[operation] = code) : delete j[operation]),
          f(
            {
              version: 1,
              sequence: ++o,
              attempt: attempt,
              adapter: adapter,
              operation: operation,
              phase: "completed",
              outcome: outcome,
              ...(outcome === "failed"
                ? { code: code }
                : outcome === "cancelled"
                  ? { code: "cancelled" }
                  : {}),
              elapsedMs: elapsedMs,
            },
            {
              active: Math.max(0, O.active - 1),
              completed: O.completed + 1,
              failures: O.failures + +(outcome === "failed"),
              cancellations: O.cancellations + +(outcome === "cancelled"),
              maxElapsedMs: Math.max(O.maxElapsedMs, elapsedMs),
              totalElapsedMs: O.totalElapsedMs + elapsedMs,
              issues: Object.freeze(j),
            },
          ));
      };
      return {
        finish: finish,
        fail(A) {
          const k = isAbortOrQuotaError(A);
          finish(k === "cancelled" ? "cancelled" : "failed", k);
        },
      };
    };
  return {
    start: start,
    run(m, v, g) {
      const b = start(m, v);
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
      } catch (error) {
        throw (b.fail(error), error);
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
function captureWorkbookView(workbook, view) {
  const sheet = workbook.sheets.find(
    (sheet) => !sheet.hidden && sheet.id === view.activeSheetId,
  );
  if (!sheet || view.selection.sheetId !== sheet.id)
    throw invalidStorageAckError();
  const rangeAddress = parseRangeAddress(view.selection.range);
  if (
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount
  )
    throw invalidStorageAckError();
  const o = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    a = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  return {
    activeSheetId: sheet.id,
    selection: { sheetId: sheet.id, range: o === a ? o : `${o}:${a}` },
  };
}
function normalizeCheckpointRecord(e, r) {
  try {
    const canonicalWorkbookSnapshot = buildCanonicalWorkbookSnapshot(
        e.workbook,
      ),
      view = captureWorkbookView(canonicalWorkbookSnapshot, e.view),
      guard = e.guard;
    if (
      !guard ||
      typeof guard.namespace != "string" ||
      !guard.namespace ||
      !Number.isSafeInteger(guard.generation) ||
      guard.generation < 0 ||
      typeof guard.digest != "string" ||
      !guard.digest ||
      guard.workbookId !== canonicalWorkbookSnapshot.id ||
      guard.revision !== canonicalWorkbookSnapshot.revision ||
      typeof e.recovered != "boolean" ||
      !Array.isArray(e.checkpoints) ||
      e.checkpoints.length > 20 ||
      e.checkpoints.some(
        (checkpoint) =>
          !checkpoint ||
          typeof checkpoint.id != "string" ||
          typeof checkpoint.label != "string" ||
          typeof checkpoint.workbookId != "string" ||
          !Number.isSafeInteger(checkpoint.revision) ||
          checkpoint.revision < 0 ||
          typeof checkpoint.createdAt != "string" ||
          !Number.isFinite(Date.parse(checkpoint.createdAt)),
      ) ||
      (r &&
        (!workbookSnapshotsEqual(
          canonicalWorkbookSnapshot,
          buildCanonicalWorkbookSnapshot(r.workbook),
        ) ||
          !workbookSnapshotsEqual(
            view,
            captureWorkbookView(r.workbook, r.view),
          )))
    )
      throw invalidStorageAckError();
    return {
      ...e,
      workbook: canonicalWorkbookSnapshot,
      view: view,
      guard: { ...guard },
      checkpoints: e.checkpoints.map((checkpoint) => ({ ...checkpoint })),
    };
  } catch {
    throw invalidStorageAckError();
  }
}
function instrumentWorkbookPersistence(e, activeReliabilityMonitor) {
  const t = (n, o) =>
    n && typeof n.then == "function"
      ? Promise.resolve(n).then((a) => normalizeCheckpointRecord(a, o))
      : normalizeCheckpointRecord(n, o);
  return {
    ...("coordination" in e ? { coordination: e.coordination } : {}),
    load: () => {
      const n = activeReliabilityMonitor.start("persistence", "load");
      try {
        const o = e.load(),
          a = o ? normalizeCheckpointRecord(o) : null;
        return (n.finish(a?.recovered ? "recovered" : "succeeded"), a);
      } catch (error) {
        throw (n.fail(error), error);
      }
    },
    save: (workbook, view, a, s) =>
      activeReliabilityMonitor.run("persistence", "save", () => {
        const d = {
          workbook: buildCanonicalWorkbookSnapshot(workbook),
          view: captureWorkbookView(workbook, view),
        };
        return t(
          e.save(d.workbook, structuredClone(d.view), a ? { ...a } : null, s),
          d,
        );
      }),
    checkpoint: (n, o, a) =>
      activeReliabilityMonitor.run("persistence", "checkpoint", () =>
        t(e.checkpoint(n, o, a)),
      ),
    listCheckpoints: () =>
      activeReliabilityMonitor.run("persistence", "list", () =>
        e.listCheckpoints(),
      ),
    compareCheckpoint: (n, o) =>
      activeReliabilityMonitor.run("persistence", "compare", () =>
        e.compareCheckpoint(n, o),
      ),
    restoreCheckpoint: (n, o, a, s, d) =>
      activeReliabilityMonitor.run("persistence", "restore", () =>
        t(e.restoreCheckpoint(n, o, a, s, d)),
      ),
  };
}
function findRecoverableCheckpoint(workbookPersistence, book, t) {
  const n = workbookPersistence.load();
  if (!n) return null;
  const checkpointRecord = normalizeCheckpointRecord(n);
  return checkpointRecord.recovered ||
    !workbookSnapshotsEqual(
      checkpointRecord.workbook,
      buildCanonicalWorkbookSnapshot(book),
    ) ||
    t
    ? null
    : checkpointRecord;
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
    workbookMessages = useWorkbookMessages();
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      o.current?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const onClick = () => {
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
      onKeyDown: (event) => {
        var u, l;
        if (
          (event.key === "Escape" && (event.stopPropagation(), onClose()),
          event.key === "Tab")
        ) {
          const c = [
            ...event.currentTarget.querySelectorAll("button:not(:disabled)"),
          ];
          event.shiftKey && document.activeElement === c[0]
            ? (event.preventDefault(), (u = c.at(-1)) == null || u.focus())
            : !event.shiftKey &&
              document.activeElement === c.at(-1) &&
              (event.preventDefault(), (l = c[0]) == null || l.focus());
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
              "aria-label": workbookMessages.text("Close workbook health"),
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
                children: Object.keys(diagnosticGroupLabels).map(
                  (diagnosticGroupLabel) => {
                    const u = t.health[diagnosticGroupLabel],
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
                              children:
                                diagnosticGroupLabels[diagnosticGroupLabel],
                            }),
                          }),
                          i.jsx("td", {
                            children: i.jsx(LocalizedText, { children: l }),
                          }),
                          i.jsx("td", { children: u.completed }),
                          i.jsx("td", { children: u.failures }),
                          i.jsxs("td", {
                            children: [
                              workbookMessages.locale
                                ? u.maxElapsedMs.toLocaleString(
                                    workbookMessages.locale,
                                  )
                                : u.maxElapsedMs,
                              " ",
                              "ms",
                            ],
                          }),
                        ],
                      },
                      diagnosticGroupLabel,
                    );
                  },
                ),
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
              onClick: onClick,
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
function snapshotSheetContent(sheet) {
  return fingerprintCanonicalValue({
    cells: Object.fromEntries(
      Object.entries(sheet.cells).map(([r, t]) => [
        r,
        { ...t, style: undefined, numFmt: undefined, format: undefined },
      ]),
    ),
    tables: sheet.tables,
    merges: sheet.merges,
    rowOrigins: sheet.rowOriginIds,
  });
}
function checkSourceSheetAccess(workbook, sheetId, t) {
  if (t.readOnly) throw new ExternalSourceError("refused");
  if (!sheetId) {
    assertWorkbookStructureEditable(workbook);
    return;
  }
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (
    !sheet ||
    sheet.hidden ||
    sheet.veryHidden ||
    sheet.protection?.enabled ||
    (t.allowedSheetIds && !t.allowedSheetIds.includes(sheetId))
  )
    throw new ExternalSourceError("refused");
  if (!sheet.externalConnection) throw new ExternalSourceError("invalid");
  if (snapshotSheetContent(sheet) !== sheet.externalConnection.outputHash)
    throw new ExternalSourceError("stale");
}
function EH(workbook, snapshot, grant, n, o = {}) {
  return evaluateExternalSourceQuery(workbook, snapshot, grant, n, o);
}
function evaluateExternalSourceQuery(
  workbook,
  snapshot,
  grant,
  n,
  o = {},
  a = [],
) {
  checkSourceSheetAccess(workbook, n.sheetId, o);
  const activeSourceApproval = parseActiveSourceApproval(grant),
    d = validateExternalSourcePayload(snapshot, activeSourceApproval.source),
    u = n.sheetId
      ? workbook.sheets.find((sheet) => sheet.id === n.sheetId)
          .externalConnection
      : undefined;
  if (u && !isSameContentHash(u.source, d.source))
    throw new ExternalSourceError("schema");
  if (
    !n.sheetId &&
    (workbook.sheets.filter((sheet) => sheet.externalConnection).length >=
      externalSourceLimits.sources ||
      !n.outputName.trim() ||
      n.outputName !== n.outputName.trim() ||
      n.outputName.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(n.outputName) ||
      /^'|'$/.test(n.outputName) ||
      workbook.sheets.some(
        (sheet) => sheet.name.toLowerCase() === n.outputName.toLowerCase(),
      ))
  )
    throw new ExternalSourceError("invalid");
  if (
    (n.sheetId || inspectOutputNameSafety(workbook, n.outputName),
    workbook.sheets
      .filter((sheet) => sheet.id !== n.sheetId && !a.includes(sheet.id))
      .reduce((h, f) => h + Object.keys(f.cells).length, 0) +
      (d.rows.length + 1) * d.source.columns.length >
      workbookLimits.maxCells)
  )
    throw new ExternalSourceError("limit");
  const c = {
    workbookHash: fingerprintCanonicalValue(workbook),
    outputName: n.outputName,
    ...(n.sheetId ? { sheetId: n.sheetId } : {}),
    snapshot: d,
    grant: activeSourceApproval,
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
    c.controls.some(
      (control) =>
        !Number.isFinite(
          control.sum -
            (c.previous?.controls.find((m) => m.field === control.field)?.sum ??
              0),
        ),
    )
  )
    throw new ExternalSourceError("limit");
  return (cellSnapshotCache.set(c, fingerprintCanonicalValue(c)), c);
}
function grantSheetSnapshotSource(workbook, r, t, n = {}) {
  if (
    !cellSnapshotCache.has(r) ||
    cellSnapshotCache.get(r) !== fingerprintCanonicalValue(r) ||
    fingerprintCanonicalValue(workbook) !== r.workbookHash
  )
    throw new ExternalSourceError("stale");
  checkSourceSheetAccess(workbook, r.sheetId, n);
  const activeSourceApproval = parseActiveSourceApproval(t);
  if (
    (parseActiveSourceApproval(r.grant),
    !isSameContentHash(activeSourceApproval.source, r.snapshot.source))
  )
    throw new ExternalSourceError("schema");
  EH(workbook, r.snapshot, activeSourceApproval, r, n);
  const sourceSheetSnapshot = buildSourceSheetSnapshot(workbook, r),
    workbookInput = normalizeWorkbookInput({
      ...workbook,
      sheets: r.sheetId
        ? workbook.sheets.map((sheet) =>
            sheet.id === sourceSheetSnapshot.id ? sourceSheetSnapshot : sheet,
          )
        : [...workbook.sheets, sourceSheetSnapshot],
    });
  return (cellSnapshotCache.delete(r), workbookInput);
}
function buildSourceSheetSnapshot(workbook, plan) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === plan.sheetId),
    n = sheet ? { ...sheet } : createWorksheetRecord(plan.outputName),
    o = {};
  ([
    plan.snapshot.source.columns.map((column) => column.name),
    ...plan.snapshot.rows,
  ].forEach((d, u) =>
    d.forEach((raw, c) => {
      const h = `${columnIndexToLetters(c)}${u + 1}`,
        f = sheet?.cells[h];
      o[h] = {
        raw: raw,
        ...(typeof raw == "string" && /^[=+@-]/.test(raw)
          ? { literal: true }
          : {}),
        ...(f?.style
          ? { style: structuredClone(f.style) }
          : u === 0
            ? { style: { bold: true, fill: "#eef2ff" } }
            : {}),
        ...(f?.numFmt ? { numFmt: f.numFmt } : {}),
        ...(f?.format ? { format: f.format } : {}),
        note: `Source ${plan.snapshot.source.id} · version ${plan.snapshot.version} · ${u === 0 ? "header" : `record ${u}`} · as of ${plan.snapshot.sourceTimestamp}`,
      };
    }),
  ),
    (n.cells = o));
  const s = new Date().toISOString();
  return (
    (n.externalConnection = {
      protocol: externalSourceFormatVersion,
      source: structuredClone(plan.snapshot.source),
      lastSuccess: {
        at: s,
        sourceTimestamp: plan.snapshot.sourceTimestamp,
        version: plan.snapshot.version,
        rows: plan.snapshot.rows.length,
        controls: structuredClone(plan.controls),
        dataHash: fingerprintCanonicalValue(plan.snapshot),
      },
      lastAttempt: { at: s, outcome: "succeeded" },
      outputHash: snapshotSheetContent(n),
    }),
    n
  );
}
const sourceGrantCache = new WeakMap();
function assertSourceGrantsValid(workbook, r, t = {}) {
  if (
    !r.length ||
    r.length > externalSourceLimits.sources ||
    new Set(r).size !== r.length
  )
    throw new ExternalSourceError("invalid");
  for (const sheetId of r) checkSourceSheetAccess(workbook, sheetId, t);
}
function planSourceGrants(workbook, r, t = {}) {
  const n = r.map((d) => d.sheetId);
  if (
    (assertSourceGrantsValid(workbook, n, t),
    new TextEncoder().encode(JSON.stringify(r)).length >
      externalSourceLimits.responseBytes)
  )
    throw new ExternalSourceError("limit");
  const plans = r.map((d) =>
    evaluateExternalSourceQuery(
      workbook,
      d.snapshot,
      d.grant,
      {
        sheetId: d.sheetId,
        outputName: workbook.sheets.find((sheet) => sheet.id === d.sheetId)
          .name,
      },
      t,
      n,
    ),
  );
  if (
    workbook.sheets
      .filter((sheet) => !n.includes(sheet.id))
      .reduce((d, u) => d + Object.keys(u.cells).length, 0) +
      plans.reduce(
        (d, u) =>
          d + (u.snapshot.rows.length + 1) * u.snapshot.source.columns.length,
        0,
      ) >
    workbookLimits.maxCells
  )
    throw new ExternalSourceError("limit");
  const s = { workbookHash: fingerprintCanonicalValue(workbook), plans: plans };
  return (sourceGrantCache.set(s, fingerprintCanonicalValue(s)), s);
}
function applySourceGrantPlans(workbook, r, t, n = {}) {
  if (
    !sourceGrantCache.has(r) ||
    sourceGrantCache.get(r) !== fingerprintCanonicalValue(r) ||
    r.workbookHash !== fingerprintCanonicalValue(workbook)
  )
    throw new ExternalSourceError("stale");
  const o = t.map((u) => parseActiveSourceApproval(u)),
    a = new Set(r.plans.map((plan) => plan.snapshot.source.id));
  if (
    o.length !== a.size ||
    new Set(o.map((u) => u.source.id)).size !== o.length ||
    o.some((u) => !a.has(u.source.id))
  )
    throw new ExternalSourceError("refused");
  for (const plan of r.plans) {
    parseActiveSourceApproval(plan.grant);
    const l = o.find((c) => c.source.id === plan.snapshot.source.id);
    if (!l || !isSameContentHash(l.source, plan.snapshot.source))
      throw new ExternalSourceError("schema");
  }
  planSourceGrants(
    workbook,
    r.plans.map((plan) => ({
      sheetId: plan.sheetId,
      snapshot: plan.snapshot,
      grant: o.find((l) => l.source.id === plan.snapshot.source.id),
    })),
    n,
  );
  const s = new Map(
      r.plans.map((plan) => [
        plan.sheetId,
        buildSourceSheetSnapshot(workbook, plan),
      ]),
    ),
    workbookInput = normalizeWorkbookInput({
      ...workbook,
      sheets: workbook.sheets.map((sheet) => s.get(sheet.id) ?? sheet),
    });
  sourceGrantCache.delete(r);
  for (const plan of r.plans) cellSnapshotCache.delete(plan);
  return workbookInput;
}
function normalizeExternalSourceState(workbook, r, t, code, o = {}) {
  if (o.readOnly) return workbook;
  const a = workbook.sheets.find((sheet) => sheet.id === r);
  if (
    !a?.externalConnection ||
    a.externalConnection.source.id !== t ||
    a.hidden ||
    a.veryHidden ||
    a.protection?.enabled ||
    (o.allowedSheetIds && !o.allowedSheetIds.includes(r))
  )
    return workbook;
  const externalConnection = a.externalConnection,
    d = new Date(
      Math.max(Date.now(), Date.parse(externalConnection.lastSuccess.at)),
    ).toISOString();
  return normalizeWorkbookInput({
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet === a
        ? {
            ...sheet,
            externalConnection: {
              ...externalConnection,
              lastAttempt: {
                at: d,
                outcome: code === "cancelled" ? "cancelled" : "failed",
                code: code,
              },
            },
          }
        : sheet,
    ),
  });
}
async function runSourceWithTimeout(e, r, t = {}) {
  const n = t.timeoutMs ?? externalSourceLimits.timeoutMs;
  if (!Number.isInteger(n) || n < 1 || n > externalSourceLimits.timeoutMs)
    throw new ExternalSourceError("invalid");
  const o = t.reliability?.start("source", e),
    a = new AbortController();
  let s,
    d = () => {};
  try {
    if (t.signal?.aborted) throw new ExternalSourceError("cancelled");
    const h = new Promise((p, m) => {
        ((d = () => {
          (a.abort(), m(new ExternalSourceError("cancelled")));
        }),
          t.signal?.addEventListener("abort", d, { once: true }),
          (s = setTimeout(() => {
            (a.abort(), m(new ExternalSourceError("timeout")));
          }, n)));
      }),
      f = await Promise.race([Promise.resolve().then(() => r(a.signal)), h]);
    if (a.signal.aborted) throw new ExternalSourceError("cancelled");
    return (o?.finish(), f);
  } catch (error) {
    const f = isExternalSourceError(error);
    throw (o?.fail(f), f);
  } finally {
    (clearTimeout(s), t.signal?.removeEventListener("abort", d));
  }
}
async function fetchSourceCatalog(adapter, r = {}) {
  return runSourceWithTimeout(
    "catalog",
    async (signal) => {
      const n = await adapter.catalog({ signal: signal });
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
async function authorizeSourceAccess(adapter, r, t = {}) {
  return runSourceWithTimeout(
    "authorize",
    async (signal) => {
      const activeSourceApproval = parseActiveSourceApproval(
        await adapter.authorize(r, { signal: signal }),
      );
      if (activeSourceApproval.source.id !== r)
        throw new ExternalSourceError("refused");
      return activeSourceApproval;
    },
    t,
  );
}
async function refreshSourceSnapshot(adapter, sourceId, t = {}) {
  const grant = await authorizeSourceAccess(adapter, sourceId, t),
    snapshot = await runSourceWithTimeout(
      "refresh",
      async (signal) => {
        const s = validateExternalSourcePayload(
          await adapter.refresh(sourceId, { signal: signal }),
          grant.source,
        );
        return (parseActiveSourceApproval(grant), s);
      },
      t,
    );
  return { grant: grant, snapshot: snapshot };
}
discriminatedUnion("operation", [
  strictObject({ operation: literal("catalog") }),
  strictObject({
    operation: _enum(["authorize", "refresh"]),
    sourceId: string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  }),
]);
async function runSourceRefreshBatch(workbook, adapter, t, n = {}) {
  assertSourceGrantsValid(workbook, t, n);
  const entries = [],
    a = new Map();
  let s;
  try {
    await runSourceWithTimeout(
      "refresh",
      async (signal) => {
        for (const sheetId of t) {
          if (signal.aborted) throw new ExternalSourceError("cancelled");
          const sourceId = workbook.sheets.find((sheet) => sheet.id === sheetId)
            .externalConnection.source.id;
          let h = a.get(sourceId);
          if (!h) {
            try {
              h = await refreshSourceSnapshot(adapter, sourceId, {
                ...n,
                signal: signal,
              });
            } catch (error) {
              h = isExternalSourceError(error).code;
            }
            if (signal.aborted) throw new ExternalSourceError("cancelled");
            a.set(sourceId, h);
          }
          if (
            (entries.push(
              typeof h == "string"
                ? { sheetId: sheetId, sourceId: sourceId, error: h }
                : { sheetId: sheetId, sourceId: sourceId, ...h },
            ),
            new TextEncoder().encode(JSON.stringify(entries)).length >
              externalSourceLimits.responseBytes)
          )
            throw new ExternalSourceError("limit");
          n.onProgress?.(entries.length, t.length);
        }
      },
      n,
    );
  } catch (error) {
    s = isExternalSourceError(error).code;
  }
  if (s) {
    for (const sheetId of t)
      entries.some((entry) => entry.sheetId === sheetId) ||
        entries.push({
          sheetId: sheetId,
          sourceId: workbook.sheets.find((sheet) => sheet.id === sheetId)
            .externalConnection.source.id,
          error: s,
        });
    return s === "limit"
      ? {
          entries: entries.map(({ sheetId, sourceId }) => ({
            sheetId,
            sourceId,
            error: s,
          })),
          error: s,
        }
      : { entries: entries, error: s };
  }
  if (entries.some((entry) => entry.error)) return { entries: entries };
  try {
    const plan = planSourceGrants(
      workbook,
      entries.map((entry) => ({
        sheetId: entry.sheetId,
        snapshot: entry.snapshot,
        grant: entry.grant,
      })),
      n,
    );
    return { entries: entries, plan: plan };
  } catch (error) {
    return { entries: entries, error: isExternalSourceError(error).code };
  }
}
async function G2e(adapter, plan, t = {}) {
  return runSourceWithTimeout(
    "authorize",
    async (signal) => {
      const o = [];
      for (const a of new Set(
        plan.plans.map((plan) => plan.snapshot.source.id),
      )) {
        if (signal.aborted) throw new ExternalSourceError("cancelled");
        o.push(
          await authorizeSourceAccess(adapter, a, { ...t, signal: signal }),
        );
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
