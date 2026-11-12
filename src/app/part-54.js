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
import { hr } from "./part-01.js";
import { Mt } from "./part-02.js";
import {
  $a,
  AC,
  Ku,
  Mp,
  OR,
  Rn,
  bU,
  ct,
  fl,
  h0,
  rh,
  rn,
  ut,
  yU,
} from "./part-03.js";
import { Dr, M } from "./part-04.js";
import { BC, Oo, fT } from "./part-11.js";
import { q1 } from "./part-12.js";
import { C_, M2e, sA } from "./part-53.js";

function A2(e, r, t, n, o = 5) {
  const a = C_(e.total, t, n),
    s = sA(r, 0, a.physicalMax),
    d = M2e(s, e.total, t, n),
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
function oN(e) {
  return `ledgerglass:workbook:${e}`;
}
function j2e() {
  try {
    return typeof navigator < "u" ? (navigator.locks ?? null) : null;
  } catch {
    return null;
  }
}
function AS(e) {
  try {
    return structuredClone(e);
  } catch {
    throw new Oo(
      "invalid",
      "The workbook save inputs could not be captured. Keep your unsaved work open and correct the invalid input before retrying.",
    );
  }
}
function yH(e, r, t = {}) {
  const n = BC(e, r, {
      maxBytes: t.maxBytes,
      maxCheckpoints: t.maxCheckpoints,
    }),
    o = t.locks === void 0 ? j2e() : t.locks,
    a = t.lockTimeoutMs ?? 5e3,
    s = oN(r);
  if (!Number.isInteger(a) || a < 1 || a > 6e4)
    throw new Oo(
      "invalid",
      "The browser save lock timeout must be between 1 and 60,000 milliseconds.",
    );
  function d(u, l) {
    return new Promise((c, h) => {
      const f = (j) => new Oo("unavailable", j),
        p = new AbortController(),
        m = performance.now() + a;
      let v = true,
        g = false,
        b = false,
        y,
        w,
        E;
      const A = () => {
          (E !== void 0 && clearTimeout(E),
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
                : j instanceof Oo
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
        (l == null || l.addEventListener("abort", F, { once: !0 }),
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
          ((g = !0), A());
          try {
            return u();
          } catch (V) {
            throw ((b = !0), (y = V), V);
          }
        });
        Promise.resolve(j).then((z) => {
          ((v = !1), A(), c(z));
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
      const f = AS({
          workbook: u,
          view: l,
          expected: c,
          options: { allowIdentityChange: h.allowIdentityChange },
        }),
        p = h.signal;
      return d(() => n.save(f.workbook, f.view, f.expected, f.options), p);
    },
    async checkpoint(u, l, c = {}) {
      const h = AS({ label: u, expected: l }),
        f = c.signal;
      return d(() => n.checkpoint(h.label, h.expected), f);
    },
    async restoreCheckpoint(u, l, c, h, f = {}) {
      const p = AS({ id: u, workbook: l, expected: c, actor: h }),
        m = f.signal;
      return d(
        () => n.restoreCheckpoint(p.id, p.workbook, p.expected, p.actor),
        m,
      );
    },
  };
}
const F2e = {
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
  iE = (e) =>
    F2e[e] ??
    e
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .replace(/[_-]/g, " ")
      .replace(/^./, (r) => r.toUpperCase()),
  wp = (e) => e !== null && typeof e == "object" && !Array.isArray(e),
  O1 = (e) =>
    Array.isArray(e)
      ? `[${e.map(O1).join(",")}]`
      : wp(e)
        ? `{${Object.keys(e)
            .sort()
            .map((r) => `${JSON.stringify(r)}:${O1(e[r])}`)
            .join(",")}}`
        : (JSON.stringify(e) ?? "undefined"),
  oE = (e, r = 220) => (e.length > r ? `${e.slice(0, r)}…` : e);
function _3(e, r = 0) {
  if (e === void 0) return "Not set";
  if (e === null || e === "") return "Blank";
  if (typeof e == "boolean") return e ? "Yes" : "No";
  if (typeof e == "string")
    return e.startsWith("data:image/")
      ? "Embedded image"
      : /^\d{4}-\d{2}-\d{2}T/.test(e) && Number.isFinite(Date.parse(e))
        ? new Date(e).toLocaleString(void 0, {
            year: "numeric",
            month: "short",
            day: "numeric",
            hour: "numeric",
            minute: "2-digit",
            second: "2-digit",
            fractionalSecondDigits: 3,
            timeZoneName: "short",
          })
        : oE(e);
  if (Array.isArray(e))
    return e.length
      ? r >= 2
        ? `${e.length} items`
        : oE(
            e
              .slice(0, 4)
              .map((t) => _3(t, r + 1))
              .join("; ") +
              (e.length > 4 ? `; ${e.length - 4} more items` : ""),
          )
      : "None";
  if (wp(e)) {
    const t = Object.entries(e);
    return t.length
      ? r >= 2
        ? `${t.length} settings`
        : oE(
            t
              .slice(0, 4)
              .map(([n, o]) => `${iE(n)}: ${_3(o, r + 1)}`)
              .join(" · ") +
              (t.length > 4 ? ` · ${t.length - 4} more settings` : ""),
          )
      : "No settings";
  }
  return String(e);
}
function B2e(e) {
  return e
    ? e.raw === null || e.raw === ""
      ? "Blank"
      : typeof e.raw == "boolean"
        ? e.raw
          ? "TRUE"
          : "FALSE"
        : typeof e.raw == "string"
          ? (e.literal ? "Text: " : e.raw.startsWith("=") ? "Formula: " : "") +
            oE(e.raw)
          : String(e.raw)
    : "Empty cell";
}
function L2e(e) {
  const r = [],
    t = (n, o, a, s) => {
      if (O1(o) !== O1(a)) {
        if (
          s < 3 &&
          (wp(o) || wp(a)) &&
          (o === void 0 || wp(o)) &&
          (a === void 0 || wp(a))
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
function P2e(e) {
  const { text: r, language: t } = Dr(),
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
            (w == null ? void 0 : w.workbook.id) === e.workbook.id
              ? fT(w.workbook, e.workbook)
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
      var E, A;
      const w = document.activeElement;
      return (
        (A = (E = p.current) == null ? void 0 : E.querySelector("button")) ==
          null || A.focus(),
        () => {
          w != null && w.isConnected && w.focus();
        }
      );
    }, []));
  const v = (w) => {
      var E, A;
      return (
        ((E = e.workbook.sheets.find((k) => k.id === w)) == null
          ? void 0
          : E.name) ??
        ((A = n == null ? void 0 : n.workbook.sheets.find((k) => k.id === w)) ==
        null
          ? void 0
          : A.name) ??
        "Removed worksheet"
      );
    },
    g = (w) =>
      w[0] === "sheets" && w[1] === "order"
        ? "Worksheet order"
        : w[0] === "sheets"
          ? [v(w[1]), ...w.slice(2).map(iE)].join(" · ")
          : w.map(iE).join(" · "),
    b = a ? [...a.metadataChanges, ...a.structuralChanges].flatMap(L2e) : [],
    y = (w, E) =>
      w[0] === "sheets" && w[1] === "order" && Array.isArray(E)
        ? E.map((A) => v(String(A))).join(" → ") || "No worksheets"
        : _3(E);
  return i.jsx("div", {
    className: "wm-backdrop",
    children: i.jsxs("div", {
      ref: p,
      className: "wm-dialog pc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": r("Workbook save conflict"),
      onKeyDown: (w) => {
        var E;
        if (
          (w.key === "Escape" &&
            (w.preventDefault(), w.stopPropagation(), e.onClose()),
          w.key === "Tab")
        ) {
          const A = [
              ...(((E = p.current) == null
                ? void 0
                : E.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
                  )) ?? []),
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
      lang: t("Workbook save conflict"),
      "data-ui-label": "Workbook save conflict",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children: i.jsx(M, { children: "Workbook save conflict" }),
            }),
            i.jsx("button", {
              "aria-label": r("Close save conflict"),
              onClick: e.onClose,
              lang: t("Close save conflict"),
              "data-ui-label": "Close save conflict",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wm-body",
          children: [
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Your edits are still in this editor. Export them to keep a portable copy, or compare them with the saved version before reloading.",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsx("strong", {
                  children: i.jsx(M, { children: "My local workbook:" }),
                }),
                " ",
                e.workbook.name,
                " ",
                i.jsx(M, { children: "· revision" }),
                " ",
                e.workbook.revision,
              ],
            }),
            n &&
              i.jsxs("p", {
                children: [
                  i.jsx("strong", {
                    children: i.jsx(M, { children: "Current saved workbook:" }),
                  }),
                  " ",
                  n.workbook.name,
                  " ",
                  i.jsx(M, { children: "· revision " }),
                  n.workbook.revision,
                ],
              }),
            i.jsxs("div", {
              className: "pc-actions",
              children: [
                i.jsx("button", {
                  onClick: e.onExport,
                  children: i.jsx(M, { children: "Export my changes" }),
                }),
                i.jsx("button", {
                  onClick: m,
                  children: i.jsx(M, { children: "Refresh comparison" }),
                }),
              ],
            }),
            n &&
              n.workbook.id !== e.workbook.id &&
              i.jsx("p", {
                children: i.jsx(M, {
                  children:
                    "The saved location now contains a different workbook. Export your local workbook before deciding whether to load it.",
                }),
              }),
            a &&
              i.jsxs("section", {
                "aria-label": r("Conflicting workbook changes"),
                lang: t("Conflicting workbook changes"),
                "data-ui-label": "Conflicting workbook changes",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(M, {
                      children: "Changes from saved to my local workbook",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      a.cellChanges.length,
                      " ",
                      i.jsx(M, { children: "cell changes ·" }),
                      " ",
                      a.metadataChanges.length,
                      " ",
                      i.jsx(M, { children: "workbook changes ·" }),
                      " ",
                      a.structuralChanges.length,
                      " ",
                      i.jsx(M, { children: "worksheet changes" }),
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
                                children: i.jsx(M, { children: "Cell" }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(M, { children: "Saved" }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(M, {
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
                              var k, F;
                              return (
                                A !== "raw" &&
                                O1((k = w.before) == null ? void 0 : k[A]) !==
                                  O1((F = w.after) == null ? void 0 : F[A])
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
                                            children: B2e(A),
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
                                                        children: [iE(F), ":"],
                                                      }),
                                                      " ",
                                                      _3(
                                                        A == null
                                                          ? void 0
                                                          : A[F],
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
                      children: i.jsx(M, {
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
                                children: i.jsx(M, { children: "Property" }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(M, { children: "Saved" }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(M, {
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
                      children: i.jsx(M, {
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
                          children: i.jsx(M, {
                            children: "Full comparison details",
                          }),
                        }),
                        h &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("p", {
                                children: i.jsx(M, {
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
                i.jsx(M, {
                  children:
                    "Discard my local changes and load the saved workbook.",
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(M, {
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
                  children: i.jsx(M, { children: "Reload saved workbook" }),
                }),
                i.jsx("button", {
                  onClick: e.onClose,
                  children: i.jsx(M, {
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
const kS = {
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
  $j = {
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
function U2e(e) {
  try {
    if (e && typeof e == "object") {
      const { code: r, name: t } = e;
      if (typeof r == "string") {
        if (wH.has(r)) return r;
        if (Object.hasOwn($j, r)) return $j[r];
      }
      if (t === "AbortError") return "cancelled";
      if (t === "TimeoutError") return "timeout";
      if (t === "QuotaExceededError") return "quota";
    }
  } catch {}
  return "unknown";
}
function xH(e = {}) {
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
    Object.keys(kS).map((m) => [
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
      var y;
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
        const w = (y = e.onEvent) == null ? void 0 : y.call(e, m);
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
      if (!Object.hasOwn(kS, m) || !kS[m].includes(v))
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
          const k = U2e(A);
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
const Ew = () =>
    new Oo(
      "invalid",
      "The storage adapter returned an invalid acknowledgement. Your local changes are retained; check the saved copy before retrying.",
    ),
  aN = (e, r) => e === r || JSON.stringify(e) === JSON.stringify(r);
function sN(e, r) {
  const t = e.sheets.find((s) => !s.hidden && s.id === r.activeSheetId);
  if (!t || r.selection.sheetId !== t.id) throw Ew();
  const n = ct(r.selection.range);
  if (n.lastRow >= t.rowCount || n.lastCol >= t.colCount) throw Ew();
  const o = `${ut(n.firstCol)}${n.firstRow + 1}`,
    a = `${ut(n.lastCol)}${n.lastRow + 1}`;
  return {
    activeSheetId: t.id,
    selection: { sheetId: t.id, range: o === a ? o : `${o}:${a}` },
  };
}
function aE(e, r) {
  try {
    const t = rh(e.workbook),
      n = sN(t, e.view),
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
      (r && (!aN(t, rh(r.workbook)) || !aN(n, sN(r.workbook, r.view))))
    )
      throw Ew();
    return {
      ...e,
      workbook: t,
      view: n,
      guard: { ...o },
      checkpoints: e.checkpoints.map((a) => ({ ...a })),
    };
  } catch {
    throw Ew();
  }
}
function $2e(e, r) {
  const t = (n, o) =>
    n && typeof n.then == "function"
      ? Promise.resolve(n).then((a) => aE(a, o))
      : aE(n, o);
  return {
    ...("coordination" in e ? { coordination: e.coordination } : {}),
    load: () => {
      const n = r.start("persistence", "load");
      try {
        const o = e.load(),
          a = o ? aE(o) : null;
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
        const d = { workbook: rh(n), view: sN(n, o) };
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
function z2e(e, r, t) {
  const n = e.load();
  if (!n) return null;
  const o = aE(n);
  return o.recovered || !aN(o.workbook, rh(r)) || t ? null : o;
}
const zj = {
  persistence: "Saving and recovery",
  calculation: "Calculation",
  file: "File processing",
  provider: "AI provider",
  source: "External sources",
  host: "Application callbacks",
};
function V2e({ monitor: e, onClose: r }) {
  const t = ne.useSyncExternalStore(e.subscribe, e.getSnapshot, e.getSnapshot),
    n = ne.useId(),
    o = ne.useRef(null),
    a = Dr();
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
              { format: "ledgerglass-reliability-v1", ...e.getSnapshot() },
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
          (d.key === "Escape" && (d.stopPropagation(), r()), d.key === "Tab")
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
              children: i.jsx(M, { children: "Workbook health" }),
            }),
            i.jsx("button", {
              ref: o,
              type: "button",
              "aria-label": a.text("Close workbook health"),
              onClick: r,
              children: i.jsx(M, { children: "Close" }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(M, {
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
                      { scope: "col", children: i.jsx(M, { children: d }) },
                      d,
                    ),
                  ),
                }),
              }),
              i.jsx("tbody", {
                children: Object.keys(zj).map((d) => {
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
                          children: i.jsx(M, { children: zj[d] }),
                        }),
                        i.jsx("td", { children: i.jsx(M, { children: l }) }),
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
            i.jsx(M, { children: "Retained events" }),
            ": ",
            t.events.length,
            " ",
            "· ",
            i.jsx(M, { children: "Discarded older events" }),
            ":",
            " ",
            t.droppedEvents,
            " ·",
            " ",
            i.jsx(M, { children: "Diagnostic delivery failures" }),
            ":",
            " ",
            t.sinkFailures,
          ],
        }),
        i.jsx("p", {
          children: i.jsx(M, {
            children:
              "Save failures remain visible in the editor. Export your local changes before reloading a saved copy. A pending save must finish before another retry.",
          }),
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              type: "button",
              onClick: s,
              children: i.jsx(M, { children: "Download diagnostics" }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: () => e.clearEvents(),
              children: i.jsx(M, { children: "Clear diagnostic history" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const Cw = new WeakMap();
function A_(e) {
  return hr({
    cells: Object.fromEntries(
      Object.entries(e.cells).map(([r, t]) => [
        r,
        { ...t, style: void 0, numFmt: void 0, format: void 0 },
      ]),
    ),
    tables: e.tables,
    merges: e.merges,
    rowOrigins: e.rowOriginIds,
  });
}
function k_(e, r, t) {
  var o;
  if (t.readOnly) throw new Rn("refused");
  if (!r) {
    $a(e);
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
    throw new Rn("refused");
  if (!n.externalConnection) throw new Rn("invalid");
  if (A_(n) !== n.externalConnection.outputHash) throw new Rn("stale");
}
function EH(e, r, t, n, o = {}) {
  return CH(e, r, t, n, o);
}
function CH(e, r, t, n, o = {}, a = []) {
  k_(e, n.sheetId, o);
  const s = h0(t),
    d = bU(r, s.source),
    u = n.sheetId
      ? e.sheets.find((h) => h.id === n.sheetId).externalConnection
      : void 0;
  if (u && !AC(u.source, d.source)) throw new Rn("schema");
  if (
    !n.sheetId &&
    (e.sheets.filter((h) => h.externalConnection).length >= fl.sources ||
      !n.outputName.trim() ||
      n.outputName !== n.outputName.trim() ||
      n.outputName.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(n.outputName) ||
      /^'|'$/.test(n.outputName) ||
      e.sheets.some((h) => h.name.toLowerCase() === n.outputName.toLowerCase()))
  )
    throw new Rn("invalid");
  if (
    (n.sheetId || q1(e, n.outputName),
    e.sheets
      .filter((h) => h.id !== n.sheetId && !a.includes(h.id))
      .reduce((h, f) => h + Object.keys(f.cells).length, 0) +
      (d.rows.length + 1) * d.source.columns.length >
      Mt.maxCells)
  )
    throw new Rn("limit");
  const c = {
    workbookHash: hr(e),
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
    controls: yU(d),
  };
  if (
    c.controls.some((h) => {
      var f, p;
      return !Number.isFinite(
        h.sum -
          (((p =
            (f = c.previous) == null
              ? void 0
              : f.controls.find((m) => m.field === h.field)) == null
            ? void 0
            : p.sum) ?? 0),
      );
    })
  )
    throw new Rn("limit");
  return (Cw.set(c, hr(c)), c);
}
function W2e(e, r, t, n = {}) {
  if (!Cw.has(r) || Cw.get(r) !== hr(r) || hr(e) !== r.workbookHash)
    throw new Rn("stale");
  k_(e, r.sheetId, n);
  const o = h0(t);
  if ((h0(r.grant), !AC(o.source, r.snapshot.source))) throw new Rn("schema");
  EH(e, r.snapshot, o, r, n);
  const a = AH(e, r),
    s = rn({
      ...e,
      sheets: r.sheetId
        ? e.sheets.map((d) => (d.id === a.id ? a : d))
        : [...e.sheets, a],
    });
  return (Cw.delete(r), s);
}
function AH(e, r) {
  const t = e.sheets.find((d) => d.id === r.sheetId),
    n = t ? { ...t } : Ku(r.outputName),
    o = {};
  ([r.snapshot.source.columns.map((d) => d.name), ...r.snapshot.rows].forEach(
    (d, u) =>
      d.forEach((l, c) => {
        const h = `${ut(c)}${u + 1}`,
          f = t == null ? void 0 : t.cells[h];
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
      protocol: OR,
      source: structuredClone(r.snapshot.source),
      lastSuccess: {
        at: s,
        sourceTimestamp: r.snapshot.sourceTimestamp,
        version: r.snapshot.version,
        rows: r.snapshot.rows.length,
        controls: structuredClone(r.controls),
        dataHash: hr(r.snapshot),
      },
      lastAttempt: { at: s, outcome: "succeeded" },
      outputHash: A_(n),
    }),
    n
  );
}
const sE = new WeakMap();
function kH(e, r, t = {}) {
  if (!r.length || r.length > fl.sources || new Set(r).size !== r.length)
    throw new Rn("invalid");
  for (const n of r) k_(e, n, t);
}
function SH(e, r, t = {}) {
  const n = r.map((d) => d.sheetId);
  if (
    (kH(e, n, t),
    new TextEncoder().encode(JSON.stringify(r)).length > fl.responseBytes)
  )
    throw new Rn("limit");
  const o = r.map((d) =>
    CH(
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
    Mt.maxCells
  )
    throw new Rn("limit");
  const s = { workbookHash: hr(e), plans: o };
  return (sE.set(s, hr(s)), s);
}
function H2e(e, r, t, n = {}) {
  if (!sE.has(r) || sE.get(r) !== hr(r) || r.workbookHash !== hr(e))
    throw new Rn("stale");
  const o = t.map((u) => h0(u)),
    a = new Set(r.plans.map((u) => u.snapshot.source.id));
  if (
    o.length !== a.size ||
    new Set(o.map((u) => u.source.id)).size !== o.length ||
    o.some((u) => !a.has(u.source.id))
  )
    throw new Rn("refused");
  for (const u of r.plans) {
    h0(u.grant);
    const l = o.find((c) => c.source.id === u.snapshot.source.id);
    if (!l || !AC(l.source, u.snapshot.source)) throw new Rn("schema");
  }
  SH(
    e,
    r.plans.map((u) => ({
      sheetId: u.sheetId,
      snapshot: u.snapshot,
      grant: o.find((l) => l.source.id === u.snapshot.source.id),
    })),
    n,
  );
  const s = new Map(r.plans.map((u) => [u.sheetId, AH(e, u)])),
    d = rn({ ...e, sheets: e.sheets.map((u) => s.get(u.id) ?? u) });
  sE.delete(r);
  for (const u of r.plans) Cw.delete(u);
  return d;
}
function lN(e, r, t, n, o = {}) {
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
  return rn({
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
async function $x(e, r, t = {}) {
  var u, l, c;
  const n = t.timeoutMs ?? fl.timeoutMs;
  if (!Number.isInteger(n) || n < 1 || n > fl.timeoutMs)
    throw new Rn("invalid");
  const o = (u = t.reliability) == null ? void 0 : u.start("source", e),
    a = new AbortController();
  let s,
    d = () => {};
  try {
    if ((l = t.signal) != null && l.aborted) throw new Rn("cancelled");
    const h = new Promise((p, m) => {
        var v;
        ((d = () => {
          (a.abort(), m(new Rn("cancelled")));
        }),
          (v = t.signal) == null ||
            v.addEventListener("abort", d, { once: !0 }),
          (s = setTimeout(() => {
            (a.abort(), m(new Rn("timeout")));
          }, n)));
      }),
      f = await Promise.race([Promise.resolve().then(() => r(a.signal)), h]);
    if (a.signal.aborted) throw new Rn("cancelled");
    return (o == null || o.finish(), f);
  } catch (h) {
    const f = Mp(h);
    throw (o == null || o.fail(f), f);
  } finally {
    (clearTimeout(s),
      (c = t.signal) == null || c.removeEventListener("abort", d));
  }
}
async function q2e(e, r = {}) {
  return $x(
    "catalog",
    async (t) => {
      const n = await e.catalog({ signal: t });
      if (!Array.isArray(n) || n.length > fl.sources) throw new Rn("invalid");
      const o = n.map((a) => h0(a));
      if (new Set(o.map((a) => a.source.id)).size !== o.length)
        throw new Rn("invalid");
      return o;
    },
    r,
  );
}
async function S_(e, r, t = {}) {
  return $x(
    "authorize",
    async (n) => {
      const o = h0(await e.authorize(r, { signal: n }));
      if (o.source.id !== r) throw new Rn("refused");
      return o;
    },
    t,
  );
}
async function TH(e, r, t = {}) {
  const n = await S_(e, r, t),
    o = await $x(
      "refresh",
      async (a) => {
        const s = bU(await e.refresh(r, { signal: a }), n.source);
        return (h0(n), s);
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
async function X2e(e, r, t, n = {}) {
  kH(e, t, n);
  const o = [],
    a = new Map();
  let s;
  try {
    await $x(
      "refresh",
      async (d) => {
        var u;
        for (const l of t) {
          if (d.aborted) throw new Rn("cancelled");
          const c = e.sheets.find((f) => f.id === l).externalConnection.source
            .id;
          let h = a.get(c);
          if (!h) {
            try {
              h = await TH(r, c, { ...n, signal: d });
            } catch (f) {
              h = Mp(f).code;
            }
            if (d.aborted) throw new Rn("cancelled");
            a.set(c, h);
          }
          if (
            (o.push(
              typeof h == "string"
                ? { sheetId: l, sourceId: c, error: h }
                : { sheetId: l, sourceId: c, ...h },
            ),
            new TextEncoder().encode(JSON.stringify(o)).length >
              fl.responseBytes)
          )
            throw new Rn("limit");
          (u = n.onProgress) == null || u.call(n, o.length, t.length);
        }
      },
      n,
    );
  } catch (d) {
    s = Mp(d).code;
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
          entries: o.map(({ sheetId: d, sourceId: u }) => ({
            sheetId: d,
            sourceId: u,
            error: s,
          })),
          error: s,
        }
      : { entries: o, error: s };
  }
  if (o.some((d) => d.error)) return { entries: o };
  try {
    const d = SH(
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
    return { entries: o, error: Mp(d).code };
  }
}
async function G2e(e, r, t = {}) {
  return $x(
    "authorize",
    async (n) => {
      const o = [];
      for (const a of new Set(r.plans.map((s) => s.snapshot.source.id))) {
        if (n.aborted) throw new Rn("cancelled");
        o.push(await S_(e, a, { ...t, signal: n }));
      }
      return o.map((a) => h0(a));
    },
    t,
  );
}

export {
  $2e,
  A2,
  A_,
  EH,
  G2e,
  H2e,
  P2e,
  S_,
  TH,
  V2e,
  W2e,
  X2e,
  lN,
  oN,
  q2e,
  xH,
  yH,
  z2e,
};
