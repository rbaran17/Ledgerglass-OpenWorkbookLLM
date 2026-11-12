import {
  X,
  ListTree,
  Check,
  MessageSquare,
  Link2,
  RotateCcw,
  ExternalLink,
} from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  AQ,
  CQ,
  EQ,
  Jl,
  NB,
  SQ,
  X3,
  _t,
  ct,
  iQ,
  kQ,
  nQ,
  oD,
  oQ,
  ut,
} from "./03-workbook-model.js";
import { Dr, M } from "./04-cell-input-paste-and-fill.js";
import { Ej, Wm, wS } from "./48-format-cells-and-move.js";

function j4e(e) {
  var z;
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e);
  a.current = e;
  const s = ne.useRef(null);
  s.current ||
    (s.current = {
      workbookId: e.workbook.id,
      revision: e.workbook.revision,
      snapshot: JSON.stringify(e.workbook),
      sheetId: e.sheetId,
      range: e.range,
    });
  const [d, u] = ne.useState(/column/i.test(e.mode) ? "column" : "row"),
    [l, c] = ne.useState(0),
    [h, f] = ne.useState(""),
    p = e.workbook.sheets.find((V) => V.id === e.sheetId),
    m = ne.useMemo(() => (p ? NB(p, d) : []), [p, d]),
    v = ct(e.range),
    g = d === "row" ? v.firstRow : v.firstCol,
    b = d === "row" ? v.lastRow : v.lastCol,
    y =
      e.readOnly ||
      !!((z = p == null ? void 0 : p.protection) != null && z.enabled),
    w = s.current,
    E = ne.useMemo(
      () => JSON.stringify(e.workbook) !== w.snapshot,
      [e.workbook, w.snapshot],
    ),
    A =
      w.workbookId !== e.workbook.id ||
      w.revision !== e.workbook.revision ||
      w.sheetId !== e.sheetId ||
      w.range !== e.range ||
      E ||
      !p,
    k = y || A,
    F = Math.max(0, ...m.map((V) => V.depth)) + 1,
    I = Math.max(1, Math.ceil(m.length / wS)),
    O = Math.min(l, I - 1);
  ne.useEffect(() => {
    var D, B;
    const V = document.activeElement;
    return (
      (B =
        (D = o.current) == null
          ? void 0
          : D.querySelector("button:not(:disabled)")) == null || B.focus(),
      () => {
        V != null && V.isConnected && V.focus();
      }
    );
  }, []);
  function j(V, D) {
    var B, _;
    try {
      const C = a.current;
      if (
        C.readOnly ||
        ((_ =
          (B = C.workbook.sheets.find((N) => N.id === C.sheetId)) == null
            ? void 0
            : B.protection) != null &&
          _.enabled)
      )
        throw Error(
          "This worksheet is read-only. Unlock it before changing outlines.",
        );
      if (
        C.workbook.id !== w.workbookId ||
        C.workbook.revision !== w.revision ||
        C.sheetId !== w.sheetId ||
        C.range !== w.range ||
        JSON.stringify(C.workbook) !== w.snapshot
      )
        throw Error(Ej);
      const T = V(C.workbook);
      (T !== C.workbook && C.onWorkbookChange(T, D), C.onClose());
    } catch (C) {
      f(C instanceof Error ? C.message : "The outline could not be changed.");
    }
  }
  return i.jsx("div", {
    className: "ot-backdrop",
    children: i.jsxs("section", {
      className: "ot-dialog",
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      tabIndex: -1,
      onKeyDown: (V) => {
        var D;
        if (
          (V.key === "Escape" &&
            (V.preventDefault(), V.stopPropagation(), a.current.onClose()),
          V.key === "Tab")
        ) {
          const B = Array.from(
              ((D = o.current) == null
                ? void 0
                : D.querySelectorAll(
                    "button:not(:disabled),select:not(:disabled),input:not(:disabled),[tabindex='0']",
                  )) ?? [],
            ),
            _ = B[0],
            C = B.at(-1);
          V.shiftKey &&
          (document.activeElement === _ || document.activeElement === o.current)
            ? (V.preventDefault(), C == null || C.focus())
            : !V.shiftKey &&
              document.activeElement === C &&
              (V.preventDefault(), _ == null || _.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "ot-header",
          children: [
            i.jsx(ListTree, { size: 20 }),
            i.jsx("h2", {
              id: n,
              children: i.jsx(M, { children: "Row and column outlines" }),
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": r("Close outline tools"),
              onClick: e.onClose,
              lang: t("Close outline tools"),
              "data-ui-label": "Close outline tools",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "ot-body",
          children: [
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Group detail rows or columns, leaving summary cells outside the group. Collapsing changes visibility only; formulas and totals remain unchanged.",
              }),
            }),
            y &&
              i.jsx("p", {
                className: "ot-notice",
                role: "status",
                children: i.jsx(M, {
                  children:
                    "This worksheet is read-only. Outline details can be inspected.",
                }),
              }),
            (h || A) &&
              i.jsx("p", {
                role: "alert",
                className: "ot-error",
                children: h || Ej,
              }),
            i.jsxs("div", {
              className: "ot-create",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Group by" }),
                    i.jsxs("select", {
                      "aria-label": r("Outline axis"),
                      value: d,
                      onChange: (V) => {
                        (u(V.target.value), c(0), f(""));
                      },
                      lang: t("Outline axis"),
                      "data-ui-label": "Outline axis",
                      children: [
                        i.jsx("option", { value: "row", children: "Rows" }),
                        i.jsx("option", {
                          value: "column",
                          children: "Columns",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  children: [
                    i.jsx("span", {
                      className: "ot-small",
                      children: i.jsx(M, { children: "Selected range" }),
                    }),
                    i.jsx("strong", { children: Wm(d, g, b) }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  className: "ot-primary",
                  disabled: k,
                  onClick: () =>
                    j(
                      (V) => EQ(V, e.sheetId, d, g, b),
                      "Group " + Wm(d, g, b).toLowerCase(),
                    ),
                  children: i.jsx(M, { children: "Group selection" }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "ot-levels",
              children: [
                i.jsx("strong", {
                  children: i.jsx(M, { children: "Show level" }),
                }),
                Array.from({ length: F }, (V, D) =>
                  i.jsx(
                    "button",
                    {
                      type: "button",
                      disabled: k || !m.length,
                      "aria-label": "Show outline level " + (D + 1),
                      onClick: () =>
                        j(
                          (B) => kQ(B, e.sheetId, d, D + 1),
                          "Show " + d + " outline level " + (D + 1),
                        ),
                      children: D + 1,
                    },
                    D,
                  ),
                ),
              ],
            }),
            i.jsxs("p", {
              className: "ot-small",
              children: [
                i.jsx(M, {
                  children:
                    "Level 1 hides all grouped details. Higher levels reveal nested detail; level ",
                }),
                F,
                " ",
                i.jsx(M, {
                  children:
                    "expands every group on this axis. Manually hidden rows and columns remain hidden.",
                }),
              ],
            }),
            i.jsxs("div", {
              className: "ot-groups-heading",
              children: [
                i.jsxs("h3", {
                  children: [
                    m.length,
                    " ",
                    d === "row" ? "row" : "column",
                    " ",
                    m.length === 1 ? "group" : "groups",
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: k || !m.length,
                  onClick: () =>
                    j(
                      (V) => SQ(V, e.sheetId, d),
                      "Clear " + d + " outline groups",
                    ),
                  children: i.jsx(M, { children: "Clear groups" }),
                }),
              ],
            }),
            m.length
              ? i.jsx("ul", {
                  className: "ot-list",
                  children: m.slice(O * wS, (O + 1) * wS).map((V) =>
                    i.jsxs(
                      "li",
                      {
                        style: { paddingLeft: 12 + (V.depth - 1) * 12 },
                        children: [
                          i.jsxs("div", {
                            children: [
                              i.jsx("strong", {
                                children: Wm(d, V.start, V.end),
                              }),
                              i.jsxs("span", {
                                className: "ot-small",
                                children: [
                                  i.jsx(M, { children: "Depth " }),
                                  V.depth,
                                  " ·",
                                  " ",
                                  V.collapsed ? "Collapsed" : "Expanded",
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("div", {
                            className: "ot-actions",
                            children: [
                              i.jsx("button", {
                                type: "button",
                                disabled: k,
                                "aria-label":
                                  (V.collapsed ? "Expand " : "Collapse ") +
                                  Wm(d, V.start, V.end).toLowerCase(),
                                onClick: () =>
                                  j(
                                    (D) => AQ(D, e.sheetId, V.id, !V.collapsed),
                                    (V.collapsed ? "Expand " : "Collapse ") +
                                      Wm(d, V.start, V.end).toLowerCase(),
                                  ),
                                children: V.collapsed ? "Expand" : "Collapse",
                              }),
                              i.jsx("button", {
                                type: "button",
                                disabled: k,
                                "aria-label":
                                  "Ungroup " +
                                  Wm(d, V.start, V.end).toLowerCase(),
                                onClick: () =>
                                  j(
                                    (D) => CQ(D, e.sheetId, V.id),
                                    "Ungroup " +
                                      Wm(d, V.start, V.end).toLowerCase(),
                                  ),
                                children: i.jsx(M, { children: "Ungroup" }),
                              }),
                            ],
                          }),
                        ],
                      },
                      V.id,
                    ),
                  ),
                })
              : i.jsx("p", {
                  className: "ot-empty",
                  children: i.jsx(M, {
                    children:
                      "Select detail cells in the worksheet, then choose Group selection.",
                  }),
                }),
            I > 1 &&
              i.jsxs("div", {
                className: "ot-pager",
                children: [
                  i.jsx("button", {
                    type: "button",
                    disabled: O === 0,
                    onClick: () => c(O - 1),
                    children: i.jsx(M, { children: "Previous groups" }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(M, { children: "Page " }),
                      O + 1,
                      " ",
                      i.jsx(M, { children: "of " }),
                      I,
                    ],
                  }),
                  i.jsx("button", {
                    type: "button",
                    disabled: O + 1 === I,
                    onClick: () => c(O + 1),
                    children: i.jsx(M, { children: "Next groups" }),
                  }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "ot-footer",
          children: [
            i.jsx("span", {
              children: i.jsx(M, {
                children: "Up to 8 nested groups · 1,000 groups per worksheet",
              }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: i.jsx(M, { children: "Done" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const by = 20;
function F4e(e) {
  const r = ct(e.range);
  if (e.active) {
    const t = _t(e.active);
    if (
      t.row >= r.firstRow &&
      t.row <= r.lastRow &&
      t.col >= r.firstCol &&
      t.col <= r.lastCol
    )
      return `${ut(t.col)}${t.row + 1}`;
  }
  return `${ut(r.firstCol)}${r.firstRow + 1}`;
}
const Cj = (e) => new Date(e).toLocaleString();
function B4e(e) {
  var ae, oe, de, he, Se, le, ce, ge;
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const s = ne.useRef({
      id: e.workbook.id,
      revision: e.workbook.revision,
      sheetId: e.sheetId,
      range: e.range,
      active: e.active,
      content: JSON.stringify(e.workbook),
    }),
    d = e.workbook.sheets.find((se) => se.id === e.sheetId),
    u = F4e(e),
    l = (ae = d == null ? void 0 : d.cells[u]) == null ? void 0 : ae.hyperlink,
    [c, h] = ne.useState(e.mode ?? "comments"),
    [f, p] = ne.useState("cell"),
    [m, v] = ne.useState(0),
    [g, b] = ne.useState(
      ((de =
        (oe = e.workbook.reviewThreads) == null
          ? void 0
          : oe.find(
              (se) =>
                !se.orphaned &&
                se.anchor.sheetId === e.sheetId &&
                se.anchor.address === u,
            )) == null
        ? void 0
        : de.id) ?? "",
    ),
    [y, w] = ne.useState(0),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState((l == null ? void 0 : l.kind) ?? "external"),
    [V, D] = ne.useState(
      (l == null ? void 0 : l.kind) === "external" ? l.url : "",
    ),
    [B, _] = ne.useState(
      (l == null ? void 0 : l.kind) === "internal" ? l.sheetId : e.sheetId,
    ),
    [C, T] = ne.useState(
      (l == null ? void 0 : l.kind) === "internal" ? l.address : "A1",
    ),
    [N, R] = ne.useState((l == null ? void 0 : l.label) ?? ""),
    L =
      s.current.id !== e.workbook.id ||
      s.current.revision !== e.workbook.revision ||
      s.current.sheetId !== e.sheetId ||
      s.current.range !== e.range ||
      s.current.active !== e.active ||
      s.current.content !== JSON.stringify(e.workbook),
    $ =
      e.readOnly ||
      !!((he = d == null ? void 0 : d.protection) != null && he.enabled) ||
      L ||
      !d,
    G = [...(e.workbook.reviewThreads ?? [])]
      .sort((se, we) => we.updatedAt.localeCompare(se.updatedAt))
      .filter(
        (se) =>
          f === "all" ||
          (f === "orphaned"
            ? !!se.orphaned
            : f === "unresolved"
              ? !se.resolved && !se.orphaned
              : !se.orphaned &&
                se.anchor.sheetId === e.sheetId &&
                se.anchor.address === u),
      ),
    H = G.find((se) => se.id === g),
    Q = e.workbook.sheets.find(
      (se) => se.id === (H == null ? void 0 : H.anchor.sheetId),
    ),
    X$1 =
      e.readOnly ||
      L ||
      !!(H != null && H.orphaned) ||
      !Q ||
      !!((Se = Q.protection) != null && Se.enabled),
    J = ((le = e.actor) == null ? void 0 : le.trim()) || "Reviewer";
  ne.useEffect(() => {
    var we, Ee;
    const se = document.activeElement;
    return (
      (Ee =
        (we = o.current) == null
          ? void 0
          : we.querySelector("button:not(:disabled)")) == null || Ee.focus(),
      () => {
        se != null && se.isConnected && se.focus();
      }
    );
  }, []);
  const q = (se, we) => {
      try {
        if ((O(""), e.readOnly)) throw new Error("This workbook is read-only.");
        if (L)
          throw new Error(
            "The workbook or selection changed. Close and reopen the review tools.",
          );
        const Ee = se();
        (Ee !== e.workbook &&
          JSON.stringify(Ee) !== JSON.stringify(e.workbook) &&
          e.onWorkbookChange(Ee, we),
          e.onClose());
      } catch (Ee) {
        O(
          Ee instanceof Error
            ? Ee.message
            : "The review change could not be saved.",
        );
      }
    },
    K = (() => {
      try {
        return X3(
          j === "external"
            ? { kind: "external", url: V, label: N }
            : { kind: "internal", sheetId: B, address: C, label: N },
        );
      } catch {
        return null;
      }
    })(),
    te = (se, we) => {
      var Ee;
      try {
        (O(""), (Ee = e.onNavigate) == null || Ee.call(e, se, we), e.onClose());
      } catch (Ae) {
        O(
          Ae instanceof Error
            ? Ae.message
            : "The link target could not be opened.",
        );
      }
    };
  return i.jsx("div", {
    className: "rv-backdrop",
    children: i.jsxs("div", {
      className: "rv-dialog",
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (se) => {
        var we;
        if (
          (se.stopPropagation(),
          se.key === "Escape" && (se.preventDefault(), a.current()),
          se.key === "Tab")
        ) {
          const Ee = [
              ...(((we = o.current) == null
                ? void 0
                : we.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),a[href],[tabindex="0"]',
                  )) ?? []),
            ],
            Ae = Ee[0],
            ye = Ee.at(-1);
          se.shiftKey && document.activeElement === Ae
            ? (se.preventDefault(), ye == null || ye.focus())
            : !se.shiftKey &&
              document.activeElement === ye &&
              (se.preventDefault(), Ae == null || Ae.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "rv-header",
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(M, { children: "Comments & links" }),
                }),
                i.jsxs("p", {
                  children: [
                    (d == null ? void 0 : d.name) ?? "Worksheet unavailable",
                    "!",
                    u,
                  ],
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close review tools"),
              onClick: e.onClose,
              lang: t("Close review tools"),
              "data-ui-label": "Close review tools",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("div", {
          className: "rv-tabs",
          role: "tablist",
          "aria-label": r("Review tools"),
          lang: t("Review tools"),
          "data-ui-label": "Review tools",
          children: ["comments", "hyperlink"].map((se, we) =>
            i.jsxs(
              "button",
              {
                role: "tab",
                id: `${n}-${se}-tab`,
                "aria-controls": `${n}-${se}-panel`,
                "aria-selected": c === se,
                tabIndex: c === se ? 0 : -1,
                onClick: () => {
                  (h(se), O(""));
                },
                onKeyDown: (Ee) => {
                  var Ae;
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(Ee.key)
                  ) {
                    Ee.preventDefault();
                    const ye =
                      Ee.key === "Home"
                        ? "comments"
                        : Ee.key === "End"
                          ? "hyperlink"
                          : we
                            ? "comments"
                            : "hyperlink";
                    (h(ye),
                      (Ae = document.getElementById(`${n}-${ye}-tab`)) ==
                        null || Ae.focus());
                  }
                },
                children: [
                  se === "comments"
                    ? i.jsx(MessageSquare, { size: 15 })
                    : i.jsx(Link2, { size: 15 }),
                  se === "comments" ? "Comments" : "Hyperlink",
                ],
              },
              se,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "rv-body",
          children: [
            L &&
              i.jsx("p", {
                className: "rv-error",
                role: "alert",
                children: i.jsx(M, {
                  children:
                    "The workbook or selection changed. Close and reopen the review tools.",
                }),
              }),
            e.readOnly &&
              i.jsx("p", {
                className: "rv-muted",
                children: i.jsx(M, {
                  children:
                    "This workbook is read-only. Discussions and links remain available for inspection.",
                }),
              }),
            !e.readOnly &&
              ((ce = d == null ? void 0 : d.protection) == null
                ? void 0
                : ce.enabled) &&
              i.jsx("p", {
                className: "rv-muted",
                children: i.jsx(M, {
                  children:
                    "This worksheet is protected. Unprotect it before adding comments or changing links.",
                }),
              }),
            c === "comments"
              ? i.jsxs("section", {
                  role: "tabpanel",
                  id: `${n}-comments-panel`,
                  "aria-labelledby": `${n}-comments-tab`,
                  children: [
                    i.jsxs("p", {
                      className: "rv-muted",
                      children: [
                        i.jsx(M, { children: "Posting as " }),
                        i.jsx("strong", { children: J }),
                        i.jsx(M, {
                          children:
                            ". Names and times come from this application; they are not verified identities. Comments stay separate from cell notes.",
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      className: "rv-filter",
                      children: [
                        i.jsx(M, { children: "Show discussions" }),
                        i.jsxs("select", {
                          "aria-label": r("Show discussions"),
                          value: f,
                          onChange: (se) => {
                            (p(se.target.value), v(0), b(""), F(""));
                          },
                          lang: t("Show discussions"),
                          "data-ui-label": "Show discussions",
                          children: [
                            i.jsx("option", {
                              value: "cell",
                              children: "This cell",
                            }),
                            i.jsx("option", {
                              value: "unresolved",
                              children: "All unresolved",
                            }),
                            i.jsx("option", {
                              value: "all",
                              children: "All discussions",
                            }),
                            i.jsx("option", {
                              value: "orphaned",
                              children: "Deleted targets",
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "rv-review-grid",
                      children: [
                        i.jsxs("div", {
                          className: "rv-thread-list",
                          "aria-label": r("Discussions"),
                          lang: t("Discussions"),
                          "data-ui-label": "Discussions",
                          children: [
                            !G.length &&
                              i.jsx("p", {
                                className: "rv-empty",
                                children: i.jsx(M, {
                                  children: "No discussions in this view.",
                                }),
                              }),
                            G.slice(m * by, (m + 1) * by).map((se) =>
                              i.jsxs(
                                "button",
                                {
                                  "aria-pressed": se.id === g,
                                  className: "rv-thread-button",
                                  onClick: () => {
                                    (b(se.id), F(""), w(0), O(""));
                                  },
                                  children: [
                                    i.jsxs("span", {
                                      children: [
                                        i.jsxs("strong", {
                                          children: [
                                            se.anchor.sheetName,
                                            "!",
                                            se.anchor.address,
                                          ],
                                        }),
                                        i.jsx("small", {
                                          children: se.orphaned
                                            ? "Deleted target"
                                            : se.resolved
                                              ? "Resolved"
                                              : "Open",
                                        }),
                                      ],
                                    }),
                                    i.jsx("span", {
                                      className: "rv-thread-summary",
                                      children: se.messages[0].text,
                                    }),
                                    i.jsxs("small", {
                                      children: [
                                        se.messages.length,
                                        " ",
                                        se.messages.length === 1
                                          ? "message"
                                          : "messages",
                                        " ",
                                        "· ",
                                        se.messages.at(-1).author,
                                      ],
                                    }),
                                  ],
                                },
                                se.id,
                              ),
                            ),
                            G.length > by &&
                              i.jsxs("div", {
                                className: "rv-pages",
                                children: [
                                  i.jsx("button", {
                                    disabled: !m,
                                    onClick: () => v(m - 1),
                                    children: i.jsx(M, {
                                      children: "Previous discussions",
                                    }),
                                  }),
                                  i.jsxs("span", {
                                    children: [
                                      m + 1,
                                      " / ",
                                      Math.ceil(G.length / by),
                                    ],
                                  }),
                                  i.jsx("button", {
                                    disabled: (m + 1) * by >= G.length,
                                    onClick: () => v(m + 1),
                                    children: i.jsx(M, {
                                      children: "Next discussions",
                                    }),
                                  }),
                                ],
                              }),
                          ],
                        }),
                        i.jsx("div", {
                          className: "rv-thread-detail",
                          children: H
                            ? i.jsxs("section", {
                                "aria-label": r("Selected discussion"),
                                lang: t("Selected discussion"),
                                "data-ui-label": "Selected discussion",
                                children: [
                                  i.jsxs("div", {
                                    className: "rv-thread-title",
                                    children: [
                                      i.jsxs("h3", {
                                        children: [
                                          H.anchor.sheetName,
                                          "!",
                                          H.anchor.address,
                                        ],
                                      }),
                                      !H.orphaned &&
                                        e.onNavigate &&
                                        i.jsx("button", {
                                          onClick: () =>
                                            te(
                                              H.anchor.sheetId,
                                              H.anchor.address,
                                            ),
                                          children: i.jsx(M, {
                                            children: "Go to cell",
                                          }),
                                        }),
                                    ],
                                  }),
                                  H.orphaned &&
                                    i.jsxs("p", {
                                      className: "rv-orphan",
                                      children: [
                                        H.orphaned.reason ===
                                        "worksheet-deleted"
                                          ? "The original worksheet was deleted."
                                          : "The original row or column was deleted.",
                                        " ",
                                        i.jsx(M, {
                                          children:
                                            "This discussion is retained for review and cannot receive new replies.",
                                        }),
                                      ],
                                    }),
                                  H.resolved &&
                                    i.jsxs("p", {
                                      className: "rv-resolution",
                                      children: [
                                        i.jsx(Check, { size: 14 }),
                                        " ",
                                        i.jsx(M, { children: "Resolved by" }),
                                        " ",
                                        H.resolved.author,
                                        " ·",
                                        " ",
                                        i.jsx("time", {
                                          dateTime: H.resolved.createdAt,
                                          title: H.resolved.createdAt,
                                          children: Cj(H.resolved.createdAt),
                                        }),
                                      ],
                                    }),
                                  H.messages
                                    .slice(y * 10, (y + 1) * 10)
                                    .map((se) =>
                                      i.jsxs(
                                        "article",
                                        {
                                          className: "rv-message",
                                          children: [
                                            i.jsxs("div", {
                                              children: [
                                                i.jsx("strong", {
                                                  children: se.author,
                                                }),
                                                i.jsx("time", {
                                                  dateTime: se.createdAt,
                                                  title: se.createdAt,
                                                  children: Cj(se.createdAt),
                                                }),
                                              ],
                                            }),
                                            i.jsx("p", { children: se.text }),
                                          ],
                                        },
                                        se.id,
                                      ),
                                    ),
                                  H.messages.length > 10 &&
                                    i.jsxs("div", {
                                      className: "rv-pages",
                                      children: [
                                        i.jsx("button", {
                                          disabled: !y,
                                          onClick: () => w(y - 1),
                                          children: i.jsx(M, {
                                            children: "Previous messages",
                                          }),
                                        }),
                                        i.jsxs("span", {
                                          children: [
                                            y + 1,
                                            " /",
                                            " ",
                                            Math.ceil(H.messages.length / 10),
                                          ],
                                        }),
                                        i.jsx("button", {
                                          disabled:
                                            (y + 1) * 10 >= H.messages.length,
                                          onClick: () => w(y + 1),
                                          children: i.jsx(M, {
                                            children: "Next messages",
                                          }),
                                        }),
                                      ],
                                    }),
                                  !H.orphaned &&
                                    i.jsx("div", {
                                      className: "rv-thread-actions",
                                      children: i.jsxs("button", {
                                        disabled: X$1,
                                        onClick: () =>
                                          q(
                                            () =>
                                              oQ(
                                                e.workbook,
                                                H.id,
                                                !H.resolved,
                                                J,
                                                { readOnly: e.readOnly },
                                              ),
                                            H.resolved
                                              ? "Reopen review discussion"
                                              : "Resolve review discussion",
                                          ),
                                        children: [
                                          H.resolved
                                            ? i.jsx(RotateCcw, { size: 14 })
                                            : i.jsx(Check, { size: 14 }),
                                          H.resolved
                                            ? "Reopen discussion"
                                            : "Resolve discussion",
                                        ],
                                      }),
                                    }),
                                  !H.resolved &&
                                    !H.orphaned &&
                                    i.jsxs("form", {
                                      className: "rv-compose",
                                      onSubmit: (se) => {
                                        (se.preventDefault(),
                                          q(
                                            () =>
                                              iQ(e.workbook, H.id, k, J, {
                                                readOnly: e.readOnly,
                                              }),
                                            "Reply to review discussion",
                                          ));
                                      },
                                      children: [
                                        i.jsxs("label", {
                                          children: [
                                            i.jsx(M, { children: "Reply" }),
                                            i.jsx("textarea", {
                                              "aria-label": r("Reply"),
                                              rows: 3,
                                              maxLength: Jl.textLength,
                                              value: k,
                                              disabled: X$1,
                                              onChange: (se) =>
                                                F(se.target.value),
                                              lang: t("Reply"),
                                              "data-ui-label": "Reply",
                                            }),
                                          ],
                                        }),
                                        i.jsx("button", {
                                          className: "rv-primary",
                                          disabled: X$1 || !k.trim(),
                                          children: i.jsx(M, {
                                            children: "Post reply",
                                          }),
                                        }),
                                      ],
                                    }),
                                ],
                              })
                            : i.jsx("p", {
                                className: "rv-empty",
                                children: i.jsx(M, {
                                  children:
                                    "Select a discussion to read its messages.",
                                }),
                              }),
                        }),
                      ],
                    }),
                    i.jsxs("form", {
                      className: "rv-compose rv-new",
                      onSubmit: (se) => {
                        (se.preventDefault(),
                          q(
                            () =>
                              nQ(e.workbook, e.sheetId, u, E, J, {
                                readOnly: e.readOnly,
                              }),
                            "Add review discussion",
                          ));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "New comment on " }),
                            d == null ? void 0 : d.name,
                            "!",
                            u,
                            i.jsx("textarea", {
                              "aria-label": r("New comment"),
                              rows: 3,
                              maxLength: Jl.textLength,
                              value: E,
                              disabled: $,
                              onChange: (se) => A(se.target.value),
                              lang: t("New comment"),
                              "data-ui-label": "New comment",
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "rv-compose-footer",
                          children: [
                            i.jsxs("small", {
                              children: [E.length, " / ", Jl.textLength],
                            }),
                            i.jsx("button", {
                              className: "rv-primary",
                              disabled: $ || !E.trim(),
                              children: i.jsx(M, { children: "Post comment" }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "rv-muted",
                      children: i.jsx(M, {
                        children:
                          "Comments are stored in this workbook. Typing an @name does not notify or send a message to anyone.",
                      }),
                    }),
                  ],
                })
              : i.jsxs("section", {
                  role: "tabpanel",
                  id: `${n}-hyperlink-panel`,
                  "aria-labelledby": `${n}-hyperlink-tab`,
                  children: [
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children:
                          "Add a link to this cell without replacing its value, formula, formatting or note.",
                      }),
                    }),
                    i.jsxs("form", {
                      className: "rv-link-form",
                      onSubmit: (se) => {
                        (se.preventDefault(),
                          q(
                            () =>
                              oD(
                                e.workbook,
                                e.sheetId,
                                u,
                                j === "external"
                                  ? { kind: "external", url: V, label: N }
                                  : {
                                      kind: "internal",
                                      sheetId: B,
                                      address: C,
                                      label: N,
                                    },
                                { readOnly: e.readOnly },
                              ),
                            "Update cell hyperlink",
                          ));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Link type" }),
                            i.jsxs("select", {
                              "aria-label": r("Link type"),
                              value: j,
                              disabled: $,
                              onChange: (se) => z(se.target.value),
                              lang: t("Link type"),
                              "data-ui-label": "Link type",
                              children: [
                                i.jsx("option", {
                                  value: "external",
                                  children: "Web or email address",
                                }),
                                i.jsx("option", {
                                  value: "internal",
                                  children: "Place in this workbook",
                                }),
                              ],
                            }),
                          ],
                        }),
                        j === "external"
                          ? i.jsxs("label", {
                              children: [
                                i.jsx(M, { children: "Address" }),
                                i.jsx("input", {
                                  "aria-label": r("Link address"),
                                  type: "text",
                                  value: V,
                                  disabled: $,
                                  placeholder: r(
                                    "https://example.com/evidence",
                                  ),
                                  maxLength: Jl.urlLength,
                                  onChange: (se) => D(se.target.value),
                                  lang: t("Link address"),
                                  "data-ui-label": "Link address",
                                }),
                              ],
                            })
                          : i.jsxs("div", {
                              className: "rv-link-target",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Worksheet" }),
                                    i.jsx("select", {
                                      "aria-label": r("Link worksheet"),
                                      value: B,
                                      disabled: $,
                                      onChange: (se) => _(se.target.value),
                                      lang: t("Link worksheet"),
                                      "data-ui-label": "Link worksheet",
                                      children: e.workbook.sheets.map((se) =>
                                        i.jsxs(
                                          "option",
                                          {
                                            value: se.id,
                                            children: [
                                              se.name,
                                              se.hidden ? " (hidden)" : "",
                                            ],
                                          },
                                          se.id,
                                        ),
                                      ),
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Cell address" }),
                                    i.jsx("input", {
                                      "aria-label": r("Link cell address"),
                                      value: C,
                                      maxLength: 12,
                                      disabled: $,
                                      onChange: (se) => T(se.target.value),
                                      lang: t("Link cell address"),
                                      "data-ui-label": "Link cell address",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Display text (optional)" }),
                            i.jsx("input", {
                              "aria-label": r("Link display text"),
                              value: N,
                              disabled: $,
                              maxLength: Jl.labelLength,
                              onChange: (se) => R(se.target.value),
                              lang: t("Link display text"),
                              "data-ui-label": "Link display text",
                            }),
                          ],
                        }),
                        K &&
                          i.jsxs("div", {
                            className: "rv-link-preview",
                            "aria-label": r("Link preview"),
                            lang: t("Link preview"),
                            "data-ui-label": "Link preview",
                            children: [
                              i.jsx(Link2, { size: 15 }),
                              i.jsx("span", {
                                children:
                                  K.label ||
                                  (K.kind === "external"
                                    ? K.url
                                    : `${((ge = e.workbook.sheets.find((se) => se.id === K.sheetId)) == null ? void 0 : ge.name) ?? "Missing worksheet"}!${K.address}`),
                              }),
                              K.kind === "external"
                                ? i.jsxs("a", {
                                    href: K.url,
                                    target: "_blank",
                                    rel: "noopener noreferrer",
                                    "aria-label": r("Open link preview"),
                                    lang: t("Open link preview"),
                                    "data-ui-label": "Open link preview",
                                    children: [
                                      i.jsx(ExternalLink, { size: 15 }),
                                      i.jsx(M, { children: "Open link" }),
                                    ],
                                  })
                                : e.onNavigate &&
                                  i.jsx("button", {
                                    type: "button",
                                    onClick: () => te(K.sheetId, K.address),
                                    children: i.jsx(M, {
                                      children: "Go to target",
                                    }),
                                  }),
                            ],
                          }),
                        i.jsx("p", {
                          className: "rv-muted",
                          children: i.jsx(M, {
                            children:
                              "Web links use https or http. Email links use one mailto recipient. Local files, embedded credentials and executable links are not accepted. Opening a link is a separate action.",
                          }),
                        }),
                        i.jsxs("div", {
                          className: "rv-link-actions",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              disabled: $ || !l,
                              onClick: () =>
                                q(
                                  () =>
                                    oD(e.workbook, e.sheetId, u, null, {
                                      readOnly: e.readOnly,
                                    }),
                                  "Remove cell hyperlink",
                                ),
                              children: i.jsx(M, { children: "Remove link" }),
                            }),
                            i.jsx("button", {
                              className: "rv-primary",
                              disabled:
                                $ || !(j === "external" ? V.trim() : C.trim()),
                              children: i.jsx(M, { children: "Save link" }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
            I &&
              i.jsx("p", { role: "alert", className: "rv-error", children: I }),
          ],
        }),
      ],
    }),
  });
}
const Aj = (e) => {
    try {
      const r = ct(e);
      return r.lastCol - r.firstCol < 500
        ? Array.from(
            { length: r.lastCol - r.firstCol + 1 },
            (t, n) => r.firstCol + n,
          )
        : [];
    } catch {
      return [];
    }
  },
  kj = (e) =>
    e === null ? "Blank" : typeof e == "string" ? JSON.stringify(e) : String(e),
  L4e = { gt: ">", gte: "≥", lt: "<", lte: "≤", eq: "=" };

export { Aj, B4e, L4e, j4e, kj };
