import { Search, X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import A2e from "react-dom";
import { Ig, Mt, b1, cx, ux } from "./part-02.js";
import {
  Af,
  Dg,
  Ip,
  Np,
  Nr,
  Og,
  Pv,
  RC,
  Vn,
  _t,
  ct,
  fB,
  hB,
  hi,
  ul,
  ut,
  vx,
  ys,
} from "./part-03.js";
import { Dr, Dx, M, tI } from "./part-04.js";
import { Lme, Pme } from "./part-33.js";
import {
  E_,
  Eo,
  aA,
  aH,
  cH,
  f2e,
  h2e,
  iH,
  lH,
  oA,
  oH,
  p2e,
  r0,
  sH,
  uH,
  xw,
} from "./part-52.js";

function m2e(e, r, t, n = {}) {
  var p, m, v;
  (n.readOnly && Eo("This workbook is read-only."),
    n.readOnly !== void 0 &&
      typeof n.readOnly != "boolean" &&
      Eo("Read-only must be boolean."),
    (typeof t != "string" || t.length > 32767) &&
      Eo("Replacement text must contain at most 32,767 characters."),
    r.lookIn &&
      !["raw", "formulas"].includes(r.lookIn) &&
      Eo(
        "Replace searches Raw contents or Formulas only. Calculated and formatted results are read-only search modes.",
      ));
  const o = cH(e, r),
    a = o.options,
    s = aH(a),
    d = oH(),
    u = n.match
      ? o.results.filter(
          (g) =>
            g.sheetId === n.match.sheetId &&
            g.address === n.match.address.toUpperCase(),
        )
      : o.results;
  n.match &&
    !u.length &&
    Eo(
      "The chosen cell does not match this search. Find it again before replacing.",
    );
  const l = [];
  let c = 0,
    h = 0;
  for (const g of u) {
    if (g.protected) {
      c++;
      continue;
    }
    const b = sH(g.text, a, s, d);
    let y = "",
      w = 0;
    g.text.length -
      b.reduce((F, [I, O]) => F + O - I, 0) +
      b.length * t.length >
      32767 &&
      Eo(
        `Replacement would exceed 32,767 characters at ${g.sheetName}!${g.address}.`,
      );
    for (const [F, I] of b) ((y += g.text.slice(w, F) + t), (w = I));
    y += g.text.slice(w);
    const A = Nr(e, g.sheetId).cells[g.address],
      k = p2e(A, y);
    JSON.stringify(k) !== JSON.stringify(hi(A)) &&
      (l.push({ sheetId: g.sheetId, address: g.address, cell: k }),
      (h += b.length),
      l.length > Mt.maxPatches &&
        Eo(
          "Replace may change at most 500 cells, including dependent formula changes. Narrow the scope or replace one result.",
        ));
  }
  const f = {
    targetSheetId:
      ((p = l[0]) == null ? void 0 : p.sheetId) ??
      ((m = e.sheets.find((g) => {
        var b;
        return !((b = g.protection) != null && b.enabled);
      })) == null
        ? void 0
        : m.id) ??
      a.sheetId,
    targetRange: ((v = l[0]) == null ? void 0 : v.address) ?? a.range,
    targetCount: l.length,
    matchedCells: u.length,
    replacedCells: l.length,
    replacedOccurrences: h,
    excludedProtected: c,
    summary: `Replace in ${a.scope}: ${l.length} cells; ${c} protected matches excluded.`,
  };
  return l.length
    ? Dx(e, l, f, n.readOnly)
    : oA({
        ...f,
        workbookId: e.id,
        revision: e.revision,
        patches: [],
        previewPatches: [],
        count: 0,
        dependentCount: 0,
        snapshotGuard: JSON.stringify(e),
      });
}
function g2e(e, r, t = {}) {
  return r.count
    ? tI(e, r, t)
    : (t.readOnly !== void 0 &&
        typeof t.readOnly != "boolean" &&
        Eo("Read-only must be boolean."),
      t.readOnly && Eo("This workbook is read-only."),
      (r.workbookId !== e.id ||
        r.revision !== e.revision ||
        r.snapshotGuard !== JSON.stringify(e)) &&
        Eo(
          "The workbook changed after this preview. Create a new preview before applying it.",
        ),
      e);
}
function v2e(e) {
  const r = [];
  for (const o of [...e].sort((a, s) => aA("rows", a, s))) {
    const a = r.at(-1);
    a && a.firstRow === o.row && a.lastCol + 1 === o.col
      ? (a.lastCol = o.col)
      : r.push({
          firstRow: o.row,
          lastRow: o.row,
          firstCol: o.col,
          lastCol: o.col,
        });
  }
  const t = [],
    n = new Map();
  for (const o of r) {
    const a = `${o.firstCol}:${o.lastCol}`,
      s = n.get(a);
    s && s.lastRow + 1 === o.firstRow
      ? (s.lastRow = o.lastRow)
      : (t.push(o), n.set(a, o));
  }
  return t.map((o) => {
    const a = xw(o.firstRow, o.firstCol),
      s = xw(o.lastRow, o.lastCol);
    return a === s ? a : `${a}:${s}`;
  });
}
function b2e(e, r, t) {
  const n = iH(e, r),
    o = Nr(e, n.sheetId),
    a = Vn(e),
    s = lH(o);
  (n.scope === "workbook" &&
    Eo("Go To Special selects cells on one sheet. Choose Selection or Sheet."),
    ["formulas", "errors", "constants", "blanks", "notes"].includes(t) ||
      Eo("Choose a supported Go To Special category."));
  const d = RC(o),
    u =
      n.scope === "selection"
        ? n.ranges
        : d
          ? [xw(d.firstRow, d.firstCol) + ":" + xw(d.lastRow, d.lastCol)]
          : [],
    l = u.map(ct),
    c = [],
    h = [],
    f = () => {
      h.sort((v, g) => aA(n.order, v, g));
      const m = v2e(h);
      return (
        m.length > r0.maxSpecialAreas &&
          Eo(
            "This selection contains more than 256 separate areas. Narrow the selection.",
          ),
        oA({
          sheetId: o.id,
          kind: t,
          count: h.length,
          addresses: h.map((v) => xw(v.row, v.col)),
          ranges: m,
          workbookId: e.id,
          revision: e.revision,
          snapshotGuard: JSON.stringify(e),
        })
      );
    };
  if (!l.length || (o.hidden && !n.includeHidden)) return f();
  if (t === "blanks")
    try {
      c.push(...Np(u, r0.maxBlankScan));
    } catch {
      Eo(
        "Blank selection scans at most 100,000 cells. Select a smaller rectangle.",
      );
    }
  else if ((c.push(...Object.keys(o.cells)), t === "errors"))
    for (const m of a.getSpillEntries(o.id))
      Object.hasOwn(o.cells, m.address) || c.push(m.address);
  c.length > Mt.maxCells && Eo("Go To Special scans at most 100,000 cells.");
  let p = 0;
  for (const m of c) {
    const v = _t(m);
    if (
      !l.some((w) => uH(v.row, v.col, w)) ||
      (!n.includeHidden && s(v.row, v.col))
    )
      continue;
    const g = o.cells[m],
      b = E_(g);
    let y = false;
    if (t === "formulas") y = b;
    else if (t === "notes") y = !!(g != null && g.note);
    else if (t === "constants") y = !!g && g.raw !== null && g.raw !== "" && !b;
    else {
      if (!b && g && g.raw !== null) continue;
      ++p > r0.maxCalculatedCells &&
        t === "errors" &&
        Eo(
          "Error selection calculates at most 10,000 cells. Narrow the selection.",
        );
      const w = a(o.id, m);
      t === "blanks"
        ? (y = ((g == null ? void 0 : g.raw) === null || !g) && w === null)
        : (y =
            (b || !!a.getSpillAnchor(o.id, m)) &&
            typeof w == "string" &&
            /^#(?:BLOCKED!|CALC!|CYCLE!|DIV\/0!|ERROR!|LIMIT!|N\/A|NAME\?|NULL!|NUM!|REF!|SPILL!|VALUE!)$/.test(
              w,
            ));
    }
    y &&
      (h.push(v),
      h.length > r0.maxSpecialCells &&
        Eo(
          "Go To Special may select at most 500 cells. Narrow the selection.",
        ));
  }
  return f();
}
const qm = (e) =>
    e instanceof Error ? e.message : "This operation could not be completed.",
  Pj = (e) =>
    e == null
      ? "(blank)"
      : typeof e == "boolean"
        ? e
          ? "TRUE"
          : "FALSE"
        : String(e),
  Bv = 20;
function y2e(e) {
  var ht;
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    s = ne.useRef(e);
  s.current = e;
  const d = ne.useRef({
      sheetId: e.sheetId,
      range: e.range,
      ranges: e.ranges ? [...e.ranges] : [e.range],
      sheetName:
        ((ht = e.workbook.sheets.find((Re) => Re.id === e.sheetId)) == null
          ? void 0
          : ht.name) ?? "Worksheet",
    }),
    [u, l] = ne.useState(e.initialTab ?? "find"),
    [c, h] = ne.useState("sheet"),
    [f, p] = ne.useState("raw"),
    [m, v] = ne.useState("rows"),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(false),
    [k, F] = ne.useState(false),
    [I, O] = ne.useState(true),
    [j, z] = ne.useState(false),
    [V, D] = ne.useState("formulas"),
    [B, _] = ne.useState(null),
    [C, T] = ne.useState(null),
    [N, R] = ne.useState(null),
    [L, $] = ne.useState(null),
    [U, G] = ne.useState(0),
    [H, Q] = ne.useState(0),
    [X$1, J] = ne.useState(""),
    [q, W] = ne.useState(""),
    [K, te] = ne.useState({}),
    ae = () => ({
      sheetId: d.current.sheetId,
      range: d.current.range,
      ranges: [...d.current.ranges],
      scope: c,
      query: g,
      lookIn: f,
      order: m,
      matchCase: E,
      wholeCell: k,
      wildcards: I,
      includeHidden: j,
    });
  function oe() {
    (_(null), T(null), R(null), $(null), G(0), Q(0), J(""), W(""));
  }
  function de(Re) {
    (l(Re),
      Re === "replace" && !["raw", "formulas"].includes(f ?? "raw") && p("raw"),
      Re === "special" && c === "workbook" && h("sheet"),
      oe());
  }
  ne.useEffect(() => {
    var Me, ke, je;
    const Re = document.activeElement;
    return (
      (je =
        a.current ??
        ((Me = o.current) == null
          ? void 0
          : Me.querySelector("select,input")) ??
        ((ke = o.current) == null ? void 0 : ke.querySelector("button"))) ==
        null || je.focus(),
      () => {
        Re != null && Re.isConnected && Re.focus();
      }
    );
  }, []);
  const he = (Re) => {
    const Me = s.current.workbook;
    if (
      Re.workbookId !== Me.id ||
      Re.revision !== Me.revision ||
      Re.snapshotGuard !== JSON.stringify(Me)
    )
      throw new Error(
        "The workbook changed. Find or preview again before using these results.",
      );
  };
  function Se() {
    try {
      (J(""), R(null), $(null));
      const Re = cH(s.current.workbook, ae());
      return (
        _(Re),
        T(null),
        G(0),
        W(
          Re.total
            ? `${Re.total.toLocaleString()} matching cells.`
            : "No matches found.",
        ),
        Re
      );
    } catch (Re) {
      return (J(qm(Re)), _(null), null);
    }
  }
  function le(Re, Me = B) {
    try {
      if (!Me) return;
      if (
        (he(Me),
        J(""),
        T(Re),
        G(
          Math.floor(
            Me.results.findIndex(
              (ke) => ke.sheetId === Re.sheetId && ke.address === Re.address,
            ) / Bv,
          ),
        ),
        Re.hidden)
      ) {
        W(
          `${Re.sheetName}!${Re.address} is in a ${Re.hiddenReason}. Unhide it before navigating; its visibility has not changed.`,
        );
        return;
      }
      (s.current.onNavigate(Re.sheetId, Re.address),
        W(
          `${Re.sheetName}!${Re.address}${Re.protected ? " · protected" : ""}`,
        ));
    } catch (ke) {
      J(qm(ke));
    }
  }
  function ce(Re) {
    const Me = B ?? Se();
    if (Me)
      try {
        he(Me);
        const ke = ct(d.current.range),
          je = C ?? {
            sheetId: d.current.sheetId,
            address: e.active ?? `${ut(ke.firstCol)}${ke.firstRow + 1}`,
          },
          $e = h2e(Me, je, Re);
        $e ? le($e, Me) : W("No matches found.");
      } catch (ke) {
        J(qm(ke));
      }
  }
  function ge(Re) {
    try {
      if ((J(""), Re && !C))
        throw new Error("Select a matching result before replacing one cell.");
      Re && B && he(B);
      const Me = m2e(s.current.workbook, ae(), y, {
        readOnly: s.current.readOnly,
        ...(Re ? { match: C } : {}),
      });
      (R(Me),
        te(
          Object.fromEntries(
            Me.previewPatches.map((ke) => {
              var $e;
              const je = Nr(s.current.workbook, ke.sheetId);
              return [
                `${ke.sheetId}:${ke.address}`,
                {
                  sheetName: je.name,
                  raw:
                    (($e = je.cells[ke.address]) == null ? void 0 : $e.raw) ??
                    null,
                },
              ];
            }),
          ),
        ),
        $(null),
        Q(0),
        W(
          Me.matchedCells
            ? `${Me.matchedCells} matched; ${Me.replacedCells} cells to replace; ${Me.excludedProtected} protected matches excluded.`
            : "No matches found.",
        ));
    } catch (Me) {
      (R(null), J(qm(Me)));
    }
  }
  function se() {
    if (N)
      try {
        J("");
        const Re = s.current.workbook,
          Me = g2e(Re, N, { readOnly: s.current.readOnly });
        (Me !== Re &&
          s.current.onWorkbookChange(Me, `Replace ${N.replacedCells} cells`),
          s.current.onClose());
      } catch (Re) {
        J(qm(Re));
      }
  }
  function we() {
    try {
      J("");
      const Re = b2e(s.current.workbook, ae(), V);
      ($(Re),
        R(null),
        G(0),
        W(
          Re.count
            ? `${Re.count} cells in ${Re.ranges.length} areas.`
            : "No matching cells found.",
        ));
    } catch (Re) {
      ($(null), J(qm(Re)));
    }
  }
  function Ee() {
    if (L != null && L.count)
      try {
        if ((J(""), he(L), !s.current.onSelectRanges))
          throw new Error(
            "This embedding host has not connected multiple-area selection.",
          );
        if (Nr(s.current.workbook, L.sheetId).hidden)
          throw new Error("Unhide this sheet before selecting its cells.");
        (s.current.onSelectRanges(L.sheetId, [...L.ranges]),
          s.current.onClose());
      } catch (Re) {
        J(qm(Re));
      }
  }
  const Ae = Math.max(1, Math.ceil(((B == null ? void 0 : B.total) ?? 0) / Bv)),
    ye = Math.min(U, Ae - 1),
    Te = Math.max(
      1,
      Math.ceil(((N == null ? void 0 : N.previewPatches.length) ?? 0) / Bv),
    ),
    Pe = Math.min(H, Te - 1),
    qe = !!(
      (B &&
        (B.workbookId !== e.workbook.id ||
          B.revision !== e.workbook.revision)) ||
      (N &&
        (N.workbookId !== e.workbook.id ||
          N.revision !== e.workbook.revision)) ||
      (L &&
        (L.workbookId !== e.workbook.id || L.revision !== e.workbook.revision))
    );
  return i.jsx("div", {
    className: "st-backdrop",
    onMouseDown: (Re) => {
      Re.target === Re.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "st-dialog",
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": n,
      onKeyDown: (Re) => {
        var Me, ke;
        if (Re.key === "Escape")
          (Re.preventDefault(), Re.stopPropagation(), s.current.onClose());
        else if (Re.key === "Tab") {
          const je = Array.from(
              ((Me = o.current) == null
                ? void 0
                : Me.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
                  )) ?? [],
            ).filter((ze) => !ze.hidden),
            $e = je[0],
            Je = je.at(-1);
          ((Re.shiftKey && document.activeElement === $e) ||
            (!Re.shiftKey && document.activeElement === Je)) &&
            (Re.preventDefault(),
            (ke = Re.shiftKey ? Je : $e) == null || ke.focus());
        } else
          Re.key === "Enter" &&
            Re.target instanceof HTMLInputElement &&
            Re.target.type !== "checkbox" &&
            (Re.preventDefault(), ce(Re.shiftKey ? "previous" : "next"));
      },
      children: [
        i.jsxs("header", {
          className: "st-header",
          children: [
            i.jsxs("h2", {
              id: n,
              children: [
                i.jsx(Search, { size: 18 }),
                i.jsx(M, { children: "Find and Replace" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": r("Close Find and Replace"),
              onClick: e.onClose,
              lang: t("Close Find and Replace"),
              "data-ui-label": "Close Find and Replace",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("div", {
          className: "st-tabs",
          role: "tablist",
          "aria-label": r("Search tools"),
          lang: t("Search tools"),
          "data-ui-label": "Search tools",
          children: [
            ["find", "Find"],
            ["replace", "Replace"],
            ["special", "Go To Special"],
          ].map(([Re, Me], ke) =>
            i.jsx(
              "button",
              {
                type: "button",
                role: "tab",
                "aria-selected": u === Re,
                tabIndex: u === Re ? 0 : -1,
                onClick: () => de(Re),
                onKeyDown: (je) => {
                  var $e, Je;
                  if (
                    ["ArrowLeft", "ArrowRight", "Home", "End"].includes(je.key)
                  ) {
                    je.preventDefault();
                    const ze = ["find", "replace", "special"],
                      at =
                        je.key === "Home"
                          ? 0
                          : je.key === "End"
                            ? 2
                            : (ke + (je.key === "ArrowRight" ? 1 : 2)) % 3;
                    (de(ze[at]),
                      (Je =
                        ($e = je.currentTarget.parentElement) == null
                          ? void 0
                          : $e.children[at]) == null || Je.focus());
                  }
                },
                children: Me,
              },
              Re,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "st-body",
          children: [
            i.jsxs("p", {
              className: "st-context",
              children: [
                i.jsx(M, { children: "Selection:" }),
                " ",
                i.jsxs("strong", {
                  children: [
                    d.current.sheetName,
                    "!",
                    d.current.ranges.join(", "),
                  ],
                }),
              ],
            }),
            i.jsxs("div", {
              className: "st-options",
              children: [
                u !== "special"
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          className: "st-wide",
                          children: [
                            i.jsx(M, { children: "Find what" }),
                            i.jsx("input", {
                              ref: a,
                              value: g,
                              maxLength: 256,
                              onChange: (Re) => {
                                (b(Re.target.value), oe());
                              },
                            }),
                          ],
                        }),
                        u === "replace" &&
                          i.jsxs("label", {
                            className: "st-wide",
                            children: [
                              i.jsx(M, { children: "Replace with" }),
                              i.jsx("input", {
                                value: y,
                                maxLength: 32767,
                                onChange: (Re) => {
                                  (w(Re.target.value), R(null), J(""));
                                },
                              }),
                            ],
                          }),
                      ],
                    })
                  : i.jsxs("label", {
                      className: "st-wide",
                      children: [
                        i.jsx(M, { children: "Select cells containing" }),
                        i.jsxs("select", {
                          value: V,
                          onChange: (Re) => {
                            (D(Re.target.value), oe());
                          },
                          children: [
                            i.jsx("option", {
                              value: "formulas",
                              children: "Formulas",
                            }),
                            i.jsx("option", {
                              value: "errors",
                              children: "Formula errors",
                            }),
                            i.jsx("option", {
                              value: "constants",
                              children: "Constants",
                            }),
                            i.jsx("option", {
                              value: "blanks",
                              children: "Blanks",
                            }),
                            i.jsx("option", {
                              value: "notes",
                              children: "Notes",
                            }),
                          ],
                        }),
                      ],
                    }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Within" }),
                    i.jsxs("select", {
                      value: c,
                      onChange: (Re) => {
                        (h(Re.target.value), oe());
                      },
                      children: [
                        i.jsx("option", {
                          value: "selection",
                          children: "Selection",
                        }),
                        i.jsx("option", { value: "sheet", children: "Sheet" }),
                        u !== "special" &&
                          i.jsx("option", {
                            value: "workbook",
                            children: "Workbook",
                          }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Search order" }),
                    i.jsxs("select", {
                      value: m,
                      onChange: (Re) => {
                        (v(Re.target.value), oe());
                      },
                      children: [
                        i.jsx("option", { value: "rows", children: "By rows" }),
                        i.jsx("option", {
                          value: "columns",
                          children: "By columns",
                        }),
                      ],
                    }),
                  ],
                }),
                u !== "special" &&
                  i.jsxs("label", {
                    className: "st-wide",
                    children: [
                      i.jsx(M, { children: "Look in" }),
                      i.jsxs("select", {
                        value: f,
                        onChange: (Re) => {
                          (p(Re.target.value), oe());
                        },
                        children: [
                          i.jsx("option", {
                            value: "raw",
                            children:
                              "Raw contents (text, numbers and formulas)",
                          }),
                          i.jsx("option", {
                            value: "formulas",
                            children: "Formulas only",
                          }),
                          u === "find" &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsx("option", {
                                  value: "calculated",
                                  children: "Calculated values",
                                }),
                                i.jsx("option", {
                                  value: "formatted",
                                  children: "Formatted display",
                                }),
                              ],
                            }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
            i.jsxs("div", {
              className: "st-switches",
              children: [
                u !== "special" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: E,
                            onChange: (Re) => {
                              (A(Re.target.checked), oe());
                            },
                          }),
                          i.jsx(M, { children: "Match case" }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: k,
                            onChange: (Re) => {
                              (F(Re.target.checked), oe());
                            },
                          }),
                          i.jsx(M, { children: "Match entire cell contents" }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: I,
                            onChange: (Re) => {
                              (O(Re.target.checked), oe());
                            },
                          }),
                          i.jsx(M, { children: "Use wildcards (* ? ~)" }),
                        ],
                      }),
                    ],
                  }),
                i.jsxs("label", {
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: j,
                      onChange: (Re) => {
                        (z(Re.target.checked), oe());
                      },
                    }),
                    i.jsx(M, { children: "Include hidden cells and sheets" }),
                  ],
                }),
              ],
            }),
            i.jsx("p", {
              className: "st-help",
              children:
                u === "special"
                  ? "Select up to 500 cells in 256 separate areas. Sheet scope uses the occupied rectangle for blanks. Formula results of empty text are not blank cells."
                  : "? matches one character; * matches any sequence. Prefix *, ? or ~ with ~ to find it literally. Search follows workbook sheet order.",
            }),
            e.readOnly &&
              i.jsx("p", {
                className: "st-notice",
                children: i.jsx(M, {
                  children:
                    "This workbook is read-only. Finding and selecting cells remain available.",
                }),
              }),
            qe &&
              i.jsx("p", {
                className: "st-notice",
                role: "status",
                children: i.jsx(M, {
                  children:
                    "The workbook changed. Find or preview again before continuing.",
                }),
              }),
            i.jsx("div", {
              className: "st-actions",
              children:
                u === "special"
                  ? i.jsx("button", {
                      type: "button",
                      onClick: we,
                      children: i.jsx(M, { children: "Preview selection" }),
                    })
                  : i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("button", {
                          type: "button",
                          onClick: Se,
                          children: i.jsx(M, { children: "Find all" }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => ce("previous"),
                          children: i.jsx(M, { children: "Find previous" }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => ce("next"),
                          children: i.jsx(M, { children: "Find next" }),
                        }),
                        u === "replace" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("button", {
                                type: "button",
                                disabled: e.readOnly || !C,
                                onClick: () => ge(true),
                                children: i.jsx(M, {
                                  children: "Preview replace selected",
                                }),
                              }),
                              i.jsx("button", {
                                type: "button",
                                disabled: e.readOnly,
                                onClick: () => ge(false),
                                children: i.jsx(M, {
                                  children: "Preview replace all",
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
            }),
            X$1 &&
              i.jsx("p", {
                className: "st-error",
                role: "alert",
                children: X$1,
              }),
            q &&
              i.jsx("p", {
                className: "st-status",
                role: "status",
                children: q,
              }),
            B &&
              i.jsxs("div", {
                className: "st-results",
                "aria-label": r("Find results"),
                lang: t("Find results"),
                "data-ui-label": "Find results",
                children: [
                  i.jsx("div", {
                    className: "st-table-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Sheet" }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Cell" }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, {
                                  children: "Matching contents",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "State" }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: f2e(B, ye, Bv).map((Re) =>
                            i.jsxs(
                              "tr",
                              {
                                className:
                                  (C == null ? void 0 : C.sheetId) ===
                                    Re.sheetId && C.address === Re.address
                                    ? "st-selected"
                                    : "",
                                children: [
                                  i.jsx("td", { children: Re.sheetName }),
                                  i.jsx("td", {
                                    children: i.jsx("button", {
                                      type: "button",
                                      "aria-label": `Go to ${Re.sheetName}!${Re.address}`,
                                      onClick: () => le(Re),
                                      children: Re.address,
                                    }),
                                  }),
                                  i.jsx("td", {
                                    className: "st-value",
                                    title: Re.text,
                                    children: Re.text,
                                  }),
                                  i.jsx("td", {
                                    children:
                                      [
                                        Re.protected ? "Protected" : "",
                                        Re.hidden ? Re.hiddenReason : "",
                                      ]
                                        .filter(Boolean)
                                        .join(", ") || "Visible",
                                  }),
                                ],
                              },
                              `${Re.sheetId}:${Re.address}`,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "st-pagination",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        "aria-label": r("Previous results page"),
                        disabled: !ye,
                        onClick: () => G(ye - 1),
                        lang: t("Previous results page"),
                        "data-ui-label": "Previous results page",
                        children: i.jsx(M, { children: "Previous" }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(M, { children: "Results page " }),
                          ye + 1,
                          " ",
                          i.jsx(M, { children: "of " }),
                          Ae,
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": r("Next results page"),
                        disabled: ye + 1 >= Ae,
                        onClick: () => G(ye + 1),
                        lang: t("Next results page"),
                        "data-ui-label": "Next results page",
                        children: i.jsx(M, { children: "Next" }),
                      }),
                    ],
                  }),
                ],
              }),
            N &&
              i.jsxs("div", {
                className: "st-preview",
                "aria-label": r("Replacement preview"),
                lang: t("Replacement preview"),
                "data-ui-label": "Replacement preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(M, { children: "Replacement preview" }),
                  }),
                  i.jsxs("p", {
                    children: [
                      N.replacedCells,
                      " ",
                      i.jsx(M, { children: "directly replaced cells · " }),
                      N.count,
                      " ",
                      i.jsx(M, { children: "total changed cells · " }),
                      N.excludedProtected,
                      " ",
                      i.jsx(M, { children: "protected matches excluded." }),
                    ],
                  }),
                  N.dependentCount > 0 &&
                    i.jsxs("p", {
                      className: "st-notice",
                      children: [
                        i.jsx(M, { children: "Includes " }),
                        N.dependentCount,
                        " ",
                        i.jsx(M, {
                          children:
                            "dependent cell changes outside the direct replacement targets. Every change is listed below.",
                        }),
                      ],
                    }),
                  i.jsx("div", {
                    className: "st-table-scroll",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(M, {
                                  children: "Sheet / cell",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Before" }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "After" }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: N.previewPatches
                            .slice(Pe * Bv, (Pe + 1) * Bv)
                            .map((Re) => {
                              var Me, ke;
                              return i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [
                                        ((Me =
                                          K[`${Re.sheetId}:${Re.address}`]) ==
                                        null
                                          ? void 0
                                          : Me.sheetName) ?? Re.sheetId,
                                        "!",
                                        Re.address,
                                      ],
                                    }),
                                    i.jsx("td", {
                                      className: "st-value",
                                      children: Pj(
                                        (ke =
                                          K[`${Re.sheetId}:${Re.address}`]) ==
                                          null
                                          ? void 0
                                          : ke.raw,
                                      ),
                                    }),
                                    i.jsx("td", {
                                      className: "st-value",
                                      children: Pj(Re.cell.raw),
                                    }),
                                  ],
                                },
                                `${Re.sheetId}:${Re.address}`,
                              );
                            }),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "st-pagination",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        "aria-label": r("Previous preview page"),
                        disabled: !Pe,
                        onClick: () => Q(Pe - 1),
                        lang: t("Previous preview page"),
                        "data-ui-label": "Previous preview page",
                        children: i.jsx(M, { children: "Previous" }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(M, { children: "Preview page " }),
                          Pe + 1,
                          " ",
                          i.jsx(M, { children: "of " }),
                          Te,
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        "aria-label": r("Next preview page"),
                        disabled: Pe + 1 >= Te,
                        onClick: () => Q(Pe + 1),
                        lang: t("Next preview page"),
                        "data-ui-label": "Next preview page",
                        children: i.jsx(M, { children: "Next" }),
                      }),
                    ],
                  }),
                ],
              }),
            L &&
              i.jsxs("div", {
                className: "st-preview",
                "aria-label": r("Selection preview"),
                lang: t("Selection preview"),
                "data-ui-label": "Selection preview",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(M, { children: "Selection preview" }),
                  }),
                  i.jsxs("p", {
                    children: [
                      L.count,
                      " ",
                      i.jsx(M, { children: "cells in " }),
                      L.ranges.length,
                      " ",
                      i.jsx(M, { children: "areas" }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "st-ranges",
                    children: L.ranges.map((Re) =>
                      i.jsx("code", { children: Re }, Re),
                    ),
                  }),
                  !e.onSelectRanges &&
                    i.jsx("p", {
                      className: "st-notice",
                      children: i.jsx(M, {
                        children:
                          "Multiple-area selection is not connected by this embedding host.",
                      }),
                    }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "st-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: i.jsx(M, { children: "Close" }),
            }),
            u === "replace" &&
              i.jsx("button", {
                type: "button",
                className: "st-primary",
                disabled: e.readOnly || !(N != null && N.count) || qe,
                onClick: se,
                children: i.jsx(M, { children: "Apply replacement" }),
              }),
            u === "special" &&
              i.jsx("button", {
                type: "button",
                className: "st-primary",
                disabled: !(L != null && L.count) || !e.onSelectRanges || qe,
                onClick: Ee,
                children: i.jsx(M, { children: "Select cells" }),
              }),
          ],
        }),
      ],
    }),
  });
}
function w2e(e, r, t, n) {
  if (!Number.isInteger(e) || !Number.isInteger(r) || r < 1 || e < 0 || e >= r)
    throw new Error("Navigation requires a valid visible cell.");
  if (n !== 1 && n !== -1) throw new Error("Choose a navigation direction.");
  const o = new Set(t.filter((d) => Number.isInteger(d) && d >= 0 && d < r)),
    a = e + n;
  if (a < 0 || a >= r) return e;
  if (o.has(e) && o.has(a)) {
    let d = a;
    for (; o.has(d + n);) d += n;
    return d;
  }
  let s = n === 1 ? r - 1 : 0;
  for (const d of o) (n === 1 ? d > e && d < s : d < e && d > s) && (s = d);
  return s;
}
const rE = ["top-left", "top-right", "bottom-left", "bottom-right"],
  x2e = () =>
    Object.fromEntries(
      rE.map((e) => [e, { top: 0, left: 0, width: 1e3, height: 600 }]),
    );
function E2e(e) {
  const [r, t] = ne.useState(x2e),
    n = ne.useRef(r),
    o = ne.useRef({}),
    [a, s] = ne.useState("top-left"),
    d = ne.useRef("top-left");
  function u(h) {
    let f = n.current;
    for (const p of rE)
      if (h[p]) {
        const m = { ...f[p], ...h[p] };
        Object.entries(m).some(([v, g]) => f[p][v] !== g) &&
          (f = { ...f, [p]: m });
      }
    f !== n.current && ((n.current = f), t(f));
  }
  const l = ne.useMemo(
    () =>
      Object.fromEntries(
        rE.map((h) => [
          h,
          (f) => {
            f ? (o.current[h] = f) : delete o.current[h];
          },
        ]),
      ),
    [],
  );
  ne.useEffect(() => {
    const h = () => {
      const p = {};
      for (const m of rE) {
        const v = o.current[m];
        v != null &&
          v.clientWidth &&
          v.clientHeight &&
          (p[m] = { width: v.clientWidth, height: v.clientHeight });
      }
      u(p);
    };
    if ((h(), typeof ResizeObserver > "u")) return;
    const f = new ResizeObserver(h);
    return (
      Object.values(o.current).forEach((p) => f.observe(p)),
      () => f.disconnect()
    );
  }, [e]);
  function c(h) {
    ((d.current = h), s(h));
  }
  return {
    views: r,
    current: n,
    refs: o,
    bindings: l,
    activeId: a,
    active: d,
    activate: c,
    patch: u,
  };
}
function C2e(e) {
  const { text: r, language: t } = Dr(),
    n = ne.useRef(null),
    o = ne.useRef(e);
  o.current = e;
  const [a, s] = ne.useState(null),
    d = ne.useRef(a);
  d.current = a;
  const [u, l] = ne.useState(""),
    c =
      a && a.snapshot === e.snapshot && !e.readOnly
        ? { ...e.split, [a.axis]: a.value }
        : e.split,
    h = (f, p) => {
      o.current.readOnly ||
        !o.current.split ||
        e.onResize({ ...o.current.split, [f]: Math.max(20, Math.min(80, p)) });
    };
  return i.jsxs("div", {
    ref: n,
    className: `lg-pane-layout ${c ? "is-split" : ""}`,
    style: {
      gridTemplateColumns:
        (c == null ? void 0 : c.vertical) !== void 0
          ? `minmax(0,${c.vertical}fr) 7px minmax(0,${100 - c.vertical}fr)`
          : "minmax(0,1fr)",
      gridTemplateRows:
        (c == null ? void 0 : c.horizontal) !== void 0
          ? `minmax(0,${c.horizontal}fr) 7px minmax(0,${100 - c.horizontal}fr)`
          : "minmax(0,1fr)",
    },
    children: [
      Pv(c).map((f) =>
        i.jsx(
          "div",
          {
            className: `lg-pane ${e.active === f ? "is-active" : ""}`,
            "data-pane": f,
            style: {
              gridRow: f.startsWith("bottom") ? 3 : 1,
              gridColumn: f.endsWith("right") ? 3 : 1,
            },
            children: e.renderPane(f),
          },
          f,
        ),
      ),
      ["horizontal", "vertical"].map((f) =>
        (c == null ? void 0 : c[f]) === void 0
          ? null
          : i.jsx(
              "div",
              {
                className: `lg-split-divider lg-split-${f}`,
                role: "separator",
                tabIndex: e.readOnly ? -1 : 0,
                "aria-label": `Resize ${f} split`,
                "aria-orientation": f,
                "aria-valuenow": c[f],
                "aria-valuemin": 20,
                "aria-valuemax": 80,
                "aria-disabled": e.readOnly,
                title: r(
                  "Drag or use arrow keys to resize; double-click to center",
                ),
                style:
                  f === "horizontal"
                    ? { gridRow: 2, gridColumn: "1 / -1" }
                    : { gridColumn: 2, gridRow: "1 / -1" },
                onDoubleClick: () => h(f, 50),
                onKeyDown: (p) => {
                  if (p.key === "Escape") {
                    ((d.current = null), s(null));
                    return;
                  }
                  const m = (
                    f === "horizontal"
                      ? ["ArrowUp", "ArrowDown"]
                      : ["ArrowLeft", "ArrowRight"]
                  ).indexOf(p.key);
                  m >= 0
                    ? (p.preventDefault(),
                      h(f, c[f] + (m ? 1 : -1) * (p.shiftKey ? 10 : 1)))
                    : (p.key === "Home" || p.key === "End") &&
                      (p.preventDefault(), h(f, p.key === "Home" ? 20 : 80));
                },
                onPointerDown: (p) => {
                  var v, g;
                  if (e.readOnly || p.button !== 0) return;
                  (p.preventDefault(), p.currentTarget.focus(), l(""));
                  const m = {
                    axis: f,
                    value: c[f],
                    snapshot: e.snapshot,
                    snapshotGuard: JSON.stringify(e.snapshot),
                    split: e.split,
                  };
                  ((d.current = m),
                    s(m),
                    (g = (v = p.currentTarget).setPointerCapture) == null ||
                      g.call(v, p.pointerId));
                },
                onPointerMove: (p) => {
                  var w;
                  const m = d.current,
                    v =
                      (w = n.current) == null
                        ? void 0
                        : w.getBoundingClientRect();
                  if (!m || !v) return;
                  const g = f === "horizontal" ? v.height : v.width;
                  if (g <= 0) return;
                  const b =
                      f === "horizontal"
                        ? p.clientY - v.top
                        : p.clientX - v.left,
                    y = {
                      ...m,
                      value: Math.max(
                        20,
                        Math.min(80, Math.round((b / g) * 100)),
                      ),
                    };
                  ((d.current = y), s(y));
                },
                onPointerUp: (p) => {
                  var v, g;
                  const m = d.current;
                  if (((d.current = null), s(null), !!m)) {
                    if (
                      m.snapshot !== o.current.snapshot ||
                      m.snapshotGuard !== JSON.stringify(o.current.snapshot) ||
                      o.current.readOnly
                    ) {
                      l("The workbook changed. Resize the split again.");
                      return;
                    }
                    (m.value !== m.split[f] && h(f, m.value),
                      (g = (v = p.currentTarget).hasPointerCapture) != null &&
                        g.call(v, p.pointerId) &&
                        p.currentTarget.releasePointerCapture(p.pointerId));
                  }
                },
                onPointerCancel: () => {
                  ((d.current = null), s(null));
                },
                onLostPointerCapture: () => {
                  ((d.current = null), s(null));
                },
                lang: t(
                  "Drag or use arrow keys to resize; double-click to center",
                ),
              },
              f,
            ),
      ),
      u &&
        i.jsx("span", {
          role: "alert",
          className: "lg-split-error",
          children: u,
        }),
    ],
  });
}

function dH(e, r) {
  var n;
  if (!(r instanceof Node)) return false;
  const t = document.getElementById(e.getAttribute("aria-controls") ?? "");
  return !!(
    (n = t == null ? void 0 : t.closest(".formula-suggestions")) != null &&
    n.contains(r)
  );
}
const fH = ne.forwardRef(function (
    {
      value: r,
      caret: t = r.length,
      workbook: n,
      sheetId: o,
      inputRef: a,
      onAccept: s,
      disabled: d,
      onExit: u,
    },
    l,
  ) {
    const c = ne.useId(),
      h = ne.useRef(null),
      f = ne.useMemo(() => Lme(r, t, n, o), [r, t, n, o]),
      [p, m] = ne.useState(-1),
      [v, g] = ne.useState(null),
      [b, y] = ne.useState(false),
      [w, E] = ne.useState({ left: 0, top: 0, width: 280 }),
      A = `${r}:${t}`,
      k = !d && b && v !== A && (f.items.length > 0 || f.hint);
    (ne.useEffect(() => {
      const I = a.current;
      if (!I) return;
      const O = ["aria-autocomplete", "aria-controls", "aria-activedescendant"],
        j = O.map((z) => I.getAttribute(z));
      return (
        k && f.items.length
          ? (I.setAttribute("aria-autocomplete", "list"),
            I.setAttribute("aria-controls", `${c}-list`),
            p >= 0 && f.items[p]
              ? I.setAttribute("aria-activedescendant", `${c}-${p}`)
              : I.removeAttribute("aria-activedescendant"))
          : (I.removeAttribute("aria-controls"),
            I.removeAttribute("aria-activedescendant")),
        () =>
          O.forEach((z, V) =>
            j[V] === null ? I.removeAttribute(z) : I.setAttribute(z, j[V]),
          )
      );
    }, [k, p, f.items.length, a, c]),
      ne.useEffect(() => m(-1), [A]),
      ne.useEffect(() => {
        const I = a.current;
        if (!I) return;
        const O = () => {
            var V;
            y(
              document.activeElement === I ||
                !!(
                  (V = h.current) != null && V.contains(document.activeElement)
                ),
            );
            const z = I.getBoundingClientRect();
            E({
              left: Math.max(8, Math.min(z.left, window.innerWidth - 350)),
              top: z.bottom + 3,
              width: Math.max(280, Math.min(400, z.width)),
            });
          },
          j = (z) => {
            var V;
            ((V = h.current) != null && V.contains(z.relatedTarget)) ||
              y(false);
          };
        return (
          O(),
          I.addEventListener("focus", O),
          I.addEventListener("blur", j),
          window.addEventListener("resize", O),
          window.addEventListener("scroll", O, true),
          () => {
            (I.removeEventListener("focus", O),
              I.removeEventListener("blur", j),
              window.removeEventListener("resize", O),
              window.removeEventListener("scroll", O, true));
          }
        );
      }, [a]),
      ne.useEffect(() => {
        if (!k || p < 0) return;
        const I = document.getElementById(`${c}-${p}`),
          O = h.current;
        if (!I || !O) return;
        const j = I.getBoundingClientRect(),
          z = O.getBoundingClientRect();
        j.bottom > z.bottom
          ? (O.scrollTop += j.bottom - z.bottom)
          : j.top < z.top && (O.scrollTop -= z.top - j.top);
      }, [k, p, c]));
    const F = (I) => {
      var z;
      const O = f.items[I];
      if (!O) return;
      (z = a.current) == null || z.focus({ preventScroll: true });
      const j = Pme(r, O);
      (s(j.value, j.caret),
        m(-1),
        g(null),
        requestAnimationFrame(() => {
          var V, D;
          ((V = a.current) == null || V.focus(),
            (D = a.current) == null || D.setSelectionRange(j.caret, j.caret));
        }));
    };
    return (
      ne.useImperativeHandle(l, () => ({
        handleKeyDown(I) {
          return I.nativeEvent.isComposing || I.defaultPrevented || !k
            ? false
            : (I.key === "ArrowDown" || I.key === "ArrowUp") && f.items.length
              ? (I.preventDefault(),
                m((O) =>
                  I.key === "ArrowDown"
                    ? (O + 1) % f.items.length
                    : O <= 0
                      ? f.items.length - 1
                      : O - 1,
                ),
                true)
              : I.key === "Tab" &&
                  !I.shiftKey &&
                  p < 0 &&
                  f.items.length &&
                  /[A-Za-z_][A-Za-z0-9_.]*(?:\[[^\]]*)?$/.test(r.slice(0, t))
                ? (I.preventDefault(), F(0), true)
                : (I.key === "Tab" || I.key === "Enter") && p >= 0 && f.items[p]
                  ? (I.preventDefault(), F(p), true)
                  : I.key === "Escape"
                    ? (I.preventDefault(), g(A), true)
                    : false;
        },
      })),
      !k || typeof document > "u"
        ? null
        : A2e.createPortal(
            i.jsxs("div", {
              className: "formula-suggestions",
              ref: h,
              role: "region",
              "aria-label": "Formula assistance",
              onBlur: (I) => {
                I.currentTarget.contains(I.relatedTarget) ||
                  I.relatedTarget === a.current ||
                  (y(false), u == null || u());
              },
              style: { position: "fixed", ...w },
              children: [
                f.items.length > 0 &&
                  i.jsx("ul", {
                    id: `${c}-list`,
                    role: "listbox",
                    "aria-label": "Formula suggestions",
                    tabIndex: 0,
                    "aria-activedescendant":
                      p >= 0 && f.items[p] ? `${c}-${p}` : void 0,
                    onFocus: () => {
                      (y(true), p < 0 && m(0));
                    },
                    onKeyDown: (I) => {
                      var O;
                      I.nativeEvent.isComposing ||
                        (I.key === "ArrowDown" || I.key === "ArrowUp"
                          ? (I.preventDefault(),
                            m((j) =>
                              I.key === "ArrowDown"
                                ? (j + 1) % f.items.length
                                : (j <= 0 ? f.items.length : j) - 1,
                            ))
                          : I.key === "Enter" ||
                              (I.key === "Tab" && !I.shiftKey)
                            ? (I.preventDefault(), F(Math.max(0, p)))
                            : (I.key === "Escape" || I.key === "Tab") &&
                              (I.preventDefault(),
                              I.key === "Escape" && g(A),
                              (O = a.current) == null || O.focus()));
                    },
                    children: f.items.map((I, O) =>
                      i.jsxs(
                        "li",
                        {
                          id: `${c}-${O}`,
                          role: "option",
                          "aria-selected": p === O,
                          onClick: (j) => {
                            j.detail === 0 && F(O);
                          },
                          onMouseDown: (j) => {
                            (j.preventDefault(), F(O));
                          },
                          onMouseEnter: () => m(O),
                          children: [
                            i.jsx("span", {
                              className: "formula-suggestion-kind",
                              children:
                                I.kind === "function"
                                  ? "ƒx"
                                  : I.kind === "table"
                                    ? "▦"
                                    : "N",
                            }),
                            i.jsxs("div", {
                              children: [
                                i.jsx("strong", { children: I.label }),
                                i.jsx("small", { children: I.detail }),
                              ],
                            }),
                          ],
                        },
                        `${I.kind}:${I.label}`,
                      ),
                    ),
                  }),
                f.hint &&
                  i.jsxs("div", {
                    className: "formula-argument-hint",
                    role: "status",
                    "aria-live": "polite",
                    "aria-atomic": "true",
                    children: [
                      i.jsx("strong", { children: f.hint.name }),
                      "(",
                      f.hint.signature,
                      ")",
                      i.jsxs("small", {
                        children: [
                          i.jsx(M, { children: "Argument " }),
                          f.hint.argumentIndex + 1,
                        ],
                      }),
                    ],
                  }),
                f.items.length > 0 &&
                  i.jsx("div", {
                    className: "formula-suggestion-help",
                    children: i.jsx(M, {
                      children: "↑ ↓ choose · Tab / Enter insert · Esc close",
                    }),
                  }),
              ],
            }),
            document.body,
          )
    );
  }),
  k2e = 6,
  S2e = 1.35,
  hH = (e, r, t) => Math.max(r, Math.min(t, e));
function pH(e) {
  const r = ys(e == null ? void 0 : e.style),
    t = r != null && r.fontSize ? `${r.fontSize}pt` : "14px",
    n =
      r != null && r.fontFamily
        ? JSON.stringify(r.fontFamily)
        : "Calibri, Aptos, Arial, sans-serif";
  return `${r != null && r.italic ? "italic" : "normal"} ${r != null && r.bold ? "700" : "400"} ${t} ${n}`;
}
function Wy(e, r) {
  const t = r.match(/([\d.]+)(px|pt)/),
    n = t ? Number(t[1]) * (t[2] === "pt" ? 4 / 3 : 1) : 14,
    o = r.includes("700") ? 1.06 : 1;
  return (
    Array.from(e).reduce(
      (a, s) =>
        a +
        (/\s/u.test(s)
          ? 0.33
          : /[ilI.,'`!:;]/.test(s)
            ? 0.28
            : /[MW@#]/.test(s)
              ? 0.85
              : /[^\u0000-\u024f]/u.test(s)
                ? 1
                : 0.56),
      0,
    ) *
    n *
    o
  );
}
function T2e() {
  if (typeof document > "u" || typeof CanvasRenderingContext2D > "u") return Wy;
  try {
    const e = document.createElement("canvas").getContext("2d");
    if (!e) return Wy;
    let r = "";
    return (t, n) => {
      r !== n && ((e.font = n), (r = n));
      const o = e.measureText(t).width;
      return Number.isFinite(o) && o >= 0 ? o : Wy(t, n);
    };
  } catch {
    return Wy;
  }
}
function nE(e, r, t) {
  try {
    const n = e(r, t);
    if (Number.isFinite(n) && n >= 0) return n;
  } catch {}
  return Wy(r, t);
}
function mH(e, r, t) {
  if (!e.includes("	")) return nE(t, e, r);
  const n = Math.max(1, nE(t, " ", r) * 8);
  let o = 0;
  return (
    e.split("	").forEach((a, s) => {
      (s && (o = (Math.floor(o / n) + 1) * n), (o += nE(t, a, r)));
    }),
    o
  );
}
function gH(e, r, t, n, o, a) {
  return o && typeof t.raw == "string" && t.raw.startsWith("=") && !t.literal
    ? t.raw
    : Af(n(e.id, r), t, a);
}
function N2e(e, r) {
  return Math.max(
    e.note ? 5 : 0,
    e.tickmark
      ? Math.min(40, nE(r, e.tickmark, "normal 400 8px Arial") + 4)
      : 0,
  );
}
function vH(e, r) {
  var t;
  return (
    k2e +
    hB(e.style).horizontal +
    (((t = e.style) == null ? void 0 : t.align) === "center"
      ? 0
      : fB(e.style)) +
    N2e(e, r)
  );
}
function bH(e, r, t) {
  for (const n of e)
    if (!Number.isInteger(n) || n < 0 || n >= r)
      throw new Error(`AutoFit ${t} index is outside the worksheet.`);
  return new Set(e);
}
function R2e(e, r, t, n, o = false, a = "1900") {
  const s = bH(r, e.colCount, "column"),
    d = Object.fromEntries([...s].map((l) => [String(l), Og(e)])),
    u = new Set();
  for (const [l, c] of Object.entries(vx(e, t))) {
    const { col: h } = _t(l);
    if (!s.has(h)) continue;
    const f = gH(e, l, c, t, o, a);
    if (!f && !c.note && !c.tickmark) continue;
    const p = pH(c),
      m = ys(c.style).wrap
        ? f.replace(
            /\r\n?/g,
            `
`,
          ).split(`
`)
        : [f.replace(/[\t\n\r\f ]+/g, " ")];
    let v = 0;
    for (const b of m) v = Math.max(v, mH(b, p, n));
    const g = hH(Math.ceil(v + vH(c, n)), ux, cx);
    ((d[String(h)] = u.has(h) ? Math.max(d[String(h)], g) : g), u.add(h));
  }
  return d;
}
function I2e(e, r, t, n, o) {
  let a = 0;
  for (const s of e.replace(
    /\r\n?/g,
    `
`,
  ).split(`
`)) {
    const d = Array.from(s);
    if (!d.length) {
      if (++a >= o) return o;
      continue;
    }
    let u = 0;
    for (; u < d.length;) {
      if (++a >= o) return o;
      let l = u + 1,
        c = d.length;
      for (; l < c;) {
        const f = Math.ceil((l + c) / 2);
        mH(d.slice(u, f).join(""), t, n) <= r ? (l = f) : (c = f - 1);
      }
      let h = l;
      if (h >= d.length) break;
      for (let f = h - 1; f > u; f--)
        if (/[\t \-\u200b]/u.test(d[f])) {
          h = f + 1;
          break;
        }
      for (; h < d.length && /[\t ]/u.test(d[h]);) h++;
      u = h;
    }
  }
  return Math.max(1, a);
}
function _2e(e, r, t, n, o = false, a = "1900") {
  const s = bH(r, e.rowCount, "row"),
    d = Object.fromEntries([...s].map((l) => [String(l), Dg(e)])),
    u = new Set();
  for (const [l, c] of Object.entries(vx(e, t))) {
    const { row: h, col: f } = _t(l);
    if (!s.has(h)) continue;
    const p = gH(e, l, c, t, o, a);
    if (!p && !ys(c.style).fontSize) continue;
    const m = pH(c),
      g = (ys(c.style).fontSize ? (ys(c.style).fontSize * 4) / 3 : 14) * S2e,
      b = hB(c.style).vertical,
      y = Math.ceil((Ig - b) / g),
      w = Math.max(1, Ip(e, f) - vH(c, n)),
      E = ys(c.style).wrap ? I2e(p, w, m, n, y) : 1,
      A = hH(Math.ceil(Math.max(20, E * g) + b), b1, Ig);
    ((d[String(h)] = u.has(h) ? Math.max(d[String(h)], A) : A), u.add(h));
  }
  return d;
}
const D2e = 16e6;
function Tp(e, r, t) {
  return r >= e.firstRow && r <= e.lastRow && t >= e.firstCol && t <= e.lastCol;
}
const sA = (e, r, t) => Math.max(r, Math.min(t, e));
function O2e(e, r) {
  let t = 0,
    n = e.length;
  for (; t < n;) {
    const o = (t + n) >>> 1;
    e[o] < r ? (t = o + 1) : (n = o);
  }
  return t;
}
function CS(e, r, t = {}, n = []) {
  const o = ul(e, n),
    a = Object.fromEntries(
      Object.entries(t).flatMap(([l, c]) => {
        const h = o.viewIndex(Number(l));
        return h === null ? [] : [[h, c]];
      }),
    );
  e = o.count;
  const s = Object.keys(a)
      .map(Number)
      .filter((l) => l >= 0 && l < e && a[l] > 0 && a[l] !== r)
      .sort((l, c) => l - c),
    d = [0];
  for (const l of s) d.push(d.at(-1) + a[l] - r);
  const u = (l) => {
    const c = sA(l, 0, e);
    return c * r + d[O2e(s, c)];
  };
  return {
    ...o,
    total: u(e),
    offset: u,
    size: (l) => a[l] ?? r,
    indexAt: (l) => {
      let c = 0,
        h = e;
      for (; c < h;) {
        const f = Math.floor((c + h) / 2);
        u(f + 1) <= l ? (c = f + 1) : (h = f);
      }
      return Math.min(Math.max(0, e - 1), c);
    },
  };
}
function Uj(e, r) {
  if (!Number.isInteger(r) || r < 0 || r > e.count)
    throw Error("Choose a valid visible-axis slice.");
  const t = e.offset(r),
    n = e.count - r,
    o = (d) =>
      Number.isInteger(d) && d >= 0 && d < n ? e.logicalIndex(d + r) : -1,
    a = (d) => {
      const u = e.viewIndex(d);
      return u === null || u < r ? null : u - r;
    },
    s = (d) => Math.max(0, e.countBefore(d) - r);
  return {
    count: n,
    logicalCount: e.logicalCount,
    logicalIndex: o,
    viewIndex: a,
    countBefore: s,
    isHidden: (d) => a(d) === null,
    next: (d, u, l = false) => {
      const c = d + (l ? 0 : u),
        h =
          u === 1 ? s(Math.max(0, c)) : s(Math.min(e.logicalCount, c + 1)) - 1,
        f = o(h);
      return f < 0 ? null : f;
    },
    total: e.total - t,
    size: (d) => e.size(d + r),
    offset: (d) => e.offset(Math.max(0, Math.min(n, d)) + r) - t,
    indexAt: (d) =>
      n ? Math.max(0, Math.min(n - 1, e.indexAt(d + t) - r)) : 0,
  };
}
function C_(e, r, t, n = D2e) {
  const o = Math.max(1, r),
    a = e * t,
    s = Math.min(a, n);
  return {
    viewport: o,
    logicalPixels: a,
    physicalPixels: s,
    logicalMax: Math.max(0, a - o),
    physicalMax: Math.max(0, s - o),
  };
}
function M2e(e, r, t, n) {
  const o = C_(r, t, n);
  return o.physicalMax
    ? ((sA(e, 0, o.physicalMax) / o.physicalMax) * o.logicalMax) / n
    : 0;
}
function up(e, r, t, n) {
  const o = C_(r, t, n);
  return o.logicalMax
    ? (sA(e * n, 0, o.logicalMax) / o.logicalMax) * o.physicalMax
    : 0;
}

export {
  C2e,
  CS,
  C_,
  E2e,
  M2e,
  R2e,
  T2e,
  Tp,
  Uj,
  _2e,
  dH,
  fH,
  sA,
  up,
  w2e,
  y2e,
};
