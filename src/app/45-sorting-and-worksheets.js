import {
  X,
  Plus,
  ArrowDown,
  ArrowUp,
  FileSpreadsheet,
  Copy,
  EyeOff,
  Eye as Eye$1,
  LockKeyhole,
  Trash2,
  LockKeyholeOpen,
} from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  AP,
  CP,
  EO,
  Hne,
  Nr,
  Pne,
  Une,
  Vne,
  Wne,
  _t,
  cn,
  ct,
  jn,
  jne,
  rn,
  ut,
  v0,
  xO,
  zne,
} from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import { LC } from "./27-data-validation-dialog.js";
import {
  BW,
  FW,
  LW,
  MW,
  N3,
  OW,
  _1,
  b_,
  c4e,
  jW,
  u4e,
  yd,
} from "./44-clean-and-reshape.js";

function d4e(e, r, t = {}) {
  var V;
  const n = BW(e, r.sheetId, t),
    o = u4e(e, r),
    a = b_(e, o.sheetId, o.range, o.headers);
  if (a.expansionRequired && !o.expandSelection)
    throw Error(
      `This selection is part of populated records. Expand the selection to ${a.expandedRange}, or cancel; partial-record sorting is blocked.`,
    );
  const s = ct(a.expandedRange),
    d = a.lastRow - a.firstDataRow + 1;
  if (d < 2)
    throw Error("Choose at least two data rows. Check the header setting.");
  if (d > yd.rows || d * (s.lastCol - s.firstCol + 1) > yd.cells)
    throw Error(
      "Sort supports at most 10,000 rows and 100,000 record cells per reviewed operation.",
    );
  const u = cn(n, "row"),
    l = Array.from({ length: d }, (B, _) => a.firstDataRow + _).filter((B) =>
      u.some((_) => B >= _.start && B <= _.end),
    ).length;
  if (l && !o.includeHidden)
    throw Error(
      `The source contains ${l} hidden rows. Include hidden rows explicitly or cancel so records stay together.`,
    );
  c4e(e, d, t);
  const c = LW(e, o, a),
    h = JSON.stringify(c),
    f = new Intl.Collator("en-US", {
      usage: "sort",
      sensitivity: o.caseSensitive ? "variant" : "accent",
      numeric: false,
    }),
    p = (B) =>
      o.caseSensitive
        ? B.normalize("NFC")
        : B.normalize("NFC").toLocaleLowerCase("en-US"),
    m = o.keys.map((B) => {
      return new Map(B.customOrder?.map((C, T) => [p(C), T]));
    }),
    v = (B) =>
      B.compare === null || B.compare === ""
        ? 4
        : B.error
          ? 3
          : typeof B.compare == "number"
            ? 0
            : typeof B.compare == "string"
              ? 1
              : 2,
    g = [...c].sort((B, _) => {
      for (let C = 0; C < o.keys.length; C++) {
        const T = B.keys[C],
          N = _.keys[C],
          R = v(T),
          L = v(N),
          $ = o.keys[C].direction === "ascending" ? 1 : -1;
        if (R === 4 || L === 4) {
          if (R !== L) return R === 4 ? 1 : -1;
          continue;
        }
        if (m[C].size) {
          const G = m[C].get(p(String(T.compare))),
            H = m[C].get(p(String(N.compare)));
          if (G !== undefined || H !== undefined) {
            if (G === undefined) return 1;
            if (H === undefined) return -1;
            if (G !== H) return (G - H) * $;
          }
        }
        const U =
          R !== L
            ? R - L
            : typeof T.compare == "number" && typeof N.compare == "number"
              ? T.compare - N.compare
              : typeof T.compare == "boolean" && typeof N.compare == "boolean"
                ? Number(T.compare) - Number(N.compare)
                : f.compare(String(T.compare), String(N.compare));
        if (U) return U * $;
      }
      return B.row - _.row;
    }),
    b = Array(d).fill(0);
  g.forEach((B, _) => (b[B.row - a.firstDataRow] = a.firstDataRow + _));
  const y = g.filter((B, _) => B.row !== a.firstDataRow + _).length,
    w = {
      axis: "row",
      index: a.firstDataRow,
      count: d,
      remove: false,
      moveBefore: a.firstDataRow,
      permutation: b,
    };
  let E = y ? LC(e, n.id, w) : structuredClone(e);
  const A = [],
    k = [],
    F = [],
    I = (B) =>
      B >= a.firstDataRow && B <= a.lastRow ? b[B - a.firstDataRow] : B;
  for (const B of e.sheets) {
    const _ = Nr(E, B.id);
    for (const [C, T] of Object.entries(B.cells))
      if (N3(T)) {
        const N = _t(C),
          R = B.id === n.id ? _1(I(N.row), N.col) : C,
          L = _.cells[R]?.raw;
        L !== T.raw &&
          k.push({
            sheet: B.name,
            source: C,
            destination: R,
            before: String(T.raw),
            after: String(L),
          });
      }
    for (const C of new Set([...Object.keys(B.cells), ...Object.keys(_.cells)]))
      JSON.stringify(B.cells[C]) !== JSON.stringify(_.cells[C]) &&
        A.push({
          sheetId: B.id,
          address: C,
          cell: _.cells[C] ?? { raw: null },
        });
    for (const C of Object.keys(_).filter((T) => T !== "cells"))
      JSON.stringify(B[C]) !== JSON.stringify(_[C]) &&
        F.push({ sheet: B.name, field: C });
  }
  if (
    (JSON.stringify(e.definedNames) !== JSON.stringify(E.definedNames) &&
      F.push({ sheet: "Workbook", field: "definedNames" }),
    JSON.stringify(e.formulaWatches) !== JSON.stringify(E.formulaWatches) &&
      F.push({ sheet: "Workbook", field: "formulaWatches" }),
    A.length > yd.cells)
  )
    throw Error(
      "The sort would change more than 100,000 cells including dependent references.",
    );
  JSON.stringify(e.reviewThreads) !== JSON.stringify(E.reviewThreads) &&
    F.push({ sheet: "Workbook", field: "reviewThreads" });
  for (const B of e.sheets)
    B.id !== n.id && JSON.stringify(B) !== JSON.stringify(Nr(E, B.id)) && jn(B);
  if ((v0(E, A), y)) {
    const B = Nr(E, n.id);
    ((E.rowOrigins = [...(E.rowOrigins ?? [])]),
      (B.rowOriginIds = { ...B.rowOriginIds }));
    const _ = new Set(
      Object.entries(n.cells)
        .filter(([, C]) => jW(C))
        .map(([C]) => _t(C).row),
    );
    for (const C of g)
      if (_.has(C.row) && !((V = n.rowOriginIds) != null && V[C.row])) {
        if (E.rowOrigins.length >= 1e5)
          throw Error("Source row identity capacity has been reached.");
        const T = crypto.randomUUID();
        ((B.rowOriginIds[I(C.row)] = T),
          E.rowOrigins.push({
            id: T,
            sourceId: e.id,
            sourceName: e.name,
            sourceSheet: n.name,
            originalRow: C.row + 1,
            kind: "snapshot",
          }));
      }
    E = rn(E);
  }
  const O = new Map(E.rowOrigins?.map((B) => [B.id, B])),
    j = {
      request: o,
      selection: a,
      snapshot: JSON.stringify(e),
      rows: g.map((B, _) => {
        const C = Nr(E, n.id).rowOriginIds?.[a.firstDataRow + _],
          T = C ? O.get(C) : undefined;
        return {
          sourceRow: B.row + 1,
          destinationRow: a.firstDataRow + _ + 1,
          ...(C
            ? {
                sourceId: C,
                sourceName: T?.sourceName,
                originalRow: T?.originalRow,
              }
            : {}),
          keys: B.keys.map((R) => R.value),
        };
      }),
      movedRows: y,
      changedCells: A.length,
      formulaChanges: k,
      metadataChanges: F,
      hiddenRows: l,
      policy: OW,
    };
  return (
    MW.set(j, {
      signature: JSON.stringify(j),
      candidate: E,
      keys: h,
      allowed: FW(e, t),
    }),
    j
  );
}
function f4e(e, r, t = {}) {
  const n = MW.get(r);
  if (!n || n.signature !== JSON.stringify(r))
    throw Error(
      "This sort preview changed or was restored from another session. Preview again.",
    );
  if (
    (BW(e, r.request.sheetId, t),
    n.allowed !== FW(e, t) || JSON.stringify(e) !== r.snapshot)
  )
    throw Error(
      "The workbook or host permissions changed after the sort preview. Preview again.",
    );
  if (n.keys !== JSON.stringify(LW(e, r.request, r.selection)))
    throw Error(
      "Calculated sort keys changed after the preview. Preview and review the new order.",
    );
  return r.movedRows ? structuredClone(n.candidate) : e;
}
const dj = (e) => ({
    column: e,
    direction: "ascending",
    type: "auto",
    customText: "",
    useCustom: false,
  }),
  h4e = (e) => {
    var t;
    try {
      const n = ct(b_(e.workbook, e.sheetId, e.range).expandedRange).firstRow;
      return (
        n === 0 ||
        !!(
          (t = e.workbook.sheets.find((o) => o.id === e.sheetId)?.tables) !=
            null && t.some((o) => ct(o.range).firstRow === n)
        )
      );
    } catch {
      return false;
    }
  };
function p4e(e) {
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    [s, d] = ne.useState(e.range),
    [u, l] = ne.useState(() => h4e(e)),
    [c, h] = ne.useState([dj(ct(e.range).firstCol)]),
    [f, p] = ne.useState(false),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState(false),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(0),
    j = e.workbook.sheets.find((X) => X.id === e.sheetId),
    z = e.allowedSheetIds ?? e.workbook.sheets.map((X) => X.id),
    V = z.includes(e.sheetId),
    D = JSON.stringify([
      e.workbook,
      e.sheetId,
      e.range,
      e.readOnly,
      [...z].sort(),
    ]),
    B = !!y && E === D && !e.readOnly;
  let _,
    C = "";
  try {
    if (!V) throw Error("The source is outside the host-approved worksheets.");
    _ = b_(e.workbook, e.sheetId, s, u);
  } catch (X) {
    C = X.message;
  }
  const T = _ ? ct(_.expandedRange) : undefined,
    N = T
      ? Array.from(
          { length: T.lastCol - T.firstCol + 1 },
          (X, J) => T.firstCol + J,
        )
      : [],
    R = () => {
      (w(null), F(""), O(0));
    },
    L = (X, J) => {
      (R(), h(c.map((q, W) => (W === X ? { ...q, ...J } : q))));
    },
    $ = (X, J) => {
      R();
      const q = [...c];
      (([q[X], q[X + J]] = [q[X + J], q[X]]), h(q));
    },
    U = () => {
      try {
        const X = d4e(
          e.workbook,
          {
            sheetId: e.sheetId,
            range: s,
            headers: u,
            keys: c.map(({ customText: J, useCustom: q, ...W }) => ({
              ...W,
              ...(q ? { customOrder: J.split(/\r?\n/) } : {}),
            })),
            expandSelection: f,
            includeHidden: m,
            caseSensitive: g,
          },
          { readOnly: e.readOnly, allowedSheetIds: e.allowedSheetIds },
        );
        (w(X), A(D), O(0), F(""));
      } catch (X) {
        (w(null), F(X.message));
      }
    },
    G = () => {
      try {
        if (!y || !B)
          throw Error("The workbook or selection changed. Preview again.");
        const X = f4e(e.workbook, y, {
          readOnly: e.readOnly,
          allowedSheetIds: e.allowedSheetIds,
        });
        (e.onWorkbookChange(X, `Sort ${j?.name}: ${y.movedRows} records moved`),
          e.onClose());
      } catch (X) {
        (w(null), F(X.message));
      }
    };
  (ne.useEffect(() => {
    var q;
    const X = document.activeElement;
    return (
      (q = o.current?.querySelector('input[data-ui-label="Source range"]')) ==
        null || q.focus(),
      () => {
        X != null && X.isConnected && X.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      var X, J, q;
      y &&
        ((J = (X = a.current) == null ? undefined : X.scrollIntoView) == null ||
          J.call(X, { block: "start" }),
        (q = a.current) == null || q.focus({ preventScroll: true }));
    }, [y]));
  const H = Math.max(1, Math.ceil((y?.rows.length ?? 0) / yd.pageSize)),
    Q = Math.min(I, H - 1);
  return i.jsx("div", {
    className: "dup-backdrop",
    onMouseDown: (X) => {
      X.target === X.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "dup-dialog sort-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      ref: o,
      onKeyDown: (X) => {
        var q, W;
        if (
          (X.stopPropagation(),
          X.key === "Escape" && (X.preventDefault(), e.onClose()),
          X.key === "Tab")
        ) {
          const K = [
            ...(o.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
            ) ?? []),
          ].filter(
            (te) =>
              !te.closest("details:not([open])") || te.tagName === "SUMMARY",
          );
          X.shiftKey && document.activeElement === K[0]
            ? (X.preventDefault(), (q = K.at(-1)) == null || q.focus())
            : !X.shiftKey &&
              document.activeElement === K.at(-1) &&
              (X.preventDefault(), (W = K[0]) == null || W.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(M, { children: "Custom sort" }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children: "Review the new record order before applying it.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close custom sort"),
              onClick: e.onClose,
              lang: t("Close custom sort"),
              "data-ui-label": "Close custom sort",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "dup-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Worksheet" }),
                i.jsx("strong", {
                  children: V ? j?.name : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": r("Source range"),
                  value: s,
                  onChange: (X) => {
                    (R(), p(false), d(X.target.value));
                  },
                  lang: t("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: u,
                  onChange: (X) => {
                    (R(), l(X.target.checked));
                  },
                }),
                i.jsx(M, { children: "My data has headers" }),
              ],
            }),
          ],
        }),
        C
          ? i.jsx("p", { role: "alert", children: C })
          : i.jsxs("div", {
              className: "sort-scope",
              children: [
                i.jsxs("strong", {
                  children: [
                    i.jsx(M, { children: "Complete records: " }),
                    _?.expandedRange,
                  ],
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      "Every populated column in these worksheet rows moves together, including fields beside the selection. Headers stay in place. Applying also clears temporary view sorting and search.",
                  }),
                }),
                _?.expansionRequired &&
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(M, { children: "Expand selection" }),
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "radio",
                            name: `${n}-expand`,
                            checked: !f,
                            onChange: () => {
                              (R(), p(false));
                            },
                          }),
                          i.jsx(M, {
                            children:
                              "Keep current selection (sorting blocked to protect records)",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "radio",
                            name: `${n}-expand`,
                            checked: f,
                            onChange: () => {
                              (R(), p(true));
                            },
                          }),
                          i.jsx(M, { children: "Expand to " }),
                          _.expandedRange,
                        ],
                      }),
                    ],
                  }),
              ],
            }),
        i.jsxs("div", {
          className: "sort-level-heading",
          children: [
            i.jsx("h3", { children: i.jsx(M, { children: "Sort levels" }) }),
            i.jsxs("button", {
              disabled: c.length >= yd.keys || c.length >= N.length,
              onClick: () => {
                (R(),
                  h([...c, dj(N.find((X) => !c.some((J) => J.column === X)))]));
              },
              children: [
                i.jsx(Plus, { size: 15 }),
                i.jsx(M, { children: "Add level" }),
              ],
            }),
          ],
        }),
        i.jsx("ol", {
          className: "sort-levels",
          children: c.map((X$1, J) =>
            i.jsxs(
              "li",
              {
                children: [
                  i.jsx("span", {
                    className: "sort-level-number",
                    children: J + 1,
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Column" }),
                      N.length > 500
                        ? i.jsx(ci, {
                            "aria-label": `Column ${J + 1}`,
                            min: T.firstCol + 1,
                            max: T.lastCol + 1,
                            value: X$1.column + 1,
                            onValueChange: (q) =>
                              L(J, { column: Number(q) - 1 }),
                          })
                        : i.jsx("select", {
                            "aria-label": `Column ${J + 1}`,
                            value: X$1.column,
                            onChange: (q) =>
                              L(J, { column: Number(q.target.value) }),
                            children: N.map((q) => {
                              return i.jsxs(
                                "option",
                                {
                                  value: q,
                                  children: [
                                    ut(q),
                                    u && T
                                      ? ` · ${String(j?.cells[`${ut(q)}${T.firstRow + 1}`]?.raw ?? "(blank header)")}`
                                      : "",
                                  ],
                                },
                                q,
                              );
                            }),
                          }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Order" }),
                      i.jsxs("select", {
                        "aria-label": `Order ${J + 1}`,
                        value: X$1.direction,
                        onChange: (q) => L(J, { direction: q.target.value }),
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
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Interpret as" }),
                      i.jsxs("select", {
                        "aria-label": `Interpretation ${J + 1}`,
                        value: X$1.type,
                        onChange: (q) =>
                          L(J, { type: q.target.value, useCustom: false }),
                        children: [
                          i.jsx("option", {
                            value: "auto",
                            children: "Automatic",
                          }),
                          i.jsx("option", {
                            value: "number",
                            children: "Number",
                          }),
                          i.jsx("option", {
                            value: "date",
                            children: "Date / time",
                          }),
                          i.jsx("option", { value: "text", children: "Text" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "sort-level-actions",
                    children: [
                      i.jsx("button", {
                        "aria-label": `Move level ${J + 1} up`,
                        disabled: J === 0,
                        onClick: () => $(J, -1),
                        children: i.jsx(ArrowUp, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": `Move level ${J + 1} down`,
                        disabled: J === c.length - 1,
                        onClick: () => $(J, 1),
                        children: i.jsx(ArrowDown, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": `Remove level ${J + 1}`,
                        disabled: c.length === 1,
                        onClick: () => {
                          (R(), h(c.filter((q, W) => W !== J)));
                        },
                        children: i.jsx(X, { size: 15 }),
                      }),
                    ],
                  }),
                  X$1.type === "text" &&
                    i.jsxs("div", {
                      className: "sort-custom",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: X$1.useCustom,
                              onChange: (q) =>
                                L(J, { useCustom: q.target.checked }),
                            }),
                            i.jsx(M, {
                              children: "Use a custom order for level ",
                            }),
                            J + 1,
                          ],
                        }),
                        X$1.useCustom &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "One value per line" }),
                              i.jsx("textarea", {
                                "aria-label": `Custom order ${J + 1}`,
                                rows: 3,
                                placeholder: `High
Medium
Low`,
                                value: X$1.customText,
                                onChange: (q) =>
                                  L(J, { customText: q.target.value }),
                              }),
                              i.jsx("small", {
                                children: i.jsx(M, {
                                  children:
                                    "Unlisted values follow the list; blanks stay last. Descending reverses the list.",
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              },
              J,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "sort-options",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: g,
                  onChange: (X) => {
                    (R(), b(X.target.checked));
                  },
                }),
                i.jsx(M, { children: "Case-sensitive text" }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: m,
                  onChange: (X) => {
                    (R(), v(X.target.checked));
                  },
                }),
                i.jsx(M, { children: "Include hidden rows" }),
              ],
            }),
          ],
        }),
        i.jsx("p", {
          className: "sort-policy",
          children: i.jsx(M, {
            children:
              "Ties keep their original row order. Blanks stay last. Automatic keeps numbers and numeric text distinct. Limit: 10,000 rows and 100,000 cells.",
          }),
        }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(M, {
                children: "Comparison and reference policy",
              }),
            }),
            Object.values(OW).map((X) => i.jsx("p", { children: X }, X)),
          ],
        }),
        e.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(M, {
              children:
                "This host is read-only. Sorting cannot change records.",
            }),
          }),
        k && i.jsx("p", { className: "dup-error", role: "alert", children: k }),
        y &&
          !B &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(M, {
              children:
                "The workbook, selection or permissions changed. Preview again.",
            }),
          }),
        y &&
          B &&
          i.jsxs("div", {
            className: "dup-preview",
            "aria-label": r("Sort preview"),
            ref: a,
            tabIndex: -1,
            lang: t("Sort preview"),
            "data-ui-label": "Sort preview",
            children: [
              i.jsxs("h3", {
                children: [
                  y.movedRows,
                  " ",
                  i.jsx(M, { children: "records move · " }),
                  y.changedCells,
                  " ",
                  i.jsx(M, { children: "cells change" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  y.rows.length,
                  " ",
                  i.jsx(M, { children: "records reviewed · " }),
                  y.hiddenRows,
                  " ",
                  i.jsx(M, { children: "hidden rows · " }),
                  y.formulaChanges.length,
                  " ",
                  i.jsx(M, { children: "formula rewrites ·" }),
                  " ",
                  y.metadataChanges.length,
                  " ",
                  i.jsx(M, { children: "metadata updates" }),
                ],
              }),
              i.jsx("div", {
                className: "dup-scroll",
                children: i.jsxs("table", {
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Source row" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Destination row" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Sort values" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Source identity" }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: y.rows
                        .slice(Q * yd.pageSize, (Q + 1) * yd.pageSize)
                        .map((X) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: X.sourceRow }),
                                i.jsx("td", { children: X.destinationRow }),
                                i.jsx("td", {
                                  children: X.keys
                                    .map((J) =>
                                      J === null || J === ""
                                        ? "(blank)"
                                        : String(J),
                                    )
                                    .join(" · "),
                                }),
                                i.jsxs("td", {
                                  children: [
                                    i.jsxs("span", {
                                      children: [
                                        X.sourceName ?? "Unchanged source",
                                        X.originalRow
                                          ? ` · row ${X.originalRow}`
                                          : "",
                                      ],
                                    }),
                                    X.sourceId &&
                                      i.jsx("small", {
                                        title: X.sourceId,
                                        children: X.sourceId,
                                      }),
                                  ],
                                }),
                              ],
                            },
                            X.sourceRow,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": r("Sort preview pages"),
                lang: t("Sort preview pages"),
                "data-ui-label": "Sort preview pages",
                children: [
                  i.jsx("button", {
                    disabled: Q === 0,
                    onClick: () => O(Q - 1),
                    children: i.jsx(M, { children: "Previous" }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(M, { children: "Page " }),
                      Q + 1,
                      " ",
                      i.jsx(M, { children: "of " }),
                      H,
                    ],
                  }),
                  i.jsx("button", {
                    disabled: Q === H - 1,
                    onClick: () => O(Q + 1),
                    children: i.jsx(M, { children: "Next" }),
                  }),
                ],
              }),
              y.formulaChanges.length > 0 &&
                i.jsxs("details", {
                  children: [
                    i.jsxs("summary", {
                      children: [
                        i.jsx(M, { children: "Review " }),
                        y.formulaChanges.length,
                        " ",
                        i.jsx(M, { children: "formula changes" }),
                      ],
                    }),
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children:
                          "First 50 changes shown; the SDK preview contains the complete list.",
                      }),
                    }),
                    y.formulaChanges.slice(0, 50).map((X, J) =>
                      i.jsxs(
                        "p",
                        {
                          children: [
                            i.jsxs("strong", {
                              children: [
                                X.sheet,
                                "!",
                                X.source,
                                " → ",
                                X.destination,
                              ],
                            }),
                            i.jsx("br", {}),
                            i.jsxs("code", {
                              children: [X.before, " → ", X.after],
                            }),
                          ],
                        },
                        J,
                      ),
                    ),
                  ],
                }),
              y.metadataChanges.length > 0 &&
                i.jsxs("details", {
                  children: [
                    i.jsx("summary", {
                      children: i.jsx(M, {
                        children: "Review metadata updates",
                      }),
                    }),
                    y.metadataChanges.map((X, J) =>
                      i.jsxs("p", { children: [X.sheet, ": ", X.field] }, J),
                    ),
                  ],
                }),
            ],
          }),
        i.jsxs("footer", {
          className: "dup-actions",
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(M, { children: "Cancel" }),
            }),
            i.jsxs("div", {
              children: [
                i.jsx("button", {
                  disabled: e.readOnly || !!C,
                  onClick: U,
                  children: i.jsx(M, { children: "Preview sort" }),
                }),
                i.jsx("button", {
                  className: "dup-primary",
                  disabled: !B || !(y != null && y.movedRows),
                  onClick: G,
                  children: i.jsx(M, { children: "Apply sort" }),
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}
const x2 = (e) =>
    e.veryHidden ? "veryHidden" : e.hidden ? "hidden" : "visible",
  fj = (e) =>
    e === "veryHidden" ? "Very hidden" : e === "hidden" ? "Hidden" : "Visible",
  bS = (e) =>
    e instanceof Error
      ? e.message
      : "The worksheet operation could not be completed.";
function m4e(e) {
  var K;
  const { text: r, language: t } = Dr(),
    [n, o] = ne.useState(e.sheetId),
    a = e.workbook.sheets.find((ae) => ae.id === n) ?? e.workbook.sheets[0],
    [s, d] = ne.useState(a.name),
    [u, l] = ne.useState(x2(a)),
    [c, h] = ne.useState(a.tabColor ?? "#107c41"),
    [f, p] = ne.useState(""),
    [m, v] = ne.useState(""),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(false),
    [E, A] = ne.useState(null),
    [k, F] = ne.useState(false),
    [I, O] = ne.useState(0),
    j = ne.useRef(null),
    z = ne.useRef(null),
    V = ne.useRef(e),
    D = ne.useRef(true);
  V.current = e;
  const B = ne.useId(),
    _ = !!((K = e.workbook.structureProtection) != null && K.enabled),
    T = ne
      .useMemo(() => jne(e.workbook), [e.workbook])
      .find((ae) => ae.sheetId === a.id),
    N = e.readOnly || _ || y,
    R = e.workbook.sheets.findIndex((ae) => ae.id === a.id),
    L = !a.hidden && e.workbook.sheets.filter((ae) => !ae.hidden).length === 1,
    $ = !N && e.workbook.sheets.length > 1 && !L;
  (ne.useEffect(() => {
    var oe;
    D.current = true;
    const ae = document.activeElement;
    return (
      (oe = z.current ?? j.current) == null || oe.focus(),
      () => {
        ((D.current = false),
          ae instanceof HTMLElement && ae.isConnected && ae.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      (o(a.id), d(a.name), l(x2(a)), h(a.tabColor ?? "#107c41"));
    }, [a.id, a.name, a.hidden, a.veryHidden, a.tabColor]));
  function U(ae) {
    (o(ae), A(null), F(false), v(""), b(""));
  }
  function G(ae, oe, de = false) {
    (v(""), b(""));
    try {
      const he = V.current;
      if (he.readOnly) throw new Error("This workbook is read-only.");
      const Se = oe(he.workbook);
      if (Se === he.workbook) return;
      (he.onWorkbookChange(Se, ae), A(null), F(false), O(0), b(ae));
      const le = de
        ? Se.sheets.find(
            (ce) => !he.workbook.sheets.some((ge) => ge.id === ce.id),
          )
        : undefined;
      (le
        ? (o(le.id), le.hidden || he.onSelectSheet(le.id))
        : Se.sheets.some((ce) => ce.id === a.id) ||
          o(Se.sheets.find((ce) => !ce.hidden).id),
        !le &&
          !Se.sheets.some((ce) => ce.id === he.sheetId && !ce.hidden) &&
          he.onSelectSheet(Se.sheets.find((ce) => !ce.hidden).id));
    } catch (he) {
      v(bS(he));
    }
  }
  function H() {
    const ae = new Set(e.workbook.sheets.map((de) => de.name.toLowerCase()));
    let oe = 1;
    for (; ae.has(`sheet${oe}`);) oe++;
    G("Add worksheet", (de) => CP(de, `Sheet${oe}`), true);
  }
  function Q() {
    (v(""), b(""), O(0));
    try {
      (A({
        workbook: e.workbook,
        sheetId: a.id,
        name: a.name,
        dependencies: AP(e.workbook, a.id),
      }),
        F(false));
    } catch (ae) {
      v(bS(ae));
    }
  }
  function X$1() {
    if (!(!E || !k)) {
      if (E.workbook !== V.current.workbook) {
        v("The workbook changed. Review current dependencies before deleting.");
        return;
      }
      G(`Delete worksheet: ${E.name}`, (ae) => Vne(ae, E.sheetId, true));
    }
  }
  async function J() {
    var de;
    const ae = V.current,
      oe = !!((de = ae.workbook.structureProtection) != null && de.enabled);
    (v(""), b(""), w(true));
    try {
      if (ae.readOnly) throw new Error("This workbook is read-only.");
      const he = await (oe ? Hne(ae.workbook, f) : Wne(ae.workbook, f));
      if (!D.current) return;
      const Se = V.current;
      if (Se.readOnly)
        throw new Error(
          "The host changed this workbook to read-only. No structure change was applied.",
        );
      if (Se.workbook !== ae.workbook)
        throw new Error(
          "The workbook changed while checking the password. Try again with the current workbook.",
        );
      (Se.onWorkbookChange(
        he,
        oe ? "Unprotect workbook structure" : "Protect workbook structure",
      ),
        p(""),
        A(null),
        b(
          oe
            ? "Workbook structure is unlocked."
            : "Workbook structure is protected.",
        ));
    } catch (he) {
      D.current && v(bS(he));
    } finally {
      D.current && w(false);
    }
  }
  const q = !!(E && E.workbook !== e.workbook),
    W = E?.dependencies.some((ae) => /unsupported/i.test(ae.kind));
  return i.jsx("div", {
    className: "sm-backdrop",
    onPointerDown: (ae) => {
      ae.target === ae.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      ref: j,
      className: "sm-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": B,
      tabIndex: -1,
      onKeyDown: (ae) => {
        var de;
        if (
          (ae.stopPropagation(),
          ae.key === "Escape" &&
            (ae.preventDefault(), E ? (A(null), F(false)) : e.onClose()),
          ae.key === "Tab")
        ) {
          const he = Array.from(
              j.current?.querySelectorAll(
                'button,input,select,textarea,[tabindex="0"]',
              ) ?? [],
            ).filter(
              (ge) =>
                !ge.matches(":disabled") &&
                !ge.closest("[hidden]") &&
                ge.tabIndex >= 0,
            ),
            Se = he[0],
            le = he.at(-1),
            ce = document.activeElement;
          he.length
            ? ae.shiftKey && (ce === Se || !he.includes(ce))
              ? (ae.preventDefault(), le == null || le.focus())
              : !ae.shiftKey &&
                (ce === le || !he.includes(ce)) &&
                (ae.preventDefault(), Se == null || Se.focus())
            : (ae.preventDefault(), (de = j.current) == null || de.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "sm-header",
          children: [
            i.jsxs("div", {
              children: [
                i.jsx(FileSpreadsheet, { size: 19 }),
                i.jsx("h2", {
                  id: B,
                  children: i.jsx(M, { children: "Manage Worksheets" }),
                }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": r("Close worksheet manager"),
              onClick: e.onClose,
              lang: t("Close worksheet manager"),
              "data-ui-label": "Close worksheet manager",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "sm-body",
          children: [
            i.jsxs("aside", {
              className: "sm-sheet-list",
              "aria-label": r("Workbook worksheets"),
              lang: t("Workbook worksheets"),
              "data-ui-label": "Workbook worksheets",
              children: [
                i.jsxs("div", {
                  className: "sm-list-heading",
                  children: [
                    i.jsxs("span", {
                      children: [
                        e.workbook.sheets.length,
                        " ",
                        i.jsx(M, { children: "worksheets" }),
                      ],
                    }),
                    i.jsx("button", {
                      type: "button",
                      "aria-label": r("Add worksheet"),
                      title: _
                        ? "Unprotect workbook structure to add worksheets"
                        : "Add worksheet",
                      disabled: N,
                      onClick: H,
                      lang: t("Add worksheet"),
                      "data-ui-label": "Add worksheet",
                      children: i.jsx(Plus, { size: 16 }),
                    }),
                  ],
                }),
                i.jsx("ul", {
                  children: e.workbook.sheets.map((ae) =>
                    i.jsx(
                      "li",
                      {
                        children: i.jsxs("button", {
                          type: "button",
                          ref: ae.id === a.id ? z : undefined,
                          "aria-label": `Manage ${ae.name}`,
                          "aria-pressed": ae.id === a.id,
                          onClick: () => U(ae.id),
                          children: [
                            i.jsx("span", {
                              className: "sm-tab-color",
                              style: { background: ae.tabColor ?? "#bcc7c0" },
                            }),
                            i.jsxs("span", {
                              className: "sm-sheet-name",
                              title: ae.name,
                              children: [
                                ae.name,
                                i.jsxs("small", {
                                  children: [
                                    fj(x2(ae)),
                                    ae.id === e.sheetId ? " · Current" : "",
                                  ],
                                }),
                              ],
                            }),
                            ae.hidden && i.jsx(EyeOff, { size: 14 }),
                          ],
                        }),
                      },
                      ae.id,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "sm-details",
              children: [
                i.jsxs("div", {
                  className: "sm-detail-title",
                  children: [
                    i.jsx("h3", { title: a.name, children: a.name }),
                    i.jsxs("button", {
                      type: "button",
                      className: "sm-view",
                      disabled: a.hidden,
                      title: a.hidden
                        ? "Make the worksheet visible before opening it"
                        : "Open this worksheet",
                      onClick: () => {
                        (e.onSelectSheet(a.id), e.onClose());
                      },
                      children: [
                        i.jsx(Eye$1, { size: 14 }),
                        i.jsx(M, { children: "Go to sheet" }),
                      ],
                    }),
                  ],
                }),
                e.readOnly &&
                  i.jsx("p", {
                    className: "sm-hint",
                    children: i.jsx(M, {
                      children:
                        "The host opened this workbook read-only. You can inspect its worksheets.",
                    }),
                  }),
                _ &&
                  i.jsxs("p", {
                    className: "sm-lock-note",
                    children: [
                      i.jsx(LockKeyhole, { size: 14 }),
                      i.jsx(M, {
                        children:
                          "Workbook structure is protected. Unlock it below to make changes.",
                      }),
                    ],
                  }),
                i.jsxs("section", {
                  className: "sm-inspection",
                  "aria-label": r("Worksheet inspection metadata"),
                  lang: t("Worksheet inspection metadata"),
                  "data-ui-label": "Worksheet inspection metadata",
                  children: [
                    i.jsx("h4", {
                      children: i.jsx(M, { children: "Worksheet inspection" }),
                    }),
                    i.jsxs("dl", {
                      children: [
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(M, { children: "Visibility" }),
                            }),
                            i.jsx("dd", { children: fj(T.visibility) }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(M, { children: "Stored cells" }),
                            }),
                            i.jsx("dd", {
                              children: T.storedCellCount.toLocaleString(),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(M, { children: "Formula cells" }),
                            }),
                            i.jsx("dd", {
                              children: T.formulaCellCount.toLocaleString(),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(M, { children: "Tab color" }),
                            }),
                            i.jsx("dd", { children: T.tabColor ?? "None" }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(M, { children: "Calculation" }),
                            }),
                            i.jsx("dd", {
                              children: i.jsx(M, {
                                children: "Included when referenced",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children:
                          "Hidden and very hidden worksheets still participate in calculation. This inspection does not evaluate formulas.",
                      }),
                    }),
                  ],
                }),
                i.jsxs("fieldset", {
                  className: "sm-edit-fields",
                  disabled: N,
                  children: [
                    i.jsx("legend", {
                      className: "sm-sr-only",
                      children: i.jsx(M, { children: "Worksheet settings" }),
                    }),
                    i.jsxs("form", {
                      className: "sm-name-form",
                      onSubmit: (ae) => {
                        (ae.preventDefault(),
                          G(`Rename worksheet: ${a.name}`, (oe) =>
                            Pne(oe, a.id, s),
                          ));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Sheet name" }),
                            i.jsx("input", {
                              "aria-label": r("Sheet name"),
                              value: s,
                              maxLength: 31,
                              onChange: (ae) => d(ae.target.value),
                              lang: t("Sheet name"),
                              "data-ui-label": "Sheet name",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "submit",
                          disabled: !s.trim() || s.trim() === a.name,
                          children: i.jsx(M, { children: "Rename sheet" }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "sm-row-actions",
                      children: [
                        i.jsxs("button", {
                          type: "button",
                          onClick: () =>
                            G(
                              `Duplicate worksheet: ${a.name}`,
                              (ae) => zne(ae, a.id),
                              true,
                            ),
                          children: [
                            i.jsx(Copy, { size: 14 }),
                            i.jsx(M, { children: "Duplicate sheet" }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "sm-order",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              "aria-label": r("Move worksheet up"),
                              title: r("Move worksheet earlier"),
                              disabled: R === 0,
                              onClick: () =>
                                G(`Move worksheet: ${a.name}`, (ae) =>
                                  xO(ae, a.id, R - 1),
                                ),
                              lang: t("Move worksheet up"),
                              "data-ui-label": "Move worksheet up",
                              children: i.jsx(ArrowUp, { size: 15 }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              "aria-label": r("Move worksheet down"),
                              title: r("Move worksheet later"),
                              disabled: R === e.workbook.sheets.length - 1,
                              onClick: () =>
                                G(`Move worksheet: ${a.name}`, (ae) =>
                                  xO(ae, a.id, R + 1),
                                ),
                              lang: t("Move worksheet down"),
                              "data-ui-label": "Move worksheet down",
                              children: i.jsx(ArrowDown, { size: 15 }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "sm-setting-row",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Visibility" }),
                            i.jsxs("select", {
                              "aria-label": r("Worksheet visibility"),
                              value: u,
                              onChange: (ae) => l(ae.target.value),
                              lang: t("Worksheet visibility"),
                              "data-ui-label": "Worksheet visibility",
                              children: [
                                i.jsx("option", {
                                  value: "visible",
                                  children: "Visible",
                                }),
                                i.jsx("option", {
                                  value: "hidden",
                                  children: "Hidden",
                                }),
                                i.jsx("option", {
                                  value: "veryHidden",
                                  children: "Very hidden",
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: u === x2(a),
                          onClick: () =>
                            G(`Change worksheet visibility: ${a.name}`, (ae) =>
                              Une(ae, a.id, u),
                            ),
                          children: i.jsx(M, { children: "Apply visibility" }),
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "sm-hint",
                      children: i.jsx(M, {
                        children:
                          "Hidden worksheets remain available here. Very hidden records a separate Excel visibility state. At least one worksheet must remain visible.",
                      }),
                    }),
                    i.jsxs("div", {
                      className: "sm-setting-row sm-color-row",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Tab color" }),
                            i.jsx("input", {
                              "aria-label": r("Worksheet tab color"),
                              type: "color",
                              value: c,
                              onChange: (ae) => h(ae.target.value),
                              lang: t("Worksheet tab color"),
                              "data-ui-label": "Worksheet tab color",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled:
                            c.toLowerCase() === a.tabColor?.toLowerCase(),
                          onClick: () =>
                            G(`Change tab color: ${a.name}`, (ae) =>
                              EO(ae, a.id, c),
                            ),
                          children: i.jsx(M, { children: "Apply tab color" }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: !a.tabColor,
                          onClick: () =>
                            G(`Remove tab color: ${a.name}`, (ae) =>
                              EO(ae, a.id),
                            ),
                          children: i.jsx(M, { children: "No color" }),
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "sm-delete-row",
                  children: [
                    i.jsxs("button", {
                      type: "button",
                      className: "sm-danger",
                      disabled: !$,
                      title: L
                        ? "Keep at least one visible worksheet"
                        : undefined,
                      onClick: Q,
                      children: [
                        i.jsx(Trash2, { size: 14 }),
                        i.jsx(M, { children: "Delete worksheet…" }),
                      ],
                    }),
                    i.jsx("small", {
                      children: i.jsx(M, {
                        children: "Review affected references before deleting.",
                      }),
                    }),
                  ],
                }),
                E &&
                  i.jsxs("section", {
                    className: "sm-deletion",
                    "aria-label": r("Delete worksheet confirmation"),
                    lang: t("Delete worksheet confirmation"),
                    "data-ui-label": "Delete worksheet confirmation",
                    children: [
                      i.jsxs("h4", {
                        children: [
                          i.jsx(M, { children: "Delete “" }),
                          E.name,
                          "”?",
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "The worksheet and its contents will be removed. Dependent formulas can become #REF!; dependent refresh definitions are removed while their result cells remain.",
                        }),
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx("strong", { children: E.dependencies.length }),
                          " ",
                          i.jsx(M, { children: "affected" }),
                          " ",
                          E.dependencies.length === 1
                            ? "reference"
                            : "references",
                        ],
                      }),
                      E.dependencies.length
                        ? i.jsx("ul", {
                            children: E.dependencies
                              .slice(I * 200, (I + 1) * 200)
                              .map((ae, oe) =>
                                i.jsxs(
                                  "li",
                                  {
                                    children: [
                                      i.jsx("span", { children: ae.label }),
                                      ae.address &&
                                        i.jsx("code", { children: ae.address }),
                                    ],
                                  },
                                  `${ae.kind}:${oe}`,
                                ),
                              ),
                          })
                        : i.jsx("p", {
                            className: "sm-hint",
                            children: i.jsx(M, {
                              children: "No dependent references were found.",
                            }),
                          }),
                      E.dependencies.length > 200 &&
                        i.jsxs("div", {
                          className: "sm-dependency-pages",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              "aria-label": r("Previous references"),
                              disabled: I === 0,
                              onClick: () => O((ae) => ae - 1),
                              lang: t("Previous references"),
                              "data-ui-label": "Previous references",
                              children: i.jsx(M, { children: "Previous" }),
                            }),
                            i.jsxs("span", {
                              "aria-live": "polite",
                              children: [
                                I * 200 + 1,
                                "–",
                                Math.min(E.dependencies.length, (I + 1) * 200),
                                " ",
                                i.jsx(M, { children: "of " }),
                                E.dependencies.length,
                                " ",
                                i.jsx(M, { children: "references" }),
                              ],
                            }),
                            i.jsx("button", {
                              type: "button",
                              "aria-label": r("Next references"),
                              disabled: (I + 1) * 200 >= E.dependencies.length,
                              onClick: () => O((ae) => ae + 1),
                              lang: t("Next references"),
                              "data-ui-label": "Next references",
                              children: i.jsx(M, { children: "Next" }),
                            }),
                          ],
                        }),
                      W &&
                        i.jsx("p", {
                          className: "sm-error",
                          children: i.jsx(M, {
                            children:
                              "Some references cannot be updated safely. Resolve the listed unsupported references first.",
                          }),
                        }),
                      q &&
                        i.jsx("p", {
                          className: "sm-error",
                          children: i.jsx(M, {
                            children:
                              "The workbook changed. Review its current dependencies before deleting.",
                          }),
                        }),
                      i.jsxs("label", {
                        className: "sm-confirm",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: k,
                            disabled: N || q || W,
                            onChange: (ae) => F(ae.target.checked),
                          }),
                          i.jsx(M, {
                            children:
                              "I understand that this removes the worksheet and may leave reference errors.",
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "sm-confirm-actions",
                        children: [
                          i.jsx("button", {
                            type: "button",
                            onClick: () => {
                              (A(null), F(false));
                            },
                            children: i.jsx(M, { children: "Keep worksheet" }),
                          }),
                          q
                            ? i.jsx("button", {
                                type: "button",
                                disabled: N,
                                onClick: Q,
                                children: i.jsx(M, {
                                  children: "Review current dependencies",
                                }),
                              })
                            : i.jsx("button", {
                                type: "button",
                                className: "sm-danger sm-delete-final",
                                disabled: !$ || !k || W,
                                onClick: X$1,
                                children: i.jsx(M, {
                                  children: "Delete worksheet",
                                }),
                              }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
          ],
        }),
        i.jsxs("section", {
          className: "sm-protection",
          "aria-label": r("Workbook structure protection"),
          lang: t("Workbook structure protection"),
          "data-ui-label": "Workbook structure protection",
          children: [
            i.jsxs("div", {
              className: "sm-protection-title",
              children: [
                _
                  ? i.jsx(LockKeyhole, { size: 16 })
                  : i.jsx(LockKeyholeOpen, { size: 16 }),
                i.jsx("strong", {
                  children: i.jsx(M, { children: "Workbook structure" }),
                }),
                i.jsx("span", { children: _ ? "Protected" : "Unlocked" }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Protect worksheet names, order, visibility, tab colors, and adding or deleting sheets.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Visibility and structure locks control editing. Access to this workbook and its hidden content is granted separately by the host application.",
              }),
            }),
            i.jsxs("form", {
              onSubmit: (ae) => {
                (ae.preventDefault(), y || J());
              },
              children: [
                i.jsxs("label", {
                  children: [
                    _ ? "Password to unlock" : "Password (optional)",
                    i.jsx("input", {
                      "aria-label": r("Structure password"),
                      type: "password",
                      autoComplete: "off",
                      value: f,
                      maxLength: 200,
                      disabled: e.readOnly || y,
                      onChange: (ae) => p(ae.target.value),
                      lang: t("Structure password"),
                      "data-ui-label": "Structure password",
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "submit",
                  disabled: e.readOnly || y,
                  children: y
                    ? "Checking…"
                    : _
                      ? "Unprotect workbook structure"
                      : "Protect workbook structure",
                }),
              ],
            }),
          ],
        }),
        (m || g) &&
          i.jsx("div", {
            className: "sm-feedback",
            children: m
              ? i.jsx("p", {
                  role: "alert",
                  className: "sm-error",
                  children: m,
                })
              : i.jsx("p", { role: "status", children: g }),
          }),
        i.jsxs("footer", {
          className: "sm-footer",
          children: [
            i.jsx("span", {
              children: i.jsx(M, {
                children: "Changes use the workbook’s undo history.",
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
const QT = {
    raw: "Value",
    literal: "Entered as text",
    style: "Formatting",
    numFmt: "Number format",
    format: "Format",
    note: "Note",
    tickmark: "Tickmark",
    createdAt: "Created",
    modifiedAt: "Modified",
    engagement: "Engagement",
    period: "Period",
    template: "Template",
    name: "Name",
    hidden: "Hidden",
    veryHidden: "Very hidden",
    tabColor: "Tab color",
    columnWidths: "Column widths",
    rowHeights: "Row heights",
    frozenRows: "Frozen rows",
    structureProtection: "Structure protection",
    protection: "Protection",
    definedNames: "Defined names",
    metadata: "Workbook properties",
    enabled: "Enabled",
    sourceRevision: "Source revision",
    version: "Version",
  },
  ox = (e) =>
    e === undefined
      ? "Not set"
      : e === null || e === ""
        ? "Blank"
        : typeof e == "boolean"
          ? e
            ? "Yes"
            : "No"
          : typeof e == "string"
            ? e.startsWith("data:image/")
              ? "Embedded picture"
              : /^\d{4}-\d{2}-\d{2}T.*Z$/.test(e) &&
                  Number.isFinite(Date.parse(e))
                ? new Date(e).toLocaleString()
                : e
            : typeof e != "object"
              ? String(e)
              : Array.isArray(e)
                ? e.length
                  ? e.map(ox).join("; ")
                  : "None"
                : Object.entries(e)
                    .filter(
                      ([r]) =>
                        ![
                          "id",
                          "sheetId",
                          "scopeSheetId",
                          "sourceWorkbookId",
                          "passwordHash",
                          "salt",
                        ].includes(r),
                    )
                    .map(([r, t]) => `${QT[r] ?? r}: ${ox(t)}`)
                    .join(" · ") || "No details",
  hj = (e) => (e ? ox(e) : "Empty cell"),
  g4e = (e) =>
    e != null && e.actor && e.timestamp
      ? `${e.actor} · ${new Date(e.timestamp).toLocaleString()}${e.basis === "operation-history" ? " · from operation history" : ""}`
      : "Author and time not recorded for this change";

export { QT, g4e, hj, m4e, ox, p4e };
