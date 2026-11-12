import i from "react/jsx-runtime";
import ne from "react";
import { hr } from "./part-01.js";
import { Mp, Rn, _t, ut, yU } from "./part-03.js";
import { Dr, M } from "./part-04.js";
import { kz } from "./part-20.js";
import { A_, EH, G2e, H2e, S_, TH, W2e, X2e, lN, q2e } from "./part-54.js";

function Z2e(e) {
  var _, C;
  const r = Dr(),
    t = ne.useId(),
    n = ne.useRef(null),
    o = ne.useRef(e);
  o.current = e;
  const a = e.workbook.sheets.filter(
      (T) =>
        T.externalConnection &&
        !T.hidden &&
        !T.veryHidden &&
        (!e.allowedSheetIds || e.allowedSheetIds.includes(T.id)),
    ),
    [s, d] = ne.useState(
      a
        .filter((T) => {
          var N;
          return !((N = T.protection) != null && N.enabled);
        })
        .map((T) => T.id),
    ),
    [u, l] = ne.useState(null),
    [c, h] = ne.useState(false),
    [f, p] = ne.useState(""),
    [m, v] = ne.useState(""),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(0),
    k = ne.useRef(null);
  ne.useEffect(() => {
    var T;
    (T = k.current) == null || T.focus();
  }, [y]);
  const F = ne.useRef(true),
    I = ne.useRef(null),
    O = JSON.stringify([e.readOnly, e.allowedSheetIds]),
    j = ne.useMemo(() => hr([e.workbook.id, e.workbook.sheets]), [e.workbook]),
    z =
      u &&
      (u.book === e.workbook || (!u.result.plan && u.contentHash === j)) &&
      u.adapter === e.adapter &&
      u.scope === O
        ? u.result
        : void 0;
  (ne.useEffect(() => {
    var N, R;
    F.current = true;
    const T = document.activeElement;
    return (
      (R = (N = n.current) == null ? void 0 : N.querySelector("button")) ==
        null || R.focus(),
      () => {
        var L;
        ((F.current = false),
          (L = I.current) == null || L.abort(),
          T != null && T.isConnected && T.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      var T;
      ((T = I.current) == null || T.abort(),
        d((N) =>
          N.filter((R) =>
            a.some((L) => {
              var $;
              return L.id === R && !(($ = L.protection) != null && $.enabled);
            }),
          ),
        ),
        l((N) =>
          N &&
          (N.book === e.workbook || (!N.result.plan && N.contentHash === j)) &&
          N.adapter === e.adapter &&
          N.scope === O
            ? N
            : null,
        ));
    }, [e.workbook, e.adapter, O]));
  function V() {
    var T;
    ((T = I.current) == null || T.abort(), e.onClose());
  }
  async function D(T = false) {
    var G, H;
    if (
      I.current ||
      !e.adapter ||
      e.readOnly ||
      !s.length ||
      (T && !(z != null && z.plan))
    )
      return;
    const N = e,
      R = e.adapter,
      L = new AbortController(),
      $ = () =>
        F.current &&
        o.current.workbook === N.workbook &&
        o.current.adapter === R &&
        JSON.stringify([o.current.readOnly, o.current.allowedSheetIds]) === O;
    ((I.current = L),
      h(true),
      v(""),
      p(""),
      w(""),
      b(
        r.text(
          T
            ? "Checking access for every source…"
            : "Preparing refresh previews…",
        ),
      ));
    const U = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      signal: L.signal,
      reliability: e.reliability,
    };
    try {
      if (T) {
        const Q = await G2e(R, z.plan, U);
        if (!$()) return;
        if (L.signal.aborted) throw new Rn("cancelled");
        const X = H2e(N.workbook, z.plan, Q, U);
        (N.onWorkbookChange(X, "Refresh all external sources"),
          l(null),
          p(
            r.text(
              "All reviewed connections were refreshed together. Undo restores the complete previous batch.",
            ),
          ));
      } else {
        const Q = await X2e(N.workbook, R, s, {
          ...U,
          onProgress: (J, q) => {
            $() &&
              b(
                r.message("Reviewed {completed} of {total} connections.", {
                  completed: J,
                  total: q,
                }),
              );
          },
        });
        if (!$()) return;
        let X = N.workbook;
        for (const J of Q.entries) {
          const q = J.error ?? Q.error;
          q && (X = lN(X, J.sheetId, J.sourceId, q, U));
        }
        (X !== N.workbook &&
          N.onWorkbookChange(
            X,
            "External source batch failed; previous data retained",
          ),
          l({
            result: Q,
            book: X,
            contentHash: hr([X.id, X.sheets]),
            adapter: R,
            scope: O,
          }),
          Q.plan ||
            v(
              r.text(
                "The batch is not ready. No source values were replaced. Resolve the failed connections, or select fewer connections and preview again.",
              ),
            ));
      }
    } catch (Q) {
      if (!$()) return;
      const X = Mp(Q);
      (v(X.message), l(null));
      let J = N.workbook;
      for (const q of s) {
        const W =
          (H =
            (G = N.workbook.sheets.find((K) => K.id === q)) == null
              ? void 0
              : G.externalConnection) == null
            ? void 0
            : H.source.id;
        W && (J = lN(J, q, W, X.code, U));
      }
      if (J !== N.workbook)
        try {
          N.onWorkbookChange(
            J,
            "External source batch failed; previous data retained",
          );
        } catch {}
    } finally {
      I.current === L && ((I.current = null), F.current && h(false));
    }
  }
  const B =
    (_ = z == null ? void 0 : z.entries.find((T) => T.sheetId === y)) == null
      ? void 0
      : _.snapshot;
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      ref: n,
      className: "lg-external-sources",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": t,
      onKeyDown: (T) => {
        var N, R;
        if (
          (T.key === "Escape" && (T.stopPropagation(), V()), T.key === "Tab")
        ) {
          const L = [
            ...T.currentTarget.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),[tabindex="0"]',
            ),
          ].filter(($) => !$.closest("[hidden]"));
          T.shiftKey && document.activeElement === L[0]
            ? (T.preventDefault(), (N = L.at(-1)) == null || N.focus())
            : !T.shiftKey &&
              document.activeElement === L.at(-1) &&
              (T.preventDefault(), (R = L[0]) == null || R.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", {
                  className: "lg-source-eyebrow",
                  children: i.jsx(M, { children: "Connections & Refresh" }),
                }),
                i.jsx("h2", {
                  id: t,
                  children: i.jsx(M, { children: "Refresh All" }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      "Compare connected worksheets, then apply all selected refreshes together.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              onClick: V,
              children: i.jsx(M, { children: "Back to connections" }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-source-body",
          children: [
            i.jsx("p", {
              className: "lg-source-help",
              children: i.jsx(M, {
                children:
                  "Only visible connections permitted by your application are listed. Protected worksheets cannot be refreshed. Each source is read once per batch; source snapshots may have different timestamps.",
              }),
            }),
            i.jsxs("fieldset", {
              className: "lg-source-batch-list",
              disabled: c || !!e.readOnly,
              children: [
                i.jsx("legend", {
                  children: i.jsx(M, { children: "Connections to refresh" }),
                }),
                a.map((T) => {
                  var N, R;
                  return i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: s.includes(T.id),
                          disabled: !!((N = T.protection) != null && N.enabled),
                          onChange: (L) => {
                            (d(($) =>
                              L.target.checked
                                ? [...$, T.id]
                                : $.filter((U) => U !== T.id),
                            ),
                              l(null),
                              v(""),
                              p(""),
                              w(""));
                          },
                        }),
                        i.jsxs("span", {
                          children: [
                            i.jsx("strong", { children: T.name }),
                            " ·",
                            " ",
                            T.externalConnection.source.label,
                            (R = T.protection) != null && R.enabled
                              ? ` · ${r.text("Protected")}`
                              : "",
                            i.jsxs("small", {
                              children: [
                                T.externalConnection.source.id,
                                " ·",
                                " ",
                                T.externalConnection.lastSuccess.rows,
                                " ",
                                r.text("records"),
                                " ·",
                                " ",
                                T.externalConnection.lastSuccess.at,
                                T.externalConnection.lastAttempt.outcome !==
                                "succeeded"
                                  ? ` · ${r.text("Stale source data")}`
                                  : "",
                              ],
                            }),
                          ],
                        }),
                      ],
                    },
                    T.id,
                  );
                }),
                !a.length &&
                  i.jsx("p", {
                    children: i.jsx(M, {
                      children: "No connected worksheets are available.",
                    }),
                  }),
              ],
            }),
            m &&
              i.jsx("p", {
                role: "alert",
                className: "lg-source-error",
                children: m,
              }),
            f && i.jsx("p", { role: "status", children: f }),
            c &&
              i.jsxs("p", {
                role: "status",
                children: [
                  g,
                  " ",
                  i.jsx("button", {
                    type: "button",
                    onClick: () => {
                      var T;
                      return (T = I.current) == null ? void 0 : T.abort();
                    },
                    children: i.jsx(M, { children: "Cancel refresh" }),
                  }),
                ],
              }),
            z &&
              i.jsx("section", {
                "aria-label": r.text("Refresh All results"),
                className: "lg-source-preview",
                children: z.entries.map((T) => {
                  const N = e.workbook.sheets.find(($) => $.id === T.sheetId),
                    R = N.externalConnection.lastSuccess,
                    L = T.error ?? z.error;
                  return i.jsxs(
                    "article",
                    {
                      className: "lg-source-history",
                      children: [
                        i.jsxs("h3", {
                          children: [
                            N.name,
                            " ·",
                            " ",
                            r.text(L ? "Not ready" : "Ready for review"),
                          ],
                        }),
                        i.jsxs("p", {
                          children: [
                            T.sourceId,
                            " · ",
                            N.externalConnection.source.kind,
                          ],
                        }),
                        L && i.jsx("p", { children: Mp(new Rn(L)).message }),
                        T.snapshot &&
                          !L &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("p", {
                                children: [
                                  r.text("Source snapshot"),
                                  ":",
                                  " ",
                                  T.snapshot.sourceTimestamp,
                                  " ·",
                                  " ",
                                  T.snapshot.version,
                                ],
                              }),
                              i.jsxs("p", {
                                children: [
                                  r.text("Rows"),
                                  ": ",
                                  R.rows,
                                  " →",
                                  " ",
                                  T.snapshot.rows.length,
                                  " · ",
                                  r.text("Change"),
                                  ":",
                                  " ",
                                  T.snapshot.rows.length - R.rows,
                                ],
                              }),
                              i.jsx("div", {
                                className: "lg-source-scroll",
                                children: i.jsxs("table", {
                                  "aria-label": r.message(
                                    "Control changes for {sheet}",
                                    { sheet: N.name },
                                  ),
                                  children: [
                                    i.jsx("thead", {
                                      children: i.jsx("tr", {
                                        children: [
                                          "Control field",
                                          "Previous total",
                                          "New total",
                                          "Change",
                                          "Numbers / blanks",
                                        ].map(($) =>
                                          i.jsx(
                                            "th",
                                            {
                                              scope: "col",
                                              children: r.text($),
                                            },
                                            $,
                                          ),
                                        ),
                                      }),
                                    }),
                                    i.jsx("tbody", {
                                      children: yU(T.snapshot).map(($) => {
                                        var G;
                                        const U =
                                          ((G = R.controls.find(
                                            (H) => H.field === $.field,
                                          )) == null
                                            ? void 0
                                            : G.sum) ?? 0;
                                        return i.jsxs(
                                          "tr",
                                          {
                                            children: [
                                              i.jsx("th", {
                                                scope: "row",
                                                children: $.field,
                                              }),
                                              i.jsx("td", { children: U }),
                                              i.jsx("td", { children: $.sum }),
                                              i.jsx("td", {
                                                children: $.sum - U,
                                              }),
                                              i.jsxs("td", {
                                                children: [
                                                  $.numeric,
                                                  " / ",
                                                  $.blanks,
                                                ],
                                              }),
                                            ],
                                          },
                                          $.field,
                                        );
                                      }),
                                    }),
                                  ],
                                }),
                              }),
                              i.jsx("button", {
                                type: "button",
                                disabled: c,
                                onClick: () => {
                                  (w(T.sheetId), A(0));
                                },
                                children: r.message("Review rows for {sheet}", {
                                  sheet: N.name,
                                }),
                              }),
                            ],
                          }),
                      ],
                    },
                    T.sheetId,
                  );
                }),
              }),
            B &&
              i.jsxs("section", {
                ref: k,
                tabIndex: -1,
                "aria-label": r.text("Batch source rows"),
                children: [
                  i.jsx("h3", {
                    children:
                      (C = e.workbook.sheets.find((T) => T.id === y)) == null
                        ? void 0
                        : C.name,
                  }),
                  i.jsx("div", {
                    className: "lg-source-scroll",
                    tabIndex: 0,
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsx("tr", {
                            children: B.source.columns.map((T) =>
                              i.jsxs(
                                "th",
                                {
                                  scope: "col",
                                  children: [T.name, " (", T.type, ")"],
                                },
                                T.name,
                              ),
                            ),
                          }),
                        }),
                        i.jsx("tbody", {
                          children: B.rows
                            .slice(E * 25, (E + 1) * 25)
                            .map((T, N) =>
                              i.jsx(
                                "tr",
                                {
                                  children: T.map((R, L) =>
                                    i.jsx(
                                      "td",
                                      {
                                        children:
                                          R === null
                                            ? r.text("(blank)")
                                            : String(R),
                                      },
                                      L,
                                    ),
                                  ),
                                },
                                N,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("nav", {
                    "aria-label": r.text("Batch row pages"),
                    children: [
                      i.jsx("button", {
                        type: "button",
                        disabled: E === 0,
                        onClick: () => A((T) => T - 1),
                        children: i.jsx(M, { children: "Previous records" }),
                      }),
                      i.jsx("span", {
                        children: r.message("Page {page} of {pages}", {
                          page: E + 1,
                          pages: Math.max(1, Math.ceil(B.rows.length / 25)),
                        }),
                      }),
                      i.jsx("button", {
                        type: "button",
                        disabled: (E + 1) * 25 >= B.rows.length,
                        onClick: () => A((T) => T + 1),
                        children: i.jsx(M, { children: "Next records" }),
                      }),
                    ],
                  }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Apply checks every source permission again. One failed connection prevents the entire batch from replacing data.",
              }),
            }),
            i.jsx("button", {
              type: "button",
              disabled: c || !!e.readOnly || !e.adapter || !s.length,
              onClick: () => void D(),
              children: i.jsx(M, { children: "Preview selected connections" }),
            }),
            i.jsx("button", {
              type: "button",
              className: "primary",
              disabled: c || !!e.readOnly || !(z != null && z.plan),
              onClick: () => void D(true),
              children: i.jsx(M, { children: "Apply all reviewed refreshes" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function K2e({ workbook: e, sheetId: r, onReview: t }) {
  const n = e.sheets.find((s) => s.id === r),
    o = n == null ? void 0 : n.externalConnection;
  if (!n || !o) return null;
  let a = true;
  try {
    a = A_(n) !== o.outputHash;
  } catch {}
  return i.jsxs("div", {
    className: "lg-source-status",
    role: "status",
    "aria-label": "External source status",
    children: [
      i.jsx("strong", {
        children:
          o.lastAttempt.outcome !== "succeeded" || a
            ? "Stale source data"
            : "Source snapshot",
      }),
      i.jsxs("span", {
        children: [
          o.source.label,
          " · ",
          o.lastSuccess.rows,
          " records · last successful refresh ",
          o.lastSuccess.at,
        ],
      }),
      i.jsx("span", {
        children: a
          ? "Local values changed; refresh cannot overwrite them."
          : o.lastAttempt.code
            ? NH(o.lastAttempt.code)
            : "Refresh to check for newer source data.",
      }),
      i.jsx("button", {
        type: "button",
        onClick: t,
        children: i.jsx(M, { children: "Connections & Refresh" }),
      }),
    ],
  });
}
function NH(e) {
  return new Rn(e).message;
}
function Y2e(e) {
  var Q, X, J, q;
  const [r, t] = ne.useState(false),
    n = Dr(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(e);
  s.current = e;
  const d = ne.useRef(true),
    u = ne.useRef(null),
    l = ne.useRef(false),
    c = ne.useRef([e.workbook, e.readOnly, JSON.stringify(e.allowedSheetIds)]),
    h =
      (Q = e.workbook.sheets.find((W) => W.id === e.sheetId)) == null
        ? void 0
        : Q.externalConnection,
    [f, p] = ne.useState(h ? e.sheetId : ""),
    [m, v] = ne.useState((h == null ? void 0 : h.source.id) ?? ""),
    [g, b] = ne.useState("External data"),
    [y, w] = ne.useState([]),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState(null),
    [V, D] = ne.useState(0),
    B = e.workbook.sheets.find((W) => W.id === f),
    _ = B == null ? void 0 : B.externalConnection,
    C =
      (j == null ? void 0 : j.book) === e.workbook &&
      j.adapter === e.adapter &&
      !e.readOnly
        ? j.plan
        : void 0,
    T = { readOnly: e.readOnly, allowedSheetIds: e.allowedSheetIds },
    N =
      !B ||
      (!B.hidden &&
        !B.veryHidden &&
        !((X = B.protection) != null && X.enabled) &&
        (!e.allowedSheetIds || e.allowedSheetIds.includes(B.id))),
    R = !!e.readOnly || !N || !e.adapter || !!E;
  (ne.useEffect(() => {
    var K, te;
    d.current = true;
    const W = document.activeElement;
    return (
      (te = (K = a.current) == null ? void 0 : K.querySelector("button")) ==
        null || te.focus(),
      () => {
        var ae;
        ((d.current = false),
          (ae = u.current) == null || ae.abort(),
          W != null && W.isConnected && W.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      var W;
      ((W = u.current) == null || W.abort(),
        (u.current = null),
        (l.current = false),
        A(""),
        w([]),
        z(null),
        e.adapter && U());
    }, [e.adapter]),
    ne.useEffect(() => {
      var K;
      const W = [e.workbook, e.readOnly, JSON.stringify(e.allowedSheetIds)];
      (W.some((te, ae) => te !== c.current[ae]) &&
        (l.current && ((K = u.current) == null || K.abort()), z(null)),
        (c.current = W));
    }, [e.workbook, e.readOnly, JSON.stringify(e.allowedSheetIds)]));
  function L() {
    var W;
    ((W = u.current) == null || W.abort(), e.onClose());
  }
  function $() {
    (z(null), D(0), F(""), O(""));
  }
  async function U() {
    const W = s.current.adapter;
    if (!W || l.current) return;
    ((l.current = true), A("Loading approved sources…"), F(""));
    const K = new AbortController();
    u.current = K;
    try {
      const te = await q2e(W, {
        signal: K.signal,
        reliability: s.current.reliability,
      });
      if (!d.current || W !== s.current.adapter) return;
      (w(te),
        v((ae) => {
          var oe;
          return ae || ((oe = te[0]) == null ? void 0 : oe.source.id) || "";
        }));
    } catch (te) {
      d.current && F(Mp(te).message);
    } finally {
      u.current === K &&
        ((u.current = null), (l.current = false), d.current && A(""));
    }
  }
  async function G(W = false) {
    if (R || l.current || !e.adapter || !m || (W && !C)) return;
    const K = s.current,
      te = e.adapter,
      ae = m,
      oe = f;
    ((l.current = true),
      A(W ? "Checking access and applying…" : "Refreshing preview…"),
      F(""),
      O(""));
    const de = new AbortController();
    u.current = de;
    try {
      if (W) {
        const he = await S_(te, ae, {
          signal: de.signal,
          reliability: e.reliability,
        });
        if (!d.current || de.signal.aborted) return;
        if (
          s.current.workbook !== K.workbook ||
          s.current.adapter !== te ||
          s.current.readOnly
        )
          throw new Rn("stale");
        const Se = W2e(s.current.workbook, C, he, {
          readOnly: s.current.readOnly,
          allowedSheetIds: s.current.allowedSheetIds,
        });
        (s.current.onWorkbookChange(Se, "Refresh external source"),
          z(null),
          O(
            "Source data refreshed. The reviewed population and control totals were applied together.",
          ));
      } else {
        const he = await TH(te, ae, {
          signal: de.signal,
          reliability: e.reliability,
        });
        if (!d.current || de.signal.aborted) return;
        if (
          s.current.workbook !== K.workbook ||
          s.current.adapter !== te ||
          s.current.readOnly
        )
          throw new Rn("stale");
        const Se = EH(
          K.workbook,
          he.snapshot,
          he.grant,
          { outputName: g, sheetId: oe || void 0 },
          T,
        );
        (z({ plan: Se, book: K.workbook, adapter: te }), D(0));
      }
    } catch (he) {
      if (!d.current) return;
      const Se = Mp(he);
      if (
        (F(Se.message),
        z(null),
        oe && s.current.adapter === te && s.current.workbook === K.workbook)
      ) {
        const le = s.current,
          ce = lN(le.workbook, oe, ae, Se.code, {
            readOnly: le.readOnly,
            allowedSheetIds: le.allowedSheetIds,
          });
        if (ce !== le.workbook)
          try {
            le.onWorkbookChange(
              ce,
              "External source refresh failed; previous data retained",
            );
          } catch {
            F(
              "The host changed while refresh failed. Previous source data remains unchanged; reopen Connections & Refresh.",
            );
          }
      }
    } finally {
      u.current === de &&
        ((u.current = null), (l.current = false), d.current && A(""));
    }
  }
  const H = e.workbook.sheets.filter(
    (W) =>
      W.externalConnection &&
      !W.hidden &&
      !W.veryHidden &&
      (!e.allowedSheetIds || e.allowedSheetIds.includes(W.id)),
  );
  return r
    ? i.jsx(Z2e, { ...e, onClose: () => t(false) })
    : i.jsx("div", {
        className: "lg-modal-backdrop",
        children: i.jsxs("section", {
          className: "lg-external-sources",
          ref: a,
          role: "dialog",
          "aria-modal": "true",
          "aria-labelledby": o,
          onKeyDown: (W) => {
            var K, te;
            if (
              (W.key === "Escape" && (W.stopPropagation(), L()),
              W.key === "Tab")
            ) {
              const ae = [
                ...W.currentTarget.querySelectorAll(
                  'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
                ),
              ];
              W.shiftKey && document.activeElement === ae[0]
                ? (W.preventDefault(), (K = ae.at(-1)) == null || K.focus())
                : !W.shiftKey &&
                  document.activeElement === ae.at(-1) &&
                  (W.preventDefault(), (te = ae[0]) == null || te.focus());
            }
          },
          children: [
            i.jsxs("header", {
              children: [
                i.jsxs("div", {
                  children: [
                    i.jsx("p", {
                      className: "lg-source-eyebrow",
                      children: i.jsx(M, { children: "Get & Transform Data" }),
                    }),
                    i.jsx("h2", {
                      id: o,
                      children: i.jsx(M, { children: "Connections & Refresh" }),
                    }),
                    i.jsx("p", {
                      children:
                        "Review an approved source before replacing its last successful population.",
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  "aria-label": n.text("Close source connections"),
                  onClick: L,
                  children: i.jsx(M, { children: "Close" }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "lg-source-body",
              children: [
                !e.adapter &&
                  i.jsx("p", {
                    className: "lg-source-message",
                    children:
                      "Your application has not connected a source service. Ask the host developer to supply approved sources. Existing snapshots remain available.",
                  }),
                e.readOnly &&
                  i.jsx("p", {
                    className: "lg-source-message",
                    children:
                      "This workbook is read-only. Source history can be inspected; refresh cannot change it.",
                  }),
                i.jsxs("div", {
                  className: "lg-source-settings",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Connection" }),
                        i.jsxs("select", {
                          "aria-label": "Connection",
                          value: f,
                          disabled: !!E,
                          onChange: (W) => {
                            var te, ae;
                            ($(), p(W.target.value));
                            const K =
                              (te = e.workbook.sheets.find(
                                (oe) => oe.id === W.target.value,
                              )) == null
                                ? void 0
                                : te.externalConnection;
                            v(
                              (K == null ? void 0 : K.source.id) ??
                                ((ae = y[0]) == null ? void 0 : ae.source.id) ??
                                "",
                            );
                          },
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "New connection",
                            }),
                            H.map((W) =>
                              i.jsxs(
                                "option",
                                {
                                  value: W.id,
                                  children: [
                                    W.name,
                                    " · ",
                                    W.externalConnection.source.label,
                                  ],
                                },
                                W.id,
                              ),
                            ),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Approved source" }),
                        i.jsxs("select", {
                          "aria-label": "Approved source",
                          disabled: !!E || !!f || !e.adapter,
                          value: m,
                          onChange: (W) => {
                            ($(), v(W.target.value));
                          },
                          children: [
                            i.jsx("option", {
                              value: "",
                              children: "Choose a source",
                            }),
                            m &&
                              !y.some((W) => W.source.id === m) &&
                              i.jsxs("option", {
                                value: m,
                                children: [
                                  (_ == null ? void 0 : _.source.label) ?? m,
                                  " · approval unavailable",
                                ],
                              }),
                            y.map((W) =>
                              i.jsxs(
                                "option",
                                {
                                  value: W.source.id,
                                  children: [
                                    W.source.label,
                                    " · ",
                                    W.source.kind,
                                  ],
                                },
                                W.source.id,
                              ),
                            ),
                          ],
                        }),
                      ],
                    }),
                    !f &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Output worksheet" }),
                          i.jsx("input", {
                            "aria-label": "Output worksheet",
                            value: g,
                            disabled: !!E,
                            onChange: (W) => {
                              ($(), b(W.target.value));
                            },
                          }),
                        ],
                      }),
                    i.jsx("button", {
                      type: "button",
                      onClick: () => void U(),
                      disabled: !e.adapter || !!E,
                      children: i.jsx(M, {
                        children: "Reload approved sources",
                      }),
                    }),
                  ],
                }),
                i.jsx("p", {
                  className: "lg-source-help",
                  children:
                    "Only the host’s approved catalog is available here. Data-source credentials, endpoints and queries stay in its backend. Refresh is manual; opening a workbook never fetches source data.",
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: !!E || !!e.readOnly || !e.adapter || !H.length,
                  onClick: () => {
                    ($(), t(true));
                  },
                  children: i.jsx(M, { children: "Refresh All" }),
                }),
                _ &&
                  i.jsxs("section", {
                    className: "lg-source-history",
                    "aria-label": "Last successful source data",
                    children: [
                      i.jsx("h3", {
                        children:
                          _.lastAttempt.outcome === "succeeded"
                            ? "Last successful snapshot"
                            : "Stale · previous successful data retained",
                      }),
                      i.jsxs("p", {
                        children: [
                          _.source.label,
                          " · ",
                          _.source.kind,
                          " · source ID",
                          " ",
                          _.source.id,
                          " · schema ",
                          _.source.schemaVersion,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          _.lastSuccess.rows,
                          " records · refreshed",
                          " ",
                          _.lastSuccess.at,
                          " · source as of",
                          " ",
                          _.lastSuccess.sourceTimestamp,
                          " · version",
                          " ",
                          _.lastSuccess.version,
                        ],
                      }),
                      i.jsx("p", {
                        children:
                          _.lastSuccess.controls
                            .map(
                              (W) =>
                                `${W.field}: ${W.sum} (${W.numeric} numbers; ${W.blanks} blanks)`,
                            )
                            .join(" · ") ||
                          "No numeric control fields declared.",
                      }),
                      _.lastAttempt.code &&
                        i.jsxs("p", {
                          children: [
                            NH(_.lastAttempt.code),
                            " · attempted",
                            " ",
                            _.lastAttempt.at,
                          ],
                        }),
                    ],
                  }),
                k &&
                  i.jsx("p", {
                    className: "lg-source-error",
                    role: "alert",
                    children: k,
                  }),
                I && i.jsx("p", { role: "status", children: I }),
                E &&
                  i.jsxs("p", {
                    role: "status",
                    children: [
                      E,
                      " ",
                      i.jsx("button", {
                        type: "button",
                        onClick: () => {
                          var W;
                          return (W = u.current) == null ? void 0 : W.abort();
                        },
                        children: i.jsx(M, { children: "Cancel refresh" }),
                      }),
                    ],
                  }),
                j &&
                  !C &&
                  i.jsx("p", {
                    className: "lg-source-message",
                    children:
                      "The workbook or access changed. Preview again before applying.",
                  }),
                C &&
                  i.jsxs("section", {
                    className: "lg-source-preview",
                    "aria-label": "External source refresh preview",
                    children: [
                      i.jsxs("h3", {
                        children: [
                          "Refresh preview · ",
                          C.snapshot.rows.length,
                          " records",
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          C.snapshot.source.label,
                          " · ",
                          C.snapshot.source.kind,
                          " ",
                          "· ",
                          C.snapshot.source.id,
                          " · as of",
                          " ",
                          C.snapshot.sourceTimestamp,
                          " · version",
                          " ",
                          C.snapshot.version,
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          "Rows: ",
                          ((J = C.previous) == null ? void 0 : J.rows) ?? 0,
                          " →",
                          " ",
                          C.snapshot.rows.length,
                          " · change",
                          " ",
                          C.snapshot.rows.length -
                            (((q = C.previous) == null ? void 0 : q.rows) ?? 0),
                        ],
                      }),
                      i.jsx("div", {
                        className: "lg-source-scroll",
                        children: i.jsxs("table", {
                          "aria-label": "Refresh control changes",
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Control field",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Previous total",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "New total",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Change",
                                  }),
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Numbers / blanks",
                                  }),
                                ],
                              }),
                            }),
                            i.jsx("tbody", {
                              children: C.controls.map((W) => {
                                var te, ae;
                                const K =
                                  ((ae =
                                    (te = C.previous) == null
                                      ? void 0
                                      : te.controls.find(
                                          (oe) => oe.field === W.field,
                                        )) == null
                                    ? void 0
                                    : ae.sum) ?? 0;
                                return i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("th", {
                                        scope: "row",
                                        children: W.field,
                                      }),
                                      i.jsx("td", { children: K }),
                                      i.jsx("td", { children: W.sum }),
                                      i.jsx("td", { children: W.sum - K }),
                                      i.jsxs("td", {
                                        children: [W.numeric, " / ", W.blanks],
                                      }),
                                    ],
                                  },
                                  W.field,
                                );
                              }),
                            }),
                          ],
                        }),
                      }),
                      !C.controls.length &&
                        i.jsx("p", {
                          children:
                            "No numeric control fields were declared by the host.",
                        }),
                      i.jsxs("p", {
                        children: [
                          "Schema:",
                          " ",
                          C.snapshot.source.columns
                            .map(
                              (W) =>
                                `${W.name} (${W.type}${W.nullable ? ", nullable" : ""})`,
                            )
                            .join(" · "),
                        ],
                      }),
                      i.jsx("div", {
                        className: "lg-source-scroll",
                        tabIndex: 0,
                        "aria-label": "Source result rows",
                        children: i.jsxs("table", {
                          children: [
                            i.jsx("thead", {
                              children: i.jsxs("tr", {
                                children: [
                                  i.jsx("th", {
                                    scope: "col",
                                    children: "Record",
                                  }),
                                  C.snapshot.source.columns.map((W) =>
                                    i.jsx(
                                      "th",
                                      { scope: "col", children: W.name },
                                      W.name,
                                    ),
                                  ),
                                ],
                              }),
                            }),
                            i.jsx("tbody", {
                              children: C.snapshot.rows
                                .slice(V * 25, (V + 1) * 25)
                                .map((W, K) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsx("th", {
                                          scope: "row",
                                          children: V * 25 + K + 1,
                                        }),
                                        W.map((te, ae) =>
                                          i.jsx(
                                            "td",
                                            {
                                              children:
                                                te === null
                                                  ? "(blank)"
                                                  : String(te),
                                            },
                                            ae,
                                          ),
                                        ),
                                      ],
                                    },
                                    K,
                                  ),
                                ),
                            }),
                          ],
                        }),
                      }),
                      i.jsxs("nav", {
                        "aria-label": "Source result pages",
                        children: [
                          i.jsx("button", {
                            type: "button",
                            disabled: V === 0,
                            onClick: () => D((W) => W - 1),
                            children: "Previous records",
                          }),
                          i.jsxs("span", {
                            children: [
                              "Page ",
                              V + 1,
                              " of",
                              " ",
                              Math.max(
                                1,
                                Math.ceil(C.snapshot.rows.length / 25),
                              ),
                              " · all records available",
                            ],
                          }),
                          i.jsx("button", {
                            type: "button",
                            disabled: (V + 1) * 25 >= C.snapshot.rows.length,
                            onClick: () => D((W) => W + 1),
                            children: "Next records",
                          }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
            i.jsxs("footer", {
              children: [
                i.jsx("p", {
                  children:
                    "Apply checks access again. Failed or cancelled refreshes retain prior data.",
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: R || !m,
                  onClick: () => void G(),
                  children: i.jsx(M, { children: "Preview refresh" }),
                }),
                i.jsx("button", {
                  type: "button",
                  className: "primary",
                  disabled: R || !C,
                  onClick: () => void G(true),
                  children: i.jsx(M, { children: "Apply reviewed refresh" }),
                }),
              ],
            }),
          ],
        }),
      });
}
const SS = new WeakMap();
let J2e = 0;
function RH(e) {
  return e ? (SS.has(e) || SS.set(e, ++J2e), SS.get(e)) : "default";
}
const Gv = (e) => !!e && typeof e.then == "function";
function TS(e, r) {
  return Gv(e) ? e.then(r) : r(e);
}
const Vj = kz(),
  Q2e = () => crypto.randomUUID(),
  NS = (e) => typeof e == "number" && Number.isFinite(e);
function Jd(e, r) {
  const t = URL.createObjectURL(e),
    n = document.createElement("a");
  ((n.href = t),
    (n.download = r),
    n.click(),
    setTimeout(() => URL.revokeObjectURL(t), 5e3));
}
function wy(e, r) {
  const t = _t(e),
    n = _t(r);
  return `${ut(Math.min(t.col, n.col))}${Math.min(t.row, n.row) + 1}:${ut(Math.max(t.col, n.col))}${Math.max(t.row, n.row) + 1}`;
}
const eEe = ne.lazy(() =>
    import("../assistant/AssistantChat.js").then((e) => ({
      default: e.AssistantChat,
    })),
  ),
  tEe = ne.lazy(() =>
    import("../assistant/AssistantChatMarkdown.js").then((e) => ({
      default: e.AssistantChatMarkdown,
    })),
  );

export { Gv, Jd, K2e, NS, Q2e, RH, TS, Vj, Y2e, eEe, tEe, wy };
