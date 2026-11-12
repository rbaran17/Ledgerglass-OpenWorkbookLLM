import i from "react/jsx-runtime";
import ne from "react";
import { fingerprintCanonicalValue } from "./01-core-schemas-and-hashing.js";
import {
  isExternalSourceError,
  ExternalSourceError,
  _t,
  columnIndexToLetters,
  computeSourceControlTotals,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { buildLocalDemoProposal } from "./20-calculation-options-and-sources.js";
import {
  snapshotSheetContent,
  EH,
  G2e,
  applySourceGrantPlans,
  authorizeSourceAccess,
  refreshSourceSnapshot,
  grantSheetSnapshotSource,
  runSourceRefreshBatch,
  normalizeExternalSourceState,
  fetchSourceCatalog,
} from "./54-workbook-save-conflicts.js";

function ExternalSourceBatchRefreshPanel(e) {
  const r = useWorkbookMessages(),
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
      a.filter((T) => !T.protection?.enabled).map((T) => T.id),
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
    k.current?.focus();
  }, [y]);
  const F = ne.useRef(true),
    I = ne.useRef(null),
    O = JSON.stringify([e.readOnly, e.allowedSheetIds]),
    j = ne.useMemo(
      () => fingerprintCanonicalValue([e.workbook.id, e.workbook.sheets]),
      [e.workbook],
    ),
    z =
      u &&
      (u.book === e.workbook || (!u.result.plan && u.contentHash === j)) &&
      u.adapter === e.adapter &&
      u.scope === O
        ? u.result
        : undefined;
  (ne.useEffect(() => {
    F.current = true;
    const T = document.activeElement;
    return (
      n.current?.querySelector("button")?.focus(),
      () => {
        ((F.current = false), I.current?.abort(), T?.isConnected && T.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      (I.current?.abort(),
        d((N) =>
          N.filter((R) => a.some((L) => L.id === R && !L.protection?.enabled)),
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
    (I.current?.abort(), e.onClose());
  }
  async function D(T = false) {
    if (I.current || !e.adapter || e.readOnly || !s.length || (T && !z?.plan))
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
        if (L.signal.aborted) throw new ExternalSourceError("cancelled");
        const X = applySourceGrantPlans(N.workbook, z.plan, Q, U);
        (N.onWorkbookChange(X, "Refresh all external sources"),
          l(null),
          p(
            r.text(
              "All reviewed connections were refreshed together. Undo restores the complete previous batch.",
            ),
          ));
      } else {
        const Q = await runSourceRefreshBatch(N.workbook, R, s, {
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
          q &&
            (X = normalizeExternalSourceState(X, J.sheetId, J.sourceId, q, U));
        }
        (X !== N.workbook &&
          N.onWorkbookChange(
            X,
            "External source batch failed; previous data retained",
          ),
          l({
            result: Q,
            book: X,
            contentHash: fingerprintCanonicalValue([X.id, X.sheets]),
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
      const X = isExternalSourceError(Q);
      (v(X.message), l(null));
      let J = N.workbook;
      for (const q of s) {
        const W = N.workbook.sheets.find((K) => K.id === q)?.externalConnection
          ?.source.id;
        W && (J = normalizeExternalSourceState(J, q, W, X.code, U));
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
  const B = z?.entries.find((T) => T.sheetId === y)?.snapshot;
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
                  children: i.jsx(LocalizedText, {
                    children: "Connections & Refresh",
                  }),
                }),
                i.jsx("h2", {
                  id: t,
                  children: i.jsx(LocalizedText, { children: "Refresh All" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Compare connected worksheets, then apply all selected refreshes together.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              onClick: V,
              children: i.jsx(LocalizedText, {
                children: "Back to connections",
              }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-source-body",
          children: [
            i.jsx("p", {
              className: "lg-source-help",
              children: i.jsx(LocalizedText, {
                children:
                  "Only visible connections permitted by your application are listed. Protected worksheets cannot be refreshed. Each source is read once per batch; source snapshots may have different timestamps.",
              }),
            }),
            i.jsxs("fieldset", {
              className: "lg-source-batch-list",
              disabled: c || !!e.readOnly,
              children: [
                i.jsx("legend", {
                  children: i.jsx(LocalizedText, {
                    children: "Connections to refresh",
                  }),
                }),
                a.map((T) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: s.includes(T.id),
                          disabled: !!T.protection?.enabled,
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
                            T.protection?.enabled
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
                  ),
                ),
                !a.length &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
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
                    onClick: () => I.current?.abort(),
                    children: i.jsx(LocalizedText, {
                      children: "Cancel refresh",
                    }),
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
                        L &&
                          i.jsx("p", {
                            children: isExternalSourceError(
                              new ExternalSourceError(L),
                            ).message,
                          }),
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
                                      children: computeSourceControlTotals(
                                        T.snapshot,
                                      ).map(($) => {
                                        const U =
                                          R.controls.find(
                                            (H) => H.field === $.field,
                                          )?.sum ?? 0;
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
                    children: e.workbook.sheets.find((T) => T.id === y)?.name,
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
                        children: i.jsx(LocalizedText, {
                          children: "Previous records",
                        }),
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
                        children: i.jsx(LocalizedText, {
                          children: "Next records",
                        }),
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
              children: i.jsx(LocalizedText, {
                children:
                  "Apply checks every source permission again. One failed connection prevents the entire batch from replacing data.",
              }),
            }),
            i.jsx("button", {
              type: "button",
              disabled: c || !!e.readOnly || !e.adapter || !s.length,
              onClick: () => void D(),
              children: i.jsx(LocalizedText, {
                children: "Preview selected connections",
              }),
            }),
            i.jsx("button", {
              type: "button",
              className: "primary",
              disabled: c || !!e.readOnly || !z?.plan,
              onClick: () => void D(true),
              children: i.jsx(LocalizedText, {
                children: "Apply all reviewed refreshes",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
function ExternalSourceStatusPanel({ workbook, sheetId, onReview }) {
  const n = workbook.sheets.find((s) => s.id === sheetId),
    o = n?.externalConnection;
  if (!n || !o) return null;
  let a = true;
  try {
    a = snapshotSheetContent(n) !== o.outputHash;
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
            ? errorMessageOf(o.lastAttempt.code)
            : "Refresh to check for newer source data.",
      }),
      i.jsx("button", {
        type: "button",
        onClick: onReview,
        children: i.jsx(LocalizedText, { children: "Connections & Refresh" }),
      }),
    ],
  });
}
function errorMessageOf(e) {
  return new ExternalSourceError(e).message;
}
function ExternalDataRefreshDialog(e) {
  const [r, t] = ne.useState(false),
    n = useWorkbookMessages(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(e);
  s.current = e;
  const d = ne.useRef(true),
    u = ne.useRef(null),
    l = ne.useRef(false),
    c = ne.useRef([e.workbook, e.readOnly, JSON.stringify(e.allowedSheetIds)]),
    h = e.workbook.sheets.find((W) => W.id === e.sheetId)?.externalConnection,
    [f, p] = ne.useState(h ? e.sheetId : ""),
    [m, v] = ne.useState(h?.source.id ?? ""),
    [g, b] = ne.useState("External data"),
    [y, w] = ne.useState([]),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState(null),
    [V, D] = ne.useState(0),
    B = e.workbook.sheets.find((W) => W.id === f),
    _ = B?.externalConnection,
    C =
      j?.book === e.workbook && j.adapter === e.adapter && !e.readOnly
        ? j.plan
        : undefined,
    T = { readOnly: e.readOnly, allowedSheetIds: e.allowedSheetIds },
    N =
      !B ||
      (!B.hidden &&
        !B.veryHidden &&
        !B.protection?.enabled &&
        (!e.allowedSheetIds || e.allowedSheetIds.includes(B.id))),
    R = !!e.readOnly || !N || !e.adapter || !!E;
  (ne.useEffect(() => {
    d.current = true;
    const W = document.activeElement;
    return (
      a.current?.querySelector("button")?.focus(),
      () => {
        ((d.current = false), u.current?.abort(), W?.isConnected && W.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      (u.current?.abort(),
        (u.current = null),
        (l.current = false),
        A(""),
        w([]),
        z(null),
        e.adapter && U());
    }, [e.adapter]),
    ne.useEffect(() => {
      const W = [e.workbook, e.readOnly, JSON.stringify(e.allowedSheetIds)];
      (W.some((te, ae) => te !== c.current[ae]) &&
        (l.current && u.current?.abort(), z(null)),
        (c.current = W));
    }, [e.workbook, e.readOnly, JSON.stringify(e.allowedSheetIds)]));
  function L() {
    (u.current?.abort(), e.onClose());
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
      const te = await fetchSourceCatalog(W, {
        signal: K.signal,
        reliability: s.current.reliability,
      });
      if (!d.current || W !== s.current.adapter) return;
      (w(te), v((ae) => ae || te[0]?.source.id || ""));
    } catch (te) {
      d.current && F(isExternalSourceError(te).message);
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
        const he = await authorizeSourceAccess(te, ae, {
          signal: de.signal,
          reliability: e.reliability,
        });
        if (!d.current || de.signal.aborted) return;
        if (
          s.current.workbook !== K.workbook ||
          s.current.adapter !== te ||
          s.current.readOnly
        )
          throw new ExternalSourceError("stale");
        const Se = grantSheetSnapshotSource(s.current.workbook, C, he, {
          readOnly: s.current.readOnly,
          allowedSheetIds: s.current.allowedSheetIds,
        });
        (s.current.onWorkbookChange(Se, "Refresh external source"),
          z(null),
          O(
            "Source data refreshed. The reviewed population and control totals were applied together.",
          ));
      } else {
        const he = await refreshSourceSnapshot(te, ae, {
          signal: de.signal,
          reliability: e.reliability,
        });
        if (!d.current || de.signal.aborted) return;
        if (
          s.current.workbook !== K.workbook ||
          s.current.adapter !== te ||
          s.current.readOnly
        )
          throw new ExternalSourceError("stale");
        const Se = EH(
          K.workbook,
          he.snapshot,
          he.grant,
          { outputName: g, sheetId: oe || undefined },
          T,
        );
        (z({ plan: Se, book: K.workbook, adapter: te }), D(0));
      }
    } catch (he) {
      if (!d.current) return;
      const Se = isExternalSourceError(he);
      if (
        (F(Se.message),
        z(null),
        oe && s.current.adapter === te && s.current.workbook === K.workbook)
      ) {
        const le = s.current,
          ce = normalizeExternalSourceState(le.workbook, oe, ae, Se.code, {
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
    ? i.jsx(ExternalSourceBatchRefreshPanel, { ...e, onClose: () => t(false) })
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
                      children: i.jsx(LocalizedText, {
                        children: "Get & Transform Data",
                      }),
                    }),
                    i.jsx("h2", {
                      id: o,
                      children: i.jsx(LocalizedText, {
                        children: "Connections & Refresh",
                      }),
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
                  children: i.jsx(LocalizedText, { children: "Close" }),
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
                        i.jsx(LocalizedText, { children: "Connection" }),
                        i.jsxs("select", {
                          "aria-label": "Connection",
                          value: f,
                          disabled: !!E,
                          onChange: (W) => {
                            ($(), p(W.target.value));
                            const K = e.workbook.sheets.find(
                              (oe) => oe.id === W.target.value,
                            )?.externalConnection;
                            v(K?.source.id ?? y[0]?.source.id ?? "");
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
                        i.jsx(LocalizedText, { children: "Approved source" }),
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
                                  _?.source.label ?? m,
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
                          i.jsx(LocalizedText, {
                            children: "Output worksheet",
                          }),
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
                      children: i.jsx(LocalizedText, {
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
                  children: i.jsx(LocalizedText, { children: "Refresh All" }),
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
                            errorMessageOf(_.lastAttempt.code),
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
                        onClick: () => u.current?.abort(),
                        children: i.jsx(LocalizedText, {
                          children: "Cancel refresh",
                        }),
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
                          C.previous?.rows ?? 0,
                          " →",
                          " ",
                          C.snapshot.rows.length,
                          " · change",
                          " ",
                          C.snapshot.rows.length - (C.previous?.rows ?? 0),
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
                                const K =
                                  C.previous?.controls.find(
                                    (oe) => oe.field === W.field,
                                  )?.sum ?? 0;
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
                  children: i.jsx(LocalizedText, {
                    children: "Preview refresh",
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  className: "primary",
                  disabled: R || !C,
                  onClick: () => void G(true),
                  children: i.jsx(LocalizedText, {
                    children: "Apply reviewed refresh",
                  }),
                }),
              ],
            }),
          ],
        }),
      });
}
const workbookInstanceIds = new WeakMap();
let workbookInstanceCounter = 0;
function workbookInstanceKey(e) {
  return e
    ? (workbookInstanceIds.has(e) ||
        workbookInstanceIds.set(e, ++workbookInstanceCounter),
      workbookInstanceIds.get(e))
    : "default";
}
const isPromiseLike = (e) => !!e && typeof e.then == "function";
function chainMaybePromise(e, r) {
  return isPromiseLike(e) ? e.then(r) : r(e);
}
const Vj = buildLocalDemoProposal();

function wy(e, r) {
  const t = _t(e),
    n = _t(r);
  return `${columnIndexToLetters(Math.min(t.col, n.col))}${Math.min(t.row, n.row) + 1}:${columnIndexToLetters(Math.max(t.col, n.col))}${Math.max(t.row, n.row) + 1}`;
}
const LazyAssistantChat = ne.lazy(() =>
    import("../assistant/AssistantChat.js").then((e) => ({
      default: e.AssistantChat,
    })),
  ),
  LazyAssistantChatMarkdown = ne.lazy(() =>
    import("../assistant/AssistantChatMarkdown.js").then((e) => ({
      default: e.AssistantChatMarkdown,
    })),
  );

export {
  isPromiseLike,
  ExternalSourceStatusPanel,
  workbookInstanceKey,
  chainMaybePromise,
  Vj,
  ExternalDataRefreshDialog,
  LazyAssistantChat,
  LazyAssistantChatMarkdown,
  wy,
};
