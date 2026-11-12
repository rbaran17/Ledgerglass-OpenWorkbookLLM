import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import { $3, Nc, bf, hr, iu, m1, ts } from "./01-core-schemas-and-hashing.js";
import {
  Aie,
  CR,
  Cf,
  ER,
  HR,
  Tx,
  Xw,
  eU,
  jn,
  kC,
  rU,
  rn,
  uh,
  v0,
  xR,
} from "./03-workbook-model.js";
import {
  A$,
  C$,
  E$,
  IC,
  Pse,
  Use,
  hI,
  k$,
  mf,
  pI,
  x$,
  y$,
} from "./04-cell-input-paste-and-fill.js";

function $se(e, r, t = {}) {
  const n = mf(e, t),
    o = E$(e, r, t),
    a = o.manifest,
    s = IC(e, n.extensions),
    d = [],
    u = bf(a);
  let l = 0;
  for (const h of pI(e))
    for (const f of h.calls.filter((p) =>
      p.name.startsWith(a.namespace + "."),
    )) {
      if ((x$(e, [h.sheetId], t), h.literal))
        throw Error(
          "Inert imported formulas must be deliberately entered as active formulas before extension execution.",
        );
      const p = C$(s, h, f, o);
      if (
        ((l += p.reduce((v, g) => {
          return v + (Array.isArray(g) ? g.length * (g[0]?.length ?? 0) : 1);
        }, 0)),
        l > Nc.inputCells || d.length >= Nc.calls)
      )
        throw Error(
          "Extension argument or call count exceeds its review limit.",
        );
      const m = $3(h.sheetId, h.address, h.formula, f.name, p, u);
      d.some((v) => v.cacheKey === m) ||
        d.push({
          sheetId: h.sheetId,
          address: h.address,
          formula: h.formula,
          cacheKey: m,
          invocation: {
            requestId: m,
            extensionId: r,
            version: a.version,
            functionName: f.name,
            arguments: p,
          },
        });
    }
  if (!d.length) throw Error("No active formula calls use this extension.");
  const c = {
    extensionId: r,
    manifest: structuredClone(a),
    requests: d,
    disclosure:
      "Only the listed argument values, function identity, extension version and request ID will be sent to the host adapter. Confirm processing separately; results require another review before changing the workbook.",
  };
  return (
    k$.set(c, {
      bookHash: hr(e),
      policyHash: n.hash(),
      planHash: hr(c),
      binding: o,
      used: false,
    }),
    c
  );
}
const rT = new WeakMap();
async function zse(e, r, t = {}, n = { approved: false }) {
  var c, h, f;
  const o = k$.get(r),
    a = mf(e, t);
  if (!o || o.used)
    throw Error("Unknown or already executed extension input review.");
  if (n.approved !== true)
    throw Error(
      "Explicit approval is required before sending extension arguments.",
    );
  if (
    t.readOnly ||
    o.bookHash !== hr(e) ||
    o.policyHash !== a.hash() ||
    o.planHash !== hr(r)
  )
    throw Error(
      "Workbook, extension or permissions changed. Review inputs again.",
    );
  const s = n.timeoutMs ?? Nc.timeoutMs;
  if (!Number.isInteger(s) || s < 1 || s > Nc.maxTimeoutMs)
    throw Error("Extension timeout must be between 1 and 60,000 milliseconds.");
  if ((c = n.signal) != null && c.aborted)
    throw Error("Extension processing cancelled.");
  o.used = true;
  const d = new AbortController(),
    u = () => d.abort();
  (h = n.signal) == null || h.addEventListener("abort", u, { once: true });
  let l;
  try {
    const p = new Promise((O, j) => {
        (d.signal.addEventListener(
          "abort",
          () => j(Error("Extension processing cancelled or timed out.")),
          { once: true },
        ),
          (l = setTimeout(() => d.abort(), s)));
      }),
      m = [];
    let v = 0;
    const g = async () => {
      for (const O of r.requests) {
        if (o.bookHash !== hr(e) || o.policyHash !== mf(e, t).hash())
          throw Error(
            "Inputs or permissions changed during extension processing.",
          );
        if (d.signal.aborted) throw Error("Extension processing cancelled.");
        let j;
        try {
          j = await o.binding.invoke(structuredClone(O.invocation), d.signal);
        } catch {
          throw Error(
            "The extension adapter failed. No workbook changes were applied.",
          );
        }
        if (d.signal.aborted) throw Error("Extension processing cancelled.");
        if (
          !j ||
          Object.keys(j).sort().join("|") !== "revision|value" ||
          typeof j.revision != "string" ||
          !j.revision.trim() ||
          j.revision.length > 200 ||
          /[\u0000-\u001f\u007f]/.test(j.revision)
        )
          throw Error("The extension returned an invalid response envelope.");
        const z = r.manifest.functions.find(
            (D) =>
              `${r.manifest.namespace}.${D.id}` === O.invocation.functionName,
          ),
          V = m1(j.value, z.result);
        if (
          ((v += Array.isArray(V)
            ? V.length * V[0].length
            : V && typeof V == "object"
              ? Object.keys(V.fields).length
              : 1),
          v > Nc.outputCells)
        )
          throw Error("Extension results exceed the total output cell budget.");
        m.push({
          key: O.cacheKey,
          sheetId: O.sheetId,
          address: O.address,
          formula: O.formula,
          functionName: O.invocation.functionName,
          arguments: structuredClone(O.invocation.arguments),
          manifestHash: bf(r.manifest),
          scopeHash: hI(e, o.binding),
          result: V,
          calculatedAt: new Date().toISOString(),
          providerRevision: j.revision,
        });
      }
      return m;
    };
    if (
      (await Promise.race([g(), p]),
      o.bookHash !== hr(e) ||
        o.policyHash !== mf(e, t).hash() ||
        o.planHash !== hr(r))
    )
      throw Error("Inputs or permissions changed during extension processing.");
    const b = rn(e),
      y = structuredClone(b.extensions);
    ((y.caches = y.caches.filter(
      (O) =>
        !m.some(
          (j) =>
            j.sheetId === O.sheetId &&
            j.address === O.address &&
            j.functionName === O.functionName,
        ),
    )),
      y.caches.push(...m),
      (b.extensions = y));
    const w = rn(b),
      E = IC(w, a.extensions),
      A = Cf(E),
      k = [],
      F = [];
    for (const O of m) {
      const j = A.getTypedValue(O.sheetId, O.address);
      if (j.type === "error")
        throw Error(
          "The resulting formula has an error or obstructed array. No results were applied.",
        );
      if (k.some((V) => V.sheetId === O.sheetId && V.address === O.address))
        continue;
      k.push({ sheetId: O.sheetId, address: O.address, value: j.value });
      const z = w.sheets.find((V) => V.id === O.sheetId).cells[O.address];
      F.push({ sheetId: O.sheetId, address: O.address, before: z, after: z });
    }
    v0(E, F);
    const I = {
      extensionId: r.extensionId,
      results: structuredClone(m),
      preview: k,
      warning:
        "These are reviewed cached results, not a live service. Changed arguments require another approved run. Provider revision and entity source labels are host assertions, not proof of authenticity.",
    };
    return (
      rT.set(I, {
        bookHash: hr(e),
        policyHash: a.hash(),
        planHash: hr(I),
        candidate: w,
      }),
      I
    );
  } finally {
    (clearTimeout(l),
      (f = n.signal) == null || f.removeEventListener("abort", u),
      d.abort());
  }
}
function Vse(e, r, t = {}) {
  const n = rT.get(r);
  if (!n) throw Error("Unknown or already applied extension results.");
  if (
    t.readOnly ||
    n.bookHash !== hr(e) ||
    n.policyHash !== mf(e, t).hash() ||
    n.planHash !== hr(r)
  )
    throw Error(
      "Workbook, extension or permissions changed. Run and review again.",
    );
  return (rT.delete(r), rn(n.candidate));
}
const Wse = [];
function Hse({ workbook: e, extensions: r, onReview: t }) {
  const n = ne.useMemo(() => {
    try {
      const o = A$(e, { extensions: r });
      return o.cells.length
        ? `${o.cells.length} custom-function cells · ${o.cells.filter((a) => a.functions.some((s) => s.state !== "reviewed-cache")).length} need review. Extension results refresh only after approval.`
        : null;
    } catch {
      return "Custom-function inventory needs review.";
    }
  }, [e, r]);
  return n
    ? i.jsxs("div", {
        className: "lg-workbook-link-status",
        role: "status",
        children: [
          i.jsx("span", { children: n }),
          i.jsx("button", { onClick: t, children: "Review custom functions" }),
        ],
      })
    : null;
}
function qse(e) {
  const r = ne.useRef(null),
    t = ne.useRef(e),
    n = ne.useRef(null),
    o = ne.useRef(0);
  t.current = e;
  const a = e.extensions ?? Wse,
    [s, d] = ne.useState(""),
    [u, l] = ne.useState(null),
    [c, h] = ne.useState(null),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState(false),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(0),
    I = ne.useMemo(() => {
      try {
        return { data: A$(e.workbook, { extensions: a }), error: "" };
      } catch (W) {
        return { data: null, error: W.message };
      }
    }, [e.workbook, a]),
    O = new Map(
      (I.data?.available ?? []).map((W) => [W.manifest.id, W.manifest]),
    );
  for (const W of I.data?.installed ?? []) O.set(W.manifest.id, W.manifest);
  const j = O.has(s) ? s : (O.keys().next().value ?? ""),
    z = O.get(j),
    V = e.workbook.extensions?.installed.find((W) => W.manifest.id === j),
    D = () => ({
      extensions: t.current.extensions,
      readOnly: t.current.readOnly,
      allowedSheetIds: t.current.allowedSheetIds,
    }),
    B = () => {
      var W;
      (o.current++,
        (W = n.current) == null || W.abort(),
        (n.current = null),
        b(false),
        h(null),
        p(null),
        l(null),
        v(false));
    };
  (ne.useEffect(() => {
    var K;
    const W = document.activeElement;
    return (
      (K = r.current) == null || K.focus(),
      () => {
        var te;
        (o.current++,
          (te = n.current) == null || te.abort(),
          W != null && W.isConnected && W.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      B();
    }, [e.workbook, a, e.readOnly, e.allowedSheetIds]));
  const _ = (W) => {
      (w(""), A(""));
      try {
        W();
      } catch (K) {
        w(K.message);
      }
    },
    C = (W) =>
      _(() => {
        (B(), l(Pse(e.workbook, j, W, D())));
      }),
    T = () =>
      _(() => {
        (B(), h($se(e.workbook, j, D())));
      }),
    N = async () => {
      if (!c || !m) return;
      const W = ++o.current,
        K = new AbortController();
      ((n.current = K), b(true), w(""));
      try {
        const te = await zse(t.current.workbook, c, D(), {
          approved: m,
          signal: K.signal,
        });
        W === o.current && (p(te), h(null), v(false));
      } catch (te) {
        W === o.current && (w(te.message), h(null));
      } finally {
        W === o.current && (b(false), (n.current = null));
      }
    },
    R = () =>
      _(() => {
        const W = t.current.workbook,
          K = u ? Use(W, u, D()) : f ? Vse(W, f, D()) : null;
        K &&
          (t.current.onWorkbookChange(
            K,
            u
              ? `${u.action} custom-function extension`
              : "Apply reviewed custom-function results",
          ),
          B(),
          A("Reviewed changes applied. Raw formulas are preserved."));
      }),
    L = I.data?.cells ?? [],
    $ = Math.max(1, Math.ceil(L.length / 10)),
    U = Math.min(k, $ - 1);
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      ref: r,
      className: "lg-reports-dialog lg-workbook-links lg-extensions",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": "Add-ins and custom functions",
      tabIndex: -1,
      onKeyDown: (W) => {
        if (
          (W.key === "Escape" && (W.stopPropagation(), B(), e.onClose()),
          W.key === "Tab")
        ) {
          const K = Array.from(
              r.current.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
              ),
            ),
            te = K[0],
            ae = K.at(-1);
          W.shiftKey &&
          (document.activeElement === te ||
            document.activeElement === r.current)
            ? (W.preventDefault(), ae == null || ae.focus())
            : !W.shiftKey &&
              (document.activeElement === ae ||
                document.activeElement === r.current) &&
              (W.preventDefault(), te == null || te.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", { children: "LEDGERGLASS · WORKBOOK EXTENSIONS" }),
                i.jsx("h2", { children: "Add-ins & custom functions" }),
              ],
            }),
            i.jsx("button", {
              "aria-label": "Close extensions",
              onClick: () => {
                (B(), e.onClose());
              },
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-reports-scroll",
          children: [
            i.jsxs("div", {
              className: "lg-links-summary",
              children: [
                i.jsx("strong", { children: "Manual, reviewed execution" }),
                i.jsxs("span", {
                  children: [
                    I.data?.installed.length ?? 0,
                    " installed extensions",
                  ],
                }),
                i.jsxs("span", { children: [L.length, " formula cells"] }),
              ],
            }),
            i.jsx("p", {
              children:
                I.data?.policy ??
                "Only host-approved adapters can process reviewed inputs.",
            }),
            i.jsx("p", {
              children:
                "Office add-in, Python and R integrations use a host adapter. This component does not run those environments or install workbook-supplied code.",
            }),
            e.readOnly &&
              i.jsx("p", {
                className: "lg-links-warning",
                children:
                  "Read-only: inspection is available; configuration and execution are disabled.",
              }),
            (y || I.error) &&
              i.jsx("p", {
                role: "alert",
                className: "lg-links-error",
                children: y || I.error,
              }),
            E && i.jsx("p", { role: "status", children: E }),
            i.jsxs("section", {
              className: "lg-links-card",
              "aria-label": "Extension manifest",
              children: [
                i.jsx("h3", { children: "1. Choose a sanctioned extension" }),
                O.size
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          children: [
                            "Extension",
                            i.jsx("select", {
                              value: j,
                              disabled: g,
                              onChange: (W) => {
                                (B(), d(W.target.value));
                              },
                              children: [...O].map(([W, K]) =>
                                i.jsxs(
                                  "option",
                                  {
                                    value: W,
                                    children: [K.label, " · ", K.version],
                                  },
                                  W,
                                ),
                              ),
                            }),
                          ],
                        }),
                        z &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("dl", {
                                children: [
                                  i.jsx("dt", {
                                    children: "Namespace / version",
                                  }),
                                  i.jsxs("dd", {
                                    children: [z.namespace, " / ", z.version],
                                  }),
                                  i.jsx("dt", { children: "Runtime" }),
                                  i.jsx("dd", { children: z.runtime }),
                                  i.jsx("dt", { children: "Permissions" }),
                                  i.jsx("dd", {
                                    children: z.permissions.join(", "),
                                  }),
                                  i.jsx("dt", { children: "Recalculation" }),
                                  i.jsx("dd", {
                                    children:
                                      "Manual · input values invalidate caches · nonvolatile",
                                  }),
                                  i.jsx("dt", { children: "State" }),
                                  i.jsx("dd", {
                                    children: V
                                      ? V.enabled
                                        ? "Enabled"
                                        : "Disabled"
                                      : "Not installed",
                                  }),
                                ],
                              }),
                              i.jsxs("details", {
                                children: [
                                  i.jsx("summary", {
                                    children: "Function signatures and schemas",
                                  }),
                                  i.jsx("pre", {
                                    children: JSON.stringify(
                                      z.functions,
                                      null,
                                      2,
                                    ),
                                  }),
                                ],
                              }),
                              i.jsxs("div", {
                                className: "lg-links-actions",
                                children: [
                                  i.jsx("button", {
                                    disabled: e.readOnly || g,
                                    onClick: () => C("install"),
                                    children: V
                                      ? "Preview install / update"
                                      : "Preview installation",
                                  }),
                                  V &&
                                    i.jsxs(i.Fragment, {
                                      children: [
                                        i.jsxs("button", {
                                          disabled: e.readOnly || g,
                                          onClick: () =>
                                            C(V.enabled ? "disable" : "enable"),
                                          children: [
                                            "Preview ",
                                            V.enabled ? "disable" : "enable",
                                          ],
                                        }),
                                        i.jsx("button", {
                                          disabled: e.readOnly || g,
                                          onClick: () => C("remove"),
                                          children: "Preview removal",
                                        }),
                                      ],
                                    }),
                                ],
                              }),
                            ],
                          }),
                      ],
                    })
                  : i.jsx("p", {
                      children:
                        "No host adapters are configured. Ask your application developer to register an approved extension. Existing formulas stay intact.",
                    }),
              ],
            }),
            u &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                "aria-label": "Extension configuration preview",
                children: [
                  i.jsxs("h3", { children: ["Review ", u.action] }),
                  i.jsxs("p", {
                    children: [
                      u.manifest.label,
                      " · version ",
                      u.manifest.version,
                    ],
                  }),
                  i.jsx("p", { children: u.warning }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: e.readOnly,
                    onClick: R,
                    children: "Apply reviewed configuration",
                  }),
                  i.jsx("button", {
                    onClick: B,
                    children: "Cancel extension preview",
                  }),
                ],
              }),
            i.jsxs("section", {
              className: "lg-links-card",
              children: [
                i.jsx("h3", {
                  children: "2. Review inputs and request results",
                }),
                i.jsx("p", {
                  children:
                    "Review every argument before allowing this host adapter to process it. Source workbooks and credentials are not included in the request.",
                }),
                i.jsx("button", {
                  disabled:
                    !(V != null && V.enabled) || e.readOnly || g || !!I.error,
                  onClick: T,
                  children: "Review function inputs",
                }),
              ],
            }),
            c &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                "aria-label": "Extension input review",
                children: [
                  i.jsxs("h3", {
                    children: ["Inputs for ", c.requests.length, " calls"],
                  }),
                  i.jsx("p", { children: c.disclosure }),
                  i.jsx("div", {
                    className: "lg-links-table",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { children: "Cell" }),
                              i.jsx("th", { children: "Function" }),
                              i.jsx("th", { children: "Exact arguments" }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: c.requests.map((W) => {
                            return i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      e.workbook.sheets.find(
                                        (te) => te.id === W.sheetId,
                                      )?.name,
                                      "!",
                                      W.address,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: W.invocation.functionName,
                                  }),
                                  i.jsx("td", {
                                    children: i.jsx("pre", {
                                      children: JSON.stringify(
                                        W.invocation.arguments,
                                        null,
                                        2,
                                      ),
                                    }),
                                  }),
                                ],
                              },
                              W.cacheKey,
                            );
                          }),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("label", {
                    className: "lg-links-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        disabled: g,
                        checked: m,
                        onChange: (W) => v(W.target.checked),
                      }),
                      "I approve sending these argument values to this host adapter.",
                    ],
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: !m || g || e.readOnly,
                    onClick: () => void N(),
                    children: "Run approved functions",
                  }),
                  i.jsx("button", {
                    onClick: B,
                    children: g
                      ? "Cancel processing"
                      : "Cancel extension preview",
                  }),
                ],
              }),
            g &&
              i.jsx("p", {
                role: "status",
                children: "Processing approved inputs…",
              }),
            f &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                "aria-label": "Extension result review",
                children: [
                  i.jsx("h3", { children: "3. Review and apply results" }),
                  i.jsx("p", { children: f.warning }),
                  i.jsx("div", {
                    className: "lg-links-table",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { children: "Cell / function" }),
                              i.jsx("th", {
                                children: "Returned value / linked fields",
                              }),
                              i.jsx("th", { children: "Provider revision" }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: f.results.map((W) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      W.address,
                                      " · ",
                                      W.functionName,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: i.jsx("pre", {
                                      children: JSON.stringify(
                                        W.result,
                                        null,
                                        2,
                                      ),
                                    }),
                                  }),
                                  i.jsx("td", { children: W.providerRevision }),
                                ],
                              },
                              W.key,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  i.jsx("h4", { children: "Calculated formula results" }),
                  i.jsx("ul", {
                    children: f.preview.map((W) =>
                      i.jsxs(
                        "li",
                        { children: [W.address, ": ", String(W.value)] },
                        W.sheetId + W.address,
                      ),
                    ),
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: e.readOnly,
                    onClick: R,
                    children: "Apply reviewed results",
                  }),
                  i.jsx("button", {
                    onClick: B,
                    children: "Cancel extension preview",
                  }),
                ],
              }),
            i.jsxs("section", {
              className: "lg-links-card",
              "aria-label": "Custom function inventory",
              children: [
                i.jsx("h3", { children: "Formula and cache inventory" }),
                L.length
                  ? i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("div", {
                          className: "lg-links-table",
                          children: i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", {
                                      children: "Cell / raw formula",
                                    }),
                                    i.jsx("th", {
                                      children: "Availability / current value",
                                    }),
                                    i.jsx("th", {
                                      children: "Retained evidence",
                                    }),
                                  ],
                                }),
                              }),
                              i.jsx("tbody", {
                                children: L.slice(U * 10, U * 10 + 10).map(
                                  (W) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              i.jsxs("strong", {
                                                children: [
                                                  W.sheetName,
                                                  "!",
                                                  W.address,
                                                ],
                                              }),
                                              i.jsx("pre", {
                                                children: W.formula,
                                              }),
                                              W.literal &&
                                                "Inert imported text",
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              W.functions.map((K) =>
                                                i.jsxs(
                                                  "div",
                                                  {
                                                    children: [
                                                      K.name,
                                                      ": ",
                                                      K.state,
                                                    ],
                                                  },
                                                  K.name,
                                                ),
                                              ),
                                              i.jsx("strong", {
                                                children: String(
                                                  W.result.value,
                                                ),
                                              }),
                                            ],
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              W.cacheLabel,
                                              W.caches.map((K) =>
                                                i.jsxs(
                                                  "details",
                                                  {
                                                    children: [
                                                      i.jsxs("summary", {
                                                        children: [
                                                          K.functionName,
                                                          " ·",
                                                          " ",
                                                          K.providerRevision,
                                                          " ·",
                                                          " ",
                                                          K.calculatedAt,
                                                        ],
                                                      }),
                                                      i.jsx("pre", {
                                                        children:
                                                          JSON.stringify(
                                                            K.result,
                                                            null,
                                                            2,
                                                          ),
                                                      }),
                                                    ],
                                                  },
                                                  K.key,
                                                ),
                                              ),
                                              W.importedCache &&
                                                i.jsxs("details", {
                                                  children: [
                                                    i.jsx("summary", {
                                                      children:
                                                        "Historical imported result — unverified",
                                                    }),
                                                    i.jsx("pre", {
                                                      children: JSON.stringify(
                                                        W.importedCache,
                                                        null,
                                                        2,
                                                      ),
                                                    }),
                                                  ],
                                                }),
                                            ],
                                          }),
                                        ],
                                      },
                                      W.sheetId + W.address,
                                    ),
                                ),
                              }),
                            ],
                          }),
                        }),
                        i.jsxs("div", {
                          className: "lg-links-actions",
                          children: [
                            i.jsx("button", {
                              disabled: U === 0,
                              onClick: () => F(U - 1),
                              children: "Previous formula page",
                            }),
                            i.jsxs("span", {
                              children: ["Page ", U + 1, " of ", $],
                            }),
                            i.jsx("button", {
                              disabled: U + 1 === $,
                              onClick: () => F(U + 1),
                              children: "Next formula page",
                            }),
                          ],
                        }),
                      ],
                    })
                  : i.jsx("p", { children: "No custom function calls found." }),
              ],
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children:
                "Native JSON retains formulas and evidence. Host execution permissions are session-only.",
            }),
            i.jsx("button", {
              onClick: () => {
                (B(), e.onClose());
              },
              children: "Done",
            }),
          ],
        }),
      ],
    }),
  });
}
function _C(e, r) {
  if (r.readOnly !== undefined && typeof r.readOnly != "boolean")
    throw Error("Read-only policy must be boolean.");
  const t = r.sources ?? [],
    n = r.allowedSheetIds ?? e.sheets.map((o) => o.id);
  if (
    !Array.isArray(t) ||
    t.length > uh.sources ||
    new Set(t.map((o) => o.target)).size !== t.length
  )
    throw Error("Approve distinct source qualifiers, at most eight workbooks.");
  for (const o of t)
    if (
      (Xw.parse(o.target),
      o.workbook.id === e.id ||
        !Array.isArray(o.allowedSheetIds) ||
        o.allowedSheetIds.some(
          (a) => !o.workbook.sheets.some((s) => s.id === a),
        ) ||
        (o.includeHidden !== undefined && typeof o.includeHidden != "boolean"))
    )
      throw Error("Invalid approved workbook identity or worksheet scope.");
  if (!Array.isArray(n) || n.some((o) => !e.sheets.some((a) => a.id === o)))
    throw Error("Invalid destination worksheet scope.");
  return {
    sources: t,
    allowed: n,
    hash: () =>
      hr([
        !!r.readOnly,
        n,
        t.map((o) => [
          o.target,
          ER(o.workbook),
          o.allowedSheetIds,
          !!o.includeHidden,
        ]),
      ]),
  };
}
function Oy(e) {
  const r = [];
  for (const t of e.sheets)
    for (const [n, o] of y$(t.cells, "external")) {
      if (typeof o.raw != "string" || !o.raw.startsWith("=")) continue;
      const a = eU(o.raw),
        s = xR(o.raw, () => "0"),
        d = CR(s, e, t.id, n, false);
      if (
        ((a.length || d?.kind === "external") &&
          r.push({
            sheetId: t.id,
            sheetName: t.name,
            address: n,
            formula: o.raw,
            literal: !!o.literal,
            references: a,
            ...(d?.kind === "external"
              ? {
                  unsupported:
                    "Unsupported external syntax. Static A1 cells and rectangles only; DDE, whole axes, external names and 3-D links remain blocked.",
                }
              : {}),
          }),
        r.length > uh.formulas)
      )
        throw Error(
          `Workbook link review is limited to ${uh.formulas} formula cells.`,
        );
    }
  return r;
}
function nT(e, r = {}) {
  const t = _C(e, r),
    n = Oy(e);
  return {
    links: [
      ...new Set(n.flatMap((s) => s.references.map((d) => d.target))),
    ].map((s) => {
      const d = t.sources.find((b) => b.target === s),
        u = e.workbookLinks?.sources.find((b) => b.target === s),
        l = n.flatMap((b) =>
          b.references
            .filter((y) => y.target === s)
            .map((y) => ({ ...y, sheetId: b.sheetId, address: b.address })),
        ),
        c =
          d &&
          l.some(
            (b) =>
              !d.workbook.sheets.some(
                (y) => y.name.toLowerCase() === b.sheet.toLowerCase(),
              ),
          ),
        h =
          d &&
          l.some((b) => {
            const y = d.workbook.sheets.find(
              (w) => w.name.toLowerCase() === b.sheet.toLowerCase(),
            );
            return (
              y &&
              (!d.allowedSheetIds.includes(y.id) ||
                ((y.hidden || y.veryHidden) && !d.includeHidden))
            );
          }),
        f = d
          ? c
            ? "missing-worksheet"
            : h
              ? "unauthorized-worksheet"
              : u
                ? u.sourceFingerprint !== ER(d.workbook)
                  ? "stale"
                  : l.some(
                        (b) =>
                          !u.sheets.find(
                            (y) =>
                              y.name.toLowerCase() === b.sheet.toLowerCase() &&
                              Tx(b.range).every((w) =>
                                Object.hasOwn(y.cells, w),
                              ),
                          ),
                      )
                    ? "incomplete-cache"
                    : "matches-approved-snapshot"
                : "unresolved"
          : "unapproved-or-missing";
      return {
        target: s,
        packageTarget: e.workbookLinks?.packageTargets?.find((b) => b.key === s)
          ?.target,
        status: f,
        references: l,
        lastResolvedRevision: u?.sourceRevision,
        lastResolvedAt: u?.resolvedAt,
        sourceName: u?.sourceName,
        sourceFingerprint: u?.sourceFingerprint,
        cached: !!u,
      };
    }),
    cells: n,
    unsupported: n.filter((s) => s.unsupported),
    cachePolicy:
      "Values are retained snapshots. Only explicit refresh reads approved workbooks. No path or URL is fetched; matching a supplied revision does not verify a live file.",
  };
}
const iT = new WeakMap();
function Xse(e, r, t, n) {
  const o = _C(e, n).sources.find((h) => h.target === t);
  if (!o)
    throw Error(
      `Source ${t} is missing or not approved. Attach the exact source and approve its worksheets.`,
    );
  const a = rn(o.workbook),
    s = new Set(o.allowedSheetIds);
  if (ts(a.datePolicy).dateSystem !== ts(e.datePolicy).dateSystem)
    throw Error(
      "Source and destination date systems differ. Convert and review the source date system before linking.",
    );
  const d = (h) => {
      const f = a.sheets.find((p) => p.id === h);
      return (
        !!f && s.has(h) && (!(f.hidden || f.veryHidden) || !!o.includeHidden)
      );
    },
    u = Cf(a, {}, { allowRead: d, allowName: () => true, step: () => {} }),
    l = [];
  let c = 0;
  for (const h of r) {
    const f = a.sheets.find(
      (m) => m.name.toLowerCase() === h.sheet.toLowerCase(),
    );
    if (!f) throw Error(`Source worksheet ${h.sheet} is missing.`);
    if (!d(f.id))
      throw Error(`Source worksheet ${h.sheet} is outside the approved scope.`);
    let p = l.find((m) => m.id === f.id);
    p || ((p = { id: f.id, name: f.name, cells: {} }), l.push(p));
    for (const m of Tx(h.range))
      if (!Object.hasOwn(p.cells, m)) {
        if (++c > uh.cells)
          throw Error("External link refresh exceeds 5,000 source cells.");
        const v = u.getTypedValue(f.id, m);
        if (
          v.type === "error" &&
          [
            "#BLOCKED!",
            "#LIMIT!",
            "#PENDING!",
            "#TIMEOUT!",
            "#CYCLE!",
          ].includes(String(v.value))
        )
          throw Error(
            "Source calculation is unavailable, unauthorized, cyclic or over its resource limit.",
          );
        p.cells[m] = v;
      }
  }
  if (u.getOperationalErrors().length)
    throw Error(
      "A source calculation attempted an unavailable or unauthorized dependency.",
    );
  return {
    target: t,
    sourceId: a.id,
    sourceName: a.name,
    sourceRevision: a.revision,
    sourceFingerprint: ER(a),
    resolvedAt: new Date().toISOString(),
    sheets: l,
  };
}
function Gse(e, r) {
  const t = e.getTypedValue(r.sheetId, r.address);
  if (t.type === "error" || e.getSpillRange(r.sheetId, r.address))
    throw Error(
      `Cannot publish ${r.sheetName}!${r.address}: resolve its error or array result first.`,
    );
  if (e.getOperationalErrors().length)
    throw Error(
      "The destination calculation includes an unresolved or unavailable dependency.",
    );
  return t;
}
function Zse(e, r, t = {}) {
  const n = _C(e, t);
  if (t.readOnly) throw Error("The workbook is read-only.");
  if (!["refresh", "replace", "break"].includes(r.action))
    throw Error("Unknown link operation.");
  if (
    (r.target !== undefined && Xw.parse(r.target),
    r.action !== "refresh" && !r.target)
  )
    throw Error("Choose a target workbook link.");
  if (
    r.action === "replace" &&
    (Xw.parse(r.replacementTarget), r.replacementTarget === r.target)
  )
    throw Error("Choose a different source qualifier.");
  const o = Oy(e),
    a = o.filter((p) =>
      r.target
        ? p.references.some((m) => m.target === r.target)
        : p.references.length || p.unsupported,
    );
  if (!a.length) throw Error("No matching workbook link cells.");
  for (const p of a) {
    if (p.unsupported) throw Error(p.unsupported);
    if (!n.allowed.includes(p.sheetId))
      throw Error("A destination worksheet is outside the host write scope.");
    if (
      (jn(e.sheets.find((m) => m.id === p.sheetId)),
      p.literal && !r.activateLiteral)
    )
      throw Error(
        "Explicitly approve activation of inert imported formulas before resolving them.",
      );
  }
  let s = rn(e);
  s.workbookLinks = structuredClone(
    s.workbookLinks ?? { schemaVersion: 1, sources: [] },
  );
  for (const p of a) {
    const m = s.sheets.find((v) => v.id === p.sheetId);
    ((m.cells[p.address] = {
      ...m.cells[p.address],
      raw:
        r.action === "replace"
          ? xR(p.formula, (v) =>
              v.target === r.target
                ? Aie(r.replacementTarget, v.sheet, v.range)
                : undefined,
            )
          : p.formula,
    }),
      delete m.cells[p.address].literal);
  }
  if (r.action !== "break") {
    const p = Oy(s)
        .filter((v) =>
          a.some((g) => g.sheetId === v.sheetId && g.address === v.address),
        )
        .flatMap((v) => v.references),
      m = [...new Set(p.map((v) => v.target))];
    for (const v of m) {
      const g = Oy(s)
          .flatMap((y) => y.references)
          .filter((y) => y.target === v),
        b = Xse(e, g, v, t);
      ((s.workbookLinks.sources = s.workbookLinks.sources.filter(
        (y) => y.target !== v,
      )),
        s.workbookLinks.sources.push(b));
    }
  } else if (!r.acknowledgeCached)
    throw Error(
      "Acknowledge that breaking links keeps the previewed cached values and removes the complete formulas.",
    );
  s = rn(s);
  const d = new Set([...n.allowed, ...rU(s).sheets.map((p) => p.id)]),
    u = Cf(s, {}, { allowRead: (p) => d.has(p), step: () => {} }),
    l = a.map((p) => {
      const m = Gse(u, p);
      return {
        sheetId: p.sheetId,
        sheetName: p.sheetName,
        address: p.address,
        before: p.formula,
        after:
          r.action === "break"
            ? m.value
            : s.sheets.find((v) => v.id === p.sheetId).cells[p.address].raw,
        result: m,
      };
    }),
    c = [];
  for (const p of l) {
    const m = s.sheets.find((g) => g.id === p.sheetId),
      v = e.sheets.find((g) => g.id === p.sheetId).cells[p.address];
    (r.action === "break" &&
      (m.cells[p.address] = {
        ...m.cells[p.address],
        raw: p.result.value,
        literal: typeof p.result.value == "string",
      }),
      c.push({
        sheetId: m.id,
        address: p.address,
        before: v,
        after: m.cells[p.address],
      }));
  }
  const h = new Set(Oy(s).flatMap((p) => p.references.map((m) => m.target)));
  ((s.workbookLinks = {
    ...s.workbookLinks,
    sources: s.workbookLinks.sources.filter((p) => h.has(p.target)),
  }),
    v0(s, c),
    (s = rn(s)));
  const f = {
    request: structuredClone(r),
    cells: l,
    sources: structuredClone(s.workbookLinks.sources),
    warning:
      r.action === "break"
        ? "Complete formulas are replaced with these scalar values, including any other references in those formulas. Undo can restore the prior workbook. Retained original files are unchanged."
        : "These are reviewed snapshots of approved workbooks, not live network connections. Future changes require another preview and refresh.",
  };
  return (
    iT.set(f, {
      bookHash: hr(e),
      contextHash: n.hash(),
      planHash: hr(f),
      candidate: s,
    }),
    f
  );
}
function Kse(e, r, t = {}) {
  const n = iT.get(r);
  if (!n) throw Error("Unknown or already applied workbook link preview.");
  if (t.readOnly) throw Error("The workbook is read-only.");
  if (
    n.bookHash !== hr(e) ||
    n.contextHash !== _C(e, t).hash() ||
    n.planHash !== hr(r)
  )
    throw Error(
      "Workbook, source, review or permissions changed. Preview links again.",
    );
  return (iT.delete(r), rn(n.candidate));
}
const S$ = new WeakMap();
function T$(e, r, t) {
  if (
    (kC(e), e.fileName !== t || e.byteLength !== r.length || e.sha256 !== iu(r))
  )
    throw Error("Source inspection does not match its original.");
  S$.set(e, { signature: JSON.stringify(e), bytes: r.slice() });
}
function DC(e) {
  const r = S$.get(e);
  if (!r || r.signature !== JSON.stringify(e))
    throw Error(
      "The source inspection changed or is untrusted. Inspect the original again.",
    );
  return r.bytes.slice();
}
const N$ = new WeakMap();
function Yse(e, r, t) {
  if (
    !e ||
    e.schemaVersion !== 1 ||
    e.fileName !== t ||
    e.fileBytes !== r.length ||
    e.sha256 !== iu(r) ||
    !HR.includes(e.format) ||
    !t.toLowerCase().endsWith("." + e.format) ||
    e.execution !== "never" ||
    !Array.isArray(e.parts) ||
    e.parts.length > 2e3
  )
    throw Error("Invalid isolated format inspection.");
  N$.set(e, { signature: JSON.stringify(e), bytes: r.slice() });
}
function oT(e) {
  const r = N$.get(e);
  if (!r || r.signature !== JSON.stringify(e))
    throw Error(
      "The format inspection changed or is untrusted. Inspect the original again.",
    );
  return r.bytes.slice();
}

export { DC, Hse, Kse, T$, Yse, Zse, nT, oT, qse };
