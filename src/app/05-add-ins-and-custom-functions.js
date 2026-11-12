import {
  formatInspectionRecords,
  recordSourceInspection,
  sourceInspectionRecords,
  validateIsolatedFormatInspection,
} from "./06c-package-and-worksheet-parsing.js";

import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import {
  fingerprintExtensionCall,
  extensionRuntimeLimits,
  hashExtensionManifest,
  fingerprintCanonicalValue,
  validateExtensionValue,
  normalizeDatePolicy,
} from "./01-schemas-date-functions-and-validation.js";
import {
  formatExternalSheetReference,
  rewriteFormulaReferences,
  evaluateWorkbookFormulas,
  encodeWorkbookFingerprint,
  parseExternalLinkRange,
  workbookQualifierSchema,
  tokenizeExternalReferences,
  assertSheetUnprotected,
  buildExternalLinkAliases,
  normalizeWorkbookInput,
  workbookLinkLimits,
  checkDataValidationBatch,
  rewriteExternalReferences,
} from "./03-workbook-model.js";
import {
  resolveExtensionCachedResult,
  evaluateExtensionFunctionCall,
  requireApprovedExtension,
  withExtensionHostScope,
  runExtensionOperation,
  applyExtensionPreview,
  resolveExtensionSheetScope,
  k$,
  validateExtensionHostPolicy,
  reviewExtensionFormulas,
  assertExtensionWriteScope,
  getFormulaScanList,
} from "./04-cell-input-paste-and-fill.js";

function buildExtensionCallReview(workbook, extensionId, t = {}) {
  const n = validateExtensionHostPolicy(workbook, t),
    binding = requireApprovedExtension(workbook, extensionId, t),
    manifest = binding.manifest,
    s = withExtensionHostScope(workbook, n.extensions),
    requests = [],
    u = hashExtensionManifest(manifest);
  let l = 0;
  for (const h of reviewExtensionFormulas(workbook))
    for (const f of h.calls.filter((call) =>
      call.name.startsWith(manifest.namespace + "."),
    )) {
      if ((assertExtensionWriteScope(workbook, [h.sheetId], t), h.literal))
        throw Error(
          "Inert imported formulas must be deliberately entered as active formulas before extension execution.",
        );
      const p = evaluateExtensionFunctionCall(s, h, f, binding);
      if (
        ((l += p.reduce(
          (v, g) => v + (Array.isArray(g) ? g.length * (g[0]?.length ?? 0) : 1),
          0,
        )),
        l > extensionRuntimeLimits.inputCells ||
          requests.length >= extensionRuntimeLimits.calls)
      )
        throw Error(
          "Extension argument or call count exceeds its review limit.",
        );
      const m = fingerprintExtensionCall(
        h.sheetId,
        h.address,
        h.formula,
        f.name,
        p,
        u,
      );
      requests.some((request) => request.cacheKey === m) ||
        requests.push({
          sheetId: h.sheetId,
          address: h.address,
          formula: h.formula,
          cacheKey: m,
          invocation: {
            requestId: m,
            extensionId: extensionId,
            version: manifest.version,
            functionName: f.name,
            arguments: p,
          },
        });
    }
  if (!requests.length)
    throw Error("No active formula calls use this extension.");
  const c = {
    extensionId: extensionId,
    manifest: structuredClone(manifest),
    requests: requests,
    disclosure:
      "Only the listed argument values, function identity, extension version and request ID will be sent to the host adapter. Confirm processing separately; results require another review before changing the workbook.",
  };
  return (
    k$.set(c, {
      bookHash: fingerprintCanonicalValue(workbook),
      policyHash: n.hash(),
      planHash: fingerprintCanonicalValue(c),
      binding: binding,
      used: false,
    }),
    c
  );
}
const extensionResultCandidates = new WeakMap();
async function runReviewedExtensionCalls(
  workbook,
  r,
  t = {},
  n = { approved: false },
) {
  const o = k$.get(r),
    a = validateExtensionHostPolicy(workbook, t);
  if (!o || o.used)
    throw Error("Unknown or already executed extension input review.");
  if (n.approved !== true)
    throw Error(
      "Explicit approval is required before sending extension arguments.",
    );
  if (
    t.readOnly ||
    o.bookHash !== fingerprintCanonicalValue(workbook) ||
    o.policyHash !== a.hash() ||
    o.planHash !== fingerprintCanonicalValue(r)
  )
    throw Error(
      "Workbook, extension or permissions changed. Review inputs again.",
    );
  const s = n.timeoutMs ?? extensionRuntimeLimits.timeoutMs;
  if (!Number.isInteger(s) || s < 1 || s > extensionRuntimeLimits.maxTimeoutMs)
    throw Error("Extension timeout must be between 1 and 60,000 milliseconds.");
  if (n.signal?.aborted) throw Error("Extension processing cancelled.");
  o.used = true;
  const d = new AbortController(),
    u = () => d.abort();
  n.signal?.addEventListener("abort", u, { once: true });
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
      for (const request of r.requests) {
        if (
          o.bookHash !== fingerprintCanonicalValue(workbook) ||
          o.policyHash !== validateExtensionHostPolicy(workbook, t).hash()
        )
          throw Error(
            "Inputs or permissions changed during extension processing.",
          );
        if (d.signal.aborted) throw Error("Extension processing cancelled.");
        let j;
        try {
          j = await o.binding.invoke(
            structuredClone(request.invocation),
            d.signal,
          );
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
              `${r.manifest.namespace}.${D.id}` ===
              request.invocation.functionName,
          ),
          V = validateExtensionValue(j.value, z.result);
        if (
          ((v += Array.isArray(V)
            ? V.length * V[0].length
            : V && typeof V == "object"
              ? Object.keys(V.fields).length
              : 1),
          v > extensionRuntimeLimits.outputCells)
        )
          throw Error("Extension results exceed the total output cell budget.");
        m.push({
          key: request.cacheKey,
          sheetId: request.sheetId,
          address: request.address,
          formula: request.formula,
          functionName: request.invocation.functionName,
          arguments: structuredClone(request.invocation.arguments),
          manifestHash: hashExtensionManifest(r.manifest),
          scopeHash: resolveExtensionSheetScope(workbook, o.binding),
          result: V,
          calculatedAt: new Date().toISOString(),
          providerRevision: j.revision,
        });
      }
      return m;
    };
    if (
      (await Promise.race([g(), p]),
      o.bookHash !== fingerprintCanonicalValue(workbook) ||
        o.policyHash !== validateExtensionHostPolicy(workbook, t).hash() ||
        o.planHash !== fingerprintCanonicalValue(r))
    )
      throw Error("Inputs or permissions changed during extension processing.");
    const workbookInput = normalizeWorkbookInput(workbook),
      y = structuredClone(workbookInput.extensions);
    ((y.caches = y.caches.filter(
      (cach) =>
        !m.some(
          (j) =>
            j.sheetId === cach.sheetId &&
            j.address === cach.address &&
            j.functionName === cach.functionName,
        ),
    )),
      y.caches.push(...m),
      (workbookInput.extensions = y));
    const w = normalizeWorkbookInput(workbookInput),
      E = withExtensionHostScope(w, a.extensions),
      A = evaluateWorkbookFormulas(E),
      preview = [],
      F = [];
    for (const O of m) {
      const j = A.getTypedValue(O.sheetId, O.address);
      if (j.type === "error")
        throw Error(
          "The resulting formula has an error or obstructed array. No results were applied.",
        );
      if (
        preview.some((V) => V.sheetId === O.sheetId && V.address === O.address)
      )
        continue;
      preview.push({ sheetId: O.sheetId, address: O.address, value: j.value });
      const z = w.sheets.find((sheet) => sheet.id === O.sheetId).cells[
        O.address
      ];
      F.push({ sheetId: O.sheetId, address: O.address, before: z, after: z });
    }
    checkDataValidationBatch(E, F);
    const I = {
      extensionId: r.extensionId,
      results: structuredClone(m),
      preview: preview,
      warning:
        "These are reviewed cached results, not a live service. Changed arguments require another approved run. Provider revision and entity source labels are host assertions, not proof of authenticity.",
    };
    return (
      extensionResultCandidates.set(I, {
        bookHash: fingerprintCanonicalValue(workbook),
        policyHash: a.hash(),
        planHash: fingerprintCanonicalValue(I),
        candidate: w,
      }),
      I
    );
  } finally {
    (clearTimeout(l), n.signal?.removeEventListener("abort", u), d.abort());
  }
}
function applyReviewedExtensionResults(workbook, r, t = {}) {
  const n = extensionResultCandidates.get(r);
  if (!n) throw Error("Unknown or already applied extension results.");
  if (
    t.readOnly ||
    n.bookHash !== fingerprintCanonicalValue(workbook) ||
    n.policyHash !== validateExtensionHostPolicy(workbook, t).hash() ||
    n.planHash !== fingerprintCanonicalValue(r)
  )
    throw Error(
      "Workbook, extension or permissions changed. Run and review again.",
    );
  return (
    extensionResultCandidates.delete(r),
    normalizeWorkbookInput(n.candidate)
  );
}
const emptyExtensionList = [];
function CustomFunctionReviewBanner({ workbook, extensions, onReview }) {
  const n = ne.useMemo(() => {
    try {
      const extensionCachedResult = resolveExtensionCachedResult(workbook, {
        extensions,
      });
      return extensionCachedResult.cells.length
        ? `${extensionCachedResult.cells.length} custom-function cells · ${extensionCachedResult.cells.filter((cell) => cell.functions.some((s) => s.state !== "reviewed-cache")).length} need review. Extension results refresh only after approval.`
        : null;
    } catch {
      return "Custom-function inventory needs review.";
    }
  }, [workbook, extensions]);
  return n
    ? i.jsxs("div", {
        className: "lg-workbook-link-status",
        role: "status",
        children: [
          i.jsx("span", { children: n }),
          i.jsx("button", {
            onClick: onReview,
            children: "Review custom functions",
          }),
        ],
      })
    : null;
}
function AddInsAndCustomFunctionsDialog(e) {
  const r = ne.useRef(null),
    t = ne.useRef(e),
    n = ne.useRef(null),
    o = ne.useRef(0);
  t.current = e;
  const extensions = e.extensions ?? emptyExtensionList,
    [s, d] = ne.useState(""),
    [u, l] = ne.useState(null),
    [c, h] = ne.useState(null),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(false),
    [disabled, b] = ne.useState(false),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(0),
    I = ne.useMemo(() => {
      try {
        return {
          data: resolveExtensionCachedResult(e.workbook, {
            extensions: extensions,
          }),
          error: "",
        };
      } catch (error) {
        return { data: null, error: error.message };
      }
    }, [e.workbook, extensions]),
    O = new Map(
      (I.data?.available ?? []).map((W) => [W.manifest.id, W.manifest]),
    );
  for (const W of I.data?.installed ?? []) O.set(W.manifest.id, W.manifest);
  const extensionId = O.has(s) ? s : (O.keys().next().value ?? ""),
    z = O.get(extensionId),
    V = e.workbook.extensions?.installed.find(
      (W) => W.manifest.id === extensionId,
    ),
    D = () => ({
      extensions: t.current.extensions,
      readOnly: t.current.readOnly,
      allowedSheetIds: t.current.allowedSheetIds,
    }),
    onClick = () => {
      (o.current++,
        n.current?.abort(),
        (n.current = null),
        b(false),
        h(null),
        p(null),
        l(null),
        v(false));
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      r.current?.focus(),
      () => {
        (o.current++,
          n.current?.abort(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      onClick();
    }, [e.workbook, extensions, e.readOnly, e.allowedSheetIds]));
  const _ = (W) => {
      (w(""), A(""));
      try {
        W();
      } catch (error) {
        w(error.message);
      }
    },
    C = (action) =>
      _(() => {
        (onClick(),
          l(runExtensionOperation(e.workbook, extensionId, action, D())));
      }),
    T = () =>
      _(() => {
        (onClick(), h(buildExtensionCallReview(e.workbook, extensionId, D())));
      }),
    N = async () => {
      if (!c || !m) return;
      const W = ++o.current,
        K = new AbortController();
      ((n.current = K), b(true), w(""));
      try {
        const te = await runReviewedExtensionCalls(t.current.workbook, c, D(), {
          approved: m,
          signal: K.signal,
        });
        W === o.current && (p(te), h(null), v(false));
      } catch (error) {
        W === o.current && (w(error.message), h(null));
      } finally {
        W === o.current && (b(false), (n.current = null));
      }
    },
    R = () =>
      _(() => {
        const workbook = t.current.workbook,
          K = u
            ? applyExtensionPreview(workbook, u, D())
            : f
              ? applyReviewedExtensionResults(workbook, f, D())
              : null;
        K &&
          (t.current.onWorkbookChange(
            K,
            u
              ? `${u.action} custom-function extension`
              : "Apply reviewed custom-function results",
          ),
          onClick(),
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
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.stopPropagation(), onClick(), e.onClose()),
          event.key === "Tab")
        ) {
          const K = Array.from(
              r.current.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
              ),
            ),
            te = K[0],
            ae = K.at(-1);
          event.shiftKey &&
          (document.activeElement === te ||
            document.activeElement === r.current)
            ? (event.preventDefault(), ae == null || ae.focus())
            : !event.shiftKey &&
              (document.activeElement === ae ||
                document.activeElement === r.current) &&
              (event.preventDefault(), te == null || te.focus());
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
                (onClick(), e.onClose());
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
                              value: extensionId,
                              disabled: disabled,
                              onChange: (event) => {
                                (onClick(), d(event.target.value));
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
                                    disabled: e.readOnly || disabled,
                                    onClick: () => C("install"),
                                    children: V
                                      ? "Preview install / update"
                                      : "Preview installation",
                                  }),
                                  V &&
                                    i.jsxs(i.Fragment, {
                                      children: [
                                        i.jsxs("button", {
                                          disabled: e.readOnly || disabled,
                                          onClick: () =>
                                            C(V.enabled ? "disable" : "enable"),
                                          children: [
                                            "Preview ",
                                            V.enabled ? "disable" : "enable",
                                          ],
                                        }),
                                        i.jsx("button", {
                                          disabled: e.readOnly || disabled,
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
                    onClick: onClick,
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
                  disabled: !V?.enabled || e.readOnly || disabled || !!I.error,
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
                          children: c.requests.map((request) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      e.workbook.sheets.find(
                                        (sheet) => sheet.id === request.sheetId,
                                      )?.name,
                                      "!",
                                      request.address,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: request.invocation.functionName,
                                  }),
                                  i.jsx("td", {
                                    children: i.jsx("pre", {
                                      children: JSON.stringify(
                                        request.invocation.arguments,
                                        null,
                                        2,
                                      ),
                                    }),
                                  }),
                                ],
                              },
                              request.cacheKey,
                            ),
                          ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("label", {
                    className: "lg-links-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        disabled: disabled,
                        checked: m,
                        onChange: (event) => v(event.target.checked),
                      }),
                      "I approve sending these argument values to this host adapter.",
                    ],
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: !m || disabled || e.readOnly,
                    onClick: () => void N(),
                    children: "Run approved functions",
                  }),
                  i.jsx("button", {
                    onClick: onClick,
                    children: disabled
                      ? "Cancel processing"
                      : "Cancel extension preview",
                  }),
                ],
              }),
            disabled &&
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
                          children: f.results.map((result) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [
                                      result.address,
                                      " · ",
                                      result.functionName,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: i.jsx("pre", {
                                      children: JSON.stringify(
                                        result.result,
                                        null,
                                        2,
                                      ),
                                    }),
                                  }),
                                  i.jsx("td", {
                                    children: result.providerRevision,
                                  }),
                                ],
                              },
                              result.key,
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
                    onClick: onClick,
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
                                              W.caches.map((cach) =>
                                                i.jsxs(
                                                  "details",
                                                  {
                                                    children: [
                                                      i.jsxs("summary", {
                                                        children: [
                                                          cach.functionName,
                                                          " ·",
                                                          " ",
                                                          cach.providerRevision,
                                                          " ·",
                                                          " ",
                                                          cach.calculatedAt,
                                                        ],
                                                      }),
                                                      i.jsx("pre", {
                                                        children:
                                                          JSON.stringify(
                                                            cach.result,
                                                            null,
                                                            2,
                                                          ),
                                                      }),
                                                    ],
                                                  },
                                                  cach.key,
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
                (onClick(), e.onClose());
              },
              children: "Done",
            }),
          ],
        }),
      ],
    }),
  });
}
function buildWorkbookLinkPolicy(workbook, r) {
  if (r.readOnly !== undefined && typeof r.readOnly != "boolean")
    throw Error("Read-only policy must be boolean.");
  const sources = r.sources ?? [],
    allowed = r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id);
  if (
    !Array.isArray(sources) ||
    sources.length > workbookLinkLimits.sources ||
    new Set(sources.map((event) => event.target)).size !== sources.length
  )
    throw Error("Approve distinct source qualifiers, at most eight workbooks.");
  for (const source of sources)
    if (
      (workbookQualifierSchema.parse(source.target),
      source.workbook.id === workbook.id ||
        !Array.isArray(source.allowedSheetIds) ||
        source.allowedSheetIds.some(
          (allowedSheetId) =>
            !source.workbook.sheets.some(
              (sheet) => sheet.id === allowedSheetId,
            ),
        ) ||
        (source.includeHidden !== undefined &&
          typeof source.includeHidden != "boolean"))
    )
      throw Error("Invalid approved workbook identity or worksheet scope.");
  if (
    !Array.isArray(allowed) ||
    allowed.some((o) => !workbook.sheets.some((sheet) => sheet.id === o))
  )
    throw Error("Invalid destination worksheet scope.");
  return {
    sources: sources,
    allowed: allowed,
    hash: () =>
      fingerprintCanonicalValue([
        !!r.readOnly,
        allowed,
        sources.map((source) => [
          source.target,
          encodeWorkbookFingerprint(source.workbook),
          source.allowedSheetIds,
          !!source.includeHidden,
        ]),
      ]),
  };
}
function collectWorkbookLinkFormulas(workbookInput) {
  const r = [];
  for (const sheet of workbookInput.sheets)
    for (const [address, cell] of getFormulaScanList(sheet.cells, "external")) {
      if (typeof cell.raw != "string" || !cell.raw.startsWith("=")) continue;
      const references = tokenizeExternalReferences(cell.raw),
        s = rewriteExternalReferences(cell.raw, () => "0"),
        d = rewriteFormulaReferences(
          s,
          workbookInput,
          sheet.id,
          address,
          false,
        );
      if (
        ((references.length || d?.kind === "external") &&
          r.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            address: address,
            formula: cell.raw,
            literal: !!cell.literal,
            references: references,
            ...(d?.kind === "external"
              ? {
                  unsupported:
                    "Unsupported external syntax. Static A1 cells and rectangles only; DDE, whole axes, external names and 3-D links remain blocked.",
                }
              : {}),
          }),
        r.length > workbookLinkLimits.formulas)
      )
        throw Error(
          `Workbook link review is limited to ${workbookLinkLimits.formulas} formula cells.`,
        );
    }
  return r;
}
function buildWorkbookLinkReview(workbook, r = {}) {
  const workbookLinkPolicy = buildWorkbookLinkPolicy(workbook, r),
    workbookLinkFormulas = collectWorkbookLinkFormulas(workbook);
  return {
    links: [
      ...new Set(
        workbookLinkFormulas.flatMap((workbookLinkFormula) =>
          workbookLinkFormula.references.map((reference) => reference.target),
        ),
      ),
    ].map((target) => {
      const d = workbookLinkPolicy.sources.find(
          (source) => source.target === target,
        ),
        u = workbook.workbookLinks?.sources.find(
          (source) => source.target === target,
        ),
        references = workbookLinkFormulas.flatMap((workbookLinkFormula) =>
          workbookLinkFormula.references
            .filter((reference) => reference.target === target)
            .map((y) => ({
              ...y,
              sheetId: workbookLinkFormula.sheetId,
              address: workbookLinkFormula.address,
            })),
        ),
        c =
          d &&
          references.some(
            (reference) =>
              !d.workbook.sheets.some(
                (sheet) =>
                  sheet.name.toLowerCase() === reference.sheet.toLowerCase(),
              ),
          ),
        h =
          d &&
          references.some((reference) => {
            const y = d.workbook.sheets.find(
              (sheet) =>
                sheet.name.toLowerCase() === reference.sheet.toLowerCase(),
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
                ? u.sourceFingerprint !== encodeWorkbookFingerprint(d.workbook)
                  ? "stale"
                  : references.some(
                        (reference) =>
                          !u.sheets.find(
                            (sheet) =>
                              sheet.name.toLowerCase() ===
                                reference.sheet.toLowerCase() &&
                              parseExternalLinkRange(reference.range).every(
                                (w) => Object.hasOwn(sheet.cells, w),
                              ),
                          ),
                      )
                    ? "incomplete-cache"
                    : "matches-approved-snapshot"
                : "unresolved"
          : "unapproved-or-missing";
      return {
        target: target,
        packageTarget: workbook.workbookLinks?.packageTargets?.find(
          (packageTarget) => packageTarget.key === target,
        )?.target,
        status: f,
        references: references,
        lastResolvedRevision: u?.sourceRevision,
        lastResolvedAt: u?.resolvedAt,
        sourceName: u?.sourceName,
        sourceFingerprint: u?.sourceFingerprint,
        cached: !!u,
      };
    }),
    cells: workbookLinkFormulas,
    unsupported: workbookLinkFormulas.filter(
      (workbookLinkFormula) => workbookLinkFormula.unsupported,
    ),
    cachePolicy:
      "Values are retained snapshots. Only explicit refresh reads approved workbooks. No path or URL is fetched; matching a supplied revision does not verify a live file.",
  };
}
const workbookLinkPreviewRecords = new WeakMap();
function resolveLinkedSourceCalculation(workbook, r, target, n) {
  const o = buildWorkbookLinkPolicy(workbook, n).sources.find(
    (source) => source.target === target,
  );
  if (!o)
    throw Error(
      `Source ${target} is missing or not approved. Attach the exact source and approve its worksheets.`,
    );
  const workbookInput = normalizeWorkbookInput(o.workbook),
    s = new Set(o.allowedSheetIds);
  if (
    normalizeDatePolicy(workbookInput.datePolicy).dateSystem !==
    normalizeDatePolicy(workbook.datePolicy).dateSystem
  )
    throw Error(
      "Source and destination date systems differ. Convert and review the source date system before linking.",
    );
  const allowRead = (h) => {
      const f = workbookInput.sheets.find((sheet) => sheet.id === h);
      return (
        !!f && s.has(h) && (!(f.hidden || f.veryHidden) || !!o.includeHidden)
      );
    },
    u = evaluateWorkbookFormulas(
      workbookInput,
      {},
      { allowRead: allowRead, allowName: () => true, step: () => {} },
    ),
    sheets = [];
  let c = 0;
  for (const h of r) {
    const f = workbookInput.sheets.find(
      (sheet) => sheet.name.toLowerCase() === h.sheet.toLowerCase(),
    );
    if (!f) throw Error(`Source worksheet ${h.sheet} is missing.`);
    if (!allowRead(f.id))
      throw Error(`Source worksheet ${h.sheet} is outside the approved scope.`);
    let p = sheets.find((sheet) => sheet.id === f.id);
    p || ((p = { id: f.id, name: f.name, cells: {} }), sheets.push(p));
    for (const m of parseExternalLinkRange(h.range))
      if (!Object.hasOwn(p.cells, m)) {
        if (++c > workbookLinkLimits.cells)
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
    target: target,
    sourceId: workbookInput.id,
    sourceName: workbookInput.name,
    sourceRevision: workbookInput.revision,
    sourceFingerprint: encodeWorkbookFingerprint(workbookInput),
    resolvedAt: new Date().toISOString(),
    sheets: sheets,
  };
}
function assertLinkDestinationPublishable(e, r) {
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
function previewWorkbookLinkOperation(workbook, r, t = {}) {
  const workbookLinkPolicy = buildWorkbookLinkPolicy(workbook, t);
  if (t.readOnly) throw Error("The workbook is read-only.");
  if (!["refresh", "replace", "break"].includes(r.action))
    throw Error("Unknown link operation.");
  if (
    (r.target !== undefined && workbookQualifierSchema.parse(r.target),
    r.action !== "refresh" && !r.target)
  )
    throw Error("Choose a target workbook link.");
  if (
    r.action === "replace" &&
    (workbookQualifierSchema.parse(r.replacementTarget),
    r.replacementTarget === r.target)
  )
    throw Error("Choose a different source qualifier.");
  const workbookLinkFormulas = collectWorkbookLinkFormulas(workbook),
    a = workbookLinkFormulas.filter((workbookLinkFormula) =>
      r.target
        ? workbookLinkFormula.references.some(
            (reference) => reference.target === r.target,
          )
        : workbookLinkFormula.references.length ||
          workbookLinkFormula.unsupported,
    );
  if (!a.length) throw Error("No matching workbook link cells.");
  for (const p of a) {
    if (p.unsupported) throw Error(p.unsupported);
    if (!workbookLinkPolicy.allowed.includes(p.sheetId))
      throw Error("A destination worksheet is outside the host write scope.");
    if (
      (assertSheetUnprotected(
        workbook.sheets.find((sheet) => sheet.id === p.sheetId),
      ),
      p.literal && !r.activateLiteral)
    )
      throw Error(
        "Explicitly approve activation of inert imported formulas before resolving them.",
      );
  }
  let workbookInput = normalizeWorkbookInput(workbook);
  workbookInput.workbookLinks = structuredClone(
    workbookInput.workbookLinks ?? { schemaVersion: 1, sources: [] },
  );
  for (const p of a) {
    const m = workbookInput.sheets.find((sheet) => sheet.id === p.sheetId);
    ((m.cells[p.address] = {
      ...m.cells[p.address],
      raw:
        r.action === "replace"
          ? rewriteExternalReferences(p.formula, (v) =>
              v.target === r.target
                ? formatExternalSheetReference(
                    r.replacementTarget,
                    v.sheet,
                    v.range,
                  )
                : undefined,
            )
          : p.formula,
    }),
      delete m.cells[p.address].literal);
  }
  if (r.action !== "break") {
    const p = collectWorkbookLinkFormulas(workbookInput)
        .filter((v) =>
          a.some((g) => g.sheetId === v.sheetId && g.address === v.address),
        )
        .flatMap((v) => v.references),
      m = [...new Set(p.map((event) => event.target))];
    for (const target of m) {
      const g = collectWorkbookLinkFormulas(workbookInput)
          .flatMap((y) => y.references)
          .filter((event) => event.target === target),
        linkedSourceCalculation = resolveLinkedSourceCalculation(
          workbook,
          g,
          target,
          t,
        );
      ((workbookInput.workbookLinks.sources =
        workbookInput.workbookLinks.sources.filter(
          (source) => source.target !== target,
        )),
        workbookInput.workbookLinks.sources.push(linkedSourceCalculation));
    }
  } else if (!r.acknowledgeCached)
    throw Error(
      "Acknowledge that breaking links keeps the previewed cached values and removes the complete formulas.",
    );
  workbookInput = normalizeWorkbookInput(workbookInput);
  const d = new Set([
      ...workbookLinkPolicy.allowed,
      ...buildExternalLinkAliases(workbookInput).sheets.map(
        (sheet) => sheet.id,
      ),
    ]),
    u = evaluateWorkbookFormulas(
      workbookInput,
      {},
      { allowRead: (p) => d.has(p), step: () => {} },
    ),
    cells = a.map((p) => {
      const m = assertLinkDestinationPublishable(u, p);
      return {
        sheetId: p.sheetId,
        sheetName: p.sheetName,
        address: p.address,
        before: p.formula,
        after:
          r.action === "break"
            ? m.value
            : workbookInput.sheets.find((sheet) => sheet.id === p.sheetId)
                .cells[p.address].raw,
        result: m,
      };
    }),
    c = [];
  for (const cell of cells) {
    const m = workbookInput.sheets.find((sheet) => sheet.id === cell.sheetId),
      before = workbook.sheets.find((sheet) => sheet.id === cell.sheetId).cells[
        cell.address
      ];
    (r.action === "break" &&
      (m.cells[cell.address] = {
        ...m.cells[cell.address],
        raw: cell.result.value,
        literal: typeof cell.result.value == "string",
      }),
      c.push({
        sheetId: m.id,
        address: cell.address,
        before: before,
        after: m.cells[cell.address],
      }));
  }
  const h = new Set(
    collectWorkbookLinkFormulas(workbookInput).flatMap((p) =>
      p.references.map((reference) => reference.target),
    ),
  );
  ((workbookInput.workbookLinks = {
    ...workbookInput.workbookLinks,
    sources: workbookInput.workbookLinks.sources.filter((source) =>
      h.has(source.target),
    ),
  }),
    checkDataValidationBatch(workbookInput, c),
    (workbookInput = normalizeWorkbookInput(workbookInput)));
  const f = {
    request: structuredClone(r),
    cells: cells,
    sources: structuredClone(workbookInput.workbookLinks.sources),
    warning:
      r.action === "break"
        ? "Complete formulas are replaced with these scalar values, including any other references in those formulas. Undo can restore the prior workbook. Retained original files are unchanged."
        : "These are reviewed snapshots of approved workbooks, not live network connections. Future changes require another preview and refresh.",
  };
  return (
    workbookLinkPreviewRecords.set(f, {
      bookHash: fingerprintCanonicalValue(workbook),
      contextHash: workbookLinkPolicy.hash(),
      planHash: fingerprintCanonicalValue(f),
      candidate: workbookInput,
    }),
    f
  );
}
function applyWorkbookLinkPreview(workbook, r, t = {}) {
  const n = workbookLinkPreviewRecords.get(r);
  if (!n) throw Error("Unknown or already applied workbook link preview.");
  if (t.readOnly) throw Error("The workbook is read-only.");
  if (
    n.bookHash !== fingerprintCanonicalValue(workbook) ||
    n.contextHash !== buildWorkbookLinkPolicy(workbook, t).hash() ||
    n.planHash !== fingerprintCanonicalValue(r)
  )
    throw Error(
      "Workbook, source, review or permissions changed. Preview links again.",
    );
  return (
    workbookLinkPreviewRecords.delete(r),
    normalizeWorkbookInput(n.candidate)
  );
}

function assertSourceInspectionCurrent(inspection) {
  const r = sourceInspectionRecords.get(inspection);
  if (!r || r.signature !== JSON.stringify(inspection))
    throw Error(
      "The source inspection changed or is untrusted. Inspect the original again.",
    );
  return r.bytes.slice();
}

function assertFormatInspectionCurrent(e) {
  const r = formatInspectionRecords.get(e);
  if (!r || r.signature !== JSON.stringify(e))
    throw Error(
      "The format inspection changed or is untrusted. Inspect the original again.",
    );
  return r.bytes.slice();
}

export {
  assertSourceInspectionCurrent,
  CustomFunctionReviewBanner,
  applyWorkbookLinkPreview,
  recordSourceInspection,
  validateIsolatedFormatInspection,
  previewWorkbookLinkOperation,
  buildWorkbookLinkReview,
  assertFormatInspectionCurrent,
  AddInsAndCustomFunctionsDialog,
};
