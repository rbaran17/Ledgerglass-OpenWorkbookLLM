import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import { fingerprintCanonicalValue } from "./01-core-schemas-and-hashing.js";
import { workbookLimits } from "./02-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  consolidationFormatVersion,
  serializeTypedKeyValues,
  createWorksheetRecord,
  sumConsolidationAmounts,
  formatOccupiedRangeAddress,
  calculateWorkbookWithSettings,
  parseRangeAddress,
  consolidationLimits,
  throwDaxError,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import { importWorkbookFile } from "./07-xlsx-archive-reader.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  readSourceTableRecords,
  buildConsolidationOutputCells,
  createConsolidationPreview,
  buildAuthorizedContextSnapshot,
} from "./12-formula-dependency-inspector.js";

function publishConsolidatedSchedule(e, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(e),
    fingerprintCanonicalValue(e) !== r.workbookHash ||
      buildAuthorizedContextSnapshot(e, t) !== r.contextHash)
  )
    throw Error(
      "Workbook or authorization changed. Preview consolidation again.",
    );
  const n = createConsolidationPreview(e, r.request, t);
  if (fingerprintCanonicalValue(n) !== fingerprintCanonicalValue(r))
    throw Error(
      "Source data or the reviewed plan changed. Preview consolidation again.",
    );
  if (!n.review.complete)
    throw Error(
      "Resolve all consolidation errors before publishing. No partial result is permitted.",
    );
  const o = n.request.definition,
    a = n.review,
    s = createWorksheetRecord(r.request.outputName),
    d = o.categoryLabels.length + 4,
    u = columnIndexToLetters(d),
    l = o.categoryLabels.length + 7,
    c = [
      "Row type",
      ...o.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
  let h = 1;
  const f = [],
    p = (E, A = -1, k = "") => {
      const F = h++;
      return (
        E.forEach((I, O) => {
          s.cells[`${columnIndexToLetters(O)}${F}`] = {
            raw: I,
            ...(typeof I == "string" && I.startsWith("=") && O !== A
              ? { literal: true }
              : {}),
            ...(k && O === d ? { note: k } : {}),
            ...(O === d && F > 1
              ? { format: "number", numFmt: "#,##0.00;[Red](#,##0.00)" }
              : {}),
            ...(F === 1 || E[0] !== "Detail"
              ? {
                  style: {
                    bold: true,
                    fill: F === 1 ? "#362C63" : "#EEEAF8",
                    color: F === 1 ? "#FFFFFF" : "#29243C",
                  },
                }
              : {}),
          };
        }),
        F
      );
    };
  p(c);
  const m = (E) =>
      p(
        [
          "Detail",
          ...E.categories,
          E.period,
          E.currency,
          E.unit,
          E.link === "live"
            ? `='${E.sheetName.replaceAll("'", "''")}'!$${E.address.replace(/\d+$/, "")}$${E.row}`
            : E.amount,
          E.locator,
          E.link === "live"
            ? "Live amount · reviewed context"
            : E.link === "approved snapshot"
              ? "Approved workbook snapshot"
              : "Reviewed snapshot",
        ],
        E.link === "live" ? d : -1,
        `Reviewed amount ${E.amount}; period ${E.period}; currency ${E.currency}; unit ${E.unit}. ${E.locator}. Source ID ${E.sourceId}; workbook ${E.workbookId}; worksheet ${E.sheetId}. Source freshness is displayed in the native workbook banner.`,
      ),
    v = (E, A) => {
      const k = new Map();
      for (const F of E) {
        const I = serializeTypedKeyValues(F.categories.slice(0, A)),
          O = k.get(I) ?? [];
        (O.push(F), k.set(I, O));
      }
      for (const F of k.values()) {
        const I = h;
        (A < o.categoryLabels.length ? v(F, A + 1) : F.forEach(m),
          f.push({
            address: `${u}${h}`,
            expected: sumConsolidationAmounts(F.map((z) => z.amount)),
            count: F.length,
          }));
        const O = h - 1,
          j = Array.from({ length: o.categoryLabels.length }, (z, V) =>
            V < A ? F[0].categories[V] : null,
          );
        (p(
          [
            `Subtotal ${A}`,
            ...j,
            o.period,
            o.currency,
            o.unit,
            `=SUBTOTAL(9,${u}${I}:${u}${O})`,
            `${F.length} contributions`,
            "Review source status",
          ],
          d,
        ),
          (s.outlineGroups ?? (s.outlineGroups = [])).push({
            id: crypto.randomUUID(),
            axis: "row",
            start: I - 1,
            end: O - 1,
            collapsed: o.collapseDetails,
          }));
      }
    };
  if (
    (v(a.contributions, 1),
    f.push({
      address: `${u}${h}`,
      expected: a.grandTotal,
      count: a.contributions.length,
    }),
    p(
      [
        "Grand total",
        ...o.categoryLabels.map(() => null),
        o.period,
        o.currency,
        o.unit,
        `=SUBTOTAL(9,${u}2:${u}${h - 1})`,
        `${a.contributions.length} contributions`,
        "Review source status",
      ],
      d,
      `Reviewed grand total ${a.grandTotal}; source controls ${a.sourceTotal}; difference ${a.difference}. No implicit conversion.`,
    ),
    h - 1 > consolidationLimits.outputRows ||
      (h - 1) * l > consolidationLimits.outputCells ||
      e.sheets.reduce((E, A) => E + Object.keys(A.cells).length, 0) +
        Object.keys(s.cells).length >
        workbookLimits.maxCells)
  )
    throw Error(
      "Consolidation exceeds the output or workbook cell budget. No worksheet was published.",
    );
  ((s.frozenRows = 1),
    (s.columnWidths = Object.fromEntries(
      c.map((E, A) => [A, A === l - 2 ? 340 : A === l - 1 ? 260 : 130]),
    )));
  const g = `A1:${columnIndexToLetters(l - 1)}${h - 1}`,
    b = {
      workbookId: e.id,
      revision: e.revision,
      at: new Date().toISOString(),
    };
  s.consolidation = {
    engine: consolidationFormatVersion,
    definition: o,
    review: a,
    origin: b,
    evidenceHash: fingerprintCanonicalValue({
      definition: o,
      review: a,
      origin: b,
    }),
    outputRange: g,
    outputHash: buildConsolidationOutputCells(s, g),
  };
  const y = normalizeWorkbookInput({ ...e, sheets: [...e.sheets, s] }),
    w = calculateWorkbookWithSettings(y);
  for (const E of f) {
    const A = w(s.id, E.address),
      k = Number.EPSILON * Math.max(1, Math.abs(E.expected)) * E.count * 4;
    if (
      typeof A != "number" ||
      !Number.isFinite(A) ||
      Math.abs(A - E.expected) > k
    )
      throw Error(
        "Calculated subtotals do not reconcile with the reviewed amounts, or a calculation limit was reached. Reduce numeric cancellation or explicitly round source amounts before publishing.",
      );
  }
  return y;
}
function checkConsolidationFreshness(e, r, t = {}) {
  const n = r.consolidation;
  if (!n) throw Error("This worksheet has no consolidation record.");
  const o = [];
  try {
    (n.origin.workbookId !== e.id &&
      o.push(
        "Historical result from another workbook identity. Review and rebind its sources before refreshing.",
      ),
      buildConsolidationOutputCells(r, n.outputRange) !== n.outputHash &&
        o.push(
          "Published result cells or source notes were edited. Reconcile or publish a fresh result.",
        ));
    for (const a of n.definition.sources)
      try {
        const s = readSourceTableRecords(e, a, t),
          d = n.review.sources.find((u) => u.sourceId === a.id);
        s.fingerprint !== d?.fingerprint &&
          o.push(
            `${a.id}: source data, context, layout or formula dependencies changed.`,
          );
      } catch (s) {
        o.push(`${a.id}: ${s.message}`);
      }
  } catch (a) {
    o.push(a.message);
  }
  return { state: o.length ? "stale" : "current", messages: o };
}
function saveBlobToFile(e, r) {
  const t = URL.createObjectURL(e),
    n = document.createElement("a");
  ((n.href = t),
    (n.download = r),
    n.click(),
    setTimeout(() => URL.revokeObjectURL(t), 5e3));
}
function findTableRangesForSource(e, r) {
  try {
    const o = e?.sheets.find((s) => s.id === r.sheetId);
    if (!o) return [];
    const a = parseRangeAddress(
      o.tables?.find((s) => s.id === r.tableId)?.range ?? r.range,
    );
    return a.lastCol - a.firstCol >= 50
      ? []
      : Array.from({ length: a.lastCol - a.firstCol + 1 }, (s, d) => {
          return o.cells[
            `${columnIndexToLetters(a.firstCol + d)}${a.firstRow + 1}`
          ]?.raw;
        }).filter((s) => typeof s == "string" && !!s);
  } catch {
    return [];
  }
}
function createConsolidationDefinition(e, r, t, n) {
  const o = {
      id: crypto.randomUUID(),
      workbookId: e.id,
      sheetId: r,
      range: t,
      categories: [...n.categoryLabels],
      amountField: "Amount",
      period: { kind: "constant", value: n.period },
      currency: { kind: "constant", value: n.currency },
      unit: { kind: "constant", value: n.unit },
      duplicatePolicy: "reject",
    },
    a = findTableRangesForSource(e, o);
  ((o.categories = n.categoryLabels.map((s, d) =>
    a.includes(s) ? s : (a[d] ?? s),
  )),
    (o.amountField = a.includes("Amount") ? "Amount" : (a.at(-1) ?? "Amount")));
  for (const s of ["period", "currency", "unit"]) {
    const d = s[0].toUpperCase() + s.slice(1);
    a.includes(d) && (o[s] = { kind: "field", field: d });
  }
  return o;
}
function ConsolidationStatusBanner({
  workbook,
  sheetId,
  options = {},
  onReview,
}) {
  const { text, language } = useWorkbookMessages(),
    s = workbook.sheets.find((l) => l.id === sheetId),
    d = s != null && s.consolidation ? JSON.stringify([workbook, options]) : "",
    u = ne.useMemo(
      () =>
        s != null && s.consolidation
          ? checkConsolidationFreshness(workbook, s, options)
          : null,
      [d, sheetId],
    );
  return u
    ? i.jsxs("div", {
        className: `lg-consolidation-status ${u.state}`,
        role: "status",
        "aria-label": text("Consolidation source status"),
        lang: language("Consolidation source status"),
        "data-ui-label": "Consolidation source status",
        children: [
          i.jsxs("span", {
            children: [
              i.jsx("strong", {
                children:
                  u.state === "stale"
                    ? "Stale consolidation"
                    : "Consolidation sources current",
              }),
              " ",
              "·",
              " ",
              u.messages.length
                ? u.messages.join(" ")
                : "Amounts reconcile to the reviewed sources. External amounts are approved snapshots.",
            ],
          }),
          i.jsx("button", {
            onClick: onReview,
            children: i.jsx(LocalizedText, { children: "Review / refresh" }),
          }),
        ],
      })
    : null;
}
function ConsolidationDialog(e) {
  var W;
  const { text, language } = useWorkbookMessages(),
    n = useCancellableOperation(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    s = ne.useRef(0),
    d = ne.useRef(""),
    u = ne.useRef(true),
    l = e.allowedSheetIds ?? e.workbook.sheets.map((K) => K.id),
    c = JSON.stringify([
      e.workbook,
      e.sheetId,
      e.range,
      e.readOnly,
      l,
      e.allowHidden,
      e.allowFileSources,
      e.approvedWorkbooks,
    ]);
  d.current = c;
  const [h, f] = ne.useState(() => {
      const K =
        e.initialDefinition ??
        e.workbook.sheets.find((Se) => Se.id === e.sheetId)?.consolidation
          ?.definition;
      if (K) return structuredClone(K);
      const te = {
          engine: consolidationFormatVersion,
          categoryLabels: ["Category"],
          period: "2026-Q1",
          currency: "USD",
          unit: "ones",
          linkedAmounts: true,
          collapseDetails: false,
          sources: [],
        },
        ae = e.workbook.sheets.find((Se) => Se.id === e.sheetId),
        oe = e.range.includes(":") ? e.range : formatOccupiedRangeAddress(ae);
      return (
        (te.sources = [
          createConsolidationDefinition(e.workbook, ae.id, oe, te),
        ]),
        te
      );
    }),
    [p, m] = ne.useState(() => {
      let K = "Consolidated",
        te = 2;
      for (
        ;
        e.workbook.sheets.some(
          (ae) => ae.name.toLowerCase() === K.toLowerCase(),
        );
      )
        K = `Consolidated ${te++}`;
      return K;
    }),
    [v, g] = ne.useState(false),
    [b, y] = ne.useState([]),
    [w, E] = ne.useState(null),
    [A, k] = ne.useState(false),
    [F, I] = ne.useState(""),
    [O, j] = ne.useState(null),
    [z, V] = ne.useState(0),
    [D, B] = ne.useState("contributions"),
    _ = [
      ...(e.approvedWorkbooks ?? []),
      ...b.filter((K) => {
        var te;
        return !(
          (te = e.approvedWorkbooks) != null &&
          te.some((ae) => ae.workbook.id === K.workbook.id)
        );
      }),
    ],
    C = {
      readOnly: e.readOnly,
      allowedSheetIds: l,
      includeHidden: !!e.allowHidden && v,
      approvedWorkbooks: _,
    },
    T = JSON.stringify([c, h, p, C]),
    N = O?.stamp === T,
    R = N ? O.plan.review : null,
    L = [
      {
        workbook: e.workbook,
        allowedSheetIds: l,
        includeHidden: !!e.allowHidden,
      },
      ..._,
    ];
  (ne.useEffect(() => {
    var ae;
    u.current = true;
    const K = document.activeElement;
    return (
      (ae = a.current?.querySelector("input,select,button")) == null ||
        ae.focus(),
      () => {
        ((u.current = false),
          s.current++,
          K != null && K.isConnected && K.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      var K, te;
      N &&
        ((te = (K = o.current) == null ? undefined : K.scrollIntoView) ==
          null ||
          te.call(K, { block: "start" }));
    }, [O, N]));
  const $ = (K) => {
      (f(K), j(null), I(""), V(0));
    },
    U = (K, te) =>
      $({ ...h, sources: h.sources.map((ae, oe) => (oe === K ? te : ae)) }),
    G = (K) => {
      try {
        (I(""), K());
      } catch (te) {
        I(te.message);
      }
    };
  async function H(K) {
    if (!K) return;
    const te = ++s.current,
      ae = d.current;
    (E(null), k(true), I(""));
    try {
      if (!/\.(json|xlsx)$/i.test(K.name))
        throw Error("Choose a native workbook JSON or an XLSX file.");
      const oe = await importWorkbookFile(K, n.begin());
      if (!u.current || te !== s.current || ae !== d.current) return;
      if (oe.workbook.id === e.workbook.id)
        throw Error(
          "This is the current workbook identity. Use the current workbook source instead.",
        );
      E({ ...oe, stamp: ae });
    } catch (oe) {
      u.current && te === s.current && ae === d.current && I(oe.message);
    } finally {
      u.current && te === s.current && k(false);
    }
  }
  function Q(K) {
    (e.onApprovedWorkbooksChange ? e.onApprovedWorkbooksChange(K) : y(K),
      j(null));
  }
  const X$1 = R
      ? D === "contributions"
        ? R.contributions
        : D === "subtotals"
          ? R.subtotals
          : D === "sources"
            ? R.sources
            : R.exceptions
      : [],
    J = Math.max(1, Math.ceil(X$1.length / 25)),
    q = Math.min(z, J - 1);
  return i.jsx("div", {
    className: "lg-modal-backdrop lg-consolidation-backdrop",
    children: i.jsxs("section", {
      ref: a,
      className: "lg-consolidation",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": "consolidation-title",
      onKeyDown: (K) => {
        if (
          (K.key === "Escape" && (K.stopPropagation(), e.onClose()),
          K.key === "Tab")
        ) {
          const te = [
              ...a.current.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),a[href]",
              ),
            ].filter((de) => !de.hasAttribute("hidden")),
            ae = te[0],
            oe = te.at(-1);
          K.shiftKey && document.activeElement === ae
            ? (K.preventDefault(), oe == null || oe.focus())
            : !K.shiftKey &&
              document.activeElement === oe &&
              (K.preventDefault(), ae == null || ae.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", {
                  className: "eyebrow",
                  children: i.jsx(LocalizedText, {
                    children: "Data · Consolidation",
                  }),
                }),
                i.jsx("h2", {
                  id: "consolidation-title",
                  children: i.jsx(LocalizedText, {
                    children: "Consolidate schedules & subtotals",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Align categories, verify source context, then publish a grouped schedule.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close consolidation"),
              onClick: e.onClose,
              lang: language("Close consolidation"),
              "data-ui-label": "Close consolidation",
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-consolidation-body",
          children: [
            i.jsxs("section", {
              className: "lg-consolidation-settings",
              "aria-label": text("Consolidation settings"),
              lang: language("Consolidation settings"),
              "data-ui-label": "Consolidation settings",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Output worksheet" }),
                    i.jsx("input", {
                      "aria-label": text("Output worksheet"),
                      value: p,
                      maxLength: 31,
                      onChange: (K) => {
                        (m(K.target.value), j(null));
                      },
                      lang: language("Output worksheet"),
                      "data-ui-label": "Output worksheet",
                    }),
                  ],
                }),
                ["period", "currency", "unit"].map((K) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx(LocalizedText, { children: "Target " }),
                        K,
                        i.jsx("input", {
                          "aria-label": `Target ${K}`,
                          value: h[K],
                          onChange: (te) => $({ ...h, [K]: te.target.value }),
                        }),
                      ],
                    },
                    K,
                  ),
                ),
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Category levels" }),
                    i.jsx("select", {
                      "aria-label": text("Category levels"),
                      value: h.categoryLabels.length,
                      onChange: (K) => {
                        const te = Number(K.target.value),
                          ae = Array.from(
                            { length: te },
                            (oe, de) =>
                              h.categoryLabels[de] ?? `Category ${de + 1}`,
                          );
                        $({
                          ...h,
                          categoryLabels: ae,
                          sources: h.sources.map((oe) => ({
                            ...oe,
                            categories: ae.map(
                              (de, he) => oe.categories[he] ?? "",
                            ),
                          })),
                        });
                      },
                      lang: language("Category levels"),
                      "data-ui-label": "Category levels",
                      children: [1, 2, 3].map((K) =>
                        i.jsx("option", { children: K }, K),
                      ),
                    }),
                  ],
                }),
                h.categoryLabels.map((K, te) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx(LocalizedText, { children: "Category " }),
                        te + 1,
                        " ",
                        i.jsx(LocalizedText, { children: "label" }),
                        i.jsx("input", {
                          "aria-label": `Category ${te + 1} label`,
                          value: K,
                          onChange: (ae) =>
                            $({
                              ...h,
                              categoryLabels: h.categoryLabels.map((oe, de) =>
                                de === te ? ae.target.value : oe,
                              ),
                            }),
                        }),
                      ],
                    },
                    te,
                  ),
                ),
              ],
            }),
            i.jsxs("div", {
              className: "lg-consolidation-options",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: h.linkedAmounts,
                      onChange: (K) =>
                        $({ ...h, linkedAmounts: K.target.checked }),
                    }),
                    i.jsx(LocalizedText, {
                      children: "Link amounts from this workbook",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx("input", {
                      type: "checkbox",
                      checked: h.collapseDetails,
                      onChange: (K) =>
                        $({ ...h, collapseDetails: K.target.checked }),
                    }),
                    i.jsx(LocalizedText, {
                      children: "Collapse detail groups",
                    }),
                  ],
                }),
                e.allowHidden &&
                  i.jsxs("label", {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: v,
                        onChange: (K) => {
                          (g(K.target.checked), j(null));
                        },
                      }),
                      i.jsx(LocalizedText, {
                        children: "Include authorized hidden source cells",
                      }),
                    ],
                  }),
              ],
            }),
            i.jsx("p", {
              className: "lg-consolidation-hint",
              children: i.jsx(LocalizedText, {
                children:
                  "Periods, currencies and units must match exactly. Amounts are summed; no implicit conversion. Same-workbook links recalculate amounts; changed source context still requires review. Approved external workbooks use refreshable snapshots.",
              }),
            }),
            i.jsxs("section", {
              "aria-label": text("Consolidation sources"),
              lang: language("Consolidation sources"),
              "data-ui-label": "Consolidation sources",
              children: [
                h.sources.map((K, te) => {
                  var ce;
                  const ae = L.find((se) => se.workbook.id === K.workbookId),
                    oe = ae?.workbook,
                    de =
                      oe?.sheets.filter(
                        (se) =>
                          ae.allowedSheetIds.includes(se.id) &&
                          ((!se.hidden && !se.veryHidden) ||
                            (v && ae.includeHidden)),
                      ) ?? [],
                    he = de.find((se) => se.id === K.sheetId),
                    Se = he ? findTableRangesForSource(oe, K) : [],
                    le = (se, we, Ee) =>
                      i.jsxs("label", {
                        children: [
                          se,
                          i.jsxs("select", {
                            "aria-label": `Source ${te + 1} ${se}`,
                            value: we,
                            onChange: (Ae) => Ee(Ae.target.value),
                            children: [
                              !Se.includes(we) &&
                                i.jsx("option", {
                                  value: we,
                                  children: we || "Choose field",
                                }),
                              Se.map((Ae) =>
                                i.jsx(
                                  "option",
                                  { value: Ae, children: Ae },
                                  Ae,
                                ),
                              ),
                            ],
                          }),
                        ],
                      });
                  return i.jsxs(
                    "article",
                    {
                      className: "lg-consolidation-source",
                      children: [
                        i.jsxs("div", {
                          className: "lg-consolidation-source-title",
                          children: [
                            i.jsxs("strong", {
                              children: [
                                i.jsx(LocalizedText, { children: "Source " }),
                                te + 1,
                              ],
                            }),
                            i.jsxs("button", {
                              disabled: h.sources.length === 1,
                              onClick: () =>
                                $({
                                  ...h,
                                  sources: h.sources.filter(
                                    (se, we) => te !== we,
                                  ),
                                }),
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Remove source ",
                                }),
                                te + 1,
                              ],
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "lg-consolidation-fields",
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Workbook" }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${te + 1} Workbook`,
                                  value: K.workbookId,
                                  onChange: (se) => {
                                    const we = L.find(
                                        (Ae) =>
                                          Ae.workbook.id === se.target.value,
                                      ),
                                      Ee = we.workbook.sheets.find(
                                        (Ae) =>
                                          we.allowedSheetIds.includes(Ae.id) &&
                                          !Ae.hidden &&
                                          !Ae.veryHidden,
                                      );
                                    Ee &&
                                      U(te, {
                                        ...createConsolidationDefinition(
                                          we.workbook,
                                          Ee.id,
                                          formatOccupiedRangeAddress(Ee),
                                          h,
                                        ),
                                        id: K.id,
                                      });
                                  },
                                  children: [
                                    !ae &&
                                      i.jsx("option", {
                                        value: K.workbookId,
                                        children:
                                          "Missing or unapproved workbook",
                                      }),
                                    L.map((se) =>
                                      i.jsxs(
                                        "option",
                                        {
                                          value: se.workbook.id,
                                          children: [
                                            se.workbook.name,
                                            se.workbook.id === e.workbook.id
                                              ? " (current)"
                                              : " (approved)",
                                          ],
                                        },
                                        se.workbook.id,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, { children: "Worksheet" }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${te + 1} Worksheet`,
                                  value: K.sheetId,
                                  onChange: (se) => {
                                    const we = de.find(
                                      (Ee) => Ee.id === se.target.value,
                                    );
                                    U(te, {
                                      ...createConsolidationDefinition(
                                        oe,
                                        we.id,
                                        formatOccupiedRangeAddress(we),
                                        h,
                                      ),
                                      id: K.id,
                                    });
                                  },
                                  children: [
                                    !he &&
                                      i.jsx("option", {
                                        value: K.sheetId,
                                        children:
                                          "Missing or unauthorized worksheet",
                                      }),
                                    de.map((se) =>
                                      i.jsx(
                                        "option",
                                        { value: se.id, children: se.name },
                                        se.id,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Source type",
                                }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${te + 1} Source type`,
                                  value: K.tableId ?? "",
                                  onChange: (se) => {
                                    const { tableId, ...Ee } = K;
                                    U(te, {
                                      ...Ee,
                                      ...(se.target.value
                                        ? { tableId: se.target.value }
                                        : {}),
                                    });
                                  },
                                  children: [
                                    i.jsx("option", {
                                      value: "",
                                      children: "Explicit range",
                                    }),
                                    K.tableId &&
                                      !(
                                        (ce = he?.tables) != null &&
                                        ce.some((se) => se.id === K.tableId)
                                      ) &&
                                      i.jsx("option", {
                                        value: K.tableId,
                                        children: "Missing table",
                                      }),
                                    he?.tables?.map((se) =>
                                      i.jsx(
                                        "option",
                                        { value: se.id, children: se.name },
                                        se.id,
                                      ),
                                    ),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Header & data range",
                                }),
                                i.jsx("input", {
                                  "aria-label": `Source ${te + 1} Header & data range`,
                                  value: K.range,
                                  disabled: !!K.tableId,
                                  onChange: (se) =>
                                    U(te, { ...K, range: se.target.value }),
                                }),
                              ],
                            }),
                            h.categoryLabels.map((se, we) =>
                              i.jsx(
                                "div",
                                {
                                  children: le(
                                    `Category ${we + 1}`,
                                    K.categories[we],
                                    (Ee) =>
                                      U(te, {
                                        ...K,
                                        categories: K.categories.map(
                                          (Ae, ye) => (ye === we ? Ee : Ae),
                                        ),
                                      }),
                                  ),
                                },
                                we,
                              ),
                            ),
                            le("Amount", K.amountField, (se) =>
                              U(te, { ...K, amountField: se }),
                            ),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Repeated categories",
                                }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${te + 1} Repeated categories`,
                                  value: K.duplicatePolicy,
                                  onChange: (se) =>
                                    U(te, {
                                      ...K,
                                      duplicatePolicy: se.target.value,
                                    }),
                                  children: [
                                    i.jsx("option", {
                                      value: "reject",
                                      children: "Reject duplicates",
                                    }),
                                    i.jsx("option", {
                                      value: "sum",
                                      children: "Sum repeated categories",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            ["period", "currency", "unit"].map((se) =>
                              i.jsxs(
                                "div",
                                {
                                  className: "lg-consolidation-context",
                                  children: [
                                    i.jsxs("label", {
                                      children: [
                                        se,
                                        " ",
                                        i.jsx(LocalizedText, {
                                          children: "source",
                                        }),
                                        i.jsxs("select", {
                                          "aria-label": `Source ${te + 1} ${se} source`,
                                          value: K[se].kind,
                                          onChange: (we) =>
                                            U(te, {
                                              ...K,
                                              [se]:
                                                we.target.value === "constant"
                                                  ? {
                                                      kind: "constant",
                                                      value: h[se],
                                                    }
                                                  : {
                                                      kind: "field",
                                                      field: Se[0] ?? "",
                                                    },
                                            }),
                                          children: [
                                            i.jsx("option", {
                                              value: "constant",
                                              children: "Declared constant",
                                            }),
                                            i.jsx("option", {
                                              value: "field",
                                              children: "Source column",
                                            }),
                                          ],
                                        }),
                                      ],
                                    }),
                                    K[se].kind === "field"
                                      ? le(se, K[se].field, (we) =>
                                          U(te, {
                                            ...K,
                                            [se]: { kind: "field", field: we },
                                          }),
                                        )
                                      : i.jsxs("label", {
                                          children: [
                                            se,
                                            i.jsx("input", {
                                              "aria-label": `Source ${te + 1} ${se}`,
                                              value: K[se].value,
                                              onChange: (we) =>
                                                U(te, {
                                                  ...K,
                                                  [se]: {
                                                    kind: "constant",
                                                    value: we.target.value,
                                                  },
                                                }),
                                            }),
                                          ],
                                        }),
                                  ],
                                },
                                se,
                              ),
                            ),
                          ],
                        }),
                      ],
                    },
                    K.id,
                  );
                }),
                i.jsx("button", {
                  disabled: h.sources.length >= 8,
                  onClick: () => {
                    const K = e.workbook.sheets.find(
                      (te) => l.includes(te.id) && !te.hidden && !te.veryHidden,
                    );
                    K &&
                      $({
                        ...h,
                        sources: [
                          ...h.sources,
                          createConsolidationDefinition(
                            e.workbook,
                            K.id,
                            formatOccupiedRangeAddress(K),
                            h,
                          ),
                        ],
                      });
                  },
                  children: i.jsx(LocalizedText, { children: "Add source" }),
                }),
              ],
            }),
            e.allowFileSources &&
              i.jsxs("details", {
                className: "lg-consolidation-import",
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children: "Approved workbook files",
                    }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Choose a local native JSON or XLSX workbook, review the import, then explicitly approve its visible worksheets. Files stay in this session; nothing is fetched automatically.",
                    }),
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Inspect workbook file",
                      }),
                      i.jsx("input", {
                        type: "file",
                        accept: ".json,.xlsx",
                        "aria-label": text("Inspect consolidation workbook"),
                        disabled: A || _.length >= 8,
                        onChange: (K) => {
                          (H(K.target.files?.[0]), (K.target.value = ""));
                        },
                        lang: language("Inspect consolidation workbook"),
                        "data-ui-label": "Inspect consolidation workbook",
                      }),
                    ],
                  }),
                  A &&
                    i.jsx("p", {
                      role: "status",
                      children: i.jsx(LocalizedText, {
                        children: "Inspecting workbook…",
                      }),
                    }),
                  w?.stamp === c &&
                    i.jsxs("div", {
                      "aria-label": text("Pending workbook approval"),
                      lang: language("Pending workbook approval"),
                      "data-ui-label": "Pending workbook approval",
                      children: [
                        i.jsx("strong", { children: w.workbook.name }),
                        i.jsx("p", {
                          children: w.workbook.sheets
                            .filter((K) => !K.hidden && !K.veryHidden)
                            .map((K) => K.name)
                            .join(", "),
                        }),
                        w.warnings.map((K, te) =>
                          i.jsx("p", { children: K }, te),
                        ),
                        i.jsx("button", {
                          onClick: () => {
                            (Q([
                              ..._.filter(
                                (K) => K.workbook.id !== w.workbook.id,
                              ),
                              {
                                workbook: w.workbook,
                                allowedSheetIds: w.workbook.sheets
                                  .filter((K) => !K.hidden && !K.veryHidden)
                                  .map((K) => K.id),
                                includeHidden: false,
                              },
                            ]),
                              E(null));
                          },
                          children: i.jsx(LocalizedText, {
                            children: "Approve workbook",
                          }),
                        }),
                        i.jsx("button", {
                          onClick: () => E(null),
                          children: i.jsx(LocalizedText, {
                            children: "Discard file",
                          }),
                        }),
                      ],
                    }),
                  _.map((K) => {
                    var te;
                    return i.jsxs(
                      "p",
                      {
                        children: [
                          K.workbook.name,
                          " · ",
                          K.allowedSheetIds.length,
                          " ",
                          i.jsx(LocalizedText, { children: "approved sheets" }),
                          " ",
                          i.jsxs("button", {
                            disabled:
                              !e.onApprovedWorkbooksChange &&
                              !!(
                                (te = e.approvedWorkbooks) != null &&
                                te.some(
                                  (ae) => ae.workbook.id === K.workbook.id,
                                )
                              ),
                            onClick: () =>
                              Q(
                                _.filter(
                                  (ae) => ae.workbook.id !== K.workbook.id,
                                ),
                              ),
                            children: [
                              i.jsx(LocalizedText, { children: "Revoke " }),
                              K.workbook.name,
                            ],
                          }),
                        ],
                      },
                      K.workbook.id,
                    );
                  }),
                ],
              }),
            F &&
              i.jsx("p", {
                role: "alert",
                className: "lg-consolidation-error",
                children: F,
              }),
            R &&
              i.jsxs("section", {
                ref: o,
                "aria-label": text("Consolidation preview"),
                className: "lg-consolidation-preview",
                lang: language("Consolidation preview"),
                "data-ui-label": "Consolidation preview",
                children: [
                  i.jsx("h3", {
                    children: R.complete
                      ? "Ready for review"
                      : "Resolve exceptions before publishing",
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-totals",
                    children: [
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Contributions " }),
                          i.jsx("strong", { children: R.contributions.length }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Grand total " }),
                          i.jsx("strong", {
                            children: R.grandTotal ?? "Blocked",
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Source controls" }),
                          " ",
                          i.jsx("strong", {
                            children: R.sourceTotal ?? "Blocked",
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Difference " }),
                          i.jsx("strong", { children: R.difference ?? "—" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      h.period,
                      " · ",
                      h.currency,
                      " · ",
                      h.unit,
                      ".",
                      " ",
                      R.exceptions.length,
                      " ",
                      i.jsx(LocalizedText, { children: "exceptions." }),
                      " ",
                      R.complete
                        ? "Review all contributions and source controls before publishing."
                        : "No totals or partial output can be published while errors remain.",
                    ],
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-tabs",
                    children: [
                      [
                        "contributions",
                        "subtotals",
                        "sources",
                        "exceptions",
                      ].map((K) =>
                        i.jsx(
                          "button",
                          {
                            "aria-pressed": D === K,
                            onClick: () => {
                              (B(K), V(0));
                            },
                            children: K[0].toUpperCase() + K.slice(1),
                          },
                          K,
                        ),
                      ),
                      i.jsx("button", {
                        onClick: () =>
                          saveBlobToFile(
                            new Blob(
                              [
                                JSON.stringify(
                                  {
                                    definition: O.plan.request.definition,
                                    review: R,
                                  },
                                  null,
                                  2,
                                ),
                              ],
                              { type: "application/json" },
                            ),
                            "Consolidation-review.json",
                          ),
                        children: i.jsx(LocalizedText, {
                          children: "Save review",
                        }),
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-table",
                    children: [
                      i.jsxs("table", {
                        "aria-label": `Consolidation ${D}`,
                        children: [
                          i.jsx("thead", {
                            children: i.jsx("tr", {
                              children: (D === "contributions"
                                ? [
                                    "Category path",
                                    "Period / currency / unit",
                                    "Amount",
                                    "Source amount",
                                    "Link",
                                  ]
                                : D === "subtotals"
                                  ? [
                                      "Category path",
                                      "Level",
                                      "Contributions",
                                      "Subtotal",
                                    ]
                                  : D === "sources"
                                    ? [
                                        "Source",
                                        "Range",
                                        "Records / accepted",
                                        "Context",
                                        "Amount",
                                        "Fingerprint",
                                      ]
                                    : ["Source / row", "Severity", "Exception"]
                              ).map((K) => i.jsx("th", { children: K }, K)),
                            }),
                          }),
                          i.jsxs("tbody", {
                            children: [
                              D === "contributions" &&
                                R.contributions
                                  .slice(q * 25, q * 25 + 25)
                                  .map((K, te) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsx("td", {
                                            children: K.categories
                                              .map(
                                                (ae) =>
                                                  `${typeof ae}: ${String(ae)}`,
                                              )
                                              .join(" / "),
                                          }),
                                          i.jsxs("td", {
                                            children: [
                                              K.period,
                                              " / ",
                                              K.currency,
                                              " / ",
                                              K.unit,
                                            ],
                                          }),
                                          i.jsx("td", { children: K.amount }),
                                          i.jsx("td", { children: K.locator }),
                                          i.jsx("td", { children: K.link }),
                                        ],
                                      },
                                      te,
                                    ),
                                  ),
                              D === "subtotals" &&
                                R.subtotals
                                  .slice(q * 25, q * 25 + 25)
                                  .map((K, te) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsx("td", {
                                            children: K.categories
                                              .map(String)
                                              .join(" / "),
                                          }),
                                          i.jsx("td", { children: K.level }),
                                          i.jsx("td", { children: K.count }),
                                          i.jsx("td", { children: K.amount }),
                                        ],
                                      },
                                      te,
                                    ),
                                  ),
                              D === "sources" &&
                                R.sources.slice(q * 25, q * 25 + 25).map((K) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsxs("td", {
                                          children: [
                                            K.workbookName,
                                            " / ",
                                            K.sheetName,
                                          ],
                                        }),
                                        i.jsx("td", { children: K.range }),
                                        i.jsxs("td", {
                                          children: [
                                            K.records,
                                            " / ",
                                            K.acceptedRecords,
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            K.periods.join(", "),
                                            " / ",
                                            K.currencies.join(", "),
                                            " ",
                                            "/ ",
                                            K.units.join(", "),
                                          ],
                                        }),
                                        i.jsx("td", { children: K.amount }),
                                        i.jsx("td", {
                                          children: i.jsx("code", {
                                            children: K.fingerprint,
                                          }),
                                        }),
                                      ],
                                    },
                                    K.sourceId,
                                  ),
                                ),
                              D === "exceptions" &&
                                R.exceptions
                                  .slice(q * 25, q * 25 + 25)
                                  .map((K, te) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              i.jsx(LocalizedText, {
                                                children: "Source",
                                              }),
                                              " ",
                                              h.sources.findIndex(
                                                (ae) => ae.id === K.sourceId,
                                              ) + 1,
                                              K.row ? ` / row ${K.row}` : "",
                                            ],
                                          }),
                                          i.jsx("td", { children: K.severity }),
                                          i.jsxs("td", {
                                            children: [K.code, ": ", K.message],
                                          }),
                                        ],
                                      },
                                      te,
                                    ),
                                  ),
                            ],
                          }),
                        ],
                      }),
                      !X$1.length &&
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, { children: "No " }),
                            D,
                            ".",
                          ],
                        }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-pages",
                    children: [
                      i.jsx("button", {
                        disabled: q === 0,
                        onClick: () => V(q - 1),
                        children: i.jsx(LocalizedText, {
                          children: "Previous page",
                        }),
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Page " }),
                          q + 1,
                          " ",
                          i.jsx(LocalizedText, { children: "of " }),
                          J,
                          " · ",
                          X$1.length,
                          " ",
                          i.jsx(LocalizedText, { children: "records" }),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: q + 1 >= J,
                        onClick: () => V(q + 1),
                        children: i.jsx(LocalizedText, {
                          children: "Next page",
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
            i.jsx("span", {
              children: e.readOnly
                ? "Read-only: preview available"
                : "Separate output · originals retained · Undo / Redo",
            }),
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              onClick: () =>
                G(() => {
                  const K = createConsolidationPreview(
                    e.workbook,
                    { definition: h, outputName: p },
                    C,
                  );
                  (j({ plan: K, stamp: T }),
                    V(0),
                    B(K.review.complete ? "contributions" : "exceptions"));
                }),
              children: i.jsx(LocalizedText, {
                children: "Preview consolidation",
              }),
            }),
            i.jsx("button", {
              className: "primary",
              disabled:
                !(R != null && R.complete) ||
                e.readOnly ||
                !!((W = e.workbook.structureProtection) != null && W.enabled),
              onClick: () =>
                G(() => {
                  if (!N) return;
                  const K = publishConsolidatedSchedule(e.workbook, O.plan, C);
                  (j(null),
                    e.onWorkbookChange(K, "Publish consolidated schedule"),
                    e.onClose());
                }),
              children: i.jsx(LocalizedText, {
                children: "Publish consolidated schedule",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const caseFoldText = (e) => e.toLowerCase(),
  buildTypedValueKey = (e) =>
    JSON.stringify([typeof e, typeof e == "string" ? caseFoldText(e) : e]),
  buildTableColumnKey = (e) => `${e.tableId}\0${e.column}`,
  requireScalarValue = (e) =>
    typeof e == "object" && e !== null
      ? throwDaxError("DAX_TYPE", "A table cannot be used as a scalar.")
      : e,
  requireLogicalOperand = (e) =>
    e === null
      ? false
      : typeof e == "boolean"
        ? e
        : typeof e == "number"
          ? e !== 0
          : throwDaxError(
              "DAX_TYPE",
              "A logical condition requires a boolean, number or BLANK.",
            ),
  requireNumericOperand = (e) =>
    e === null
      ? 0
      : typeof e == "number"
        ? e
        : throwDaxError(
            "DAX_TYPE",
            "Numeric expressions require numbers or BLANK; implicit text/date conversion is unsupported.",
          ),
  requireFiniteNumber = (e) =>
    Number.isFinite(e)
      ? e
      : throwDaxError(
          "DAX_NUMBER",
          "The result is not finite. Use DIVIDE for a blank or zero denominator.",
        );
function compareDaxValues(e, r, t = false) {
  if (t && (e === null || r === null)) return e === r ? 0 : undefined;
  if (e === null && r === null) return 0;
  if (
    (e === null &&
      (e = typeof r == "string" ? "" : typeof r == "boolean" ? false : 0),
    r === null &&
      (r = typeof e == "string" ? "" : typeof e == "boolean" ? false : 0),
    typeof e != typeof r)
  )
    return throwDaxError(
      "DAX_TYPE",
      "Comparison operands must have compatible types.",
    );
  const n = typeof e == "string" ? caseFoldText(e) : e,
    o = typeof r == "string" ? caseFoldText(r) : r;
  return n === o ? 0 : n < o ? -1 : 1;
}
function sumWithCompensation(e) {
  let r = 0,
    t = 0,
    n = 0;
  for (const o of e) {
    if (typeof o != "number") continue;
    const a = r + o;
    ((t += Math.abs(r) >= Math.abs(o) ? r - a + o : o - a + r), (r = a), n++);
  }
  return n ? requireFiniteNumber(r + t) : null;
}

export {
  requireNumericOperand,
  requireLogicalOperand,
  buildTableColumnKey,
  requireScalarValue,
  requireFiniteNumber,
  ConsolidationDialog,
  ConsolidationStatusBanner,
  sumWithCompensation,
  compareDaxValues,
  caseFoldText,
  buildTypedValueKey,
};
export { saveBlobToFile };
