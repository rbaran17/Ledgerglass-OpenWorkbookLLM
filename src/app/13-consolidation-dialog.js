import i from "react/jsx-runtime";
import ne from "react";
import { X } from "lucide-react";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
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
import { importWorkbookFile } from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  readSourceTableRecords,
  buildConsolidationOutputCells,
  createConsolidationPreview,
  buildAuthorizedContextSnapshot,
} from "./12-consolidation-and-dependency-inspection.js";

function publishConsolidatedSchedule(workbook, plan, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    fingerprintCanonicalValue(workbook) !== plan.workbookHash ||
      buildAuthorizedContextSnapshot(workbook, options) !== plan.contextHash)
  )
    throw Error(
      "Workbook or authorization changed. Preview consolidation again.",
    );
  const consolidationPreview = createConsolidationPreview(
    workbook,
    plan.request,
    options,
  );
  if (
    fingerprintCanonicalValue(consolidationPreview) !==
    fingerprintCanonicalValue(plan)
  )
    throw Error(
      "Source data or the reviewed plan changed. Preview consolidation again.",
    );
  if (!consolidationPreview.review.complete)
    throw Error(
      "Resolve all consolidation errors before publishing. No partial result is permitted.",
    );
  const definition = consolidationPreview.request.definition,
    review = consolidationPreview.review,
    worksheetRecord = createWorksheetRecord(plan.request.outputName),
    d = definition.categoryLabels.length + 4,
    u = columnIndexToLetters(d),
    l = definition.categoryLabels.length + 7,
    c = [
      "Row type",
      ...definition.categoryLabels,
      "Period",
      "Currency",
      "Unit",
      "Amount",
      "Source",
      "Link state",
    ];
  let h = 1;
  const f = [],
    p = (E, A = -1, note = "") => {
      const F = h++;
      return (
        E.forEach((raw, O) => {
          worksheetRecord.cells[`${columnIndexToLetters(O)}${F}`] = {
            raw: raw,
            ...(typeof raw == "string" && raw.startsWith("=") && O !== A
              ? { literal: true }
              : {}),
            ...(note && O === d ? { note: note } : {}),
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
        (A < definition.categoryLabels.length ? v(F, A + 1) : F.forEach(m),
          f.push({
            address: `${u}${h}`,
            expected: sumConsolidationAmounts(F.map((z) => z.amount)),
            count: F.length,
          }));
        const O = h - 1,
          j = Array.from(
            { length: definition.categoryLabels.length },
            (z, V) => (V < A ? F[0].categories[V] : null),
          );
        (p(
          [
            `Subtotal ${A}`,
            ...j,
            definition.period,
            definition.currency,
            definition.unit,
            `=SUBTOTAL(9,${u}${I}:${u}${O})`,
            `${F.length} contributions`,
            "Review source status",
          ],
          d,
        ),
          (worksheetRecord.outlineGroups ??= []).push({
            id: crypto.randomUUID(),
            axis: "row",
            start: I - 1,
            end: O - 1,
            collapsed: definition.collapseDetails,
          }));
      }
    };
  if (
    (v(review.contributions, 1),
    f.push({
      address: `${u}${h}`,
      expected: review.grandTotal,
      count: review.contributions.length,
    }),
    p(
      [
        "Grand total",
        ...definition.categoryLabels.map(() => null),
        definition.period,
        definition.currency,
        definition.unit,
        `=SUBTOTAL(9,${u}2:${u}${h - 1})`,
        `${review.contributions.length} contributions`,
        "Review source status",
      ],
      d,
      `Reviewed grand total ${review.grandTotal}; source controls ${review.sourceTotal}; difference ${review.difference}. No implicit conversion.`,
    ),
    h - 1 > consolidationLimits.outputRows ||
      (h - 1) * l > consolidationLimits.outputCells ||
      workbook.sheets.reduce((E, A) => E + Object.keys(A.cells).length, 0) +
        Object.keys(worksheetRecord.cells).length >
        workbookLimits.maxCells)
  )
    throw Error(
      "Consolidation exceeds the output or workbook cell budget. No worksheet was published.",
    );
  ((worksheetRecord.frozenRows = 1),
    (worksheetRecord.columnWidths = Object.fromEntries(
      c.map((E, A) => [A, A === l - 2 ? 340 : A === l - 1 ? 260 : 130]),
    )));
  const outputRange = `A1:${columnIndexToLetters(l - 1)}${h - 1}`,
    b = {
      workbookId: workbook.id,
      revision: workbook.revision,
      at: new Date().toISOString(),
    };
  worksheetRecord.consolidation = {
    engine: consolidationFormatVersion,
    definition: definition,
    review: review,
    origin: b,
    evidenceHash: fingerprintCanonicalValue({
      definition: definition,
      review: review,
      origin: b,
    }),
    outputRange: outputRange,
    outputHash: buildConsolidationOutputCells(worksheetRecord, outputRange),
  };
  const workbookInput = normalizeWorkbookInput({
      ...workbook,
      sheets: [...workbook.sheets, worksheetRecord],
    }),
    workbookWithSettings = calculateWorkbookWithSettings(workbookInput);
  for (const E of f) {
    const A = workbookWithSettings(worksheetRecord.id, E.address),
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
  return workbookInput;
}
function checkConsolidationFreshness(workbook, worksheetRecord, options = {}) {
  const consolidation = worksheetRecord.consolidation;
  if (!consolidation)
    throw Error("This worksheet has no consolidation record.");
  const messages = [];
  try {
    (consolidation.origin.workbookId !== workbook.id &&
      messages.push(
        "Historical result from another workbook identity. Review and rebind its sources before refreshing.",
      ),
      buildConsolidationOutputCells(
        worksheetRecord,
        consolidation.outputRange,
      ) !== consolidation.outputHash &&
        messages.push(
          "Published result cells or source notes were edited. Reconcile or publish a fresh result.",
        ));
    for (const source of consolidation.definition.sources)
      try {
        const sourceTableRecords = readSourceTableRecords(
            workbook,
            source,
            options,
          ),
          d = consolidation.review.sources.find(
            (u) => u.sourceId === source.id,
          );
        sourceTableRecords.fingerprint !== d?.fingerprint &&
          messages.push(
            `${source.id}: source data, context, layout or formula dependencies changed.`,
          );
      } catch (error) {
        messages.push(`${source.id}: ${error.message}`);
      }
  } catch (error) {
    messages.push(error.message);
  }
  return { state: messages.length ? "stale" : "current", messages: messages };
}
function saveBlobToFile(e, r) {
  const t = URL.createObjectURL(e),
    n = document.createElement("a");
  ((n.href = t),
    (n.download = r),
    n.click(),
    setTimeout(() => URL.revokeObjectURL(t), 5e3));
}
function findTableRangesForSource(workbook, source) {
  try {
    const o = workbook?.sheets.find((sheet) => sheet.id === source.sheetId);
    if (!o) return [];
    const rangeAddress = parseRangeAddress(
      o.tables?.find((table) => table.id === source.tableId)?.range ??
        source.range,
    );
    return rangeAddress.lastCol - rangeAddress.firstCol >= 50
      ? []
      : Array.from(
          { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
          (s, d) =>
            o.cells[
              `${columnIndexToLetters(rangeAddress.firstCol + d)}${rangeAddress.firstRow + 1}`
            ]?.raw,
        ).filter((s) => typeof s == "string" && !!s);
  } catch {
    return [];
  }
}
function createConsolidationDefinition(workbook, sheetId, range, definition) {
  const source = {
      id: crypto.randomUUID(),
      workbookId: workbook.id,
      sheetId: sheetId,
      range: range,
      categories: [...definition.categoryLabels],
      amountField: "Amount",
      period: { kind: "constant", value: definition.period },
      currency: { kind: "constant", value: definition.currency },
      unit: { kind: "constant", value: definition.unit },
      duplicatePolicy: "reject",
    },
    tableRangesForSource = findTableRangesForSource(workbook, source);
  ((source.categories = definition.categoryLabels.map((categoryLabel, d) =>
    tableRangesForSource.includes(categoryLabel)
      ? categoryLabel
      : (tableRangesForSource[d] ?? categoryLabel),
  )),
    (source.amountField = tableRangesForSource.includes("Amount")
      ? "Amount"
      : (tableRangesForSource.at(-1) ?? "Amount")));
  for (const s of ["period", "currency", "unit"]) {
    const field = s[0].toUpperCase() + s.slice(1);
    tableRangesForSource.includes(field) &&
      (source[s] = { kind: "field", field: field });
  }
  return source;
}
function ConsolidationStatusBanner({
  workbook,
  sheetId,
  options = {},
  onReview,
}) {
  const { text, language } = useWorkbookMessages(),
    sheet = workbook.sheets.find((sheet) => sheet.id === sheetId),
    d = sheet?.consolidation ? JSON.stringify([workbook, options]) : "",
    u = ne.useMemo(
      () =>
        sheet?.consolidation
          ? checkConsolidationFreshness(workbook, sheet, options)
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
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    s = ne.useRef(0),
    d = ne.useRef(""),
    u = ne.useRef(true),
    allowedSheetIds =
      e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
    c = JSON.stringify([
      e.workbook,
      e.sheetId,
      e.range,
      e.readOnly,
      allowedSheetIds,
      e.allowHidden,
      e.allowFileSources,
      e.approvedWorkbooks,
    ]);
  d.current = c;
  const [definition, f] = ne.useState(() => {
      const K =
        e.initialDefinition ??
        e.workbook.sheets.find((sheet) => sheet.id === e.sheetId)?.consolidation
          ?.definition;
      if (K) return structuredClone(K);
      const definition = {
          engine: consolidationFormatVersion,
          categoryLabels: ["Category"],
          period: "2026-Q1",
          currency: "USD",
          unit: "ones",
          linkedAmounts: true,
          collapseDetails: false,
          sources: [],
        },
        sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
        range = e.range.includes(":")
          ? e.range
          : formatOccupiedRangeAddress(sheet);
      return (
        (definition.sources = [
          createConsolidationDefinition(
            e.workbook,
            sheet.id,
            range,
            definition,
          ),
        ]),
        definition
      );
    }),
    [outputName, m] = ne.useState(() => {
      let K = "Consolidated",
        te = 2;
      for (
        ;
        e.workbook.sheets.some(
          (sheet) => sheet.name.toLowerCase() === K.toLowerCase(),
        );
      )
        K = `Consolidated ${te++}`;
      return K;
    }),
    [checked, g] = ne.useState(false),
    [b, y] = ne.useState([]),
    [w, E] = ne.useState(null),
    [A, k] = ne.useState(false),
    [F, I] = ne.useState(""),
    [O, j] = ne.useState(null),
    [z, V] = ne.useState(0),
    [D, B] = ne.useState("contributions"),
    approvedWorkbooks = [
      ...(e.approvedWorkbooks ?? []),
      ...b.filter(
        (K) =>
          !e.approvedWorkbooks?.some(
            (approvedWorkbook) =>
              approvedWorkbook.workbook.id === K.workbook.id,
          ),
      ),
    ],
    options = {
      readOnly: e.readOnly,
      allowedSheetIds: allowedSheetIds,
      includeHidden: !!e.allowHidden && checked,
      approvedWorkbooks: approvedWorkbooks,
    },
    stamp = JSON.stringify([c, definition, outputName, options]),
    N = O?.stamp === stamp,
    review = N ? O.plan.review : null,
    L = [
      {
        workbook: e.workbook,
        allowedSheetIds: allowedSheetIds,
        includeHidden: !!e.allowHidden,
      },
      ...approvedWorkbooks,
    ];
  (ne.useEffect(() => {
    u.current = true;
    const activeElement = document.activeElement;
    return (
      a.current?.querySelector("input,select,button")?.focus(),
      () => {
        ((u.current = false),
          s.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      var K;
      N &&
        ((K = o.current) == null ? undefined : K.scrollIntoView)?.call(K, {
          block: "start",
        });
    }, [O, N]));
  const $ = (K) => {
      (f(K), j(null), I(""), V(0));
    },
    U = (K, te) =>
      $({
        ...definition,
        sources: definition.sources.map((source, oe) =>
          oe === K ? te : source,
        ),
      }),
    G = (K) => {
      try {
        (I(""), K());
      } catch (error) {
        I(error.message);
      }
    };
  async function H(K) {
    if (!K) return;
    const te = ++s.current,
      stamp = d.current;
    (E(null), k(true), I(""));
    try {
      if (!/\.(json|xlsx)$/i.test(K.name))
        throw Error("Choose a native workbook JSON or an XLSX file.");
      const oe = await importWorkbookFile(K, cancellableOperation.begin());
      if (!u.current || te !== s.current || stamp !== d.current) return;
      if (oe.workbook.id === e.workbook.id)
        throw Error(
          "This is the current workbook identity. Use the current workbook source instead.",
        );
      E({ ...oe, stamp: stamp });
    } catch (error) {
      u.current && te === s.current && stamp === d.current && I(error.message);
    } finally {
      u.current && te === s.current && k(false);
    }
  }
  function Q(K) {
    (e.onApprovedWorkbooksChange ? e.onApprovedWorkbooksChange(K) : y(K),
      j(null));
  }
  const X$1 = review
      ? D === "contributions"
        ? review.contributions
        : D === "subtotals"
          ? review.subtotals
          : D === "sources"
            ? review.sources
            : review.exceptions
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
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), e.onClose()),
          event.key === "Tab")
        ) {
          const te = [
              ...a.current.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),a[href]",
              ),
            ].filter((de) => !de.hasAttribute("hidden")),
            ae = te[0],
            oe = te.at(-1);
          event.shiftKey && document.activeElement === ae
            ? (event.preventDefault(), oe == null || oe.focus())
            : !event.shiftKey &&
              document.activeElement === oe &&
              (event.preventDefault(), ae == null || ae.focus());
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
                      value: outputName,
                      maxLength: 31,
                      onChange: (event) => {
                        (m(event.target.value), j(null));
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
                          value: definition[K],
                          onChange: (event) =>
                            $({ ...definition, [K]: event.target.value }),
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
                      value: definition.categoryLabels.length,
                      onChange: (event) => {
                        const te = Number(event.target.value),
                          categoryLabels = Array.from(
                            { length: te },
                            (oe, de) =>
                              definition.categoryLabels[de] ??
                              `Category ${de + 1}`,
                          );
                        $({
                          ...definition,
                          categoryLabels: categoryLabels,
                          sources: definition.sources.map((source) => ({
                            ...source,
                            categories: categoryLabels.map(
                              (categoryLabel, he) =>
                                source.categories[he] ?? "",
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
                definition.categoryLabels.map((categoryLabel, te) =>
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
                          value: categoryLabel,
                          onChange: (event) =>
                            $({
                              ...definition,
                              categoryLabels: definition.categoryLabels.map(
                                (categoryLabel, de) =>
                                  de === te
                                    ? event.target.value
                                    : categoryLabel,
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
                      checked: definition.linkedAmounts,
                      onChange: (event) =>
                        $({
                          ...definition,
                          linkedAmounts: event.target.checked,
                        }),
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
                      checked: definition.collapseDetails,
                      onChange: (event) =>
                        $({
                          ...definition,
                          collapseDetails: event.target.checked,
                        }),
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
                        checked: checked,
                        onChange: (event) => {
                          (g(event.target.checked), j(null));
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
                definition.sources.map((source, te) => {
                  const ae = L.find(
                      (se) => se.workbook.id === source.workbookId,
                    ),
                    workbook = ae?.workbook,
                    de =
                      workbook?.sheets.filter(
                        (sheet) =>
                          ae.allowedSheetIds.includes(sheet.id) &&
                          ((!sheet.hidden && !sheet.veryHidden) ||
                            (checked && ae.includeHidden)),
                      ) ?? [],
                    he = de.find((se) => se.id === source.sheetId),
                    Se = he ? findTableRangesForSource(workbook, source) : [],
                    le = (se, we, Ee) =>
                      i.jsxs("label", {
                        children: [
                          se,
                          i.jsxs("select", {
                            "aria-label": `Source ${te + 1} ${se}`,
                            value: we,
                            onChange: (event) => Ee(event.target.value),
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
                              disabled: definition.sources.length === 1,
                              onClick: () =>
                                $({
                                  ...definition,
                                  sources: definition.sources.filter(
                                    (source, we) => te !== we,
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
                                  value: source.workbookId,
                                  onChange: (event) => {
                                    const we = L.find(
                                        (Ae) =>
                                          Ae.workbook.id === event.target.value,
                                      ),
                                      sheet = we.workbook.sheets.find(
                                        (sheet) =>
                                          we.allowedSheetIds.includes(
                                            sheet.id,
                                          ) &&
                                          !sheet.hidden &&
                                          !sheet.veryHidden,
                                      );
                                    sheet &&
                                      U(te, {
                                        ...createConsolidationDefinition(
                                          we.workbook,
                                          sheet.id,
                                          formatOccupiedRangeAddress(sheet),
                                          definition,
                                        ),
                                        id: source.id,
                                      });
                                  },
                                  children: [
                                    !ae &&
                                      i.jsx("option", {
                                        value: source.workbookId,
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
                                  value: source.sheetId,
                                  onChange: (event) => {
                                    const we = de.find(
                                      (Ee) => Ee.id === event.target.value,
                                    );
                                    U(te, {
                                      ...createConsolidationDefinition(
                                        workbook,
                                        we.id,
                                        formatOccupiedRangeAddress(we),
                                        definition,
                                      ),
                                      id: source.id,
                                    });
                                  },
                                  children: [
                                    !he &&
                                      i.jsx("option", {
                                        value: source.sheetId,
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
                                  value: source.tableId ?? "",
                                  onChange: (event) => {
                                    const { tableId, ...Ee } = source;
                                    U(te, {
                                      ...Ee,
                                      ...(event.target.value
                                        ? { tableId: event.target.value }
                                        : {}),
                                    });
                                  },
                                  children: [
                                    i.jsx("option", {
                                      value: "",
                                      children: "Explicit range",
                                    }),
                                    source.tableId &&
                                      !he?.tables?.some(
                                        (table) => table.id === source.tableId,
                                      ) &&
                                      i.jsx("option", {
                                        value: source.tableId,
                                        children: "Missing table",
                                      }),
                                    he?.tables?.map((table) =>
                                      i.jsx(
                                        "option",
                                        {
                                          value: table.id,
                                          children: table.name,
                                        },
                                        table.id,
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
                                  value: source.range,
                                  disabled: !!source.tableId,
                                  onChange: (event) =>
                                    U(te, {
                                      ...source,
                                      range: event.target.value,
                                    }),
                                }),
                              ],
                            }),
                            definition.categoryLabels.map((categoryLabel, we) =>
                              i.jsx(
                                "div",
                                {
                                  children: le(
                                    `Category ${we + 1}`,
                                    source.categories[we],
                                    (Ee) =>
                                      U(te, {
                                        ...source,
                                        categories: source.categories.map(
                                          (category, ye) =>
                                            ye === we ? Ee : category,
                                        ),
                                      }),
                                  ),
                                },
                                we,
                              ),
                            ),
                            le("Amount", source.amountField, (amountField) =>
                              U(te, { ...source, amountField: amountField }),
                            ),
                            i.jsxs("label", {
                              children: [
                                i.jsx(LocalizedText, {
                                  children: "Repeated categories",
                                }),
                                i.jsxs("select", {
                                  "aria-label": `Source ${te + 1} Repeated categories`,
                                  value: source.duplicatePolicy,
                                  onChange: (event) =>
                                    U(te, {
                                      ...source,
                                      duplicatePolicy: event.target.value,
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
                                          value: source[se].kind,
                                          onChange: (event) =>
                                            U(te, {
                                              ...source,
                                              [se]:
                                                event.target.value ===
                                                "constant"
                                                  ? {
                                                      kind: "constant",
                                                      value: definition[se],
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
                                    source[se].kind === "field"
                                      ? le(se, source[se].field, (field) =>
                                          U(te, {
                                            ...source,
                                            [se]: {
                                              kind: "field",
                                              field: field,
                                            },
                                          }),
                                        )
                                      : i.jsxs("label", {
                                          children: [
                                            se,
                                            i.jsx("input", {
                                              "aria-label": `Source ${te + 1} ${se}`,
                                              value: source[se].value,
                                              onChange: (event) =>
                                                U(te, {
                                                  ...source,
                                                  [se]: {
                                                    kind: "constant",
                                                    value: event.target.value,
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
                    source.id,
                  );
                }),
                i.jsx("button", {
                  disabled: definition.sources.length >= 8,
                  onClick: () => {
                    const sheet = e.workbook.sheets.find(
                      (sheet) =>
                        allowedSheetIds.includes(sheet.id) &&
                        !sheet.hidden &&
                        !sheet.veryHidden,
                    );
                    sheet &&
                      $({
                        ...definition,
                        sources: [
                          ...definition.sources,
                          createConsolidationDefinition(
                            e.workbook,
                            sheet.id,
                            formatOccupiedRangeAddress(sheet),
                            definition,
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
                        disabled: A || approvedWorkbooks.length >= 8,
                        onChange: (event) => {
                          (H(event.target.files?.[0]),
                            (event.target.value = ""));
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
                            .filter(
                              (sheet) => !sheet.hidden && !sheet.veryHidden,
                            )
                            .map((K) => K.name)
                            .join(", "),
                        }),
                        w.warnings.map((warning, te) =>
                          i.jsx("p", { children: warning }, te),
                        ),
                        i.jsx("button", {
                          onClick: () => {
                            (Q([
                              ...approvedWorkbooks.filter(
                                (approvedWorkbook) =>
                                  approvedWorkbook.workbook.id !==
                                  w.workbook.id,
                              ),
                              {
                                workbook: w.workbook,
                                allowedSheetIds: w.workbook.sheets
                                  .filter(
                                    (sheet) =>
                                      !sheet.hidden && !sheet.veryHidden,
                                  )
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
                  approvedWorkbooks.map((K) =>
                    i.jsxs(
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
                              !!e.approvedWorkbooks?.some(
                                (approvedWorkbook) =>
                                  approvedWorkbook.workbook.id ===
                                  K.workbook.id,
                              ),
                            onClick: () =>
                              Q(
                                approvedWorkbooks.filter(
                                  (approvedWorkbook) =>
                                    approvedWorkbook.workbook.id !==
                                    K.workbook.id,
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
                    ),
                  ),
                ],
              }),
            F &&
              i.jsx("p", {
                role: "alert",
                className: "lg-consolidation-error",
                children: F,
              }),
            review &&
              i.jsxs("section", {
                ref: o,
                "aria-label": text("Consolidation preview"),
                className: "lg-consolidation-preview",
                lang: language("Consolidation preview"),
                "data-ui-label": "Consolidation preview",
                children: [
                  i.jsx("h3", {
                    children: review.complete
                      ? "Ready for review"
                      : "Resolve exceptions before publishing",
                  }),
                  i.jsxs("div", {
                    className: "lg-consolidation-totals",
                    children: [
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Contributions " }),
                          i.jsx("strong", {
                            children: review.contributions.length,
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Grand total " }),
                          i.jsx("strong", {
                            children: review.grandTotal ?? "Blocked",
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Source controls" }),
                          " ",
                          i.jsx("strong", {
                            children: review.sourceTotal ?? "Blocked",
                          }),
                        ],
                      }),
                      i.jsxs("span", {
                        children: [
                          i.jsx(LocalizedText, { children: "Difference " }),
                          i.jsx("strong", {
                            children: review.difference ?? "—",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      definition.period,
                      " · ",
                      definition.currency,
                      " · ",
                      definition.unit,
                      ".",
                      " ",
                      review.exceptions.length,
                      " ",
                      i.jsx(LocalizedText, { children: "exceptions." }),
                      " ",
                      review.complete
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
                                    review: review,
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
                                review.contributions
                                  .slice(q * 25, q * 25 + 25)
                                  .map((K, te) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsx("td", {
                                            children: K.categories
                                              .map(
                                                (category) =>
                                                  `${typeof category}: ${String(category)}`,
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
                                review.subtotals
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
                                review.sources
                                  .slice(q * 25, q * 25 + 25)
                                  .map((K) =>
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
                                review.exceptions
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
                                              definition.sources.findIndex(
                                                (source) =>
                                                  source.id === K.sourceId,
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
                  const consolidationPreview = createConsolidationPreview(
                    e.workbook,
                    { definition: definition, outputName: outputName },
                    options,
                  );
                  (j({ plan: consolidationPreview, stamp: stamp }),
                    V(0),
                    B(
                      consolidationPreview.review.complete
                        ? "contributions"
                        : "exceptions",
                    ));
                }),
              children: i.jsx(LocalizedText, {
                children: "Preview consolidation",
              }),
            }),
            i.jsx("button", {
              className: "primary",
              disabled:
                !review?.complete ||
                e.readOnly ||
                !!e.workbook.structureProtection?.enabled,
              onClick: () =>
                G(() => {
                  if (!N) return;
                  const K = publishConsolidatedSchedule(
                    e.workbook,
                    O.plan,
                    options,
                  );
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
