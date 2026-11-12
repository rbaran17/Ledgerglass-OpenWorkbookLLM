import { listRangeColumnIndexes } from "./21-aging-and-materiality-screens.js";
import { throwErrorMessage, isNullOrEmptyValue } from "./03-workbook-model.js";
import { cloneStructuredValue } from "./11-workbook-formats-dialog.js";
import { parseNumericText } from "./37-linear-solver-engine.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  string,
  number,
  boolean,
  array,
  enum as _enum,
  discriminatedUnion,
  literal,
  object,
} from "zod";
import { sha256HexDigest } from "./01-core-schemas-and-hashing.js";
import {
  findRowOrigin,
  parseRangeAddress,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  importNativeJsonFile,
  toFileProcessingError,
} from "./07-xlsx-archive-reader.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  materialityScreenLimits,
  loadMaterialityConfiguration,
  checkMaterialityReportFreshness,
  buildMaterialitySourceEvidence,
  exportMaterialityReportJson,
  runMaterialityScreen,
} from "./21-aging-and-materiality-screens.js";
import {
  comparatorSymbols,
  formatCellValueLabel,
} from "./49-outlines-and-comments.js";

function MaterialityRiskScreen(e) {
  const { text, language } = useWorkbookMessages(),
    n = useCancellableOperation(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(""),
    d = ne.useRef(0),
    u = ne.useRef(true),
    l = listRangeColumnIndexes(e.range),
    [c, h] = ne.useState(e.range),
    [f, p] = ne.useState(true),
    [m, v] = ne.useState(l[0] ?? 0),
    [g, b] = ne.useState(l[1] ?? 1),
    [y, w] = ne.useState(l[2] ?? 2),
    [E, A] = ne.useState("USD"),
    [k, F] = ne.useState("50000"),
    [I, O] = ne.useState("absolute"),
    [j, z] = ne.useState("gte"),
    [V, D] = ne.useState("1900-01-01"),
    [B, _] = ne.useState("9999-12-31"),
    [C, T] = ne.useState(false),
    [N, R] = ne.useState(null),
    [L, $] = ne.useState(""),
    [U, G] = ne.useState(""),
    [H, Q] = ne.useState(""),
    [X$1, J] = ne.useState(""),
    [q, W] = ne.useState(0),
    [K, te] = ne.useState("all"),
    ae = {
      sheetId: e.sheetId,
      range: c,
      headers: f,
      amountColumn: m,
      currencyColumn: g,
      dateColumn: y,
      currency: E,
      threshold: k.trim() === "" ? NaN : Number(k),
      amountMode: I,
      comparator: j,
      effectiveFrom: V,
      effectiveTo: B,
    },
    oe = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && C,
      allowedSheetIds: e.allowedSheetIds,
    },
    de = e.workbook.sheets.find((Me) => Me.id === e.sheetId),
    he = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((Me) => Me.id)),
    ].sort(),
    Se =
      !!de &&
      he.includes(de.id) &&
      (oe.includeHidden || (!de.hidden && !de.veryHidden)),
    le = JSON.stringify([e.sheetId, e.range]),
    ce = JSON.stringify([e.workbook, he, oe.includeHidden, le]),
    ge = JSON.stringify([ce, ae]);
  s.current = ge;
  const se = N
      ? L !== le
        ? {
            current: false,
            disclose: false,
            reason: "The source selection changed. Run the screen again.",
          }
        : checkMaterialityReportFreshness(e.workbook, N, ae, oe)
      : null,
    we = listRangeColumnIndexes(c),
    Ee = se != null && se.disclose ? N : null,
    Ae = Ee?.rows.filter((Me) => K === "all" || Me.disposition === K) ?? [],
    ye = Math.max(1, Math.ceil(Ae.length / materialityScreenLimits.pageSize)),
    Te = Math.min(q, ye - 1),
    Pe = (Me) => {
      (G(
        Me instanceof Error
          ? Me.message
          : "The screen could not complete. Review the settings and source.",
      ),
        Q(ce),
        J(""));
    };
  ne.useEffect(() => {
    var je;
    u.current = true;
    const Me = document.activeElement;
    return (
      (je = a.current?.querySelector('[data-ui-label="Source range"]')) ==
        null || je.focus(),
      () => {
        ((u.current = false),
          d.current++,
          Me != null && Me.isConnected && Me.focus());
      }
    );
  }, []);
  const qe = () => {
      d.current++;
      try {
        const Me = runMaterialityScreen(e.workbook, ae, oe);
        (R(Me), $(le), G(""), J(""), W(0), te("all"));
      } catch (Me) {
        (R(null), Pe(Me));
      }
    },
    ht = () => {
      try {
        if (!N || !(se != null && se.current))
          throw Error("Run a current screen before downloading a report.");
        const Me = exportMaterialityReportJson(e.workbook, N, ae, oe),
          ke = URL.createObjectURL(
            new Blob([Me], { type: "application/json" }),
          ),
          je = document.createElement("a");
        ((je.href = ke),
          (je.download = "materiality-screen.json"),
          je.click(),
          URL.revokeObjectURL(ke),
          J(
            "Report downloaded with settings, source rows, exclusions and control totals.",
          ),
          G(""));
      } catch (Me) {
        Pe(Me);
      }
    },
    Re = async (Me) => {
      if (!Me) return;
      const ke = ge,
        je = ++d.current,
        $e = () => u.current && je === d.current && s.current === ke;
      try {
        if (Me.size > 2e6)
          throw Error(
            "Choose a configuration/report no larger than 2,000,000 bytes.",
          );
        const Je = await importNativeJsonFile(Me, n.begin());
        if (!$e()) return;
        const ze = loadMaterialityConfiguration(e.workbook, Je, oe);
        if (ze.sheetId !== e.sheetId)
          throw Error(
            "Open the saved configuration's source worksheet before loading it.",
          );
        (h(ze.range),
          p(ze.headers),
          v(ze.amountColumn),
          b(ze.currencyColumn),
          w(ze.dateColumn),
          A(ze.currency),
          F(String(ze.threshold)),
          O(ze.amountMode),
          z(ze.comparator),
          D(ze.effectiveFrom),
          _(ze.effectiveTo),
          R(null),
          G(""),
          J(
            "Configuration loaded. Run the screen against the current source; saved conclusions are not reused.",
          ));
      } catch (Je) {
        $e() && Pe(toFileProcessingError(Je));
      }
    };
  return i.jsx("div", {
    className: "materiality-backdrop",
    onMouseDown: (Me) => {
      Me.target === Me.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "materiality-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${o}-title`,
      ref: a,
      onKeyDown: (Me) => {
        var je, $e;
        if (
          (Me.stopPropagation(),
          Me.key === "Escape" && (Me.preventDefault(), e.onClose()),
          Me.key === "Tab")
        ) {
          const Je = Array.from(
            a.current?.querySelectorAll(
              'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex="0"]',
            ) ?? [],
          ).filter(
            (ze) =>
              !ze.closest("details:not([open])") || ze.tagName === "SUMMARY",
          );
          Me.shiftKey && document.activeElement === Je[0]
            ? (Me.preventDefault(), (je = Je.at(-1)) == null || je.focus())
            : !Me.shiftKey &&
              document.activeElement === Je.at(-1) &&
              (Me.preventDefault(), ($e = Je[0]) == null || $e.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${o}-title`,
                  children: i.jsx(LocalizedText, {
                    children: "Materiality and risk screen",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Review a user-defined quantitative threshold against a declared currency and period.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close materiality screen"),
              onClick: e.onClose,
              lang: language("Close materiality screen"),
              "data-ui-label": "Close materiality screen",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "This screen applies your settings; it does not recommend audit materiality. It reads the source without editing it. Protected and read-only workbooks support inspection.",
          }),
        }),
        i.jsxs("div", {
          className: "materiality-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source worksheet" }),
                i.jsx("strong", {
                  children: Se ? de.name : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: Se ? c : "",
                  onChange: (Me) => h(Me.target.value),
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "materiality-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: f,
                  onChange: (Me) => p(Me.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            [
              ["Amount column", m, v],
              ["Currency column", g, b],
              ["Date column", y, w],
            ].map(([Me, ke, je]) =>
              i.jsxs(
                "label",
                {
                  children: [
                    Me,
                    i.jsx("select", {
                      "aria-label": Me,
                      value: ke,
                      onChange: ($e) => je(Number($e.target.value)),
                      children: we.map(($e) =>
                        i.jsx(
                          "option",
                          { value: $e, children: columnIndexToLetters($e) },
                          $e,
                        ),
                      ),
                    }),
                  ],
                },
                Me,
              ),
            ),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Declared currency" }),
                i.jsx("input", {
                  "aria-label": text("Declared currency"),
                  value: E,
                  maxLength: 3,
                  onChange: (Me) => A(Me.target.value),
                  lang: language("Declared currency"),
                  "data-ui-label": "Declared currency",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Threshold" }),
                i.jsx(LocalizedNumberField, {
                  "aria-label": text("Threshold"),
                  min: "0",
                  step: "any",
                  value: k,
                  onValueChange: (Me) => F(Me),
                  lang: language("Threshold"),
                  "data-ui-label": "Threshold",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Amount comparison" }),
                i.jsxs("select", {
                  "aria-label": text("Amount comparison"),
                  value: I,
                  onChange: (Me) => O(Me.target.value),
                  lang: language("Amount comparison"),
                  "data-ui-label": "Amount comparison",
                  children: [
                    i.jsx("option", {
                      value: "absolute",
                      children: "Absolute amount",
                    }),
                    i.jsx("option", {
                      value: "signed",
                      children: "Signed amount",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Comparison rule" }),
                i.jsxs("select", {
                  "aria-label": text("Comparison rule"),
                  value: j,
                  onChange: (Me) => z(Me.target.value),
                  lang: language("Comparison rule"),
                  "data-ui-label": "Comparison rule",
                  children: [
                    i.jsx("option", {
                      value: "gte",
                      children: "Greater than or equal to (≥)",
                    }),
                    i.jsx("option", {
                      value: "gt",
                      children: "Greater than (>)",
                    }),
                    i.jsx("option", {
                      value: "lte",
                      children: "Less than or equal to (≤)",
                    }),
                    i.jsx("option", { value: "lt", children: "Less than (<)" }),
                    i.jsx("option", {
                      value: "eq",
                      children: "Exactly equal to (=)",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Effective from (inclusive)",
                }),
                i.jsx("input", {
                  "aria-label": text("Effective from"),
                  type: "date",
                  value: V,
                  onChange: (Me) => D(Me.target.value),
                  lang: language("Effective from"),
                  "data-ui-label": "Effective from",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Effective to (inclusive)" }),
                i.jsx("input", {
                  "aria-label": text("Effective to"),
                  type: "date",
                  value: B,
                  onChange: (Me) => _(Me.target.value),
                  lang: language("Effective to"),
                  "data-ui-label": "Effective to",
                }),
              ],
            }),
          ],
        }),
        e.allowHidden &&
          i.jsxs("label", {
            className: "materiality-check",
            children: [
              i.jsx("input", {
                "aria-label": text("Include hidden content"),
                type: "checkbox",
                checked: C,
                onChange: (Me) => T(Me.target.checked),
                lang: language("Include hidden content"),
                "data-ui-label": "Include hidden content",
              }),
              i.jsx(LocalizedText, {
                children:
                  "Explicitly include hidden source and dependency cells within the permitted host context",
              }),
            ],
          }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Up to 500 selected cells. Numeric text is an exception. Dates must be ISO calendar days or integer Excel 1900-system serials. Currency identifiers are case-insensitive, with no trimming or conversion. Missing or mixed currency blocks the entire screen.",
          }),
        }),
        i.jsxs("div", {
          className: "materiality-actions",
          children: [
            i.jsx("button", {
              className: "materiality-primary",
              disabled: !Se,
              onClick: qe,
              children: i.jsx(LocalizedText, {
                children: "Run materiality screen",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Load saved configuration/report",
                }),
                i.jsx("input", {
                  "aria-label": text("Load saved configuration"),
                  type: "file",
                  accept: "application/json,.json",
                  onChange: (Me) => {
                    const ke = Me.target.files?.[0];
                    ((Me.target.value = ""), Re(ke));
                  },
                  lang: language("Load saved configuration"),
                  "data-ui-label": "Load saved configuration",
                }),
              ],
            }),
            i.jsx("button", {
              onClick: ht,
              disabled: !(se != null && se.current),
              children: i.jsx(LocalizedText, {
                children: "Download current report",
              }),
            }),
          ],
        }),
        U &&
          H === ce &&
          i.jsx("p", {
            role: "alert",
            className: "materiality-error",
            children: U,
          }),
        X$1 && i.jsx("p", { role: "status", children: X$1 }),
        N &&
          se &&
          i.jsx("p", {
            role: "status",
            className: se.current ? "" : "materiality-stale",
            children: se.current ? se.reason : `Stale result. ${se.reason}`,
          }),
        Ee &&
          i.jsxs("section", {
            className: "materiality-result",
            "aria-label": text("Materiality screen results"),
            lang: language("Materiality screen results"),
            "data-ui-label": "Materiality screen results",
            children: [
              i.jsx("h3", {
                children:
                  Ee.status === "blocked"
                    ? "Blocked — resolve currency exceptions"
                    : Ee.status === "complete-with-exclusions"
                      ? "Complete with exclusions — review every excluded row"
                      : "Screen complete",
              }),
              i.jsxs("p", {
                children: [
                  Ee.request.amountMode === "absolute"
                    ? "Absolute amount"
                    : "Signed amount",
                  " ",
                  comparatorSymbols[Ee.request.comparator],
                  " ",
                  Ee.request.threshold,
                  " ",
                  Ee.request.currency,
                  " ",
                  "· ",
                  Ee.request.effectiveFrom,
                  " ",
                  i.jsx(LocalizedText, { children: "through" }),
                  " ",
                  Ee.request.effectiveTo,
                  i.jsx(LocalizedText, { children: ", inclusive." }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  Ee.counts.population,
                  " ",
                  i.jsx(LocalizedText, { children: "source rows ·" }),
                  " ",
                  Ee.counts.screened,
                  " ",
                  i.jsx(LocalizedText, { children: "screened ·" }),
                  " ",
                  Ee.counts.flagged,
                  " ",
                  i.jsx(LocalizedText, { children: "flagged ·" }),
                  " ",
                  Ee.counts.notFlagged,
                  " ",
                  i.jsx(LocalizedText, { children: "not flagged ·" }),
                  " ",
                  Ee.counts.excluded,
                  " ",
                  i.jsx(LocalizedText, { children: "excluded ·" }),
                  " ",
                  Ee.counts.notScreened,
                  " ",
                  i.jsx(LocalizedText, { children: "not screened" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  Ee.counts.currencyExceptions,
                  " ",
                  i.jsx(LocalizedText, { children: "currency exceptions ·" }),
                  " ",
                  Ee.counts.invalidAmounts,
                  " ",
                  i.jsx(LocalizedText, { children: "invalid amounts ·" }),
                  " ",
                  Ee.counts.invalidDates,
                  " ",
                  i.jsx(LocalizedText, { children: "invalid dates ·" }),
                  " ",
                  Ee.counts.outsidePeriod,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "outside period. Exception categories may overlap.",
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Signed " }),
                  Ee.totals.currency,
                  " ",
                  i.jsx(LocalizedText, {
                    children: "control totals: known source ",
                  }),
                  Ee.totals.knownSourceAmount,
                  i.jsx(LocalizedText, { children: "; eligible" }),
                  " ",
                  Ee.totals.eligibleAmount,
                  i.jsx(LocalizedText, { children: "; excluded known amount" }),
                  " ",
                  Ee.totals.excludedKnownAmount,
                  i.jsx(LocalizedText, { children: "; flagged" }),
                  " ",
                  Ee.totals.flaggedAmount ?? "unavailable (blocked)",
                  ".",
                  " ",
                  Ee.totals.unknownOrOtherCurrencyRows,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "rows have unknown amounts or another/unknown currency and are not added to these totals.",
                  }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(LocalizedText, { children: "Source: " }),
                  Ee.sheetName,
                  "!",
                  Ee.request.range,
                  " ",
                  i.jsx(LocalizedText, { children: "· workbook " }),
                  Ee.workbookId,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  Ee.revision,
                  " · ",
                  Ee.algorithm,
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Rows to review" }),
                  i.jsxs("select", {
                    "aria-label": text("Rows to review"),
                    value: K,
                    onChange: (Me) => {
                      (te(Me.target.value), W(0));
                    },
                    lang: language("Rows to review"),
                    "data-ui-label": "Rows to review",
                    children: [
                      i.jsx("option", {
                        value: "all",
                        children: "All source rows",
                      }),
                      i.jsx("option", {
                        value: "flagged",
                        children: "Flagged",
                      }),
                      i.jsx("option", {
                        value: "not-flagged",
                        children: "Not flagged",
                      }),
                      i.jsx("option", {
                        value: "excluded",
                        children: "Excluded",
                      }),
                      i.jsx("option", {
                        value: "not-screened",
                        children: "Not screened",
                      }),
                    ],
                  }),
                ],
              }),
              i.jsx("div", {
                className: "materiality-scroll",
                children: i.jsxs("table", {
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Source row / range",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Disposition",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Amount / compared",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Currency / date",
                            }),
                          }),
                          i.jsx("th", {
                            scope: "col",
                            children: i.jsx(LocalizedText, {
                              children: "Reasons",
                            }),
                          }),
                          Ee.headers.map((Me, ke) =>
                            i.jsxs(
                              "th",
                              {
                                scope: "col",
                                children: [
                                  Me,
                                  " (",
                                  columnIndexToLetters(
                                    parseRangeAddress(Ee.request.range)
                                      .firstCol + ke,
                                  ),
                                  ")",
                                ],
                              },
                              ke,
                            ),
                          ),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: Ae.slice(Te * 20, Te * 20 + 20).map((Me) =>
                        i.jsxs(
                          "tr",
                          {
                            children: [
                              i.jsxs("th", {
                                scope: "row",
                                children: [Me.sourceRow, " · ", Me.sourceRange],
                              }),
                              i.jsx("td", { children: Me.disposition }),
                              i.jsxs("td", {
                                children: [
                                  Me.amount ?? "Invalid",
                                  " /",
                                  " ",
                                  Me.comparedAmount ?? "Unavailable",
                                ],
                              }),
                              i.jsxs("td", {
                                children: [
                                  Me.currency ?? "Invalid",
                                  " / ",
                                  Me.date ?? "Invalid",
                                ],
                              }),
                              i.jsx("td", {
                                children:
                                  Me.reasons.join(" ") ||
                                  (Me.disposition === "flagged"
                                    ? "Matches the configured comparison."
                                    : "Does not match the configured comparison."),
                              }),
                              Me.cells.map((ke, je) =>
                                i.jsx(
                                  "td",
                                  {
                                    children: i.jsxs("details", {
                                      children: [
                                        i.jsx("summary", {
                                          children: formatCellValueLabel(
                                            Me.values[je],
                                          ),
                                        }),
                                        i.jsxs("p", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Stored: ",
                                            }),
                                            formatCellValueLabel(ke.raw),
                                          ],
                                        }),
                                        i.jsxs("p", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Literal: ",
                                            }),
                                            ke.literal ? "Yes" : "No",
                                          ],
                                        }),
                                        i.jsx("pre", {
                                          children: JSON.stringify(
                                            Object.fromEntries(
                                              Object.entries(ke).filter(
                                                ([$e]) =>
                                                  !["raw", "literal"].includes(
                                                    $e,
                                                  ),
                                              ),
                                            ),
                                            null,
                                            2,
                                          ),
                                        }),
                                      ],
                                    }),
                                  },
                                  je,
                                ),
                              ),
                            ],
                          },
                          Me.sourceRow,
                        ),
                      ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Screen result pages"),
                lang: language("Screen result pages"),
                "data-ui-label": "Screen result pages",
                children: [
                  i.jsx("button", {
                    disabled: Te === 0,
                    onClick: () => W(Te - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous rows",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      Te + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      ye,
                      " · ",
                      Ae.length,
                      " ",
                      i.jsx(LocalizedText, { children: "rows in this view" }),
                    ],
                  }),
                  i.jsx("button", {
                    disabled: Te + 1 >= ye,
                    onClick: () => W(Te + 1),
                    children: i.jsx(LocalizedText, { children: "Next rows" }),
                  }),
                ],
              }),
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children: "Screening rules and limitations",
                    }),
                  }),
                  Ee.notes.map((Me) => i.jsx("p", { children: Me }, Me)),
                ],
              }),
            ],
          }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: e.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const gf = Object.freeze({
    sourceCells: 500,
    rules: 12,
    keyColumns: 8,
    holidays: 1e3,
    reportBytes: 2e6,
    pageSize: 20,
    sequenceMaximum: 999999999999999,
  }),
  nullableColumnIndexSchema = number().int().min(0).max(16383).nullable(),
  keyColumnsSchema = array(number().int().min(0).max(16383)).max(gf.keyColumns),
  ruleIdentitySchema = {
    id: string().regex(/^[A-Za-z][A-Za-z0-9_-]{0,39}$/),
    label: string().trim().min(1).max(80),
  },
  textComparisonSchema = { caseSensitive: boolean(), trimText: boolean() },
  dateSystemSchema = _enum(["1900", "1904"]),
  sequenceValueSchema = number().int().min(0).max(gf.sequenceMaximum),
  screenRuleSchema = discriminatedUnion("kind", [
    object({
      ...ruleIdentitySchema,
      kind: literal("duplicates"),
      keyColumns: keyColumnsSchema,
      ...textComparisonSchema,
      blankKeys: _enum(["fail", "exclude"]),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("sequence"),
      column: nullableColumnIndexSchema,
      groupColumns: keyColumnsSchema,
      ...textComparisonSchema,
      prefix: string().max(40),
      suffix: string().max(40),
      step: sequenceValueSchema.min(1),
      start: sequenceValueSchema.nullable(),
      end: sequenceValueSchema.nullable(),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("round-amount"),
      column: nullableColumnIndexSchema,
      increment: number().finite().positive().max(1e15),
      precision: number().int().min(0).max(6),
      tolerance: number().finite().nonnegative(),
      includeZero: boolean(),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("weekend"),
      column: nullableColumnIndexSchema,
      dateSystem: dateSystemSchema,
      days: array(number().int().min(0).max(6)).min(1).max(7),
    }).strict(),
    object({
      ...ruleIdentitySchema,
      kind: literal("holiday"),
      column: nullableColumnIndexSchema,
      dateSystem: dateSystemSchema,
      calendarName: string().trim().min(1).max(80),
      calendarVersion: string().trim().min(1).max(40),
      dates: array(string().max(10)).max(gf.holidays),
    }).strict(),
  ]),
  screenSettingsSchema = object({
    headers: boolean(),
    comparison: _enum(["raw", "calculated"]),
    excludeBlankRows: boolean(),
    excludedRows: array(number().int().min(1).max(1048576)).max(500),
    rules: array(screenRuleSchema).min(1).max(gf.rules),
    combination: object({
      mode: _enum(["any", "all"]),
      ruleIds: array(string().max(40)).max(gf.rules),
    }).strict(),
  }).strict(),
  sourceRangeSchema = screenSettingsSchema
    .extend({
      sheetId: string().min(1).max(100),
      range: string().regex(
        /^\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?$/,
      ),
    })
    .strict(),
  KW = new WeakMap(),
  rN = (e, r) => `${columnIndexToLetters(r)}${e}`;
function parseIsoCalendarDate(e) {
  if (
    typeof e != "string" ||
    !/^\d{4}-\d{2}-\d{2}$/.test(e) ||
    e.startsWith("0000-")
  )
    return null;
  const r = Date.parse(`${e}T00:00:00.000Z`);
  return Number.isFinite(r) && new Date(r).toISOString().slice(0, 10) === e
    ? r / 864e5
    : null;
}
function convertSerialToIsoDate(e, r) {
  if (typeof e == "string") return parseIsoCalendarDate(e) === null ? null : e;
  if (typeof e != "number" || !Number.isFinite(e)) return null;
  const t = Math.floor(e),
    n = r === "1900" ? 2958465 : 2957003;
  return t < (r === "1900" ? 1 : 0) || t > n || (r === "1900" && t === 60)
    ? null
    : new Date(
        (r === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1)) +
          (r === "1900" && t > 60 ? t - 1 : t) * 864e5,
      )
        .toISOString()
        .slice(0, 10);
}
function toScaledIntegerAmount(e, r) {
  if (!Number.isFinite(e) || Math.abs(e) > 1e15) return null;
  const [t, n = "0"] = Math.abs(e).toString().split("e"),
    o = t.split("."),
    a = BigInt(o.join("")),
    s = r - (o[1]?.length ?? 0) + Number(n);
  if (s >= 0) return a * 10n ** BigInt(s);
  const d = 10n ** BigInt(-s);
  return a % d === 0n ? a / d : null;
}
function ruleKeyColumns(e) {
  return e.kind === "duplicates"
    ? e.keyColumns
    : e.kind === "sequence"
      ? [...(e.column === null ? [] : [e.column]), ...e.groupColumns]
      : e.column === null
        ? []
        : [e.column];
}
function normalizePatternCheckSettings(e, r) {
  const t = sourceRangeSchema.safeParse(r);
  if (!t.success)
    throw Error(
      `Invalid pattern-check settings: ${t.error.issues
        .map((s) => `${s.path.join(".")}: ${s.message}`)
        .join("; ")
        .slice(0, 1500)}`,
    );
  const n = t.data,
    o = parseRangeAddress(n.range),
    a = e.sheets.find((s) => s.id === n.sheetId);
  ((!a || o.lastRow >= a.rowCount || o.lastCol >= a.colCount) &&
    throwErrorMessage(
      "Pattern source is outside current worksheet dimensions.",
    ),
    (o.lastRow - o.firstRow + 1) * (o.lastCol - o.firstCol + 1) >
      gf.sourceCells &&
      throwErrorMessage(
        "Pattern checks support at most 500 selected source cells.",
      ),
    new Set(n.rules.map((s) => s.id)).size !== n.rules.length &&
      throwErrorMessage("Every rule needs a unique ID."),
    (new Set(n.excludedRows).size !== n.excludedRows.length ||
      n.excludedRows.some(
        (s) => s <= o.firstRow + Number(n.headers) || s > o.lastRow + 1,
      )) &&
      throwErrorMessage(
        "Explicit exclusions must be distinct data-row numbers inside the selected range.",
      ));
  for (const s of n.rules) {
    const d = ruleKeyColumns(s);
    if (
      (new Set(d).size !== d.length &&
        throwErrorMessage(`Rule ${s.id} has duplicate field mappings.`),
      s.kind === "sequence" &&
        s.start !== null &&
        s.end !== null &&
        (s.start > s.end || (s.end - s.start) % s.step !== 0) &&
        throwErrorMessage(
          "Sequence bounds must be chronological integers on the declared step.",
        ),
      s.kind === "round-amount")
    ) {
      const u = toScaledIntegerAmount(s.increment, s.precision),
        l = toScaledIntegerAmount(s.tolerance, s.precision);
      (u === null || u === 0n || l === null || l * 2n >= u) &&
        throwErrorMessage(
          "Round-amount increment/tolerance must fit the declared precision; tolerance must be less than half the increment.",
        );
    }
    (s.kind === "weekend" &&
      new Set(s.days).size !== s.days.length &&
      throwErrorMessage(
        "Weekend weekdays must be distinct (0=Sunday through 6=Saturday).",
      ),
      s.kind === "holiday" &&
        (new Set(s.dates).size !== s.dates.length ||
          s.dates.some((u) => parseIsoCalendarDate(u) === null)) &&
        throwErrorMessage(
          "Holiday dates must be distinct valid ISO calendar dates.",
        ));
  }
  return (
    (new Set(n.combination.ruleIds).size !== n.combination.ruleIds.length ||
      n.combination.ruleIds.some(
        (s) => !n.rules.some((d) => d.id === s && d.kind !== "sequence"),
      )) &&
      throwErrorMessage(
        "Row combinations require distinct existing row rules. Sequence gaps are separate interval findings, not missing source rows.",
      ),
    {
      ...n,
      range: `${rN(o.firstRow + 1, o.firstCol)}:${rN(o.lastRow + 1, o.lastCol)}`,
    }
  );
}
const createEmptyCounts = (e) => ({
  population: e,
  excluded: 0,
  invalid: 0,
  tested: 0,
  matched: 0,
  notMatched: 0,
  notTested: 0,
  groups: 0,
  findings: 0,
  gapValues: 0,
});
function summarizeRuleOutcome(e, r) {
  const t = createEmptyCounts(e.length);
  for (const o of e)
    o.disposition === "excluded"
      ? t.excluded++
      : o.disposition === "invalid"
        ? t.invalid++
        : o.disposition === "matched"
          ? t.matched++
          : o.disposition === "not-matched"
            ? t.notMatched++
            : t.notTested++;
  ((t.tested = t.matched + t.notMatched),
    (t.groups = new Set(r.map((o) => o.group.id)).size),
    (t.findings = r.length));
  const n = r.reduce((o, a) => {
    return o + BigInt(a.missing?.count ?? 0);
  }, 0n);
  return (
    n > BigInt(Number.MAX_SAFE_INTEGER) &&
      throwErrorMessage(
        "The total missing-value count exceeds the exact integer reporting limit. Split sequence groups into smaller checks.",
      ),
    (t.gapValues = Number(n)),
    t
  );
}
function buildRowMatchKey(e, r) {
  const t = e.map((n) =>
    typeof n == "string"
      ? r.caseSensitive
        ? r.trimText
          ? n.trim()
          : n
        : (r.trimText ? n.trim() : n).toLowerCase()
      : n,
  );
  return t.some(
    (n) =>
      isNullOrEmptyValue(n) || (typeof n == "number" && !Number.isFinite(n)),
  )
    ? null
    : { key: JSON.stringify(t.map((n) => [typeof n, n])), values: t };
}
function runPatternChecks(e, r, t = {}) {
  const n = normalizePatternCheckSettings(e, r),
    o = parseRangeAddress(n.range),
    a = { sheetId: n.sheetId, range: n.range, headers: n.headers, columns: [] };
  let s = buildMaterialitySourceEvidence(e, a, t),
    d = "";
  if (n.comparison === "calculated") {
    const y = [...new Set(n.rules.flatMap(ruleKeyColumns))].filter(
      (w) => w >= o.firstCol && w <= o.lastCol,
    );
    try {
      s = buildMaterialitySourceEvidence(e, { ...a, columns: y }, t);
    } catch (w) {
      d =
        w instanceof Error
          ? w.message
          : "Selected fields could not be calculated.";
    }
  }
  const u = s.rows.map((y) => ({
      ...y,
      origin: findRowOrigin(e, s.sheetId, y.sourceRow - 1),
      exclusionReason: n.excludedRows.includes(y.sourceRow)
        ? "Explicitly excluded source row."
        : n.excludeBlankRows && y.values.every(isNullOrEmptyValue)
          ? "Entire selected row is blank; excluded under the visible blank-row policy."
          : null,
    })),
    l = {
      workbookId: e.id,
      revision: e.revision,
      sheetId: s.sheetId,
      sheetName: s.sheetName,
      range: n.range,
      sha256: sha256HexDigest(
        new TextEncoder().encode(
          JSON.stringify({ headers: s.headers, rows: u }),
        ),
      ),
    },
    c = new Map(u.map((y) => [y.sourceRow, y])),
    h = (y, w) =>
      y.flatMap((E) =>
        w.map((A) => {
          const k = c.get(E);
          return {
            address: rN(E, A),
            raw: k.cells[A - o.firstCol].raw,
            value: k.values[A - o.firstCol],
            origin: k.origin,
          };
        }),
      ),
    f = () =>
      u.map((y) => ({
        sourceRow: y.sourceRow,
        disposition: y.exclusionReason ? "excluded" : "not-tested",
        reasons: y.exclusionReason ? [y.exclusionReason] : [],
        findingIds: [],
      })),
    p = n.rules.map((y) => {
      const w = {
          rule: cloneStructuredValue(y),
          version: "1.0.0",
          status: "complete",
          messages: [],
          counts: createEmptyCounts(u.length),
          rows: f(),
          findings: [],
        },
        E = new Map(w.rows.map((C) => [C.sourceRow, C])),
        A = ruleKeyColumns(y),
        k = A.every((C) => C >= o.firstCol && C <= o.lastCol),
        F = y.kind === "duplicates" ? !A.length : y.column === null,
        I = () => {
          if (w.status === "failed") {
            w.findings = [];
            for (const C of w.rows)
              C.disposition !== "excluded" &&
                C.disposition !== "invalid" &&
                ((C.disposition = "not-tested"),
                (C.findingIds = []),
                C.reasons.push(
                  "Rule failed; no clean or exception conclusion is available.",
                ));
          }
          w.counts = summarizeRuleOutcome(w.rows, w.findings);
          for (const C of w.findings) C.counts = cloneStructuredValue(w.counts);
          return (
            w.status === "complete" &&
              !w.counts.tested &&
              ((w.status = "not-applicable"),
              w.messages.push(
                "No eligible records or sequence groups could be tested.",
              )),
            w
          );
        };
      if (F || !k || d)
        return (
          (w.status = "failed"),
          w.messages.push(
            d ||
              "Missing field mapping: select all required columns inside the source range.",
          ),
          I()
        );
      if (y.kind === "holiday" && !y.dates.length)
        return (
          (w.status = "not-applicable"),
          w.messages.push(
            "No holiday dates supplied. A calendar name alone is not a holiday test.",
          ),
          I()
        );
      const O = (C, T) => {
          const N = E.get(C);
          ((N.disposition = "invalid"),
            N.reasons.push(T),
            (w.status = "failed"));
        },
        j = u.filter((C) => !C.exclusionReason),
        z = new Map(),
        V = (C, T, N) => {
          let R = z.get(C);
          (R || ((R = { key: T, members: [] }), z.set(C, R)),
            R.members.push(N));
        },
        D = new Map(),
        B = new Set(y.kind === "holiday" ? y.dates : []);
      for (const C of j) {
        const T = E.get(C.sourceRow),
          N = (R) => C.values[R - o.firstCol];
        if (y.kind === "duplicates") {
          const R = y.keyColumns.map(N),
            L = buildRowMatchKey(R, y);
          if (!L) {
            y.blankKeys === "exclude" &&
            R.some(
              ($) =>
                isNullOrEmptyValue($) ||
                (typeof $ == "string" && y.trimText && !$.trim()),
            )
              ? ((T.disposition = "excluded"),
                T.reasons.push(
                  "Blank duplicate key excluded under this rule policy.",
                ))
              : O(
                  C.sourceRow,
                  "Duplicate key is blank or has an invalid value.",
                );
            continue;
          }
          ((T.disposition = "not-matched"), V(L.key, L.values, C.sourceRow));
        } else if (y.kind === "sequence") {
          const R = N(y.column),
            L = y.groupColumns.length
              ? buildRowMatchKey(y.groupColumns.map(N), y)
              : { key: "[]", values: [] };
          let $ = null;
          if (
            (typeof R == "number" &&
              Number.isSafeInteger(R) &&
              R >= 0 &&
              R <= gf.sequenceMaximum &&
              !y.prefix &&
              !y.suffix &&
              ($ = R),
            typeof R == "string" &&
              R.startsWith(y.prefix) &&
              R.endsWith(y.suffix))
          ) {
            const U = R.slice(
              y.prefix.length,
              y.suffix ? R.length - y.suffix.length : undefined,
            );
            /^\d{1,15}$/.test(U) && ($ = Number(U));
          }
          if ($ === null || !L) {
            O(
              C.sourceRow,
              "Sequence requires a nonnegative integer or matching prefix/digit/suffix text, and nonblank grouping keys.",
            );
            continue;
          }
          if (
            (y.start !== null && $ < y.start) ||
            (y.end !== null && $ > y.end)
          ) {
            O(
              C.sourceRow,
              "Sequence value is outside the declared start/end bounds.",
            );
            continue;
          }
          (D.set(C.sourceRow, $), V(L.key, L.values, C.sourceRow));
        } else if (y.kind === "round-amount") {
          const R = N(y.column),
            L =
              typeof R == "number"
                ? toScaledIntegerAmount(R, y.precision)
                : null;
          if (L === null) {
            O(
              C.sourceRow,
              "Amount must be a finite number of at most 1e15 magnitude, exactly representable at the declared decimal precision; numeric text is not coerced.",
            );
            continue;
          }
          if (((T.disposition = "not-matched"), L === 0n && !y.includeZero))
            continue;
          const $ = toScaledIntegerAmount(y.increment, y.precision),
            U = toScaledIntegerAmount(y.tolerance, y.precision),
            G = L % $;
          (G < $ - G ? G : $ - G) <= U &&
            V(JSON.stringify(R), [R], C.sourceRow);
        } else {
          const R = convertSerialToIsoDate(N(y.column), y.dateSystem);
          if (!R) {
            O(
              C.sourceRow,
              "Date must be strict ISO YYYY-MM-DD or a valid Excel date/time serial in the chosen system; serial 60 in the 1900 system is invalid.",
            );
            continue;
          }
          T.disposition = "not-matched";
          const L = new Date(parseIsoCalendarDate(R) * 864e5).getUTCDay();
          (y.kind === "weekend" ? y.days.includes(L) : B.has(R)) &&
            V(R, [R], C.sourceRow);
        }
      }
      if (w.status === "failed")
        return (
          w.messages.push(
            "Invalid field values prevent this rule from completing. Inspect the invalid-row ledger; other rows were not treated as a clean population.",
          ),
          I()
        );
      const _ = (C, T, N, R, L, $) => {
        const U = `${y.id}-G${w.findings.length + 1}`,
          G = {
            id: U,
            classification: "review-candidate",
            title: R,
            detail:
              L +
              " Review candidate only; this is not proof of error or fraud.",
            rule: { id: y.id, kind: y.kind, version: "1.0.0" },
            parameters: cloneStructuredValue(y),
            source: cloneStructuredValue(l),
            sourceCells: h(N, A),
            group: { id: U, key: C, memberRows: [...T] },
            counts: createEmptyCounts(u.length),
            ...($ ? { missing: $ } : {}),
          };
        if ((w.findings.push(G), y.kind !== "sequence"))
          for (const H of T)
            ((E.get(H).disposition = "matched"), E.get(H).findingIds.push(U));
        else for (const H of N) E.get(H).findingIds.push(U);
      };
      if (y.kind === "sequence") {
        let C = 0;
        for (const T of z.values()) {
          const N = new Map();
          for (const Q of T.members) {
            const X = D.get(Q);
            N.set(X, [...(N.get(X) ?? []), Q]);
          }
          const R = [...N.keys()].sort((Q, X) => Q - X),
            L = y.start ?? R[0],
            $ = y.end ?? R.at(-1);
          if (R.some((Q) => (Q - L) % y.step !== 0) || ($ - L) % y.step !== 0) {
            for (const Q of T.members)
              O(
                Q,
                "Sequence values or bounds do not align with the declared step.",
              );
            continue;
          }
          if (R.length < 2 && (y.start === null || y.end === null)) {
            C++;
            for (const Q of T.members)
              E.get(Q).reasons.push(
                "One distinct sequence value without both bounds cannot establish missing values.",
              );
            continue;
          }
          for (const Q of T.members) E.get(Q).disposition = "not-matched";
          let U = null,
            G = L;
          const H = (Q, X, J, q) => {
            if (Q > X) return;
            const W = (X - Q) / y.step + 1,
              K = [
                ...(J === null ? [] : (N.get(J) ?? [])),
                ...(q === null ? [] : (N.get(q) ?? [])),
              ],
              te = Array.from(
                { length: Math.min(20, W) },
                (ae, oe) => y.prefix + String(Q + oe * y.step) + y.suffix,
              );
            _(
              T.key,
              T.members,
              K,
              `Missing sequence ${Q === X ? Q : `${Q}–${X}`} (${W})`,
              "Missing interval under the declared bounds/step. Cited existing cells are boundary evidence, not erroneous rows.",
              {
                first: Q,
                last: X,
                step: y.step,
                count: W,
                examples: te,
                examplesTruncated: W > 20,
              },
            );
          };
          for (const Q of R)
            (H(G, Q - y.step, U, Q), (U = Q), (G = Q + y.step));
          H(G, $, U, null);
        }
        (C &&
          w.messages.push(
            `${C} group(s) lack enough distinct values/bounds; their rows are not tested. No unseen groups or leading/trailing gaps are inferred without bounds.`,
          ),
          w.messages.push(
            "Gaps are missing intervals, never fabricated source rows. Duplicate sequence values are collapsed for gap detection; add a duplicate rule to review multiplicity.",
          ));
      } else
        for (const C of z.values())
          (y.kind === "duplicates" && C.members.length < 2) ||
            _(
              C.key,
              C.members,
              C.members,
              `${y.label}: ${C.members.length} row(s)`,
              y.kind === "duplicates"
                ? "All members of the typed matching-key group are included."
                : y.kind === "round-amount"
                  ? "Amount meets the declared increment, decimal precision, tolerance and zero policy."
                  : `Posting calendar date ${C.key[0]} matches the supplied ${y.kind === "holiday" ? "holiday calendar" : "weekend weekdays"}.`,
            );
      return I();
    }),
    m = {
      status: "complete",
      messages: [],
      counts: createEmptyCounts(u.length),
      rows: f(),
      findings: [],
    },
    v = p.filter((y) => n.combination.ruleIds.includes(y.rule.id));
  if (!v.length)
    ((m.status = "not-applicable"),
      m.messages.push(
        "No row combination configured. Sequence gaps remain independent interval findings.",
      ));
  else if (v.some((y) => y.status !== "complete"))
    ((m.status = v.some((y) => y.status === "failed")
      ? "failed"
      : "not-applicable"),
      m.messages.push(
        "Every selected rule must complete before a combined result is available.",
      ));
  else
    for (const y of m.rows) {
      if (y.disposition === "excluded") continue;
      const w = v.map((I) => ({
        result: I,
        row: I.rows.find((O) => O.sourceRow === y.sourceRow),
      }));
      if (
        w.some((I) => !["matched", "not-matched"].includes(I.row.disposition))
      ) {
        y.reasons.push(
          "Combination uses only rows tested by every selected rule (complete cases).",
        );
        continue;
      }
      const E = w.filter((I) => I.row.disposition === "matched"),
        A = n.combination.mode === "any" ? E.length > 0 : E.length === v.length;
      if (((y.disposition = A ? "matched" : "not-matched"), !A)) continue;
      const k = `combined-R${y.sourceRow}`;
      y.findingIds.push(k);
      const F = [...new Set(v.flatMap((I) => ruleKeyColumns(I.rule)))];
      m.findings.push({
        id: k,
        classification: "review-candidate",
        title: `Combined review candidate — row ${y.sourceRow}`,
        detail: `${n.combination.mode === "all" ? "All" : "At least one"} selected row rule matched. Review candidate only; not proof of error or fraud.`,
        rule: { id: "combined", kind: n.combination.mode, version: "1.0.0" },
        parameters: cloneStructuredValue(n.combination),
        source: cloneStructuredValue(l),
        sourceCells: h([y.sourceRow], F),
        group: { id: k, key: [y.sourceRow], memberRows: [y.sourceRow] },
        counts: createEmptyCounts(u.length),
        matchedRules: E.map((I) => ({
          ruleId: I.result.rule.id,
          findingIds: [...I.row.findingIds],
        })),
      });
    }
  ((m.counts = summarizeRuleOutcome(m.rows, m.findings)),
    m.status === "complete" &&
      !m.counts.tested &&
      ((m.status = "not-applicable"),
      m.messages.push("No row was eligible for every selected check.")));
  for (const y of m.findings) y.counts = cloneStructuredValue(m.counts);
  const g = {
    population: u.length,
    excluded: u.filter((y) => y.exclusionReason).length,
    eligible: u.filter((y) => !y.exclusionReason).length,
    rulesComplete: p.filter((y) => y.status === "complete").length,
    rulesFailed: p.filter((y) => y.status === "failed").length,
    rulesNotApplicable: p.filter((y) => y.status === "not-applicable").length,
    findings: p.reduce((y, w) => y + w.findings.length, 0),
    gapValues: p.reduce((y, w) => y + w.counts.gapValues, 0),
  };
  Number.isSafeInteger(g.gapValues) ||
    throwErrorMessage(
      "The aggregate missing-value count exceeds the exact integer reporting limit. Run fewer sequence rules.",
    );
  const b = {
    schemaVersion: 1,
    kind: "pattern-audit",
    algorithm: "ledgerglass-patterns-1",
    generatedAt: new Date().toISOString(),
    request: n,
    context: s.context,
    source: l,
    status: g.rulesFailed
      ? "failed"
      : g.rulesComplete
        ? "complete"
        : "not-applicable",
    messages: d
      ? [
          "Calculation failed. Evidence contains raw fields; no rule conclusion is available.",
          d,
        ]
      : [],
    headers: s.headers,
    rows: u,
    counts: g,
    results: p,
    combination: m,
    notes: [
      "Review candidates are not proof of error or fraud. No cells are edited, records removed or automatic totals excluded.",
      "Population includes every selected data-row position. Only the visible blank-row/explicit-row and rule blank-key policies exclude records. Hidden/filter content is included only when explicitly authorized; it is never silently skipped.",
      "Typed keys keep text and numbers distinct. Numeric text is not coerced for amount tests. Round amounts use exact decimal units of the stored number; no currency conversion, floating epsilon or silent rounding is applied.",
      "Dates use calendar days, independent of browser timezone. Fractional Excel serials retain their posting calendar day; ambiguous locale/timestamp strings must be converted explicitly. Holiday calendars are user supplied, not a current official jurisdiction calendar.",
      "Calculated mode validates mapped formula fields across the whole selected evidence, including explicitly excluded rows. Formula errors or unsupported dependency closure fail the run; select raw mode only when intentionally testing stored expressions. Sequence examples do not infer zero-padding width.",
      "Row combinations use the intersection of evaluable rows. Sequence gaps are independent intervals with observed boundary citations; examples may be abbreviated, but counts and full intervals are complete.",
      "Reports are current-snapshot evidence, not an immutable audit service. Store original source and the exported report under host access controls; native workbook JSON does not automatically retain a separately downloaded report.",
    ],
  };
  return (
    new TextEncoder().encode(JSON.stringify(b, null, 2)).length >
      gf.reportBytes &&
      throwErrorMessage(
        "The complete pattern report exceeds 2,000,000 bytes. Select fewer cells/rules; results are never silently truncated.",
      ),
    KW.set(b, {
      snapshot: JSON.stringify(e),
      request: JSON.stringify(n),
      context: JSON.stringify(s.context),
      content: JSON.stringify(b),
    }),
    b
  );
}
function checkReportCurrency(e, r, t, n = {}) {
  const o = KW.get(r),
    a = (s) => ({ current: false, disclose: false, reason: s });
  if (!o || o.content !== JSON.stringify(r))
    return a(
      "This report is untrusted or changed. Run the pattern checks again.",
    );
  try {
    const s = normalizePatternCheckSettings(e, t),
      d = buildMaterialitySourceEvidence(
        e,
        { sheetId: s.sheetId, range: s.range, headers: s.headers, columns: [] },
        n,
      );
    return o.snapshot !== JSON.stringify(e) ||
      o.context !== JSON.stringify(d.context) ||
      s.sheetId !== r.request.sheetId ||
      s.range !== r.request.range
      ? a(
          "Source or authorized scope changed. Previous evidence is hidden; run checks again.",
        )
      : o.request !== JSON.stringify(s)
        ? {
            current: false,
            disclose: true,
            reason:
              "Settings changed. This is a previous result; run checks again before export.",
          }
        : {
            current: true,
            disclose: true,
            reason: "Current source and rule configuration.",
          };
  } catch {
    return a(
      "Source, settings or authorization are no longer valid. Run checks again.",
    );
  }
}
function exportScreenReportJson(e, r, t, n = {}) {
  const o = checkReportCurrency(e, r, t, n);
  return (o.current || throwErrorMessage(o.reason), JSON.stringify(r, null, 2));
}
function parseScreenReportRequest(e, r) {
  const t = r;
  return (
    (!t || t.schemaVersion !== 1 || t.kind !== "pattern-audit") &&
      throwErrorMessage("Choose a pattern-audit configuration or report."),
    normalizePatternCheckSettings(e, t.request)
  );
}
const screenRuleLabels = {
    duplicates: "Duplicate keys",
    sequence: "Sequence gaps",
    "round-amount": "Round amounts",
    weekend: "Weekend postings",
    holiday: "Holiday postings",
  },
  weekdayNames = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ],
  buildScreenRule = (e, r) => {
    const t = { id: r, label: screenRuleLabels[e] };
    switch (e) {
      case "duplicates":
        return {
          ...t,
          kind: e,
          keyColumns: [],
          caseSensitive: false,
          trimText: true,
          blankKeys: "fail",
        };
      case "sequence":
        return {
          ...t,
          kind: e,
          column: null,
          groupColumns: [],
          caseSensitive: false,
          trimText: true,
          prefix: "",
          suffix: "",
          step: 1,
          start: null,
          end: null,
        };
      case "round-amount":
        return {
          ...t,
          kind: e,
          column: null,
          increment: 100,
          precision: 2,
          tolerance: 0,
          includeZero: false,
        };
      case "weekend":
        return {
          ...t,
          kind: e,
          column: null,
          dateSystem: "1900",
          days: [0, 6],
        };
      case "holiday":
        return {
          ...t,
          kind: e,
          column: null,
          dateSystem: "1900",
          calendarName: "User-supplied calendar",
          calendarVersion: "1",
          dates: [],
        };
    }
  },
  formatScreenCountsSummary = (e) =>
    `${e.population} population · ${e.excluded} excluded · ${e.invalid} invalid · ${e.tested} tested · ${e.matched} matched · ${e.notMatched} not matched · ${e.notTested} not tested · ${e.findings} findings · ${e.gapValues} missing sequence values`,
  nanToEmptyText = (e) => (e === null || Number.isNaN(e) ? "" : e);

export {
  formatScreenCountsSummary,
  exportScreenReportJson,
  nanToEmptyText,
  MaterialityRiskScreen,
  runPatternChecks,
  weekdayNames,
  checkReportCurrency,
  buildScreenRule,
  gf,
  normalizePatternCheckSettings,
  screenRuleLabels,
  parseScreenReportRequest,
};
