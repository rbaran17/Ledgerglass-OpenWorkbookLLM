import { listRangeColumnIndexes } from "./21-aging-and-materiality-screens.js";
import { throwErrorMessage, isNullOrEmptyValue } from "./03-workbook-model.js";
import { cloneStructuredValue } from "./11-local-workbook-store-and-versions.js";

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
import { sha256HexDigest } from "./01-schemas-date-functions-and-validation.js";
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
} from "./07-file-import-and-worker-dispatch.js";
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
} from "./49-review-tools-and-outlines.js";

function MaterialityRiskScreen(e) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(""),
    d = ne.useRef(0),
    u = ne.useRef(true),
    rangeColumnIndexes = listRangeColumnIndexes(e.range),
    [range, h] = ne.useState(e.range),
    [f, p] = ne.useState(true),
    [amountColumn, v] = ne.useState(rangeColumnIndexes[0] ?? 0),
    [currencyColumn, b] = ne.useState(rangeColumnIndexes[1] ?? 1),
    [dateColumn, w] = ne.useState(rangeColumnIndexes[2] ?? 2),
    [currency, A] = ne.useState("USD"),
    [k, F] = ne.useState("50000"),
    [amountMode, O] = ne.useState("absolute"),
    [comparator, z] = ne.useState("gte"),
    [effectiveFrom, D] = ne.useState("1900-01-01"),
    [effectiveTo, _] = ne.useState("9999-12-31"),
    [checked, T] = ne.useState(false),
    [N, R] = ne.useState(null),
    [L, $] = ne.useState(""),
    [U, G] = ne.useState(""),
    [H, Q] = ne.useState(""),
    [X$1, J] = ne.useState(""),
    [q, W] = ne.useState(0),
    [K, te] = ne.useState("all"),
    ae = {
      sheetId: e.sheetId,
      range: range,
      headers: f,
      amountColumn: amountColumn,
      currencyColumn: currencyColumn,
      dateColumn: dateColumn,
      currency: currency,
      threshold: k.trim() === "" ? NaN : Number(k),
      amountMode: amountMode,
      comparator: comparator,
      effectiveFrom: effectiveFrom,
      effectiveTo: effectiveTo,
    },
    oe = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && checked,
      allowedSheetIds: e.allowedSheetIds,
    },
    de = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    he = [
      ...new Set(
        e.allowedSheetIds ?? e.workbook.sheets.map((sheet) => sheet.id),
      ),
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
    we = listRangeColumnIndexes(range),
    Ee = se?.disclose ? N : null,
    Ae = Ee?.rows.filter((row) => K === "all" || row.disposition === K) ?? [],
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
    u.current = true;
    const activeElement = document.activeElement;
    return (
      a.current?.querySelector('[data-ui-label="Source range"]')?.focus(),
      () => {
        ((u.current = false),
          d.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []);
  const onClick = () => {
      d.current++;
      try {
        const Me = runMaterialityScreen(e.workbook, ae, oe);
        (R(Me), $(le), G(""), J(""), W(0), te("all"));
      } catch (error) {
        (R(null), Pe(error));
      }
    },
    ht = () => {
      try {
        if (!N || !se?.current)
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
      } catch (error) {
        Pe(error);
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
        const Je = await importNativeJsonFile(Me, cancellableOperation.begin());
        if (!$e()) return;
        const materialityConfiguration = loadMaterialityConfiguration(
          e.workbook,
          Je,
          oe,
        );
        if (materialityConfiguration.sheetId !== e.sheetId)
          throw Error(
            "Open the saved configuration's source worksheet before loading it.",
          );
        (h(materialityConfiguration.range),
          p(materialityConfiguration.headers),
          v(materialityConfiguration.amountColumn),
          b(materialityConfiguration.currencyColumn),
          w(materialityConfiguration.dateColumn),
          A(materialityConfiguration.currency),
          F(String(materialityConfiguration.threshold)),
          O(materialityConfiguration.amountMode),
          z(materialityConfiguration.comparator),
          D(materialityConfiguration.effectiveFrom),
          _(materialityConfiguration.effectiveTo),
          R(null),
          G(""),
          J(
            "Configuration loaded. Run the screen against the current source; saved conclusions are not reused.",
          ));
      } catch (error) {
        $e() && Pe(toFileProcessingError(error));
      }
    };
  return i.jsx("div", {
    className: "materiality-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "materiality-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${o}-title`,
      ref: a,
      onKeyDown: (event) => {
        var je, $e;
        if (
          (event.stopPropagation(),
          event.key === "Escape" && (event.preventDefault(), e.onClose()),
          event.key === "Tab")
        ) {
          const Je = Array.from(
            a.current?.querySelectorAll(
              'button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex="0"]',
            ) ?? [],
          ).filter(
            (ze) =>
              !ze.closest("details:not([open])") || ze.tagName === "SUMMARY",
          );
          event.shiftKey && document.activeElement === Je[0]
            ? (event.preventDefault(), (je = Je.at(-1)) == null || je.focus())
            : !event.shiftKey &&
              document.activeElement === Je.at(-1) &&
              (event.preventDefault(), ($e = Je[0]) == null || $e.focus());
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
                  value: Se ? range : "",
                  onChange: (event) => h(event.target.value),
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
                  onChange: (event) => p(event.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            [
              ["Amount column", amountColumn, v],
              ["Currency column", currencyColumn, b],
              ["Date column", dateColumn, w],
            ].map(([Me, ke, je]) =>
              i.jsxs(
                "label",
                {
                  children: [
                    Me,
                    i.jsx("select", {
                      "aria-label": Me,
                      value: ke,
                      onChange: (event) => je(Number(event.target.value)),
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
                  value: currency,
                  maxLength: 3,
                  onChange: (event) => A(event.target.value),
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
                  value: amountMode,
                  onChange: (event) => O(event.target.value),
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
                  value: comparator,
                  onChange: (event) => z(event.target.value),
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
                  value: effectiveFrom,
                  onChange: (event) => D(event.target.value),
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
                  value: effectiveTo,
                  onChange: (event) => _(event.target.value),
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
                checked: checked,
                onChange: (event) => T(event.target.checked),
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
              onClick: onClick,
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
                  onChange: (event) => {
                    const ke = event.target.files?.[0];
                    ((event.target.value = ""), Re(ke));
                  },
                  lang: language("Load saved configuration"),
                  "data-ui-label": "Load saved configuration",
                }),
              ],
            }),
            i.jsx("button", {
              onClick: ht,
              disabled: !se?.current,
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
                    onChange: (event) => {
                      (te(event.target.value), W(0));
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
                          Ee.headers.map((header, ke) =>
                            i.jsxs(
                              "th",
                              {
                                scope: "col",
                                children: [
                                  header,
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
                              Me.cells.map((cell, je) =>
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
                                            formatCellValueLabel(cell.raw),
                                          ],
                                        }),
                                        i.jsxs("p", {
                                          children: [
                                            i.jsx(LocalizedText, {
                                              children: "Literal: ",
                                            }),
                                            cell.literal ? "Yes" : "No",
                                          ],
                                        }),
                                        i.jsx("pre", {
                                          children: JSON.stringify(
                                            Object.fromEntries(
                                              Object.entries(cell).filter(
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
                  Ee.notes.map((note) => i.jsx("p", { children: note }, note)),
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
function convertSerialToIsoDate(e, dateSystem) {
  if (typeof e == "string") return parseIsoCalendarDate(e) === null ? null : e;
  if (typeof e != "number" || !Number.isFinite(e)) return null;
  const t = Math.floor(e),
    n = dateSystem === "1900" ? 2958465 : 2957003;
  return t < (dateSystem === "1900" ? 1 : 0) ||
    t > n ||
    (dateSystem === "1900" && t === 60)
    ? null
    : new Date(
        (dateSystem === "1900"
          ? Date.UTC(1899, 11, 31)
          : Date.UTC(1904, 0, 1)) +
          (dateSystem === "1900" && t > 60 ? t - 1 : t) * 864e5,
      )
        .toISOString()
        .slice(0, 10);
}
function toScaledIntegerAmount(e, precision) {
  if (!Number.isFinite(e) || Math.abs(e) > 1e15) return null;
  const [t, n = "0"] = Math.abs(e).toString().split("e"),
    o = t.split("."),
    a = BigInt(o.join("")),
    s = precision - (o[1]?.length ?? 0) + Number(n);
  if (s >= 0) return a * 10n ** BigInt(s);
  const d = 10n ** BigInt(-s);
  return a % d === 0n ? a / d : null;
}
function ruleKeyColumns(rule) {
  return rule.kind === "duplicates"
    ? rule.keyColumns
    : rule.kind === "sequence"
      ? [...(rule.column === null ? [] : [rule.column]), ...rule.groupColumns]
      : rule.column === null
        ? []
        : [rule.column];
}
function normalizePatternCheckSettings(workbook, r) {
  const t = sourceRangeSchema.safeParse(r);
  if (!t.success)
    throw Error(
      `Invalid pattern-check settings: ${t.error.issues
        .map((issue) => `${issue.path.join(".")}: ${issue.message}`)
        .join("; ")
        .slice(0, 1500)}`,
    );
  const data = t.data,
    rangeAddress = parseRangeAddress(data.range),
    a = workbook.sheets.find((sheet) => sheet.id === data.sheetId);
  ((!a ||
    rangeAddress.lastRow >= a.rowCount ||
    rangeAddress.lastCol >= a.colCount) &&
    throwErrorMessage(
      "Pattern source is outside current worksheet dimensions.",
    ),
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      gf.sourceCells &&
      throwErrorMessage(
        "Pattern checks support at most 500 selected source cells.",
      ),
    new Set(data.rules.map((rule) => rule.id)).size !== data.rules.length &&
      throwErrorMessage("Every rule needs a unique ID."),
    (new Set(data.excludedRows).size !== data.excludedRows.length ||
      data.excludedRows.some(
        (excludedRow) =>
          excludedRow <= rangeAddress.firstRow + Number(data.headers) ||
          excludedRow > rangeAddress.lastRow + 1,
      )) &&
      throwErrorMessage(
        "Explicit exclusions must be distinct data-row numbers inside the selected range.",
      ));
  for (const rule of data.rules) {
    const d = ruleKeyColumns(rule);
    if (
      (new Set(d).size !== d.length &&
        throwErrorMessage(`Rule ${rule.id} has duplicate field mappings.`),
      rule.kind === "sequence" &&
        rule.start !== null &&
        rule.end !== null &&
        (rule.start > rule.end || (rule.end - rule.start) % rule.step !== 0) &&
        throwErrorMessage(
          "Sequence bounds must be chronological integers on the declared step.",
        ),
      rule.kind === "round-amount")
    ) {
      const scaledIntegerAmount = toScaledIntegerAmount(
          rule.increment,
          rule.precision,
        ),
        l = toScaledIntegerAmount(rule.tolerance, rule.precision);
      (scaledIntegerAmount === null ||
        scaledIntegerAmount === 0n ||
        l === null ||
        l * 2n >= scaledIntegerAmount) &&
        throwErrorMessage(
          "Round-amount increment/tolerance must fit the declared precision; tolerance must be less than half the increment.",
        );
    }
    (rule.kind === "weekend" &&
      new Set(rule.days).size !== rule.days.length &&
      throwErrorMessage(
        "Weekend weekdays must be distinct (0=Sunday through 6=Saturday).",
      ),
      rule.kind === "holiday" &&
        (new Set(rule.dates).size !== rule.dates.length ||
          rule.dates.some((date) => parseIsoCalendarDate(date) === null)) &&
        throwErrorMessage(
          "Holiday dates must be distinct valid ISO calendar dates.",
        ));
  }
  return (
    (new Set(data.combination.ruleIds).size !==
      data.combination.ruleIds.length ||
      data.combination.ruleIds.some(
        (ruleId) =>
          !data.rules.some(
            (rule) => rule.id === ruleId && rule.kind !== "sequence",
          ),
      )) &&
      throwErrorMessage(
        "Row combinations require distinct existing row rules. Sequence gaps are separate interval findings, not missing source rows.",
      ),
    {
      ...data,
      range: `${rN(rangeAddress.firstRow + 1, rangeAddress.firstCol)}:${rN(rangeAddress.lastRow + 1, rangeAddress.lastCol)}`,
    }
  );
}
const createEmptyCounts = (population) => ({
  population: population,
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
function summarizeRuleOutcome(rows, findings) {
  const emptyCounts = createEmptyCounts(rows.length);
  for (const row of rows)
    row.disposition === "excluded"
      ? emptyCounts.excluded++
      : row.disposition === "invalid"
        ? emptyCounts.invalid++
        : row.disposition === "matched"
          ? emptyCounts.matched++
          : row.disposition === "not-matched"
            ? emptyCounts.notMatched++
            : emptyCounts.notTested++;
  ((emptyCounts.tested = emptyCounts.matched + emptyCounts.notMatched),
    (emptyCounts.groups = new Set(
      findings.map((finding) => finding.group.id),
    ).size),
    (emptyCounts.findings = findings.length));
  const n = findings.reduce((o, a) => o + BigInt(a.missing?.count ?? 0), 0n);
  return (
    n > BigInt(Number.MAX_SAFE_INTEGER) &&
      throwErrorMessage(
        "The total missing-value count exceeds the exact integer reporting limit. Split sequence groups into smaller checks.",
      ),
    (emptyCounts.gapValues = Number(n)),
    emptyCounts
  );
}
function buildRowMatchKey(e, rule) {
  const values = e.map((n) =>
    typeof n == "string"
      ? rule.caseSensitive
        ? rule.trimText
          ? n.trim()
          : n
        : (rule.trimText ? n.trim() : n).toLowerCase()
      : n,
  );
  return values.some(
    (value) =>
      isNullOrEmptyValue(value) ||
      (typeof value == "number" && !Number.isFinite(value)),
  )
    ? null
    : {
        key: JSON.stringify(values.map((value) => [typeof value, value])),
        values: values,
      };
}
function runPatternChecks(workbook, r, t = {}) {
  const patternCheckSettings = normalizePatternCheckSettings(workbook, r),
    rangeAddress = parseRangeAddress(patternCheckSettings.range),
    a = {
      sheetId: patternCheckSettings.sheetId,
      range: patternCheckSettings.range,
      headers: patternCheckSettings.headers,
      columns: [],
    };
  let materialitySourceEvidence = buildMaterialitySourceEvidence(
      workbook,
      a,
      t,
    ),
    d = "";
  if (patternCheckSettings.comparison === "calculated") {
    const columns = [
      ...new Set(patternCheckSettings.rules.flatMap(ruleKeyColumns)),
    ].filter((w) => w >= rangeAddress.firstCol && w <= rangeAddress.lastCol);
    try {
      materialitySourceEvidence = buildMaterialitySourceEvidence(
        workbook,
        { ...a, columns: columns },
        t,
      );
    } catch (error) {
      d =
        error instanceof Error
          ? error.message
          : "Selected fields could not be calculated.";
    }
  }
  const rows = materialitySourceEvidence.rows.map((row) => ({
      ...row,
      origin: findRowOrigin(
        workbook,
        materialitySourceEvidence.sheetId,
        row.sourceRow - 1,
      ),
      exclusionReason: patternCheckSettings.excludedRows.includes(row.sourceRow)
        ? "Explicitly excluded source row."
        : patternCheckSettings.excludeBlankRows &&
            row.values.every(isNullOrEmptyValue)
          ? "Entire selected row is blank; excluded under the visible blank-row policy."
          : null,
    })),
    source = {
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetId: materialitySourceEvidence.sheetId,
      sheetName: materialitySourceEvidence.sheetName,
      range: patternCheckSettings.range,
      sha256: sha256HexDigest(
        new TextEncoder().encode(
          JSON.stringify({
            headers: materialitySourceEvidence.headers,
            rows: rows,
          }),
        ),
      ),
    },
    c = new Map(rows.map((row) => [row.sourceRow, row])),
    h = (y, w) =>
      y.flatMap((E) =>
        w.map((A) => {
          const k = c.get(E);
          return {
            address: rN(E, A),
            raw: k.cells[A - rangeAddress.firstCol].raw,
            value: k.values[A - rangeAddress.firstCol],
            origin: k.origin,
          };
        }),
      ),
    f = () =>
      rows.map((row) => ({
        sourceRow: row.sourceRow,
        disposition: row.exclusionReason ? "excluded" : "not-tested",
        reasons: row.exclusionReason ? [row.exclusionReason] : [],
        findingIds: [],
      })),
    results = patternCheckSettings.rules.map((rule) => {
      const w = {
          rule: cloneStructuredValue(rule),
          version: "1.0.0",
          status: "complete",
          messages: [],
          counts: createEmptyCounts(rows.length),
          rows: f(),
          findings: [],
        },
        E = new Map(w.rows.map((row) => [row.sourceRow, row])),
        A = ruleKeyColumns(rule),
        k = A.every(
          (C) => C >= rangeAddress.firstCol && C <= rangeAddress.lastCol,
        ),
        F = rule.kind === "duplicates" ? !A.length : rule.column === null,
        I = () => {
          if (w.status === "failed") {
            w.findings = [];
            for (const row of w.rows)
              row.disposition !== "excluded" &&
                row.disposition !== "invalid" &&
                ((row.disposition = "not-tested"),
                (row.findingIds = []),
                row.reasons.push(
                  "Rule failed; no clean or exception conclusion is available.",
                ));
          }
          w.counts = summarizeRuleOutcome(w.rows, w.findings);
          for (const finding of w.findings)
            finding.counts = cloneStructuredValue(w.counts);
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
      if (rule.kind === "holiday" && !rule.dates.length)
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
        j = rows.filter((row) => !row.exclusionReason),
        z = new Map(),
        V = (C, T, N) => {
          let R = z.get(C);
          (R || ((R = { key: T, members: [] }), z.set(C, R)),
            R.members.push(N));
        },
        D = new Map(),
        B = new Set(rule.kind === "holiday" ? rule.dates : []);
      for (const C of j) {
        const T = E.get(C.sourceRow),
          N = (R) => C.values[R - rangeAddress.firstCol];
        if (rule.kind === "duplicates") {
          const R = rule.keyColumns.map(N),
            rowMatchKey = buildRowMatchKey(R, rule);
          if (!rowMatchKey) {
            rule.blankKeys === "exclude" &&
            R.some(
              ($) =>
                isNullOrEmptyValue($) ||
                (typeof $ == "string" && rule.trimText && !$.trim()),
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
          ((T.disposition = "not-matched"),
            V(rowMatchKey.key, rowMatchKey.values, C.sourceRow));
        } else if (rule.kind === "sequence") {
          const R = N(rule.column),
            L = rule.groupColumns.length
              ? buildRowMatchKey(rule.groupColumns.map(N), rule)
              : { key: "[]", values: [] };
          let $ = null;
          if (
            (typeof R == "number" &&
              Number.isSafeInteger(R) &&
              R >= 0 &&
              R <= gf.sequenceMaximum &&
              !rule.prefix &&
              !rule.suffix &&
              ($ = R),
            typeof R == "string" &&
              R.startsWith(rule.prefix) &&
              R.endsWith(rule.suffix))
          ) {
            const U = R.slice(
              rule.prefix.length,
              rule.suffix ? R.length - rule.suffix.length : undefined,
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
            (rule.start !== null && $ < rule.start) ||
            (rule.end !== null && $ > rule.end)
          ) {
            O(
              C.sourceRow,
              "Sequence value is outside the declared start/end bounds.",
            );
            continue;
          }
          (D.set(C.sourceRow, $), V(L.key, L.values, C.sourceRow));
        } else if (rule.kind === "round-amount") {
          const R = N(rule.column),
            L =
              typeof R == "number"
                ? toScaledIntegerAmount(R, rule.precision)
                : null;
          if (L === null) {
            O(
              C.sourceRow,
              "Amount must be a finite number of at most 1e15 magnitude, exactly representable at the declared decimal precision; numeric text is not coerced.",
            );
            continue;
          }
          if (((T.disposition = "not-matched"), L === 0n && !rule.includeZero))
            continue;
          const scaledIntegerAmount = toScaledIntegerAmount(
              rule.increment,
              rule.precision,
            ),
            U = toScaledIntegerAmount(rule.tolerance, rule.precision),
            G = L % scaledIntegerAmount;
          (G < scaledIntegerAmount - G ? G : scaledIntegerAmount - G) <= U &&
            V(JSON.stringify(R), [R], C.sourceRow);
        } else {
          const R = convertSerialToIsoDate(N(rule.column), rule.dateSystem);
          if (!R) {
            O(
              C.sourceRow,
              "Date must be strict ISO YYYY-MM-DD or a valid Excel date/time serial in the chosen system; serial 60 in the 1900 system is invalid.",
            );
            continue;
          }
          T.disposition = "not-matched";
          const L = new Date(parseIsoCalendarDate(R) * 864e5).getUTCDay();
          (rule.kind === "weekend" ? rule.days.includes(L) : B.has(R)) &&
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
      const _ = (C, T, N, R, L, missing) => {
        const U = `${rule.id}-G${w.findings.length + 1}`,
          G = {
            id: U,
            classification: "review-candidate",
            title: R,
            detail:
              L +
              " Review candidate only; this is not proof of error or fraud.",
            rule: { id: rule.id, kind: rule.kind, version: "1.0.0" },
            parameters: cloneStructuredValue(rule),
            source: cloneStructuredValue(source),
            sourceCells: h(N, A),
            group: { id: U, key: C, memberRows: [...T] },
            counts: createEmptyCounts(rows.length),
            ...(missing ? { missing: missing } : {}),
          };
        if ((w.findings.push(G), rule.kind !== "sequence"))
          for (const H of T)
            ((E.get(H).disposition = "matched"), E.get(H).findingIds.push(U));
        else for (const H of N) E.get(H).findingIds.push(U);
      };
      if (rule.kind === "sequence") {
        let C = 0;
        for (const T of z.values()) {
          const N = new Map();
          for (const member of T.members) {
            const X = D.get(member);
            N.set(X, [...(N.get(X) ?? []), member]);
          }
          const R = [...N.keys()].sort((Q, X) => Q - X),
            L = rule.start ?? R[0],
            $ = rule.end ?? R.at(-1);
          if (
            R.some((Q) => (Q - L) % rule.step !== 0) ||
            ($ - L) % rule.step !== 0
          ) {
            for (const member of T.members)
              O(
                member,
                "Sequence values or bounds do not align with the declared step.",
              );
            continue;
          }
          if (R.length < 2 && (rule.start === null || rule.end === null)) {
            C++;
            for (const member of T.members)
              E.get(member).reasons.push(
                "One distinct sequence value without both bounds cannot establish missing values.",
              );
            continue;
          }
          for (const member of T.members)
            E.get(member).disposition = "not-matched";
          let U = null,
            G = L;
          const H = (first, last, J, q) => {
            if (first > last) return;
            const count = (last - first) / rule.step + 1,
              K = [
                ...(J === null ? [] : (N.get(J) ?? [])),
                ...(q === null ? [] : (N.get(q) ?? [])),
              ],
              examples = Array.from(
                { length: Math.min(20, count) },
                (ae, oe) =>
                  rule.prefix + String(first + oe * rule.step) + rule.suffix,
              );
            _(
              T.key,
              T.members,
              K,
              `Missing sequence ${first === last ? first : `${first}–${last}`} (${count})`,
              "Missing interval under the declared bounds/step. Cited existing cells are boundary evidence, not erroneous rows.",
              {
                first: first,
                last: last,
                step: rule.step,
                count: count,
                examples: examples,
                examplesTruncated: count > 20,
              },
            );
          };
          for (const Q of R)
            (H(G, Q - rule.step, U, Q), (U = Q), (G = Q + rule.step));
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
          (rule.kind === "duplicates" && C.members.length < 2) ||
            _(
              C.key,
              C.members,
              C.members,
              `${rule.label}: ${C.members.length} row(s)`,
              rule.kind === "duplicates"
                ? "All members of the typed matching-key group are included."
                : rule.kind === "round-amount"
                  ? "Amount meets the declared increment, decimal precision, tolerance and zero policy."
                  : `Posting calendar date ${C.key[0]} matches the supplied ${rule.kind === "holiday" ? "holiday calendar" : "weekend weekdays"}.`,
            );
      return I();
    }),
    combination = {
      status: "complete",
      messages: [],
      counts: createEmptyCounts(rows.length),
      rows: f(),
      findings: [],
    },
    v = results.filter((result) =>
      patternCheckSettings.combination.ruleIds.includes(result.rule.id),
    );
  if (!v.length)
    ((combination.status = "not-applicable"),
      combination.messages.push(
        "No row combination configured. Sequence gaps remain independent interval findings.",
      ));
  else if (v.some((y) => y.status !== "complete"))
    ((combination.status = v.some((y) => y.status === "failed")
      ? "failed"
      : "not-applicable"),
      combination.messages.push(
        "Every selected rule must complete before a combined result is available.",
      ));
  else
    for (const row of combination.rows) {
      if (row.disposition === "excluded") continue;
      const w = v.map((I) => ({
        result: I,
        row: I.rows.find((O) => O.sourceRow === row.sourceRow),
      }));
      if (
        w.some((I) => !["matched", "not-matched"].includes(I.row.disposition))
      ) {
        row.reasons.push(
          "Combination uses only rows tested by every selected rule (complete cases).",
        );
        continue;
      }
      const E = w.filter((I) => I.row.disposition === "matched"),
        A =
          patternCheckSettings.combination.mode === "any"
            ? E.length > 0
            : E.length === v.length;
      if (((row.disposition = A ? "matched" : "not-matched"), !A)) continue;
      const k = `combined-R${row.sourceRow}`;
      row.findingIds.push(k);
      const F = [...new Set(v.flatMap((I) => ruleKeyColumns(I.rule)))];
      combination.findings.push({
        id: k,
        classification: "review-candidate",
        title: `Combined review candidate — row ${row.sourceRow}`,
        detail: `${patternCheckSettings.combination.mode === "all" ? "All" : "At least one"} selected row rule matched. Review candidate only; not proof of error or fraud.`,
        rule: {
          id: "combined",
          kind: patternCheckSettings.combination.mode,
          version: "1.0.0",
        },
        parameters: cloneStructuredValue(patternCheckSettings.combination),
        source: cloneStructuredValue(source),
        sourceCells: h([row.sourceRow], F),
        group: { id: k, key: [row.sourceRow], memberRows: [row.sourceRow] },
        counts: createEmptyCounts(rows.length),
        matchedRules: E.map((I) => ({
          ruleId: I.result.rule.id,
          findingIds: [...I.row.findingIds],
        })),
      });
    }
  ((combination.counts = summarizeRuleOutcome(
    combination.rows,
    combination.findings,
  )),
    combination.status === "complete" &&
      !combination.counts.tested &&
      ((combination.status = "not-applicable"),
      combination.messages.push(
        "No row was eligible for every selected check.",
      )));
  for (const finding of combination.findings)
    finding.counts = cloneStructuredValue(combination.counts);
  const counts = {
    population: rows.length,
    excluded: rows.filter((row) => row.exclusionReason).length,
    eligible: rows.filter((row) => !row.exclusionReason).length,
    rulesComplete: results.filter((result) => result.status === "complete")
      .length,
    rulesFailed: results.filter((result) => result.status === "failed").length,
    rulesNotApplicable: results.filter(
      (result) => result.status === "not-applicable",
    ).length,
    findings: results.reduce((y, w) => y + w.findings.length, 0),
    gapValues: results.reduce((y, w) => y + w.counts.gapValues, 0),
  };
  Number.isSafeInteger(counts.gapValues) ||
    throwErrorMessage(
      "The aggregate missing-value count exceeds the exact integer reporting limit. Run fewer sequence rules.",
    );
  const b = {
    schemaVersion: 1,
    kind: "pattern-audit",
    algorithm: "ledgerglass-patterns-1",
    generatedAt: new Date().toISOString(),
    request: patternCheckSettings,
    context: materialitySourceEvidence.context,
    source: source,
    status: counts.rulesFailed
      ? "failed"
      : counts.rulesComplete
        ? "complete"
        : "not-applicable",
    messages: d
      ? [
          "Calculation failed. Evidence contains raw fields; no rule conclusion is available.",
          d,
        ]
      : [],
    headers: materialitySourceEvidence.headers,
    rows: rows,
    counts: counts,
    results: results,
    combination: combination,
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
      snapshot: JSON.stringify(workbook),
      request: JSON.stringify(patternCheckSettings),
      context: JSON.stringify(materialitySourceEvidence.context),
      content: JSON.stringify(b),
    }),
    b
  );
}
function checkReportCurrency(workbook, r, t, n = {}) {
  const o = KW.get(r),
    a = (reason) => ({ current: false, disclose: false, reason: reason });
  if (!o || o.content !== JSON.stringify(r))
    return a(
      "This report is untrusted or changed. Run the pattern checks again.",
    );
  try {
    const patternCheckSettings = normalizePatternCheckSettings(workbook, t),
      materialitySourceEvidence = buildMaterialitySourceEvidence(
        workbook,
        {
          sheetId: patternCheckSettings.sheetId,
          range: patternCheckSettings.range,
          headers: patternCheckSettings.headers,
          columns: [],
        },
        n,
      );
    return o.snapshot !== JSON.stringify(workbook) ||
      o.context !== JSON.stringify(materialitySourceEvidence.context) ||
      patternCheckSettings.sheetId !== r.request.sheetId ||
      patternCheckSettings.range !== r.request.range
      ? a(
          "Source or authorized scope changed. Previous evidence is hidden; run checks again.",
        )
      : o.request !== JSON.stringify(patternCheckSettings)
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
function exportScreenReportJson(workbook, r, t, n = {}) {
  const o = checkReportCurrency(workbook, r, t, n);
  return (o.current || throwErrorMessage(o.reason), JSON.stringify(r, null, 2));
}
function parseScreenReportRequest(workbook, r) {
  const t = r;
  return (
    (!t || t.schemaVersion !== 1 || t.kind !== "pattern-audit") &&
      throwErrorMessage("Choose a pattern-audit configuration or report."),
    normalizePatternCheckSettings(workbook, t.request)
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
