import { parseNumericText } from "./37-statistics-and-linear-solver.js";
import { listRangeColumnIndexes } from "./21-aging-and-materiality-screens.js";

import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  buildCalculationClock,
  sha256HexDigest,
} from "./01-schemas-date-functions-and-validation.js";
import {
  findRowOrigin,
  _t,
  cn,
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
  validateCitationRange,
  validateProposedCalculation,
  answerGroundingLimits,
  buildAssistantSelectionContext,
} from "./20-calculation-options-and-assistant.js";
import {
  formatScreenCountsSummary,
  exportScreenReportJson,
  nanToEmptyText,
  runPatternChecks,
  weekdayNames,
  checkReportCurrency,
  buildScreenRule,
  gf,
  normalizePatternCheckSettings,
  screenRuleLabels,
  parseScreenReportRequest,
} from "./50-materiality-screen-and-pattern-checks.js";

function PatternCheckDialog(e) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    o = ne.useId(),
    a = ne.useRef(null),
    s = ne.useRef(0),
    d = ne.useRef(true),
    u = ne.useRef(""),
    l = ne.useRef(2),
    [c, h] = ne.useState(() =>
      structuredClone(
        e.initialConfiguration ?? {
          sheetId: e.sheetId,
          range: e.range,
          headers: true,
          comparison: "calculated",
          excludeBlankRows: true,
          excludedRows: [],
          rules: [buildScreenRule("duplicates", "rule1")],
          combination: { mode: "all", ruleIds: [] },
        },
      ),
    ),
    [f, p] = ne.useState(() => c.excludedRows.join(", ")),
    [checked, v] = ne.useState(false),
    [g, b] = ne.useState("duplicates"),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState(""),
    [V, D] = ne.useState(""),
    [B, _] = ne.useState(0),
    C = { ...c, sheetId: e.sheetId },
    T = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      includeHidden: !!e.allowHidden && checked,
    },
    N = JSON.stringify([e.sheetId, e.range]),
    R = JSON.stringify([e.workbook, e.allowedSheetIds, T, N]),
    L = JSON.stringify([R, C, f]);
  u.current = L;
  const $ = y
      ? E !== N
        ? {
            current: false,
            disclose: false,
            reason: "The host selection changed. Run checks again.",
          }
        : checkReportCurrency(e.workbook, y, C, T)
      : null,
    U = $?.disclose ? y : null,
    G = U?.results.find((result) => result.rule.id === V) ?? U?.results[0],
    H = V === "combined" ? U?.combination : G,
    Q = Math.max(1, Math.ceil((H?.rows.length ?? 0) / gf.pageSize)),
    X$1 = Math.min(B, Q - 1),
    rangeColumnIndexes = listRangeColumnIndexes(C.range),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    W = (ce) => h((ge) => ({ ...ge, ...ce })),
    K = (ce, ge) =>
      h((se) => ({
        ...se,
        rules: se.rules.map((rule, Ee) =>
          Ee === ce ? { ...rule, ...ge } : rule,
        ),
      })),
    te = (ce) => {
      (F(
        ce instanceof Error
          ? ce.message
          : "The check failed. Review source and settings.",
      ),
        O(R),
        z(""));
    };
  ne.useEffect(() => {
    d.current = true;
    const activeElement = document.activeElement;
    return (
      a.current?.querySelector('[data-ui-label="Source range"]')?.focus(),
      () => {
        ((d.current = false),
          s.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []);
  const ae = (ce, ge) => {
      const se = URL.createObjectURL(
          new Blob([ce], { type: "application/json" }),
        ),
        we = document.createElement("a");
      ((we.href = se), (we.download = ge), we.click(), URL.revokeObjectURL(se));
    },
    onClick = () => {
      s.current++;
      try {
        const ce = runPatternChecks(e.workbook, C, T);
        (w(ce), A(N), F(""), z(""), D(ce.results[0].rule.id), _(0));
      } catch (error) {
        (w(null), te(error));
      }
    },
    de = async (ce) => {
      if (!ce) return;
      const ge = ++s.current,
        se = L;
      try {
        if (ce.size > gf.reportBytes)
          throw Error("Choose a configuration no larger than 2,000,000 bytes.");
        const we = await importNativeJsonFile(ce, cancellableOperation.begin());
        if (!d.current || s.current !== ge || u.current !== se) return;
        const screenReportRequest = parseScreenReportRequest(e.workbook, we);
        if (screenReportRequest.sheetId !== e.sheetId)
          throw Error("The saved configuration belongs to another worksheet.");
        (h(screenReportRequest),
          p(screenReportRequest.excludedRows.join(", ")),
          w(null),
          F(""),
          z(
            "Configuration loaded. Run checks against the current source; saved findings are not trusted.",
          ));
      } catch (error) {
        d.current &&
          s.current === ge &&
          u.current === se &&
          te(toFileProcessingError(error));
      }
    },
    he = (ce) => {
      try {
        const rangeAddress = parseRangeAddress(C.range),
          we =
            sheet?.cells[
              `${columnIndexToLetters(ce)}${rangeAddress.firstRow + 1}`
            ]?.raw;
        return `${columnIndexToLetters(ce)}${C.headers && we !== null && we !== undefined ? " — " + String(we).slice(0, 60) : ""}`;
      } catch {
        return columnIndexToLetters(ce);
      }
    },
    Se = (ce, ge) =>
      i.jsxs("label", {
        children: [
          i.jsx(LocalizedText, { children: "Field" }),
          i.jsxs("select", {
            "aria-label": `${ce.id} field`,
            value: "column" in ce ? (ce.column ?? "") : "",
            onChange: (event) =>
              K(ge, {
                column:
                  event.target.value === "" ? null : Number(event.target.value),
              }),
            children: [
              i.jsx("option", { value: "", children: "Select a column" }),
              rangeColumnIndexes.map((rangeColumnIndex) =>
                i.jsx(
                  "option",
                  { value: rangeColumnIndex, children: he(rangeColumnIndex) },
                  rangeColumnIndex,
                ),
              ),
            ],
          }),
        ],
      }),
    le = (ce, ge, se, we) =>
      i.jsxs("fieldset", {
        className: "pattern-checks",
        children: [
          i.jsx("legend", {
            children:
              se === "keyColumns"
                ? "Matching key fields"
                : "Sequence grouping fields (optional)",
          }),
          rangeColumnIndexes.map((rangeColumnIndex) =>
            i.jsxs(
              "label",
              {
                children: [
                  i.jsx("input", {
                    "aria-label": `${ce.id} ${se} ${columnIndexToLetters(rangeColumnIndex)}`,
                    type: "checkbox",
                    checked: we.includes(rangeColumnIndex),
                    onChange: (event) =>
                      K(ge, {
                        [se]: event.target.checked
                          ? [...we, rangeColumnIndex]
                          : we.filter((ye) => ye !== rangeColumnIndex),
                      }),
                  }),
                  he(rangeColumnIndex),
                ],
              },
              rangeColumnIndex,
            ),
          ),
        ],
      });
  return i.jsx("div", {
    className: "materiality-backdrop",
    children: i.jsxs("section", {
      className: "materiality-dialog pattern-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": o,
      ref: a,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.stopPropagation(), e.onClose()),
          event.key === "Tab")
        ) {
          const se = Array.from(
              a.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? [],
            ).filter((Ae) => Ae.getClientRects().length),
            we = se[0],
            Ee = se.at(-1);
          event.shiftKey && document.activeElement === we
            ? (event.preventDefault(), Ee == null || Ee.focus())
            : !event.shiftKey &&
              document.activeElement === Ee &&
              (event.preventDefault(), we == null || we.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: o,
                  children: i.jsx(LocalizedText, {
                    children: "Pattern checks",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Duplicates, sequence gaps and unusual postings · read-only analysis",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close pattern checks"),
              onClick: e.onClose,
              lang: language("Close pattern checks"),
              "data-ui-label": "Close pattern checks",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Choose up to 500 source cells and map each rule explicitly. Results are review candidates, not proof of error or fraud. Every data-row position remains in the evidence ledger.",
          }),
        }),
        i.jsxs("div", {
          className: "materiality-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: C.range,
                  onChange: (event) => W({ range: event.target.value }),
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Compare" }),
                i.jsxs("select", {
                  "aria-label": text("Comparison basis"),
                  value: C.comparison,
                  onChange: (event) => W({ comparison: event.target.value }),
                  lang: language("Comparison basis"),
                  "data-ui-label": "Comparison basis",
                  children: [
                    i.jsx("option", {
                      value: "calculated",
                      children: "Calculated scalar values",
                    }),
                    i.jsx("option", {
                      value: "raw",
                      children: "Stored raw values / expressions",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, {
                  children: "Exclude row numbers (comma separated)",
                }),
                i.jsx("input", {
                  "aria-label": text("Excluded rows"),
                  value: f,
                  onChange: (event) => {
                    (p(event.target.value),
                      W({
                        excludedRows: event.target.value.trim()
                          ? event.target.value.split(",").map(parseNumericText)
                          : [],
                      }));
                  },
                  lang: language("Excluded rows"),
                  "data-ui-label": "Excluded rows",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "materiality-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: C.headers,
                  onChange: (event) => W({ headers: event.target.checked }),
                }),
                i.jsx(LocalizedText, {
                  children: "First row contains headers",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "materiality-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: C.excludeBlankRows,
                  onChange: (event) =>
                    W({ excludeBlankRows: event.target.checked }),
                }),
                i.jsx(LocalizedText, {
                  children: "Exclude entirely blank rows",
                }),
              ],
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "materiality-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => v(event.target.checked),
                  }),
                  i.jsx(LocalizedText, {
                    children:
                      "Include authorized hidden / filtered rows and columns",
                  }),
                ],
              }),
          ],
        }),
        C.rules.map((rule, ge) =>
          i.jsxs(
            "fieldset",
            {
              className: "pattern-rule",
              children: [
                i.jsxs("legend", {
                  children: [
                    ge + 1,
                    ". ",
                    screenRuleLabels[rule.kind],
                    " · ",
                    rule.id,
                  ],
                }),
                i.jsxs("div", {
                  className: "pattern-rule-heading",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Rule label" }),
                        i.jsx("input", {
                          "aria-label": `${rule.id} label`,
                          maxLength: 80,
                          value: rule.label,
                          onChange: (event) =>
                            K(ge, { label: event.target.value }),
                        }),
                      ],
                    }),
                    i.jsxs("button", {
                      disabled: C.rules.length === 1,
                      onClick: () =>
                        W({
                          rules: C.rules.filter((rule, we) => we !== ge),
                          combination: {
                            ...C.combination,
                            ruleIds: C.combination.ruleIds.filter(
                              (ruleId) => ruleId !== rule.id,
                            ),
                          },
                        }),
                      children: [
                        i.jsx(LocalizedText, { children: "Remove " }),
                        rule.id,
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "materiality-fields",
                  children: [
                    rule.kind !== "duplicates" && Se(rule, ge),
                    (rule.kind === "duplicates" || rule.kind === "sequence") &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            className: "materiality-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: rule.caseSensitive,
                                onChange: (event) =>
                                  K(ge, {
                                    caseSensitive: event.target.checked,
                                  }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Case sensitive keys",
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            className: "materiality-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: rule.trimText,
                                onChange: (event) =>
                                  K(ge, { trimText: event.target.checked }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Trim key whitespace",
                              }),
                            ],
                          }),
                        ],
                      }),
                    rule.kind === "duplicates" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Blank key policy",
                          }),
                          i.jsxs("select", {
                            "aria-label": `${rule.id} blank key policy`,
                            value: rule.blankKeys,
                            onChange: (event) =>
                              K(ge, { blankKeys: event.target.value }),
                            children: [
                              i.jsx("option", {
                                value: "fail",
                                children: "Fail the rule",
                              }),
                              i.jsx("option", {
                                value: "exclude",
                                children: "Exclude with reason",
                              }),
                            ],
                          }),
                        ],
                      }),
                    rule.kind === "sequence" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          ["prefix", "suffix"].map((se) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  se,
                                  i.jsx("input", {
                                    "aria-label": `${rule.id} ${se}`,
                                    value: rule[se],
                                    onChange: (event) =>
                                      K(ge, { [se]: event.target.value }),
                                  }),
                                ],
                              },
                              se,
                            ),
                          ),
                          ["step", "start", "end"].map((se) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  se,
                                  se !== "step" ? " (optional)" : "",
                                  i.jsx(LocalizedNumberField, {
                                    allowEmpty: se !== "step",
                                    "aria-label": `${rule.id} ${se}`,
                                    min: se === "step" ? 1 : 0,
                                    step: "1",
                                    value: nanToEmptyText(rule[se]),
                                    onValueChange: (we) =>
                                      K(ge, {
                                        [se]:
                                          se !== "step" && we === ""
                                            ? null
                                            : parseNumericText(we),
                                      }),
                                  }),
                                ],
                              },
                              se,
                            ),
                          ),
                        ],
                      }),
                    rule.kind === "round-amount" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          ["increment", "precision", "tolerance"].map((se) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  se,
                                  i.jsx(LocalizedNumberField, {
                                    "aria-label": `${rule.id} ${se}`,
                                    min: "0",
                                    step: se === "precision" ? "1" : "any",
                                    value: nanToEmptyText(rule[se]),
                                    onValueChange: (we) =>
                                      K(ge, { [se]: parseNumericText(we) }),
                                  }),
                                ],
                              },
                              se,
                            ),
                          ),
                          i.jsxs("label", {
                            className: "materiality-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: rule.includeZero,
                                onChange: (event) =>
                                  K(ge, { includeZero: event.target.checked }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Flag zero amounts",
                              }),
                            ],
                          }),
                        ],
                      }),
                    (rule.kind === "weekend" || rule.kind === "holiday") &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Excel date system",
                          }),
                          i.jsxs("select", {
                            "aria-label": `${rule.id} date system`,
                            value: rule.dateSystem,
                            onChange: (event) =>
                              K(ge, { dateSystem: event.target.value }),
                            children: [
                              i.jsx("option", {
                                value: "1900",
                                children: "1900",
                              }),
                              i.jsx("option", {
                                value: "1904",
                                children: "1904",
                              }),
                            ],
                          }),
                        ],
                      }),
                    rule.kind === "holiday" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calendar name",
                              }),
                              i.jsx("input", {
                                "aria-label": `${rule.id} calendar name`,
                                value: rule.calendarName,
                                onChange: (event) =>
                                  K(ge, { calendarName: event.target.value }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calendar version",
                              }),
                              i.jsx("input", {
                                "aria-label": `${rule.id} calendar version`,
                                value: rule.calendarVersion,
                                onChange: (event) =>
                                  K(ge, {
                                    calendarVersion: event.target.value,
                                  }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children:
                                  "Holiday dates (YYYY-MM-DD, one per line)",
                              }),
                              i.jsx("textarea", {
                                "aria-label": `${rule.id} holiday dates`,
                                rows: 4,
                                value: rule.dates.join(`
`),
                                onChange: (event) =>
                                  K(ge, {
                                    dates:
                                      event.target.value === ""
                                        ? []
                                        : event.target.value.split(`
`),
                                  }),
                              }),
                            ],
                          }),
                        ],
                      }),
                  ],
                }),
                rule.kind === "duplicates" &&
                  le(rule, ge, "keyColumns", rule.keyColumns),
                rule.kind === "sequence" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      le(rule, ge, "groupColumns", rule.groupColumns),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Without start/end bounds, only gaps between observed values can be inferred. Sequence gaps are intervals with boundary citations, not missing source rows.",
                        }),
                      }),
                    ],
                  }),
                rule.kind === "weekend" &&
                  i.jsxs("fieldset", {
                    className: "pattern-checks",
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Weekend weekdays",
                        }),
                      }),
                      weekdayNames.map((weekdayName, we) =>
                        i.jsxs(
                          "label",
                          {
                            children: [
                              i.jsx("input", {
                                "aria-label": `${rule.id} ${weekdayName}`,
                                type: "checkbox",
                                checked: rule.days.includes(we),
                                onChange: (event) =>
                                  K(ge, {
                                    days: event.target.checked
                                      ? [...rule.days, we]
                                      : rule.days.filter((day) => day !== we),
                                  }),
                              }),
                              weekdayName,
                            ],
                          },
                          weekdayName,
                        ),
                      ),
                    ],
                  }),
                rule.kind === "holiday" &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Use an approved, versioned calendar for the relevant jurisdiction. No official holiday dates are inferred.",
                    }),
                  }),
              ],
            },
            rule.id,
          ),
        ),
        i.jsxs("div", {
          className: "materiality-actions",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Add a rule" }),
                i.jsx("select", {
                  "aria-label": text("New rule type"),
                  value: g,
                  onChange: (event) => b(event.target.value),
                  lang: language("New rule type"),
                  "data-ui-label": "New rule type",
                  children: Object.entries(screenRuleLabels).map(([ce, ge]) =>
                    i.jsx("option", { value: ce, children: ge }, ce),
                  ),
                }),
              ],
            }),
            i.jsx("button", {
              disabled: C.rules.length >= gf.rules,
              onClick: () => {
                let ce = `rule${l.current++}`;
                for (; C.rules.some((rule) => rule.id === ce);)
                  ce = `rule${l.current++}`;
                W({ rules: [...C.rules, buildScreenRule(g, ce)] });
              },
              children: i.jsx(LocalizedText, { children: "Add rule" }),
            }),
            i.jsxs("span", {
              children: [
                C.rules.length,
                " ",
                i.jsx(LocalizedText, { children: "/ 12 rules" }),
              ],
            }),
          ],
        }),
        i.jsxs("fieldset", {
          className: "pattern-rule",
          children: [
            i.jsx("legend", {
              children: i.jsx(LocalizedText, {
                children: "Combine row rules (optional)",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Match" }),
                " ",
                i.jsxs("select", {
                  "aria-label": text("Combination mode"),
                  value: C.combination.mode,
                  onChange: (event) =>
                    W({
                      combination: {
                        ...C.combination,
                        mode: event.target.value,
                      },
                    }),
                  lang: language("Combination mode"),
                  "data-ui-label": "Combination mode",
                  children: [
                    i.jsx("option", {
                      value: "all",
                      children: "All selected rules",
                    }),
                    i.jsx("option", {
                      value: "any",
                      children: "Any selected rule",
                    }),
                  ],
                }),
              ],
            }),
            i.jsx("div", {
              className: "pattern-checks",
              children: C.rules
                .filter((rule) => rule.kind !== "sequence")
                .map((ce) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          "aria-label": `Combine ${ce.id}`,
                          checked: C.combination.ruleIds.includes(ce.id),
                          onChange: (event) =>
                            W({
                              combination: {
                                ...C.combination,
                                ruleIds: event.target.checked
                                  ? [...C.combination.ruleIds, ce.id]
                                  : C.combination.ruleIds.filter(
                                      (ruleId) => ruleId !== ce.id,
                                    ),
                              },
                            }),
                        }),
                        ce.label,
                        " (",
                        ce.id,
                        ")",
                      ],
                    },
                    ce.id,
                  ),
                ),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "A row must be tested by every selected rule. Failed or not-applicable rules block a combined conclusion.",
              }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "materiality-actions",
          children: [
            i.jsx("button", {
              className: "materiality-primary",
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Run pattern checks",
              }),
            }),
            i.jsx("button", {
              disabled: !$?.current,
              onClick: () => {
                try {
                  if (!y) throw Error("Run a current report first.");
                  (ae(
                    exportScreenReportJson(e.workbook, y, C, T),
                    "pattern-checks-report.json",
                  ),
                    z(
                      "Report download requested with source evidence, settings, counts and findings.",
                    ));
                } catch (error) {
                  te(error);
                }
              },
              children: i.jsx(LocalizedText, { children: "Download report" }),
            }),
            i.jsx("button", {
              onClick: () => {
                try {
                  (ae(
                    JSON.stringify(
                      {
                        schemaVersion: 1,
                        kind: "pattern-audit",
                        request: normalizePatternCheckSettings(e.workbook, C),
                      },
                      null,
                      2,
                    ),
                    "pattern-checks-configuration.json",
                  ),
                    z("Configuration download requested."));
                } catch (error) {
                  te(error);
                }
              },
              children: i.jsx(LocalizedText, {
                children: "Save configuration",
              }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Load configuration" }),
                i.jsx("input", {
                  "aria-label": text("Load pattern configuration"),
                  type: "file",
                  accept: ".json,application/json",
                  onChange: (event) => {
                    (de(event.target.files?.[0]), (event.target.value = ""));
                  },
                  lang: language("Load pattern configuration"),
                  "data-ui-label": "Load pattern configuration",
                }),
              ],
            }),
          ],
        }),
        k &&
          I === R &&
          i.jsxs("p", {
            className: "materiality-error",
            role: "alert",
            children: [i.jsx(LocalizedText, { children: "Check failed: " }), k],
          }),
        j && i.jsx("p", { role: "status", children: j }),
        $ &&
          !$.current &&
          i.jsx("p", {
            role: "status",
            className: "materiality-stale",
            children: $.reason,
          }),
        U &&
          i.jsxs("section", {
            className: "materiality-result",
            "aria-label": text("Pattern report"),
            lang: language("Pattern report"),
            "data-ui-label": "Pattern report",
            children: [
              i.jsxs("h3", {
                children: [
                  $?.current ? "Current" : "Previous",
                  " ",
                  i.jsx(LocalizedText, { children: "report —" }),
                  " ",
                  U.status,
                ],
              }),
              i.jsxs("p", {
                children: [
                  U.source.sheetName,
                  "!",
                  U.source.range,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision" }),
                  " ",
                  U.source.revision,
                  " · ",
                  U.counts.population,
                  " ",
                  i.jsx(LocalizedText, { children: "population · " }),
                  U.counts.excluded,
                  " ",
                  i.jsx(LocalizedText, { children: "globally excluded ·" }),
                  " ",
                  U.counts.rulesComplete,
                  " ",
                  i.jsx(LocalizedText, { children: "complete /" }),
                  " ",
                  U.counts.rulesFailed,
                  " ",
                  i.jsx(LocalizedText, { children: "failed /" }),
                  " ",
                  U.counts.rulesNotApplicable,
                  " ",
                  i.jsx(LocalizedText, { children: "not applicable rules" }),
                ],
              }),
              U.messages.map((message, ge) =>
                i.jsx("p", { role: "alert", children: message }, ge),
              ),
              U.results.map((result) =>
                i.jsxs(
                  "article",
                  {
                    className: "pattern-rule-result",
                    children: [
                      i.jsxs("h4", {
                        children: [
                          result.rule.label,
                          " (",
                          result.rule.id,
                          ") — ",
                          result.status,
                        ],
                      }),
                      i.jsx("p", {
                        children: formatScreenCountsSummary(result.counts),
                      }),
                      result.messages.map((message, se) =>
                        i.jsx("p", { children: message }, se),
                      ),
                    ],
                  },
                  result.rule.id,
                ),
              ),
              i.jsxs("h4", {
                children: [
                  i.jsx(LocalizedText, { children: "Combined result — " }),
                  U.combination.status,
                ],
              }),
              i.jsx("p", {
                children: formatScreenCountsSummary(U.combination.counts),
              }),
              U.combination.messages.map((message, ge) =>
                i.jsx("p", { children: message }, ge),
              ),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Evidence ledger" }),
                  i.jsxs("select", {
                    "aria-label": text("Evidence rule"),
                    value: V,
                    onChange: (event) => {
                      (D(event.target.value), _(0));
                    },
                    lang: language("Evidence rule"),
                    "data-ui-label": "Evidence rule",
                    children: [
                      U.results.map((result) =>
                        i.jsxs(
                          "option",
                          {
                            value: result.rule.id,
                            children: [
                              result.rule.label,
                              " (",
                              result.rule.id,
                              ")",
                            ],
                          },
                          result.rule.id,
                        ),
                      ),
                      i.jsx("option", {
                        value: "combined",
                        children: "Combined result",
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
                            children: i.jsx(LocalizedText, {
                              children: "Source row",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Disposition",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Reasons / finding IDs",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source values",
                            }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: H?.rows
                        .slice(X$1 * 20, (X$1 + 1) * 20)
                        .map((ce) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: ce.sourceRow }),
                                i.jsx("td", { children: ce.disposition }),
                                i.jsx("td", {
                                  children:
                                    [...ce.reasons, ...ce.findingIds].join(
                                      " · ",
                                    ) || "—",
                                }),
                                i.jsx("td", {
                                  children: JSON.stringify(
                                    U.rows.find(
                                      (row) => row.sourceRow === ce.sourceRow,
                                    )?.values,
                                  ),
                                }),
                              ],
                            },
                            ce.sourceRow,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Evidence pages"),
                lang: language("Evidence pages"),
                "data-ui-label": "Evidence pages",
                children: [
                  i.jsx("button", {
                    disabled: !X$1,
                    onClick: () => _(X$1 - 1),
                    children: i.jsx(LocalizedText, {
                      children: "Previous rows",
                    }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      X$1 + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      Q,
                    ],
                  }),
                  i.jsx("button", {
                    disabled: X$1 >= Q - 1,
                    onClick: () => _(X$1 + 1),
                    children: i.jsx(LocalizedText, { children: "Next rows" }),
                  }),
                ],
              }),
              i.jsx("h4", {
                children: i.jsx(LocalizedText, {
                  children: "Review candidates",
                }),
              }),
              i.jsxs("p", {
                children: [
                  H?.findings.length ?? 0,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "finding groups / intervals. Open a finding for complete rule settings, counts, matching group and raw/evaluated source-cell citations.",
                  }),
                ],
              }),
              i.jsx("div", {
                className: "pattern-findings",
                children: H?.findings.map((finding) =>
                  i.jsxs(
                    "details",
                    {
                      children: [
                        i.jsxs("summary", {
                          children: [finding.title, " · ", finding.id],
                        }),
                        i.jsx("p", { children: finding.detail }),
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, { children: "Rule " }),
                            finding.rule.id,
                            " ",
                            i.jsx(LocalizedText, { children: "v" }),
                            finding.rule.version,
                            i.jsx(LocalizedText, { children: "; group " }),
                            finding.group.id,
                            i.jsx(LocalizedText, {
                              children: "; source rows ",
                            }),
                            finding.group.memberRows.join(", "),
                          ],
                        }),
                        finding.missing &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Missing values: ",
                              }),
                              finding.missing.count,
                              i.jsx(LocalizedText, { children: "; interval" }),
                              " ",
                              finding.missing.first,
                              "–",
                              finding.missing.last,
                              i.jsx(LocalizedText, { children: "; step " }),
                              finding.missing.step,
                              i.jsx(LocalizedText, { children: "; examples " }),
                              finding.missing.examples.join(", "),
                              finding.missing.examplesTruncated
                                ? " (first 20 only; complete interval above)"
                                : "",
                            ],
                          }),
                        i.jsx("pre", {
                          children: JSON.stringify(finding, null, 2),
                        }),
                      ],
                    },
                    finding.id,
                  ),
                ),
              }),
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Source ledger, settings and interpretation notes",
                    }),
                  }),
                  U.notes.map((note) => i.jsx("p", { children: note }, note)),
                  i.jsx("pre", {
                    children: JSON.stringify(
                      { request: U.request, source: U.source, rows: U.rows },
                      null,
                      2,
                    ),
                  }),
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
const hashCitationText = (e) => sha256HexDigest(new TextEncoder().encode(e));
function buildCitationContext(assistantSelectionContext) {
  return hashCitationText(
    JSON.stringify({
      workbookId: assistantSelectionContext.workbookId,
      revision: assistantSelectionContext.revision,
      sheetId: assistantSelectionContext.sheetId,
      sheetName: assistantSelectionContext.sheetName,
      range: assistantSelectionContext.range,
      ...buildCalculationClock(assistantSelectionContext),
      cells: assistantSelectionContext.cells.map(({ address, raw, value }) => ({
        address,
        raw,
        value,
      })),
    }),
  );
}
const citationRangeCache = new WeakMap(),
  citationRangeKey = (e) =>
    JSON.stringify([
      e.sheetId,
      validateCitationRange(e.range),
      e.ranges && e.ranges.length !== 1 ? e.ranges : undefined,
    ]);
function validateCitationVisibility(workbook, sheetId, t) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (!sheet || sheet.hidden || sheet.veryHidden)
    throw Error("Citation worksheet is unavailable or hidden.");
  const o = cn(sheet, "row"),
    a = cn(sheet, "column");
  for (const s of t) {
    const d = _t(s);
    if (
      o.some((u) => d.row >= u.start && d.row <= u.end) ||
      a.some((u) => d.col >= u.start && d.col <= u.end)
    )
      throw Error(
        "Citation includes hidden cells. Make the cited evidence visible and request a fresh answer.",
      );
  }
  return sheet;
}
function buildAnswerCitations(
  activeWorkbook,
  assistantSelection,
  t,
  n,
  o = {},
) {
  const a = validateProposedCalculation(activeWorkbook, assistantSelection, t),
    assistantSelectionContext = buildAssistantSelectionContext(
      activeWorkbook,
      assistantSelection,
    );
  if (JSON.stringify(assistantSelectionContext) !== JSON.stringify(n))
    throw Error(
      "Answer source is stale or outside the original approved context. Request a fresh answer.",
    );
  if (assistantSelection.ranges && assistantSelection.ranges.length !== 1)
    throw Error("Citations require one approved rectangle.");
  const sheet = activeWorkbook.sheets.find(
      (sheet) => sheet.id === assistantSelectionContext.sheetId,
    ),
    citationContext = buildCitationContext(assistantSelectionContext),
    l = new Map(
      assistantSelectionContext.cells.map((cell) => [cell.address, cell]),
    ),
    citations = [],
    h = new Map(),
    f = (g) => {
      const b = validateCitationRange(g);
      if (b.some((E) => !l.has(E)))
        throw Error("Citation is outside approved context.");
      validateCitationVisibility(
        activeWorkbook,
        assistantSelectionContext.sheetId,
        b,
      );
      const range = b.length === 1 ? b[0] : `${b[0]}:${b.at(-1)}`;
      let w = h.get(range);
      return (
        w ||
        ((w = `cite-${citations.length + 1}`),
        h.set(range, w),
        citations.push({
          id: w,
          workbookId: activeWorkbook.id,
          revision: activeWorkbook.revision,
          sheetId: assistantSelectionContext.sheetId,
          sheetName: assistantSelectionContext.sheetName,
          range: range,
          contextSha256: citationContext,
          cells: b.map((E) => {
            const cell = l.get(E),
              k = sheet.cells[E],
              literal = !!k?.literal;
            return {
              ...cell,
              literal: literal,
              formula:
                !literal &&
                typeof cell.raw == "string" &&
                cell.raw.startsWith("=")
                  ? cell.raw
                  : null,
              sourceOrigin:
                o.includeSourceOrigins === false
                  ? null
                  : findRowOrigin(activeWorkbook, sheet.id, _t(E).row),
            };
          }),
        }),
        w)
      );
    },
    claims = a.grounding
      ? a.grounding.claims.map((claim) => ({
          id: claim.id,
          kind: claim.kind,
          text: claim.text,
          citationIds: claim.kind === "workbook" ? claim.ranges.map(f) : [],
          sourceIds: claim.kind === "external" ? [...claim.sourceIds] : [],
        }))
      : [
          {
            id: "legacy",
            kind: "commentary",
            text: a.explanation,
            citationIds: [...new Set(a.citations.map(f))],
            sourceIds: [],
          },
        ],
    m = {
      schemaVersion: 1,
      kind: "assistant-answer-evidence",
      capturedAt: new Date().toISOString(),
      proposalId: a.id,
      mode: a.grounding ? "linked-claims" : "legacy-unverified",
      contextSha256: citationContext,
      approvedSelection: {
        sheetId: assistantSelection.sheetId,
        range: assistantSelectionContext.range,
      },
      citations: citations,
      claims: claims,
      externalSources: (a.grounding?.externalSources ?? []).map((g) => ({
        ...g,
        status: "supplied-not-verified",
      })),
      sourceOriginsIncluded: o.includeSourceOrigins !== false,
      notes: [
        "Workbook citation identity, approved scope, captured values and revision are checked locally. These checks do not prove that the cited cells logically support every claim; the reviewer must inspect the evidence.",
        "Legacy narratives and commentary are unverified. A workbook citation cannot substantiate an external fact. External URLs are adapter-supplied references, not automatically fetched or verified.",
        "Source-row linkage is copied only from the current host workbook registry. It does not certify the original file bytes or identify an original column after transformations; missing linkage is explicitly unavailable.",
        "Only host-local evidence can include row-origin metadata. Provider context remains the selected raw/calculated cells. Downloaded evidence is mutable JSON, not an immutable audit service or a replayable authority.",
      ],
    };
  if (
    new TextEncoder().encode(JSON.stringify(m, null, 2)).length >
    answerGroundingLimits.receiptBytes
  )
    throw Error(
      "Complete answer evidence exceeds 2,000,000 bytes. Reduce the selection or answer; evidence is never silently truncated.",
    );
  return (
    citationRangeCache.set(m, {
      snapshot: JSON.stringify(activeWorkbook),
      content: JSON.stringify(m),
      selection: citationRangeKey(assistantSelection),
    }),
    m
  );
}
function evaluateAnswerEvidence(workbook, r, t) {
  const n = citationRangeCache.get(t),
    o = (reason) => ({ current: false, disclose: false, reason: reason });
  if (!n || JSON.stringify(t) !== n.content)
    return o(
      "Answer evidence is untrusted or altered. Request a fresh answer.",
    );
  try {
    if (citationRangeKey(r) !== n.selection)
      return o(
        "Approved selection changed. Previous answer and citations are unavailable.",
      );
    if (JSON.stringify(workbook) !== n.snapshot)
      return o(
        "Workbook content or revision changed. Previous answer and citations are stale.",
      );
    for (const citation of t.citations)
      validateCitationVisibility(
        workbook,
        citation.sheetId,
        citation.cells.map((cell) => cell.address),
      );
    return buildCitationContext(buildAssistantSelectionContext(workbook, r)) !==
      t.contextSha256
      ? o("Calculated source values changed. Request a fresh answer.")
      : {
          current: true,
          disclose: true,
          reason: "Current approved source and revision.",
        };
  } catch {
    return o(
      "Citation source is unavailable or outside the current authorized selection.",
    );
  }
}
function getCitationStatus(workbook, selection, evidence, n) {
  const o = evaluateAnswerEvidence(workbook, selection, evidence);
  if (!o.current) throw Error(o.reason);
  const citation = evidence.citations.find((citation) => citation.id === n);
  if (!citation) throw Error("Unknown citation.");
  return structuredClone(citation);
}
function serializeCurrentCitation(workbook, selection, evidence) {
  const n = evaluateAnswerEvidence(workbook, selection, evidence);
  if (!n.current) throw Error(n.reason);
  return JSON.stringify(evidence, null, 2);
}
const formatCellValueForCitation = (e) =>
  e === null ? "(blank)" : typeof e == "string" ? JSON.stringify(e) : String(e);
function AssistantAnswerEvidence(e) {
  const { text, language } = useWorkbookMessages(),
    [n, o] = ne.useState(null),
    [a, s] = ne.useState(""),
    [d, u] = ne.useState(""),
    [l, c] = ne.useState(0),
    h = evaluateAnswerEvidence(e.workbook, e.selection, e.evidence);
  if (!h.disclose)
    return i.jsx("p", {
      role: "status",
      className: "lg-evidence-warning",
      children: h.reason,
    });
  const evidence = e.evidence,
    citation = evidence.citations.find((citation) => citation.id === n),
    m = Math.max(1, Math.ceil((citation?.cells.length ?? 0) / 20)),
    v = Math.min(l, m - 1),
    g = (b, y) => {
      try {
        const citationStatus = getCitationStatus(
          e.workbook,
          e.selection,
          evidence,
          b,
        );
        if (y && !citationStatus.cells.some((cell) => cell.address === y))
          throw Error("Cell is outside this citation.");
        (e.onNavigate(
          citationStatus.sheetId,
          y ?? citationStatus.cells[0].address,
        ),
          o(b),
          y || c(0),
          s(""),
          u(""));
      } catch (error) {
        s(error instanceof Error ? error.message : "Citation unavailable.");
      }
    };
  return i.jsxs("section", {
    className: "lg-answer-evidence",
    "aria-label": text("Assistant answer evidence"),
    lang: language("Assistant answer evidence"),
    "data-ui-label": "Assistant answer evidence",
    children: [
      evidence.mode === "legacy-unverified" &&
        i.jsx("p", {
          className: "lg-evidence-warning",
          children: i.jsx(LocalizedText, {
            children:
              "Unverified narrative: this adapter did not link individual claims to evidence. The listed cells are references only.",
          }),
        }),
      i.jsx("p", {
        className: "lg-evidence-note",
        children: i.jsx(LocalizedText, {
          children:
            "Workbook links verify source identity, scope and revision. Review whether the evidence supports each statement.",
        }),
      }),
      evidence.claims.map((claim) =>
        i.jsxs(
          "article",
          {
            className: "lg-answer-claim",
            children: [
              i.jsx("strong", {
                children:
                  claim.kind === "workbook"
                    ? "Workbook claim"
                    : claim.kind === "external"
                      ? "External claim · source not verified"
                      : "Commentary · unverified",
              }),
              e.renderText
                ? i.jsx("div", {
                    className: "lg-answer-text",
                    children: e.renderText(claim.text),
                  })
                : i.jsx("p", { children: claim.text }),
              i.jsx("div", {
                className: "lg-citations",
                children: claim.citationIds.map((citationId) => {
                  const citation = evidence.citations.find(
                    (citation) => citation.id === citationId,
                  );
                  return i.jsxs(
                    "button",
                    {
                      type: "button",
                      "aria-label": `Inspect citation ${citation.sheetName}!${citation.range}`,
                      onClick: () => g(citationId),
                      children: [
                        citation.sheetName,
                        "!",
                        citation.range,
                        " ",
                        i.jsx(LocalizedText, { children: "· revision " }),
                        citation.revision,
                      ],
                    },
                    citationId,
                  );
                }),
              }),
              claim.sourceIds.map((sourceId) => {
                const externalSource = evidence.externalSources.find(
                  (externalSource) => externalSource.id === sourceId,
                );
                return i.jsxs(
                  "p",
                  {
                    className: "lg-external-source",
                    children: [
                      i.jsx("a", {
                        href: externalSource.url,
                        target: "_blank",
                        rel: "noopener noreferrer",
                        referrerPolicy: "no-referrer",
                        children: externalSource.title,
                      }),
                      i.jsx("br", {}),
                      externalSource.url,
                      i.jsx("br", {}),
                      i.jsx("small", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Adapter-supplied external source; not fetched or verified.",
                        }),
                      }),
                    ],
                  },
                  sourceId,
                );
              }),
            ],
          },
          claim.id,
        ),
      ),
      a &&
        i.jsx("p", {
          role: "alert",
          className: "lg-evidence-warning",
          children: a,
        }),
      d && i.jsx("p", { role: "status", children: d }),
      citation &&
        i.jsxs("section", {
          className: "lg-citation-inspector",
          "aria-label": text("Citation source details"),
          lang: language("Citation source details"),
          "data-ui-label": "Citation source details",
          children: [
            i.jsxs("h4", {
              children: [citation.sheetName, "!", citation.range],
            }),
            i.jsxs("dl", {
              children: [
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Workbook" }),
                }),
                i.jsx("dd", { children: citation.workbookId }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Revision" }),
                }),
                i.jsx("dd", { children: citation.revision }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Sheet ID" }),
                }),
                i.jsx("dd", { children: citation.sheetId }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Captured" }),
                }),
                i.jsx("dd", { children: evidence.capturedAt }),
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, {
                  children:
                    "The cited cell is focused in the grid. The approved shared selection remains ",
                }),
                evidence.approvedSelection.range,
                ".",
              ],
            }),
            citation.cells.slice(v * 20, (v + 1) * 20).map((cell) =>
              i.jsxs(
                "article",
                {
                  className: "lg-citation-cell",
                  children: [
                    i.jsx("button", {
                      onClick: () => g(citation.id, cell.address),
                      "aria-label": `Focus cited cell ${cell.address}`,
                      children: cell.address,
                    }),
                    i.jsxs("dl", {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Stored input",
                          }),
                        }),
                        i.jsx("dd", {
                          children: formatCellValueForCitation(cell.raw),
                        }),
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Formula",
                          }),
                        }),
                        i.jsx("dd", {
                          children:
                            cell.formula ??
                            (cell.literal
                              ? "Literal text; not a formula"
                              : "No formula"),
                        }),
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Captured value",
                          }),
                        }),
                        i.jsx("dd", {
                          children: formatCellValueForCitation(cell.value),
                        }),
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Source linkage",
                          }),
                        }),
                        i.jsx("dd", {
                          children: cell.sourceOrigin
                            ? i.jsxs(i.Fragment, {
                                children: [
                                  cell.sourceOrigin.sourceName,
                                  " ·",
                                  " ",
                                  cell.sourceOrigin.sourceSheet,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "· original row",
                                  }),
                                  " ",
                                  cell.sourceOrigin.originalRow,
                                  i.jsx("br", {}),
                                  i.jsx(LocalizedText, {
                                    children: "Source ID: ",
                                  }),
                                  cell.sourceOrigin.sourceId,
                                  i.jsx("br", {}),
                                  i.jsx(LocalizedText, {
                                    children: "Origin ID: ",
                                  }),
                                  cell.sourceOrigin.id,
                                  " (",
                                  cell.sourceOrigin.kind,
                                  ")",
                                  i.jsx("br", {}),
                                  i.jsx("small", {
                                    children: i.jsx(LocalizedText, {
                                      children:
                                        "Recorded row lineage; original file bytes and original column are not verified here.",
                                    }),
                                  }),
                                ],
                              })
                            : "No original source-row linkage is available in this evidence.",
                        }),
                      ],
                    }),
                  ],
                },
                cell.address,
              ),
            ),
            i.jsxs("nav", {
              "aria-label": text("Citation evidence pages"),
              lang: language("Citation evidence pages"),
              "data-ui-label": "Citation evidence pages",
              children: [
                i.jsx("button", {
                  disabled: v === 0,
                  onClick: () => c(v - 1),
                  children: i.jsx(LocalizedText, {
                    children: "Previous cells",
                  }),
                }),
                i.jsxs("span", {
                  children: [
                    i.jsx(LocalizedText, { children: "Page " }),
                    v + 1,
                    " ",
                    i.jsx(LocalizedText, { children: "of " }),
                    m,
                  ],
                }),
                i.jsx("button", {
                  disabled: v >= m - 1,
                  onClick: () => c(v + 1),
                  children: i.jsx(LocalizedText, { children: "Next cells" }),
                }),
              ],
            }),
          ],
        }),
      i.jsx("button", {
        className: "lg-button full",
        onClick: () => {
          try {
            const b = serializeCurrentCitation(
                e.workbook,
                e.selection,
                evidence,
              ),
              y = URL.createObjectURL(
                new Blob([b], { type: "application/json" }),
              ),
              w = document.createElement("a");
            ((w.href = y),
              (w.download = "assistant-answer-evidence.json"),
              w.click(),
              URL.revokeObjectURL(y),
              u(
                "Evidence download requested. Preserve it with the original workpaper.",
              ),
              s(""));
          } catch (error) {
            s(error instanceof Error ? error.message : "Evidence unavailable.");
          }
        },
        children: i.jsx(LocalizedText, {
          children: "Download answer evidence",
        }),
      }),
    ],
  });
}
const patternCheckBudgets = {
    mappings: 20,
    selectedCells: 1e4,
    rowsPerRange: 1e3,
    columnsPerRange: 100,
    scannedCells: 2e4,
    formulaCharacters: 2e6,
    findings: 3e4,
    nameDepth: 16,
    keys: 8,
  },
  readFormulaText = (cell) => cell?.raw ?? null,
  isBlankFormulaCell = (e) =>
    readFormulaText(e) === null || readFormulaText(e) === "",
  ww = (e, r) => {
    if (e === r) return true;
    if (!e || !r || typeof e != "object" || typeof r != "object") return false;
    if (Array.isArray(e) || Array.isArray(r))
      return (
        Array.isArray(e) &&
        Array.isArray(r) &&
        e.length === r.length &&
        e.every((s, d) => ww(s, r[d]))
      );
    const t = e,
      n = r,
      o = Object.keys(t)
        .filter((s) => t[s] !== undefined)
        .sort(),
      a = Object.keys(n)
        .filter((s) => n[s] !== undefined)
        .sort();
    return (
      o.length === a.length && o.every((s, d) => s === a[d] && ww(t[s], n[s]))
    );
  },
  unquotedSheetNamePattern = "[A-Za-z_.\\d\\u007F-\\uFFFF]+",
  I3 = `(?:'(?:[^']|'')*'|${unquotedSheetNamePattern})`,
  a1RangePattern =
    "(?:\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?:\\s*:\\s*\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)?|\\$?[A-Za-z]{1,3}\\s*:\\s*\\$?[A-Za-z]{1,3}|\\$?[1-9]\\d*\\s*:\\s*\\$?[1-9]\\d*)",
  referenceTokenPattern = new RegExp(
    `^(${I3}\\s*!\\s*)?(${a1RangePattern})(#)?`,
  ),
  definedNameReferencePattern = new RegExp(
    `^(${I3})\\s*!\\s*([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*)`,
  ),
  sheetSpanPattern = new RegExp(`^${I3}\\s*:\\s*${I3}\\s*!`),
  Bj = (e) =>
    e
      .trim()
      .replace(/!\s*$/, "")
      .trim()
      .replace(/^'|'$/g, "")
      .replace(/''/g, "'"),
  identifierPattern = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
  isNonReferenceToken = (e) => !/[A-Za-z0-9_.!:]/.test(e),
  o2e = (e, r) => {
    const t = _t(r);
    return e
      .replace(/\s/g, "")
      .replace(
        /(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)|(?:(\$?)([A-Za-z]{1,3}))|(?:(\$?)([1-9]\d*))/g,
        (n, o, a, s, d, u, l, c, h) => {
          const f = (p) =>
            [...p.toUpperCase()].reduce(
              (m, v) => m * 26 + v.charCodeAt(0) - 64,
              0,
            ) - 1;
          return a
            ? `${o ? "C$" + f(a) : "C[" + (f(a) - t.col) + "]"}${s ? "R$" + d : "R[" + (Number(d) - 1 - t.row) + "]"}`
            : l
              ? u
                ? "C$" + f(l)
                : "C[" + (f(l) - t.col) + "]"
              : c
                ? "R$" + h
                : "R[" + (Number(h) - 1 - t.row) + "]";
        },
      );
  };
function computeFormulaSignature(raw, r) {
  if (!raw.startsWith("=") || raw.length > 4096)
    throw Error(
      "Pattern comparison requires a formula of at most 4,096 characters.",
    );
  _t(r);
  let t = "",
    n = 0;
  for (; n < raw.length;) {
    const a = raw.slice(n);
    if (a[0] === '"') {
      const l = /^"(?:[^"]|"")*"/.exec(a);
      if (!l)
        throw Error(
          "An unterminated text literal has no reliable formula signature.",
        );
      ((t += l[0]), (n += l[0].length));
      continue;
    }
    const s = /^([A-Za-z_][A-Za-z0-9_.]*)?\[/.exec(a);
    if (s) {
      const l = n;
      n += s[0].length;
      let c = 1;
      for (; n < raw.length && c;)
        (raw[n] === "[" && c++, raw[n] === "]" && c--, n++);
      if (c)
        throw Error(
          "An incomplete structured reference has no reliable formula signature.",
        );
      t += raw.slice(l, n);
      continue;
    }
    const d = referenceTokenPattern.exec(a);
    if (
      d &&
      isNonReferenceToken(a[d[0].length] ?? "") &&
      !/^\s*\(/.test(a.slice(d[0].length))
    ) {
      ((t += (d[1]?.toUpperCase() ?? "") + o2e(d[2], r) + (d[3] ?? "")),
        (n += d[0].length));
      continue;
    }
    const u =
      /^(?:[A-Za-z_\\][A-Za-z0-9_.\\]*|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?)/.exec(
        a,
      );
    if (u) {
      ((t += u[0].toUpperCase()), (n += u[0].length));
      continue;
    }
    t += raw[n++];
  }
  return t;
}

export {
  patternCheckBudgets,
  Bj,
  isBlankFormulaCell,
  buildAnswerCitations,
  AssistantAnswerEvidence,
  PatternCheckDialog,
  computeFormulaSignature,
  evaluateAnswerEvidence,
  identifierPattern,
  sheetSpanPattern,
  isNonReferenceToken,
  definedNameReferencePattern,
  referenceTokenPattern,
  ww,
  readFormulaText,
};
