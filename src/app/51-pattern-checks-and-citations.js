import { parseNumericText } from "./37-linear-solver-engine.js";
import { listRangeColumnIndexes } from "./21-aging-and-materiality-screens.js";

import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  buildCalculationClock,
  sha256HexDigest,
} from "./01-core-schemas-and-hashing.js";
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
} from "./07-xlsx-archive-reader.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  validateCitationRange,
  validateProposedCalculation,
  answerGroundingLimits,
  buildAssistantSelectionContext,
} from "./20-calculation-options-and-sources.js";
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
} from "./50-materiality-screen.js";

function PatternCheckDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = useCancellableOperation(),
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
    [m, v] = ne.useState(false),
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
      includeHidden: !!e.allowHidden && m,
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
    G = U?.results.find((ce) => ce.rule.id === V) ?? U?.results[0],
    H = V === "combined" ? U?.combination : G,
    Q = Math.max(1, Math.ceil((H?.rows.length ?? 0) / gf.pageSize)),
    X$1 = Math.min(B, Q - 1),
    J = listRangeColumnIndexes(C.range),
    q = e.workbook.sheets.find((ce) => ce.id === e.sheetId),
    W = (ce) => h((ge) => ({ ...ge, ...ce })),
    K = (ce, ge) =>
      h((se) => ({
        ...se,
        rules: se.rules.map((we, Ee) => (Ee === ce ? { ...we, ...ge } : we)),
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
    const ce = document.activeElement;
    return (
      a.current?.querySelector('[data-ui-label="Source range"]')?.focus(),
      () => {
        ((d.current = false), s.current++, ce?.isConnected && ce.focus());
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
    oe = () => {
      s.current++;
      try {
        const ce = runPatternChecks(e.workbook, C, T);
        (w(ce), A(N), F(""), z(""), D(ce.results[0].rule.id), _(0));
      } catch (ce) {
        (w(null), te(ce));
      }
    },
    de = async (ce) => {
      if (!ce) return;
      const ge = ++s.current,
        se = L;
      try {
        if (ce.size > gf.reportBytes)
          throw Error("Choose a configuration no larger than 2,000,000 bytes.");
        const we = await importNativeJsonFile(ce, n.begin());
        if (!d.current || s.current !== ge || u.current !== se) return;
        const Ee = parseScreenReportRequest(e.workbook, we);
        if (Ee.sheetId !== e.sheetId)
          throw Error("The saved configuration belongs to another worksheet.");
        (h(Ee),
          p(Ee.excludedRows.join(", ")),
          w(null),
          F(""),
          z(
            "Configuration loaded. Run checks against the current source; saved findings are not trusted.",
          ));
      } catch (we) {
        d.current &&
          s.current === ge &&
          u.current === se &&
          te(toFileProcessingError(we));
      }
    },
    he = (ce) => {
      try {
        const se = parseRangeAddress(C.range),
          we = q?.cells[`${columnIndexToLetters(ce)}${se.firstRow + 1}`]?.raw;
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
            onChange: (se) =>
              K(ge, {
                column: se.target.value === "" ? null : Number(se.target.value),
              }),
            children: [
              i.jsx("option", { value: "", children: "Select a column" }),
              J.map((se) =>
                i.jsx("option", { value: se, children: he(se) }, se),
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
          J.map((Ee) =>
            i.jsxs(
              "label",
              {
                children: [
                  i.jsx("input", {
                    "aria-label": `${ce.id} ${se} ${columnIndexToLetters(Ee)}`,
                    type: "checkbox",
                    checked: we.includes(Ee),
                    onChange: (Ae) =>
                      K(ge, {
                        [se]: Ae.target.checked
                          ? [...we, Ee]
                          : we.filter((ye) => ye !== Ee),
                      }),
                  }),
                  he(Ee),
                ],
              },
              Ee,
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
      onKeyDown: (ce) => {
        if (
          (ce.key === "Escape" && (ce.stopPropagation(), e.onClose()),
          ce.key === "Tab")
        ) {
          const se = Array.from(
              a.current?.querySelectorAll(
                "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
              ) ?? [],
            ).filter((Ae) => Ae.getClientRects().length),
            we = se[0],
            Ee = se.at(-1);
          ce.shiftKey && document.activeElement === we
            ? (ce.preventDefault(), Ee == null || Ee.focus())
            : !ce.shiftKey &&
              document.activeElement === Ee &&
              (ce.preventDefault(), we == null || we.focus());
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
                  onChange: (ce) => W({ range: ce.target.value }),
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
                  onChange: (ce) => W({ comparison: ce.target.value }),
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
                  onChange: (ce) => {
                    (p(ce.target.value),
                      W({
                        excludedRows: ce.target.value.trim()
                          ? ce.target.value.split(",").map(parseNumericText)
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
                  onChange: (ce) => W({ headers: ce.target.checked }),
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
                  onChange: (ce) => W({ excludeBlankRows: ce.target.checked }),
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
                    checked: m,
                    onChange: (ce) => v(ce.target.checked),
                  }),
                  i.jsx(LocalizedText, {
                    children:
                      "Include authorized hidden / filtered rows and columns",
                  }),
                ],
              }),
          ],
        }),
        C.rules.map((ce, ge) =>
          i.jsxs(
            "fieldset",
            {
              className: "pattern-rule",
              children: [
                i.jsxs("legend", {
                  children: [
                    ge + 1,
                    ". ",
                    screenRuleLabels[ce.kind],
                    " · ",
                    ce.id,
                  ],
                }),
                i.jsxs("div", {
                  className: "pattern-rule-heading",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Rule label" }),
                        i.jsx("input", {
                          "aria-label": `${ce.id} label`,
                          maxLength: 80,
                          value: ce.label,
                          onChange: (se) => K(ge, { label: se.target.value }),
                        }),
                      ],
                    }),
                    i.jsxs("button", {
                      disabled: C.rules.length === 1,
                      onClick: () =>
                        W({
                          rules: C.rules.filter((se, we) => we !== ge),
                          combination: {
                            ...C.combination,
                            ruleIds: C.combination.ruleIds.filter(
                              (se) => se !== ce.id,
                            ),
                          },
                        }),
                      children: [
                        i.jsx(LocalizedText, { children: "Remove " }),
                        ce.id,
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "materiality-fields",
                  children: [
                    ce.kind !== "duplicates" && Se(ce, ge),
                    (ce.kind === "duplicates" || ce.kind === "sequence") &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            className: "materiality-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: ce.caseSensitive,
                                onChange: (se) =>
                                  K(ge, { caseSensitive: se.target.checked }),
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
                                checked: ce.trimText,
                                onChange: (se) =>
                                  K(ge, { trimText: se.target.checked }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Trim key whitespace",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ce.kind === "duplicates" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Blank key policy",
                          }),
                          i.jsxs("select", {
                            "aria-label": `${ce.id} blank key policy`,
                            value: ce.blankKeys,
                            onChange: (se) =>
                              K(ge, { blankKeys: se.target.value }),
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
                    ce.kind === "sequence" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          ["prefix", "suffix"].map((se) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  se,
                                  i.jsx("input", {
                                    "aria-label": `${ce.id} ${se}`,
                                    value: ce[se],
                                    onChange: (we) =>
                                      K(ge, { [se]: we.target.value }),
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
                                    "aria-label": `${ce.id} ${se}`,
                                    min: se === "step" ? 1 : 0,
                                    step: "1",
                                    value: nanToEmptyText(ce[se]),
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
                    ce.kind === "round-amount" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          ["increment", "precision", "tolerance"].map((se) =>
                            i.jsxs(
                              "label",
                              {
                                children: [
                                  se,
                                  i.jsx(LocalizedNumberField, {
                                    "aria-label": `${ce.id} ${se}`,
                                    min: "0",
                                    step: se === "precision" ? "1" : "any",
                                    value: nanToEmptyText(ce[se]),
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
                                checked: ce.includeZero,
                                onChange: (se) =>
                                  K(ge, { includeZero: se.target.checked }),
                              }),
                              i.jsx(LocalizedText, {
                                children: "Flag zero amounts",
                              }),
                            ],
                          }),
                        ],
                      }),
                    (ce.kind === "weekend" || ce.kind === "holiday") &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Excel date system",
                          }),
                          i.jsxs("select", {
                            "aria-label": `${ce.id} date system`,
                            value: ce.dateSystem,
                            onChange: (se) =>
                              K(ge, { dateSystem: se.target.value }),
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
                    ce.kind === "holiday" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calendar name",
                              }),
                              i.jsx("input", {
                                "aria-label": `${ce.id} calendar name`,
                                value: ce.calendarName,
                                onChange: (se) =>
                                  K(ge, { calendarName: se.target.value }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Calendar version",
                              }),
                              i.jsx("input", {
                                "aria-label": `${ce.id} calendar version`,
                                value: ce.calendarVersion,
                                onChange: (se) =>
                                  K(ge, { calendarVersion: se.target.value }),
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
                                "aria-label": `${ce.id} holiday dates`,
                                rows: 4,
                                value: ce.dates.join(`
`),
                                onChange: (se) =>
                                  K(ge, {
                                    dates:
                                      se.target.value === ""
                                        ? []
                                        : se.target.value.split(`
`),
                                  }),
                              }),
                            ],
                          }),
                        ],
                      }),
                  ],
                }),
                ce.kind === "duplicates" &&
                  le(ce, ge, "keyColumns", ce.keyColumns),
                ce.kind === "sequence" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      le(ce, ge, "groupColumns", ce.groupColumns),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Without start/end bounds, only gaps between observed values can be inferred. Sequence gaps are intervals with boundary citations, not missing source rows.",
                        }),
                      }),
                    ],
                  }),
                ce.kind === "weekend" &&
                  i.jsxs("fieldset", {
                    className: "pattern-checks",
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Weekend weekdays",
                        }),
                      }),
                      weekdayNames.map((se, we) =>
                        i.jsxs(
                          "label",
                          {
                            children: [
                              i.jsx("input", {
                                "aria-label": `${ce.id} ${se}`,
                                type: "checkbox",
                                checked: ce.days.includes(we),
                                onChange: (Ee) =>
                                  K(ge, {
                                    days: Ee.target.checked
                                      ? [...ce.days, we]
                                      : ce.days.filter((Ae) => Ae !== we),
                                  }),
                              }),
                              se,
                            ],
                          },
                          se,
                        ),
                      ),
                    ],
                  }),
                ce.kind === "holiday" &&
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
                      children:
                        "Use an approved, versioned calendar for the relevant jurisdiction. No official holiday dates are inferred.",
                    }),
                  }),
              ],
            },
            ce.id,
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
                  onChange: (ce) => b(ce.target.value),
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
                for (; C.rules.some((ge) => ge.id === ce);)
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
                  onChange: (ce) =>
                    W({
                      combination: { ...C.combination, mode: ce.target.value },
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
                .filter((ce) => ce.kind !== "sequence")
                .map((ce) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          "aria-label": `Combine ${ce.id}`,
                          checked: C.combination.ruleIds.includes(ce.id),
                          onChange: (ge) =>
                            W({
                              combination: {
                                ...C.combination,
                                ruleIds: ge.target.checked
                                  ? [...C.combination.ruleIds, ce.id]
                                  : C.combination.ruleIds.filter(
                                      (se) => se !== ce.id,
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
              onClick: oe,
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
                } catch (ce) {
                  te(ce);
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
                } catch (ce) {
                  te(ce);
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
                  onChange: (ce) => {
                    (de(ce.target.files?.[0]), (ce.target.value = ""));
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
              U.messages.map((ce, ge) =>
                i.jsx("p", { role: "alert", children: ce }, ge),
              ),
              U.results.map((ce) =>
                i.jsxs(
                  "article",
                  {
                    className: "pattern-rule-result",
                    children: [
                      i.jsxs("h4", {
                        children: [
                          ce.rule.label,
                          " (",
                          ce.rule.id,
                          ") — ",
                          ce.status,
                        ],
                      }),
                      i.jsx("p", {
                        children: formatScreenCountsSummary(ce.counts),
                      }),
                      ce.messages.map((ge, se) =>
                        i.jsx("p", { children: ge }, se),
                      ),
                    ],
                  },
                  ce.rule.id,
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
              U.combination.messages.map((ce, ge) =>
                i.jsx("p", { children: ce }, ge),
              ),
              i.jsxs("label", {
                children: [
                  i.jsx(LocalizedText, { children: "Evidence ledger" }),
                  i.jsxs("select", {
                    "aria-label": text("Evidence rule"),
                    value: V,
                    onChange: (ce) => {
                      (D(ce.target.value), _(0));
                    },
                    lang: language("Evidence rule"),
                    "data-ui-label": "Evidence rule",
                    children: [
                      U.results.map((ce) =>
                        i.jsxs(
                          "option",
                          {
                            value: ce.rule.id,
                            children: [ce.rule.label, " (", ce.rule.id, ")"],
                          },
                          ce.rule.id,
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
                                      (se) => se.sourceRow === ce.sourceRow,
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
                children: H?.findings.map((ce) =>
                  i.jsxs(
                    "details",
                    {
                      children: [
                        i.jsxs("summary", {
                          children: [ce.title, " · ", ce.id],
                        }),
                        i.jsx("p", { children: ce.detail }),
                        i.jsxs("p", {
                          children: [
                            i.jsx(LocalizedText, { children: "Rule " }),
                            ce.rule.id,
                            " ",
                            i.jsx(LocalizedText, { children: "v" }),
                            ce.rule.version,
                            i.jsx(LocalizedText, { children: "; group " }),
                            ce.group.id,
                            i.jsx(LocalizedText, {
                              children: "; source rows ",
                            }),
                            ce.group.memberRows.join(", "),
                          ],
                        }),
                        ce.missing &&
                          i.jsxs("p", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Missing values: ",
                              }),
                              ce.missing.count,
                              i.jsx(LocalizedText, { children: "; interval" }),
                              " ",
                              ce.missing.first,
                              "–",
                              ce.missing.last,
                              i.jsx(LocalizedText, { children: "; step " }),
                              ce.missing.step,
                              i.jsx(LocalizedText, { children: "; examples " }),
                              ce.missing.examples.join(", "),
                              ce.missing.examplesTruncated
                                ? " (first 20 only; complete interval above)"
                                : "",
                            ],
                          }),
                        i.jsx("pre", {
                          children: JSON.stringify(ce, null, 2),
                        }),
                      ],
                    },
                    ce.id,
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
                  U.notes.map((ce) => i.jsx("p", { children: ce }, ce)),
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
function buildCitationContext(e) {
  return hashCitationText(
    JSON.stringify({
      workbookId: e.workbookId,
      revision: e.revision,
      sheetId: e.sheetId,
      sheetName: e.sheetName,
      range: e.range,
      ...buildCalculationClock(e),
      cells: e.cells.map(({ address, raw, value }) => ({
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
function validateCitationVisibility(e, r, t) {
  const n = e.sheets.find((s) => s.id === r);
  if (!n || n.hidden || n.veryHidden)
    throw Error("Citation worksheet is unavailable or hidden.");
  const o = cn(n, "row"),
    a = cn(n, "column");
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
  return n;
}
function buildAnswerCitations(e, r, t, n, o = {}) {
  const a = validateProposedCalculation(e, r, t),
    s = buildAssistantSelectionContext(e, r);
  if (JSON.stringify(s) !== JSON.stringify(n))
    throw Error(
      "Answer source is stale or outside the original approved context. Request a fresh answer.",
    );
  if (r.ranges && r.ranges.length !== 1)
    throw Error("Citations require one approved rectangle.");
  const d = e.sheets.find((g) => g.id === s.sheetId),
    u = buildCitationContext(s),
    l = new Map(s.cells.map((g) => [g.address, g])),
    c = [],
    h = new Map(),
    f = (g) => {
      const b = validateCitationRange(g);
      if (b.some((E) => !l.has(E)))
        throw Error("Citation is outside approved context.");
      validateCitationVisibility(e, s.sheetId, b);
      const y = b.length === 1 ? b[0] : `${b[0]}:${b.at(-1)}`;
      let w = h.get(y);
      return (
        w ||
        ((w = `cite-${c.length + 1}`),
        h.set(y, w),
        c.push({
          id: w,
          workbookId: e.id,
          revision: e.revision,
          sheetId: s.sheetId,
          sheetName: s.sheetName,
          range: y,
          contextSha256: u,
          cells: b.map((E) => {
            const A = l.get(E),
              k = d.cells[E],
              F = !!k?.literal;
            return {
              ...A,
              literal: F,
              formula:
                !F && typeof A.raw == "string" && A.raw.startsWith("=")
                  ? A.raw
                  : null,
              sourceOrigin:
                o.includeSourceOrigins === false
                  ? null
                  : findRowOrigin(e, d.id, _t(E).row),
            };
          }),
        }),
        w)
      );
    },
    p = a.grounding
      ? a.grounding.claims.map((g) => ({
          id: g.id,
          kind: g.kind,
          text: g.text,
          citationIds: g.kind === "workbook" ? g.ranges.map(f) : [],
          sourceIds: g.kind === "external" ? [...g.sourceIds] : [],
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
      contextSha256: u,
      approvedSelection: { sheetId: r.sheetId, range: s.range },
      citations: c,
      claims: p,
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
      snapshot: JSON.stringify(e),
      content: JSON.stringify(m),
      selection: citationRangeKey(r),
    }),
    m
  );
}
function evaluateAnswerEvidence(e, r, t) {
  const n = citationRangeCache.get(t),
    o = (a) => ({ current: false, disclose: false, reason: a });
  if (!n || JSON.stringify(t) !== n.content)
    return o(
      "Answer evidence is untrusted or altered. Request a fresh answer.",
    );
  try {
    if (citationRangeKey(r) !== n.selection)
      return o(
        "Approved selection changed. Previous answer and citations are unavailable.",
      );
    if (JSON.stringify(e) !== n.snapshot)
      return o(
        "Workbook content or revision changed. Previous answer and citations are stale.",
      );
    for (const a of t.citations)
      validateCitationVisibility(
        e,
        a.sheetId,
        a.cells.map((s) => s.address),
      );
    return buildCitationContext(buildAssistantSelectionContext(e, r)) !==
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
function getCitationStatus(e, r, t, n) {
  const o = evaluateAnswerEvidence(e, r, t);
  if (!o.current) throw Error(o.reason);
  const a = t.citations.find((s) => s.id === n);
  if (!a) throw Error("Unknown citation.");
  return structuredClone(a);
}
function serializeCurrentCitation(e, r, t) {
  const n = evaluateAnswerEvidence(e, r, t);
  if (!n.current) throw Error(n.reason);
  return JSON.stringify(t, null, 2);
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
  const f = e.evidence,
    p = f.citations.find((b) => b.id === n),
    m = Math.max(1, Math.ceil((p?.cells.length ?? 0) / 20)),
    v = Math.min(l, m - 1),
    g = (b, y) => {
      try {
        const w = getCitationStatus(e.workbook, e.selection, f, b);
        if (y && !w.cells.some((E) => E.address === y))
          throw Error("Cell is outside this citation.");
        (e.onNavigate(w.sheetId, y ?? w.cells[0].address),
          o(b),
          y || c(0),
          s(""),
          u(""));
      } catch (w) {
        s(w instanceof Error ? w.message : "Citation unavailable.");
      }
    };
  return i.jsxs("section", {
    className: "lg-answer-evidence",
    "aria-label": text("Assistant answer evidence"),
    lang: language("Assistant answer evidence"),
    "data-ui-label": "Assistant answer evidence",
    children: [
      f.mode === "legacy-unverified" &&
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
      f.claims.map((b) =>
        i.jsxs(
          "article",
          {
            className: "lg-answer-claim",
            children: [
              i.jsx("strong", {
                children:
                  b.kind === "workbook"
                    ? "Workbook claim"
                    : b.kind === "external"
                      ? "External claim · source not verified"
                      : "Commentary · unverified",
              }),
              e.renderText
                ? i.jsx("div", {
                    className: "lg-answer-text",
                    children: e.renderText(b.text),
                  })
                : i.jsx("p", { children: b.text }),
              i.jsx("div", {
                className: "lg-citations",
                children: b.citationIds.map((y) => {
                  const w = f.citations.find((E) => E.id === y);
                  return i.jsxs(
                    "button",
                    {
                      type: "button",
                      "aria-label": `Inspect citation ${w.sheetName}!${w.range}`,
                      onClick: () => g(y),
                      children: [
                        w.sheetName,
                        "!",
                        w.range,
                        " ",
                        i.jsx(LocalizedText, { children: "· revision " }),
                        w.revision,
                      ],
                    },
                    y,
                  );
                }),
              }),
              b.sourceIds.map((y) => {
                const w = f.externalSources.find((E) => E.id === y);
                return i.jsxs(
                  "p",
                  {
                    className: "lg-external-source",
                    children: [
                      i.jsx("a", {
                        href: w.url,
                        target: "_blank",
                        rel: "noopener noreferrer",
                        referrerPolicy: "no-referrer",
                        children: w.title,
                      }),
                      i.jsx("br", {}),
                      w.url,
                      i.jsx("br", {}),
                      i.jsx("small", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "Adapter-supplied external source; not fetched or verified.",
                        }),
                      }),
                    ],
                  },
                  y,
                );
              }),
            ],
          },
          b.id,
        ),
      ),
      a &&
        i.jsx("p", {
          role: "alert",
          className: "lg-evidence-warning",
          children: a,
        }),
      d && i.jsx("p", { role: "status", children: d }),
      p &&
        i.jsxs("section", {
          className: "lg-citation-inspector",
          "aria-label": text("Citation source details"),
          lang: language("Citation source details"),
          "data-ui-label": "Citation source details",
          children: [
            i.jsxs("h4", { children: [p.sheetName, "!", p.range] }),
            i.jsxs("dl", {
              children: [
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Workbook" }),
                }),
                i.jsx("dd", { children: p.workbookId }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Revision" }),
                }),
                i.jsx("dd", { children: p.revision }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Sheet ID" }),
                }),
                i.jsx("dd", { children: p.sheetId }),
                i.jsx("dt", {
                  children: i.jsx(LocalizedText, { children: "Captured" }),
                }),
                i.jsx("dd", { children: f.capturedAt }),
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(LocalizedText, {
                  children:
                    "The cited cell is focused in the grid. The approved shared selection remains ",
                }),
                f.approvedSelection.range,
                ".",
              ],
            }),
            p.cells.slice(v * 20, (v + 1) * 20).map((b) =>
              i.jsxs(
                "article",
                {
                  className: "lg-citation-cell",
                  children: [
                    i.jsx("button", {
                      onClick: () => g(p.id, b.address),
                      "aria-label": `Focus cited cell ${b.address}`,
                      children: b.address,
                    }),
                    i.jsxs("dl", {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Stored input",
                          }),
                        }),
                        i.jsx("dd", {
                          children: formatCellValueForCitation(b.raw),
                        }),
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Formula",
                          }),
                        }),
                        i.jsx("dd", {
                          children:
                            b.formula ??
                            (b.literal
                              ? "Literal text; not a formula"
                              : "No formula"),
                        }),
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Captured value",
                          }),
                        }),
                        i.jsx("dd", {
                          children: formatCellValueForCitation(b.value),
                        }),
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Source linkage",
                          }),
                        }),
                        i.jsx("dd", {
                          children: b.sourceOrigin
                            ? i.jsxs(i.Fragment, {
                                children: [
                                  b.sourceOrigin.sourceName,
                                  " ·",
                                  " ",
                                  b.sourceOrigin.sourceSheet,
                                  " ",
                                  i.jsx(LocalizedText, {
                                    children: "· original row",
                                  }),
                                  " ",
                                  b.sourceOrigin.originalRow,
                                  i.jsx("br", {}),
                                  i.jsx(LocalizedText, {
                                    children: "Source ID: ",
                                  }),
                                  b.sourceOrigin.sourceId,
                                  i.jsx("br", {}),
                                  i.jsx(LocalizedText, {
                                    children: "Origin ID: ",
                                  }),
                                  b.sourceOrigin.id,
                                  " (",
                                  b.sourceOrigin.kind,
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
                b.address,
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
            const b = serializeCurrentCitation(e.workbook, e.selection, f),
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
          } catch (b) {
            s(b instanceof Error ? b.message : "Evidence unavailable.");
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
  readFormulaText = (e) => e?.raw ?? null,
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
function computeFormulaSignature(e, r) {
  if (!e.startsWith("=") || e.length > 4096)
    throw Error(
      "Pattern comparison requires a formula of at most 4,096 characters.",
    );
  _t(r);
  let t = "",
    n = 0;
  for (; n < e.length;) {
    const a = e.slice(n);
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
      for (; n < e.length && c;)
        (e[n] === "[" && c++, e[n] === "]" && c--, n++);
      if (c)
        throw Error(
          "An incomplete structured reference has no reliable formula signature.",
        );
      t += e.slice(l, n);
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
    t += e[n++];
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
