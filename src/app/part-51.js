import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { PJ, iu } from "./part-01.js";
import { U1, _t, cn, ct, ut } from "./part-03.js";
import { Dr, M, ci } from "./part-04.js";
import { MC, OC } from "./part-07.js";
import { yh } from "./part-09.js";
import { TI, X2, r1, s3 } from "./part-20.js";
import {
  Dj,
  ES,
  G4e,
  H4e,
  Oj,
  W4e,
  X4e,
  YW,
  _j,
  gf,
  iA,
  iN,
  q4e,
} from "./part-50.js";

function Z4e(e) {
  const { text: r, language: t } = Dr(),
    n = yh(),
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
          rules: [_j("duplicates", "rule1")],
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
        : YW(e.workbook, y, C, T)
      : null,
    U = $ != null && $.disclose ? y : null,
    G =
      (U == null ? void 0 : U.results.find((ce) => ce.rule.id === V)) ??
      (U == null ? void 0 : U.results[0]),
    H = V === "combined" ? (U == null ? void 0 : U.combination) : G,
    Q = Math.max(
      1,
      Math.ceil(((H == null ? void 0 : H.rows.length) ?? 0) / gf.pageSize),
    ),
    X$1 = Math.min(B, Q - 1),
    J = G4e(C.range),
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
    var ge, se;
    d.current = true;
    const ce = document.activeElement;
    return (
      (se =
        (ge = a.current) == null
          ? void 0
          : ge.querySelector('[data-ui-label="Source range"]')) == null ||
        se.focus(),
      () => {
        ((d.current = false),
          s.current++,
          ce != null && ce.isConnected && ce.focus());
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
        const ce = W4e(e.workbook, C, T);
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
        const we = await MC(ce, n.begin());
        if (!d.current || s.current !== ge || u.current !== se) return;
        const Ee = q4e(e.workbook, we);
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
        d.current && s.current === ge && u.current === se && te(OC(we));
      }
    },
    he = (ce) => {
      var ge;
      try {
        const se = ct(C.range),
          we =
            (ge =
              q == null ? void 0 : q.cells[`${ut(ce)}${se.firstRow + 1}`]) ==
            null
              ? void 0
              : ge.raw;
        return `${ut(ce)}${C.headers && we !== null && we !== void 0 ? " — " + String(we).slice(0, 60) : ""}`;
      } catch {
        return ut(ce);
      }
    },
    Se = (ce, ge) =>
      i.jsxs("label", {
        children: [
          i.jsx(M, { children: "Field" }),
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
                    "aria-label": `${ce.id} ${se} ${ut(Ee)}`,
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
        var ge;
        if (
          (ce.key === "Escape" && (ce.stopPropagation(), e.onClose()),
          ce.key === "Tab")
        ) {
          const se = Array.from(
              ((ge = a.current) == null
                ? void 0
                : ge.querySelectorAll(
                    "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
                  )) ?? [],
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
                  children: i.jsx(M, { children: "Pattern checks" }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      "Duplicates, sequence gaps and unusual postings · read-only analysis",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close pattern checks"),
              onClick: e.onClose,
              lang: t("Close pattern checks"),
              "data-ui-label": "Close pattern checks",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(M, {
            children:
              "Choose up to 500 source cells and map each rule explicitly. Results are review candidates, not proof of error or fraud. Every data-row position remains in the evidence ledger.",
          }),
        }),
        i.jsxs("div", {
          className: "materiality-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": r("Source range"),
                  value: C.range,
                  onChange: (ce) => W({ range: ce.target.value }),
                  lang: t("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Compare" }),
                i.jsxs("select", {
                  "aria-label": r("Comparison basis"),
                  value: C.comparison,
                  onChange: (ce) => W({ comparison: ce.target.value }),
                  lang: t("Comparison basis"),
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
                i.jsx(M, { children: "Exclude row numbers (comma separated)" }),
                i.jsx("input", {
                  "aria-label": r("Excluded rows"),
                  value: f,
                  onChange: (ce) => {
                    (p(ce.target.value),
                      W({
                        excludedRows: ce.target.value.trim()
                          ? ce.target.value.split(",").map(ES)
                          : [],
                      }));
                  },
                  lang: t("Excluded rows"),
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
                i.jsx(M, { children: "First row contains headers" }),
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
                i.jsx(M, { children: "Exclude entirely blank rows" }),
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
                  i.jsx(M, {
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
                  children: [ge + 1, ". ", iN[ce.kind], " · ", ce.id],
                }),
                i.jsxs("div", {
                  className: "pattern-rule-heading",
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Rule label" }),
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
                      children: [i.jsx(M, { children: "Remove " }), ce.id],
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
                              i.jsx(M, { children: "Case sensitive keys" }),
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
                              i.jsx(M, { children: "Trim key whitespace" }),
                            ],
                          }),
                        ],
                      }),
                    ce.kind === "duplicates" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Blank key policy" }),
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
                                  i.jsx(ci, {
                                    allowEmpty: se !== "step",
                                    "aria-label": `${ce.id} ${se}`,
                                    min: se === "step" ? 1 : 0,
                                    step: "1",
                                    value: Oj(ce[se]),
                                    onValueChange: (we) =>
                                      K(ge, {
                                        [se]:
                                          se !== "step" && we === ""
                                            ? null
                                            : ES(we),
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
                                  i.jsx(ci, {
                                    "aria-label": `${ce.id} ${se}`,
                                    min: "0",
                                    step: se === "precision" ? "1" : "any",
                                    value: Oj(ce[se]),
                                    onValueChange: (we) =>
                                      K(ge, { [se]: ES(we) }),
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
                              i.jsx(M, { children: "Flag zero amounts" }),
                            ],
                          }),
                        ],
                      }),
                    (ce.kind === "weekend" || ce.kind === "holiday") &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Excel date system" }),
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
                              i.jsx(M, { children: "Calendar name" }),
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
                              i.jsx(M, { children: "Calendar version" }),
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
                              i.jsx(M, {
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
                        children: i.jsx(M, {
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
                        children: i.jsx(M, { children: "Weekend weekdays" }),
                      }),
                      X4e.map((se, we) =>
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
                    children: i.jsx(M, {
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
                i.jsx(M, { children: "Add a rule" }),
                i.jsx("select", {
                  "aria-label": r("New rule type"),
                  value: g,
                  onChange: (ce) => b(ce.target.value),
                  lang: t("New rule type"),
                  "data-ui-label": "New rule type",
                  children: Object.entries(iN).map(([ce, ge]) =>
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
                W({ rules: [...C.rules, _j(g, ce)] });
              },
              children: i.jsx(M, { children: "Add rule" }),
            }),
            i.jsxs("span", {
              children: [
                C.rules.length,
                " ",
                i.jsx(M, { children: "/ 12 rules" }),
              ],
            }),
          ],
        }),
        i.jsxs("fieldset", {
          className: "pattern-rule",
          children: [
            i.jsx("legend", {
              children: i.jsx(M, { children: "Combine row rules (optional)" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Match" }),
                " ",
                i.jsxs("select", {
                  "aria-label": r("Combination mode"),
                  value: C.combination.mode,
                  onChange: (ce) =>
                    W({
                      combination: { ...C.combination, mode: ce.target.value },
                    }),
                  lang: t("Combination mode"),
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
              children: i.jsx(M, {
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
              children: i.jsx(M, { children: "Run pattern checks" }),
            }),
            i.jsx("button", {
              disabled: !($ != null && $.current),
              onClick: () => {
                try {
                  if (!y) throw Error("Run a current report first.");
                  (ae(H4e(e.workbook, y, C, T), "pattern-checks-report.json"),
                    z(
                      "Report download requested with source evidence, settings, counts and findings.",
                    ));
                } catch (ce) {
                  te(ce);
                }
              },
              children: i.jsx(M, { children: "Download report" }),
            }),
            i.jsx("button", {
              onClick: () => {
                try {
                  (ae(
                    JSON.stringify(
                      {
                        schemaVersion: 1,
                        kind: "pattern-audit",
                        request: iA(e.workbook, C),
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
              children: i.jsx(M, { children: "Save configuration" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Load configuration" }),
                i.jsx("input", {
                  "aria-label": r("Load pattern configuration"),
                  type: "file",
                  accept: ".json,application/json",
                  onChange: (ce) => {
                    var ge;
                    (de((ge = ce.target.files) == null ? void 0 : ge[0]),
                      (ce.target.value = ""));
                  },
                  lang: t("Load pattern configuration"),
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
            children: [i.jsx(M, { children: "Check failed: " }), k],
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
            "aria-label": r("Pattern report"),
            lang: t("Pattern report"),
            "data-ui-label": "Pattern report",
            children: [
              i.jsxs("h3", {
                children: [
                  $ != null && $.current ? "Current" : "Previous",
                  " ",
                  i.jsx(M, { children: "report —" }),
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
                  i.jsx(M, { children: "· revision" }),
                  " ",
                  U.source.revision,
                  " · ",
                  U.counts.population,
                  " ",
                  i.jsx(M, { children: "population · " }),
                  U.counts.excluded,
                  " ",
                  i.jsx(M, { children: "globally excluded ·" }),
                  " ",
                  U.counts.rulesComplete,
                  " ",
                  i.jsx(M, { children: "complete /" }),
                  " ",
                  U.counts.rulesFailed,
                  " ",
                  i.jsx(M, { children: "failed /" }),
                  " ",
                  U.counts.rulesNotApplicable,
                  " ",
                  i.jsx(M, { children: "not applicable rules" }),
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
                      i.jsx("p", { children: Dj(ce.counts) }),
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
                  i.jsx(M, { children: "Combined result — " }),
                  U.combination.status,
                ],
              }),
              i.jsx("p", { children: Dj(U.combination.counts) }),
              U.combination.messages.map((ce, ge) =>
                i.jsx("p", { children: ce }, ge),
              ),
              i.jsxs("label", {
                children: [
                  i.jsx(M, { children: "Evidence ledger" }),
                  i.jsxs("select", {
                    "aria-label": r("Evidence rule"),
                    value: V,
                    onChange: (ce) => {
                      (D(ce.target.value), _(0));
                    },
                    lang: t("Evidence rule"),
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
                            children: i.jsx(M, { children: "Source row" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Disposition" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, {
                              children: "Reasons / finding IDs",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Source values" }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children:
                        H == null
                          ? void 0
                          : H.rows.slice(X$1 * 20, (X$1 + 1) * 20).map((ce) => {
                              var ge;
                              return i.jsxs(
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
                                        (ge = U.rows.find(
                                          (se) => se.sourceRow === ce.sourceRow,
                                        )) == null
                                          ? void 0
                                          : ge.values,
                                      ),
                                    }),
                                  ],
                                },
                                ce.sourceRow,
                              );
                            }),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": r("Evidence pages"),
                lang: t("Evidence pages"),
                "data-ui-label": "Evidence pages",
                children: [
                  i.jsx("button", {
                    disabled: !X$1,
                    onClick: () => _(X$1 - 1),
                    children: i.jsx(M, { children: "Previous rows" }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(M, { children: "Page " }),
                      X$1 + 1,
                      " ",
                      i.jsx(M, { children: "of " }),
                      Q,
                    ],
                  }),
                  i.jsx("button", {
                    disabled: X$1 >= Q - 1,
                    onClick: () => _(X$1 + 1),
                    children: i.jsx(M, { children: "Next rows" }),
                  }),
                ],
              }),
              i.jsx("h4", {
                children: i.jsx(M, { children: "Review candidates" }),
              }),
              i.jsxs("p", {
                children: [
                  (H == null ? void 0 : H.findings.length) ?? 0,
                  " ",
                  i.jsx(M, {
                    children:
                      "finding groups / intervals. Open a finding for complete rule settings, counts, matching group and raw/evaluated source-cell citations.",
                  }),
                ],
              }),
              i.jsx("div", {
                className: "pattern-findings",
                children:
                  H == null
                    ? void 0
                    : H.findings.map((ce) =>
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
                                  i.jsx(M, { children: "Rule " }),
                                  ce.rule.id,
                                  " ",
                                  i.jsx(M, { children: "v" }),
                                  ce.rule.version,
                                  i.jsx(M, { children: "; group " }),
                                  ce.group.id,
                                  i.jsx(M, { children: "; source rows " }),
                                  ce.group.memberRows.join(", "),
                                ],
                              }),
                              ce.missing &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(M, { children: "Missing values: " }),
                                    ce.missing.count,
                                    i.jsx(M, { children: "; interval" }),
                                    " ",
                                    ce.missing.first,
                                    "–",
                                    ce.missing.last,
                                    i.jsx(M, { children: "; step " }),
                                    ce.missing.step,
                                    i.jsx(M, { children: "; examples " }),
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
                    children: i.jsx(M, {
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
            children: i.jsx(M, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const K4e = (e) => iu(new TextEncoder().encode(e));
function JW(e) {
  return K4e(
    JSON.stringify({
      workbookId: e.workbookId,
      revision: e.revision,
      sheetId: e.sheetId,
      sheetName: e.sheetName,
      range: e.range,
      ...PJ(e),
      cells: e.cells.map(({ address: r, raw: t, value: n }) => ({
        address: r,
        raw: t,
        value: n,
      })),
    }),
  );
}
const QW = new WeakMap(),
  eH = (e) =>
    JSON.stringify([
      e.sheetId,
      TI(e.range),
      e.ranges && e.ranges.length !== 1 ? e.ranges : void 0,
    ]);
function tH(e, r, t) {
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
function Mj(e, r, t, n, o = {}) {
  var v;
  const a = X2(e, r, t),
    s = s3(e, r);
  if (JSON.stringify(s) !== JSON.stringify(n))
    throw Error(
      "Answer source is stale or outside the original approved context. Request a fresh answer.",
    );
  if (r.ranges && r.ranges.length !== 1)
    throw Error("Citations require one approved rectangle.");
  const d = e.sheets.find((g) => g.id === s.sheetId),
    u = JW(s),
    l = new Map(s.cells.map((g) => [g.address, g])),
    c = [],
    h = new Map(),
    f = (g) => {
      const b = TI(g);
      if (b.some((E) => !l.has(E)))
        throw Error("Citation is outside approved context.");
      tH(e, s.sheetId, b);
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
              F = !!(k != null && k.literal);
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
                  : U1(e, d.id, _t(E).row),
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
      externalSources: (
        ((v = a.grounding) == null ? void 0 : v.externalSources) ?? []
      ).map((g) => ({ ...g, status: "supplied-not-verified" })),
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
    r1.receiptBytes
  )
    throw Error(
      "Complete answer evidence exceeds 2,000,000 bytes. Reduce the selection or answer; evidence is never silently truncated.",
    );
  return (
    QW.set(m, {
      snapshot: JSON.stringify(e),
      content: JSON.stringify(m),
      selection: eH(r),
    }),
    m
  );
}
function fg(e, r, t) {
  const n = QW.get(t),
    o = (a) => ({ current: false, disclose: false, reason: a });
  if (!n || JSON.stringify(t) !== n.content)
    return o(
      "Answer evidence is untrusted or altered. Request a fresh answer.",
    );
  try {
    if (eH(r) !== n.selection)
      return o(
        "Approved selection changed. Previous answer and citations are unavailable.",
      );
    if (JSON.stringify(e) !== n.snapshot)
      return o(
        "Workbook content or revision changed. Previous answer and citations are stale.",
      );
    for (const a of t.citations)
      tH(
        e,
        a.sheetId,
        a.cells.map((s) => s.address),
      );
    return JW(s3(e, r)) !== t.contextSha256
      ? o("Calculated source values changed. Request a fresh answer.")
      : {
          current: !0,
          disclose: !0,
          reason: "Current approved source and revision.",
        };
  } catch {
    return o(
      "Citation source is unavailable or outside the current authorized selection.",
    );
  }
}
function Y4e(e, r, t, n) {
  const o = fg(e, r, t);
  if (!o.current) throw Error(o.reason);
  const a = t.citations.find((s) => s.id === n);
  if (!a) throw Error("Unknown citation.");
  return structuredClone(a);
}
function J4e(e, r, t) {
  const n = fg(e, r, t);
  if (!n.current) throw Error(n.reason);
  return JSON.stringify(t, null, 2);
}
const jj = (e) =>
  e === null ? "(blank)" : typeof e == "string" ? JSON.stringify(e) : String(e);
function Q4e(e) {
  const { text: r, language: t } = Dr(),
    [n, o] = ne.useState(null),
    [a, s] = ne.useState(""),
    [d, u] = ne.useState(""),
    [l, c] = ne.useState(0),
    h = fg(e.workbook, e.selection, e.evidence);
  if (!h.disclose)
    return i.jsx("p", {
      role: "status",
      className: "lg-evidence-warning",
      children: h.reason,
    });
  const f = e.evidence,
    p = f.citations.find((b) => b.id === n),
    m = Math.max(
      1,
      Math.ceil(((p == null ? void 0 : p.cells.length) ?? 0) / 20),
    ),
    v = Math.min(l, m - 1),
    g = (b, y) => {
      try {
        const w = Y4e(e.workbook, e.selection, f, b);
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
    "aria-label": r("Assistant answer evidence"),
    lang: t("Assistant answer evidence"),
    "data-ui-label": "Assistant answer evidence",
    children: [
      f.mode === "legacy-unverified" &&
        i.jsx("p", {
          className: "lg-evidence-warning",
          children: i.jsx(M, {
            children:
              "Unverified narrative: this adapter did not link individual claims to evidence. The listed cells are references only.",
          }),
        }),
      i.jsx("p", {
        className: "lg-evidence-note",
        children: i.jsx(M, {
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
                        i.jsx(M, { children: "· revision " }),
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
                        children: i.jsx(M, {
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
          "aria-label": r("Citation source details"),
          lang: t("Citation source details"),
          "data-ui-label": "Citation source details",
          children: [
            i.jsxs("h4", { children: [p.sheetName, "!", p.range] }),
            i.jsxs("dl", {
              children: [
                i.jsx("dt", { children: i.jsx(M, { children: "Workbook" }) }),
                i.jsx("dd", { children: p.workbookId }),
                i.jsx("dt", { children: i.jsx(M, { children: "Revision" }) }),
                i.jsx("dd", { children: p.revision }),
                i.jsx("dt", { children: i.jsx(M, { children: "Sheet ID" }) }),
                i.jsx("dd", { children: p.sheetId }),
                i.jsx("dt", { children: i.jsx(M, { children: "Captured" }) }),
                i.jsx("dd", { children: f.capturedAt }),
              ],
            }),
            i.jsxs("p", {
              children: [
                i.jsx(M, {
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
                          children: i.jsx(M, { children: "Stored input" }),
                        }),
                        i.jsx("dd", { children: jj(b.raw) }),
                        i.jsx("dt", {
                          children: i.jsx(M, { children: "Formula" }),
                        }),
                        i.jsx("dd", {
                          children:
                            b.formula ??
                            (b.literal
                              ? "Literal text; not a formula"
                              : "No formula"),
                        }),
                        i.jsx("dt", {
                          children: i.jsx(M, { children: "Captured value" }),
                        }),
                        i.jsx("dd", { children: jj(b.value) }),
                        i.jsx("dt", {
                          children: i.jsx(M, { children: "Source linkage" }),
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
                                  i.jsx(M, { children: "· original row" }),
                                  " ",
                                  b.sourceOrigin.originalRow,
                                  i.jsx("br", {}),
                                  i.jsx(M, { children: "Source ID: " }),
                                  b.sourceOrigin.sourceId,
                                  i.jsx("br", {}),
                                  i.jsx(M, { children: "Origin ID: " }),
                                  b.sourceOrigin.id,
                                  " (",
                                  b.sourceOrigin.kind,
                                  ")",
                                  i.jsx("br", {}),
                                  i.jsx("small", {
                                    children: i.jsx(M, {
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
              "aria-label": r("Citation evidence pages"),
              lang: t("Citation evidence pages"),
              "data-ui-label": "Citation evidence pages",
              children: [
                i.jsx("button", {
                  disabled: v === 0,
                  onClick: () => c(v - 1),
                  children: i.jsx(M, { children: "Previous cells" }),
                }),
                i.jsxs("span", {
                  children: [
                    i.jsx(M, { children: "Page " }),
                    v + 1,
                    " ",
                    i.jsx(M, { children: "of " }),
                    m,
                  ],
                }),
                i.jsx("button", {
                  disabled: v >= m - 1,
                  onClick: () => c(v + 1),
                  children: i.jsx(M, { children: "Next cells" }),
                }),
              ],
            }),
          ],
        }),
      i.jsx("button", {
        className: "lg-button full",
        onClick: () => {
          try {
            const b = J4e(e.workbook, e.selection, f),
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
        children: i.jsx(M, { children: "Download answer evidence" }),
      }),
    ],
  });
}
const $h = {
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
  Vf = (e) =>
    !!e && !e.literal && typeof e.raw == "string" && e.raw.startsWith("="),
  yp = (e) => (e == null ? void 0 : e.raw) ?? null,
  Fj = (e) => yp(e) === null || yp(e) === "",
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
        .filter((s) => t[s] !== void 0)
        .sort(),
      a = Object.keys(n)
        .filter((s) => n[s] !== void 0)
        .sort();
    return (
      o.length === a.length && o.every((s, d) => s === a[d] && ww(t[s], n[s]))
    );
  },
  xc = (e, r) => `${ut(r)}${e + 1}`,
  e2e = "[A-Za-z_.\\d\\u007F-\\uFFFF]+",
  I3 = `(?:'(?:[^']|'')*'|${e2e})`,
  t2e =
    "(?:\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*(?:\\s*:\\s*\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)?|\\$?[A-Za-z]{1,3}\\s*:\\s*\\$?[A-Za-z]{1,3}|\\$?[1-9]\\d*\\s*:\\s*\\$?[1-9]\\d*)",
  rH = new RegExp(`^(${I3}\\s*!\\s*)?(${t2e})(#)?`),
  r2e = new RegExp(`^(${I3})\\s*!\\s*([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*)`),
  n2e = new RegExp(`^${I3}\\s*:\\s*${I3}\\s*!`),
  Bj = (e) =>
    e
      .trim()
      .replace(/!\s*$/, "")
      .trim()
      .replace(/^'|'$/g, "")
      .replace(/''/g, "'"),
  i2e = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
  nH = (e) => !/[A-Za-z0-9_.!:]/.test(e),
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
function a2e(e, r) {
  var o;
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
    const d = rH.exec(a);
    if (d && nH(a[d[0].length] ?? "") && !/^\s*\(/.test(a.slice(d[0].length))) {
      ((t +=
        (((o = d[1]) == null ? void 0 : o.toUpperCase()) ?? "") +
        o2e(d[2], r) +
        (d[3] ?? "")),
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
  $h,
  Bj,
  Fj,
  Mj,
  Q4e,
  Vf,
  Z4e,
  a2e,
  fg,
  i2e,
  n2e,
  nH,
  r2e,
  rH,
  ww,
  xc,
  yp,
};
