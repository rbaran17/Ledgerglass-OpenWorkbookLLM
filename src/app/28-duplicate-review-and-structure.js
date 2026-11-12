import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { Mt } from "./02-financial-functions.js";
import { U1, Vn, _t, ct, ut } from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import { LC, Uz, Z7, cg, lw, qk, rfe } from "./27-data-validation-dialog.js";

function $z(e) {
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const [s, d] = ne.useState(e.range),
    [u, l] = ne.useState(qk(e.range)),
    [c, h] = ne.useState(true),
    [f, p] = ne.useState("raw"),
    [m, v] = ne.useState(true),
    [g, b] = ne.useState(false),
    [y, w] = ne.useState("keep"),
    [E, A] = ne.useState("Unique review"),
    [k, F] = ne.useState(false),
    [I, O] = ne.useState(null),
    [j, z] = ne.useState(""),
    [V, D] = ne.useState(""),
    [B, _] = ne.useState(""),
    [C, T] = ne.useState("all"),
    [N, R] = ne.useState(0),
    L = qk(s),
    $ = e.workbook.sheets.find((le) => le.id === e.sheetId),
    U = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((le) => le.id)),
    ].sort(),
    G = {
      includeHidden: !!e.allowHidden && k,
      allowedSheetIds: e.allowedSheetIds,
    },
    H =
      !!$ &&
      U.includes($.id) &&
      (G.includeHidden || (!$.hidden && !$.veryHidden)),
    Q = JSON.stringify([e.sheetId, e.range]),
    X$1 = JSON.stringify([e.workbook, e.readOnly, U, G.includeHidden, Q]),
    J =
      !!I &&
      !e.readOnly &&
      I.snapshot === JSON.stringify(e.workbook) &&
      j === Q &&
      I.context.includeHidden === G.includeHidden &&
      JSON.stringify(I.context.allowedSheetIds) === JSON.stringify(U),
    q = () => {
      (O(null), D(""), R(0));
    },
    W = (le) => {
      (_(X$1),
        D(
          le instanceof Error
            ? le.message
            : "Duplicate review could not complete the operation.",
        ));
    };
  ne.useEffect(() => {
    var ce, ge;
    const le = document.activeElement;
    return (
      (ge =
        (ce = o.current) == null
          ? void 0
          : ce.querySelector('[data-ui-label="Source range"]')) == null ||
        ge.focus(),
      () => {
        le != null && le.isConnected && le.focus();
      }
    );
  }, []);
  const K = () => {
      try {
        const le = Uz(
          e.workbook,
          {
            sheetId: e.sheetId,
            range: s,
            keyColumns: u,
            headers: c,
            comparison: f,
            caseSensitive: m,
            trimText: g,
            blankPolicy: y,
            outputSheetName: E,
          },
          { ...G, readOnly: e.readOnly },
        );
        (O(le), z(Q), D(""), R(0), T("all"));
      } catch (le) {
        (O(null), W(le));
      }
    },
    te = () => {
      if (I)
        try {
          if (!J)
            throw Error(
              "The source selection or preview changed. Preview and review again.",
            );
          const le = rfe(e.workbook, I, { ...G, readOnly: e.readOnly });
          (e.onWorkbookChange(
            le,
            `Create duplicate-reviewed population: ${I.counts.retained} retained, ${I.counts.excluded} excluded`,
          ),
            O(null),
            D(""));
        } catch (le) {
          W(le);
        }
    },
    ae =
      (I == null
        ? void 0
        : I.rows.filter(
            (le) =>
              C === "all" ||
              (C === "groups" ? !!le.groupId : le.disposition === C),
          )) ?? [],
    oe = Math.max(1, Math.ceil(ae.length / lw.pageSize)),
    de = Math.min(N, oe - 1),
    he = ae.slice(de * 20, de * 20 + 20),
    Se = (le, ce) =>
      i.jsxs("details", {
        children: [
          i.jsx("summary", { children: Z7(ce) }),
          i.jsxs("dl", {
            children: [
              i.jsx("dt", { children: i.jsx(M, { children: "Stored value" }) }),
              i.jsx("dd", { children: Z7(le.raw) }),
              i.jsx("dt", {
                children: i.jsx(M, { children: "Stored as literal" }),
              }),
              i.jsx("dd", { children: le.literal ? "Yes" : "No" }),
              i.jsx("dt", { children: i.jsx(M, { children: "Note" }) }),
              i.jsx("dd", { children: le.note ?? "None" }),
              i.jsx("dt", { children: i.jsx(M, { children: "Tickmark" }) }),
              i.jsx("dd", { children: le.tickmark ?? "None" }),
              i.jsx("dt", {
                children: i.jsx(M, { children: "Cell formatting / link" }),
              }),
              i.jsx("dd", {
                children: i.jsx("code", {
                  children: JSON.stringify(
                    Object.fromEntries(
                      Object.entries(le).filter(
                        ([ge]) =>
                          !["raw", "literal", "note", "tickmark"].includes(ge),
                      ),
                    ),
                  ),
                }),
              }),
            ],
          }),
        ],
      });
  return i.jsx("div", {
    className: "dup-backdrop",
    onMouseDown: (le) => {
      le.target === le.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      className: "dup-dialog",
      ref: o,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (le) => {
        var ce, ge, se;
        if (
          (le.stopPropagation(),
          le.key === "Escape" && (le.preventDefault(), a.current()),
          le.key === "Tab")
        ) {
          const we = [
            ...(((ce = o.current) == null
              ? void 0
              : ce.querySelectorAll(
                  "button:not(:disabled),input:not(:disabled),select:not(:disabled),summary",
                )) ?? []),
          ].filter(
            (Ee) =>
              !Ee.closest("details:not([open])") || Ee.tagName === "SUMMARY",
          );
          le.shiftKey && document.activeElement === we[0]
            ? (le.preventDefault(), (ge = we.at(-1)) == null || ge.focus())
            : !le.shiftKey &&
              document.activeElement === we.at(-1) &&
              (le.preventDefault(), (se = we[0]) == null || se.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: `${n}-title`,
                  children: i.jsx(M, { children: "Remove duplicates" }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      "Create a reviewed unique population on a new worksheet.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close duplicate review"),
              onClick: e.onClose,
              lang: t("Close duplicate review"),
              "data-ui-label": "Close duplicate review",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(M, {
            children:
              "The original worksheet stays intact. The new worksheet contains retained rows and an exclusion appendix with every later duplicate and its original source locator.",
          }),
        }),
        i.jsxs("div", {
          className: "dup-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Source worksheet" }),
                i.jsx("strong", {
                  children: H ? $.name : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": r("Source range"),
                  value: H ? s : "",
                  readOnly: e.readOnly,
                  onChange: (le) => {
                    const ce = le.target.value;
                    (d(ce), l(qk(ce)), q());
                  },
                  lang: t("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "New worksheet name" }),
                i.jsx("input", {
                  "aria-label": r("New worksheet name"),
                  value: E,
                  readOnly: e.readOnly,
                  onChange: (le) => {
                    (A(le.target.value), q());
                  },
                  lang: t("New worksheet name"),
                  "data-ui-label": "New worksheet name",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Compare" }),
                i.jsxs("select", {
                  "aria-label": r("Comparison basis"),
                  value: f,
                  disabled: e.readOnly,
                  onChange: (le) => {
                    (p(le.target.value), q());
                  },
                  lang: t("Comparison basis"),
                  "data-ui-label": "Comparison basis",
                  children: [
                    i.jsx("option", {
                      value: "raw",
                      children: "Typed raw values / exact formula text",
                    }),
                    i.jsx("option", {
                      value: "calculated",
                      children: "Typed calculated scalar values",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Blank key policy" }),
                i.jsxs("select", {
                  "aria-label": r("Blank key policy"),
                  value: y,
                  disabled: e.readOnly,
                  onChange: (le) => {
                    (w(le.target.value), q());
                  },
                  lang: t("Blank key policy"),
                  "data-ui-label": "Blank key policy",
                  children: [
                    i.jsx("option", {
                      value: "keep",
                      children: "Keep every row with any blank key",
                    }),
                    i.jsx("option", {
                      value: "match",
                      children: "Match blanks together",
                    }),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: c,
                  disabled: e.readOnly,
                  onChange: (le) => {
                    (h(le.target.checked), q());
                  },
                }),
                i.jsx(M, { children: "First row contains headers" }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: m,
                  disabled: e.readOnly,
                  onChange: (le) => {
                    (v(le.target.checked), q());
                  },
                }),
                i.jsx(M, { children: "Case-sensitive text" }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: g,
                  disabled: e.readOnly,
                  onChange: (le) => {
                    (b(le.target.checked), q());
                  },
                }),
                i.jsx(M, { children: "Trim outer text whitespace" }),
              ],
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "dup-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: k,
                    onChange: (le) => {
                      (F(le.target.checked), q());
                    },
                  }),
                  i.jsx(M, { children: "Include hidden sources" }),
                ],
              }),
          ],
        }),
        i.jsxs("fieldset", {
          disabled: e.readOnly || !H,
          children: [
            i.jsx("legend", {
              children: i.jsx(M, {
                children: "Key columns — all selected keys must match",
              }),
            }),
            i.jsx("div", {
              className: "dup-keys",
              children:
                H &&
                L.map((le) =>
                  i.jsxs(
                    "label",
                    {
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          "aria-label": `Key column ${ut(le)}`,
                          checked: u.includes(le),
                          onChange: (ce) => {
                            (l(
                              ce.target.checked
                                ? [...u, le]
                                : u.filter((ge) => ge !== le),
                            ),
                              q());
                          },
                        }),
                        ut(le),
                      ],
                    },
                    le,
                  ),
                ),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(M, {
            children:
              "First source row wins. Numbers, text and Booleans remain distinct. Matching rules affect keys only; exported values preserve their original text. Raw formulas compare exactly and become literal text in the derived snapshot. Calculated mode freezes supported scalar results; cell display formats do not determine equality.",
          }),
        }),
        V &&
          B === X$1 &&
          i.jsx("p", { className: "dup-error", role: "alert", children: V }),
        i.jsx("button", {
          className: "dup-primary",
          disabled: e.readOnly || !H || !u.length,
          onClick: K,
          children: i.jsx(M, { children: "Preview duplicate review" }),
        }),
        I &&
          i.jsxs("section", {
            className: "dup-preview",
            "aria-label": r("Duplicate review preview"),
            lang: t("Duplicate review preview"),
            "data-ui-label": "Duplicate review preview",
            children: [
              J
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("h3", {
                        children: [
                          I.counts.input,
                          " ",
                          i.jsx(M, { children: "source rows · " }),
                          I.counts.retained,
                          " ",
                          i.jsx(M, { children: "retained · " }),
                          I.counts.excluded,
                          " ",
                          i.jsx(M, { children: "excluded" }),
                        ],
                      }),
                      i.jsxs("p", {
                        role: "status",
                        children: [
                          I.counts.duplicateGroups,
                          " ",
                          i.jsx(M, { children: "duplicate groups;" }),
                          " ",
                          I.counts.blankKeyRows,
                          " ",
                          i.jsx(M, {
                            children: "rows with a blank key. New worksheet: ",
                          }),
                          I.request.outputSheetName,
                          i.jsx(M, { children: ". Unique population:" }),
                          " ",
                          I.populationRange,
                          ".",
                          " ",
                          I.excludedRange
                            ? `Exclusion appendix: ${I.excludedRange}.`
                            : "No duplicate rows were found; the derived population still includes source locators.",
                        ],
                      }),
                      I.request.headers &&
                        i.jsxs("details", {
                          children: [
                            i.jsx("summary", {
                              children: i.jsx(M, {
                                children: "Source header evidence",
                              }),
                            }),
                            i.jsx("div", {
                              className: "dup-headers",
                              children: I.headerCells.map((le, ce) =>
                                i.jsxs(
                                  "div",
                                  {
                                    children: [
                                      i.jsx("strong", {
                                        children: ut(
                                          I.request.keyColumns.length
                                            ? ct(I.request.range).firstCol + ce
                                            : ce,
                                        ),
                                      }),
                                      Se(le, le.raw),
                                    ],
                                  },
                                  ce,
                                ),
                              ),
                            }),
                          ],
                        }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Preview view" }),
                          i.jsxs("select", {
                            "aria-label": r("Preview view"),
                            value: C,
                            onChange: (le) => {
                              (T(le.target.value), R(0));
                            },
                            lang: t("Preview view"),
                            "data-ui-label": "Preview view",
                            children: [
                              i.jsx("option", {
                                value: "all",
                                children: "Before — every source row",
                              }),
                              i.jsx("option", {
                                value: "retained",
                                children: "After — retained population",
                              }),
                              i.jsx("option", {
                                value: "excluded",
                                children: "Excluded — every later duplicate",
                              }),
                              i.jsx("option", {
                                value: "groups",
                                children: "Duplicate groups — every member",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "dup-scroll",
                        children: [
                          i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", {
                                      children: i.jsx(M, {
                                        children: "Source row / range",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(M, {
                                        children: "Decision / derived row",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(M, {
                                        children: "Group / first row / members",
                                      }),
                                    }),
                                    i.jsx("th", {
                                      children: i.jsx(M, {
                                        children: "Normalized typed key",
                                      }),
                                    }),
                                    Array.from(
                                      {
                                        length:
                                          ct(I.request.range).lastCol -
                                          ct(I.request.range).firstCol +
                                          1,
                                      },
                                      (le, ce) =>
                                        i.jsxs(
                                          "th",
                                          {
                                            children: [
                                              i.jsx(M, { children: "Column" }),
                                              " ",
                                              ut(
                                                ct(I.request.range).firstCol +
                                                  ce,
                                              ),
                                            ],
                                          },
                                          ce,
                                        ),
                                    ),
                                  ],
                                }),
                              }),
                              i.jsx("tbody", {
                                children: he.map((le) =>
                                  i.jsxs(
                                    "tr",
                                    {
                                      children: [
                                        i.jsxs("th", {
                                          scope: "row",
                                          children: [
                                            le.sourceRow,
                                            " · ",
                                            le.sourceRange,
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            le.disposition,
                                            " · ",
                                            le.outputRow,
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            le.groupId || "Unique",
                                            " · ",
                                            le.firstSourceRow,
                                            " ·",
                                            " ",
                                            le.groupSize,
                                          ],
                                        }),
                                        i.jsxs("td", {
                                          children: [
                                            i.jsx("code", {
                                              children: le.normalizedKey,
                                            }),
                                            le.blankKey &&
                                              i.jsxs("span", {
                                                children: [
                                                  " ",
                                                  i.jsx(M, {
                                                    children: "· Blank key",
                                                  }),
                                                ],
                                              }),
                                          ],
                                        }),
                                        le.cells.map((ce, ge) =>
                                          i.jsx(
                                            "td",
                                            { children: Se(ce, le.values[ge]) },
                                            ge,
                                          ),
                                        ),
                                      ],
                                    },
                                    le.sourceRow,
                                  ),
                                ),
                              }),
                            ],
                          }),
                          !ae.length &&
                            i.jsx("p", {
                              children: i.jsx(M, {
                                children: "No rows in this view.",
                              }),
                            }),
                        ],
                      }),
                      i.jsxs("nav", {
                        "aria-label": r("Duplicate preview pages"),
                        lang: t("Duplicate preview pages"),
                        "data-ui-label": "Duplicate preview pages",
                        children: [
                          i.jsx("button", {
                            disabled: de === 0,
                            onClick: () => R(de - 1),
                            children: i.jsx(M, { children: "Previous page" }),
                          }),
                          i.jsxs("span", {
                            role: "status",
                            children: [
                              i.jsx(M, { children: "Page " }),
                              de + 1,
                              " ",
                              i.jsx(M, { children: "of " }),
                              oe,
                              " · ",
                              ae.length,
                              " ",
                              i.jsx(M, { children: "rows in this view" }),
                            ],
                          }),
                          i.jsx("button", {
                            disabled: de + 1 >= oe,
                            onClick: () => R(de + 1),
                            children: i.jsx(M, { children: "Next page" }),
                          }),
                        ],
                      }),
                    ],
                  })
                : i.jsx("p", {
                    role: "status",
                    children: i.jsx(M, {
                      children:
                        "This preview is stale. Source locations, values and groups are withheld. Preview again in the current workbook and access context.",
                    }),
                  }),
              i.jsxs("div", {
                className: "dup-actions",
                children: [
                  i.jsx("button", {
                    onClick: () => O(null),
                    children: i.jsx(M, { children: "Discard preview" }),
                  }),
                  i.jsx("button", {
                    className: "dup-primary",
                    disabled: e.readOnly || !J,
                    onClick: te,
                    children: i.jsx(M, {
                      children: "Create reviewed population",
                    }),
                  }),
                ],
              }),
            ],
          }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(M, { children: "Scope and evidence limits" }),
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Maximum 500 selected cells. Calculated comparison requires a complete static scalar dependency closure within 1,000 cells, 5,000 scanned formulas and depth 8. Volatile, dynamic, array/spill and unresolved calculations are excluded. Formula work limits cannot be hidden by error handlers. Calculated sources and dependencies must always be populated to avoid mistaking runtime spill children for blanks. Use raw comparison for null or absent source cells.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Null and empty normalized text are blanks. Case folding uses Unicode lowercase; trimming removes outer whitespace only. Formatting, locale-specific collation and numerical display rounding do not merge distinct values. Sheet objects, table styles, conditional rules and discussion threads remain on the original source; direct cell metadata is copied. Source locators describe this snapshot and are not immutable lineage or live links. A protected source may be read, but read-only hosts and workbook structure protection block creation.",
              }),
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
const zz = new WeakMap(),
  K7 = (e) =>
    typeof (e == null ? void 0 : e.raw) == "string" &&
    e.raw.startsWith("=") &&
    !e.literal,
  n2 = (e) => /#REF!/i.test(e.replace(/"(?:[^"]|"")*"/g, '""')),
  nfe = (e, r) => (e.axis === "row" ? String(r + 1) : ut(r));
function ife(e, r, t, n = {}) {
  var f, p, m;
  if (n.readOnly) throw Error("This workbook is read-only.");
  const o = Vn(e),
    a = LC(e, r, t),
    s = Vn(a);
  if (
    e.sheets.reduce((v, g) => v + Object.values(g.cells).filter(K7).length, 0) >
    5e3
  )
    throw Error(
      "Structural preview can verify up to 5,000 formula cells. Reduce or separate the formula population before changing worksheet structure.",
    );
  const u = e.sheets.find((v) => v.id === r),
    l =
      t.moveBefore !== void 0
        ? `Move ${t.count} ${t.axis}${t.count === 1 ? "" : "s"} before ${t.moveBefore === (t.axis === "row" ? u.rowCount : u.colCount) ? "worksheet end" : nfe(t, t.moveBefore)}`
        : `${t.remove ? "Delete" : "Insert"} ${t.count} ${t.axis}${t.count === 1 ? "" : "s"}`,
    c = {
      sheetId: r,
      edit: { ...t },
      snapshot: JSON.stringify(e),
      action: l,
      deletedCells: 0,
      movedCells: 0,
      formulaChanges: [],
      referenceIssues: [],
      originChanges: [],
      metadataChanges: [],
    };
  for (const v of e.sheets) {
    const g = a.sheets.find((b) => b.id === v.id);
    for (const [b, y] of Object.entries(v.cells)) {
      const w = v.id === r ? cg(b, t) : b;
      if (!w) {
        c.deletedCells++;
        continue;
      }
      w !== b && c.movedCells++;
      const E = g.cells[w];
      K7(y) &&
        ((E == null ? void 0 : E.raw) !== y.raw &&
          c.formulaChanges.push({
            sheetId: v.id,
            sheetName: v.name,
            beforeAddress: b,
            address: w,
            before: String(y.raw),
            after: String((E == null ? void 0 : E.raw) ?? ""),
          }),
        s(v.id, w) === "#REF!" &&
          o(v.id, b) !== "#REF!" &&
          c.referenceIssues.push({
            kind: "cell",
            sheetName: v.name,
            address: w,
            before: String(y.raw),
            after: String((E == null ? void 0 : E.raw) ?? ""),
          }));
    }
    for (const b of g.conditionalFormats ?? []) {
      const y =
        (f = v.conditionalFormats) == null
          ? void 0
          : f.find((w) => w.id === b.id);
      b.formula &&
        n2(b.formula) &&
        !n2((y == null ? void 0 : y.formula) ?? "") &&
        c.referenceIssues.push({
          kind: "conditional rule",
          sheetName: g.name,
          address: b.range,
          before: (y == null ? void 0 : y.formula) ?? "",
          after: b.formula,
        });
    }
    for (const b of [
      "rowHeights",
      "columnWidths",
      "hiddenRows",
      "hiddenCols",
      "frozenRows",
      "frozenCols",
      "merges",
      "validations",
      "conditionalFormats",
      "tables",
      "pivots",
      "queries",
      "objects",
      "pageLayout",
      "outlineGroups",
    ])
      JSON.stringify(v[b]) !== JSON.stringify(g[b]) &&
        c.metadataChanges.push({ sheetName: v.name, field: b });
  }
  for (const v of a.definedNames ?? []) {
    const g =
      (p = e.definedNames) == null
        ? void 0
        : p.find((b) => b.name === v.name && b.scopeSheetId === v.scopeSheetId);
    n2(v.refersTo) &&
      !n2((g == null ? void 0 : g.refersTo) ?? "") &&
      c.referenceIssues.push({
        kind: "name",
        sheetName:
          ((m = e.sheets.find((b) => b.id === v.scopeSheetId)) == null
            ? void 0
            : m.name) ?? "Workbook",
        address: v.name,
        before: (g == null ? void 0 : g.refersTo) ?? "",
        after: v.refersTo,
      });
  }
  const h = new Map((e.rowOrigins ?? []).map((v) => [v.id, v]));
  if (t.axis === "row")
    for (const [v, g] of Object.entries(u.rowOriginIds ?? {})) {
      const b = cg(`A${Number(v) + 1}`, t),
        y = b ? _t(b).row + 1 : null;
      if (y === Number(v) + 1) continue;
      const w = h.get(g);
      w &&
        c.originChanges.push({
          id: g,
          sourceName: w.sourceName,
          sourceSheet: w.sourceSheet,
          originalRow: w.originalRow,
          beforeRow: Number(v) + 1,
          afterRow: y,
        });
    }
  return (zz.set(c, JSON.stringify(c)), c);
}
function ofe(e, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (zz.get(r) !== JSON.stringify(r))
    throw Error(
      "This structural preview was changed or restored from another session. Create a fresh preview before applying.",
    );
  if (JSON.stringify(e) !== r.snapshot)
    throw Error(
      "The workbook changed after this preview. Refresh the preview before applying.",
    );
  return LC(e, r.sheetId, r.edit);
}
function afe({
  workbook: e,
  sheetId: r,
  range: t,
  mode: n,
  readOnly: o,
  onClose: a,
  onWorkbookChange: s,
}) {
  const { text: d, language: u } = Dr(),
    l = n === "insertRow" ? "row" : "column",
    c = ct(t),
    h = l === "row" ? Mt.maxRows : Mt.maxCols,
    [f, p] = ne.useState(l === "row" ? String(c.firstRow + 1) : ut(c.firstCol)),
    [m, v] = ne.useState(1),
    [g, b] = ne.useState("insert"),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(null),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(0),
    [j, z] = ne.useState("formulas"),
    [V, D] = ne.useState(false),
    B = ne.useRef(null),
    _ = e.sheets.find((G) => G.id === r);
  ne.useEffect(() => {
    var H, Q;
    const G = document.activeElement;
    return (
      (Q = (H = B.current) == null ? void 0 : H.querySelector("select")) ==
        null || Q.focus(),
      () => (G == null ? void 0 : G.focus())
    );
  }, []);
  const C = () => {
      (A(null), D(false), F(""), O(0));
    },
    T = (G, H = false) => {
      if (H && G.trim().toLowerCase() === "end") return h;
      if (l === "row") {
        const Q = Number(G);
        if (!/^[1-9]\d*$/.test(G.trim()) || !Number.isInteger(Q) || Q > h)
          throw Error(
            "Enter an existing row number, or use end for the move destination.",
          );
        return Q - 1;
      }
      if (!/^[a-z]{1,3}$/i.test(G.trim()))
        throw Error(
          "Enter a column letter such as A or XFD, or use end for the move destination.",
        );
      return ct(G.trim().toUpperCase() + "1").firstCol;
    };
  function N() {
    try {
      F("");
      const G = ife(
        e,
        r,
        {
          axis: l,
          index: T(f),
          count: m,
          remove: g === "delete",
          ...(g === "move" ? { moveBefore: T(y, !0) } : {}),
        },
        { readOnly: o },
      );
      (A(G), D(!1), O(0));
    } catch (G) {
      (F(G.message), A(null));
    }
  }
  function R() {
    try {
      if (!E) throw Error("Create a preview first.");
      if (L && !V)
        throw Error(
          "Confirm the deletion and reference changes before applying.",
        );
      const G = ofe(e, E, { readOnly: o });
      (s(
        G,
        `${E.action}${E.referenceIssues.length ? `; ${E.referenceIssues.length} new reference error(s) — reviewed` : ""}`,
      ),
        a());
    } catch (G) {
      F(G.message);
    }
  }
  const L =
      !!E &&
      (E.deletedCells > 0 ||
        E.referenceIssues.length > 0 ||
        E.originChanges.some((G) => G.afterRow === null)),
    $ = E
      ? j === "errors"
        ? E.referenceIssues.map((G) => ({
            key: `${G.kind}-${G.sheetName}-${G.address}`,
            title: `${G.kind}: ${G.sheetName}!${G.address}`,
            detail: `${G.before} → ${G.after}`,
          }))
        : j === "origins"
          ? E.originChanges.map((G) => ({
              key: G.id + "-" + G.beforeRow,
              title: `Row ${G.beforeRow} → ${G.afterRow ?? "deleted"}`,
              detail: `${G.sourceName} · ${G.sourceSheet}, original row ${G.originalRow} · ID ${G.id}`,
            }))
          : E.formulaChanges.map((G) => ({
              key: G.sheetId + G.beforeAddress,
              title: `${G.sheetName}!${G.beforeAddress} → ${G.address}`,
              detail: `${G.before} → ${G.after}`,
            }))
      : [];
  let U = null;
  try {
    U = U1(e, r, l === "row" ? T(f) : c.firstRow);
  } catch {}
  return i.jsx("div", {
    className: "st-backdrop",
    onPointerDown: (G) => {
      G.target === G.currentTarget && a();
    },
    children: i.jsxs("section", {
      className: "st-dialog",
      ref: B,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": `Insert / Delete / Move ${l === "row" ? "Rows" : "Columns"}`,
      onKeyDown: (G) => {
        var H;
        if (
          (G.key === "Escape" && (G.stopPropagation(), a()), G.key === "Tab")
        ) {
          const Q = Array.from(
              ((H = B.current) == null
                ? void 0
                : H.querySelectorAll(
                    "button:not([disabled]),input:not([disabled]),select:not([disabled])",
                  )) ?? [],
            ),
            X = Q[0],
            J = Q.at(-1);
          G.shiftKey && document.activeElement === X
            ? (G.preventDefault(), J == null || J.focus())
            : !G.shiftKey &&
              document.activeElement === J &&
              (G.preventDefault(), X == null || X.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "st-header",
          children: [
            i.jsxs("h2", {
              children: [l === "row" ? "Rows" : "Columns", " · ", _.name],
            }),
            i.jsx("button", {
              "aria-label": d("Close structure tool"),
              onClick: a,
              lang: u("Close structure tool"),
              "data-ui-label": "Close structure tool",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "st-body",
          children: [
            i.jsxs("div", {
              className: "st-controls",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Operation" }),
                    i.jsxs("select", {
                      "aria-label": d("Structure operation"),
                      value: g,
                      onChange: (G) => {
                        (C(), b(G.target.value));
                      },
                      lang: u("Structure operation"),
                      "data-ui-label": "Structure operation",
                      children: [
                        i.jsx("option", {
                          value: "insert",
                          children: "Insert",
                        }),
                        i.jsx("option", {
                          value: "delete",
                          children: "Delete",
                        }),
                        i.jsx("option", {
                          value: "move",
                          children: "Move before",
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    l === "row" ? "First row" : "First column",
                    i.jsx("input", {
                      "aria-label": d("Structure start"),
                      value: f,
                      onChange: (G) => {
                        (C(), p(G.target.value));
                      },
                      lang: u("Structure start"),
                      "data-ui-label": "Structure start",
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Count" }),
                    i.jsx(ci, {
                      "aria-label": d("Structure count"),
                      min: 1,
                      max: 1e4,
                      value: m,
                      onValueChange: (G) => {
                        (C(), v(Number(G)));
                      },
                      lang: u("Structure count"),
                      "data-ui-label": "Structure count",
                    }),
                  ],
                }),
              ],
            }),
            g === "move" &&
              i.jsxs("label", {
                className: "st-destination",
                children: [
                  i.jsx(M, { children: "Before " }),
                  l,
                  " ",
                  i.jsx(M, { children: "(original position)" }),
                  i.jsx("input", {
                    "aria-label": d("Structure destination"),
                    value: y,
                    placeholder: l === "row" ? "10, or end" : "F, or end",
                    onChange: (G) => {
                      (C(), w(G.target.value));
                    },
                    lang: u("Structure destination"),
                    "data-ui-label": "Structure destination",
                  }),
                ],
              }),
            i.jsxs("p", {
              children: [
                i.jsx(M, { children: "Entire " }),
                l === "row" ? "rows" : "columns",
                i.jsx(M, {
                  children:
                    ", including hidden cells, move together. A move inserts the selection before the destination and shifts the intervening data.",
                }),
              ],
            }),
            i.jsxs("div", {
              className: "st-source",
              children: [
                i.jsx("strong", {
                  children: i.jsx(M, { children: "Source row" }),
                }),
                i.jsx("p", {
                  children: U
                    ? `${U.sourceName} · ${U.sourceSheet}, original row ${U.originalRow} · ID ${U.id}`
                    : "No source identity is recorded for this row. Inserted and editor-created rows do not acquire a fabricated file origin.",
                }),
              ],
            }),
            k &&
              i.jsx("p", { role: "alert", className: "st-error", children: k }),
            E &&
              i.jsxs("section", {
                "aria-label": d("Structure preview"),
                lang: u("Structure preview"),
                "data-ui-label": "Structure preview",
                children: [
                  i.jsx("h3", { children: E.action }),
                  i.jsxs("p", {
                    children: [
                      E.movedCells,
                      " ",
                      i.jsx(M, { children: "stored cells relocate · " }),
                      E.deletedCells,
                      " ",
                      i.jsx(M, { children: "cells deleted · " }),
                      E.formulaChanges.length,
                      " ",
                      i.jsx(M, { children: "formulas rewritten ·" }),
                      " ",
                      i.jsxs("strong", {
                        children: [
                          E.referenceIssues.length,
                          " ",
                          i.jsx(M, { children: "new reference errors" }),
                        ],
                      }),
                      " ",
                      "· ",
                      E.originChanges.length,
                      " ",
                      i.jsx(M, { children: "source-row associations change." }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(M, {
                        children:
                          "Original source IDs and original row numbers stay fixed. Deleted-row origins remain in the native source registry.",
                      }),
                      " ",
                      E.metadataChanges.length,
                      " ",
                      i.jsx(M, { children: "worksheet settings update." }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Review" }),
                      i.jsxs("select", {
                        "aria-label": d("Structure review category"),
                        value: j,
                        onChange: (G) => {
                          (z(G.target.value), O(0));
                        },
                        lang: u("Structure review category"),
                        "data-ui-label": "Structure review category",
                        children: [
                          i.jsxs("option", {
                            value: "formulas",
                            children: [
                              "Rewritten formulas (",
                              E.formulaChanges.length,
                              ")",
                            ],
                          }),
                          i.jsxs("option", {
                            value: "errors",
                            children: [
                              "New reference errors (",
                              E.referenceIssues.length,
                              ")",
                            ],
                          }),
                          i.jsxs("option", {
                            value: "origins",
                            children: [
                              "Source rows (",
                              E.originChanges.length,
                              ")",
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("div", {
                    className: "st-changes",
                    children:
                      $.length === 0
                        ? i.jsx("p", {
                            children: i.jsx(M, {
                              children: "No changes in this category.",
                            }),
                          })
                        : $.slice(I * 25, I * 25 + 25).map((G) =>
                            i.jsxs(
                              "article",
                              {
                                children: [
                                  i.jsx("strong", { children: G.title }),
                                  i.jsx("p", { children: G.detail }),
                                ],
                              },
                              G.key,
                            ),
                          ),
                  }),
                  $.length > 25 &&
                    i.jsxs("nav", {
                      "aria-label": d("Structure review pages"),
                      lang: u("Structure review pages"),
                      "data-ui-label": "Structure review pages",
                      children: [
                        i.jsx("button", {
                          disabled: I === 0,
                          onClick: () => O((G) => G - 1),
                          children: i.jsx(M, { children: "Previous" }),
                        }),
                        i.jsxs("span", {
                          children: [
                            I * 25 + 1,
                            "–",
                            Math.min($.length, (I + 1) * 25),
                            " ",
                            i.jsx(M, { children: "of" }),
                            " ",
                            $.length,
                          ],
                        }),
                        i.jsx("button", {
                          disabled: (I + 1) * 25 >= $.length,
                          onClick: () => O((G) => G + 1),
                          children: i.jsx(M, { children: "Next" }),
                        }),
                      ],
                    }),
                  E.metadataChanges.length > 0 &&
                    i.jsxs("details", {
                      children: [
                        i.jsx("summary", {
                          children: i.jsx(M, {
                            children: "Changed worksheet settings",
                          }),
                        }),
                        i.jsx("ul", {
                          children: E.metadataChanges.map((G) =>
                            i.jsxs(
                              "li",
                              {
                                children: [
                                  G.sheetName,
                                  ":",
                                  " ",
                                  {
                                    rowHeights: "Row heights",
                                    columnWidths: "Column widths",
                                    hiddenRows: "Hidden rows",
                                    hiddenCols: "Hidden columns",
                                    frozenRows: "Frozen rows",
                                    frozenCols: "Frozen columns",
                                    merges: "Merged cells",
                                    validations: "Validation rules",
                                    conditionalFormats: "Conditional formats",
                                    tables: "Tables",
                                    pivots: "PivotTables",
                                    queries: "Saved queries",
                                    objects: "Pictures, charts and shapes",
                                    pageLayout: "Print settings",
                                    outlineGroups: "Outline groups",
                                  }[G.field],
                                ],
                              },
                              G.sheetName + G.field,
                            ),
                          ),
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("p", {
              className: "st-hint",
              children: i.jsx(M, {
                children:
                  "Up to 10,000 positions and 5,000 formula cells can be reviewed. Split references or structures are refused. Native JSON retains row provenance; XLSX and CSV do not preserve the source registry.",
              }),
            }),
          ],
        }),
        i.jsxs("footer", {
          className: "st-footer",
          children: [
            E &&
              L &&
              i.jsxs("label", {
                className: "st-consent",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: V,
                    onChange: (G) => D(G.target.checked),
                  }),
                  i.jsx(M, {
                    children:
                      "I reviewed the deletions and resulting reference errors.",
                  }),
                ],
              }),
            i.jsxs("div", {
              children: [
                i.jsx("button", {
                  onClick: a,
                  children: i.jsx(M, { children: "Cancel" }),
                }),
                i.jsx("button", {
                  disabled: o,
                  onClick: N,
                  children: E ? "Refresh preview" : "Preview changes",
                }),
                i.jsx("button", {
                  className: "st-primary",
                  disabled: !E || o || (L && !V),
                  onClick: R,
                  children: i.jsx(M, { children: "Apply changes" }),
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}

export { $z, afe };
