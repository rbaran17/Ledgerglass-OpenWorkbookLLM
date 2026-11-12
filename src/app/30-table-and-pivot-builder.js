import i from "react/jsx-runtime";
import ne from "react";
import {
  $P,
  Eie,
  FP,
  QP,
  ZP,
  cR,
  die,
  fie,
  mie,
  pR,
  vie,
  yR,
  yie,
} from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import {
  BI,
  Xk,
  Y7,
  cfe,
  ffe,
  hfe,
  ufe,
} from "./29-conditional-formatting-dialog.js";

function pfe({
  mode: e,
  workbook: r,
  sheetId: t,
  range: n,
  readOnly: o,
  onClose: a,
  onWorkbookChange: s,
}) {
  var br;
  const { text: d, language: u } = Dr(),
    l = ne.useRef(null);
  ne.useEffect(() => {
    const rt = document.activeElement,
      Le = l.current,
      et = Le?.querySelector(
        "input:not(:disabled), select:not(:disabled), button:not(:disabled)",
      );
    return (
      et == null || et.focus(),
      () => {
        rt != null && rt.isConnected && rt.focus();
      }
    );
  }, []);
  const c = ne.useMemo(() => {
      try {
        return die(r, t, n);
      } catch {
        return n;
      }
    }, []),
    h = ne.useMemo(() => hfe(r, t, c), []),
    [f, p] = ne.useState(t),
    [m, v] = ne.useState(c),
    [g, b] = ne.useState(h),
    [y, w] = ne.useState(
      e === "table"
        ? Y7(r)
        : e === "pivot"
          ? "Pivot summary"
          : "Extracted data",
    ),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState("blue"),
    [I, O] = ne.useState(false),
    [j, z] = ne.useState({}),
    [V, D] = ne.useState(h.slice(0, 1)),
    [B, _] = ne.useState([]),
    [C, T] = ne.useState(
      h.length ? [{ field: h.at(-1), operation: "sum" }] : [],
    ),
    [N, R] = ne.useState(true),
    [L, $] = ne.useState(false),
    [U, G] = ne.useState([]),
    [H, Q] = ne.useState([]);
  function X(rt) {
    (T((Le) =>
      Le.flatMap((et) => {
        const st = H.findIndex((kt) => kt.name === et.field);
        return st < 0
          ? [et]
          : rt.length === H.length
            ? [{ ...et, field: rt[st].name }]
            : rt.some((kt) => kt.name === et.field)
              ? [et]
              : [];
      }),
    ),
      Q(rt));
  }
  const [J, q] = ne.useState([]),
    [W, K] = ne.useState(h),
    [te, ae] = ne.useState({}),
    [oe, de] = ne.useState(""),
    [he, Se] = ne.useState("ascending"),
    [le, ce] = ne.useState([]),
    [ge, se] = ne.useState(false),
    [we, Ee] = ne.useState(t),
    [Ae, ye] = ne.useState("A1"),
    [Te, Pe] = ne.useState(""),
    [qe, ht] = ne.useState(""),
    [Re, Me] = ne.useState(""),
    [ke, je] = ne.useState("1"),
    [$e, Je] = ne.useState(null),
    [ze, at] = ne.useState(0),
    _e = JSON.stringify({
      sheetId: t,
      range: n,
      tableId: E,
      name: y,
      sourceRange: m,
      style: k,
      showTotals: I,
      totals: j,
      columnId: qe,
      columnFormula: Re,
      appendCount: ke,
    }),
    Ze = !!$e && ($e.plan.snapshot !== JSON.stringify(r) || $e.context !== _e),
    [lt, be] = ne.useState(""),
    [ft, pt] = ne.useState(""),
    [Fe, nt] = ne.useState(null),
    mt = r.sheets.find((rt) => rt.id === t)?.tables ?? [],
    Wt = r.sheets.flatMap((rt) =>
      (e === "pivot" ? (rt.pivots ?? []) : (rt.queries ?? [])).map((Le) => ({
        sheetName: rt.name,
        definition: Le,
      })),
    ),
    Zt = () => (ge ? { sheetId: we, address: Ae } : undefined);
  function xt(rt, Le) {
    if (!o) {
      (be(""), pt(""));
      try {
        const et = rt();
        (s(et, Le), pt(`${Le} completed.`));
      } catch (et) {
        be(
          et instanceof Error
            ? et.message
            : "The operation could not be completed.",
        );
      }
    }
  }
  function or() {
    be("");
    try {
      const rt = pR(r, f, m);
      (b(rt),
        K(rt),
        D(rt.slice(0, 1)),
        _([]),
        T(rt.length ? [{ field: rt.at(-1), operation: "sum" }] : []),
        q([]),
        ae({}),
        ce([]),
        de(""),
        G([]),
        Q([]));
    } catch (rt) {
      be(rt.message);
    }
  }
  function Wr(rt) {
    (nt(null), A(rt), Je(null));
    const Le = mt.find((kt) => kt.id === rt);
    if (!Le) {
      (w(Y7(r)), v(c), F("blue"), O(false), z({}));
      return;
    }
    (w(Le.name),
      ht(Le.columns[0]?.id ?? ""),
      Me(Le.columns[0]?.calculatedFormula ?? ""),
      v(Le.range),
      b(Le.columns.map((kt) => kt.name)),
      F(Le.style),
      O(Le.showTotals),
      z(Le.totals));
  }
  function Rr() {
    return J.map((rt) => ({
      ...rt,
      value: ["gt", "gte", "lt", "lte"].includes(rt.operator)
        ? Number(rt.value)
        : rt.value,
    }));
  }
  function qt(rt) {
    (be(""), pt(""), Je(null), at(0));
    try {
      Je({ plan: QP(r, t, E, rt, { readOnly: o }), context: _e });
    } catch (Le) {
      be(Le.message);
    }
  }
  function Ye() {
    if ($e) {
      be("");
      try {
        if (Ze)
          throw Error(
            "Table settings or the current selection changed. Preview again.",
          );
        const rt = Eie(r, $e.plan, { readOnly: o });
        (rt !== r && s(rt, "Apply reviewed table operation"),
          v($e.plan.afterRange),
          Je(null),
          pt(
            rt === r
              ? "No table changes were needed."
              : "Reviewed table operation completed.",
          ));
      } catch (rt) {
        be(rt.message);
      }
    }
  }
  function bt() {
    e === "table"
      ? E
        ? qt({
            kind: "updateTable",
            options: { name: y, range: m, style: k, showTotals: I, totals: j },
          })
        : xt(
            () =>
              fie(r, t, {
                name: y,
                range: m,
                style: k,
                showTotals: I,
                totals: j,
                ...(Fe
                  ? {
                      headers: (() => {
                        if (Fe.snapshot !== JSON.stringify(r) || Fe.range !== m)
                          throw Error(
                            "The workbook or source range changed. Review headers again.",
                          );
                        return Fe.headers.map((rt) => rt.suggested);
                      })(),
                    }
                  : {}),
              }),
            "Create table",
          )
      : e === "pivot"
        ? xt(
            () =>
              yR(r, {
                name: y,
                source: { sheetId: f, range: m },
                rowFields: V,
                columnFields: B,
                values: C,
                filters: Rr(),
                grandTotals: N,
                subtotals: L,
                dateGroups: U,
                calculatedFields: H,
                destination: Zt(),
              }),
            "Create pivot summary",
          )
        : xt(
            () =>
              ZP(r, {
                name: y,
                source: { sheetId: f, range: m },
                fields: W,
                rename: te,
                filters: Rr(),
                ...(oe ? { sort: { field: oe, direction: he } } : {}),
                derived: le,
                destination: Zt(),
              }),
            "Extract data",
          );
  }
  const St = (rt) =>
      i.jsxs(i.Fragment, {
        children: [
          rt !== undefined && i.jsx("option", { value: "", children: rt }),
          g.map((Le) => i.jsx("option", { value: Le, children: Le }, Le)),
        ],
      }),
    cr = (rt, Le, et) =>
      i.jsxs("fieldset", {
        className: "lg-table-fields",
        children: [
          i.jsx("legend", { children: rt }),
          g.length
            ? g.map((st) =>
                i.jsxs(
                  "label",
                  {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: Le.includes(st),
                        onChange: () => et(ffe(Le, st)),
                      }),
                      st,
                    ],
                  },
                  st,
                ),
              )
            : i.jsx("p", {
                children: i.jsx(M, {
                  children: "Read the source fields first.",
                }),
              }),
        ],
      }),
    dr =
      e === "table"
        ? "Table tools"
        : e === "pivot"
          ? "PivotTable builder"
          : "From Table / Range";
  return i.jsxs("div", {
    ref: l,
    className: "lg-table-tools",
    role: "dialog",
    "aria-modal": "true",
    "aria-label": dr,
    onKeyDown: (rt) => {
      if (
        (rt.key === "Escape" &&
          (rt.preventDefault(), rt.stopPropagation(), a()),
        rt.key === "Tab")
      ) {
        const et = Array.from(
            l.current?.querySelectorAll(
              "button:not(:disabled), input:not(:disabled), select:not(:disabled), summary, [tabindex='0']",
            ) ?? [],
          ).filter(
            (gr) =>
              (!gr.closest("fieldset:disabled") &&
                !gr.closest("details:not([open])")) ||
              gr.tagName === "SUMMARY",
          ),
          st = et[0],
          kt = et.at(-1);
        rt.shiftKey && document.activeElement === st && kt
          ? (rt.preventDefault(), kt.focus())
          : !rt.shiftKey &&
            document.activeElement === kt &&
            st &&
            (rt.preventDefault(), st.focus());
      }
    },
    children: [
      i.jsxs("header", {
        children: [
          i.jsxs("div", {
            children: [
              i.jsx("h2", { children: dr }),
              i.jsx("p", {
                children:
                  e === "table"
                    ? "Organize a headed range with a name, style and totals."
                    : e === "pivot"
                      ? "Group source values into a summary with refreshable settings."
                      : "Select, filter, sort and transform values into a separate result.",
              }),
            ],
          }),
          i.jsx("button", {
            type: "button",
            "aria-label": `Close ${dr}`,
            onClick: a,
            children: "×",
          }),
        ],
      }),
      lt &&
        i.jsx("p", {
          role: "alert",
          className: "lg-table-error",
          children: lt,
        }),
      ft && i.jsx("p", { role: "status", children: ft }),
      o &&
        i.jsx("p", {
          className: "lg-table-readonly",
          children: i.jsx(M, { children: "This workbook is read-only." }),
        }),
      i.jsxs("fieldset", {
        disabled: o,
        className: "lg-table-form",
        children: [
          e === "table" &&
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Existing table" }),
                i.jsxs("select", {
                  "aria-label": d("Existing table"),
                  value: E,
                  onChange: (rt) => Wr(rt.target.value),
                  lang: u("Existing table"),
                  "data-ui-label": "Existing table",
                  children: [
                    i.jsx("option", {
                      value: "",
                      children: "Create a new table",
                    }),
                    mt.map((rt) =>
                      i.jsx(
                        "option",
                        { value: rt.id, children: rt.name },
                        rt.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
          i.jsxs("label", {
            children: [
              e === "table" ? "Table name" : "Result name",
              i.jsx("input", {
                "aria-label": e === "table" ? "Table name" : "Result name",
                value: y,
                onChange: (rt) => w(rt.target.value),
              }),
            ],
          }),
          e !== "table" &&
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Source sheet" }),
                i.jsx("select", {
                  value: f,
                  onChange: (rt) => p(rt.target.value),
                  children: r.sheets.map((rt) =>
                    i.jsx("option", { value: rt.id, children: rt.name }, rt.id),
                  ),
                }),
              ],
            }),
          i.jsxs("div", {
            className: "lg-table-inline",
            children: [
              i.jsxs("label", {
                children: [
                  e === "table" ? "Table range" : "Source range",
                  i.jsx("input", {
                    "aria-label":
                      e === "table" ? "Table range" : "Source range",
                    value: m,
                    onChange: (rt) => v(rt.target.value),
                  }),
                ],
              }),
              i.jsx("button", {
                type: "button",
                onClick: or,
                children: i.jsx(M, { children: "Read fields" }),
              }),
            ],
          }),
          e === "table"
            ? i.jsxs(i.Fragment, {
                children: [
                  !E &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("button", {
                          type: "button",
                          onClick: () => {
                            be("");
                            try {
                              const rt = $P(r, t, m);
                              (nt({
                                snapshot: JSON.stringify(r),
                                range: m,
                                headers: rt,
                              }),
                                b(rt.map((Le) => Le.suggested)));
                            } catch (rt) {
                              be(rt.message);
                            }
                          },
                          children: i.jsx(M, { children: "Review headers" }),
                        }),
                        Fe &&
                          i.jsxs("section", {
                            "aria-label": d("Table header resolution"),
                            className: "lg-table-list",
                            lang: u("Table header resolution"),
                            "data-ui-label": "Table header resolution",
                            children: [
                              i.jsx("h3", {
                                children: i.jsx(M, {
                                  children: "Resolve table headers",
                                }),
                              }),
                              i.jsx("p", {
                                children: i.jsx(M, {
                                  children:
                                    "Review each source value and replacement. Create table applies these names and retains changed originals in native source evidence.",
                                }),
                              }),
                              Fe.headers.map((rt, Le) =>
                                i.jsxs(
                                  "label",
                                  {
                                    children: [
                                      rt.address,
                                      " ",
                                      i.jsx(M, { children: "original: " }),
                                      JSON.stringify(rt.original),
                                      i.jsx("input", {
                                        "aria-label": `Reviewed header ${Le + 1}`,
                                        value: rt.suggested,
                                        onChange: (et) => {
                                          const st = Fe.headers.map((kt, gr) =>
                                            gr === Le
                                              ? {
                                                  ...kt,
                                                  suggested: et.target.value,
                                                }
                                              : kt,
                                          );
                                          (nt({ ...Fe, headers: st }),
                                            b(st.map((kt) => kt.suggested)));
                                        },
                                      }),
                                    ],
                                  },
                                  rt.address,
                                ),
                              ),
                            ],
                          }),
                      ],
                    }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Table style" }),
                      i.jsxs("select", {
                        value: k,
                        onChange: (rt) => F(rt.target.value),
                        children: [
                          i.jsx("option", { value: "blue", children: "Blue" }),
                          i.jsx("option", {
                            value: "green",
                            children: "Green",
                          }),
                          i.jsx("option", { value: "gray", children: "Gray" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-table-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: I,
                        onChange: (rt) => O(rt.target.checked),
                      }),
                      i.jsx(M, { children: "Show totals row" }),
                    ],
                  }),
                  I &&
                    i.jsx("div", {
                      className: "lg-table-list",
                      children: g.map((rt) =>
                        i.jsxs(
                          "label",
                          {
                            children: [
                              rt,
                              " ",
                              i.jsx(M, { children: "total" }),
                              i.jsxs("select", {
                                "aria-label": `${rt} total`,
                                value: Object.hasOwn(j, rt) ? j[rt] : "",
                                onChange: (Le) =>
                                  z((et) => {
                                    const st = { ...et };
                                    return (
                                      Le.target.value
                                        ? Object.defineProperty(st, rt, {
                                            value: Le.target.value,
                                            enumerable: true,
                                            writable: true,
                                            configurable: true,
                                          })
                                        : delete st[rt],
                                      st
                                    );
                                  }),
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "None",
                                  }),
                                  cR.map((Le) =>
                                    i.jsx(
                                      "option",
                                      { value: Le, children: Xk(Le) },
                                      Le,
                                    ),
                                  ),
                                ],
                              }),
                            ],
                          },
                          rt,
                        ),
                      ),
                    }),
                  E &&
                    i.jsx(
                      ufe,
                      {
                        workbook: r,
                        sheetId: t,
                        tableId: E,
                        context: _e,
                        readOnly: o,
                        onWorkbookChange: s,
                      },
                      E,
                    ),
                  E &&
                    i.jsxs("fieldset", {
                      className: "lg-table-list",
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(M, {
                            children: "Calculated columns and rows",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Calculated column" }),
                            i.jsx("select", {
                              "aria-label": d("Calculated column"),
                              value: qe,
                              onChange: (rt) => {
                                (ht(rt.target.value),
                                  Me(
                                    mt
                                      .find((st) => st.id === E)
                                      ?.columns.find(
                                        (st) => st.id === rt.target.value,
                                      )?.calculatedFormula ?? "",
                                  ),
                                  Je(null));
                              },
                              lang: u("Calculated column"),
                              "data-ui-label": "Calculated column",
                              children: mt
                                .find((rt) => rt.id === E)
                                ?.columns.map((rt) =>
                                  i.jsxs(
                                    "option",
                                    {
                                      value: rt.id,
                                      children: [
                                        rt.name,
                                        rt.calculatedFormula
                                          ? " · calculated"
                                          : "",
                                      ],
                                    },
                                    rt.id,
                                  ),
                                ),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "First data-row formula" }),
                            i.jsx("input", {
                              "aria-label": d("First data-row formula"),
                              value: Re,
                              onChange: (rt) => Me(rt.target.value),
                              placeholder: d("=[@Quantity]*[@Rate]"),
                              lang: u("First data-row formula"),
                              "data-ui-label": "First data-row formula",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "lg-table-help",
                          children: i.jsx(M, {
                            children:
                              "The formula is based on this column’s first data row. Preview lists each overwritten value. Other columns and existing evidence stay intact.",
                          }),
                        }),
                        i.jsxs("div", {
                          className: "lg-table-inline",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              onClick: () =>
                                qt({
                                  kind: "setColumnFormula",
                                  columnId: qe,
                                  formula: Re,
                                }),
                              children: i.jsx(M, {
                                children: "Preview column formula",
                              }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              disabled: !(
                                (br = mt
                                  .find((rt) => rt.id === E)
                                  ?.columns.find((rt) => rt.id === qe)) !=
                                  null && br.calculatedFormula
                              ),
                              onClick: () =>
                                qt({
                                  kind: "clearColumnFormula",
                                  columnId: qe,
                                }),
                              children: i.jsx(M, {
                                children: "Preview clear rule",
                              }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Rows to append" }),
                            i.jsx(ci, {
                              "aria-label": d("Rows to append"),
                              min: "1",
                              max: "100",
                              value: ke,
                              onValueChange: (rt) => je(rt),
                              lang: u("Rows to append"),
                              "data-ui-label": "Rows to append",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () =>
                            qt({ kind: "appendRows", count: Number(ke) }),
                          children: i.jsx(M, {
                            children: "Preview append rows",
                          }),
                        }),
                        i.jsx("p", {
                          className: "lg-table-help",
                          children: i.jsx(M, {
                            children:
                              "Append 1–100 rows within the 500-cell change limit. Saved rules fill added rows. Adjacent cells must be empty. Clearing a rule keeps its existing cell formulas. Ordinary typing does not create rules or expand tables.",
                          }),
                        }),
                      ],
                    }),
                  $e &&
                    i.jsxs("section", {
                      className: "lg-table-preview",
                      "aria-label": d("Table change preview"),
                      lang: u("Table change preview"),
                      "data-ui-label": "Table change preview",
                      children: [
                        i.jsx("h3", {
                          children: i.jsx(M, {
                            children: "Review table changes",
                          }),
                        }),
                        Ze
                          ? i.jsx("p", {
                              role: "alert",
                              children: i.jsx(M, {
                                children:
                                  "The workbook, selection or table settings changed. Preview again.",
                              }),
                            })
                          : i.jsxs(i.Fragment, {
                              children: [
                                i.jsxs("p", {
                                  children: [
                                    $e.plan.beforeRange,
                                    " →",
                                    " ",
                                    $e.plan.afterRange,
                                    ";",
                                    " ",
                                    $e.plan.changes.length,
                                    " ",
                                    i.jsx(M, { children: "cell changes." }),
                                  ],
                                }),
                                i.jsxs("details", {
                                  children: [
                                    i.jsx("summary", {
                                      children: i.jsx(M, {
                                        children:
                                          "Saved column rules before and after",
                                      }),
                                    }),
                                    $e.plan.afterColumns.map((rt, Le) => {
                                      return i.jsxs(
                                        "p",
                                        {
                                          children: [
                                            i.jsx("strong", {
                                              children: rt.name,
                                            }),
                                            i.jsx("br", {}),
                                            i.jsx(M, { children: "Before:" }),
                                            " ",
                                            $e.plan.beforeColumns[Le]
                                              ?.calculatedFormula ?? "No rule",
                                            i.jsx("br", {}),
                                            i.jsx(M, { children: "After: " }),
                                            rt.calculatedFormula ?? "No rule",
                                          ],
                                        },
                                        rt.id,
                                      );
                                    }),
                                  ],
                                }),
                                i.jsx("ol", {
                                  start: ze * 20 + 1,
                                  children: $e.plan.changes
                                    .slice(ze * 20, ze * 20 + 20)
                                    .map((rt) => {
                                      return i.jsxs(
                                        "li",
                                        {
                                          children: [
                                            i.jsxs("strong", {
                                              children: [
                                                r.sheets.find(
                                                  (kt) => kt.id === rt.sheetId,
                                                )?.name,
                                                "!",
                                                rt.address,
                                              ],
                                            }),
                                            i.jsx("br", {}),
                                            i.jsx(M, { children: "Before: " }),
                                            JSON.stringify(
                                              rt.before?.raw ?? null,
                                            ),
                                            i.jsx("br", {}),
                                            i.jsx(M, { children: "After: " }),
                                            JSON.stringify(
                                              rt.after?.raw ?? null,
                                            ),
                                            i.jsxs("details", {
                                              children: [
                                                i.jsx("summary", {
                                                  children: i.jsx(M, {
                                                    children:
                                                      "Full cell metadata",
                                                  }),
                                                }),
                                                i.jsx("pre", {
                                                  children: JSON.stringify(
                                                    {
                                                      before: rt.before ?? null,
                                                      after: rt.after ?? null,
                                                    },
                                                    null,
                                                    2,
                                                  ),
                                                }),
                                              ],
                                            }),
                                          ],
                                        },
                                        `${rt.sheetId}:${rt.address}`,
                                      );
                                    }),
                                }),
                                $e.plan.changes.length > 20 &&
                                  i.jsxs("div", {
                                    className: "lg-table-inline",
                                    children: [
                                      i.jsx("button", {
                                        type: "button",
                                        disabled: ze === 0,
                                        onClick: () => at((rt) => rt - 1),
                                        children: i.jsx(M, {
                                          children: "Previous changes",
                                        }),
                                      }),
                                      i.jsxs("span", {
                                        children: [
                                          i.jsx(M, { children: "Page " }),
                                          ze + 1,
                                          " ",
                                          i.jsx(M, { children: "of" }),
                                          " ",
                                          Math.ceil(
                                            $e.plan.changes.length / 20,
                                          ),
                                        ],
                                      }),
                                      i.jsx("button", {
                                        type: "button",
                                        disabled:
                                          (ze + 1) * 20 >=
                                          $e.plan.changes.length,
                                        onClick: () => at((rt) => rt + 1),
                                        children: i.jsx(M, {
                                          children: "Next changes",
                                        }),
                                      }),
                                    ],
                                  }),
                                $e.plan.anchorChanges.map((rt) =>
                                  i.jsxs(
                                    "p",
                                    {
                                      children: [
                                        rt.kind,
                                        " ",
                                        rt.id,
                                        ": ",
                                        rt.before,
                                        " →",
                                        " ",
                                        rt.after,
                                      ],
                                    },
                                    rt.kind + rt.id,
                                  ),
                                ),
                              ],
                            }),
                        i.jsxs("div", {
                          className: "lg-table-inline",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              className: "lg-table-primary",
                              disabled: Ze,
                              onClick: Ye,
                              children: i.jsx(M, {
                                children: "Apply table changes",
                              }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              onClick: () => Je(null),
                              children: i.jsx(M, {
                                children: "Cancel table preview",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsx("p", {
                    className: "lg-table-help",
                    children: i.jsx(M, {
                      children:
                        "Creating a totals row requires an empty row below the source. When resizing an existing table, include its totals row in the range.",
                    }),
                  }),
                ],
              })
            : i.jsxs(i.Fragment, {
                children: [
                  e === "pivot"
                    ? i.jsxs(i.Fragment, {
                        children: [
                          cr("Row groups", V, D),
                          cr("Column groups", B, _),
                          i.jsxs("fieldset", {
                            className: "lg-table-list",
                            children: [
                              i.jsx("legend", {
                                children: i.jsx(M, {
                                  children: "Value fields",
                                }),
                              }),
                              C.map((rt, Le) =>
                                i.jsxs(
                                  "div",
                                  {
                                    className: "lg-table-value",
                                    children: [
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(M, {
                                            children: "Value field ",
                                          }),
                                          Le + 1,
                                          i.jsx("select", {
                                            value: rt.field,
                                            onChange: (et) =>
                                              T((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        field: et.target.value,
                                                      }
                                                    : kt,
                                                ),
                                              ),
                                            children: g
                                              .concat(H.map((et) => et.name))
                                              .map((et) =>
                                                i.jsx(
                                                  "option",
                                                  { value: et, children: et },
                                                  et,
                                                ),
                                              ),
                                          }),
                                        ],
                                      }),
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(M, {
                                            children: "Summarize by",
                                          }),
                                          i.jsx("select", {
                                            "aria-label": `Aggregation ${Le + 1}`,
                                            value: rt.operation,
                                            onChange: (et) =>
                                              T((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        operation:
                                                          et.target.value,
                                                      }
                                                    : kt,
                                                ),
                                              ),
                                            children: FP.map((et) =>
                                              i.jsx(
                                                "option",
                                                { value: et, children: Xk(et) },
                                                et,
                                              ),
                                            ),
                                          }),
                                        ],
                                      }),
                                      i.jsx("button", {
                                        type: "button",
                                        "aria-label": `Remove value ${Le + 1}`,
                                        onClick: () =>
                                          T((et) =>
                                            et.filter((st, kt) => kt !== Le),
                                          ),
                                        children: i.jsx(M, {
                                          children: "Remove",
                                        }),
                                      }),
                                    ],
                                  },
                                  Le,
                                ),
                              ),
                              i.jsx("button", {
                                type: "button",
                                disabled: !g.length || C.length >= 100,
                                onClick: () =>
                                  T((rt) => [
                                    ...rt,
                                    { field: g.at(-1), operation: "sum" },
                                  ]),
                                children: i.jsx(M, {
                                  children: "Add value field",
                                }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            className: "lg-table-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: N,
                                onChange: (rt) => R(rt.target.checked),
                              }),
                              i.jsx(M, { children: "Grand totals" }),
                            ],
                          }),
                          i.jsxs("label", {
                            className: "lg-table-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: L,
                                onChange: (rt) => $(rt.target.checked),
                              }),
                              i.jsx(M, {
                                children: "Row and column subtotals",
                              }),
                            ],
                          }),
                          i.jsx(cfe, {
                            fields: g,
                            groups: [...V, ...B],
                            dateGroups: U,
                            setDateGroups: G,
                            calculatedFields: H,
                            setCalculatedFields: X,
                            onAdd: (rt) =>
                              T((Le) => [
                                ...Le,
                                { field: rt, operation: "sum" },
                              ]),
                          }),
                        ],
                      })
                    : i.jsxs(i.Fragment, {
                        children: [
                          cr("Output fields", W, (rt) => {
                            (K(rt),
                              ae((Le) =>
                                Object.fromEntries(
                                  Object.entries(Le).filter(([et]) =>
                                    rt.includes(et),
                                  ),
                                ),
                              ));
                          }),
                          i.jsxs("details", {
                            children: [
                              i.jsx("summary", {
                                children: i.jsx(M, {
                                  children: "Rename output fields",
                                }),
                              }),
                              i.jsx("div", {
                                className: "lg-table-list",
                                children: W.map((rt) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        rt,
                                        i.jsx("input", {
                                          "aria-label": `Rename ${rt}`,
                                          value: Object.hasOwn(te, rt)
                                            ? te[rt]
                                            : rt,
                                          onChange: (Le) =>
                                            ae((et) =>
                                              Object.defineProperty(
                                                { ...et },
                                                rt,
                                                {
                                                  value: Le.target.value,
                                                  enumerable: true,
                                                  writable: true,
                                                  configurable: true,
                                                },
                                              ),
                                            ),
                                        }),
                                      ],
                                    },
                                    rt,
                                  ),
                                ),
                              }),
                            ],
                          }),
                          i.jsxs("div", {
                            className: "lg-table-inline",
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(M, { children: "Sort by" }),
                                  i.jsx("select", {
                                    value: oe,
                                    onChange: (rt) => de(rt.target.value),
                                    children: St("Source order"),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(M, { children: "Order" }),
                                  i.jsxs("select", {
                                    value: he,
                                    onChange: (rt) => Se(rt.target.value),
                                    children: [
                                      i.jsx("option", {
                                        value: "ascending",
                                        children: "Ascending",
                                      }),
                                      i.jsx("option", {
                                        value: "descending",
                                        children: "Descending",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("fieldset", {
                            className: "lg-table-list",
                            children: [
                              i.jsx("legend", {
                                children: i.jsx(M, {
                                  children: "Derived fields",
                                }),
                              }),
                              le.map((rt, Le) =>
                                i.jsxs(
                                  "div",
                                  {
                                    className: "lg-table-derived",
                                    children: [
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(M, {
                                            children: "New field name",
                                          }),
                                          i.jsx("input", {
                                            value: rt.name,
                                            onChange: (et) =>
                                              ce((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        name: et.target.value,
                                                      }
                                                    : kt,
                                                ),
                                              ),
                                          }),
                                        ],
                                      }),
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(M, {
                                            children: "Source field",
                                          }),
                                          i.jsx("select", {
                                            value: rt.field,
                                            onChange: (et) =>
                                              ce((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        field: et.target.value,
                                                      }
                                                    : kt,
                                                ),
                                              ),
                                            children: St(),
                                          }),
                                        ],
                                      }),
                                      i.jsxs("label", {
                                        children: [
                                          i.jsx(M, { children: "Operation" }),
                                          i.jsx("select", {
                                            value: rt.operation,
                                            onChange: (et) =>
                                              ce((st) =>
                                                st.map((kt, gr) =>
                                                  gr === Le
                                                    ? {
                                                        ...kt,
                                                        operation:
                                                          et.target.value,
                                                      }
                                                    : kt,
                                                ),
                                              ),
                                            children: [
                                              "add",
                                              "subtract",
                                              "multiply",
                                              "divide",
                                              "concat",
                                              "upper",
                                              "lower",
                                              "trim",
                                            ].map((et) =>
                                              i.jsx(
                                                "option",
                                                { value: et, children: Xk(et) },
                                                et,
                                              ),
                                            ),
                                          }),
                                        ],
                                      }),
                                      !["upper", "lower", "trim"].includes(
                                        rt.operation,
                                      ) &&
                                        i.jsxs(i.Fragment, {
                                          children: [
                                            i.jsxs("label", {
                                              children: [
                                                i.jsx(M, {
                                                  children: "Second field",
                                                }),
                                                i.jsx("select", {
                                                  value: rt.otherField ?? "",
                                                  onChange: (et) =>
                                                    ce((st) =>
                                                      st.map((kt, gr) =>
                                                        gr === Le
                                                          ? {
                                                              ...kt,
                                                              otherField:
                                                                et.target
                                                                  .value ||
                                                                undefined,
                                                            }
                                                          : kt,
                                                      ),
                                                    ),
                                                  children:
                                                    St("Use a constant"),
                                                }),
                                              ],
                                            }),
                                            !rt.otherField &&
                                              i.jsxs("label", {
                                                children: [
                                                  i.jsx(M, {
                                                    children: "Constant",
                                                  }),
                                                  rt.operation === "concat"
                                                    ? i.jsx("input", {
                                                        value:
                                                          rt.constant ?? "",
                                                        onChange: (et) =>
                                                          ce((st) =>
                                                            st.map((kt, gr) =>
                                                              gr === Le
                                                                ? {
                                                                    ...kt,
                                                                    constant:
                                                                      et.target
                                                                        .value,
                                                                  }
                                                                : kt,
                                                            ),
                                                          ),
                                                      })
                                                    : i.jsx(ci, {
                                                        "aria-label":
                                                          d("Constant"),
                                                        value:
                                                          rt.constant ?? "",
                                                        onValueChange: (et) =>
                                                          ce((st) =>
                                                            st.map((kt, gr) =>
                                                              gr === Le
                                                                ? {
                                                                    ...kt,
                                                                    constant:
                                                                      Number(
                                                                        et,
                                                                      ),
                                                                  }
                                                                : kt,
                                                            ),
                                                          ),
                                                      }),
                                                ],
                                              }),
                                          ],
                                        }),
                                      i.jsx("button", {
                                        type: "button",
                                        onClick: () =>
                                          ce((et) =>
                                            et.filter((st, kt) => kt !== Le),
                                          ),
                                        children: i.jsx(M, {
                                          children: "Remove derived field",
                                        }),
                                      }),
                                    ],
                                  },
                                  Le,
                                ),
                              ),
                              i.jsx("button", {
                                type: "button",
                                disabled: !g.length || le.length >= 100,
                                onClick: () =>
                                  ce((rt) => [
                                    ...rt,
                                    {
                                      name: `Calculated${rt.length + 1}`,
                                      field: g.at(-1),
                                      operation: "multiply",
                                      constant: 1,
                                    },
                                  ]),
                                children: i.jsx(M, {
                                  children: "Add derived field",
                                }),
                              }),
                            ],
                          }),
                        ],
                      }),
                  i.jsxs("fieldset", {
                    className: "lg-table-list",
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(M, { children: "Source filters" }),
                      }),
                      J.map((rt, Le) =>
                        i.jsxs(
                          "div",
                          {
                            className: "lg-table-filter",
                            children: [
                              i.jsxs("label", {
                                children: [
                                  i.jsx(M, { children: "Filter field " }),
                                  Le + 1,
                                  i.jsx("select", {
                                    value: rt.field,
                                    onChange: (et) =>
                                      q((st) =>
                                        st.map((kt, gr) =>
                                          gr === Le
                                            ? { ...kt, field: et.target.value }
                                            : kt,
                                        ),
                                      ),
                                    children: St(),
                                  }),
                                ],
                              }),
                              i.jsxs("label", {
                                children: [
                                  i.jsx(M, { children: "Condition" }),
                                  i.jsx("select", {
                                    value: rt.operator,
                                    onChange: (et) =>
                                      q((st) =>
                                        st.map((kt, gr) =>
                                          gr === Le
                                            ? {
                                                ...kt,
                                                operator: et.target.value,
                                              }
                                            : kt,
                                        ),
                                      ),
                                    children: [
                                      ["equals", "Equals"],
                                      ["notEquals", "Does not equal"],
                                      ["contains", "Contains"],
                                      ["gt", "Greater than"],
                                      ["gte", "At least"],
                                      ["lt", "Less than"],
                                      ["lte", "At most"],
                                      ["blank", "Is blank"],
                                      ["notBlank", "Is not blank"],
                                    ].map(([et, st]) =>
                                      i.jsx(
                                        "option",
                                        { value: et, children: st },
                                        et,
                                      ),
                                    ),
                                  }),
                                ],
                              }),
                              !["blank", "notBlank"].includes(rt.operator) &&
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Filter value" }),
                                    i.jsx("input", {
                                      value: String(rt.value ?? ""),
                                      onChange: (et) =>
                                        q((st) =>
                                          st.map((kt, gr) =>
                                            gr === Le
                                              ? {
                                                  ...kt,
                                                  value: et.target.value,
                                                }
                                              : kt,
                                          ),
                                        ),
                                    }),
                                  ],
                                }),
                              i.jsx("button", {
                                type: "button",
                                "aria-label": `Remove filter ${Le + 1}`,
                                onClick: () =>
                                  q((et) => et.filter((st, kt) => kt !== Le)),
                                children: i.jsx(M, { children: "Remove" }),
                              }),
                            ],
                          },
                          Le,
                        ),
                      ),
                      i.jsx("button", {
                        type: "button",
                        disabled: !g.length || J.length >= 100,
                        onClick: () =>
                          q((rt) => [
                            ...rt,
                            { field: g[0], operator: "equals", value: "" },
                          ]),
                        children: i.jsx(M, { children: "Add source filter" }),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    className: "lg-table-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: ge,
                        onChange: (rt) => se(rt.target.checked),
                      }),
                      i.jsx(M, {
                        children: "Use an empty area in an existing sheet",
                      }),
                    ],
                  }),
                  ge &&
                    i.jsxs("div", {
                      className: "lg-table-inline",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Destination sheet" }),
                            i.jsx("select", {
                              value: we,
                              onChange: (rt) => Ee(rt.target.value),
                              children: r.sheets.map((rt) =>
                                i.jsx(
                                  "option",
                                  { value: rt.id, children: rt.name },
                                  rt.id,
                                ),
                              ),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Destination cell" }),
                            i.jsx("input", {
                              value: Ae,
                              onChange: (rt) => ye(rt.target.value),
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsx("p", {
                    className: "lg-table-help",
                    children: i.jsx(M, {
                      children:
                        "Results preserve the source. An existing destination must be empty. Refresh creates a new result so earlier work remains available.",
                    }),
                  }),
                  Wt.length > 0 &&
                    i.jsxs("fieldset", {
                      className: "lg-table-list",
                      children: [
                        i.jsx("legend", {
                          children: i.jsx(M, {
                            children: "Refresh a saved result",
                          }),
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Saved result" }),
                            i.jsxs("select", {
                              value: Te,
                              onChange: (rt) => Pe(rt.target.value),
                              children: [
                                i.jsx("option", {
                                  value: "",
                                  children: "Choose a result",
                                }),
                                Wt.map(({ sheetName: rt, definition: Le }) =>
                                  i.jsxs(
                                    "option",
                                    {
                                      value: Le.id,
                                      children: [rt, ": ", Le.name],
                                    },
                                    Le.id,
                                  ),
                                ),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: !Te,
                          onClick: () =>
                            xt(
                              () =>
                                e === "pivot"
                                  ? mie(r, Te, Zt())
                                  : yie(r, Te, Zt()),
                              e === "pivot"
                                ? "Refresh pivot summary"
                                : "Refresh extracted data",
                            ),
                          children: i.jsx(M, {
                            children: "Refresh into a new result",
                          }),
                        }),
                        e === "pivot" &&
                          Te &&
                          i.jsx(BI, { workbook: r, pivotId: Te }),
                      ],
                    }),
                ],
              }),
          i.jsx("footer", {
            children: i.jsx("button", {
              className: "lg-table-primary",
              type: "button",
              onClick: bt,
              children:
                e === "table"
                  ? E
                    ? "Update table"
                    : "Create table"
                  : e === "pivot"
                    ? "Create pivot summary"
                    : "Extract data",
            }),
          }),
        ],
      }),
    ],
  });
}
const Vz = {
    value: "No calculation",
    percentGrand: "% of grand total",
    percentRow: "% of row total",
    percentColumn: "% of column total",
  },
  mfe = (e) =>
    e === null
      ? "(blank)"
      : e === ""
        ? "(empty text)"
        : `${typeof e}: ${String(e)}`;
function gfe({ book: e, id: r }) {
  const { text: t, language: n } = Dr();
  let o;
  try {
    o = vie(e, r);
  } catch (m) {
    return i.jsx("p", { role: "alert", children: m.message });
  }
  const a = o.points.map((m) => m.value),
    s = a.filter((m) => m !== null),
    d = a
      .reduce(
        (m, v, g) => (
          v === null ? m.push([]) : m[m.length - 1].push({ v, i: g }),
          m
        ),
        [[]],
      )
      .filter((m) => m.length),
    u = Math.min(0, ...s),
    l = Math.max(1, ...s),
    c = (m) => 180 - ((m - u) / (l - u)) * 145,
    h = 420 / Math.max(1, a.length),
    f = c(0),
    p = o.showAs !== "value";
  return i.jsxs("figure", {
    className: "lg-pivot-chart",
    "aria-label": t("PivotChart"),
    lang: n("PivotChart"),
    "data-ui-label": "PivotChart",
    children: [
      i.jsxs("figcaption", {
        children: [
          o.measure,
          p ? ` — ${Vz[o.showAs]}` : "",
          " ",
          i.jsx(M, { children: "· source revision" }),
          " ",
          o.sourceRevision,
        ],
      }),
      a.length
        ? i.jsxs("svg", {
            role: "img",
            "aria-label": `${o.kind} PivotChart with ${a.length} groups; values are listed below`,
            viewBox: "0 0 470 210",
            children: [
              o.kind === "bar"
                ? i.jsx("line", {
                    x1: 35 - (u / (l - u)) * 420,
                    x2: 35 - (u / (l - u)) * 420,
                    y1: "28",
                    y2: "180",
                    stroke: "#607780",
                  })
                : i.jsx("line", {
                    x1: "35",
                    x2: "455",
                    y1: f,
                    y2: f,
                    stroke: "#607780",
                  }),
              i.jsx("text", {
                x: o.kind === "bar" ? 455 : 30,
                y: o.kind === "bar" ? 198 : c(l) + 4,
                textAnchor: "end",
                fontSize: "10",
                children: p ? `${(l * 100).toFixed(0)}%` : l.toFixed(1),
              }),
              i.jsx("text", {
                x: o.kind === "bar" ? 35 : 30,
                y: o.kind === "bar" ? 198 : c(u) + 4,
                textAnchor: o.kind === "bar" ? "start" : "end",
                fontSize: "10",
                children: p ? `${(u * 100).toFixed(0)}%` : u.toFixed(1),
              }),
              o.kind === "line"
                ? d.map((m, v) =>
                    i.jsx(
                      "polyline",
                      {
                        fill: "none",
                        stroke: "#217346",
                        strokeWidth: "3",
                        points: m
                          .map(
                            ({ v: g, i: b }) => `${35 + h * (b + 0.5)},${c(g)}`,
                          )
                          .join(" "),
                      },
                      v,
                    ),
                  )
                : a.map((m, v) =>
                    m === null
                      ? null
                      : o.kind === "bar"
                        ? i.jsx(
                            "rect",
                            {
                              x: 35 + ((Math.min(m, 0) - u) / (l - u)) * 420,
                              y: 28 + v * (150 / a.length),
                              width: Math.max(1, (Math.abs(m) / (l - u)) * 420),
                              height: Math.max(1, 120 / a.length),
                              fill: "#217346",
                              children: i.jsxs("title", {
                                children: [o.points[v].label, ": ", m],
                              }),
                            },
                            v,
                          )
                        : i.jsx(
                            "rect",
                            {
                              x: 36 + h * v,
                              width: Math.max(1, h - 3),
                              y: Math.min(c(m), f),
                              height: Math.max(1, Math.abs(c(m) - f)),
                              fill: "#217346",
                              children: i.jsxs("title", {
                                children: [o.points[v].label, ": ", m],
                              }),
                            },
                            v,
                          ),
                  ),
            ],
          })
        : i.jsx("p", {
            children: i.jsx(M, { children: "No records match these filters." }),
          }),
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(M, { children: "Chart values and categories" }),
          }),
          i.jsxs("table", {
            children: [
              i.jsx("thead", {
                children: i.jsxs("tr", {
                  children: [
                    i.jsx("th", { children: i.jsx(M, { children: "Group" }) }),
                    i.jsx("th", { children: i.jsx(M, { children: "Value" }) }),
                  ],
                }),
              }),
              i.jsx("tbody", {
                children: o.points.map((m, v) =>
                  i.jsxs(
                    "tr",
                    {
                      children: [
                        i.jsx("td", { children: m.label }),
                        i.jsx("td", {
                          children:
                            m.value === null
                              ? "Blank"
                              : p
                                ? `${(Number(m.value) * 100).toFixed(2)}%`
                                : String(m.value),
                        }),
                      ],
                    },
                    v,
                  ),
                ),
              }),
            ],
          }),
        ],
      }),
      i.jsx("p", {
        children: i.jsx(M, {
          children:
            "Leaf row groups only; subtotals and grand totals are excluded. Column groups are combined using the selected aggregate.",
        }),
      }),
    ],
  });
}

export { Vz, gfe, mfe, pfe };
