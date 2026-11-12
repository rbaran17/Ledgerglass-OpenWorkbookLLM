import i from "react/jsx-runtime";
import { X, ShieldCheck, Grid2x2, Paintbrush, CheckCheck } from "lucide-react";
import ne from "react";
import {
  CE,
  Gne,
  HP,
  HQ,
  RP,
  TP,
  Ta,
  WQ,
  Zne,
  ct,
  hR,
  kE,
  pR,
  qQ,
  ut,
} from "./part-03.js";
import { Dr, M, ci } from "./part-04.js";
import { LC, qde } from "./part-27.js";
import { $z, afe } from "./part-28.js";

function sfe(e) {
  return e.mode === "validation"
    ? i.jsx(qde, { ...e })
    : e.mode === "duplicates"
      ? i.jsx($z, { ...e })
      : e.mode === "insertRow" || e.mode === "insertColumn"
        ? i.jsx(afe, { ...e })
        : i.jsx(lfe, { ...e });
}
function lfe({
  mode: e,
  workbook: r,
  sheetId: t,
  range: n,
  readOnly: o,
  onClose: a,
  onWorkbookChange: s,
}) {
  var Ee, Ae, ye, Te, Pe, qe, ht;
  const { text: d, language: u } = Dr(),
    l = r.sheets.find((Re) => Re.id === t),
    c = n.split(":")[0],
    [h, f] = ne.useState(e === "duplicates" && n === c ? Ta(l) : n),
    [p, m] = ne.useState(e === "conditional" ? "greater" : "list"),
    [v, g] = ne.useState(""),
    [b, y] = ne.useState("100"),
    [w, E] = ne.useState("#fff2cc"),
    [A, k] = ne.useState("#9c5700"),
    [F, I] = ne.useState("between"),
    [O, j] = ne.useState(true),
    [z, V] = ne.useState("Enter a value that satisfies this rule."),
    [D, B] = ne.useState(true),
    [_, C] = ne.useState(null),
    [T, N] = ne.useState(""),
    [R, L] = ne.useState(1),
    [$, U] = ne.useState(false),
    [G, H] = ne.useState(""),
    [Q, X$1] = ne.useState(""),
    [J, q] = ne.useState(false),
    [W, K] = ne.useState(""),
    te = ne.useRef(null),
    ae = Object.entries(l.cells)
      .filter(([Re, Me]) => {
        try {
          const ke = ct(h),
            je = ct(Re);
          return (
            typeof Me.raw == "string" &&
            !Me.raw.startsWith("=") &&
            je.firstRow >= ke.firstRow &&
            je.lastRow <= ke.lastRow &&
            je.firstCol >= ke.firstCol &&
            je.lastCol <= ke.lastCol
          );
        } catch {
          return false;
        }
      })
      .slice(0, 500),
    [oe, de] = ne.useState(0),
    [he, Se] = ne.useState(((Ee = ae[0]) == null ? void 0 : Ee[1].raw) ?? "");
  ne.useEffect(() => {
    var Me, ke;
    const Re = document.activeElement;
    return (
      (ke =
        (Me = te.current) == null
          ? void 0
          : Me.querySelector("input,textarea,select,button")) == null ||
        ke.focus(),
      () => (Re == null ? void 0 : Re.focus())
    );
  }, []);
  const le = {
      conditional: "Conditional Formatting",
      validation: "Data Validation",
      duplicates: "Remove Duplicates",
      protection:
        (Ae = l.protection) != null && Ae.enabled
          ? "Unprotect Sheet"
          : "Protect Sheet",
      merge: "Merge & Center",
      paint: "Format Painter",
      spelling: "Spelling",
      insertRow: "Insert / Delete Rows",
      insertColumn: "Insert / Delete Columns",
    },
    ce = (Re, Me) => {
      (s(Re, Me), a());
    };
  async function ge() {
    var Re, Me;
    (H(""), q(true));
    try {
      if (o && e !== "spelling") throw Error("This workbook is read-only.");
      if (e === "merge") ce(kE(r, t, h, !1), "Merge cells");
      else if (e === "paint") ce(qQ(r, t, c, W), "Paint cell formatting");
      else if (e === "protection") {
        const ke = await HQ(
          r,
          t,
          T,
          !((Re = l.protection) != null && Re.enabled),
        );
        ce(
          ke,
          (Me = l.protection) != null && Me.enabled
            ? "Unprotect sheet"
            : "Protect sheet",
        );
      } else if (e === "conditional") {
        if (
          ["greater", "less", "between"].includes(p) &&
          (!v.trim() ||
            !Number.isFinite(Number(v)) ||
            (p === "between" && !Number.isFinite(Number(b))))
        )
          throw Error("Enter finite comparison values.");
        const ke = {
          id: crypto.randomUUID(),
          range: CE(h),
          kind: p,
          value: v,
          second: b,
          fill: w,
          color: A,
        };
        ce(
          {
            ...r,
            sheets: r.sheets.map((je) =>
              je.id === t
                ? {
                    ...je,
                    conditionalFormats: [...(je.conditionalFormats ?? []), ke],
                  }
                : je,
            ),
          },
          "Add conditional formatting",
        );
      } else if (e === "validation") {
        const ke = {
          id: crypto.randomUUID(),
          range: CE(h),
          kind: p,
          operator: F,
          min: Number(v || 0),
          max: Number(b),
          list: v
            .split(",")
            .map((je) => je.trim())
            .filter(Boolean),
          allowBlank: O,
          message: z,
        };
        ce(
          {
            ...r,
            sheets: r.sheets.map((je) =>
              je.id === t
                ? { ...je, validations: [...(je.validations ?? []), ke] }
                : je,
            ),
          },
          "Add data validation",
        );
      } else if (e === "duplicates") {
        const ke = ct(h),
          je =
            _ ??
            Array.from(
              { length: ke.lastCol - ke.firstCol + 1 },
              (Je, ze) => ke.firstCol + ze,
            ),
          $e = WQ(r, t, h, je, D);
        if (!$e.removed) {
          X$1("No duplicate rows found.");
          return;
        }
        ce($e.workbook, `Remove ${$e.removed} duplicate rows from selection`);
      } else if (e === "insertRow" || e === "insertColumn") {
        const ke = ct(h),
          je = {
            axis: e === "insertRow" ? "row" : "column",
            index: e === "insertRow" ? ke.firstRow : ke.firstCol,
            count: R,
            remove: $,
          };
        ce(
          LC(r, t, je),
          `${$ ? "Delete" : "Insert"} ${R} ${je.axis}${R === 1 ? "" : "s"}`,
        );
      } else if (e === "spelling") {
        if (o) throw Error("This workbook is read-only.");
        const ke = ae[oe];
        if (!ke) throw Error("No text cells in the selected range.");
        const je = { ...l.cells, [ke[0]]: { ...ke[1], raw: he } };
        ce(
          hR(r, {
            ...r,
            sheets: r.sheets.map(($e) =>
              $e.id === t ? { ...$e, cells: je } : $e,
            ),
          }),
          `Spelling correction at ${ke[0]}`,
        );
      }
    } catch (ke) {
      H(
        ke instanceof Error
          ? ke.message
          : "The operation could not be completed.",
      );
    } finally {
      q(false);
    }
  }
  let se;
  try {
    se = ct(h);
  } catch {
    se = ct(n);
  }
  const we =
    se.lastCol - se.firstCol < 100
      ? Array.from(
          { length: se.lastCol - se.firstCol + 1 },
          (Re, Me) => se.firstCol + Me,
        )
      : [];
  return i.jsx("div", {
    className: "wt-backdrop",
    onPointerDown: (Re) => {
      Re.target === Re.currentTarget && a();
    },
    children: i.jsxs("div", {
      className: "wt-dialog",
      ref: te,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": le[e] ?? e,
      onKeyDown: (Re) => {
        var Me;
        if (
          (Re.key === "Escape" && (Re.stopPropagation(), a()), Re.key === "Tab")
        ) {
          const ke = Array.from(
              ((Me = te.current) == null
                ? void 0
                : Me.querySelectorAll(
                    "button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled])",
                  )) ?? [],
            ),
            je = ke[0],
            $e = ke.at(-1);
          Re.shiftKey && document.activeElement === je
            ? (Re.preventDefault(), $e == null || $e.focus())
            : !Re.shiftKey &&
              document.activeElement === $e &&
              (Re.preventDefault(), je == null || je.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("span", {
              children: [
                i.jsx(Grid2x2, { size: 19 }),
                i.jsx("h2", { children: le[e] ?? e }),
              ],
            }),
            i.jsx("button", {
              "aria-label": d("Close worksheet tool"),
              onClick: a,
              lang: u("Close worksheet tool"),
              "data-ui-label": "Close worksheet tool",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wt-body",
          children: [
            e !== "protection" &&
              e !== "paint" &&
              i.jsxs("label", {
                children: [
                  i.jsx(M, { children: "Range" }),
                  i.jsx("input", {
                    "aria-label": d("Tool range"),
                    value: h,
                    onChange: (Re) => {
                      (f(Re.target.value), C(null));
                    },
                    lang: u("Tool range"),
                    "data-ui-label": "Tool range",
                  }),
                ],
              }),
            e === "paint" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("p", {
                    children: [
                      i.jsx(Paintbrush, { size: 16 }),
                      " ",
                      i.jsx(M, { children: "Copy formatting from" }),
                      " ",
                      i.jsxs("strong", { children: [l.name, "!", c] }),
                      ".",
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Apply to range" }),
                      i.jsx("input", {
                        "aria-label": d("Format destination"),
                        value: W,
                        onChange: (Re) => K(Re.target.value),
                        placeholder: d("B2:D10"),
                        lang: u("Format destination"),
                        "data-ui-label": "Format destination",
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(M, {
                      children:
                        "Values, formulas, notes and review marks stay in the destination.",
                    }),
                  }),
                ],
              }),
            e === "merge" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    children: i.jsx(M, {
                      children:
                        "The selected rectangle becomes one displayed cell, centered on its first value. Covered values remain stored and reappear when unmerged.",
                    }),
                  }),
                  i.jsx("button", {
                    className: "wt-secondary",
                    disabled: o,
                    onClick: () => {
                      try {
                        ce(kE(r, t, h, !0), "Unmerge cells");
                      } catch (Re) {
                        H(String(Re));
                      }
                    },
                    children: i.jsx(M, { children: "Unmerge selected cells" }),
                  }),
                ],
              }),
            e === "conditional" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Rule" }),
                      i.jsxs("select", {
                        "aria-label": d("Conditional rule"),
                        value: p,
                        onChange: (Re) => m(Re.target.value),
                        lang: u("Conditional rule"),
                        "data-ui-label": "Conditional rule",
                        children: [
                          i.jsx("option", {
                            value: "greater",
                            children: "Greater than",
                          }),
                          i.jsx("option", {
                            value: "less",
                            children: "Less than",
                          }),
                          i.jsx("option", {
                            value: "equal",
                            children: "Equal to",
                          }),
                          i.jsx("option", {
                            value: "between",
                            children: "Between",
                          }),
                          i.jsx("option", {
                            value: "contains",
                            children: "Text contains",
                          }),
                          i.jsx("option", {
                            value: "duplicate",
                            children: "Duplicate values",
                          }),
                          i.jsx("option", {
                            value: "colorScale",
                            children: "Green color scale",
                          }),
                        ],
                      }),
                    ],
                  }),
                  !["duplicate", "colorScale"].includes(p) &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Value" }),
                        i.jsx("input", {
                          "aria-label": d("Conditional value"),
                          value: v,
                          onChange: (Re) => g(Re.target.value),
                          lang: u("Conditional value"),
                          "data-ui-label": "Conditional value",
                        }),
                      ],
                    }),
                  p === "between" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Maximum" }),
                        i.jsx("input", {
                          "aria-label": d("Conditional maximum"),
                          value: b,
                          onChange: (Re) => y(Re.target.value),
                          lang: u("Conditional maximum"),
                          "data-ui-label": "Conditional maximum",
                        }),
                      ],
                    }),
                  i.jsxs("div", {
                    className: "wt-columns",
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Fill" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": d("Conditional fill"),
                            value: w,
                            onChange: (Re) => E(Re.target.value),
                            lang: u("Conditional fill"),
                            "data-ui-label": "Conditional fill",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Text" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": d("Conditional color"),
                            value: A,
                            onChange: (Re) => k(Re.target.value),
                            lang: u("Conditional color"),
                            "data-ui-label": "Conditional color",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(M, {
                      children:
                        "Rules update when calculated cell values change.",
                    }),
                  }),
                  (ye = l.conditionalFormats) == null
                    ? void 0
                    : ye.map((Re) =>
                        i.jsxs(
                          "div",
                          {
                            className: "wt-existing",
                            children: [
                              i.jsxs("span", {
                                children: [Re.range, " · ", Re.kind],
                              }),
                              i.jsx("button", {
                                disabled: o,
                                onClick: () =>
                                  ce(
                                    {
                                      ...r,
                                      sheets: r.sheets.map((Me) => {
                                        var ke;
                                        return Me.id === t
                                          ? {
                                              ...Me,
                                              conditionalFormats:
                                                (ke = Me.conditionalFormats) ==
                                                null
                                                  ? void 0
                                                  : ke.filter(
                                                      (je) => je.id !== Re.id,
                                                    ),
                                            }
                                          : Me;
                                      }),
                                    },
                                    "Remove conditional formatting",
                                  ),
                                children: i.jsx(M, { children: "Remove rule" }),
                              }),
                            ],
                          },
                          Re.id,
                        ),
                      ),
                ],
              }),
            e === "validation" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Allow" }),
                      i.jsxs("select", {
                        "aria-label": d("Validation type"),
                        value: p,
                        onChange: (Re) => m(Re.target.value),
                        lang: u("Validation type"),
                        "data-ui-label": "Validation type",
                        children: [
                          i.jsx("option", { value: "list", children: "List" }),
                          i.jsx("option", {
                            value: "number",
                            children: "Decimal",
                          }),
                          i.jsx("option", {
                            value: "integer",
                            children: "Whole number",
                          }),
                          i.jsx("option", {
                            value: "date",
                            children: "Date (Excel serial)",
                          }),
                          i.jsx("option", {
                            value: "textLength",
                            children: "Text length",
                          }),
                        ],
                      }),
                    ],
                  }),
                  p !== "list" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Condition" }),
                        i.jsxs("select", {
                          "aria-label": d("Validation operator"),
                          value: F,
                          onChange: (Re) => I(Re.target.value),
                          lang: u("Validation operator"),
                          "data-ui-label": "Validation operator",
                          children: [
                            i.jsx("option", {
                              value: "between",
                              children: "Between",
                            }),
                            i.jsx("option", {
                              value: "greater",
                              children: "Greater than",
                            }),
                            i.jsx("option", {
                              value: "less",
                              children: "Less than",
                            }),
                            i.jsx("option", {
                              value: "equal",
                              children: "Equal to",
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsxs("label", {
                    children: [
                      p === "list"
                        ? "Values, separated by commas"
                        : "Minimum / comparison value",
                      i.jsx("input", {
                        "aria-label": d("Validation values"),
                        value: v,
                        onChange: (Re) => g(Re.target.value),
                        placeholder:
                          p === "list" ? "Pending, Complete, Reviewed" : "0",
                        lang: u("Validation values"),
                        "data-ui-label": "Validation values",
                      }),
                    ],
                  }),
                  p !== "list" &&
                    F === "between" &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Maximum" }),
                        i.jsx("input", {
                          "aria-label": d("Validation maximum"),
                          value: b,
                          onChange: (Re) => y(Re.target.value),
                          lang: u("Validation maximum"),
                          "data-ui-label": "Validation maximum",
                        }),
                      ],
                    }),
                  i.jsxs("label", {
                    className: "wt-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: O,
                        onChange: (Re) => j(Re.target.checked),
                      }),
                      i.jsx(M, { children: "Allow blank cells" }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Error message" }),
                      i.jsx("input", {
                        "aria-label": d("Validation message"),
                        value: z,
                        onChange: (Re) => V(Re.target.value),
                        lang: u("Validation message"),
                        "data-ui-label": "Validation message",
                      }),
                    ],
                  }),
                  (Te = l.validations) == null
                    ? void 0
                    : Te.map((Re) =>
                        i.jsxs(
                          "div",
                          {
                            className: "wt-existing",
                            children: [
                              i.jsxs("span", {
                                children: [Re.range, " · ", Re.kind],
                              }),
                              i.jsx("button", {
                                disabled: o,
                                onClick: () =>
                                  ce(
                                    {
                                      ...r,
                                      sheets: r.sheets.map((Me) => {
                                        var ke;
                                        return Me.id === t
                                          ? {
                                              ...Me,
                                              validations:
                                                (ke = Me.validations) == null
                                                  ? void 0
                                                  : ke.filter(
                                                      (je) => je.id !== Re.id,
                                                    ),
                                            }
                                          : Me;
                                      }),
                                    },
                                    "Remove data validation",
                                  ),
                                children: i.jsx(M, { children: "Remove rule" }),
                              }),
                            ],
                          },
                          Re.id,
                        ),
                      ),
                ],
              }),
            e === "duplicates" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    className: "wt-check",
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: D,
                        onChange: (Re) => B(Re.target.checked),
                      }),
                      i.jsx(M, { children: "My data has headers" }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(M, {
                          children: "Compare these columns",
                        }),
                      }),
                      we.map((Re) => {
                        var Me;
                        return i.jsxs(
                          "label",
                          {
                            className: "wt-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: (_ ?? we).includes(Re),
                                onChange: (ke) =>
                                  C(
                                    ke.target.checked
                                      ? [...(_ ?? we), Re]
                                      : (_ ?? we).filter((je) => je !== Re),
                                  ),
                              }),
                              ut(Re),
                              " ·",
                              " ",
                              String(
                                ((Me =
                                  l.cells[`${ut(Re)}${se.firstRow + 1}`]) ==
                                null
                                  ? void 0
                                  : Me.raw) ?? "",
                              ),
                            ],
                          },
                          Re,
                        );
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(M, {
                      children:
                        "The first occurrence is retained. Later duplicates are removed from the selected rectangle, remaining rows move up within it, and the vacated cells are cleared. Undo restores the full selection.",
                    }),
                  }),
                ],
              }),
            e === "protection" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("p", {
                    children: [
                      i.jsx(ShieldCheck, { size: 17 }),
                      " ",
                      (Pe = l.protection) != null && Pe.enabled
                        ? "Unlock editing on this worksheet."
                        : "Prevent changes to this worksheet until it is unprotected.",
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      (qe = l.protection) != null && qe.hash
                        ? "Password"
                        : "Password (optional)",
                      i.jsx("input", {
                        type: "password",
                        "aria-label": d("Sheet password"),
                        value: T,
                        onChange: (Re) => N(Re.target.value),
                        autoComplete: "new-password",
                        lang: u("Sheet password"),
                        "data-ui-label": "Sheet password",
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(M, {
                      children:
                        "This prevents edits in the workbook editor. The embedding application controls access to files and data.",
                    }),
                  }),
                ],
              }),
            (e === "insertRow" || e === "insertColumn") &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Operation" }),
                      i.jsxs("select", {
                        "aria-label": d("Structure operation"),
                        value: $ ? "delete" : "insert",
                        onChange: (Re) => U(Re.target.value === "delete"),
                        lang: u("Structure operation"),
                        "data-ui-label": "Structure operation",
                        children: [
                          i.jsx("option", {
                            value: "insert",
                            children: "Insert before selection",
                          }),
                          i.jsx("option", {
                            value: "delete",
                            children: "Delete starting at selection",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Count" }),
                      i.jsx(ci, {
                        min: "1",
                        max: "10000",
                        "aria-label": d("Structure count"),
                        value: R,
                        onValueChange: (Re) => L(Number(Re)),
                        lang: u("Structure count"),
                        "data-ui-label": "Structure count",
                      }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(M, {
                      children:
                        "Values, supported references, row/column sizes and supported range metadata move with the edit. Undo restores the previous workbook.",
                    }),
                  }),
                ],
              }),
            e === "spelling" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsxs("p", {
                    children: [
                      i.jsx(CheckCheck, { size: 17 }),
                      " ",
                      i.jsx(M, {
                        children:
                          "Review text with your browser’s spelling dictionary. Right-click an underlined word to see available corrections.",
                      }),
                    ],
                  }),
                  ae.length
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Cell" }),
                              i.jsx("select", {
                                "aria-label": d("Spelling cell"),
                                value: oe,
                                onChange: (Re) => {
                                  const Me = Number(Re.target.value);
                                  (de(Me), Se(String(ae[Me][1].raw)));
                                },
                                lang: u("Spelling cell"),
                                "data-ui-label": "Spelling cell",
                                children: ae.map(([Re], Me) =>
                                  i.jsx(
                                    "option",
                                    { value: Me, children: Re },
                                    Re,
                                  ),
                                ),
                              }),
                            ],
                          }),
                          i.jsx("textarea", {
                            "aria-label": d("Spelling text"),
                            spellCheck: true,
                            autoCorrect: "on",
                            value: he,
                            onChange: (Re) => Se(Re.target.value),
                            readOnly: o,
                            lang: u("Spelling text"),
                            "data-ui-label": "Spelling text",
                          }),
                        ],
                      })
                    : i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "No text cells were found in this selection.",
                        }),
                      }),
                  i.jsx("p", {
                    className: "wt-hint",
                    children: i.jsx(M, {
                      children:
                        "Spelling suggestions depend on the browser and its installed languages.",
                    }),
                  }),
                ],
              }),
            Q && i.jsx("p", { role: "status", children: Q }),
            G &&
              i.jsx("p", { className: "wt-error", role: "alert", children: G }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: a,
              children: i.jsx(M, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "wt-primary",
              disabled: J || o,
              onClick: () => void ge(),
              children: J
                ? "Applying…"
                : e === "spelling"
                  ? "Save correction"
                  : e === "protection"
                    ? (ht = l.protection) != null && ht.enabled
                      ? "Unprotect"
                      : "Protect"
                    : "Apply",
            }),
          ],
        }),
      ],
    }),
  });
}
function ufe({
  workbook: e,
  sheetId: r,
  tableId: t,
  context: n,
  readOnly: o,
  onWorkbookChange: a,
}) {
  var C;
  const { text: s, language: d } = Dr(),
    u = e.sheets.find((T) => T.id === r).tables.find((T) => T.id === t),
    [l, c] = ne.useState(u.columns[0].id),
    [h, f] = ne.useState(u.columns[0].dataType ?? "any"),
    [p, m] = ne.useState("convert"),
    [v, g] = ne.useState("reject"),
    [b, y] = ne.useState("."),
    [w, E] = ne.useState(","),
    [A, k] = ne.useState(null),
    [F, I] = ne.useState(""),
    [O, j] = ne.useState(0),
    z = { columnId: l, dataType: h, mode: p, invalid: v, decimal: b, group: w },
    V = JSON.stringify({ context: n, request: z }),
    D = !!A && (A.context !== V || A.plan.snapshot !== JSON.stringify(e)),
    B = ne.useMemo(() => {
      try {
        return Gne(e, r, t, l);
      } catch {
        return null;
      }
    }, [e, r, t, l]),
    _ =
      ((C = e.tableEvidence) == null
        ? void 0
        : C.filter((T) => T.tableId === t)) ?? [];
  return i.jsxs("details", {
    children: [
      i.jsx("summary", {
        children: i.jsx(M, {
          children: "Column types and original source evidence",
        }),
      }),
      i.jsxs("fieldset", {
        className: "lg-table-list",
        children: [
          i.jsx("legend", {
            children: i.jsx(M, { children: "Review column types" }),
          }),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Typed column" }),
              i.jsx("select", {
                "aria-label": s("Typed column"),
                value: l,
                onChange: (T) => {
                  var N;
                  (c(T.target.value),
                    f(
                      ((N = u.columns.find((R) => R.id === T.target.value)) ==
                      null
                        ? void 0
                        : N.dataType) ?? "any",
                    ),
                    k(null));
                },
                lang: d("Typed column"),
                "data-ui-label": "Typed column",
                children: u.columns.map((T) =>
                  i.jsx("option", { value: T.id, children: T.name }, T.id),
                ),
              }),
            ],
          }),
          i.jsxs("p", {
            "aria-label": s("Current column type"),
            lang: d("Current column type"),
            "data-ui-label": "Current column type",
            children: [
              i.jsx(M, { children: "Saved type: " }),
              (B == null ? void 0 : B.dataType) ?? "unknown",
              "; ",
              (B == null ? void 0 : B.checked) ?? 0,
              " ",
              i.jsx(M, { children: "rows checked; " }),
              (B == null ? void 0 : B.exceptions.length) ?? 0,
              " ",
              i.jsx(M, { children: "exceptions" }),
              B != null && B.truncated
                ? " — partial inspection (first 500 rows)"
                : "",
              ".",
            ],
          }),
          !!(B != null && B.exceptions.length) &&
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(M, { children: "Current type exceptions" }),
                }),
                B.exceptions
                  .slice(0, 50)
                  .map((T) =>
                    i.jsxs(
                      "p",
                      { children: [T.address, ": ", JSON.stringify(T.value)] },
                      T.address,
                    ),
                  ),
              ],
            }),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Column type" }),
              i.jsx("select", {
                "aria-label": s("Column type"),
                value: h,
                onChange: (T) => f(T.target.value),
                lang: d("Column type"),
                "data-ui-label": "Column type",
                children: TP.map((T) =>
                  i.jsx(
                    "option",
                    {
                      value: T,
                      children: T === "date" ? "Date (YYYY-MM-DD text)" : T,
                    },
                    T,
                  ),
                ),
              }),
            ],
          }),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Type operation" }),
              i.jsxs("select", {
                value: p,
                onChange: (T) => m(T.target.value),
                children: [
                  i.jsx("option", {
                    value: "convert",
                    children: "Convert compatible values",
                  }),
                  i.jsx("option", {
                    value: "tag",
                    children: "Set type; keep every source value",
                  }),
                ],
              }),
            ],
          }),
          h === "number" &&
            p === "convert" &&
            i.jsxs("div", {
              className: "lg-table-inline",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Decimal separator" }),
                    i.jsxs("select", {
                      value: b,
                      onChange: (T) => y(T.target.value),
                      children: [
                        i.jsx("option", { value: ".", children: "Period" }),
                        i.jsx("option", { value: ",", children: "Comma" }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Group separator" }),
                    i.jsxs("select", {
                      value: w,
                      onChange: (T) => E(T.target.value),
                      children: [
                        i.jsx("option", { value: ",", children: "Comma" }),
                        i.jsx("option", { value: ".", children: "Period" }),
                        i.jsx("option", { value: " ", children: "Space" }),
                        i.jsx("option", { value: "", children: "None" }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Unconvertible values" }),
              i.jsxs("select", {
                value: v,
                onChange: (T) => g(T.target.value),
                children: [
                  i.jsx("option", {
                    value: "reject",
                    children: "Block apply until resolved",
                  }),
                  i.jsx("option", {
                    value: "keep",
                    children: "Keep exceptions unchanged",
                  }),
                ],
              }),
            ],
          }),
          i.jsx("p", {
            className: "lg-table-help",
            children: i.jsx(M, {
              children:
                "Preview up to 500 rows. Blanks and formulas are preserved. Types report current exceptions; ordinary edits are not coerced. Explicit row appends must match the saved types. Native JSON retains original headers and converted cells, including their notes and formats.",
            }),
          }),
          F && i.jsx("p", { role: "alert", children: F }),
          i.jsx("button", {
            type: "button",
            disabled: o,
            onClick: () => {
              (I(""), j(0), k(null));
              try {
                k({ plan: RP(e, r, t, z, { readOnly: o }), context: V });
              } catch (T) {
                I(T.message);
              }
            },
            children: i.jsx(M, { children: "Preview column type" }),
          }),
          A &&
            i.jsxs("section", {
              "aria-label": s("Column type preview"),
              className: "lg-table-preview",
              lang: d("Column type preview"),
              "data-ui-label": "Column type preview",
              children: [
                i.jsx("h3", {
                  children: i.jsx(M, { children: "Review type changes" }),
                }),
                i.jsxs("p", {
                  children: [
                    A.plan.beforeType,
                    " → ",
                    A.plan.request.dataType,
                    ";",
                    " ",
                    A.plan.changes.length,
                    " ",
                    i.jsx(M, { children: "cell changes;" }),
                    " ",
                    A.plan.exceptions.length,
                    " ",
                    i.jsx(M, { children: "exceptions." }),
                  ],
                }),
                D &&
                  i.jsx("p", {
                    role: "alert",
                    children: i.jsx(M, {
                      children:
                        "The workbook, selection or settings changed. Preview again.",
                    }),
                  }),
                !A.plan.canApply &&
                  i.jsx("p", {
                    role: "alert",
                    children: i.jsx(M, {
                      children:
                        "Resolve the exceptions or explicitly choose Keep exceptions unchanged, then preview again.",
                    }),
                  }),
                i.jsx("ol", {
                  start: O * 20 + 1,
                  children: [
                    ...A.plan.changes.map((T) => {
                      var N;
                      return {
                        address: T.address,
                        text: `${JSON.stringify(((N = T.before) == null ? void 0 : N.raw) ?? null)} → ${JSON.stringify(T.after.raw)}`,
                      };
                    }),
                    ...A.plan.exceptions.map((T) => ({
                      address: T.address,
                      text: `Retained ${JSON.stringify(T.value)}: ${T.reason}`,
                    })),
                  ]
                    .slice(O * 20, O * 20 + 20)
                    .map((T) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            i.jsx("strong", { children: T.address }),
                            ": ",
                            T.text,
                          ],
                        },
                        T.address,
                      ),
                    ),
                }),
                i.jsxs("div", {
                  className: "lg-table-inline",
                  children: [
                    i.jsx("button", {
                      type: "button",
                      disabled: !O,
                      onClick: () => j((T) => T - 1),
                      children: i.jsx(M, { children: "Previous type changes" }),
                    }),
                    i.jsxs("span", {
                      children: [i.jsx(M, { children: "Page " }), O + 1],
                    }),
                    i.jsx("button", {
                      type: "button",
                      disabled:
                        (O + 1) * 20 >=
                        A.plan.changes.length + A.plan.exceptions.length,
                      onClick: () => j((T) => T + 1),
                      children: i.jsx(M, { children: "Next type changes" }),
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: o || D || !A.plan.canApply,
                  onClick: () => {
                    I("");
                    try {
                      if (D) throw Error("Preview again.");
                      const T = Zne(e, A.plan, { readOnly: o });
                      (T !== e && a(T, "Apply reviewed table column type"),
                        k(null));
                    } catch (T) {
                      I(T.message);
                    }
                  },
                  children: i.jsx(M, { children: "Apply column type" }),
                }),
                i.jsx("button", {
                  type: "button",
                  onClick: () => k(null),
                  children: i.jsx(M, { children: "Cancel type preview" }),
                }),
              ],
            }),
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(M, { children: "Saved original source evidence (" }),
                  _.length,
                  ")",
                ],
              }),
              i.jsx("p", {
                children: i.jsx(M, {
                  children:
                    "Historical addresses and source revisions. Save native JSON to retain these records; XLSX carries current values only.",
                }),
              }),
              _.map((T) =>
                i.jsxs(
                  "details",
                  {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          T.kind,
                          " ",
                          i.jsx(M, { children: "· source revision " }),
                          T.sourceRevision,
                          " · ",
                          T.changes.length,
                          " ",
                          i.jsx(M, { children: "cells" }),
                        ],
                      }),
                      i.jsx("pre", { children: JSON.stringify(T, null, 2) }),
                    ],
                  },
                  T.id,
                ),
              ),
            ],
          }),
        ],
      }),
    ],
  });
}
function cfe({
  fields: e,
  groups: r,
  dateGroups: t,
  setDateGroups: n,
  calculatedFields: o,
  setCalculatedFields: a,
  onAdd: s,
}) {
  const { text: d } = Dr();
  return i.jsxs(i.Fragment, {
    children: [
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(M, { children: "Date grouping" }),
          }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Choose year, quarter or month for selected row/column fields. ISO requires YYYY-MM-DD text. Numeric dates require an explicit 1900 or 1904 date system.",
            }),
          }),
          [...new Set(r)].map((u) => {
            const l = t.find((c) => c.field === u);
            return i.jsxs(
              "div",
              {
                className: "lg-table-inline",
                children: [
                  i.jsxs("label", {
                    children: [
                      u,
                      " ",
                      i.jsx(M, { children: "grouping" }),
                      i.jsxs("select", {
                        "aria-label": `Date grouping ${u}`,
                        value: (l == null ? void 0 : l.unit) ?? "",
                        onChange: (c) =>
                          n([
                            ...t.filter((h) => h.field !== u),
                            ...(c.target.value
                              ? [
                                  {
                                    field: u,
                                    unit: c.target.value,
                                    dateSystem:
                                      (l == null ? void 0 : l.dateSystem) ??
                                      "iso",
                                  },
                                ]
                              : []),
                          ]),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "No date grouping",
                          }),
                          i.jsx("option", { value: "year", children: "Year" }),
                          i.jsx("option", {
                            value: "quarter",
                            children: "Quarter",
                          }),
                          i.jsx("option", {
                            value: "month",
                            children: "Month",
                          }),
                        ],
                      }),
                    ],
                  }),
                  l &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Date system for " }),
                        u,
                        i.jsxs("select", {
                          value: l.dateSystem,
                          onChange: (c) =>
                            n(
                              t.map((h) =>
                                h.field === u
                                  ? { ...h, dateSystem: c.target.value }
                                  : h,
                              ),
                            ),
                          children: [
                            i.jsx("option", {
                              value: "iso",
                              children: "ISO date text",
                            }),
                            i.jsx("option", {
                              value: "1900",
                              children: "Excel 1900 serial",
                            }),
                            i.jsx("option", {
                              value: "1904",
                              children: "Excel 1904 serial",
                            }),
                          ],
                        }),
                      ],
                    }),
                ],
              },
              u,
            );
          }),
        ],
      }),
      i.jsxs("details", {
        children: [
          i.jsx("summary", {
            children: i.jsx(M, { children: "Calculated pivot fields" }),
          }),
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Each calculation uses source-field sums within each group and total. These are aggregate calculations. Use a table calculated column for arithmetic on individual records. Choose sum in Value fields.",
            }),
          }),
          o.map((u, l) => {
            const c = (h) => a(o.map((f, p) => (l === p ? { ...f, ...h } : f)));
            return i.jsxs(
              "fieldset",
              {
                className: "lg-table-list",
                children: [
                  i.jsxs("legend", {
                    children: [i.jsx(M, { children: "Calculation " }), l + 1],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Calculated field name " }),
                      l + 1,
                      i.jsx("input", {
                        value: u.name,
                        onChange: (h) => c({ name: h.target.value }),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Left source field " }),
                      l + 1,
                      i.jsx("select", {
                        value: u.field,
                        onChange: (h) => c({ field: h.target.value }),
                        children: e.map((h) =>
                          i.jsx("option", { children: h }, h),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Calculation operation " }),
                      l + 1,
                      i.jsx("select", {
                        value: u.operation,
                        onChange: (h) => c({ operation: h.target.value }),
                        children: ["add", "subtract", "multiply", "divide"].map(
                          (h) => i.jsx("option", { children: h }, h),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Right operand " }),
                      l + 1,
                      i.jsxs("select", {
                        value: u.otherField ?? "",
                        onChange: (h) =>
                          a(
                            o.map((f, p) =>
                              p !== l
                                ? f
                                : {
                                    name: f.name,
                                    field: f.field,
                                    operation: f.operation,
                                    ...(h.target.value
                                      ? { otherField: h.target.value }
                                      : { constant: 1 }),
                                  },
                            ),
                          ),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Numeric constant",
                          }),
                          e.map((h) => i.jsx("option", { children: h }, h)),
                        ],
                      }),
                    ],
                  }),
                  !u.otherField &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Calculation constant " }),
                        l + 1,
                        i.jsx(ci, {
                          "aria-label": d("Calculation constant ") + (l + 1),
                          value: Number.isFinite(u.constant) ? u.constant : "",
                          onValueChange: (h) =>
                            c({ constant: h === "" ? NaN : Number(h) }),
                        }),
                      ],
                    }),
                  i.jsxs("button", {
                    type: "button",
                    onClick: () => a(o.filter((h, f) => f !== l)),
                    children: [
                      i.jsx(M, { children: "Remove calculation " }),
                      l + 1,
                    ],
                  }),
                ],
              },
              l,
            );
          }),
          i.jsx("button", {
            type: "button",
            disabled: !e.length || o.length >= 100,
            onClick: () => {
              let u = 1;
              for (
                ;
                e.includes(`Calculated${u}`) ||
                o.some((c) => c.name === `Calculated${u}`);
              )
                u++;
              const l = `Calculated${u}`;
              (a([
                ...o,
                {
                  name: l,
                  field: e.at(-1),
                  operation: "multiply",
                  constant: 1,
                },
              ]),
                s(l));
            },
            children: i.jsx(M, { children: "Add calculated pivot field" }),
          }),
        ],
      }),
    ],
  });
}
function BI({ workbook: e, pivotId: r, compact: t = false }) {
  const n = ne.useMemo(() => HP(e, r), [e, r]),
    o = n.definition,
    a = o.refresh;
  return i.jsxs("section", {
    "aria-label": `Pivot status ${o.name}`,
    className: `lg-pivot-status ${n.state}`,
    children: [
      i.jsxs("strong", {
        children: [
          o.name,
          ":",
          " ",
          n.state === "current"
            ? "Current"
            : n.state === "stale"
              ? "Stale — refresh required"
              : n.state === "unverified"
                ? "Unverified — refresh required"
                : "Source unavailable",
        ],
      }),
      i.jsxs("p", {
        children: [
          (a == null ? void 0 : a.sourceSheetName) ?? o.source.sheetId,
          "!",
          (a == null ? void 0 : a.sourceRange) ?? o.source.range,
          " ",
          i.jsx(M, { children: "· Source revision " }),
          o.sourceRevision,
          " ",
          i.jsx(M, { children: "· Refreshed " }),
          (a == null ? void 0 : a.at) ?? "not recorded",
        ],
      }),
      n.state !== "current" &&
        i.jsxs("p", {
          children: [
            n.reason,
            " ",
            i.jsx(M, {
              children: "Existing totals are retained as historical output.",
            }),
          ],
        }),
      !t &&
        a &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsxs("p", {
              children: [
                a.includedRows,
                " ",
                i.jsx(M, { children: "included + " }),
                a.excludedRows,
                " ",
                i.jsx(M, { children: "excluded =" }),
                " ",
                a.sourceRows,
                " ",
                i.jsx(M, {
                  children:
                    "source records. Blank rows are omitted; hidden rows remain in the population unless explicitly filtered.",
                }),
              ],
            }),
            i.jsx("ul", {
              children: a.measures.map((s, d) =>
                i.jsxs(
                  "li",
                  {
                    children: [
                      s.name,
                      ": ",
                      JSON.stringify(s.result),
                      " · ",
                      s.numericCount,
                      " ",
                      i.jsx(M, { children: "numeric / " }),
                      s.nonblankCount,
                      " ",
                      i.jsx(M, { children: "nonblank inputs" }),
                    ],
                  },
                  d,
                ),
              ),
            }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Distinct counts exclude blanks, compare text without case and retain value types. Averages and calculated totals are recomputed from the included source records.",
              }),
            }),
          ],
        }),
    ],
  });
}
function dfe({ workbook: e, sheetId: r, onOpen: t, onExplore: n }) {
  var d;
  const { text: o, language: a } = Dr(),
    s =
      ((d = e.sheets.find((u) => u.id === r)) == null ? void 0 : d.pivots) ??
      [];
  return s.length
    ? i.jsxs("div", {
        className: "lg-pivot-status-bar",
        "aria-label": o("Pivot result freshness"),
        lang: a("Pivot result freshness"),
        "data-ui-label": "Pivot result freshness",
        children: [
          s
            .slice(0, 3)
            .map((u) =>
              i.jsx(BI, { workbook: e, pivotId: u.id, compact: true }, u.id),
            ),
          s.length > 3 &&
            i.jsxs("p", {
              children: [
                s.length - 3,
                " ",
                i.jsx(M, {
                  children:
                    "additional results — review each in PivotTable tools.",
                }),
              ],
            }),
          i.jsx("button", {
            type: "button",
            onClick: t,
            children: i.jsx(M, { children: "Review / refresh PivotTable" }),
          }),
          n &&
            i.jsx("button", {
              type: "button",
              onClick: n,
              children: i.jsx(M, { children: "Filter / explore PivotTable" }),
            }),
        ],
      })
    : null;
}
const ffe = (e, r) => (e.includes(r) ? e.filter((t) => t !== r) : [...e, r]),
  Xk = (e) =>
    e === "distinctCount" ? "Distinct count" : e[0].toUpperCase() + e.slice(1);
function hfe(e, r, t) {
  try {
    return pR(e, r, t);
  } catch {
    return [];
  }
}
function Y7(e) {
  let r = 1;
  for (
    ;
    e.sheets.some((t) => {
      var n;
      return (n = t.tables) == null
        ? void 0
        : n.some((o) => o.name.toLowerCase() === `table${r}`);
    });
  )
    r++;
  return `Table${r}`;
}

export { BI, Xk, Y7, cfe, dfe, ffe, hfe, sfe, ufe };
