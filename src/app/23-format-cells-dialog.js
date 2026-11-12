import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { dx } from "./02-financial-functions.js";
import {
  $Q,
  Af,
  BB,
  Dw,
  FN,
  J3,
  KJ,
  LB,
  Nr,
  RN,
  Vn,
  W3,
  ct,
  ff,
  jn,
  kE,
  oh,
  pB,
  rn,
  s0,
  uB,
  ut,
  y1,
  ys,
} from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import { $k, Tde, jI, oy } from "./22-themes-and-cell-styles.js";

function Nde(e) {
  var $, U, G;
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const s = e.workbook.sheets.find((H) => H.id === e.sheetId),
    d = ff(e.ranges ?? [e.range]),
    u = ct(e.range),
    l =
      e.active && W3(d, e.active)
        ? e.active
        : `${ut(u.firstCol)}${u.firstRow + 1}`,
    c = ne.useRef({
      book: e.workbook.id,
      revision: e.workbook.revision,
      content: JSON.stringify(e.workbook),
      sheet: e.sheetId,
      selection: JSON.stringify(e.ranges ?? [e.range]),
      active: e.active,
    }),
    [h] = ne.useState(() => {
      var H;
      return oh(
        ys((H = s == null ? void 0 : s.cells[l]) == null ? void 0 : H.style),
      );
    }),
    [f, p] = ne.useState(
      e.mode === "alignment"
        ? "Alignment"
        : e.mode === "borders"
          ? "Border"
          : "Font",
    ),
    [m, v] = ne.useState({}),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(),
    [E, A] = ne.useState({ style: "thin", color: "#000000" }),
    [k, F] = ne.useState(h.fontSize === void 0 ? "" : String(h.fontSize)),
    [I, O] = ne.useState(String(h.indent ?? 0)),
    j = y1(d),
    z = ne.useMemo(() => Vn(e.workbook), [e.workbook]),
    V = s ? z(s.id, l) : null,
    D =
      c.current.book !== e.workbook.id ||
      c.current.revision !== e.workbook.revision ||
      c.current.content !== JSON.stringify(e.workbook) ||
      c.current.sheet !== e.sheetId ||
      c.current.selection !== JSON.stringify(e.ranges ?? [e.range]) ||
      c.current.active !== e.active,
    B =
      e.readOnly ||
      D ||
      !!(($ = s == null ? void 0 : s.protection) != null && $.enabled) ||
      !s ||
      j > 500;
  let _ = h,
    C = "";
  try {
    _ = oh({
      ...h,
      ...m,
      ...(m.borders ? { borders: { ...h.borders, ...m.borders } } : {}),
    });
  } catch (H) {
    C = H instanceof Error ? H.message : "Choose supported cell styles.";
  }
  function T(H) {
    (v((Q) => ({ ...Q, ...H })), b(""));
  }
  function N(H, Q = false) {
    (v((X) => ({
      ...X,
      borders: { ...X.borders, [H]: Q ? { style: "none" } : { ...E } },
    })),
      b(""));
  }
  function R(H) {
    (w(H),
      v((Q) => {
        const X = { ...Q };
        return (delete X.borders, X);
      }),
      b(""));
  }
  const L = (H) => ({
    ..._,
    borders: { ...h.borders, ...(y ? pB(H, Tde, y, E) : {}), ...m.borders },
  });
  return (
    ne.useEffect(() => {
      var Q, X;
      const H = document.activeElement;
      return (
        (X =
          (Q = o.current) == null
            ? void 0
            : Q.querySelector('[role="tab"][aria-selected="true"]')) == null ||
          X.focus(),
        () => {
          H != null && H.isConnected && H.focus();
        }
      );
    }, []),
    i.jsx("div", {
      className: "cs-backdrop",
      children: i.jsxs("div", {
        ref: o,
        className: "cs-dialog",
        role: "dialog",
        "aria-modal": "true",
        "aria-labelledby": `${n}-title`,
        onKeyDown: (H) => {
          var Q;
          if (
            (H.stopPropagation(),
            H.key === "Escape" && (H.preventDefault(), a.current()),
            H.key === "Tab")
          ) {
            const X = [
                ...(((Q = o.current) == null
                  ? void 0
                  : Q.querySelectorAll(
                      'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
                    )) ?? []),
              ],
              J = X[0],
              q = X.at(-1);
            H.shiftKey && document.activeElement === J
              ? (H.preventDefault(), q == null || q.focus())
              : !H.shiftKey &&
                document.activeElement === q &&
                (H.preventDefault(), J == null || J.focus());
          }
        },
        children: [
          i.jsxs("header", {
            className: "cs-header",
            children: [
              i.jsx("h2", {
                id: `${n}-title`,
                children: i.jsx(M, { children: "Format Cells" }),
              }),
              i.jsx("button", {
                "aria-label": r("Close cell styles"),
                onClick: e.onClose,
                lang: t("Close cell styles"),
                "data-ui-label": "Close cell styles",
                children: i.jsx(X, { size: 18 }),
              }),
            ],
          }),
          i.jsx("div", {
            className: "cs-tabs",
            role: "tablist",
            "aria-label": r("Cell style settings"),
            lang: t("Cell style settings"),
            "data-ui-label": "Cell style settings",
            children: oy.map((H, Q) =>
              i.jsx(
                "button",
                {
                  role: "tab",
                  id: `${n}-${H}`,
                  "aria-controls": `${n}-panel`,
                  "aria-selected": f === H,
                  tabIndex: f === H ? 0 : -1,
                  onClick: () => p(H),
                  onKeyDown: (X) => {
                    var J, q;
                    if (
                      ["ArrowLeft", "ArrowRight", "Home", "End"].includes(X.key)
                    ) {
                      X.preventDefault();
                      const W =
                        X.key === "Home"
                          ? 0
                          : X.key === "End"
                            ? oy.length - 1
                            : (Q +
                                (X.key === "ArrowRight" ? 1 : oy.length - 1)) %
                              oy.length;
                      (p(oy[W]),
                        (q =
                          (J = X.currentTarget.parentElement) == null
                            ? void 0
                            : J.children[W]) == null || q.focus());
                    }
                  },
                  children: H,
                },
                H,
              ),
            ),
          }),
          i.jsxs("form", {
            onSubmit: (H) => {
              (H.preventDefault(), b(""));
              try {
                if (e.readOnly) throw Error("This workbook is read-only.");
                if (D)
                  throw Error(
                    "The workbook or selection changed. Close and reopen Format Cells.",
                  );
                if (C) throw Error(C);
                const Q = KJ(e.workbook, e.sheetId, d, m, {
                  readOnly: e.readOnly,
                  ...(y ? { borders: { preset: y, line: E } } : {}),
                });
                (Q !== e.workbook &&
                  e.onWorkbookChange(Q, "Format cell styles"),
                  e.onClose());
              } catch (Q) {
                b(
                  Q instanceof Error
                    ? Q.message
                    : "The cell styles could not be applied.",
                );
              }
            },
            children: [
              i.jsxs("div", {
                className: "cs-body",
                children: [
                  i.jsxs("p", {
                    className: "cs-selection",
                    children: [
                      j.toLocaleString(),
                      " ",
                      i.jsx(M, { children: "selected cells · " }),
                      d.length,
                      " ",
                      d.length === 1 ? "area" : "areas",
                      i.jsx(M, {
                        children:
                          ". Only changed controls apply; existing differences are preserved.",
                      }),
                    ],
                  }),
                  e.readOnly &&
                    i.jsx("p", {
                      className: "cs-notice",
                      children: i.jsx(M, {
                        children:
                          "This workbook is read-only. You can inspect style previews.",
                      }),
                    }),
                  ((U = s == null ? void 0 : s.protection) == null
                    ? void 0
                    : U.enabled) &&
                    i.jsx("p", {
                      className: "cs-notice",
                      children: i.jsx(M, {
                        children:
                          "This worksheet is protected. Unprotect it before changing styles.",
                      }),
                    }),
                  D &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: i.jsx(M, {
                        children:
                          "The workbook or selection changed. Close and reopen Format Cells.",
                      }),
                    }),
                  j > 500 &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: i.jsx(M, {
                        children:
                          "Select at most 500 unique cells to apply styles.",
                      }),
                    }),
                  i.jsxs("div", {
                    className: "cs-panel",
                    role: "tabpanel",
                    id: `${n}-panel`,
                    "aria-labelledby": `${n}-${f}`,
                    children: [
                      f === "Font" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Font" }),
                                    i.jsxs("select", {
                                      "aria-label": r("Font family"),
                                      value: _.fontFamily ?? "",
                                      onChange: (H) =>
                                        T({
                                          fontFamily: H.target.value || void 0,
                                        }),
                                      lang: t("Font family"),
                                      "data-ui-label": "Font family",
                                      children: [
                                        i.jsx("option", {
                                          value: "",
                                          children: "Default (Calibri)",
                                        }),
                                        dx.map((H) =>
                                          i.jsx("option", { children: H }, H),
                                        ),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Size" }),
                                    i.jsx(ci, {
                                      allowEmpty: true,
                                      "aria-label": r("Font size"),
                                      min: 8,
                                      max: 72,
                                      step: 1,
                                      placeholder: r("Default (14 px)"),
                                      value: k,
                                      onValueChange: (H) => {
                                        (F(H),
                                          T({
                                            fontSize:
                                              H === "" ? void 0 : Number(H),
                                          }));
                                      },
                                      lang: t("Font size"),
                                      "data-ui-label": "Font size",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              className: "cs-effects",
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(M, { children: "Effects" }),
                                }),
                                [
                                  ["bold", "Bold"],
                                  ["italic", "Italic"],
                                  ["underline", "Underline"],
                                  ["strikethrough", "Strikethrough"],
                                ].map(([H, Q]) =>
                                  i.jsxs(
                                    "label",
                                    {
                                      children: [
                                        i.jsx("input", {
                                          type: "checkbox",
                                          "aria-label": Q,
                                          checked: _[H] ?? false,
                                          onChange: (X) =>
                                            T({ [H]: X.target.checked }),
                                        }),
                                        Q,
                                      ],
                                    },
                                    H,
                                  ),
                                ),
                              ],
                            }),
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Font color" }),
                                    i.jsxs("span", {
                                      className: "cs-color",
                                      children: [
                                        i.jsx("input", {
                                          type: "color",
                                          "aria-label": r("Font color"),
                                          value: $k(_.color, "#000000"),
                                          onChange: (H) =>
                                            T({ color: H.target.value }),
                                          lang: t("Font color"),
                                          "data-ui-label": "Font color",
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          onClick: () => T({ color: void 0 }),
                                          children: i.jsx(M, {
                                            children: "Automatic",
                                          }),
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Fill color" }),
                                    i.jsxs("span", {
                                      className: "cs-color",
                                      children: [
                                        i.jsx("input", {
                                          type: "color",
                                          "aria-label": r("Fill color"),
                                          value: $k(_.fill, "#ffffff"),
                                          onChange: (H) =>
                                            T({ fill: H.target.value }),
                                          lang: t("Fill color"),
                                          "data-ui-label": "Fill color",
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          onClick: () => T({ fill: void 0 }),
                                          children: i.jsx(M, {
                                            children: "No fill",
                                          }),
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                              ],
                            }),
                          ],
                        }),
                      f === "Alignment" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Horizontal" }),
                                    i.jsxs("select", {
                                      "aria-label": r("Horizontal alignment"),
                                      value: _.align ?? "",
                                      onChange: (H) =>
                                        T({ align: H.target.value || void 0 }),
                                      lang: t("Horizontal alignment"),
                                      "data-ui-label": "Horizontal alignment",
                                      children: [
                                        i.jsx("option", {
                                          value: "",
                                          children: "General",
                                        }),
                                        i.jsx("option", {
                                          value: "left",
                                          children: "Left",
                                        }),
                                        i.jsx("option", {
                                          value: "center",
                                          children: "Center",
                                        }),
                                        i.jsx("option", {
                                          value: "right",
                                          children: "Right",
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Vertical" }),
                                    i.jsxs("select", {
                                      "aria-label": r("Vertical alignment"),
                                      value: _.verticalAlign ?? "",
                                      onChange: (H) =>
                                        T({
                                          verticalAlign:
                                            H.target.value || void 0,
                                        }),
                                      lang: t("Vertical alignment"),
                                      "data-ui-label": "Vertical alignment",
                                      children: [
                                        i.jsx("option", {
                                          value: "",
                                          children: "Default",
                                        }),
                                        i.jsx("option", {
                                          value: "top",
                                          children: "Top",
                                        }),
                                        i.jsx("option", {
                                          value: "middle",
                                          children: "Center",
                                        }),
                                        i.jsx("option", {
                                          value: "bottom",
                                          children: "Bottom",
                                        }),
                                      ],
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Indent" }),
                                    i.jsx(ci, {
                                      "aria-label": r("Indent level"),
                                      min: 0,
                                      max: 15,
                                      step: 1,
                                      value: I,
                                      onValueChange: (H) => {
                                        (O(H),
                                          T({
                                            indent: H === "" ? NaN : Number(H),
                                          }));
                                      },
                                      lang: t("Indent level"),
                                      "data-ui-label": "Indent level",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              className: "cs-check",
                              children: [
                                i.jsx("input", {
                                  type: "checkbox",
                                  "aria-label": r("Wrap text"),
                                  checked: _.wrap ?? false,
                                  onChange: (H) =>
                                    T({ wrap: H.target.checked }),
                                  lang: t("Wrap text"),
                                  "data-ui-label": "Wrap text",
                                }),
                                i.jsx(M, { children: "Wrap text" }),
                              ],
                            }),
                            i.jsx("p", {
                              className: "cs-help",
                              children: i.jsx(M, {
                                children:
                                  "Indent applies to the left or right edge according to alignment. Centered text ignores indent. Explicit row heights still clip content; use AutoFit when more space is needed.",
                              }),
                            }),
                          ],
                        }),
                      f === "Border" &&
                        i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("div", {
                              className: "cs-fields",
                              children: [
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Line" }),
                                    i.jsx("select", {
                                      "aria-label": r("Border line style"),
                                      value: E.style,
                                      onChange: (H) =>
                                        A({ ...E, style: H.target.value }),
                                      lang: t("Border line style"),
                                      "data-ui-label": "Border line style",
                                      children: uB.map((H) =>
                                        i.jsx(
                                          "option",
                                          {
                                            value: H,
                                            children:
                                              H.charAt(0).toUpperCase() +
                                              H.slice(1),
                                          },
                                          H,
                                        ),
                                      ),
                                    }),
                                  ],
                                }),
                                i.jsxs("label", {
                                  children: [
                                    i.jsx(M, { children: "Color" }),
                                    i.jsx("input", {
                                      type: "color",
                                      "aria-label": r("Border color"),
                                      value: $k(E.color, "#000000"),
                                      onChange: (H) =>
                                        A({ ...E, color: H.target.value }),
                                      lang: t("Border color"),
                                      "data-ui-label": "Border color",
                                    }),
                                  ],
                                }),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              className: "cs-presets",
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(M, {
                                    children: "Presets for selected cells",
                                  }),
                                }),
                                i.jsx("button", {
                                  type: "button",
                                  onClick: () => {
                                    (w(void 0),
                                      v((H) => {
                                        const Q = { ...H };
                                        return (delete Q.borders, Q);
                                      }),
                                      b(""));
                                  },
                                  children: i.jsx(M, {
                                    children: "Keep existing borders",
                                  }),
                                }),
                                [
                                  ["clear", "Clear borders"],
                                  ["outline", "Outline borders"],
                                  ["inside", "Inside borders"],
                                  ["all", "All borders"],
                                ].map(([H, Q]) =>
                                  i.jsx(
                                    "button",
                                    {
                                      type: "button",
                                      "aria-pressed": y === H,
                                      onClick: () => R(H),
                                      children: Q,
                                    },
                                    H,
                                  ),
                                ),
                              ],
                            }),
                            i.jsxs("fieldset", {
                              className: "cs-presets",
                              children: [
                                i.jsx("legend", {
                                  children: i.jsx(M, {
                                    children: "Individual edges",
                                  }),
                                }),
                                s0.map((H) =>
                                  i.jsxs(
                                    "span",
                                    {
                                      children: [
                                        i.jsx("button", {
                                          type: "button",
                                          "aria-label": `Set ${H} border`,
                                          onClick: () => N(H),
                                          children:
                                            H.charAt(0).toUpperCase() +
                                            H.slice(1),
                                        }),
                                        i.jsx("button", {
                                          type: "button",
                                          className: "cs-edge-clear",
                                          "aria-label": `Clear ${H} border`,
                                          onClick: () => N(H, true),
                                          children: "×",
                                        }),
                                      ],
                                    },
                                    H,
                                  ),
                                ),
                              ],
                            }),
                            i.jsx("p", {
                              className: "cs-help",
                              children: i.jsx(M, {
                                children:
                                  "Choose a line and color, then a preset or edge. Presets follow the selected cells, including gaps. Outline and Inside retain other existing edges; Clear removes every selected edge. Borders on unselected neighboring cells remain intact.",
                              }),
                            }),
                          ],
                        }),
                    ],
                  }),
                  i.jsxs("section", {
                    className: "cs-preview",
                    "aria-label": r("Cell style preview"),
                    lang: t("Cell style preview"),
                    "data-ui-label": "Cell style preview",
                    children: [
                      i.jsxs("h3", {
                        children: [
                          i.jsx(M, { children: "Preview · " }),
                          s == null ? void 0 : s.name,
                          "!",
                          l,
                        ],
                      }),
                      f === "Border"
                        ? i.jsx("table", {
                            className: "cs-border-preview",
                            "aria-label": r("Border preset preview"),
                            lang: t("Border preset preview"),
                            "data-ui-label": "Border preset preview",
                            children: i.jsx("tbody", {
                              children: [0, 1, 2].map((H) =>
                                i.jsx(
                                  "tr",
                                  {
                                    children: [0, 1, 2].map((Q) => {
                                      const X = `${ut(Q)}${H + 1}`,
                                        J = L(X);
                                      return i.jsx(
                                        "td",
                                        {
                                          "data-preview-cell": X,
                                          style: Dw(J, "", {
                                            defaultBorder: "1px dotted #d5d5d5",
                                          }),
                                          children:
                                            H === 1 && Q === 1 ? "Text" : "",
                                        },
                                        Q,
                                      );
                                    }),
                                  },
                                  H,
                                ),
                              ),
                            }),
                          })
                        : i.jsx("table", {
                            className: "cs-text-preview",
                            children: i.jsx("tbody", {
                              children: i.jsx("tr", {
                                children: i.jsx("td", {
                                  style: Dw(L("B2"), V, {
                                    defaultBorder: "1px solid #d5d5d5",
                                  }),
                                  children: i.jsx("div", {
                                    style: {
                                      ...RN(_),
                                      maxHeight: 94,
                                      overflow: "hidden",
                                      lineHeight: 1.35,
                                    },
                                    children: i.jsx(jI, {
                                      dateSystem:
                                        (G = e.workbook.datePolicy) == null
                                          ? void 0
                                          : G.dateSystem,
                                      value: V === null ? "Audit evidence" : V,
                                      cell: {
                                        ...((s == null
                                          ? void 0
                                          : s.cells[l]) ?? { raw: null }),
                                        style: _,
                                      },
                                    }),
                                  }),
                                }),
                              }),
                            }),
                          }),
                      i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "Values, formulas, number formats and review evidence remain unchanged.",
                        }),
                      }),
                    ],
                  }),
                  (g || C) &&
                    i.jsx("p", {
                      className: "cs-error",
                      role: "alert",
                      children: g || C,
                    }),
                ],
              }),
              i.jsxs("footer", {
                className: "cs-actions",
                children: [
                  i.jsx("button", {
                    type: "button",
                    onClick: e.onClose,
                    children: i.jsx(M, { children: "Cancel" }),
                  }),
                  i.jsx("button", {
                    className: "cs-primary",
                    disabled: B || !!C,
                    children: i.jsx(M, { children: "Apply styles" }),
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    })
  );
}
const zk = {
  greater: "Greater than",
  less: "Less than",
  equal: "Equal to",
  between: "Between",
  contains: "Text contains",
  duplicate: "Duplicate values",
  unique: "Unique values",
  blank: "Blank cells",
  notBlank: "Nonblank cells",
  error: "Errors",
  notError: "No errors",
  formula: "Formula is true",
  colorScale: "Two-color scale",
  dataBar: "Data bars",
  iconSet: "Three-level icons",
};
function t2(e) {
  return {
    id: "",
    range: e,
    kind: "greater",
    value: "0",
    second: "100",
    fill: "#fff2cc",
    color: "#9c5700",
    stopIfTrue: false,
  };
}
function W7(e) {
  return [
    e.style.fill ? `Fill ${e.style.fill}` : "No fill",
    e.style.color ? `Text ${e.style.color}` : "",
    e.dataBar
      ? `Bar ${e.dataBar.width.toFixed(1)}% at ${e.dataBar.start.toFixed(1)}%${e.dataBar.negative ? " negative" : ""}`
      : "",
    e.icon ? `${e.icon.label} icon` : "",
    ...e.errors.map((r) => r.message),
  ]
    .filter(Boolean)
    .join("; ");
}
function Rde(e) {
  var O, j, z, V;
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef({
      book: e.workbook,
      sheetId: e.sheetId,
      range: e.range,
      guard: JSON.stringify(e.workbook),
    }),
    [s, d] = ne.useState(() => {
      var D;
      return structuredClone(
        ((D = e.workbook.sheets.find((B) => B.id === e.sheetId)) == null
          ? void 0
          : D.conditionalFormats) ?? [],
      );
    }),
    [u, l] = ne.useState(() => t2(e.range)),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(0),
    g = e.workbook.sheets.find((D) => D.id === e.sheetId),
    b =
      a.current.guard !== JSON.stringify(e.workbook) ||
      a.current.sheetId !== e.sheetId ||
      a.current.range !== e.range,
    y = e.readOnly || b || !g || !!((O = g.protection) != null && O.enabled),
    w = J3(s),
    E = ["colorScale", "dataBar", "iconSet"].includes(u.kind);
  ne.useEffect(() => {
    var B, _;
    const D = document.activeElement;
    return (
      (_ = (B = o.current) == null ? void 0 : B.querySelector("button")) ==
        null || _.focus(),
      () => {
        D != null && D.isConnected && D.focus();
      }
    );
  }, []);
  const A = (D) => {
      (d(D.map((B, _) => ({ ...B, priority: _ + 1 }))), p(null), h(""));
    },
    k = (D) => {
      (l({ ...u, ...D }), p(null), h(""));
    },
    F = () => {
      try {
        if (y) throw Error("Rules cannot be changed in this view.");
        const D = {
            ...u,
            id: u.id || crypto.randomUUID(),
            ...(E ? { stopIfTrue: !1 } : {}),
          },
          B = u.id ? w.map((C) => (C.id === u.id ? D : C)) : [D, ...w],
          _ = FN(B.map((C, T) => ({ ...C, priority: T + 1 })));
        (A(_), l(t2(e.range)));
      } catch (D) {
        h(D instanceof Error ? D.message : "Rule could not be staged.");
      }
    },
    I = () => {
      try {
        if (y) throw Error("Rules cannot be changed in this view.");
        (p(LB(e.workbook, e.sheetId, s, { readOnly: e.readOnly })),
          v(0),
          h(""));
      } catch (D) {
        h(D instanceof Error ? D.message : "Preview failed.");
      }
    };
  return i.jsx("div", {
    className: "cf-backdrop",
    children: i.jsxs("div", {
      ref: o,
      className: "cf-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      onKeyDown: (D) => {
        var B;
        if (
          (D.stopPropagation(),
          D.key === "Escape" && (D.preventDefault(), e.onClose()),
          D.key === "Tab")
        ) {
          const _ = [
              ...(((B = o.current) == null
                ? void 0
                : B.querySelectorAll(
                    "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
                  )) ?? []),
            ],
            C = _[0],
            T = _.at(-1);
          D.shiftKey && document.activeElement === C
            ? (D.preventDefault(), T == null || T.focus())
            : !D.shiftKey &&
              document.activeElement === T &&
              (D.preventDefault(), C == null || C.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: `${n}-title`,
              children: i.jsx(M, {
                children: "Conditional Formatting Rules Manager",
              }),
            }),
            i.jsx("button", {
              "aria-label": r("Close conditional formatting"),
              onClick: e.onClose,
              lang: t("Close conditional formatting"),
              "data-ui-label": "Close conditional formatting",
              children: "×",
            }),
          ],
        }),
        i.jsxs("p", {
          children: [
            (g == null ? void 0 : g.name) ?? "Worksheet removed",
            " ",
            i.jsx(M, {
              children:
                "· Highest priority first. Stage changes, inspect the preview, then Apply once.",
            }),
          ],
        }),
        b &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(M, {
              children:
                "The workbook or selection changed. Close and reopen this manager to refresh.",
            }),
          }),
        e.readOnly &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(M, { children: "This workbook is read-only." }),
          }),
        ((j = g == null ? void 0 : g.protection) == null
          ? void 0
          : j.enabled) &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(M, {
              children: "Unprotect this worksheet before changing rules.",
            }),
          }),
        s.some((D) => D.priority === void 0) &&
          i.jsx("p", {
            children: i.jsx(M, {
              children:
                "Legacy rules retain their last-wins behavior. Editing the list assigns explicit priorities that preserve the displayed order; the preview shows this migration.",
            }),
          }),
        i.jsxs("div", {
          className: "cf-workspace",
          children: [
            i.jsxs("section", {
              "aria-label": r("Conditional rules"),
              lang: t("Conditional rules"),
              "data-ui-label": "Conditional rules",
              children: [
                i.jsxs("h3", {
                  children: [
                    i.jsx(M, { children: "Rules (" }),
                    s.length,
                    "/100)",
                  ],
                }),
                w.length === 0
                  ? i.jsx("p", {
                      children: i.jsx(M, { children: "No conditional rules." }),
                    })
                  : i.jsx("ol", {
                      className: "cf-rules",
                      children: w.map((D, B) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsx("strong", { children: zk[D.kind] }),
                              " · ",
                              D.range,
                              i.jsxs("span", {
                                children: [
                                  D.formula ?? D.value,
                                  D.kind === "between" ? ` to ${D.second}` : "",
                                  D.stopIfTrue ? " · Stop if true" : "",
                                ],
                              }),
                              i.jsxs("div", {
                                children: [
                                  i.jsxs("button", {
                                    disabled: y,
                                    onClick: () => {
                                      (l(structuredClone(D)), h(""));
                                    },
                                    children: [
                                      i.jsx(M, { children: "Edit rule " }),
                                      B + 1,
                                    ],
                                  }),
                                  i.jsxs("button", {
                                    disabled: y,
                                    onClick: () => {
                                      (A(w.filter((_) => _.id !== D.id)),
                                        u.id === D.id && l(t2(e.range)));
                                    },
                                    children: [
                                      i.jsx(M, { children: "Delete rule " }),
                                      B + 1,
                                    ],
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move rule ${B + 1} up`,
                                    disabled: y || B === 0,
                                    onClick: () => {
                                      const _ = [...w];
                                      (([_[B - 1], _[B]] = [_[B], _[B - 1]]),
                                        A(_));
                                    },
                                    children: "↑",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move rule ${B + 1} down`,
                                    disabled: y || B === w.length - 1,
                                    onClick: () => {
                                      const _ = [...w];
                                      (([_[B + 1], _[B]] = [_[B], _[B + 1]]),
                                        A(_));
                                    },
                                    children: "↓",
                                  }),
                                ],
                              }),
                            ],
                          },
                          D.id,
                        ),
                      ),
                    }),
              ],
            }),
            i.jsxs("form", {
              className: "cf-editor",
              onSubmit: (D) => {
                (D.preventDefault(), F());
              },
              children: [
                i.jsx("h3", { children: u.id ? "Edit rule" : "New rule" }),
                i.jsxs("fieldset", {
                  disabled: y,
                  children: [
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Rule type" }),
                        i.jsx("select", {
                          value: u.kind,
                          onChange: (D) => {
                            const B = D.target.value;
                            k({
                              kind: B,
                              ...(B === "formula"
                                ? { formula: u.formula ?? "=A1>0" }
                                : {}),
                              ...(["colorScale", "dataBar", "iconSet"].includes(
                                B,
                              )
                                ? { stopIfTrue: false }
                                : {}),
                            });
                          },
                          children: BB.map((D) =>
                            i.jsx("option", { value: D, children: zk[D] }, D),
                          ),
                        }),
                      ],
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Applies to" }),
                        i.jsx("input", {
                          value: u.range,
                          onChange: (D) => k({ range: D.target.value }),
                        }),
                      ],
                    }),
                    u.kind === "formula"
                      ? i.jsxs(i.Fragment, {
                          children: [
                            i.jsxs("label", {
                              children: [
                                i.jsx(M, { children: "Condition formula" }),
                                i.jsx("input", {
                                  value: u.formula ?? "",
                                  onChange: (D) =>
                                    k({ formula: D.target.value }),
                                }),
                              ],
                            }),
                            i.jsx("small", {
                              children: i.jsx(M, {
                                children:
                                  "Write relative references for the top-left cell of Applies to. Use $ to anchor rows or columns.",
                              }),
                            }),
                          ],
                        })
                      : [
                          "greater",
                          "less",
                          "equal",
                          "between",
                          "contains",
                        ].includes(u.kind) &&
                        i.jsxs("label", {
                          children: [
                            u.kind === "contains"
                              ? "Text to find"
                              : "Rule value",
                            i.jsx("input", {
                              value: u.value,
                              onChange: (D) => k({ value: D.target.value }),
                            }),
                          ],
                        }),
                    u.kind === "between" &&
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Second value" }),
                          i.jsx("input", {
                            value: u.second,
                            onChange: (D) => k({ second: D.target.value }),
                          }),
                        ],
                      }),
                    i.jsxs("div", {
                      className: "cf-colors",
                      children: [
                        i.jsxs("label", {
                          children: [
                            u.kind === "dataBar"
                              ? "Positive bar color"
                              : "Fill color",
                            i.jsx("input", {
                              type: "color",
                              value: u.fill,
                              onChange: (D) => k({ fill: D.target.value }),
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Text color" }),
                            i.jsx("input", {
                              type: "color",
                              value: u.color,
                              onChange: (D) => k({ color: D.target.value }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    u.kind === "colorScale" &&
                      i.jsxs("div", {
                        className: "cf-colors",
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Minimum color" }),
                              i.jsx("input", {
                                type: "color",
                                value: u.minColor ?? "#f8e99b",
                                onChange: (D) =>
                                  k({ minColor: D.target.value }),
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Maximum color" }),
                              i.jsx("input", {
                                type: "color",
                                value: u.maxColor ?? "#63be7b",
                                onChange: (D) =>
                                  k({ maxColor: D.target.value }),
                              }),
                            ],
                          }),
                        ],
                      }),
                    u.kind === "dataBar" &&
                      i.jsx("small", {
                        children: i.jsx(M, {
                          children:
                            "Automatic numeric minimum and maximum, a zero axis, red negative bars. Values stay visible.",
                        }),
                      }),
                    u.kind === "iconSet" &&
                      i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Icon threshold units" }),
                              i.jsxs("select", {
                                value: u.thresholdType ?? "percent",
                                onChange: (D) =>
                                  k({ thresholdType: D.target.value }),
                                children: [
                                  i.jsx("option", {
                                    value: "percent",
                                    children: "Percent of numeric span",
                                  }),
                                  i.jsx("option", {
                                    value: "number",
                                    children: "Number",
                                  }),
                                ],
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "Middle icon starts at" }),
                              i.jsx(ci, {
                                "aria-label": r("Middle icon starts at"),
                                value:
                                  ((z = u.iconThresholds) == null
                                    ? void 0
                                    : z[0]) ?? 33,
                                onValueChange: (D) => {
                                  var B;
                                  return k({
                                    iconThresholds: [
                                      Number(D),
                                      ((B = u.iconThresholds) == null
                                        ? void 0
                                        : B[1]) ?? 67,
                                    ],
                                  });
                                },
                              }),
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              i.jsx(M, { children: "High icon starts at" }),
                              i.jsx(ci, {
                                "aria-label": r("High icon starts at"),
                                value:
                                  ((V = u.iconThresholds) == null
                                    ? void 0
                                    : V[1]) ?? 67,
                                onValueChange: (D) => {
                                  var B;
                                  return k({
                                    iconThresholds: [
                                      ((B = u.iconThresholds) == null
                                        ? void 0
                                        : B[0]) ?? 33,
                                      Number(D),
                                    ],
                                  });
                                },
                              }),
                            ],
                          }),
                        ],
                      }),
                    i.jsxs("label", {
                      className: "cf-checkbox",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: u.stopIfTrue ?? false,
                          disabled: E || y,
                          onChange: (D) => k({ stopIfTrue: D.target.checked }),
                        }),
                        i.jsx(M, { children: "Stop if true" }),
                      ],
                    }),
                    i.jsx("button", {
                      type: "submit",
                      children: u.id ? "Stage rule edit" : "Stage new rule",
                    }),
                    i.jsx("button", {
                      type: "button",
                      onClick: () => l(t2(e.range)),
                      children: i.jsx(M, { children: "Reset editor" }),
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        c && i.jsx("p", { role: "alert", children: c }),
        i.jsx("button", {
          disabled: y,
          onClick: I,
          children: i.jsx(M, { children: "Preview rule changes" }),
        }),
        f &&
          i.jsxs("section", {
            className: "cf-preview",
            "aria-label": r("Conditional formatting preview"),
            lang: t("Conditional formatting preview"),
            "data-ui-label": "Conditional formatting preview",
            children: [
              i.jsxs("h3", {
                children: [
                  f.changes.length,
                  " ",
                  i.jsx(M, { children: "rule changes" }),
                ],
              }),
              f.changes.length === 0 &&
                i.jsx("p", {
                  children: i.jsx(M, { children: "No changes to apply." }),
                }),
              i.jsxs("p", {
                children: [
                  i.jsx(M, {
                    children:
                      "Exact rule metadata is listed below. The visual sample shows up to",
                  }),
                  " ",
                  f.previewLimit,
                  " ",
                  i.jsx(M, {
                    children:
                      "anchors and populated cells; it is not an exhaustive scan of the range.",
                  }),
                ],
              }),
              f.changes.map((D) =>
                i.jsxs(
                  "details",
                  {
                    children: [
                      i.jsxs("summary", {
                        children: [
                          D.before && D.after
                            ? "Edit"
                            : D.after
                              ? "Add"
                              : "Delete",
                          " ",
                          zk[(D.after ?? D.before).kind],
                          " ·",
                          " ",
                          (D.after ?? D.before).range,
                        ],
                      }),
                      i.jsxs("div", {
                        className: "cf-change",
                        children: [
                          i.jsxs("pre", {
                            children: [
                              i.jsx(M, { children: "Before" }),
                              `
`,
                              D.before
                                ? JSON.stringify(D.before, null, 2)
                                : "None",
                            ],
                          }),
                          i.jsxs("pre", {
                            children: [
                              i.jsx(M, { children: "After" }),
                              `
`,
                              D.after
                                ? JSON.stringify(D.after, null, 2)
                                : "None",
                            ],
                          }),
                        ],
                      }),
                    ],
                  },
                  D.id,
                ),
              ),
              f.errors.length > 0 &&
                i.jsxs("div", {
                  role: "alert",
                  children: [
                    i.jsx("strong", {
                      children: i.jsx(M, { children: "Rule errors" }),
                    }),
                    i.jsx("ul", {
                      children: f.errors.map((D, B) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              D.address || "Range",
                              " · ",
                              D.ruleId,
                              ": ",
                              D.message,
                            ],
                          },
                          B,
                        ),
                      ),
                    }),
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children:
                          "Cells with a rule error receive no formatting from that rule. Apply keeps the rule available for later repair.",
                      }),
                    }),
                  ],
                }),
              i.jsx("div", {
                className: "cf-table",
                children: i.jsxs("table", {
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Cell" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Value" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Before" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "After" }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: f.samples.slice(m * 10, m * 10 + 10).map((D) =>
                        i.jsxs(
                          "tr",
                          {
                            children: [
                              i.jsx("th", { children: D.address }),
                              i.jsx("td", {
                                children: String(D.value ?? ""),
                              }),
                              i.jsx("td", { children: W7(D.before) }),
                              i.jsx("td", { children: W7(D.after) }),
                            ],
                          },
                          D.address,
                        ),
                      ),
                    }),
                  ],
                }),
              }),
              f.samples.length > 10 &&
                i.jsxs("nav", {
                  "aria-label": r("Conditional preview pages"),
                  lang: t("Conditional preview pages"),
                  "data-ui-label": "Conditional preview pages",
                  children: [
                    i.jsx("button", {
                      disabled: m === 0,
                      onClick: () => v(m - 1),
                      children: i.jsx(M, { children: "Previous samples" }),
                    }),
                    i.jsxs("span", {
                      children: [
                        i.jsx(M, { children: "Page " }),
                        m + 1,
                        " ",
                        i.jsx(M, { children: "of " }),
                        Math.ceil(f.samples.length / 10),
                      ],
                    }),
                    i.jsx("button", {
                      disabled: (m + 1) * 10 >= f.samples.length,
                      onClick: () => v(m + 1),
                      children: i.jsx(M, { children: "Next samples" }),
                    }),
                  ],
                }),
            ],
          }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(M, { children: "Cancel" }),
            }),
            i.jsx("button", {
              disabled: y || !f || f.changes.length === 0,
              onClick: () => {
                try {
                  if (!f || y) return;
                  const D = $Q(e.workbook, f, { readOnly: e.readOnly });
                  (D !== e.workbook &&
                    e.onWorkbookChange(
                      D,
                      "Update conditional formatting rules",
                    ),
                    e.onClose());
                } catch (D) {
                  h(
                    D instanceof Error
                      ? D.message
                      : "Rules could not be applied.",
                  );
                }
              },
              children: i.jsx(M, { children: "Apply rules" }),
            }),
          ],
        }),
      ],
    }),
  });
}
function Fz({ decoration: e, children: r }) {
  const { text: t, language: n } = Dr();
  if (
    !(e != null && e.dataBar) &&
    !(e != null && e.icon) &&
    !(e != null && e.errors.length)
  )
    return i.jsx(i.Fragment, { children: r });
  const o = e.dataBar;
  return i.jsxs("span", {
    style: { display: "block", position: "relative", minWidth: 0 },
    children: [
      o &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsx("span", {
              "data-testid": "conditional-data-bar",
              "aria-hidden": "true",
              style: {
                position: "absolute",
                pointerEvents: "none",
                left: `${o.start}%`,
                width: `${o.width}%`,
                top: "12%",
                bottom: "12%",
                background: o.color,
                opacity: 0.28,
              },
            }),
            i.jsx("span", {
              "aria-hidden": "true",
              style: {
                position: "absolute",
                left: `${o.axis}%`,
                top: 0,
                bottom: 0,
                borderLeft: "1px solid #777",
                pointerEvents: "none",
              },
            }),
          ],
        }),
      i.jsxs("span", {
        style: { position: "relative", zIndex: 1 },
        children: [
          e.icon &&
            i.jsx("span", {
              role: "img",
              "aria-label": `${e.icon.label} conditional indicator`,
              title: `${e.icon.label} conditional indicator`,
              style: { color: e.icon.color, marginRight: 4 },
              children: e.icon.symbol,
            }),
          r,
          e.errors.length > 0 &&
            i.jsx("span", {
              role: "img",
              "aria-label": t("Conditional formatting error"),
              title: e.errors.map((a) => `${a.ruleId}: ${a.message}`).join(`
`),
              style: { color: "#a4262c", marginLeft: 3 },
              lang: n("Conditional formatting error"),
              "data-ui-label": "Conditional formatting error",
              children: "⚠",
            }),
        ],
      }),
    ],
  });
}
const r2 = (e, r) => {
    const t = ct(e),
      n = ct(r);
    return (
      t.firstRow <= n.lastRow &&
      t.lastRow >= n.firstRow &&
      t.firstCol <= n.lastCol &&
      t.lastCol >= n.firstCol
    );
  },
  Ide = (e) => {
    const r = ct(e);
    return `${ut(r.firstCol)}${r.firstRow + 1}:${ut(r.lastCol)}${r.lastRow + 1}`;
  },
  Vk = (e) => {
    const r = ct(e);
    return (r.lastRow - r.firstRow + 1) * (r.lastCol - r.firstCol + 1);
  };
function Bz(e, r, t, n = false, o = {}) {
  var f;
  if (o.readOnly) throw Error("This workbook is read-only.");
  const a = Nr(e, r);
  jn(a);
  const s = Ide(t);
  if (Vk(s) > 500)
    throw Error("Preview a merge or unmerge of at most 500 cells.");
  if (!n && Vk(s) < 2) throw Error("Select at least two cells to merge.");
  const d = (a.merges ?? []).filter((p) => r2(p, s));
  if (!n && d.length) throw Error("Unmerge overlapping areas first.");
  if (n && !d.length)
    throw Error("The selected range does not intersect a merged area.");
  const u = n ? d : [s];
  if (u.reduce((p, m) => p + Vk(m), 0) > 500)
    throw Error(
      "Unmerge fewer areas; this preview supports at most 500 affected cells.",
    );
  if (!n) {
    if ((a.tables ?? []).some((p) => r2(p.range, s)))
      throw Error(
        "Merging cells inside a table is unsupported. Convert the table first.",
      );
    if (
      [...(a.pivots ?? []), ...(a.queries ?? [])].some((p) =>
        r2(p.outputRange, s),
      )
    )
      throw Error("Merging generated analysis output is unsupported.");
  }
  rn(kE(e, r, s, n));
  const l = Vn(e);
  if (!n) {
    let p = 0;
    for (const [m, v] of Object.entries(a.cells))
      if (typeof v.raw == "string" && v.raw.startsWith("=") && !v.literal) {
        if (++p > 1e4)
          throw Error(
            "Merge checking supports at most 10,000 active formulas per worksheet.",
          );
        const g = l.getSpillRange(r, m);
        if (g && r2(g, s))
          throw Error(
            `Unmergeable array output at ${g}. Move or remove its formula first.`,
          );
      }
  }
  const c = [];
  for (const p of u) {
    const m = ct(p),
      v = `${ut(m.firstCol)}${m.firstRow + 1}`;
    for (let g = m.firstRow; g <= m.lastRow; g++)
      for (let b = m.firstCol; b <= m.lastCol; b++) {
        const y = `${ut(b)}${g + 1}`,
          w = a.cells[y],
          E = (e.reviewThreads ?? []).filter(
            (A) =>
              !A.orphaned && A.anchor.sheetId === r && A.anchor.address === y,
          ).length;
        (w || E || y === v) &&
          c.push({
            address: y,
            cell: structuredClone(w ?? { raw: null }),
            display: Af(
              l(r, y),
              w,
              (f = e.datePolicy) == null ? void 0 : f.dateSystem,
            ),
            discussions: E,
            effect:
              y === v
                ? "Anchor remains visible"
                : n
                  ? "Revealed by unmerge"
                  : "Hidden until unmerge",
          });
      }
  }
  const h = c.filter(
    (p) =>
      p.effect === "Hidden until unmerge" &&
      p.cell.raw !== null &&
      p.cell.raw !== "",
  ).length;
  return {
    version: 1,
    snapshot: JSON.stringify(e),
    sheetId: r,
    sheetName: a.name,
    range: s,
    unmerge: n,
    affectedRanges: [...u],
    cells: c,
    hiddenValues: h,
    xlsxBlocked: c.some(
      (p) =>
        p.effect === "Hidden until unmerge" &&
        ((p.cell.raw !== null && p.cell.raw !== "") ||
          p.cell.note ||
          p.cell.tickmark ||
          p.cell.hyperlink),
    ),
  };
}
function _de(e, r, t = {}) {
  if (r.snapshot !== JSON.stringify(e))
    throw Error(
      "The workbook changed. Close this preview and review the current cells again.",
    );
  const n = Bz(e, r.sheetId, r.range, r.unmerge, t),
    o = (a) => ({ ...a, cells: a.cells.map(({ display: s, ...d }) => d) });
  if (JSON.stringify(o(n)) !== JSON.stringify(o(r)))
    throw Error("The merge preview changed. Create a fresh preview.");
  return kE(e, r.sheetId, r.range, r.unmerge);
}

export { Bz, Fz, Nde, Rde, _de };
