import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { VU, cn } from "./03-workbook-model.js";
import { Dr, M } from "./04-cell-input-paste-and-fill.js";
import { hz } from "./16-text-split-tools.js";
import { Hi, IW, mS, o4e, vy } from "./42-combine-populations.js";

function s4e(e) {
  const { text, language } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(false),
    [s, d] = ne.useState(e.range),
    [u, l] = ne.useState("Cleaned text"),
    [c, h] = ne.useState(false),
    [f, p] = ne.useState([mS("CLEAN"), mS("TRIM")]),
    [m, v] = ne.useState("SUBSTITUTE"),
    [g, b] = ne.useState(null),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(0),
    [k, F] = ne.useState(e.workbook.textCleanups?.at(-1)?.id ?? ""),
    [I, O] = ne.useState(0),
    j = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((H) => H.id)),
    ].sort(),
    z = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && c,
      allowedSheetIds: j,
      allowNewSheet: e.allowNewSheet ?? true,
    },
    V = JSON.stringify({
      sourceRange: s,
      outputName: u,
      steps: f,
      options: z,
      selection: [e.sheetId, e.range],
    }),
    D = JSON.stringify(e.workbook),
    B = g?.key === V && g.plan.snapshot === D ? g.plan : null,
    _ = ne.useRef({ p: e, key: V, options: z });
  _.current = { p: e, key: V, options: z };
  const C = () => {
    (b(null), w(""), (a.current = false), A(0));
  };
  (ne.useEffect(() => {
    (d(e.range), b(null), w(""), (a.current = false));
  }, [e.sheetId, e.range]),
    ne.useEffect(() => {
      var X;
      const H = document.activeElement;
      return (
        (X = o.current?.querySelector("input:not(:disabled),button")) == null ||
          X.focus(),
        () => {
          H != null && H.isConnected && H.focus();
        }
      );
    }, []));
  const T = (H) =>
      w(
        H instanceof Error
          ? H.message
          : "Cleanup failed. Review the source and steps.",
      ),
    N = () => {
      C();
      try {
        const H = f.map((Q) => ({
          function: Q.name,
          args: Q.fields
            .filter((X) => !(X.optional && X.value === ""))
            .map((X) =>
              X.kind === "number"
                ? X.value.trim()
                  ? Number(X.value)
                  : NaN
                : X.kind === "boolean"
                  ? X.value === "TRUE"
                  : X.value,
            ),
        }));
        b({
          key: V,
          plan: IW(
            e.workbook,
            {
              sourceSheetId: e.sheetId,
              sourceRange: s,
              outputName: u,
              steps: H,
            },
            z,
          ),
        });
      } catch (H) {
        T(H);
      }
    },
    R = () => {
      if (!(!B || a.current))
        try {
          if (_.current.key !== g?.key)
            throw Error("The cleanup settings changed. Preview again.");
          const H = o4e(_.current.p.workbook, B, _.current.options);
          ((a.current = true),
            _.current.p.onWorkbookChange(
              H,
              `Text cleanup: ${B.sourceSheetName}!${B.request.sourceRange} → ${B.request.outputName}`,
            ),
            b(null),
            F(H.textCleanups.at(-1).id),
            O(0));
        } catch (H) {
          T(H);
        }
    },
    L = (e.workbook.textCleanups ?? []).filter((H) => {
      const Q = e.workbook.sheets.find((X) => X.id === H.sourceSheetId);
      return (
        (e.allowedSheetIds === undefined || j.includes(H.sourceSheetId)) &&
        (z.includeHidden ||
          !Q ||
          (!Q.hidden &&
            !Q.veryHidden &&
            !cn(Q, "row").length &&
            !cn(Q, "column").length))
      );
    }),
    $ = L.find((H) => H.id === k) ?? L.at(-1);
  return i.jsx("div", {
    className: "ttc-backdrop",
    onMouseDown: (H) => {
      H.target === H.currentTarget && e.onClose();
    },
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${n}-title`,
      className: "ttc-dialog cleanup-dialog",
      ref: o,
      onKeyDown: (H) => {
        var X, J;
        if (
          (H.stopPropagation(),
          H.key === "Escape" && (H.preventDefault(), e.onClose()),
          H.key === "Tab")
        ) {
          const q = [
            ...(o.current?.querySelectorAll("button,input,select,summary") ??
              []),
          ].filter(
            (W) =>
              !W.matches(":disabled") &&
              !W.closest("details:not([open]) > :not(summary)"),
          );
          H.shiftKey && document.activeElement === q[0]
            ? (H.preventDefault(), (X = q.at(-1)) == null || X.focus())
            : !H.shiftKey &&
              document.activeElement === q.at(-1) &&
              (H.preventDefault(), (J = q[0]) == null || J.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: `${n}-title`,
              children: i.jsx(M, { children: "Clean up text" }),
            }),
            i.jsx("button", {
              "aria-label": text("Close text cleanup"),
              onClick: e.onClose,
              lang: language("Close text cleanup"),
              "data-ui-label": "Close text cleanup",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(M, {
            children:
              "Review text changes before creating a separate worksheet. Original cells, evidence notes and source identities stay in place.",
          }),
        }),
        i.jsxs("fieldset", {
          className: "ttc-fields",
          disabled: e.readOnly || e.allowNewSheet === false,
          children: [
            i.jsx("legend", {
              children: i.jsx(M, { children: "Source and output" }),
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Cleanup source range"),
                  value: s,
                  onChange: (H) => {
                    (d(H.target.value), C());
                  },
                  lang: language("Cleanup source range"),
                  "data-ui-label": "Cleanup source range",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(M, { children: "New worksheet name" }),
                i.jsx("input", {
                  value: u,
                  onChange: (H) => {
                    (l(H.target.value), C());
                  },
                }),
              ],
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "ttc-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: c,
                    onChange: (H) => {
                      (h(H.target.checked), C());
                    },
                  }),
                  i.jsx(M, { children: "Include hidden source cells" }),
                ],
              }),
          ],
        }),
        e.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(M, {
              children:
                "This workbook is read-only. Saved cleanup evidence can still be reviewed.",
            }),
          }),
        i.jsxs("fieldset", {
          disabled: e.readOnly || e.allowNewSheet === false,
          className: "cleanup-steps",
          children: [
            i.jsx("legend", {
              children: i.jsx(M, { children: "Steps in order" }),
            }),
            f.map((H, Q) =>
              i.jsxs(
                "div",
                {
                  className: "cleanup-step",
                  children: [
                    i.jsxs("strong", { children: [Q + 1, ". ", H.name] }),
                    H.fields.map((X, J) =>
                      i.jsxs(
                        "label",
                        {
                          children: [
                            X.label,
                            X.kind === "boolean"
                              ? i.jsxs("select", {
                                  "aria-label": `Step ${Q + 1} ${X.label}`,
                                  value: X.value,
                                  onChange: (q) => {
                                    (p(
                                      f.map((W, K) =>
                                        K === Q
                                          ? {
                                              ...W,
                                              fields: W.fields.map((te, ae) =>
                                                ae === J
                                                  ? {
                                                      ...te,
                                                      value: q.target.value,
                                                    }
                                                  : te,
                                              ),
                                            }
                                          : W,
                                      ),
                                    ),
                                      C());
                                  },
                                  children: [
                                    i.jsx("option", { children: "TRUE" }),
                                    i.jsx("option", { children: "FALSE" }),
                                  ],
                                })
                              : i.jsx("input", {
                                  "aria-label": `Step ${Q + 1} ${X.label}`,
                                  value: X.value,
                                  onChange: (q) => {
                                    (p(
                                      f.map((W, K) =>
                                        K === Q
                                          ? {
                                              ...W,
                                              fields: W.fields.map((te, ae) =>
                                                ae === J
                                                  ? {
                                                      ...te,
                                                      value: q.target.value,
                                                    }
                                                  : te,
                                              ),
                                            }
                                          : W,
                                      ),
                                    ),
                                      C());
                                  },
                                }),
                          ],
                        },
                        J,
                      ),
                    ),
                    i.jsx("button", {
                      "aria-label": `Move step ${Q + 1} up`,
                      disabled: Q === 0,
                      onClick: () => {
                        const X = [...f];
                        (([X[Q - 1], X[Q]] = [X[Q], X[Q - 1]]), p(X), C());
                      },
                      children: "↑",
                    }),
                    i.jsx("button", {
                      "aria-label": `Remove step ${Q + 1}`,
                      onClick: () => {
                        (p(f.filter((X, J) => Q !== J)), C());
                      },
                      children: i.jsx(M, { children: "Remove" }),
                    }),
                  ],
                },
                Q,
              ),
            ),
            i.jsxs("div", {
              className: "cleanup-add",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Add function" }),
                    i.jsx("select", {
                      value: m,
                      onChange: (H) => v(H.target.value),
                      children: VU.map((H) =>
                        i.jsx("option", { children: H }, H),
                      ),
                    }),
                  ],
                }),
                i.jsx("button", {
                  disabled: f.length >= 12,
                  onClick: () => {
                    (p([...f, mS(m)]), C());
                  },
                  children: i.jsx(M, { children: "Add step" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (p([
                      ...f,
                      {
                        name: "SUBSTITUTE",
                        fields: [
                          Hi("Old text", " "),
                          Hi("New text", " "),
                          Hi("Occurrence (blank = all)", "", "number", true),
                        ],
                      },
                    ]),
                      C());
                  },
                  disabled: f.length >= 12,
                  children: i.jsx(M, {
                    children: "Replace nonbreaking spaces",
                  }),
                }),
              ],
            }),
          ],
        }),
        i.jsx("p", {
          className: "ttc-boundary",
          children: i.jsx(M, {
            children:
              "TRIM removes ordinary spaces; CLEAN removes ASCII control characters. Nonbreaking spaces need an explicit SUBSTITUTE step before TRIM. Character positions count Unicode code points. VALUE uses US numeric notation; NUMBERVALUE accepts explicit decimal and group separators. Numeric conversion removes leading zeros.",
          }),
        }),
        i.jsx("button", {
          onClick: N,
          disabled: e.readOnly || e.allowNewSheet === false,
          children: i.jsx(M, { children: "Preview cleanup" }),
        }),
        y && i.jsx("p", { className: "ttc-error", role: "alert", children: y }),
        g &&
          !B &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(M, {
              children:
                "The workbook or settings changed. Preview again before applying.",
            }),
          }),
        B &&
          i.jsxs("section", {
            className: "ttc-preview",
            "aria-label": text("Text cleanup preview"),
            lang: language("Text cleanup preview"),
            "data-ui-label": "Text cleanup preview",
            children: [
              i.jsxs("h3", {
                children: [
                  B.rows.length,
                  " ",
                  i.jsx(M, { children: "results · " }),
                  B.errors.length,
                  " ",
                  i.jsx(M, { children: "errors" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  i.jsx(M, { children: "Source revision " }),
                  e.workbook.revision,
                  i.jsx(M, {
                    children:
                      ". Quoted values expose spaces; nonbreaking spaces appear as \\u00a0.",
                  }),
                ],
              }),
              i.jsx("code", { children: B.formulas[0] }),
              i.jsx("div", {
                className: "ttc-scroll",
                children: i.jsxs("table", {
                  children: [
                    i.jsx("thead", {
                      children: i.jsxs("tr", {
                        children: [
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Source cell" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, {
                              children: "Original evidence value",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Result" }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(M, { children: "Type / error" }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: B.rows.slice(E * 20, E * 20 + 20).map((H) => {
                        return i.jsxs(
                          "tr",
                          {
                            children: [
                              i.jsx("td", { children: H.sourceAddress }),
                              i.jsx("td", {
                                children: i.jsx("code", {
                                  children: vy(H.original?.raw ?? null),
                                }),
                              }),
                              i.jsx("td", {
                                children: i.jsx("code", {
                                  children: vy(H.result),
                                }),
                              }),
                              i.jsx("td", {
                                children:
                                  B.errors.find(
                                    (J) => J.sourceAddress === H.sourceAddress,
                                  )?.error ?? typeof H.result,
                              }),
                            ],
                          },
                          H.sourceAddress,
                        );
                      }),
                    }),
                  ],
                }),
              }),
              i.jsxs("div", {
                className: "ttc-pages",
                children: [
                  i.jsx("button", {
                    disabled: E === 0,
                    onClick: () => A(E - 1),
                    children: i.jsx(M, { children: "Previous results" }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(M, { children: "Page " }),
                      E + 1,
                      " ",
                      i.jsx(M, { children: "of " }),
                      Math.ceil(B.rows.length / 20),
                    ],
                  }),
                  i.jsx("button", {
                    disabled: (E + 1) * 20 >= B.rows.length,
                    onClick: () => A(E + 1),
                    children: i.jsx(M, { children: "Next results" }),
                  }),
                ],
              }),
              B.errors.length > 0 &&
                i.jsx("p", {
                  role: "alert",
                  children: i.jsx(M, {
                    children:
                      "Resolve all errors before applying. No partial output will be created.",
                  }),
                }),
            ],
          }),
        L.length > 0 &&
          i.jsxs("details", {
            children: [
              i.jsxs("summary", {
                children: [
                  i.jsx(M, { children: "Saved text cleanup evidence (" }),
                  L.length,
                  ")",
                ],
              }),
              i.jsxs("label", {
                children: [
                  i.jsx(M, { children: "Cleanup record" }),
                  i.jsx("select", {
                    value: $?.id ?? "",
                    onChange: (H) => {
                      (F(H.target.value), O(0));
                    },
                    children: L.map((H) =>
                      i.jsxs(
                        "option",
                        {
                          value: H.id,
                          children: [
                            H.outputSheetName,
                            " · source revision ",
                            H.sourceRevision,
                            " ·",
                            " ",
                            H.createdAt,
                          ],
                        },
                        H.id,
                      ),
                    ),
                  }),
                ],
              }),
              $ &&
                i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("p", {
                      children: [
                        i.jsx(M, { children: "Source: " }),
                        $.sourceSheetName,
                        " ",
                        i.jsx(M, { children: "· workbook" }),
                        " ",
                        $.workbookId,
                        " ",
                        i.jsx(M, { children: "· revision " }),
                        $.sourceRevision,
                        i.jsx(M, {
                          children:
                            ". Addresses and values are snapshots from the cleanup, even if cells later move or are deleted.",
                        }),
                      ],
                    }),
                    i.jsx("code", { children: vy($.steps) }),
                    i.jsx("div", {
                      className: "ttc-scroll",
                      children: i.jsxs("table", {
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Source at cleanup",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Original value and note",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Result at cleanup",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(M, {
                                    children: "Original file row",
                                  }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: $.rows
                              .slice(I * 20, I * 20 + 20)
                              .map((H) =>
                                i.jsxs(
                                  "tr",
                                  {
                                    children: [
                                      i.jsx("td", {
                                        children: H.sourceAddress,
                                      }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: vy(
                                            H.original ?? { raw: null },
                                          ),
                                        }),
                                      }),
                                      i.jsx("td", {
                                        children: i.jsx("code", {
                                          children: vy(H.result),
                                        }),
                                      }),
                                      i.jsx("td", {
                                        children: H.origin
                                          ? `${H.origin.sourceName} · ${H.origin.sourceSheet} · row ${H.origin.originalRow}`
                                          : "Editor value snapshot",
                                      }),
                                    ],
                                  },
                                  H.sourceAddress,
                                ),
                              ),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("button", {
                      disabled: I === 0,
                      onClick: () => O(I - 1),
                      children: i.jsx(M, { children: "Previous evidence" }),
                    }),
                    i.jsx("button", {
                      disabled: (I + 1) * 20 >= $.rows.length,
                      onClick: () => O(I + 1),
                      children: i.jsx(M, { children: "Next evidence" }),
                    }),
                  ],
                }),
            ],
          }),
        i.jsx("p", {
          className: "ttc-boundary",
          children: i.jsx(M, {
            children:
              "Up to 500 ordinary source cells, 200,000 source characters and 12 steps. Keep a native JSON backup to retain cleanup history; CSV/XLSX do not preserve that history.",
          }),
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(M, { children: "Close" }),
            }),
            i.jsx("button", {
              className: "ttc-primary",
              disabled: !B || !!B.errors.length || e.readOnly || a.current,
              onClick: R,
              children: i.jsx(M, { children: "Create cleaned worksheet" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const DW = [
    ["trim", "Trim whitespace"],
    ["clean", "Remove control characters"],
    ["case", "Change case"],
    ["replace", "Replace text"],
    ["split", "Split fields"],
    ["concat", "Concatenate fields"],
    ["convert", "Convert type"],
    ["fill", "Fill gaps"],
    ["select", "Select and reorder fields"],
    ["pattern", "Infer from examples"],
  ],
  gS = (e, r = false) => {
    if (!e.trim() && r) return [];
    if (!/^\s*\d+(?:\s*,\s*\d+)*\s*$/.test(e))
      throw Error("Enter field positions separated by commas, such as 1, 3.");
    const t = e.split(",").map((n) => Number(n) - 1);
    if (t.some((n) => n < 0 || n > 49) || new Set(t).size !== t.length)
      throw Error("Choose distinct field positions from 1 to 50.");
    return t;
  },
  vS = (e) => JSON.stringify(e);
function cj(e) {
  if (e.kind === "pattern")
    return `Infer “${e.name}”: ${hz(e.pattern)} (${e.examples.length} examples)`;
  if (e.kind === "split")
    return `Split field ${e.column + 1}: ${e.mode === "fixed" ? "breaks " + e.breaks.join(", ") : "delimiters " + JSON.stringify(e.delimiters) + (e.collapse ? ", collapse consecutive" : "")}`;
  const r = e.columns.map((n) => n + 1).join(", ");
  return `${DW.find((n) => n[0] === e.kind)?.[1]} · fields ${r}${e.kind === "convert" ? ` → ${e.to} (decimal ${e.decimal}, groups ${e.group || "none"}, dates ${e.dateOrder})` : e.kind === "fill" ? ` → ${e.direction}${e.groupColumns.length ? " within fields " + e.groupColumns.map((n) => n + 1).join(", ") : ""}` : e.kind === "concat" ? ` → ${e.name}; separator ${JSON.stringify(e.separator)}` : e.kind === "case" || e.kind === "trim" ? ` → ${e.mode}` : e.kind === "replace" ? ` ${JSON.stringify(e.find)} → ${JSON.stringify(e.replacement)}` : ""}`;
}

export { DW, cj, gS, s4e, vS };
