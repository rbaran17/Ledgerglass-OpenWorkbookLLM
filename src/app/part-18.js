import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { Yl, hr, pg } from "./part-01.js";
import { Ta } from "./part-03.js";
import { Dr, M } from "./part-04.js";
import { CT, ET, Lce, Pce, Q4, Uce, Vce, Wce, yz, zce } from "./part-17.js";

function Hce(e) {
  const { text: r, language: t } = Dr(),
    n = ne.useRef(null),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const [s, d] = ne.useState(() => ({
      engine: pg,
      source: {
        sheetId: e.sheetId,
        range: e.range.includes(":")
          ? e.range
          : Ta(e.workbook.sheets.find((ye) => ye.id === e.sheetId)),
      },
      steps: [],
    })),
    [u, l] = ne.useState("Transformation recipe"),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState(0),
    [m, v] = ne.useState(""),
    [g, b] = ne.useState(1),
    [y, w] = ne.useState(-1),
    [E, A] = ne.useState("types"),
    [k, F] = ne.useState("Recipe result"),
    [I, O] = ne.useState(false),
    [j, z] = ne.useState(null),
    [V, D] = ne.useState(null),
    [B, _] = ne.useState(""),
    [C, T] = ne.useState("output"),
    [N, R] = ne.useState(0),
    L = [
      ...new Set(e.allowedSheetIds ?? e.workbook.sheets.map((ye) => ye.id)),
    ].sort(),
    $ = {
      readOnly: e.readOnly,
      includeHidden: !!e.allowHidden && I,
      allowedSheetIds: L,
    },
    U = e.workbook.sheets.filter(
      (ye) =>
        L.includes(ye.id) &&
        ($.includeHidden || (!ye.hidden && !ye.veryHidden)),
    ),
    G = e.workbook.transformationRecipes ?? [],
    H = G.find((ye) => ye.id === m),
    Q = G.find((ye) => ye.id === c),
    X$1 = Q == null ? void 0 : Q.versions.find((ye) => ye.version === f),
    J = !X$1 || X$1.definitionHash !== hr(s),
    q = JSON.stringify([e.workbook, e.sheetId, e.range, s, c, f, k, $]),
    W = (j == null ? void 0 : j.stamp) === q,
    K = W ? j.plan.run : void 0;
  function te() {
    (z(null), D(null), R(0));
  }
  function ae(ye) {
    (d(ye), te(), _(""));
  }
  function oe(ye) {
    try {
      const Te = ET(e.workbook, { ...s, steps: s.steps.slice(0, ye) }, $);
      return Te.status === "complete"
        ? { fields: Te.output.fields, error: "" }
        : {
            fields: Te.failure.inputSchema.map((Pe) => Pe.name),
            error: Te.failure.message,
          };
    } catch (Te) {
      return { fields: [], error: Te.message };
    }
  }
  const de = ne.useMemo(
      () => oe(Math.max(0, y)),
      [e.workbook, s, y, JSON.stringify($)],
    ),
    he = s.steps[y],
    Se = ne.useMemo(() => {
      if (!he || !["merge", "append"].includes(he.kind)) return [];
      try {
        const ye = ET(
          e.workbook,
          { engine: pg, source: he.source, steps: [] },
          $,
        );
        return ye.status === "complete" ? ye.output.fields : [];
      } catch {
        return [];
      }
    }, [e.workbook, he, JSON.stringify($)]);
  (ne.useEffect(() => {
    var Te, Pe;
    const ye = document.activeElement;
    return (
      (Pe =
        (Te = n.current) == null
          ? void 0
          : Te.querySelector("input,select,button")) == null || Pe.focus(),
      () => {
        ye != null && ye.isConnected && ye.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      var ye, Te;
      W &&
        ((Te = (ye = o.current) == null ? void 0 : ye.scrollIntoView) == null ||
          Te.call(ye, { block: "start" }));
    }, [j, W]));
  const le = (ye) => {
    try {
      ye();
    } catch (Te) {
      D({ text: Te.message, stamp: q });
    }
  };
  function ce() {
    le(() => {
      const ye = G.find((Pe) => Pe.id === m),
        Te = ye == null ? void 0 : ye.versions.find((Pe) => Pe.version === g);
      if (!ye || !Te) throw Error("Choose a saved recipe version.");
      (h(ye.id),
        p(Te.version),
        l(ye.name),
        d(structuredClone(Te.definition)),
        w(Te.definition.steps.length ? 0 : -1),
        te(),
        _(
          "Saved coordinates are unchanged by structural edits. Review the source and steps, then preview current data.",
        ));
    });
  }
  function ge() {
    le(() => {
      const ye = Lce(e.workbook, u, s, c || void 0, $),
        Te = ye.transformationRecipes.find((Pe) =>
          c ? Pe.id === c : Pe.name === u.trim(),
        );
      (e.onWorkbookChange(
        ye,
        `Saved recipe ${Te.name} version ${Te.versions.length}`,
      ),
        d(Te.versions.at(-1).definition),
        h(Te.id),
        p(Te.versions.length),
        v(Te.id),
        b(Te.versions.length),
        te(),
        _(
          `Saved version ${Te.versions.length}. Preview this version before publishing.`,
        ));
    });
  }
  function se() {
    le(() => {
      const ye = yz(
        e.workbook,
        {
          definition: s,
          ...(!J && c ? { recipeId: c, version: f } : {}),
          outputName: k,
        },
        $,
      );
      (z({ plan: ye, stamp: q }),
        R(0),
        T(ye.run.status === "failed" ? "steps" : "output"),
        D(null));
    });
  }
  function we() {
    le(() => {
      if (!W || !j)
        throw Error("Preview the current recipe before publishing.");
      const ye = Uce(e.workbook, j.plan, $);
      (te(),
        e.onWorkbookChange(ye, `Published recipe ${u} version ${f}`),
        e.onClose());
    });
  }
  function Ee() {
    if (!W || !j) return;
    const { snapshot: ye, ...Te } = j.plan,
      Pe = URL.createObjectURL(
        new Blob([JSON.stringify(Te, null, 2)], { type: "application/json" }),
      ),
      qe = document.createElement("a");
    ((qe.href = Pe),
      (qe.download = "Recipe-execution-review.json"),
      qe.click(),
      setTimeout(() => URL.revokeObjectURL(Pe), 5e3));
  }
  const Ae = e.workbook.sheets.filter((ye) => {
    var Te;
    return ((Te = ye.transformationRun) == null ? void 0 : Te.recipeId) === c;
  });
  return i.jsx("div", {
    className: "recipe-backdrop",
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": "recipe-title",
      className: "recipe-dialog",
      ref: n,
      onKeyDown: (ye) => {
        var Te;
        if (
          (ye.key === "Escape" && (ye.preventDefault(), a.current()),
          ye.key === "Tab")
        ) {
          const Pe = Array.from(
              ((Te = n.current) == null
                ? void 0
                : Te.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary,[tabindex="0"]',
                  )) ?? [],
            ).filter(
              (Re) =>
                !Re.closest("details:not([open])") || Re.tagName === "SUMMARY",
            ),
            qe = Pe[0],
            ht = Pe.at(-1);
          ye.shiftKey && document.activeElement === qe
            ? (ye.preventDefault(), ht == null || ht.focus())
            : !ye.shiftKey &&
              document.activeElement === ht &&
              (ye.preventDefault(), qe == null || qe.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("h2", {
                  id: "recipe-title",
                  children: i.jsx(M, { children: "Transformation recipes" }),
                }),
                i.jsx("p", {
                  children: i.jsx(M, {
                    children:
                      "Save ordered steps. Review each run before publishing a separate result.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close transformation recipes"),
              onClick: e.onClose,
              lang: t("Close transformation recipes"),
              "data-ui-label": "Close transformation recipes",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "recipe-body",
          children: [
            i.jsxs("section", {
              className: "recipe-saved",
              "aria-label": r("Saved transformation recipes"),
              lang: t("Saved transformation recipes"),
              "data-ui-label": "Saved transformation recipes",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Saved recipe" }),
                    i.jsxs("select", {
                      "aria-label": r("Saved recipe"),
                      value: m,
                      onChange: (ye) => {
                        var Te;
                        (v(ye.target.value),
                          b(
                            ((Te = G.find((Pe) => Pe.id === ye.target.value)) ==
                            null
                              ? void 0
                              : Te.versions.length) ?? 1,
                          ));
                      },
                      lang: t("Saved recipe"),
                      "data-ui-label": "Saved recipe",
                      children: [
                        i.jsx("option", {
                          value: "",
                          children: "Choose a recipe",
                        }),
                        G.map((ye) =>
                          i.jsx(
                            "option",
                            { value: ye.id, children: ye.name },
                            ye.id,
                          ),
                        ),
                      ],
                    }),
                  ],
                }),
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Version" }),
                    i.jsx("select", {
                      "aria-label": r("Saved recipe version"),
                      value: g,
                      onChange: (ye) => b(Number(ye.target.value)),
                      lang: t("Saved recipe version"),
                      "data-ui-label": "Saved recipe version",
                      children: ((H == null ? void 0 : H.versions) ?? []).map(
                        (ye) =>
                          i.jsx(
                            "option",
                            { value: ye.version, children: ye.version },
                            ye.version,
                          ),
                      ),
                    }),
                  ],
                }),
                i.jsx("button", {
                  disabled: !H,
                  onClick: ce,
                  children: i.jsx(M, { children: "Load version" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (h(""),
                      p(0),
                      l("Transformation recipe"),
                      d({ engine: pg, source: s.source, steps: [] }),
                      w(-1),
                      te(),
                      _("New recipe draft."));
                  },
                  children: i.jsx(M, { children: "New recipe" }),
                }),
                i.jsx("button", {
                  disabled: !H || e.readOnly,
                  onClick: () =>
                    le(() => {
                      const ye = Pce(e.workbook, m, $);
                      (e.onWorkbookChange(
                        ye,
                        "Deleted saved recipe definition; outputs retained",
                      ),
                        c === m && (h(""), p(0)),
                        v(""),
                        te(),
                        _(
                          "Saved definition deleted. Historical outputs and their run evidence are retained.",
                        ));
                    }),
                  children: i.jsx(M, { children: "Delete saved recipe" }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "recipe-line",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "Recipe name" }),
                    i.jsx("input", {
                      "aria-label": r("Recipe name"),
                      value: u,
                      onChange: (ye) => l(ye.target.value),
                      maxLength: 128,
                      lang: t("Recipe name"),
                      "data-ui-label": "Recipe name",
                    }),
                  ],
                }),
                i.jsx("p", {
                  className: "recipe-version",
                  children: c
                    ? `Version ${f}${J ? " · edited draft" : " · saved"}`
                    : "Unsaved draft",
                }),
                i.jsx("button", {
                  disabled: e.readOnly,
                  onClick: ge,
                  children: c ? "Save next version" : "Save new recipe",
                }),
              ],
            }),
            i.jsx(CT, {
              label: r("Primary source"),
              source: s.source,
              sheets: U,
              onChange: (ye) => ae({ ...s, source: ye }),
            }),
            e.allowHidden &&
              i.jsxs("label", {
                className: "recipe-check",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: I,
                    onChange: (ye) => {
                      (O(ye.target.checked), te());
                    },
                  }),
                  i.jsx(M, {
                    children:
                      "Include hidden source cells and formula dependencies",
                  }),
                ],
              }),
            i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "Sources include every record in the declared ranges, regardless of current filter views. Tables use their current headers and data body. Source cells stay unchanged.",
              }),
            }),
            i.jsxs("div", {
              className: "recipe-layout",
              children: [
                i.jsxs("aside", {
                  "aria-label": r("Ordered recipe steps"),
                  lang: t("Ordered recipe steps"),
                  "data-ui-label": "Ordered recipe steps",
                  children: [
                    i.jsx("h3", {
                      children: i.jsx(M, { children: "Applied steps" }),
                    }),
                    i.jsx("ol", {
                      children: s.steps.map((ye, Te) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              i.jsxs("button", {
                                "aria-pressed": y === Te,
                                onClick: () => w(Te),
                                children: [Te + 1, ". ", Q4[ye.kind]],
                              }),
                              i.jsxs("span", {
                                children: [
                                  i.jsx("button", {
                                    "aria-label": `Move step ${Te + 1} up`,
                                    disabled: Te === 0,
                                    onClick: () => {
                                      const Pe = [...s.steps];
                                      (([Pe[Te - 1], Pe[Te]] = [
                                        Pe[Te],
                                        Pe[Te - 1],
                                      ]),
                                        ae({ ...s, steps: Pe }),
                                        w(Te - 1));
                                    },
                                    children: "↑",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Move step ${Te + 1} down`,
                                    disabled: Te === s.steps.length - 1,
                                    onClick: () => {
                                      const Pe = [...s.steps];
                                      (([Pe[Te], Pe[Te + 1]] = [
                                        Pe[Te + 1],
                                        Pe[Te],
                                      ]),
                                        ae({ ...s, steps: Pe }),
                                        w(Te + 1));
                                    },
                                    children: "↓",
                                  }),
                                  i.jsx("button", {
                                    "aria-label": `Remove step ${Te + 1}`,
                                    onClick: () => {
                                      (ae({
                                        ...s,
                                        steps: s.steps.filter(
                                          (Pe, qe) => qe !== Te,
                                        ),
                                      }),
                                        w(Math.min(Te, s.steps.length - 2)));
                                    },
                                    children: "×",
                                  }),
                                ],
                              }),
                            ],
                          },
                          ye.id,
                        ),
                      ),
                    }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Add transformation" }),
                        i.jsx("select", {
                          "aria-label": r("New transformation type"),
                          value: E,
                          onChange: (ye) => A(ye.target.value),
                          lang: t("New transformation type"),
                          "data-ui-label": "New transformation type",
                          children: Object.entries(Q4).map(([ye, Te]) =>
                            i.jsx("option", { value: ye, children: Te }, ye),
                          ),
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      disabled: s.steps.length >= Yl.steps,
                      onClick: () => {
                        const ye = oe(s.steps.length).fields;
                        (ae({
                          ...s,
                          steps: [...s.steps, Vce(E, ye, s.source)],
                        }),
                          w(s.steps.length));
                      },
                      children: i.jsx(M, { children: "Add step" }),
                    }),
                    i.jsxs("p", {
                      children: [
                        s.steps.length,
                        " / ",
                        Yl.steps,
                        " ",
                        i.jsx(M, { children: "steps" }),
                      ],
                    }),
                  ],
                }),
                i.jsx("section", {
                  className: "recipe-editor",
                  "aria-label": r("Recipe step editor"),
                  lang: t("Recipe step editor"),
                  "data-ui-label": "Recipe step editor",
                  children: he
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("h3", {
                            children: [y + 1, ". ", Q4[he.kind]],
                          }),
                          de.error &&
                            i.jsxs("p", {
                              role: "status",
                              children: [
                                i.jsx(M, {
                                  children:
                                    "Earlier step or source needs attention: ",
                                }),
                                de.error,
                              ],
                            }),
                          i.jsx(zce, {
                            step: he,
                            fields: de.fields,
                            rightFields: Se,
                            sheets: U,
                            onChange: (ye) =>
                              ae({
                                ...s,
                                steps: s.steps.map((Te, Pe) =>
                                  Pe === y ? ye : Te,
                                ),
                              }),
                          }),
                        ],
                      })
                    : i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "Add a transformation to define the sequence. Each step reads the previous step's complete result.",
                        }),
                      }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "recipe-line",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(M, { children: "New output worksheet" }),
                    i.jsx("input", {
                      "aria-label": r("Recipe output worksheet"),
                      value: k,
                      onChange: (ye) => {
                        (F(ye.target.value), te());
                      },
                      maxLength: 31,
                      lang: t("Recipe output worksheet"),
                      "data-ui-label": "Recipe output worksheet",
                    }),
                  ],
                }),
                i.jsx("button", {
                  className: "recipe-primary",
                  onClick: se,
                  children: i.jsx(M, { children: "Preview recipe" }),
                }),
              ],
            }),
            B && i.jsx("p", { role: "status", children: B }),
            (V == null ? void 0 : V.stamp) === q &&
              i.jsx("p", {
                role: "alert",
                className: "recipe-error",
                children: V.text,
              }),
            j &&
              !W &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(M, {
                  children:
                    "Workbook, selection or permissions changed. Preview again; previous results are withheld.",
                }),
              }),
            Ae.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(M, { children: "Previous saved outputs (" }),
                      Ae.length,
                      ")",
                    ],
                  }),
                  Ae.map((ye) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          ye.name,
                          " ",
                          i.jsx(M, { children: "· version " }),
                          ye.transformationRun.version,
                          " ·",
                          " ",
                          ye.transformationRun.outputRows,
                          " ",
                          i.jsx(M, { children: "rows ·" }),
                          " ",
                          ye.transformationRun.at,
                          i.jsx(M, {
                            children:
                              ". Historical result; use a fresh preview to compare.",
                          }),
                        ],
                      },
                      ye.id,
                    ),
                  ),
                ],
              }),
            K &&
              i.jsxs("section", {
                className: "recipe-review",
                "aria-label": r("Recipe execution preview"),
                ref: o,
                lang: t("Recipe execution preview"),
                "data-ui-label": "Recipe execution preview",
                children: [
                  i.jsx("h3", {
                    children:
                      K.status === "complete"
                        ? `Complete preview · ${K.output.rows.length} records`
                        : `Stopped at ${K.failure.index === 0 ? "source" : `step ${K.failure.index}`}`,
                  }),
                  K.status === "failed" &&
                    i.jsxs("p", {
                      role: "alert",
                      className: "recipe-error",
                      children: [
                        K.failure.message,
                        " ",
                        i.jsx(M, {
                          children:
                            "No output will be published. Downstream steps were skipped.",
                        }),
                      ],
                    }),
                  J &&
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children:
                          "Draft preview. Save a version and preview it before publishing.",
                      }),
                    }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(M, { children: "Definition fingerprint " }),
                      i.jsx("code", { children: K.definitionHash }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(M, { children: "Source fingerprint " }),
                      i.jsx("code", { children: K.sourceHash }),
                    ],
                  }),
                  i.jsx("nav", {
                    "aria-label": r("Recipe review views"),
                    lang: t("Recipe review views"),
                    "data-ui-label": "Recipe review views",
                    children: [
                      ["output", "Output records"],
                      ["steps", "Step schemas & results"],
                      ["sources", "Sources & control totals"],
                    ].map(([ye, Te]) =>
                      i.jsx(
                        "button",
                        {
                          "aria-pressed": C === ye,
                          disabled: ye === "output" && K.status !== "complete",
                          onClick: () => T(ye),
                          children: Te,
                        },
                        ye,
                      ),
                    ),
                  }),
                  C === "output" &&
                    K.status === "complete" &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("div", {
                          className: "recipe-table",
                          tabIndex: 0,
                          "aria-label": r("Complete recipe output"),
                          lang: t("Complete recipe output"),
                          "data-ui-label": "Complete recipe output",
                          children: [
                            i.jsxs("table", {
                              children: [
                                i.jsx("thead", {
                                  children: i.jsxs("tr", {
                                    children: [
                                      K.output.fields.map((ye) =>
                                        i.jsx("th", { children: ye }, ye),
                                      ),
                                      i.jsx("th", {
                                        children: i.jsx(M, {
                                          children: "Original source rows",
                                        }),
                                      }),
                                    ],
                                  }),
                                }),
                                i.jsx("tbody", {
                                  children: K.output.rows
                                    .slice(N * 25, (N + 1) * 25)
                                    .map((ye, Te) =>
                                      i.jsxs(
                                        "tr",
                                        {
                                          children: [
                                            ye.values.map((Pe, qe) =>
                                              i.jsx(
                                                "td",
                                                { children: Wce(Pe) },
                                                qe,
                                              ),
                                            ),
                                            i.jsx("td", {
                                              children: i.jsxs("details", {
                                                children: [
                                                  i.jsxs("summary", {
                                                    children: [
                                                      ye.sources.length,
                                                      " ",
                                                      i.jsx(M, {
                                                        children: "source row",
                                                      }),
                                                      ye.sources.length === 1
                                                        ? ""
                                                        : "s",
                                                    ],
                                                  }),
                                                  ye.sources.map((Pe) =>
                                                    i.jsx(
                                                      "p",
                                                      { children: Pe },
                                                      Pe,
                                                    ),
                                                  ),
                                                ],
                                              }),
                                            }),
                                          ],
                                        },
                                        Te,
                                      ),
                                    ),
                                }),
                              ],
                            }),
                            !K.output.rows.length &&
                              i.jsx("p", {
                                children: i.jsx(M, {
                                  children:
                                    "No records remain. Source records are retained.",
                                }),
                              }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "recipe-line",
                          children: [
                            i.jsx("button", {
                              disabled: N === 0,
                              onClick: () => R((ye) => ye - 1),
                              children: i.jsx(M, {
                                children: "Previous recipe records",
                              }),
                            }),
                            i.jsxs("span", {
                              children: [
                                i.jsx(M, { children: "Page " }),
                                N + 1,
                                " ",
                                i.jsx(M, { children: "of" }),
                                " ",
                                Math.max(
                                  1,
                                  Math.ceil(K.output.rows.length / 25),
                                ),
                                " ·",
                                " ",
                                K.output.rows.length,
                                " ",
                                i.jsx(M, { children: "records" }),
                              ],
                            }),
                            i.jsx("button", {
                              disabled: (N + 1) * 25 >= K.output.rows.length,
                              onClick: () => R((ye) => ye + 1),
                              children: i.jsx(M, {
                                children: "Next recipe records",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  C === "steps" &&
                    i.jsx("div", {
                      "aria-label": r("Recipe step trace"),
                      lang: t("Recipe step trace"),
                      "data-ui-label": "Recipe step trace",
                      children: K.steps.map((ye) => {
                        var Te;
                        return i.jsxs(
                          "details",
                          {
                            open: ye.status === "failed",
                            children: [
                              i.jsxs("summary", {
                                children: [
                                  ye.index,
                                  ".",
                                  " ",
                                  Q4[ye.kind] ?? ye.kind,
                                  " ",
                                  "· ",
                                  ye.status,
                                  " · ",
                                  ye.inputRows,
                                  " → ",
                                  ye.outputRows ?? "—",
                                  " ",
                                  i.jsx(M, { children: "rows" }),
                                ],
                              }),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(M, { children: "Input schema:" }),
                                  " ",
                                  ye.inputSchema
                                    .map(
                                      (Pe) =>
                                        `${Pe.name} (${Pe.types.join(" / ") || "empty"}${Pe.declaredType ? `; declared ${Pe.declaredType}` : ""})`,
                                    )
                                    .join(", ") || "not executed",
                                ],
                              }),
                              ye.outputSchema &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(M, { children: "Output schema:" }),
                                    " ",
                                    ye.outputSchema
                                      .map(
                                        (Pe) =>
                                          `${Pe.name} (${Pe.types.join(" / ") || "empty"})`,
                                      )
                                      .join(", "),
                                  ],
                                }),
                              ye.error &&
                                i.jsx("p", {
                                  className: "recipe-error",
                                  children: ye.error,
                                }),
                              ye.controls &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(M, { children: "Control totals:" }),
                                    " ",
                                    ye.controls
                                      .map(
                                        (Pe) =>
                                          `${Pe.field}: ${Pe.sum} across ${Pe.numericCount} numeric values; ${Pe.blankCount} blanks`,
                                      )
                                      .join(" · "),
                                  ],
                                }),
                              ((Te = ye.details) == null
                                ? void 0
                                : Te.matchPairs) !== void 0 &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(M, { children: "Merge: " }),
                                    String(ye.details.matchPairs),
                                    " ",
                                    i.jsx(M, { children: "pairs · matched" }),
                                    " ",
                                    String(ye.details.matchedLeft),
                                    " ",
                                    i.jsx(M, { children: "left /" }),
                                    " ",
                                    String(ye.details.matchedRight),
                                    " ",
                                    i.jsx(M, { children: "right · unmatched" }),
                                    " ",
                                    ye.details.unmatchedLeft.length,
                                    " ",
                                    i.jsx(M, { children: "left /" }),
                                    " ",
                                    ye.details.unmatchedRight.length,
                                    " ",
                                    i.jsx(M, { children: "right · " }),
                                    String(ye.details.manyToMany),
                                    " ",
                                    i.jsx(M, {
                                      children: "many-to-many groups ·",
                                    }),
                                    " ",
                                    String(ye.details.extraPairs),
                                    " ",
                                    i.jsx(M, {
                                      children:
                                        "extra pairs. The saved review contains every match explanation and exception.",
                                    }),
                                  ],
                                }),
                              ye.outputHash &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(M, {
                                      children: "Output fingerprint ",
                                    }),
                                    i.jsx("code", { children: ye.outputHash }),
                                  ],
                                }),
                            ],
                          },
                          ye.id,
                        );
                      }),
                    }),
                  C === "sources" &&
                    i.jsxs("div", {
                      "aria-label": r("Recipe sources and totals"),
                      lang: t("Recipe sources and totals"),
                      "data-ui-label": "Recipe sources and totals",
                      children: [
                        K.sources.map((ye, Te) =>
                          i.jsxs(
                            "details",
                            {
                              open: true,
                              children: [
                                i.jsxs("summary", {
                                  children: [
                                    ye.sheetName,
                                    "!",
                                    ye.resolvedRange,
                                    " · ",
                                    ye.rows,
                                    " ",
                                    i.jsx(M, { children: "source records" }),
                                  ],
                                }),
                                i.jsx("p", {
                                  children: ye.schema
                                    .map(
                                      (Pe) =>
                                        `${Pe.name}: ${Pe.types.join(" / ") || "empty"}`,
                                    )
                                    .join(", "),
                                }),
                                i.jsx("p", {
                                  children: ye.controls
                                    .map(
                                      (Pe) =>
                                        `${Pe.field}: ${Pe.sum} (${Pe.numericCount} numbers, ${Pe.blankCount} blanks)`,
                                    )
                                    .join(" · "),
                                }),
                                i.jsx("code", { children: ye.hash }),
                              ],
                            },
                            Te,
                          ),
                        ),
                        K.status === "complete" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("h4", {
                                children: i.jsx(M, {
                                  children: "Final controls",
                                }),
                              }),
                              K.controls.map((ye) =>
                                i.jsxs(
                                  "p",
                                  {
                                    children: [
                                      ye.field,
                                      i.jsx(M, { children: ": sum " }),
                                      ye.sum,
                                      " · ",
                                      ye.numericCount,
                                      " ",
                                      i.jsx(M, {
                                        children: "numeric values · ",
                                      }),
                                      ye.blankCount,
                                      " ",
                                      i.jsx(M, { children: "blanks" }),
                                    ],
                                  },
                                  ye.field,
                                ),
                              ),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(M, { children: "Output fingerprint " }),
                                  i.jsx("code", { children: K.outputHash }),
                                ],
                              }),
                            ],
                          }),
                      ],
                    }),
                  i.jsx("button", {
                    onClick: Ee,
                    children: i.jsx(M, { children: "Save recipe review" }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(M, {
                      children:
                        "Output is a separate static snapshot. Publishing preserves prior outputs and records one Undo action. Saved steps, fingerprints and lineage are retained in native JSON; XLSX preserves the supported values and notes.",
                    }),
                  }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: e.readOnly
                ? "Read-only host"
                : J
                  ? "Save a version before publishing"
                  : "Review the current version before publishing",
            }),
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(M, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "recipe-primary",
              disabled:
                e.readOnly ||
                !W ||
                J ||
                (K == null ? void 0 : K.status) !== "complete",
              onClick: we,
              children: i.jsx(M, { children: "Publish recipe output" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const Bk = (e) => ({
    column: e,
    type: "text",
    operator: "contains",
    value: "",
  }),
  qce = {
    equals: "Equals",
    notEquals: "Does not equal",
    contains: "Contains",
    notContains: "Does not contain",
    startsWith: "Begins with",
    endsWith: "Ends with",
    wildcard: "Matches wildcard (* ? ~)",
    gt: "Greater than / after",
    gte: "At least / on or after",
    lt: "Less than / before",
    lte: "At most / on or before",
    blank: "Is blank",
    notBlank: "Is not blank",
  };

export { Bk, Hce, qce };
