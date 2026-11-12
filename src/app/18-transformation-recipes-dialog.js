import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  transformationRecipeLimits,
  fingerprintCanonicalValue,
  recipeStoreFormatId,
} from "./01-schemas-date-functions-and-validation.js";
import { formatOccupiedRangeAddress } from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  RecipeSourceSelector,
  runTransformationRecipe,
  saveTransformationRecipe,
  deleteTransformationRecipe,
  recipeStepKindLabels,
  publishRecipeOutput,
  createRecipeStep,
  formatCellValueForDisplay,
  normalizeRecipeRunRequest,
  RecipeStepEditor,
} from "./17-recipe-step-editors.js";

function TransformationRecipesDialog(e) {
  const { text, language } = useWorkbookMessages(),
    n = ne.useRef(null),
    o = ne.useRef(null),
    a = ne.useRef(e.onClose);
  a.current = e.onClose;
  const [s, d] = ne.useState(() => ({
      engine: recipeStoreFormatId,
      source: {
        sheetId: e.sheetId,
        range: e.range.includes(":")
          ? e.range
          : formatOccupiedRangeAddress(
              e.workbook.sheets.find((ye) => ye.id === e.sheetId),
            ),
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
    X$1 = Q?.versions.find((ye) => ye.version === f),
    J = !X$1 || X$1.definitionHash !== fingerprintCanonicalValue(s),
    q = JSON.stringify([e.workbook, e.sheetId, e.range, s, c, f, k, $]),
    W = j?.stamp === q,
    K = W ? j.plan.run : undefined;
  function te() {
    (z(null), D(null), R(0));
  }
  function ae(ye) {
    (d(ye), te(), _(""));
  }
  function oe(ye) {
    try {
      const Te = runTransformationRecipe(
        e.workbook,
        { ...s, steps: s.steps.slice(0, ye) },
        $,
      );
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
        const ye = runTransformationRecipe(
          e.workbook,
          { engine: recipeStoreFormatId, source: he.source, steps: [] },
          $,
        );
        return ye.status === "complete" ? ye.output.fields : [];
      } catch {
        return [];
      }
    }, [e.workbook, he, JSON.stringify($)]);
  (ne.useEffect(() => {
    const ye = document.activeElement;
    return (
      n.current?.querySelector("input,select,button")?.focus(),
      () => {
        ye?.isConnected && ye.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      var ye;
      W &&
        ((ye = o.current) == null ? undefined : ye.scrollIntoView)?.call(ye, {
          block: "start",
        });
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
        Te = ye?.versions.find((Pe) => Pe.version === g);
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
      const ye = saveTransformationRecipe(e.workbook, u, s, c || undefined, $),
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
      const ye = normalizeRecipeRunRequest(
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
      const ye = publishRecipeOutput(e.workbook, j.plan, $);
      (te(),
        e.onWorkbookChange(ye, `Published recipe ${u} version ${f}`),
        e.onClose());
    });
  }
  function Ee() {
    if (!W || !j) return;
    const { snapshot, ...Te } = j.plan,
      Pe = URL.createObjectURL(
        new Blob([JSON.stringify(Te, null, 2)], { type: "application/json" }),
      ),
      qe = document.createElement("a");
    ((qe.href = Pe),
      (qe.download = "Recipe-execution-review.json"),
      qe.click(),
      setTimeout(() => URL.revokeObjectURL(Pe), 5e3));
  }
  const Ae = e.workbook.sheets.filter(
    (ye) => ye.transformationRun?.recipeId === c,
  );
  return i.jsx("div", {
    className: "recipe-backdrop",
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": "recipe-title",
      className: "recipe-dialog",
      ref: n,
      onKeyDown: (ye) => {
        if (
          (ye.key === "Escape" && (ye.preventDefault(), a.current()),
          ye.key === "Tab")
        ) {
          const Pe = Array.from(
              n.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary,[tabindex="0"]',
              ) ?? [],
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
                  children: i.jsx(LocalizedText, {
                    children: "Transformation recipes",
                  }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Save ordered steps. Review each run before publishing a separate result.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close transformation recipes"),
              onClick: e.onClose,
              lang: language("Close transformation recipes"),
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
              "aria-label": text("Saved transformation recipes"),
              lang: language("Saved transformation recipes"),
              "data-ui-label": "Saved transformation recipes",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Saved recipe" }),
                    i.jsxs("select", {
                      "aria-label": text("Saved recipe"),
                      value: m,
                      onChange: (ye) => {
                        (v(ye.target.value),
                          b(
                            G.find((Pe) => Pe.id === ye.target.value)?.versions
                              .length ?? 1,
                          ));
                      },
                      lang: language("Saved recipe"),
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
                    i.jsx(LocalizedText, { children: "Version" }),
                    i.jsx("select", {
                      "aria-label": text("Saved recipe version"),
                      value: g,
                      onChange: (ye) => b(Number(ye.target.value)),
                      lang: language("Saved recipe version"),
                      "data-ui-label": "Saved recipe version",
                      children: (H?.versions ?? []).map((ye) =>
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
                  children: i.jsx(LocalizedText, { children: "Load version" }),
                }),
                i.jsx("button", {
                  onClick: () => {
                    (h(""),
                      p(0),
                      l("Transformation recipe"),
                      d({
                        engine: recipeStoreFormatId,
                        source: s.source,
                        steps: [],
                      }),
                      w(-1),
                      te(),
                      _("New recipe draft."));
                  },
                  children: i.jsx(LocalizedText, { children: "New recipe" }),
                }),
                i.jsx("button", {
                  disabled: !H || e.readOnly,
                  onClick: () =>
                    le(() => {
                      const ye = deleteTransformationRecipe(e.workbook, m, $);
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
                  children: i.jsx(LocalizedText, {
                    children: "Delete saved recipe",
                  }),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "recipe-line",
              children: [
                i.jsxs("label", {
                  children: [
                    i.jsx(LocalizedText, { children: "Recipe name" }),
                    i.jsx("input", {
                      "aria-label": text("Recipe name"),
                      value: u,
                      onChange: (ye) => l(ye.target.value),
                      maxLength: 128,
                      lang: language("Recipe name"),
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
            i.jsx(RecipeSourceSelector, {
              label: text("Primary source"),
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
                  i.jsx(LocalizedText, {
                    children:
                      "Include hidden source cells and formula dependencies",
                  }),
                ],
              }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Sources include every record in the declared ranges, regardless of current filter views. Tables use their current headers and data body. Source cells stay unchanged.",
              }),
            }),
            i.jsxs("div", {
              className: "recipe-layout",
              children: [
                i.jsxs("aside", {
                  "aria-label": text("Ordered recipe steps"),
                  lang: language("Ordered recipe steps"),
                  "data-ui-label": "Ordered recipe steps",
                  children: [
                    i.jsx("h3", {
                      children: i.jsx(LocalizedText, {
                        children: "Applied steps",
                      }),
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
                                children: [
                                  Te + 1,
                                  ". ",
                                  recipeStepKindLabels[ye.kind],
                                ],
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
                        i.jsx(LocalizedText, {
                          children: "Add transformation",
                        }),
                        i.jsx("select", {
                          "aria-label": text("New transformation type"),
                          value: E,
                          onChange: (ye) => A(ye.target.value),
                          lang: language("New transformation type"),
                          "data-ui-label": "New transformation type",
                          children: Object.entries(recipeStepKindLabels).map(
                            ([ye, Te]) =>
                              i.jsx("option", { value: ye, children: Te }, ye),
                          ),
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      disabled:
                        s.steps.length >= transformationRecipeLimits.steps,
                      onClick: () => {
                        const ye = oe(s.steps.length).fields;
                        (ae({
                          ...s,
                          steps: [
                            ...s.steps,
                            createRecipeStep(E, ye, s.source),
                          ],
                        }),
                          w(s.steps.length));
                      },
                      children: i.jsx(LocalizedText, { children: "Add step" }),
                    }),
                    i.jsxs("p", {
                      children: [
                        s.steps.length,
                        " / ",
                        transformationRecipeLimits.steps,
                        " ",
                        i.jsx(LocalizedText, { children: "steps" }),
                      ],
                    }),
                  ],
                }),
                i.jsx("section", {
                  className: "recipe-editor",
                  "aria-label": text("Recipe step editor"),
                  lang: language("Recipe step editor"),
                  "data-ui-label": "Recipe step editor",
                  children: he
                    ? i.jsxs(i.Fragment, {
                        children: [
                          i.jsxs("h3", {
                            children: [
                              y + 1,
                              ". ",
                              recipeStepKindLabels[he.kind],
                            ],
                          }),
                          de.error &&
                            i.jsxs("p", {
                              role: "status",
                              children: [
                                i.jsx(LocalizedText, {
                                  children:
                                    "Earlier step or source needs attention: ",
                                }),
                                de.error,
                              ],
                            }),
                          i.jsx(RecipeStepEditor, {
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
                        children: i.jsx(LocalizedText, {
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
                    i.jsx(LocalizedText, { children: "New output worksheet" }),
                    i.jsx("input", {
                      "aria-label": text("Recipe output worksheet"),
                      value: k,
                      onChange: (ye) => {
                        (F(ye.target.value), te());
                      },
                      maxLength: 31,
                      lang: language("Recipe output worksheet"),
                      "data-ui-label": "Recipe output worksheet",
                    }),
                  ],
                }),
                i.jsx("button", {
                  className: "recipe-primary",
                  onClick: se,
                  children: i.jsx(LocalizedText, {
                    children: "Preview recipe",
                  }),
                }),
              ],
            }),
            B && i.jsx("p", { role: "status", children: B }),
            V?.stamp === q &&
              i.jsx("p", {
                role: "alert",
                className: "recipe-error",
                children: V.text,
              }),
            j &&
              !W &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children:
                    "Workbook, selection or permissions changed. Preview again; previous results are withheld.",
                }),
              }),
            Ae.length > 0 &&
              i.jsxs("details", {
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Previous saved outputs (",
                      }),
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
                          i.jsx(LocalizedText, { children: "· version " }),
                          ye.transformationRun.version,
                          " ·",
                          " ",
                          ye.transformationRun.outputRows,
                          " ",
                          i.jsx(LocalizedText, { children: "rows ·" }),
                          " ",
                          ye.transformationRun.at,
                          i.jsx(LocalizedText, {
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
                "aria-label": text("Recipe execution preview"),
                ref: o,
                lang: language("Recipe execution preview"),
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
                        i.jsx(LocalizedText, {
                          children:
                            "No output will be published. Downstream steps were skipped.",
                        }),
                      ],
                    }),
                  J &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Draft preview. Save a version and preview it before publishing.",
                      }),
                    }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, {
                        children: "Definition fingerprint ",
                      }),
                      i.jsx("code", { children: K.definitionHash }),
                    ],
                  }),
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Source fingerprint " }),
                      i.jsx("code", { children: K.sourceHash }),
                    ],
                  }),
                  i.jsx("nav", {
                    "aria-label": text("Recipe review views"),
                    lang: language("Recipe review views"),
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
                          "aria-label": text("Complete recipe output"),
                          lang: language("Complete recipe output"),
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
                                        children: i.jsx(LocalizedText, {
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
                                                {
                                                  children:
                                                    formatCellValueForDisplay(
                                                      Pe,
                                                    ),
                                                },
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
                                                      i.jsx(LocalizedText, {
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
                                children: i.jsx(LocalizedText, {
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
                              children: i.jsx(LocalizedText, {
                                children: "Previous recipe records",
                              }),
                            }),
                            i.jsxs("span", {
                              children: [
                                i.jsx(LocalizedText, { children: "Page " }),
                                N + 1,
                                " ",
                                i.jsx(LocalizedText, { children: "of" }),
                                " ",
                                Math.max(
                                  1,
                                  Math.ceil(K.output.rows.length / 25),
                                ),
                                " ·",
                                " ",
                                K.output.rows.length,
                                " ",
                                i.jsx(LocalizedText, { children: "records" }),
                              ],
                            }),
                            i.jsx("button", {
                              disabled: (N + 1) * 25 >= K.output.rows.length,
                              onClick: () => R((ye) => ye + 1),
                              children: i.jsx(LocalizedText, {
                                children: "Next recipe records",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                  C === "steps" &&
                    i.jsx("div", {
                      "aria-label": text("Recipe step trace"),
                      lang: language("Recipe step trace"),
                      "data-ui-label": "Recipe step trace",
                      children: K.steps.map((ye) =>
                        i.jsxs(
                          "details",
                          {
                            open: ye.status === "failed",
                            children: [
                              i.jsxs("summary", {
                                children: [
                                  ye.index,
                                  ".",
                                  " ",
                                  recipeStepKindLabels[ye.kind] ?? ye.kind,
                                  " ",
                                  "· ",
                                  ye.status,
                                  " · ",
                                  ye.inputRows,
                                  " → ",
                                  ye.outputRows ?? "—",
                                  " ",
                                  i.jsx(LocalizedText, { children: "rows" }),
                                ],
                              }),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Input schema:",
                                  }),
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
                                    i.jsx(LocalizedText, {
                                      children: "Output schema:",
                                    }),
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
                                    i.jsx(LocalizedText, {
                                      children: "Control totals:",
                                    }),
                                    " ",
                                    ye.controls
                                      .map(
                                        (Pe) =>
                                          `${Pe.field}: ${Pe.sum} across ${Pe.numericCount} numeric values; ${Pe.blankCount} blanks`,
                                      )
                                      .join(" · "),
                                  ],
                                }),
                              ye.details?.matchPairs !== undefined &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Merge: ",
                                    }),
                                    String(ye.details.matchPairs),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "pairs · matched",
                                    }),
                                    " ",
                                    String(ye.details.matchedLeft),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "left /",
                                    }),
                                    " ",
                                    String(ye.details.matchedRight),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "right · unmatched",
                                    }),
                                    " ",
                                    ye.details.unmatchedLeft.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "left /",
                                    }),
                                    " ",
                                    ye.details.unmatchedRight.length,
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "right · ",
                                    }),
                                    String(ye.details.manyToMany),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children: "many-to-many groups ·",
                                    }),
                                    " ",
                                    String(ye.details.extraPairs),
                                    " ",
                                    i.jsx(LocalizedText, {
                                      children:
                                        "extra pairs. The saved review contains every match explanation and exception.",
                                    }),
                                  ],
                                }),
                              ye.outputHash &&
                                i.jsxs("p", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Output fingerprint ",
                                    }),
                                    i.jsx("code", { children: ye.outputHash }),
                                  ],
                                }),
                            ],
                          },
                          ye.id,
                        ),
                      ),
                    }),
                  C === "sources" &&
                    i.jsxs("div", {
                      "aria-label": text("Recipe sources and totals"),
                      lang: language("Recipe sources and totals"),
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
                                    i.jsx(LocalizedText, {
                                      children: "source records",
                                    }),
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
                                children: i.jsx(LocalizedText, {
                                  children: "Final controls",
                                }),
                              }),
                              K.controls.map((ye) =>
                                i.jsxs(
                                  "p",
                                  {
                                    children: [
                                      ye.field,
                                      i.jsx(LocalizedText, {
                                        children: ": sum ",
                                      }),
                                      ye.sum,
                                      " · ",
                                      ye.numericCount,
                                      " ",
                                      i.jsx(LocalizedText, {
                                        children: "numeric values · ",
                                      }),
                                      ye.blankCount,
                                      " ",
                                      i.jsx(LocalizedText, {
                                        children: "blanks",
                                      }),
                                    ],
                                  },
                                  ye.field,
                                ),
                              ),
                              i.jsxs("p", {
                                children: [
                                  i.jsx(LocalizedText, {
                                    children: "Output fingerprint ",
                                  }),
                                  i.jsx("code", { children: K.outputHash }),
                                ],
                              }),
                            ],
                          }),
                      ],
                    }),
                  i.jsx("button", {
                    onClick: Ee,
                    children: i.jsx(LocalizedText, {
                      children: "Save recipe review",
                    }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(LocalizedText, {
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
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "recipe-primary",
              disabled: e.readOnly || !W || J || K?.status !== "complete",
              onClick: we,
              children: i.jsx(LocalizedText, {
                children: "Publish recipe output",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const createDefaultFilterCriterion = (e) => ({
    column: e,
    type: "text",
    operator: "contains",
    value: "",
  }),
  filterOperatorLabels = {
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

export {
  createDefaultFilterCriterion,
  TransformationRecipesDialog,
  filterOperatorLabels,
};
