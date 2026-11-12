import i from "react/jsx-runtime";
import ne from "react";
import { ts } from "./part-01.js";
import { FO, Kie, Qie, ct, eh, eoe, f0, toe, ut } from "./part-03.js";
import { Dr, M, ci } from "./part-04.js";
import { Bk, qce } from "./part-18.js";

function Xce(e) {
  var R, L, $, U, G, H, Q, X, J, q, W, K, te, ae;
  const { text: r, language: t } = Dr(),
    n = ne.useId(),
    o = ne.useRef(null),
    a = ne.useRef(null),
    [s, d] = ne.useState(() =>
      e.definition
        ? structuredClone(e.definition)
        : { range: Kie(e.workbook, e.sheetId, e.range), columns: [] },
    ),
    [u, l] = ne.useState(e.column ?? ct(s.range).firstCol),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState(null),
    [m, v] = ne.useState(""),
    [g, b] = ne.useState(""),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(""),
    k =
      (R = e.workbook.savedFilterViews) == null
        ? void 0
        : R.find((oe) => oe.id === g),
    F =
      !!e.onWorkbookChange &&
      !e.readOnly &&
      !((L = e.workbook.structureProtection) != null && L.enabled) &&
      !(
        (U =
          ($ = e.workbook.sheets.find((oe) => oe.id === e.sheetId)) == null
            ? void 0
            : $.protection) != null && U.enabled
      ),
    I =
      !!e.onWorkbookChange &&
      !e.readOnly &&
      !((G = e.workbook.structureProtection) != null && G.enabled) &&
      !!k &&
      !(
        (Q =
          (H = e.workbook.sheets.find((oe) => oe.id === k.sheetId)) == null
            ? void 0
            : H.protection) != null && Q.enabled
      ),
    O = ne.useMemo(() => {
      try {
        const oe = ct(s.range);
        return Array.from(
          { length: Math.min(64, oe.lastCol - oe.firstCol + 1) },
          (de, he) => {
            var le, ce;
            const Se = oe.firstCol + he;
            return {
              column: Se,
              label: `${ut(Se)} · ${String(((ce = (le = e.workbook.sheets.find((ge) => ge.id === e.sheetId)) == null ? void 0 : le.cells[`${ut(Se)}${oe.firstRow + 1}`]) == null ? void 0 : ce.raw) ?? "Column")}`,
            };
          },
        );
      } catch {
        return [];
      }
    }, [s.range, e.workbook, e.sheetId]),
    j = ne.useMemo(() => {
      try {
        return { result: FO(e.workbook, e.sheetId, s, u, c), error: "" };
      } catch (oe) {
        return {
          result: void 0,
          error: oe instanceof Error ? oe.message : "Cannot inspect values.",
        };
      }
    }, [e.workbook, e.sheetId, s, u, c]),
    z = s.columns.find((oe) => oe.column === u) ?? { column: u },
    V =
      (f == null ? void 0 : f.book) === e.workbook &&
      f.definition === JSON.stringify(s)
        ? f.result
        : null;
  function D() {
    if (!V)
      throw Error(
        "Preview the current population before saving or applying a view.",
      );
    const oe = f0(e.workbook, e.sheetId, s);
    if (JSON.stringify(oe.context) !== JSON.stringify(V.context))
      throw Error(
        "Population membership changed. Preview again before saving or applying.",
      );
    return oe;
  }
  function B(oe = false) {
    try {
      if (!F || (oe && (!k || k.sheetId !== e.sheetId)))
        throw Error("This saved view cannot be changed here.");
      const de = D(),
        he = Qie(
          e.workbook,
          e.sheetId,
          de.context.definition,
          y,
          oe ? g : void 0,
        );
      (e.onWorkbookChange(
        he,
        oe ? "Replace saved filter criteria" : "Save named filter view",
      ),
        b(he.savedFilterViews.at(-1).id),
        p(null),
        v(""),
        A(
          "Saved in the native workbook. Preview again to apply; saved counts are historical and source cells are unchanged.",
        ));
    } catch (de) {
      v(de instanceof Error ? de.message : "Cannot save this view.");
    }
  }
  ne.useEffect(() => {
    var oe, de;
    f &&
      ((de = (oe = a.current) == null ? void 0 : oe.scrollIntoView) == null ||
        de.call(oe, { block: "nearest" }));
  }, [f]);
  function _(oe) {
    (d((de) => ({
      ...de,
      columns: [...de.columns.filter((he) => he.column !== u), oe],
    })),
      v(""));
  }
  function C(oe) {
    const de = FO(e.workbook, e.sheetId, s, u);
    if (z.values === void 0 && de.truncated) {
      v(
        "Choose Select no values first, then select specific values, or use a custom condition.",
      );
      return;
    }
    const he = z.values ?? de.values.map((ce) => ce.value),
      Se = eh(oe, s.caseSensitive),
      le = he.some((ce) => eh(ce, s.caseSensitive) === Se);
    _({
      ...z,
      values: le
        ? he.filter((ce) => eh(ce, s.caseSensitive) !== Se)
        : [...he, oe],
    });
  }
  function T(oe, de, he) {
    d((Se) => {
      var le;
      return {
        ...Se,
        advanced:
          (le = Se.advanced) == null
            ? void 0
            : le.map((ce, ge) =>
                ge === oe ? ce.map((se, we) => (we === de ? he : se)) : ce,
              ),
      };
    });
  }
  function N(oe, de, he, Se = false) {
    return i.jsxs("div", {
      className: "lg-filter-condition",
      children: [
        Se &&
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Column" }),
              i.jsx("select", {
                "aria-label": `${he} column`,
                value: oe.column,
                onChange: (le) =>
                  de({ ...oe, column: Number(le.target.value) }),
                children: O.map((le) =>
                  i.jsx(
                    "option",
                    { value: le.column, children: le.label },
                    le.column,
                  ),
                ),
              }),
            ],
          }),
        i.jsxs("label", {
          children: [
            i.jsx(M, { children: "Type" }),
            i.jsx("select", {
              "aria-label": `${he} type`,
              value: oe.type,
              onChange: (le) =>
                de({
                  ...oe,
                  type: le.target.value,
                  operator: "equals",
                  value: le.target.value === "number" ? 0 : "",
                }),
              children: ["value", "text", "number", "date"].map((le) =>
                i.jsx(
                  "option",
                  {
                    value: le,
                    children:
                      le === "value"
                        ? "Typed value"
                        : le[0].toUpperCase() + le.slice(1),
                  },
                  le,
                ),
              ),
            }),
          ],
        }),
        i.jsxs("label", {
          children: [
            i.jsx(M, { children: "Comparison" }),
            i.jsx("select", {
              "aria-label": `${he} comparison`,
              value: oe.operator,
              onChange: (le) => de({ ...oe, operator: le.target.value }),
              children: Object.entries(qce)
                .filter(
                  ([le]) =>
                    oe.type === "text" ||
                    [
                      "equals",
                      "notEquals",
                      "blank",
                      "notBlank",
                      ...(oe.type === "value"
                        ? []
                        : ["gt", "gte", "lt", "lte"]),
                    ].includes(le),
                )
                .map(([le, ce]) =>
                  i.jsx("option", { value: le, children: ce }, le),
                ),
            }),
          ],
        }),
        !["blank", "notBlank"].includes(oe.operator) &&
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Value" }),
              oe.type === "number"
                ? i.jsx(ci, {
                    "aria-label": `${he} value`,
                    value: typeof oe.value == "number" ? oe.value : "",
                    onValueChange: (le) => de({ ...oe, value: Number(le) }),
                  })
                : i.jsx("input", {
                    "aria-label": `${he} value`,
                    value: String(oe.value ?? ""),
                    type: "text",
                    placeholder:
                      oe.type === "date" ? "2026-09-17" : "Match value",
                    onChange: (le) => de({ ...oe, value: le.target.value }),
                  }),
            ],
          }),
      ],
    });
  }
  return (
    ne.useEffect(() => {
      var de;
      const oe = document.activeElement;
      return (
        (de = o.current) == null || de.focus(),
        () => {
          oe != null && oe.isConnected && oe.focus();
        }
      );
    }, []),
    i.jsx("div", {
      className: "lg-filter-overlay",
      children: i.jsxs("section", {
        className: "lg-filter-dialog",
        role: "dialog",
        "aria-modal": "true",
        "aria-labelledby": n,
        ref: o,
        tabIndex: -1,
        onKeyDown: (oe) => {
          var de, he, Se;
          if (
            (oe.stopPropagation(),
            oe.key === "Escape" && (oe.preventDefault(), e.onClose()),
            oe.key === "Tab")
          ) {
            const le = [
              ...(((de = o.current) == null
                ? void 0
                : de.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
                  )) ?? []),
            ];
            oe.shiftKey && [le[0], o.current].includes(document.activeElement)
              ? (oe.preventDefault(), (he = le.at(-1)) == null || he.focus())
              : !oe.shiftKey &&
                document.activeElement === le.at(-1) &&
                (oe.preventDefault(), (Se = le[0]) == null || Se.focus());
          }
        },
        children: [
          i.jsxs("header", {
            children: [
              i.jsxs("div", {
                children: [
                  i.jsx("h2", {
                    id: n,
                    children: i.jsx(M, { children: "Filter population" }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(M, {
                      children:
                        "Keep complete records and their original worksheet addresses.",
                    }),
                  }),
                ],
              }),
              i.jsx("button", {
                "aria-label": r("Close population filter"),
                onClick: e.onClose,
                lang: t("Close population filter"),
                "data-ui-label": "Close population filter",
                children: "×",
              }),
            ],
          }),
          i.jsxs("div", {
            className: "lg-filter-scroll",
            children: [
              i.jsx("p", {
                className: "lg-filter-disclosure",
                children: i.jsx(M, {
                  children:
                    "View only · first range row is the header. Edits pause this view until you review and reapply it. Native JSON, XLSX, printing, other analysis tools and AI selections use full source data. CSV/TSV and Audit Checks offer an explicit full or visible population.",
                }),
              }),
              i.jsxs("details", {
                className: "lg-saved-filter-views",
                children: [
                  i.jsxs("summary", {
                    children: [
                      i.jsx(M, { children: "Saved filter views (" }),
                      ((X = e.workbook.savedFilterViews) == null
                        ? void 0
                        : X.length) ?? 0,
                      ")",
                    ],
                  }),
                  i.jsx("p", {
                    children: i.jsx(M, {
                      children:
                        "Save named criteria in the native workbook and browser save. Loading never hides rows automatically: check the range, headers and date rules, then preview and apply. XLSX does not retain saved views.",
                    }),
                  }),
                  i.jsxs("div", {
                    className: "lg-filter-controls",
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Saved view" }),
                          i.jsxs("select", {
                            "aria-label": r("Saved filter view"),
                            value: g,
                            onChange: (oe) => {
                              var de, he;
                              (b(oe.target.value),
                                w(
                                  ((he =
                                    (de = e.workbook.savedFilterViews) == null
                                      ? void 0
                                      : de.find(
                                          (Se) => Se.id === oe.target.value,
                                        )) == null
                                    ? void 0
                                    : he.name) ?? "",
                                ),
                                A(""),
                                v(""));
                            },
                            lang: t("Saved filter view"),
                            "data-ui-label": "Saved filter view",
                            children: [
                              i.jsx("option", {
                                value: "",
                                children: "Choose a saved view",
                              }),
                              (J = e.workbook.savedFilterViews) == null
                                ? void 0
                                : J.map((oe) => {
                                    var de;
                                    return i.jsxs(
                                      "option",
                                      {
                                        value: oe.id,
                                        children: [
                                          oe.name,
                                          " ·",
                                          " ",
                                          ((de = e.workbook.sheets.find(
                                            (he) => he.id === oe.sheetId,
                                          )) == null
                                            ? void 0
                                            : de.name) ??
                                            `${oe.sheetName} (sheet missing)`,
                                        ],
                                      },
                                      oe.id,
                                    );
                                  }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "View name" }),
                          i.jsx("input", {
                            "aria-label": r("Filter view name"),
                            maxLength: 80,
                            value: y,
                            onChange: (oe) => w(oe.target.value),
                            lang: t("Filter view name"),
                            "data-ui-label": "Filter view name",
                          }),
                        ],
                      }),
                    ],
                  }),
                  k &&
                    i.jsxs("p", {
                      children: [
                        i.jsx(M, { children: "Saved range " }),
                        k.definition.range,
                        " ·",
                        " ",
                        k.population.visibleRows,
                        " ",
                        i.jsx(M, { children: "visible /" }),
                        " ",
                        k.population.totalRows,
                        " ",
                        i.jsx(M, { children: "total at source revision" }),
                        " ",
                        k.sourceRevision,
                        i.jsx(M, { children: ". Current revision" }),
                        " ",
                        e.workbook.revision,
                        i.jsx(M, {
                          children:
                            ". These are historical counts, not current results.",
                        }),
                        k.sourceWorkbookId !== e.workbook.id &&
                          " The preview was saved in the original workbook; this is an independent copy.",
                        k.sheetId !== e.sheetId &&
                          " Open the source worksheet to load these criteria. A missing worksheet must be restored first.",
                        JSON.stringify(k.datePolicy) !==
                          JSON.stringify(ts(e.workbook.datePolicy)) &&
                          " Date policy changed. Check all date criteria against the current workbook policy.",
                      ],
                    }),
                  i.jsxs("div", {
                    className: "lg-filter-actions",
                    children: [
                      i.jsx("button", {
                        disabled: !k || k.sheetId !== e.sheetId,
                        onClick: () => {
                          !k ||
                            k.sheetId !== e.sheetId ||
                            (d(structuredClone(k.definition)),
                            l(ct(k.definition.range).firstCol),
                            h(""),
                            p(null),
                            v(""),
                            A(
                              "Criteria loaded for review. Saved coordinates do not move with inserted, deleted or sorted records. Check the range and headers, then preview current results.",
                            ));
                        },
                        children: i.jsx(M, {
                          children: "Load criteria for review",
                        }),
                      }),
                      i.jsx("button", {
                        disabled: !F || !V || !y.trim(),
                        onClick: () => B(),
                        children: i.jsx(M, { children: "Save as new view" }),
                      }),
                      i.jsx("button", {
                        disabled:
                          !F ||
                          !V ||
                          !k ||
                          k.sheetId !== e.sheetId ||
                          !y.trim(),
                        onClick: () => B(true),
                        children: i.jsx(M, {
                          children: "Replace saved criteria",
                        }),
                      }),
                      i.jsx("button", {
                        disabled: !I || !y.trim(),
                        onClick: () => {
                          try {
                            if (!I) return;
                            (e.onWorkbookChange(
                              eoe(e.workbook, g, y),
                              "Rename saved filter view",
                            ),
                              A(
                                "Saved view renamed. Its criteria and historical counts are unchanged.",
                              ),
                              v(""));
                          } catch (oe) {
                            v(
                              oe instanceof Error
                                ? oe.message
                                : "Cannot rename this view.",
                            );
                          }
                        },
                        children: i.jsx(M, { children: "Rename view" }),
                      }),
                      i.jsx("button", {
                        disabled: !I,
                        onClick: () => {
                          try {
                            if (!I) return;
                            (e.onWorkbookChange(
                              toe(e.workbook, g),
                              "Delete saved filter view",
                            ),
                              b(""),
                              w(""),
                              A(
                                "Saved view deleted. Source cells and criteria in this editor are retained. Undo can restore the saved view.",
                              ),
                              v(""));
                          } catch (oe) {
                            v(
                              oe instanceof Error
                                ? oe.message
                                : "Cannot delete this view.",
                            );
                          }
                        },
                        children: i.jsx(M, { children: "Delete view" }),
                      }),
                    ],
                  }),
                  !F &&
                    i.jsx("p", {
                      children: i.jsx(M, {
                        children:
                          "Saving is unavailable in this read-only, protected or view-only host. Loading, previewing and applying session views remain available.",
                      }),
                    }),
                  E && i.jsx("p", { role: "status", children: E }),
                ],
              }),
              i.jsxs("div", {
                className: "lg-filter-controls",
                children: [
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Population range" }),
                      i.jsx("input", {
                        "aria-label": r("Population range"),
                        value: s.range,
                        onChange: (oe) =>
                          d((de) => ({
                            ...de,
                            range: oe.target.value,
                            columns: [],
                            advanced: [],
                            unique: false,
                          })),
                        lang: t("Population range"),
                        "data-ui-label": "Population range",
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Column" }),
                      i.jsx("select", {
                        "aria-label": r("Filter column"),
                        value: u,
                        onChange: (oe) => {
                          (l(Number(oe.target.value)), h(""));
                        },
                        lang: t("Filter column"),
                        "data-ui-label": "Filter column",
                        children: O.map((oe) =>
                          i.jsx(
                            "option",
                            { value: oe.column, children: oe.label },
                            oe.column,
                          ),
                        ),
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: !!s.caseSensitive,
                        onChange: (oe) =>
                          d((de) => ({
                            ...de,
                            caseSensitive: oe.target.checked,
                          })),
                      }),
                      i.jsx(M, { children: "Case sensitive" }),
                    ],
                  }),
                ],
              }),
              i.jsxs("div", {
                className: "lg-filter-columns",
                children: [
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(M, { children: "Values / slicer" }),
                      }),
                      i.jsx("input", {
                        "aria-label": r("Search filter values"),
                        placeholder: r("Search available values"),
                        value: c,
                        onChange: (oe) => h(oe.target.value),
                        lang: t("Search filter values"),
                        "data-ui-label": "Search filter values",
                      }),
                      i.jsxs("div", {
                        className: "lg-filter-actions",
                        children: [
                          i.jsx("button", {
                            onClick: () => _({ ...z, values: void 0 }),
                            children: i.jsx(M, {
                              children: "Select all values",
                            }),
                          }),
                          i.jsx("button", {
                            onClick: () => _({ ...z, values: [] }),
                            children: i.jsx(M, {
                              children: "Select no values",
                            }),
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "Other column and advanced criteria limit available values. Search narrows this list only. Manual hiding does not change the value list.",
                        }),
                      }),
                      j.error &&
                        i.jsx("p", { role: "alert", children: j.error }),
                      ((q = j.result) == null ? void 0 : q.truncated) &&
                        i.jsxs("p", {
                          role: "status",
                          children: [
                            i.jsx(M, { children: "Showing 1,000 of " }),
                            j.result.distinct,
                            " ",
                            i.jsx(M, {
                              children:
                                "values. Search or use conditions. Clear this column to include all values.",
                            }),
                          ],
                        }),
                      i.jsx("div", {
                        className: "lg-filter-values",
                        role: "group",
                        "aria-label": r("Value slicer"),
                        lang: t("Value slicer"),
                        "data-ui-label": "Value slicer",
                        children:
                          (W = j.result) == null
                            ? void 0
                            : W.values.map((oe) => {
                                var he;
                                const de =
                                  z.values === void 0 ||
                                  z.values.some(
                                    (Se) => eh(Se, s.caseSensitive) === oe.key,
                                  );
                                return i.jsxs(
                                  "button",
                                  {
                                    "aria-pressed": de,
                                    disabled:
                                      z.values === void 0 &&
                                      !!(
                                        (he = j.result) != null && he.truncated
                                      ),
                                    onClick: () => C(oe.value),
                                    children: [
                                      i.jsxs("span", {
                                        children: [
                                          oe.value === null || oe.value === ""
                                            ? "(Blanks)"
                                            : String(oe.value),
                                          " ",
                                          i.jsx("small", {
                                            children:
                                              oe.value === null ||
                                              oe.value === ""
                                                ? ""
                                                : typeof oe.value,
                                          }),
                                        ],
                                      }),
                                      i.jsx("span", { children: oe.count }),
                                    ],
                                  },
                                  oe.key,
                                );
                              }),
                      }),
                      i.jsx("button", {
                        onClick: () =>
                          d((oe) => ({
                            ...oe,
                            columns: oe.columns.filter((de) => de.column !== u),
                          })),
                        children: i.jsx(M, { children: "Clear this column" }),
                      }),
                    ],
                  }),
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(M, {
                          children: "Custom column filter",
                        }),
                      }),
                      i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "Column filters intersect. Value selections and custom conditions both apply.",
                        }),
                      }),
                      (z.conditions ?? []).map((oe, de) =>
                        i.jsxs(
                          "div",
                          {
                            children: [
                              N(
                                oe,
                                (he) => {
                                  var Se;
                                  return _({
                                    ...z,
                                    conditions:
                                      (Se = z.conditions) == null
                                        ? void 0
                                        : Se.map((le, ce) =>
                                            ce === de ? he : le,
                                          ),
                                  });
                                },
                                `Column condition ${de + 1}`,
                              ),
                              i.jsxs("button", {
                                onClick: () => {
                                  var he;
                                  return _({
                                    ...z,
                                    conditions:
                                      (he = z.conditions) == null
                                        ? void 0
                                        : he.filter((Se, le) => le !== de),
                                  });
                                },
                                children: [
                                  i.jsx(M, { children: "Remove condition " }),
                                  de + 1,
                                ],
                              }),
                            ],
                          },
                          de,
                        ),
                      ),
                      i.jsx("button", {
                        disabled:
                          (((K = z.conditions) == null ? void 0 : K.length) ??
                            0) >= 2,
                        onClick: () =>
                          _({
                            ...z,
                            conditions: [...(z.conditions ?? []), Bk(u)],
                          }),
                        children: i.jsx(M, {
                          children: "Add column condition",
                        }),
                      }),
                      (((te = z.conditions) == null ? void 0 : te.length) ??
                        0) > 1 &&
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Combine conditions" }),
                            i.jsxs("select", {
                              "aria-label": r("Combine column conditions"),
                              value: z.join ?? "and",
                              onChange: (oe) =>
                                _({ ...z, join: oe.target.value }),
                              lang: t("Combine column conditions"),
                              "data-ui-label": "Combine column conditions",
                              children: [
                                i.jsx("option", {
                                  value: "and",
                                  children: "AND · both",
                                }),
                                i.jsx("option", {
                                  value: "or",
                                  children: "OR · either",
                                }),
                              ],
                            }),
                          ],
                        }),
                      i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "Numbers match numeric cells only; text numbers remain text. Dates use the workbook's 1900/1904 and MDY/DMY policy. Blank means an empty cell or empty text.",
                        }),
                      }),
                    ],
                  }),
                ],
              }),
              i.jsxs("fieldset", {
                children: [
                  i.jsx("legend", {
                    children: i.jsx(M, { children: "Advanced criteria" }),
                  }),
                  i.jsx("p", {
                    children: i.jsx(M, {
                      children:
                        "Meet every condition within a group (AND), and at least one group (OR). These groups also intersect with all column filters.",
                    }),
                  }),
                  (ae = s.advanced) == null
                    ? void 0
                    : ae.map((oe, de) =>
                        i.jsxs(
                          "div",
                          {
                            className: "lg-filter-group",
                            children: [
                              i.jsxs("strong", {
                                children: [
                                  i.jsx(M, { children: "OR group " }),
                                  de + 1,
                                ],
                              }),
                              oe.map((he, Se) =>
                                i.jsxs(
                                  "div",
                                  {
                                    children: [
                                      N(
                                        he,
                                        (le) => T(de, Se, le),
                                        `Group ${de + 1} condition ${Se + 1}`,
                                        true,
                                      ),
                                      i.jsxs("button", {
                                        onClick: () =>
                                          d((le) => {
                                            var ce;
                                            return {
                                              ...le,
                                              advanced:
                                                (ce = le.advanced) == null
                                                  ? void 0
                                                  : ce
                                                      .map((ge, se) =>
                                                        se === de
                                                          ? ge.filter(
                                                              (we, Ee) =>
                                                                Ee !== Se,
                                                            )
                                                          : ge,
                                                      )
                                                      .filter(
                                                        (ge) => ge.length,
                                                      ),
                                            };
                                          }),
                                        children: [
                                          i.jsx(M, {
                                            children: "Remove group ",
                                          }),
                                          de + 1,
                                          " ",
                                          i.jsx(M, { children: "condition " }),
                                          Se + 1,
                                        ],
                                      }),
                                    ],
                                  },
                                  Se,
                                ),
                              ),
                              i.jsxs("button", {
                                onClick: () =>
                                  d((he) => {
                                    var Se;
                                    return {
                                      ...he,
                                      advanced:
                                        (Se = he.advanced) == null
                                          ? void 0
                                          : Se.map((le, ce) =>
                                              ce === de ? [...le, Bk(u)] : le,
                                            ),
                                    };
                                  }),
                                children: [
                                  i.jsx(M, {
                                    children: "Add AND condition to group ",
                                  }),
                                  de + 1,
                                ],
                              }),
                              i.jsxs("button", {
                                onClick: () =>
                                  d((he) => {
                                    var Se;
                                    return {
                                      ...he,
                                      advanced:
                                        (Se = he.advanced) == null
                                          ? void 0
                                          : Se.filter((le, ce) => ce !== de),
                                    };
                                  }),
                                children: [
                                  i.jsx(M, { children: "Remove group " }),
                                  de + 1,
                                ],
                              }),
                            ],
                          },
                          de,
                        ),
                      ),
                  i.jsx("button", {
                    onClick: () =>
                      d((oe) => ({
                        ...oe,
                        advanced: [...(oe.advanced ?? []), [Bk(u)]],
                      })),
                    children: i.jsx(M, { children: "Add OR group" }),
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx("input", {
                        type: "checkbox",
                        checked: !!s.unique,
                        onChange: (oe) =>
                          d((de) => ({ ...de, unique: oe.target.checked })),
                      }),
                      i.jsx(M, {
                        children:
                          "Unique complete records only (keep first matching source row)",
                      }),
                    ],
                  }),
                ],
              }),
              i.jsxs("div", {
                className: "lg-filter-actions",
                children: [
                  i.jsx("button", {
                    onClick: () =>
                      d((oe) => ({
                        ...oe,
                        columns: [],
                        advanced: [],
                        unique: false,
                      })),
                    children: i.jsx(M, { children: "Clear all criteria" }),
                  }),
                  i.jsx("button", {
                    onClick: () => {
                      try {
                        const oe = f0(e.workbook, e.sheetId, s);
                        (p({
                          book: e.workbook,
                          definition: JSON.stringify(s),
                          result: oe,
                        }),
                          v(""));
                      } catch (oe) {
                        (v(
                          oe instanceof Error
                            ? oe.message
                            : "Filter preview failed.",
                        ),
                          p(null));
                      }
                    },
                    children: i.jsx(M, { children: "Preview population" }),
                  }),
                ],
              }),
              m && i.jsx("p", { role: "alert", children: m }),
              V &&
                i.jsxs("div", {
                  className: "lg-filter-preview",
                  role: "status",
                  ref: a,
                  children: [
                    i.jsxs("strong", {
                      children: [
                        V.context.visibleRows,
                        " ",
                        i.jsx(M, { children: "visible / " }),
                        V.context.totalRows,
                        " ",
                        i.jsx(M, { children: "total data rows" }),
                      ],
                    }),
                    i.jsxs("p", {
                      children: [
                        V.context.matchedRows,
                        " ",
                        i.jsx(M, { children: "match criteria ·" }),
                        " ",
                        V.context.filteredRows,
                        " ",
                        i.jsx(M, { children: "filtered out ·" }),
                        " ",
                        V.context.hiddenMatchingRows,
                        " ",
                        i.jsx(M, {
                          children: "matching rows manually hidden.",
                        }),
                      ],
                    }),
                    V.context.visibleRows === 0
                      ? i.jsx("p", {
                          children: i.jsx(M, {
                            children:
                              "No matching visible records. The full source population is unchanged.",
                          }),
                        })
                      : i.jsxs("p", {
                          children: [
                            i.jsx(M, { children: "Original worksheet rows:" }),
                            " ",
                            V.context.visibleSourceRows.slice(0, 25).join(", "),
                            V.context.visibleRows > 25
                              ? " … (first 25 shown)"
                              : "",
                          ],
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
                className: "lg-filter-primary",
                disabled: !V,
                onClick: () => {
                  if (V)
                    try {
                      const oe = D();
                      e.onApply(oe.context.definition);
                    } catch (oe) {
                      (v(
                        oe instanceof Error
                          ? oe.message
                          : "Filter apply failed.",
                      ),
                        p(null));
                    }
                },
                children: i.jsx(M, { children: "Apply filter view" }),
              }),
            ],
          }),
        ],
      }),
    })
  );
}
function Gce({ error: e, readOnly: r, onAccept: t, onCancel: n, onRetry: o }) {
  const a = ne.useRef(null),
    s = e.issues.some((d) => d.style === "stop" || d.evaluationError);
  return (
    ne.useEffect(() => {
      var d, u;
      (u = (d = a.current) == null ? void 0 : d.querySelector("button")) ==
        null || u.focus();
    }, []),
    i.jsx("div", {
      className: "wt-backdrop",
      style: { zIndex: 900 },
      onKeyDown: (d) => {
        var u;
        if ((d.stopPropagation(), d.key === "Escape" && n(), d.key === "Tab")) {
          const l = [
              ...(((u = a.current) == null
                ? void 0
                : u.querySelectorAll("button:not(:disabled)")) ?? []),
            ],
            c = l[0],
            h = l.at(-1);
          d.shiftKey && document.activeElement === c
            ? (d.preventDefault(), h == null || h.focus())
            : !d.shiftKey &&
              document.activeElement === h &&
              (d.preventDefault(), c == null || c.focus());
        }
      },
      children: i.jsxs("div", {
        className: "wt-dialog",
        role: "alertdialog",
        "aria-modal": "true",
        "aria-label": s
          ? "Data validation stopped this change"
          : "Review validation warnings",
        ref: a,
        children: [
          i.jsx("header", {
            children: i.jsx("h2", {
              children: s
                ? "Correct invalid entries"
                : "Review validation warnings",
            }),
          }),
          i.jsxs("div", {
            className: "wt-body",
            children: [
              i.jsxs("p", {
                children: [
                  i.jsx(M, { children: "No cells have changed." }),
                  " ",
                  s
                    ? "Correct these entries or discard the entire batch."
                    : "Accepting keeps the invalid entries and records your decision for the complete batch.",
                ],
              }),
              i.jsxs("p", {
                children: [
                  e.issues.length,
                  " ",
                  i.jsx(M, { children: "rule violations" }),
                ],
              }),
              e.issues.slice(0, 500).map((d, u) =>
                i.jsxs(
                  "div",
                  {
                    children: [
                      i.jsxs("strong", {
                        children: [d.sheetName, "!", d.address, " · ", d.style],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsxs("strong", { children: [d.title, ": "] }),
                          d.message,
                          " ",
                          d.reason,
                        ],
                      }),
                    ],
                  },
                  u,
                ),
              ),
            ],
          }),
          i.jsxs("footer", {
            children: [
              i.jsx("button", {
                onClick: o,
                children: i.jsx(M, { children: "Return to edit" }),
              }),
              i.jsx("button", {
                onClick: n,
                children: i.jsx(M, { children: "Discard changes" }),
              }),
              !s &&
                i.jsx("button", {
                  disabled: r,
                  onClick: t,
                  children: i.jsx(M, {
                    children: "Accept warnings and apply batch",
                  }),
                }),
            ],
          }),
        ],
      }),
    })
  );
}

export { Gce, Xce };
