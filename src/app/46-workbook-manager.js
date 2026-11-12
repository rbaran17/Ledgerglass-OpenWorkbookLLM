import { X, ArrowLeft, ArrowRight, ArrowDown, ArrowUp } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { IB, IQ } from "./03-workbook-model.js";
import { Dr, M } from "./04-cell-input-paste-and-fill.js";
import { QT, g4e, hj, ox } from "./45-sorting-and-worksheets.js";

function v4e(e) {
  var N, R, L;
  const { text, language } = Dr(),
    n = ne.useRef(null),
    o = ne.useRef(e.onClose);
  o.current = e.onClose;
  const [a, s] = ne.useState(
      e.mode === "template"
        ? `${e.workbook.name.slice(0, 185)} — copy`
        : e.workbook.name,
    ),
    [d, u] = ne.useState(e.workbook.metadata?.engagement ?? ""),
    [l, c] = ne.useState(e.workbook.metadata?.period ?? ""),
    [h, f] = ne.useState(""),
    [p, m] = ne.useState(""),
    [v, g] = ne.useState(null),
    [b, y] = ne.useState(""),
    [w, E] = ne.useState(""),
    [A, k] = ne.useState(false),
    F = ne.useRef(true);
  ne.useEffect(
    () => (
      (F.current = true),
      () => {
        F.current = false;
      }
    ),
    [],
  );
  const [I, O] = ne.useState(false),
    [j, z] = ne.useState(0),
    V = !!(
      ((N = e.workbook.structureProtection) != null && N.enabled) ||
      e.workbook.sheets.some((U) => {
        return U.protection?.enabled;
      })
    );
  (ne.useEffect(() => {
    var H;
    const U = document.activeElement;
    return (
      (H = n.current?.querySelector("input, button")) == null || H.focus(),
      () => {
        U != null && U.isConnected && U.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      (g(null), O(false), z(0));
    }, [e.workbook.revision, e.workbook.id]));
  const D = (U) => {
      return (
        e.workbook.sheets.find((X) => X.id === U)?.name ??
        v?.structuralChanges.find((X) => X.path === `sheets.${U}`)?.before
          ?.name ??
        "Removed worksheet"
      );
    },
    B = v
      ? [
          ...v.cellChanges.map((U) => ({
            label: `${D(U.sheetId)}!${U.address}`,
            before: hj(U.before),
            after: hj(U.after),
            attribution: U.attribution,
          })),
          ...[...v.structuralChanges, ...v.metadataChanges].map((U) => {
            const G = U.path.split(".");
            return {
              label:
                U.path === "sheets.order"
                  ? "Worksheet order"
                  : G[0] === "sheets"
                    ? `${D(G[1])}${G[2] ? ` · ${QT[G[2]] ?? G[2]}` : ""}`
                    : (QT[U.path] ?? U.path),
              before:
                U.path === "sheets.order" && Array.isArray(U.before)
                  ? U.before.map(D).join(" → ")
                  : ox(U.before),
              after:
                U.path === "sheets.order" && Array.isArray(U.after)
                  ? U.after.map(D).join(" → ")
                  : ox(U.after),
              attribution: U.attribution,
            };
          }),
        ]
      : [],
    _ = (U) => {
      try {
        (y(""), E(""));
        const G = U();
        G &&
          typeof G.then == "function" &&
          (k(true),
          G.catch((H) => {
            F.current &&
              y(
                H instanceof Error
                  ? H.message
                  : "The workbook operation could not be completed.",
              );
          }).finally(() => {
            F.current && k(false);
          }));
      } catch (G) {
        y(
          G instanceof Error
            ? G.message
            : "The workbook operation could not be completed.",
        );
      }
    };
  return i.jsx("div", {
    className: "wm-backdrop",
    children: i.jsxs("div", {
      ref: n,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": "wm-title",
      className: "wm-dialog",
      onKeyDown: (U) => {
        if (
          (U.key === "Escape" &&
            (U.preventDefault(), U.stopPropagation(), o.current()),
          U.key === "Tab")
        ) {
          const H = Array.from(
              n.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
            Q = H[0],
            X = H.at(-1);
          U.shiftKey && document.activeElement === Q
            ? (U.preventDefault(), X == null || X.focus())
            : !U.shiftKey &&
              document.activeElement === X &&
              (U.preventDefault(), Q == null || Q.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: "wm-title",
              children:
                e.mode === "versions"
                  ? "Version history"
                  : e.mode === "template"
                    ? "Create from this template"
                    : "Workbook properties",
            }),
            i.jsx("button", {
              "aria-label": text("Close workbook manager"),
              onClick: e.onClose,
              lang: language("Close workbook manager"),
              "data-ui-label": "Close workbook manager",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wm-body",
          children: [
            e.mode !== "versions"
              ? i.jsxs("form", {
                  onSubmit: (U) => {
                    (U.preventDefault(),
                      _(() => {
                        if (e.readOnly)
                          throw new Error("This workbook is read-only.");
                        e.mode === "template"
                          ? (e.onCreateWorkbook(
                              IQ(e.workbook, {
                                name: a,
                                engagement: d,
                                period: l,
                              }),
                            ),
                            e.onClose())
                          : (e.onWorkbookChange(
                              IB(e.workbook, {
                                name: a,
                                engagement: d,
                                period: l,
                              }),
                              "Update workbook properties",
                            ),
                            E("Workbook properties saved."));
                      }));
                  },
                  children: [
                    e.mode === "template" &&
                      i.jsx("p", {
                        children: i.jsx(M, {
                          children:
                            "Download a new workbook with the current sheets, formulas and supported layout. Its identity and history will be independent, with a record of this template and its revision. The current workbook stays open.",
                        }),
                      }),
                    i.jsxs("label", {
                      children: [
                        i.jsx(M, { children: "Workbook name" }),
                        i.jsx("input", {
                          "aria-label": text("Workbook name"),
                          required: true,
                          maxLength: 200,
                          value: a,
                          disabled: e.readOnly || A,
                          onChange: (U) => s(U.target.value),
                          lang: language("Workbook name"),
                          "data-ui-label": "Workbook name",
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "wm-fields",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Engagement" }),
                            e.mode === "template" ? " *" : "",
                            i.jsx("input", {
                              "aria-label": text("Engagement"),
                              required: e.mode === "template",
                              value: d,
                              maxLength: 200,
                              disabled: e.readOnly || A,
                              onChange: (U) => u(U.target.value),
                              lang: language("Engagement"),
                              "data-ui-label": "Engagement",
                            }),
                          ],
                        }),
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Period" }),
                            e.mode === "template" ? " *" : "",
                            i.jsx("input", {
                              "aria-label": text("Period"),
                              required: e.mode === "template",
                              value: l,
                              maxLength: 100,
                              placeholder: text("For example, FY2026"),
                              disabled: e.readOnly || A,
                              onChange: (U) => c(U.target.value),
                              lang: language("Period"),
                              "data-ui-label": "Period",
                            }),
                          ],
                        }),
                      ],
                    }),
                    e.mode === "properties" &&
                      i.jsxs("dl", {
                        children: [
                          i.jsx("dt", {
                            children: i.jsx(M, { children: "Created" }),
                          }),
                          i.jsx("dd", {
                            children:
                              (R = e.workbook.metadata) != null && R.createdAt
                                ? new Date(
                                    e.workbook.metadata.createdAt,
                                  ).toLocaleString()
                                : "Not recorded in this file",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(M, { children: "Modified" }),
                          }),
                          i.jsx("dd", {
                            children:
                              (L = e.workbook.metadata) != null && L.modifiedAt
                                ? new Date(
                                    e.workbook.metadata.modifiedAt,
                                  ).toLocaleString()
                                : "Not recorded in this file",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(M, { children: "Revision" }),
                          }),
                          i.jsx("dd", { children: e.workbook.revision }),
                          e.workbook.metadata?.template &&
                            i.jsxs(i.Fragment, {
                              children: [
                                i.jsx("dt", {
                                  children: i.jsx(M, { children: "Template" }),
                                }),
                                i.jsxs("dd", {
                                  children: [
                                    e.workbook.metadata.template.name,
                                    " ",
                                    i.jsx(M, { children: "· version" }),
                                    " ",
                                    e.workbook.metadata.template.version,
                                    e.workbook.metadata.template
                                      .sourceRevision !== undefined
                                      ? ` · source revision ${e.workbook.metadata.template.sourceRevision}`
                                      : "",
                                  ],
                                }),
                              ],
                            }),
                        ],
                      }),
                    i.jsx("button", {
                      type: "submit",
                      className: "wm-primary",
                      disabled: e.readOnly || A,
                      children:
                        e.mode === "template"
                          ? "Create independent workbook"
                          : "Save properties",
                    }),
                  ],
                })
              : i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("p", {
                      children: [
                        e.persistent
                          ? "Named checkpoints are saved in this browser with this workbook."
                          : "Named checkpoints are kept for this editor session. The embedding application can provide a persistent version store.",
                        " ",
                        i.jsx(M, {
                          children:
                            "Restoring records a new revision and keeps intervening change history.",
                        }),
                      ],
                    }),
                    i.jsxs("form", {
                      className: "wm-checkpoint",
                      onSubmit: (U) => {
                        (U.preventDefault(),
                          _(() => {
                            if (e.readOnly)
                              throw new Error("This workbook is read-only.");
                            const G = e.onCheckpoint(h),
                              H = () => {
                                F.current && (f(""), E("Checkpoint saved."));
                              };
                            return G ? G.then(H) : H();
                          }));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(M, { children: "Checkpoint name" }),
                            i.jsx("input", {
                              "aria-label": text("Checkpoint name"),
                              value: h,
                              required: true,
                              maxLength: 100,
                              disabled: e.readOnly || A,
                              onChange: (U) => f(U.target.value),
                              lang: language("Checkpoint name"),
                              "data-ui-label": "Checkpoint name",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: e.readOnly || A,
                          type: "submit",
                          className: "wm-primary",
                          children: i.jsx(M, { children: "Save checkpoint" }),
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "wm-muted",
                      children: i.jsx(M, {
                        children:
                          "Up to 20 checkpoints are retained. Saving another checkpoint replaces the oldest checkpoint when the limit is reached.",
                      }),
                    }),
                    e.checkpoints.length
                      ? i.jsx("div", {
                          className: "wm-versions",
                          "aria-label": text("Saved checkpoints"),
                          lang: language("Saved checkpoints"),
                          "data-ui-label": "Saved checkpoints",
                          children: [...e.checkpoints].reverse().map((U) =>
                            i.jsxs(
                              "button",
                              {
                                "aria-pressed": p === U.id,
                                onClick: () =>
                                  _(() => {
                                    (m(U.id),
                                      O(false),
                                      z(0),
                                      g(e.onCompare(U.id)));
                                  }),
                                children: [
                                  i.jsx("strong", { children: U.label }),
                                  i.jsxs("span", {
                                    children: [
                                      i.jsx(M, { children: "Revision " }),
                                      U.revision,
                                      " ·",
                                      " ",
                                      new Date(U.createdAt).toLocaleString(),
                                    ],
                                  }),
                                ],
                              },
                              U.id,
                            ),
                          ),
                        })
                      : i.jsx("p", {
                          children: i.jsx(M, {
                            children: "No checkpoints have been saved.",
                          }),
                        }),
                    v &&
                      i.jsxs("section", {
                        className: "wm-comparison",
                        "aria-label": text("Version comparison"),
                        lang: language("Version comparison"),
                        "data-ui-label": "Version comparison",
                        children: [
                          i.jsx("h3", {
                            children: i.jsx(M, {
                              children: "Changes since the selected checkpoint",
                            }),
                          }),
                          i.jsxs("p", {
                            children: [
                              v.cellChanges.length,
                              " ",
                              i.jsx(M, { children: "cell changes ·" }),
                              " ",
                              v.structuralChanges.length,
                              " ",
                              i.jsx(M, { children: "structure changes ·" }),
                              " ",
                              v.metadataChanges.length,
                              " ",
                              i.jsx(M, { children: "property changes" }),
                            ],
                          }),
                          i.jsx("div", {
                            className: "wm-diffs",
                            children: B.slice(j * 100, (j + 1) * 100).map(
                              (U, G) =>
                                i.jsxs(
                                  "div",
                                  {
                                    children: [
                                      i.jsx("strong", { children: U.label }),
                                      i.jsxs("span", {
                                        children: [U.before, " → ", U.after],
                                      }),
                                      i.jsx("small", {
                                        children: g4e(U.attribution),
                                      }),
                                    ],
                                  },
                                  `${j}-${G}`,
                                ),
                            ),
                          }),
                          B.length > 100 &&
                            i.jsxs("div", {
                              className: "wm-pagination",
                              children: [
                                i.jsx("button", {
                                  disabled: j === 0,
                                  onClick: () => z((U) => U - 1),
                                  children: i.jsx(M, {
                                    children: "Previous changes",
                                  }),
                                }),
                                i.jsxs("span", {
                                  children: [
                                    i.jsx(M, { children: "Page " }),
                                    j + 1,
                                    " ",
                                    i.jsx(M, { children: "of " }),
                                    Math.ceil(B.length / 100),
                                  ],
                                }),
                                i.jsx("button", {
                                  disabled: (j + 1) * 100 >= B.length,
                                  onClick: () => z((U) => U + 1),
                                  children: i.jsx(M, {
                                    children: "Next changes",
                                  }),
                                }),
                              ],
                            }),
                          V &&
                            i.jsx("p", {
                              children: i.jsx(M, {
                                children:
                                  "Unprotect the workbook and its worksheets before restoring a checkpoint.",
                              }),
                            }),
                          i.jsxs("label", {
                            className: "wm-confirm",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: I,
                                disabled: e.readOnly || V || A,
                                onChange: (U) => O(U.target.checked),
                              }),
                              i.jsx(M, {
                                children:
                                  "Restore this checkpoint as a new revision.",
                              }),
                            ],
                          }),
                          i.jsx("button", {
                            className: "wm-primary",
                            disabled: e.readOnly || V || !I || A,
                            onClick: () =>
                              _(() => {
                                const U = e.onRestore(p),
                                  G = () => {
                                    F.current &&
                                      (g(null),
                                      O(false),
                                      E(
                                        "Checkpoint restored as a new revision.",
                                      ));
                                  };
                                return U ? U.then(G) : G();
                              }),
                            children: i.jsx(M, {
                              children: "Restore checkpoint",
                            }),
                          }),
                        ],
                      }),
                  ],
                }),
            A &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(M, { children: "Saving workbook changes…" }),
              }),
            b &&
              i.jsx("p", { role: "alert", className: "wm-error", children: b }),
            w &&
              i.jsx("p", {
                role: "status",
                className: "wm-notice",
                children: w,
              }),
          ],
        }),
      ],
    }),
  });
}
const b4e = [
    { value: "down", label: "Fill down", icon: ArrowDown },
    { value: "right", label: "Fill right", icon: ArrowRight },
    { value: "up", label: "Fill up", icon: ArrowUp },
    { value: "left", label: "Fill left", icon: ArrowLeft },
  ],
  pj = (e) =>
    e instanceof Error ? e.message : "This change could not be completed.";

export { b4e, pj, v4e };
