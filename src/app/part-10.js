import { X, Download, CircleCheck, FileInput } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { Cd, FU, ut } from "./part-03.js";
import { Dr, M, ci } from "./part-04.js";
import { oT } from "./part-05.js";
import { H1, MC } from "./part-07.js";
import { Due, H$, Iue, Mue, Oue, Rue, _ue, jue, q$, yh } from "./part-09.js";

function Fue(e) {
  var we, Ee;
  const r = ne.useId(),
    t = ne.useRef(null),
    n = ne.useRef(e),
    o = ne.useRef(0),
    a = ne.useRef(true),
    s = yh();
  n.current = e;
  const [d, u] = ne.useState(null),
    [l, c] = ne.useState(""),
    [h, f] = ne.useState("Imported data"),
    [p, m] = ne.useState(false),
    [v, g] = ne.useState(false),
    [b, y] = ne.useState({}),
    [w, E] = ne.useState(null),
    [A, k] = ne.useState(null),
    [F, I] = ne.useState(false),
    [O, j] = ne.useState(""),
    [z, V] = ne.useState(""),
    [D, B] = ne.useState(0),
    [_, C] = ne.useState(0),
    [T, N] = ne.useState({}),
    [R, L] = ne.useState(""),
    [$, U] = ne.useState("0:0");
  ne.useEffect(() => {
    var ye;
    a.current = true;
    const Ae = document.activeElement;
    return (
      (ye = t.current) == null || ye.focus(),
      () => {
        ((a.current = false),
          o.current++,
          Ae != null && Ae.isConnected && Ae.focus());
      }
    );
  }, []);
  const G = async (Ae) => {
      const ye = ++o.current;
      (I(true), j(""), V(""));
      try {
        await Ae(ye);
      } catch (Te) {
        a.current &&
          ye === o.current &&
          j(
            Te instanceof Error
              ? Te.message
              : "The source could not be processed.",
          );
      } finally {
        a.current && ye === o.current && I(false);
      }
    },
    H = (Ae) => {
      var ye;
      (u(Ae),
        c(((ye = Ae.tables[0]) == null ? void 0 : ye.id) ?? ""),
        E(null),
        y({}),
        N({}),
        m(false),
        B(0),
        C(0));
    },
    Q = (Ae) => (
      u(null),
      E(null),
      g(false),
      G(async (ye) => {
        const Te = await Rue(Ae, s.begin());
        a.current && o.current === ye && H(Te);
      })
    );
  ne.useEffect(() => {
    e.initialFile && Q(e.initialFile);
  }, [e.initialFile]);
  const X$1 = (Ae) =>
      G(async (ye) => {
        const Te = d,
          Pe = await MC(Ae, s.begin());
        a.current && o.current === ye && H(q$(Te, Pe));
      }),
    J = () =>
      G(async (Ae) => {
        const ye = await Iue(d, e.extractor, v, s.begin());
        a.current && o.current === Ae && H(ye);
      }),
    q = (Ae, ye) => (n.current.onDownload ?? jue)({ filename: Ae, blob: ye }),
    W = (Ae) =>
      G(async () => {
        const ye = Ae ? Oue(n.current.workbook, R) : H$(d);
        await q(
          ye.fileName,
          new Blob([new Uint8Array(ye.bytes)], {
            type: "application/octet-stream",
          }),
        );
      }),
    K = () => E(null),
    te = d == null ? void 0 : d.tables.find((Ae) => Ae.id === l),
    ae =
      (te == null
        ? void 0
        : te.rows
            .flatMap((Ae, ye) =>
              Ae.map((Te, Pe) => ({
                cell: Te,
                row: ye,
                column: Pe,
                key: `${ye}:${Pe}`,
              })),
            )
            .filter((Ae) => Ae.cell.confidence < Cd.confidence)) ?? [],
    oe = ae.filter((Ae) => !b[Ae.key]).length,
    de =
      (we = e.workbook.sourceImports) == null
        ? void 0
        : we.find((Ae) => Ae.id === R),
    [he, Se] = $.split(":").map(Number),
    le = de ? Mue(e.workbook, de.id, he, Se) : null,
    ce = () => {
      j("");
      try {
        if (Object.values(T).some(Boolean))
          throw Error("Finish each corrected value before previewing.");
        const Ae = _ue(e.workbook, d, {
          tableId: l,
          sheetName: h,
          ambiguityReviewed: p,
          resolutions: b,
        });
        (E(Ae), k(e.workbook));
      } catch (Ae) {
        j(Ae.message);
      }
    },
    ge = () => {
      j("");
      try {
        if (n.current.workbook !== A)
          throw Error("The workbook changed. Preview the import again.");
        const Ae = Due(n.current.workbook, w, { readOnly: n.current.readOnly });
        (e.onWorkbookChange(Ae, "Import reviewed source"),
          L(w.record.id),
          U("0:0"),
          E(null),
          V(
            "Reviewed source imported into a new worksheet. The original and cell evidence are retained in native JSON.",
          ));
      } catch (Ae) {
        j(Ae.message);
      }
    },
    se = (Ae, ye) => {
      (K(),
        N((Te) => ({ ...Te, [Ae]: false })),
        y((Te) => ({ ...Te, [Ae]: ye })));
    };
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": r,
      lang: "en",
      tabIndex: -1,
      ref: t,
      className: "lg-reports-dialog lg-formats-dialog lg-source-dialog",
      onKeyDown: (Ae) => {
        var ye, Te;
        if (
          (Ae.key === "Escape" && (Ae.stopPropagation(), e.onClose()),
          Ae.key === "Tab")
        ) {
          const Pe = [
            ...t.current.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
            ),
          ].filter((qe) => qe.getClientRects().length > 0);
          Ae.shiftKey &&
          (document.activeElement === Pe[0] ||
            document.activeElement === t.current)
            ? (Ae.preventDefault(), (ye = Pe.at(-1)) == null || ye.focus())
            : !Ae.shiftKey &&
              document.activeElement === Pe.at(-1) &&
              (Ae.preventDefault(), (Te = Pe[0]) == null || Te.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", { children: "DATA & SOURCE EVIDENCE" }),
                i.jsx("h2", { id: r, children: "Import sources" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              "aria-label": "Close import sources",
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-reports-body",
          children: [
            i.jsx("p", {
              children:
                "Bring JSON or XML tables into a new worksheet. Document extraction can be supplied by your host or a matching extraction JSON file. Review first; the current workbook stays unchanged until import.",
            }),
            O &&
              i.jsx("p", {
                role: "alert",
                className: "lg-formats-error",
                children: O,
              }),
            z &&
              i.jsxs("p", {
                role: "status",
                className: "lg-source-success",
                children: [i.jsx(CircleCheck, { size: 17 }), z],
              }),
            i.jsxs("div", {
              className: "lg-formats-actions",
              children: [
                i.jsxs("label", {
                  className: "lg-source-file",
                  children: [
                    i.jsx(FileInput, { size: 18 }),
                    "Choose source file",
                    i.jsx("input", {
                      "aria-label": "Import source file",
                      type: "file",
                      accept: ".json,.xml,.pdf,.png,.jpg,.jpeg,.html,.htm",
                      disabled: F,
                      onChange: (Ae) => {
                        var Te;
                        const ye =
                          (Te = Ae.target.files) == null ? void 0 : Te[0];
                        ((Ae.target.value = ""), ye && Q(ye));
                      },
                    }),
                  ],
                }),
                F &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("span", {
                        role: "status",
                        children: "Processing source…",
                      }),
                      i.jsx("button", {
                        type: "button",
                        onClick: () => {
                          (s.cancel(),
                            o.current++,
                            I(false),
                            V("Source processing cancelled."));
                        },
                        children: "Cancel processing",
                      }),
                    ],
                  }),
              ],
            }),
            d &&
              i.jsxs("section", {
                className: "lg-formats-source",
                children: [
                  i.jsxs("div", {
                    className: "lg-source-heading",
                    children: [
                      i.jsxs("div", {
                        children: [
                          i.jsx("h3", { children: d.fileName }),
                          i.jsxs("p", {
                            children: [
                              d.format.toUpperCase(),
                              " ·",
                              " ",
                              d.byteLength.toLocaleString(),
                              " bytes ·",
                              " ",
                              d.tables.length,
                              " table candidates",
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("button", {
                        type: "button",
                        disabled: F,
                        onClick: () => void W(false),
                        children: [
                          i.jsx(Download, { size: 16 }),
                          "Download original source",
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("p", {
                    className: "lg-formats-hash",
                    children: ["SHA-256: ", d.sha256],
                  }),
                  i.jsx("ul", {
                    children: d.warnings.map((Ae, ye) =>
                      i.jsx("li", { children: Ae }, ye),
                    ),
                  }),
                  d.method === "document-adapter" &&
                    i.jsxs("div", {
                      className: "lg-source-extractor",
                      children: [
                        i.jsx("h4", { children: "Document extraction" }),
                        i.jsx("p", {
                          children:
                            "No OCR or document provider runs automatically. Regions use a 1-based page and normalized top-left coordinates.",
                        }),
                        e.extractor
                          ? i.jsxs(i.Fragment, {
                              children: [
                                i.jsxs("label", {
                                  className: "lg-formats-ack",
                                  children: [
                                    i.jsx("input", {
                                      type: "checkbox",
                                      checked: v,
                                      disabled: F,
                                      onChange: (Ae) => g(Ae.target.checked),
                                    }),
                                    "I authorize ",
                                    e.extractor.label,
                                    " to process this selected source.",
                                  ],
                                }),
                                i.jsx("button", {
                                  type: "button",
                                  disabled: !v || F,
                                  onClick: () => void J(),
                                  children: "Run host extraction",
                                }),
                              ],
                            })
                          : i.jsx("p", {
                              children:
                                "Your host has not connected a document extractor. Load a matching extraction JSON result below.",
                            }),
                        i.jsxs("label", {
                          children: [
                            "Load extraction JSON",
                            i.jsx("input", {
                              type: "file",
                              accept: ".json",
                              disabled: F,
                              "aria-label": "Load extraction JSON",
                              onChange: (Ae) => {
                                var Te;
                                const ye =
                                  (Te = Ae.target.files) == null
                                    ? void 0
                                    : Te[0];
                                ((Ae.target.value = ""), ye && X$1(ye));
                              },
                            }),
                          ],
                        }),
                      ],
                    }),
                ],
              }),
            te &&
              i.jsxs("section", {
                className: "lg-source-review",
                children: [
                  i.jsxs("div", {
                    className: "lg-source-options",
                    children: [
                      i.jsxs("label", {
                        children: [
                          "Source table",
                          i.jsx("select", {
                            "aria-label": "Source table",
                            value: l,
                            disabled: F,
                            onChange: (Ae) => {
                              (c(Ae.target.value),
                                E(null),
                                y({}),
                                m(false),
                                B(0));
                            },
                            children: d.tables.map((Ae) =>
                              i.jsxs(
                                "option",
                                {
                                  value: Ae.id,
                                  children: [
                                    Ae.name,
                                    " · ",
                                    Ae.rows.length,
                                    " rows",
                                  ],
                                },
                                Ae.id,
                              ),
                            ),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "New worksheet name",
                          i.jsx("input", {
                            "aria-label": "New worksheet name",
                            value: h,
                            maxLength: 31,
                            onChange: (Ae) => {
                              (f(Ae.target.value), K());
                            },
                          }),
                        ],
                      }),
                    ],
                  }),
                  (te.ambiguous || d.tables.length > 1) &&
                    i.jsxs("label", {
                      className: "lg-formats-ack",
                      children: [
                        i.jsx("input", {
                          type: "checkbox",
                          checked: p,
                          onChange: (Ae) => {
                            (m(Ae.target.checked), K());
                          },
                        }),
                        "I reviewed the candidates and confirm this table selection.",
                      ],
                    }),
                  i.jsxs("p", {
                    children: [
                      te.rows.length,
                      " rows · ",
                      te.columns.length,
                      " fields ·",
                      " ",
                      oe,
                      " low-confidence cells awaiting review. Missing fields remain distinguishable from null in source evidence.",
                    ],
                  }),
                  i.jsx("div", {
                    className: "lg-formats-table lg-source-preview",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("caption", {
                          children:
                            "Source preview — values are not yet in the workbook",
                        }),
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { scope: "col", children: "Row" }),
                              te.columns.map((Ae, ye) =>
                                i.jsxs(
                                  "th",
                                  {
                                    scope: "col",
                                    children: [
                                      Ae.name,
                                      i.jsx("small", {
                                        children: Ae.path || "(root)",
                                      }),
                                    ],
                                  },
                                  ye,
                                ),
                              ),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: te.rows
                            .slice(D * 20, D * 20 + 20)
                            .map((Ae, ye) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      scope: "row",
                                      children: D * 20 + ye + 1,
                                    }),
                                    Ae.map((Te, Pe) =>
                                      i.jsxs(
                                        "td",
                                        {
                                          className:
                                            Te.confidence < 0.9
                                              ? "lg-source-low"
                                              : "",
                                          children: [
                                            String(
                                              Te.value ??
                                                (Te.missing
                                                  ? "(missing)"
                                                  : "(null)"),
                                            ),
                                            i.jsxs("small", {
                                              children: [
                                                Te.type,
                                                Te.declaredType
                                                  ? " · " + Te.declaredType
                                                  : "",
                                                " · ",
                                                Math.round(Te.confidence * 100),
                                                "%",
                                                Te.region
                                                  ? " · page " + Te.region.page
                                                  : "",
                                              ],
                                            }),
                                          ],
                                        },
                                        Pe,
                                      ),
                                    ),
                                  ],
                                },
                                ye,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "lg-formats-actions",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        disabled: D === 0,
                        onClick: () => B((Ae) => Ae - 1),
                        children: "Previous source rows",
                      }),
                      i.jsxs("span", {
                        children: [
                          "Page ",
                          D + 1,
                          " of ",
                          Math.ceil(te.rows.length / 20),
                        ],
                      }),
                      i.jsx("button", {
                        type: "button",
                        disabled: (D + 1) * 20 >= te.rows.length,
                        onClick: () => B((Ae) => Ae + 1),
                        children: "Next source rows",
                      }),
                    ],
                  }),
                  ae.length > 0 &&
                    i.jsxs("section", {
                      className: "lg-source-flags",
                      children: [
                        i.jsx("h3", { children: "Review uncertain cells" }),
                        i.jsx("p", {
                          children:
                            "Accept after comparison with the source, correct the value, or exclude it. Excluded cells become blanks and stay documented in the review record.",
                        }),
                        ae
                          .slice(_ * 10, _ * 10 + 10)
                          .map(({ cell: Ae, row: ye, column: Te, key: Pe }) => {
                            var qe, ht;
                            return i.jsxs(
                              "div",
                              {
                                className: "lg-source-flag",
                                children: [
                                  i.jsxs("div", {
                                    children: [
                                      i.jsxs("strong", {
                                        children: [
                                          ut(Te),
                                          ye + 2,
                                          " · ",
                                          Math.round(Ae.confidence * 100),
                                          "% confidence",
                                        ],
                                      }),
                                      i.jsx("p", {
                                        children: String(Ae.value ?? "(null)"),
                                      }),
                                      i.jsxs("small", {
                                        children: [
                                          Ae.path,
                                          " ·",
                                          " ",
                                          Ae.region
                                            ? `Page ${Ae.region.page} · x ${Ae.region.x}, y ${Ae.region.y}, w ${Ae.region.width}, h ${Ae.region.height}`
                                            : "Structured field",
                                        ],
                                      }),
                                    ],
                                  }),
                                  i.jsxs("label", {
                                    children: [
                                      "Review ",
                                      Pe,
                                      i.jsxs("select", {
                                        "aria-label": "Review " + Pe,
                                        value:
                                          ((qe = b[Pe]) == null
                                            ? void 0
                                            : qe.decision) ?? "",
                                        onChange: (Re) => {
                                          Re.target.value
                                            ? se(Pe, {
                                                decision: Re.target.value,
                                                ...(Re.target.value ===
                                                "correct"
                                                  ? { value: Ae.value }
                                                  : {}),
                                              })
                                            : (y((Me) => {
                                                const ke = { ...Me };
                                                return (delete ke[Pe], ke);
                                              }),
                                              K());
                                        },
                                        children: [
                                          i.jsx("option", {
                                            value: "",
                                            children: "Review required",
                                          }),
                                          i.jsx("option", {
                                            value: "accept",
                                            children: "Accept source value",
                                          }),
                                          i.jsx("option", {
                                            value: "exclude",
                                            children: "Exclude from worksheet",
                                          }),
                                          i.jsx("option", {
                                            value: "correct",
                                            children: "Correct value",
                                          }),
                                        ],
                                      }),
                                    ],
                                  }),
                                  ((ht = b[Pe]) == null
                                    ? void 0
                                    : ht.decision) === "correct" &&
                                    i.jsxs("label", {
                                      children: [
                                        "Corrected value (JSON scalar)",
                                        i.jsx("input", {
                                          "aria-label": "Corrected value " + Pe,
                                          "aria-invalid": T[Pe] || void 0,
                                          defaultValue: JSON.stringify(
                                            Ae.value,
                                          ),
                                          onChange: (Re) => {
                                            try {
                                              const Me = JSON.parse(
                                                Re.target.value,
                                              );
                                              if (
                                                Me !== null &&
                                                ![
                                                  "string",
                                                  "number",
                                                  "boolean",
                                                ].includes(typeof Me)
                                              )
                                                throw Error();
                                              (se(Pe, {
                                                decision: "correct",
                                                value: Me,
                                              }),
                                                Re.target.setCustomValidity(
                                                  "",
                                                ));
                                            } catch {
                                              (N((Me) => ({
                                                ...Me,
                                                [Pe]: true,
                                              })),
                                                E(null),
                                                Re.target.setCustomValidity(
                                                  "Enter a JSON string, number, boolean or null.",
                                                ));
                                            }
                                          },
                                        }),
                                      ],
                                    }),
                                ],
                              },
                              Pe,
                            );
                          }),
                        i.jsxs("div", {
                          className: "lg-formats-actions",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              disabled: _ === 0,
                              onClick: () => C((Ae) => Ae - 1),
                              children: "Previous uncertain cells",
                            }),
                            i.jsxs("span", {
                              children: [
                                "Review page ",
                                _ + 1,
                                " of",
                                " ",
                                Math.ceil(ae.length / 10),
                              ],
                            }),
                            i.jsx("button", {
                              type: "button",
                              disabled: (_ + 1) * 10 >= ae.length,
                              onClick: () => C((Ae) => Ae + 1),
                              children: "Next uncertain cells",
                            }),
                          ],
                        }),
                      ],
                    }),
                  i.jsx("div", {
                    className: "lg-formats-actions",
                    children: i.jsx("button", {
                      type: "button",
                      className: "lg-source-primary",
                      disabled:
                        F ||
                        e.readOnly ||
                        Object.values(T).some(Boolean) ||
                        oe > 0 ||
                        ((te.ambiguous || d.tables.length > 1) && !p),
                      onClick: ce,
                      children: "Preview import",
                    }),
                  }),
                  w &&
                    i.jsxs("section", {
                      className: "lg-source-confirm",
                      "aria-label": "Import confirmation",
                      children: [
                        i.jsxs("h3", {
                          children: ["Ready for ", w.sheet.name],
                        }),
                        i.jsxs("p", {
                          children: [
                            w.accepted,
                            " accepted · ",
                            w.corrected,
                            " corrected ·",
                            " ",
                            w.excluded,
                            " excluded cells. Source evidence will be retained in native JSON.",
                          ],
                        }),
                        A !== e.workbook &&
                          i.jsx("p", {
                            role: "alert",
                            children: "The workbook changed. Preview again.",
                          }),
                        i.jsx("button", {
                          type: "button",
                          className: "lg-source-primary",
                          disabled: e.readOnly || A !== e.workbook || F,
                          onClick: ge,
                          children: "Import reviewed table",
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => E(null),
                          children: "Cancel import preview",
                        }),
                      ],
                    }),
                ],
              }),
            !!((Ee = e.workbook.sourceImports) != null && Ee.length) &&
              i.jsxs("section", {
                className: "lg-formats-source",
                children: [
                  i.jsx("h3", { children: "Retained sources & cell evidence" }),
                  i.jsxs("label", {
                    children: [
                      "Retained import",
                      i.jsxs("select", {
                        "aria-label": "Retained import",
                        value: R,
                        onChange: (Ae) => {
                          (L(Ae.target.value), U("0:0"));
                        },
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Select an imported source",
                          }),
                          e.workbook.sourceImports.map((Ae) =>
                            i.jsxs(
                              "option",
                              {
                                value: Ae.id,
                                children: [
                                  Ae.sheetName,
                                  " · ",
                                  Ae.original.fileName,
                                ],
                              },
                              Ae.id,
                            ),
                          ),
                        ],
                      }),
                    ],
                  }),
                  de &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("p", {
                          className: "lg-formats-hash",
                          children: ["SHA-256: ", de.original.sha256],
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () => void W(true),
                          children: "Download retained source",
                        }),
                        i.jsx("button", {
                          type: "button",
                          onClick: () =>
                            void G(async () => {
                              const { base64: Ae, ...ye } = de.original;
                              await q(
                                de.original.fileName + ".import-review.json",
                                new Blob(
                                  [
                                    JSON.stringify(
                                      { ...de, original: ye },
                                      null,
                                      2,
                                    ),
                                  ],
                                  { type: "application/json" },
                                ),
                              );
                            }),
                          children: "Download source review",
                        }),
                        i.jsxs("div", {
                          className: "lg-source-options",
                          children: [
                            i.jsxs("label", {
                              children: [
                                "Source record",
                                i.jsx(ci, {
                                  "aria-label": "Source record",
                                  min: 1,
                                  max: de.table.rows.length,
                                  value: he + 1,
                                  onValueChange: (Ae) => {
                                    const ye = Number(Ae);
                                    Number.isInteger(ye) &&
                                      ye >= 1 &&
                                      ye <= de.table.rows.length &&
                                      U(ye - 1 + ":" + Se);
                                  },
                                }),
                              ],
                            }),
                            i.jsxs("label", {
                              children: [
                                "Source field",
                                i.jsx("select", {
                                  "aria-label": "Source field",
                                  value: Se,
                                  onChange: (Ae) =>
                                    U(he + ":" + Ae.target.value),
                                  children: de.table.columns.map((Ae, ye) =>
                                    i.jsxs(
                                      "option",
                                      {
                                        value: ye,
                                        children: [ut(ye), " · ", Ae.name],
                                      },
                                      ye,
                                    ),
                                  ),
                                }),
                              ],
                            }),
                          ],
                        }),
                        le &&
                          i.jsxs("dl", {
                            className: "lg-source-evidence",
                            children: [
                              i.jsx("dt", { children: "Link status" }),
                              i.jsx("dd", { children: le.status }),
                              i.jsx("dt", { children: "Source field" }),
                              i.jsx("dd", {
                                children: le.source.path || "(root)",
                              }),
                              i.jsx("dt", {
                                children: "Original type & value",
                              }),
                              i.jsxs("dd", {
                                children: [
                                  le.source.type,
                                  " ·",
                                  " ",
                                  String(le.source.value ?? "(null)"),
                                ],
                              }),
                              i.jsx("dt", { children: "Review decision" }),
                              i.jsx("dd", { children: le.decision }),
                              i.jsx("dt", { children: "Source region" }),
                              i.jsx("dd", {
                                children: le.source.region
                                  ? `Page ${le.source.region.page} · x ${le.source.region.x}, y ${le.source.region.y}, width ${le.source.region.width}, height ${le.source.region.height}`
                                  : "Structured field path",
                              }),
                            ],
                          }),
                        i.jsx("p", {
                          children:
                            "Evidence records the import position. Edits, moves or sorts can change the link status; compare the original before relying on a changed cell.",
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("p", {
              className: "lg-source-limits",
              children:
                "Limits: 4 MiB per source; 20 table candidates; 1,000 rows, 100 columns and 20,000 cells across candidates; 10 retained imports and 12 MiB original bytes. Native JSON preserves originals and provenance; XLSX and delimited exports carry worksheet values without this evidence.",
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: e.readOnly
                ? "Read-only: inspection and original downloads are available."
                : "Import appends a new worksheet. Undo restores the prior workbook.",
            }),
            i.jsx("button", {
              type: "button",
              onClick: e.onClose,
              children: "Close",
            }),
          ],
        }),
      ],
    }),
  });
}
async function Bue(e, r, t, n) {
  const o = oT(e);
  if (t !== true)
    throw Error(
      "Acknowledge conversion losses and compare the source before importing the converted workbook.",
    );
  if (!/\.xlsx$/i.test(r.name))
    throw Error(
      "Choose the converted XLSX copy; renaming the original does not convert it.",
    );
  const a = await H1(r, n);
  if ((oT(e), !a.workbook.xlsxSource || !a.report))
    throw Error(
      "A screened XLSX conversion and compatibility report are required.",
    );
  a.workbook.alternateSource = FU(o, e.fileName, e.format, {
    method: "external-conversion",
    convertedFileName: r.name,
    convertedSha256: a.workbook.xlsxSource.sha256,
    reviewedAt: new Date().toISOString(),
    equivalence: "not-verified",
  });
  const s =
    "An externally converted XLSX was selected and conversion losses acknowledged. This application has not verified that its formulas, dates, styles or special parts match the retained original. Compare them in the source application. The original is retained under File > Workbook formats and in native JSON.";
  return (
    (a.warnings = [...new Set([...a.warnings, s])]),
    (a.report.warnings = a.warnings),
    (a.report.alternateOriginal = {
      fileName: e.fileName,
      byteLength: e.fileBytes,
      sha256: e.sha256,
      format: e.format,
    }),
    a
  );
}
function Lue(e) {
  return { bytes: oT(e), fileName: e.fileName, sha256: e.sha256 };
}
const Tk = {
  editable: "Editable",
  "preserved-opaque": "Preserved opaque",
  flattened: "Flattened",
  dropped: "Dropped",
};
function Pue({ inventory: e }) {
  const { text: r, language: t } = Dr(),
    [n, o] = ne.useState(""),
    [a, s] = ne.useState("all"),
    [d, u] = ne.useState(0),
    l = e.parts.filter(
      (f) =>
        (a === "all" || f.disposition === a) &&
        [
          f.path,
          f.family,
          ...f.locations.map((p) => p.sheet ?? ""),
          ...f.features.map(
            (p) =>
              `${p.feature} ${p.location.range ?? ""} ${p.location.detail ?? ""}`,
          ),
        ]
          .join(" ")
          .toLowerCase()
          .includes(n.toLowerCase()),
    ),
    c = Math.max(0, Math.ceil(l.length / 25) - 1),
    h = Math.min(d, c);
  return i.jsxs("section", {
    className: "lg-xlsx-inventory",
    "aria-label": r("Source package inventory"),
    lang: t("Source package inventory"),
    "data-ui-label": "Source package inventory",
    children: [
      i.jsx("h3", {
        children: i.jsx(M, { children: "Source parts and export outcome" }),
      }),
      i.jsxs("p", {
        children: [
          e.parts.length,
          " ",
          i.jsx(M, { children: "entries inventoried ·" }),
          " ",
          e.counts.editable,
          " ",
          i.jsx(M, { children: "editable ·" }),
          " ",
          e.counts["preserved-opaque"],
          " ",
          i.jsx(M, { children: "preserved opaque ·" }),
          " ",
          e.counts.flattened,
          " ",
          i.jsx(M, { children: "flattened · " }),
          e.counts.dropped,
          " ",
          i.jsx(M, { children: "dropped from the derived copy." }),
        ],
      }),
      i.jsx("p", { children: e.note }),
      i.jsx("p", {
        children:
          e.phase === "export-result"
            ? "These are the checked outcomes for this generated file."
            : "Opaque parts shown here are eligible for export. Export checks the current package and limits again before reporting the final outcome.",
      }),
      i.jsxs("div", {
        className: "lg-xlsx-filters",
        children: [
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Find a part, sheet or feature" }),
              i.jsx("input", {
                value: n,
                onChange: (f) => {
                  (o(f.target.value), u(0));
                },
                type: "search",
              }),
            ],
          }),
          i.jsxs("label", {
            children: [
              i.jsx(M, { children: "Export outcome" }),
              i.jsxs("select", {
                value: a,
                onChange: (f) => {
                  (s(f.target.value), u(0));
                },
                children: [
                  i.jsx("option", { value: "all", children: "All outcomes" }),
                  Object.entries(Tk).map(([f, p]) =>
                    i.jsx("option", { value: f, children: p }, f),
                  ),
                ],
              }),
            ],
          }),
        ],
      }),
      i.jsxs("table", {
        "aria-label": r("Workbook part inventory"),
        lang: t("Workbook part inventory"),
        "data-ui-label": "Workbook part inventory",
        children: [
          i.jsx("thead", {
            children: i.jsxs("tr", {
              children: [
                i.jsx("th", {
                  children: i.jsx(M, { children: "Part and source location" }),
                }),
                i.jsx("th", {
                  children: i.jsx(M, { children: "Derived copy" }),
                }),
                i.jsx("th", { children: i.jsx(M, { children: "Details" }) }),
              ],
            }),
          }),
          i.jsx("tbody", {
            children: l.slice(h * 25, h * 25 + 25).map((f) =>
              i.jsxs(
                "tr",
                {
                  children: [
                    i.jsxs("td", {
                      children: [
                        i.jsx("strong", { children: f.path }),
                        i.jsxs("small", {
                          children: [
                            f.family,
                            " · ",
                            f.bytes.toLocaleString(),
                            " ",
                            i.jsx(M, { children: "bytes" }),
                          ],
                        }),
                        f.locations.map((p, m) =>
                          i.jsxs(
                            "small",
                            {
                              children: [
                                p.sheet ??
                                  (p.scope === "unresolved"
                                    ? "Unresolved location"
                                    : "Workbook"),
                                p.range ? `!${p.range}` : "",
                                p.detail ? ` — ${p.detail}` : "",
                              ],
                            },
                            m,
                          ),
                        ),
                      ],
                    }),
                    i.jsxs("td", {
                      children: [
                        i.jsx("span", {
                          className: `lg-xlsx-outcome ${f.disposition}`,
                          children: Tk[f.disposition],
                        }),
                        i.jsx("small", {
                          children: i.jsx(M, {
                            children: "Retained in original",
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("td", {
                      children: [
                        f.reason,
                        i.jsxs("details", {
                          children: [
                            i.jsxs("summary", {
                              children: [
                                i.jsx(M, { children: "Part evidence" }),
                                f.features.length
                                  ? ` (${f.features.length} features)`
                                  : "",
                              ],
                            }),
                            i.jsx("p", { children: f.contentType }),
                            i.jsxs("p", {
                              children: [
                                i.jsx(M, { children: "SHA-256: " }),
                                i.jsx("code", { children: f.sha256 }),
                              ],
                            }),
                            f.features.slice(0, 200).map((p, m) =>
                              i.jsxs(
                                "p",
                                {
                                  children: [
                                    p.location.sheet
                                      ? `${p.location.sheet}!`
                                      : "",
                                    p.location.range ??
                                      p.location.detail ??
                                      "Workbook",
                                    " · ",
                                    p.feature,
                                    " · ",
                                    Tk[p.disposition],
                                    " —",
                                    " ",
                                    p.reason,
                                  ],
                                },
                                m,
                              ),
                            ),
                            f.features.length > 200 &&
                              i.jsxs("p", {
                                children: [
                                  i.jsx(M, {
                                    children:
                                      "Showing 200 features. Save the compatibility report for all ",
                                  }),
                                  f.features.length,
                                  " ",
                                  i.jsx(M, { children: "entries." }),
                                ],
                              }),
                          ],
                        }),
                      ],
                    }),
                  ],
                },
                f.path,
              ),
            ),
          }),
        ],
      }),
      !l.length &&
        i.jsx("p", {
          role: "status",
          children: i.jsx(M, { children: "No matching source parts." }),
        }),
      i.jsxs("div", {
        className: "lg-xlsx-pages",
        children: [
          i.jsx("button", {
            disabled: h === 0,
            onClick: () => u(h - 1),
            children: i.jsx(M, { children: "Previous parts" }),
          }),
          i.jsxs("span", {
            "aria-live": "polite",
            children: [
              l.length,
              " ",
              i.jsx(M, { children: "matching parts · Page " }),
              h + 1,
              " ",
              i.jsx(M, { children: "of " }),
              c + 1,
            ],
          }),
          i.jsx("button", {
            disabled: h === c,
            onClick: () => u(h + 1),
            children: i.jsx(M, { children: "Next parts" }),
          }),
        ],
      }),
    ],
  });
}
function X$({
  report: e,
  blocked: r,
  onConfirm: t,
  onClose: n,
  onReport: o,
  onOriginal: a,
  originalLabel: s = "Download original XLSX",
  inspectOnly: d = false,
}) {
  var m;
  const { text: u, language: l } = Dr(),
    c = ne.useId(),
    h = ne.useRef(null);
  ne.useEffect(() => {
    var g;
    const v = document.activeElement;
    return (
      (g = h.current) == null || g.focus(),
      () => {
        v != null && v.isConnected && v.focus();
      }
    );
  }, []);
  const f = e.direction === "import",
    p = e.caches;
  return i.jsx("div", {
    className: "lg-xlsx-overlay",
    children: i.jsxs("section", {
      ref: h,
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": c,
      tabIndex: -1,
      className: "lg-xlsx-review",
      onKeyDown: (v) => {
        var g;
        if (
          (v.key === "Escape" && (v.preventDefault(), v.stopPropagation(), n()),
          v.key === "Tab")
        ) {
          const b = [
              ...(((g = h.current) == null
                ? void 0
                : g.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),select:not(:disabled),a[href],summary,[tabindex="0"]',
                  )) ?? []),
            ],
            y = b[0],
            w = b.at(-1);
          v.shiftKey &&
          (document.activeElement === y || document.activeElement === h.current)
            ? (v.preventDefault(), w == null || w.focus())
            : !v.shiftKey &&
              document.activeElement === w &&
              (v.preventDefault(), y == null || y.focus());
        }
      },
      children: [
        i.jsx("h2", {
          id: c,
          children: d
            ? "Original XLSX and compatibility"
            : `Review ${e.format === "xltx" ? "XLTX" : "XLSX"} ${f ? "import" : "export"}`,
        }),
        i.jsxs("p", {
          children: [
            i.jsx("strong", { children: e.fileName }),
            " ·",
            " ",
            (e.fileBytes / 1024).toFixed(1),
            " ",
            i.jsx(M, { children: "KiB · " }),
            e.sheets.length,
            " ",
            i.jsx(M, { children: "worksheets in current conversion ·" }),
            " ",
            e.sheets.reduce((v, g) => v + g.storedCells, 0).toLocaleString(),
            " ",
            i.jsx(M, { children: "stored cells" }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(M, {
            children:
              "This is a derived copy with supported content. It is not a lossless Excel conversion. Keep the original file and a native JSON backup.",
          }),
        }),
        d
          ? i.jsx("p", {
              children: i.jsx(M, {
                children:
                  "This is the retained source snapshot. It does not include your subsequent edits. Source part locations refer to the original sheet names and addresses.",
              }),
            })
          : f
            ? i.jsxs("p", {
                children: [
                  i.jsx(M, {
                    children:
                      "Opening replaces the current workbook. Cancel keeps it unchanged; Undo can restore it after opening.",
                  }),
                  " ",
                  e.originalFileRetained
                    ? "The exact original is retained in this workbook and native JSON; access it later from the File menu."
                    : "No source bytes are attached to this report.",
                ],
              })
            : i.jsxs("p", {
                children: [
                  i.jsx(M, { children: "The download contains revision " }),
                  e.revision,
                  " ",
                  i.jsx(M, { children: "of" }),
                  " ",
                  e.workbookName,
                  i.jsx(M, {
                    children:
                      ". Unsupported native features remain in your current workbook and JSON backup.",
                  }),
                ],
              }),
        e.originalSource &&
          i.jsxs("p", {
            className: "lg-xlsx-source",
            children: [
              i.jsx(M, {
                children: e.alternateOriginal
                  ? "Converted XLSX retained: "
                  : "Original retained: ",
              }),
              i.jsx("strong", { children: e.originalSource.fileName }),
              " ·",
              " ",
              e.originalSource.byteLength.toLocaleString(),
              " ",
              i.jsx(M, { children: "bytes" }),
              i.jsx("br", {}),
              i.jsxs("small", {
                children: [
                  i.jsx(M, { children: "SHA-256: " }),
                  i.jsx("code", { children: e.originalSource.sha256 }),
                ],
              }),
            ],
          }),
        e.alternateOriginal &&
          i.jsxs("p", {
            className: "lg-xlsx-source",
            children: [
              i.jsxs("strong", {
                children: [
                  "Original ",
                  e.alternateOriginal.format.toUpperCase(),
                  ": ",
                  e.alternateOriginal.fileName,
                ],
              }),
              " · ",
              e.alternateOriginal.byteLength.toLocaleString(),
              " bytes",
              i.jsx("br", {}),
              i.jsxs("small", {
                children: [
                  "SHA-256: ",
                  i.jsx("code", { children: e.alternateOriginal.sha256 }),
                ],
              }),
            ],
          }),
        i.jsxs("div", {
          className: "lg-xlsx-scroll",
          children: [
            e.partInventory &&
              i.jsx(
                Pue,
                { inventory: e.partInventory },
                e.partInventory.sourceSha256,
              ),
            !d &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("h3", {
                    children: i.jsx(M, { children: "Worksheets" }),
                  }),
                  i.jsx("ul", {
                    children: e.sheets.map((v) =>
                      i.jsxs(
                        "li",
                        {
                          children: [
                            v.name,
                            " — ",
                            v.storedCells.toLocaleString(),
                            " ",
                            i.jsx(M, { children: "cells" }),
                            v.hidden ? " · hidden" : "",
                          ],
                        },
                        v.name,
                      ),
                    ),
                  }),
                  i.jsx("h3", {
                    children: i.jsx(M, { children: "Formula results" }),
                  }),
                  f
                    ? i.jsxs("p", {
                        children: [
                          p.pending ? `${p.pending} pending comparisons; ` : "",
                          p.total,
                          " ",
                          i.jsx(M, { children: "formulas: " }),
                          p.matched,
                          " ",
                          i.jsx(M, { children: "caches match, " }),
                          p.different,
                          " ",
                          i.jsx(M, { children: "differ, " }),
                          p.missing,
                          " ",
                          i.jsx(M, { children: "missing, " }),
                          p.neutralized,
                          " ",
                          i.jsx(M, {
                            children:
                              "converted to text. Local recalculation is used; a saved source cache is historical evidence, not proof of correctness.",
                          }),
                        ],
                      })
                    : i.jsxs("p", {
                        children: [
                          p.written,
                          " ",
                          i.jsx(M, { children: "formula caches written; " }),
                          p.omitted,
                          " ",
                          i.jsx(M, { children: "omitted;" }),
                          " ",
                          p.neutralized,
                          " ",
                          i.jsx(M, {
                            children:
                              "formulas converted to text. Excel is asked to recalculate on open.",
                          }),
                        ],
                      }),
                  !!p.issues.length &&
                    i.jsxs("table", {
                      "aria-label": u("Formula cache differences"),
                      lang: l("Formula cache differences"),
                      "data-ui-label": "Formula cache differences",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Cell" }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, {
                                  children: "Saved result",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, {
                                  children: "Local result",
                                }),
                              }),
                              i.jsx("th", {
                                children: i.jsx(M, { children: "Status" }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: p.issues.map((v) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("td", {
                                    children: [v.sheet, "!", v.address],
                                  }),
                                  i.jsx("td", {
                                    children: v.source
                                      ? `${v.source.type}: ${String(v.source.value)}`
                                      : "No cache",
                                  }),
                                  i.jsxs("td", {
                                    children: [
                                      v.calculated.type,
                                      ": ",
                                      String(v.calculated.value),
                                    ],
                                  }),
                                  i.jsx("td", { children: v.status }),
                                ],
                              },
                              `${v.sheet}!${v.address}`,
                            ),
                          ),
                        }),
                      ],
                    }),
                  !!p.detailsOmitted &&
                    i.jsxs("p", {
                      children: [
                        p.detailsOmitted,
                        " ",
                        i.jsx(M, {
                          children:
                            "additional cache details omitted from the bounded report; counts include all formulas.",
                        }),
                      ],
                    }),
                ],
              }),
            i.jsx("h3", {
              children: i.jsx(M, { children: "Compatibility notices" }),
            }),
            i.jsxs("ul", {
              "aria-label": u("Compatibility notices"),
              lang: l("Compatibility notices"),
              "data-ui-label": "Compatibility notices",
              children: [
                (m = e.embeddedObjectLosses) != null && m.length
                  ? i.jsxs("li", {
                      children: [
                        i.jsxs("strong", {
                          children: [
                            e.embeddedObjectLosses.length,
                            " unsupported embedded-object item(s) omitted",
                          ],
                        }),
                        i.jsx("ul", {
                          children: e.embeddedObjectLosses.map((v, g) =>
                            i.jsxs(
                              "li",
                              {
                                children: [
                                  v.sourceFile,
                                  " · ",
                                  v.path,
                                  " · ",
                                  v.kind,
                                  " · dropped · never executed",
                                ],
                              },
                              g,
                            ),
                          ),
                        }),
                      ],
                    })
                  : null,
                e.warnings.map((v, g) => i.jsx("li", { children: v }, g)),
              ],
            }),
          ],
        }),
        r && i.jsx("p", { role: "alert", children: r }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", { onClick: n, children: d ? "Close" : "Cancel" }),
            e.originalFileRetained &&
              a &&
              i.jsx("button", {
                onClick: a,
                children: i.jsx(M, { children: s }),
              }),
            i.jsx("button", {
              onClick: o,
              children: i.jsx(M, { children: "Save compatibility report" }),
            }),
            !d &&
              i.jsx("button", {
                disabled: !!r,
                onClick: t,
                children: f
                  ? "Open derived workbook"
                  : `Download derived ${e.format === "xltx" ? "XLTX" : "XLSX"}`,
              }),
          ],
        }),
      ],
    }),
  });
}
const Uue = ({ filename: e, blob: r }) => {
  const t = URL.createObjectURL(r),
    n = document.createElement("a");
  ((n.href = t),
    (n.download = e),
    n.click(),
    setTimeout(() => URL.revokeObjectURL(t), 1e3));
};

export { Bue, Fue, Lue, Uue, X$ };
