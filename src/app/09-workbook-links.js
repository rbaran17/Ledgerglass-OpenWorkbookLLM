import { unzipSync } from "fflate";
import ne from "react";
import i from "react/jsx-runtime";
import { X } from "lucide-react";
import { string, number, array, literal, unknown, object } from "zod";
import { hr } from "./01-core-schemas-and-hashing.js";
import { Mt } from "./02-financial-functions.js";
import { $R, $a, Cd, WR, kC, rn, uh, ut, zR, zoe } from "./03-workbook-model.js";
import { DC, Kse, T$, Zse, nT } from "./05-add-ins-and-custom-functions.js";
import { H1, im } from "./07-xlsx-archive-reader.js";
import { F$, P$, W2, cT } from "./08-xlsx-export-and-preservation.js";

function Sue(e) {
  const r = WR(e),
    t = W2(r.bytes, cT()),
    n = P$("import", r.fileName, r.bytes.length, e, t);
  return ((n.partInventory = F$(unzipSync(r.bytes), r.sha256)), n);
}
const W$ = ne.createContext({}),
  FC = () => ne.useContext(W$);
function yh() {
  const e = FC(),
    r = ne.useRef(null),
    t = ne.useCallback(() => {
      var n;
      return (n = r.current) == null ? void 0 : n.abort();
    }, []);
  return (
    ne.useEffect(() => () => t(), [t]),
    {
      cancel: t,
      begin: () => {
        t();
        const n = new AbortController();
        return (
          (r.current = n),
          {
            ...e,
            signal: e.signal ? AbortSignal.any([n.signal, e.signal]) : n.signal,
          }
        );
      },
    }
  );
}
function Tue({ workbook: e, sources: r, onReview: t }) {
  const n = ne.useMemo(() => {
    try {
      const o = nT(e, { sources: r });
      return !o.links.length && !o.unsupported.length
        ? null
        : `${o.links.length} workbook link targets · ${o.links.filter((a) => a.status !== "matches-approved-snapshot").length} need source review · ${o.unsupported.length} unsupported. Displayed linked values are cached snapshots.`;
    } catch {
      return "Workbook link inventory needs review; links exceed the supported review limit.";
    }
  }, [e, r]);
  return n
    ? i.jsxs("div", {
        className: "lg-workbook-link-status",
        role: "status",
        children: [
          i.jsx("span", { children: n }),
          i.jsx("button", { onClick: t, children: "Review workbook links" }),
        ],
      })
    : null;
}
function Nue(e) {
  var ce, ge;
  const r = ne.useRef(e);
  r.current = e;
  const t = ne.useRef(null),
    n = yh(),
    o = ne.useRef(0),
    a = ne.useRef(true),
    [s, d] = ne.useState([]),
    [u, l] = ne.useState(""),
    [c, h] = ne.useState(""),
    [f, p] = ne.useState("refresh"),
    [m, v] = ne.useState(false),
    [g, b] = ne.useState(false),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(0),
    [j, z] = ne.useState(0),
    [V, D] = ne.useState(false),
    [B, _] = ne.useState(null),
    [C, T] = ne.useState(""),
    [N, R] = ne.useState([]),
    [L, $] = ne.useState([]),
    [U, G] = ne.useState(false),
    H = e.approvedSources ?? s,
    Q = ne.useMemo(() => {
      try {
        return { inventory: nT(e.workbook, { sources: H }), error: "" };
      } catch (se) {
        return { inventory: null, error: se.message };
      }
    }, [e.workbook, H]),
    X$1 = Q.inventory,
    J =
      X$1 == null
        ? void 0
        : X$1.links.find((se) => {
            var we;
            return (
              se.target ===
              (u || ((we = X$1.links[0]) == null ? void 0 : we.target))
            );
          }),
    q = (J == null ? void 0 : J.target) ?? "",
    W =
      (X$1 == null
        ? void 0
        : X$1.cells.filter((se) =>
            se.references.some((we) => we.target === q),
          )) ?? [],
    K = (se) => {
      var we;
      if (e.approvedSources !== void 0 && !e.onApprovedSourcesChange)
        throw Error("Source approvals are controlled by your host.");
      ((we = e.onApprovedSourcesChange) == null || we.call(e, se),
        e.approvedSources === void 0 && d(se),
        w(null));
    },
    te = () => ({
      sources: r.current.approvedSources ?? s,
      readOnly: r.current.readOnly,
      allowedSheetIds: r.current.allowedSheetIds,
    });
  (ne.useEffect(() => {
    var we;
    a.current = true;
    const se = document.activeElement;
    return (
      (we = t.current) == null || we.focus(),
      () => {
        ((a.current = false),
          o.current++,
          n.cancel(),
          se != null && se.isConnected && se.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      w(null);
    }, [e.workbook, e.approvedSources]));
  const ae = (se) => {
      (A(""), F(""));
      try {
        se();
      } catch (we) {
        A(we instanceof Error ? we.message : "Workbook link operation failed.");
      }
    },
    oe = () =>
      ae(() => {
        (w(
          Zse(
            e.workbook,
            {
              action: f,
              target: q,
              replacementTarget: c,
              activateLiteral: m,
              acknowledgeCached: g,
            },
            te(),
          ),
        ),
          O(0),
          z(0));
      }),
    de = () =>
      ae(() => {
        if (!y) return;
        const se = Kse(r.current.workbook, y, te());
        (r.current.onWorkbookChange(se, `${y.request.action} workbook links`),
          w(null),
          F(
            "Reviewed workbook link changes applied. Undo restores the prior workbook.",
          ));
      }),
    he = async (se) => {
      const we = ++o.current;
      (D(true), A(""), _(null), w(null));
      try {
        const Ee = await H1(se, n.begin());
        if (!a.current || we !== o.current) return;
        (_(Ee.workbook),
          T(q || `[${se.name}]`),
          R(
            Ee.workbook.sheets
              .filter((Ae) => !Ae.hidden && !Ae.veryHidden)
              .map((Ae) => Ae.id),
          ),
          $(Ee.warnings),
          G(!1));
      } catch (Ee) {
        a.current && we === o.current && A(Ee.message);
      } finally {
        a.current && we === o.current && D(false);
      }
    },
    Se =
      e.allowFileSources !== false &&
      (e.approvedSources === void 0 || !!e.onApprovedSourcesChange),
    le = () => {
      (w(null), A(""), F(""), O(0));
    };
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      className: "lg-reports-dialog lg-workbook-links",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": "Workbook links",
      tabIndex: -1,
      ref: t,
      onKeyDown: (se) => {
        var we;
        if (
          (se.key === "Escape" && (se.preventDefault(), e.onClose()),
          se.key === "Tab")
        ) {
          const Ee = Array.from(
              ((we = t.current) == null
                ? void 0
                : we.querySelectorAll(
                    'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
                  )) ?? [],
            ).filter((Te) => !Te.hidden),
            Ae = Ee[0],
            ye = Ee.at(-1);
          se.shiftKey &&
          (document.activeElement === Ae ||
            document.activeElement === t.current)
            ? (se.preventDefault(), ye == null || ye.focus())
            : !se.shiftKey &&
              document.activeElement === ye &&
              (se.preventDefault(), Ae == null || Ae.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", { children: "DATA & WORKBOOK CONNECTIONS" }),
                i.jsx("h2", { children: "Workbook links" }),
              ],
            }),
            i.jsx("button", {
              "aria-label": "Close workbook links",
              onClick: e.onClose,
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "lg-reports-scroll",
          children: [
            i.jsx("p", {
              className: "lg-reports-intro",
              children:
                "Inspect and review linked values before refreshing, changing a source or breaking links. Workbook paths are labels; this editor never fetches them.",
            }),
            (E || Q.error) &&
              i.jsx("p", {
                role: "alert",
                className: "lg-links-error",
                children: E || Q.error,
              }),
            k && i.jsx("p", { role: "status", children: k }),
            i.jsxs("div", {
              className: "lg-links-summary",
              children: [
                i.jsxs("strong", {
                  children: [
                    (X$1 == null ? void 0 : X$1.links.length) ?? 0,
                    " linked sources",
                  ],
                }),
                i.jsx("span", { children: "Manual refresh only" }),
                i.jsx("span", {
                  children:
                    "Native JSON retains cached values and revision evidence",
                }),
              ],
            }),
            X$1 != null && X$1.links.length
              ? i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("label", {
                      children: [
                        "Linked workbook",
                        i.jsx("select", {
                          "aria-label": "Linked workbook",
                          value: q,
                          onChange: (se) => {
                            (l(se.target.value), v(false), b(false), le());
                          },
                          children: X$1.links.map((se) =>
                            i.jsxs(
                              "option",
                              {
                                value: se.target,
                                children: [se.target, " · ", se.status],
                              },
                              se.target,
                            ),
                          ),
                        }),
                      ],
                    }),
                    i.jsxs("section", {
                      className: "lg-links-card",
                      children: [
                        i.jsx("h3", { children: "Source and cache" }),
                        i.jsxs("dl", {
                          children: [
                            i.jsx("dt", { children: "Target workbook" }),
                            i.jsx("dd", { children: q }),
                            (J == null ? void 0 : J.packageTarget) &&
                              i.jsxs(i.Fragment, {
                                children: [
                                  i.jsx("dt", {
                                    children: "Original XLSX target",
                                  }),
                                  i.jsx("dd", { children: J.packageTarget }),
                                ],
                              }),
                            i.jsx("dt", { children: "Status" }),
                            i.jsx("dd", {
                              children: J == null ? void 0 : J.status,
                            }),
                            i.jsx("dt", { children: "Last resolved revision" }),
                            i.jsx("dd", {
                              children:
                                (J == null ? void 0 : J.lastResolvedRevision) ??
                                "Never resolved",
                            }),
                            i.jsx("dt", { children: "Resolved at" }),
                            i.jsx("dd", {
                              children:
                                (J == null ? void 0 : J.lastResolvedAt) ??
                                "Not available",
                            }),
                            i.jsx("dt", { children: "Source snapshot" }),
                            i.jsx("dd", {
                              children:
                                (J == null ? void 0 : J.sourceName) ??
                                "No approved result",
                            }),
                          ],
                        }),
                        i.jsx("p", { children: X$1.cachePolicy }),
                        (J == null ? void 0 : J.cached) &&
                          J.status !== "matches-approved-snapshot" &&
                          i.jsx("p", {
                            className: "lg-links-warning",
                            children:
                              "The previous cached values remain in this workbook. They are not verified against a currently authorized source.",
                          }),
                      ],
                    }),
                    i.jsxs("section", {
                      children: [
                        i.jsx("h3", { children: "Linked cells and targets" }),
                        i.jsx("div", {
                          className: "lg-links-table",
                          children: i.jsxs("table", {
                            children: [
                              i.jsx("thead", {
                                children: i.jsxs("tr", {
                                  children: [
                                    i.jsx("th", { children: "Destination" }),
                                    i.jsx("th", {
                                      children: "Source worksheet / range",
                                    }),
                                    i.jsx("th", { children: "Formula" }),
                                    i.jsx("th", { children: "State" }),
                                  ],
                                }),
                              }),
                              i.jsx("tbody", {
                                children: W.slice(I * 20, I * 20 + 20).map(
                                  (se) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              se.sheetName,
                                              "!",
                                              se.address,
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: se.references
                                              .filter((we) => we.target === q)
                                              .map(
                                                (we) =>
                                                  `${we.sheet}!${we.range}`,
                                              )
                                              .join(", "),
                                          }),
                                          i.jsx("td", { children: se.formula }),
                                          i.jsx("td", {
                                            children: se.literal
                                              ? "Inert text — activation required"
                                              : "Formula",
                                          }),
                                        ],
                                      },
                                      se.sheetId + se.address,
                                    ),
                                ),
                              }),
                            ],
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "lg-links-actions",
                      children: [
                        i.jsx("button", {
                          disabled: I === 0,
                          onClick: () => O(I - 1),
                          children: "Previous linked cells",
                        }),
                        i.jsxs("span", {
                          children: [
                            "Page ",
                            I + 1,
                            " of ",
                            Math.max(1, Math.ceil(W.length / 20)),
                          ],
                        }),
                        i.jsx("button", {
                          disabled: (I + 1) * 20 >= W.length,
                          onClick: () => O(I + 1),
                          children: "Next linked cells",
                        }),
                      ],
                    }),
                    i.jsxs("section", {
                      className: "lg-links-card",
                      children: [
                        i.jsx("h3", { children: "Review a change" }),
                        i.jsxs("label", {
                          children: [
                            "Operation",
                            i.jsxs("select", {
                              "aria-label": "Link operation",
                              value: f,
                              onChange: (se) => {
                                (p(se.target.value), le());
                              },
                              children: [
                                i.jsx("option", {
                                  value: "refresh",
                                  children: "Resolve / refresh approved source",
                                }),
                                i.jsx("option", {
                                  value: "replace",
                                  children: "Replace source workbook",
                                }),
                                i.jsx("option", {
                                  value: "break",
                                  children: "Break links and keep values",
                                }),
                              ],
                            }),
                          ],
                        }),
                        f === "replace" &&
                          i.jsxs("label", {
                            children: [
                              "Replacement source",
                              i.jsxs("select", {
                                "aria-label": "Replacement source",
                                value: c,
                                onChange: (se) => {
                                  (h(se.target.value), le());
                                },
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "Choose an approved source",
                                  }),
                                  H.filter((se) => se.target !== q).map((se) =>
                                    i.jsxs(
                                      "option",
                                      {
                                        value: se.target,
                                        children: [
                                          se.target,
                                          " · ",
                                          se.workbook.name,
                                        ],
                                      },
                                      se.target,
                                    ),
                                  ),
                                ],
                              }),
                            ],
                          }),
                        W.some((se) => se.literal) &&
                          i.jsxs("label", {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: m,
                                onChange: (se) => {
                                  (v(se.target.checked), le());
                                },
                              }),
                              "Activate these inert formulas using only reviewed cached source values",
                            ],
                          }),
                        f === "break" &&
                          i.jsxs("label", {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: g,
                                onChange: (se) => {
                                  (b(se.target.checked), le());
                                },
                              }),
                              "I understand that complete formulas will be replaced with the previewed cached values",
                            ],
                          }),
                        i.jsx("button", {
                          className: "lg-reports-primary",
                          disabled: e.readOnly || V,
                          onClick: oe,
                          children:
                            f === "refresh"
                              ? "Preview refresh"
                              : f === "replace"
                                ? "Preview replacement"
                                : "Preview break links",
                        }),
                      ],
                    }),
                  ],
                })
              : i.jsx("p", {
                  children:
                    "No supported external workbook references were found. Static links such as =SUM('[Source.xlsx]Sheet1'!A1:A3) can be reviewed here.",
                }),
            !!(X$1 != null && X$1.unsupported.length) &&
              i.jsxs("section", {
                className: "lg-links-warning",
                children: [
                  i.jsx("h3", { children: "Unsupported external syntax" }),
                  X$1.unsupported.slice(0, 20).map((se) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          se.sheetName,
                          "!",
                          se.address,
                          ": ",
                          se.unsupported,
                        ],
                      },
                      se.sheetId + se.address,
                    ),
                  ),
                ],
              }),
            !!(
              (ge =
                (ce = e.workbook.workbookLinks) == null
                  ? void 0
                  : ce.packageTargets) != null && ge.length
            ) &&
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: "Retained XLSX target inventory",
                  }),
                  e.workbook.workbookLinks.packageTargets.map((se) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          se.key,
                          " → ",
                          se.target,
                          " (",
                          se.part,
                          "). Retained original metadata; never fetched.",
                        ],
                      },
                      se.key,
                    ),
                  ),
                ],
              }),
            i.jsxs("section", {
              className: "lg-links-card",
              children: [
                i.jsx("h3", { children: "Approved source snapshots" }),
                H.length
                  ? i.jsx("ul", {
                      children: H.map((se) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              se.target,
                              " · ",
                              se.workbook.name,
                              " · revision",
                              " ",
                              se.workbook.revision,
                              " · ",
                              se.allowedSheetIds.length,
                              " approved worksheets",
                              " ",
                              (!e.approvedSources ||
                                e.onApprovedSourcesChange) &&
                                i.jsxs("button", {
                                  onClick: () =>
                                    ae(() => K(H.filter((we) => we !== se))),
                                  children: ["Remove approval for ", se.target],
                                }),
                            ],
                          },
                          se.target,
                        ),
                      ),
                    })
                  : i.jsx("p", {
                      children:
                        "No source workbook is approved in this session.",
                    }),
                Se
                  ? i.jsxs("label", {
                      children: [
                        "Attach source workbook",
                        i.jsx("input", {
                          "aria-label": "Attach source workbook",
                          type: "file",
                          accept: ".xlsx,.xltx,.json,.csv,.tsv",
                          disabled: V,
                          onChange: (se) => {
                            var Ee;
                            const we =
                              (Ee = se.target.files) == null ? void 0 : Ee[0];
                            ((se.target.value = ""), we && he(we));
                          },
                        }),
                      ],
                    })
                  : i.jsx("p", {
                      children:
                        "Your host supplies source approvals. Ask the host application to attach an authorized snapshot.",
                    }),
                V &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("p", {
                        role: "status",
                        children: "Inspecting selected source…",
                      }),
                      i.jsx("button", {
                        onClick: () => {
                          (o.current++, n.cancel(), D(false), _(null));
                        },
                        children: "Cancel source inspection",
                      }),
                    ],
                  }),
                B &&
                  i.jsxs("div", {
                    className: "lg-links-pending",
                    children: [
                      i.jsxs("h4", { children: ["Approve ", B.name] }),
                      i.jsxs("label", {
                        children: [
                          "Exact formula qualifier",
                          i.jsx("input", {
                            "aria-label": "Source qualifier",
                            value: C,
                            onChange: (se) => T(se.target.value),
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        children:
                          "Match the selected reference exactly. A matching filename alone does not grant access.",
                      }),
                      B.sheets.map((se) =>
                        i.jsxs(
                          "label",
                          {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: N.includes(se.id),
                                disabled: !!(se.hidden || se.veryHidden),
                                onChange: (we) =>
                                  R(
                                    we.target.checked
                                      ? [...N, se.id]
                                      : N.filter((Ee) => Ee !== se.id),
                                  ),
                              }),
                              se.name,
                              se.hidden || se.veryHidden
                                ? " (hidden; host approval required)"
                                : "",
                            ],
                          },
                          se.id,
                        ),
                      ),
                      L.length > 0 &&
                        i.jsxs("details", {
                          children: [
                            i.jsxs("summary", {
                              children: [L.length, " import notices"],
                            }),
                            L.map((se, we) => i.jsx("p", { children: se }, we)),
                          ],
                        }),
                      i.jsxs("label", {
                        className: "lg-links-check",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: U,
                            onChange: (se) => G(se.target.checked),
                          }),
                          "I approve this source and the selected worksheets for explicit refresh",
                        ],
                      }),
                      i.jsx("button", {
                        disabled: !U || !N.length,
                        onClick: () =>
                          ae(() => {
                            const se = [
                              ...H.filter((we) => we.target !== C),
                              { target: C, workbook: B, allowedSheetIds: N },
                            ];
                            (nT(e.workbook, { sources: se }), K(se), _(null));
                          }),
                        children: "Approve selected source",
                      }),
                      i.jsx("button", {
                        onClick: () => _(null),
                        children: "Discard source",
                      }),
                    ],
                  }),
              ],
            }),
            y &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                role: "region",
                "aria-label": "Workbook link preview",
                children: [
                  i.jsx("h3", {
                    children:
                      y.request.action === "break"
                        ? "Values that will replace formulas"
                        : "Reviewed results",
                  }),
                  i.jsx("p", { children: y.warning }),
                  i.jsx("div", {
                    className: "lg-links-table",
                    children: i.jsxs("table", {
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", { children: "Cell" }),
                              i.jsx("th", { children: "Before" }),
                              i.jsx("th", { children: "After" }),
                              i.jsx("th", { children: "Calculated value" }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: y.cells
                            .slice(j * 20, j * 20 + 20)
                            .map((se) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [se.sheetName, "!", se.address],
                                    }),
                                    i.jsx("td", { children: se.before }),
                                    i.jsx("td", {
                                      children: String(se.after ?? "(blank)"),
                                    }),
                                    i.jsxs("td", {
                                      children: [
                                        se.result.type,
                                        " ·",
                                        " ",
                                        String(se.result.value ?? "(blank)"),
                                      ],
                                    }),
                                  ],
                                },
                                se.sheetId + se.address,
                              ),
                            ),
                        }),
                      ],
                    }),
                  }),
                  i.jsxs("div", {
                    className: "lg-links-actions",
                    children: [
                      i.jsx("button", {
                        disabled: j === 0,
                        onClick: () => z(j - 1),
                        children: "Previous preview cells",
                      }),
                      i.jsxs("span", {
                        children: [
                          y.cells.length,
                          " affected cells · preview page",
                          " ",
                          j + 1,
                          " of ",
                          Math.ceil(y.cells.length / 20),
                        ],
                      }),
                      i.jsx("button", {
                        disabled: (j + 1) * 20 >= y.cells.length,
                        onClick: () => z(j + 1),
                        children: "Next preview cells",
                      }),
                    ],
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: e.readOnly || V,
                    onClick: de,
                    children: "Apply reviewed link changes",
                  }),
                  i.jsx("button", {
                    onClick: () => w(null),
                    children: "Cancel link preview",
                  }),
                ],
              }),
            i.jsxs("p", {
              children: [
                "Limits: ",
                uh.sources,
                " sources,",
                " ",
                uh.cells.toLocaleString(),
                " cached source cells and ",
                uh.formulas,
                " destination formulas. Static A1 cells/ranges only. Linked arrays, external names, DDE and automatic/network refresh are unsupported. XLSX exports link formulas as inert text; native JSON retains the reviewed cache. Breaking links does not alter an attached original file.",
              ],
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: e.readOnly
                ? "Read-only: inspection remains available."
                : "Every mutation requires a preview and confirmation.",
            }),
            i.jsx("button", { onClick: e.onClose, children: "Close" }),
          ],
        }),
      ],
    }),
  });
}
function Rue(e, r) {
  return im("source", e, r);
}
function H$(e) {
  return { fileName: e.fileName, bytes: DC(e), sha256: e.sha256 };
}
function q$(e, r) {
  const t = DC(e);
  if (e.method !== "document-adapter")
    throw Error("JSON/XML use the native structured parser.");
  const n = object({
    schemaVersion: literal(1),
    sourceSha256: string(),
    extractor: string().min(1).max(200),
    pageCount: number().int().min(1).max(1e4),
    tables: array(unknown()).min(1).max(Cd.tables),
  })
    .strict()
    .parse(r);
  if (n.sourceSha256 !== e.sha256)
    throw Error("Extraction does not match the selected source hash.");
  if (["png", "jpeg"].includes(e.format) && n.pageCount !== 1)
    throw Error("An image has one source page.");
  const o = kC({
    ...e,
    extractor: n.extractor,
    pageCount: n.pageCount,
    tables: n.tables,
    warnings: [
      ...e.warnings,
      "Extraction values, confidence and regions are adapter assertions. Review the original; hashes do not prove extraction accuracy.",
    ],
  });
  for (const a of o.tables)
    for (const s of a.rows)
      for (const d of s)
        if (d.region.page > n.pageCount)
          throw Error("A source region exceeds the declared page count.");
  return (
    o.tables.length > 1 && o.tables.forEach((a) => (a.ambiguous = true)),
    T$(o, t, e.fileName),
    o
  );
}
async function Iue(e, r, t, n = {}) {
  var l, c, h;
  if (t !== true)
    throw Error(
      "Approve the host extractor before sharing the selected source.",
    );
  if (e.method !== "document-adapter")
    throw Error("Choose a document source for extraction.");
  const o = H$(e);
  if ((l = n.signal) != null && l.aborted)
    throw Error("Document extraction cancelled.");
  const a = n.timeoutMs ?? 3e4;
  if (!Number.isInteger(a) || a < 1 || a > 6e4)
    throw Error("Invalid document extraction time budget.");
  const s = new AbortController(),
    d = () => s.abort();
  (c = n.signal) == null || c.addEventListener("abort", d, { once: true });
  let u;
  try {
    return await new Promise((f, p) => {
      let m = !1;
      const v = (b, y) => {
          m ||
            ((m = !0),
            s.signal.removeEventListener("abort", g),
            s.abort(),
            b ? p(b) : f(y));
        },
        g = () => v(Error("Document extraction cancelled."));
      (s.signal.addEventListener("abort", g, { once: !0 }),
        (u = setTimeout(
          () => v(Error("Document extraction exceeded its time budget.")),
          a,
        )),
        Promise.resolve()
          .then(() =>
            r.extract({
              file: new File([new Uint8Array(o.bytes)], o.fileName),
              sha256: e.sha256,
              signal: s.signal,
            }),
          )
          .then(
            (b) => {
              if (!m)
                try {
                  v(void 0, q$(e, b));
                } catch {
                  v(
                    Error(
                      "Extraction has invalid types, confidence, regions or source identity.",
                    ),
                  );
                }
            },
            () =>
              v(
                Error("Document extraction failed. The workbook is unchanged."),
              ),
          ));
    });
  } finally {
    (clearTimeout(u),
      (h = n.signal) == null || h.removeEventListener("abort", d),
      s.abort());
  }
}
const dT = new WeakMap();
function _ue(e, r, t) {
  const n = DC(r);
  if (
    (kC(r),
    $a(e),
    typeof t.sheetName != "string" ||
      !t.sheetName.trim() ||
      t.sheetName.length > 31 ||
      /[\\/?*\[\]:\x00-\x1f]/.test(t.sheetName) ||
      e.sheets.some(
        (p) => p.name.toLowerCase() === t.sheetName.trim().toLowerCase(),
      ))
  )
    throw Error("Choose a unique worksheet name of 1–31 characters.");
  const o = r.tables.find((p) => p.id === t.tableId);
  if (!o) throw Error("Select a source table.");
  if ((o.ambiguous || r.tables.length > 1) && t.ambiguityReviewed !== true)
    throw Error("Review and confirm the ambiguous table selection.");
  if (!t.resolutions || typeof t.resolutions != "object")
    throw Error("Invalid source review.");
  const a = {},
    s = new Set();
  let d = 0,
    u = 0,
    l = 0;
  if (
    (o.columns.forEach((p, m) => {
      a[ut(m) + "1"] = { raw: p.name, literal: true, style: { bold: true } };
    }),
    o.rows.forEach((p, m) =>
      p.forEach((v, g) => {
        const b = `${m}:${g}`;
        s.add(b);
        const y = t.resolutions[b],
          w = zR(v, y);
        ((y == null ? void 0 : y.decision) === "exclude"
          ? u++
          : (y == null ? void 0 : y.decision) === "correct"
            ? l++
            : d++,
          (a[ut(g) + (m + 2)] = {
            raw: w,
            ...(typeof w == "string" ? { literal: true } : {}),
          }));
      }),
    ),
    Object.keys(t.resolutions).some((p) => !s.has(p)))
  )
    throw Error("A review decision refers to an unknown cell.");
  const c = {
      id: crypto.randomUUID(),
      name: t.sheetName.trim(),
      cells: a,
      rowCount: Mt.maxRows,
      colCount: Mt.maxCols,
    },
    h = {
      id: crypto.randomUUID(),
      sheetId: c.id,
      sheetName: c.name,
      importedAt: new Date().toISOString(),
      original: zoe(n, r.fileName, r.format),
      method: r.method,
      extractor: r.extractor,
      ...(r.pageCount ? { pageCount: r.pageCount } : {}),
      table: structuredClone(o),
      resolutions: structuredClone(t.resolutions),
      ambiguityReviewed: t.ambiguityReviewed,
    };
  rn({
    ...e,
    sheets: [...e.sheets, c],
    sourceImports: [...(e.sourceImports ?? []), h],
  });
  const f = {
    workbookHash: hr(e),
    inspection: r,
    review: structuredClone(t),
    record: h,
    sheet: c,
    accepted: d,
    excluded: u,
    corrected: l,
  };
  return (dT.set(f, hr(f)), f);
}
function Due(e, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (($a(e), dT.get(r) !== hr(r) || r.workbookHash !== hr(e)))
    throw Error("The workbook or review changed. Preview the import again.");
  DC(r.inspection);
  const n = rn({
    ...e,
    sheets: [...e.sheets, r.sheet],
    sourceImports: [...(e.sourceImports ?? []), r.record],
  });
  return (dT.delete(r), n);
}
function Oue(e, r) {
  var n;
  const t = (n = e.sourceImports) == null ? void 0 : n.find((o) => o.id === r);
  if (!t) throw Error("Source import is unavailable.");
  return {
    bytes: $R(t.original),
    fileName: t.original.fileName,
    sha256: t.original.sha256,
  };
}
function Mue(e, r, t, n) {
  var h, f;
  const o = (h = e.sourceImports) == null ? void 0 : h.find((p) => p.id === r),
    a = (f = o == null ? void 0 : o.table.rows[t]) == null ? void 0 : f[n];
  if (!o || !a) throw Error("Imported source cell is unavailable.");
  const s = ut(n) + (t + 2),
    d = e.sheets.find((p) => p.id === o.sheetId),
    u = d == null ? void 0 : d.cells[s],
    l = o.resolutions[`${t}:${n}`],
    c = zR(a, l);
  return {
    address: s,
    sheetId: o.sheetId,
    sourceHash: o.original.sha256,
    source: a,
    decision: (l == null ? void 0 : l.decision) ?? "accepted",
    importedValue: c,
    currentValue: (u == null ? void 0 : u.raw) ?? null,
    status: d
      ? (u == null ? void 0 : u.raw) === c &&
        (typeof c != "string" || u.literal === true)
        ? "matches-import"
        : "changed"
      : "missing-sheet",
  };
}
const jue = ({ filename: e, blob: r }) => {
  const t = URL.createObjectURL(r),
    n = document.createElement("a");
  ((n.href = t),
    (n.download = e),
    n.click(),
    setTimeout(() => URL.revokeObjectURL(t), 1e3));
};

export { Due, FC, H$, Iue, Mue, Nue, Oue, Rue, Sue, Tue, W$, _ue, jue, q$, yh };
