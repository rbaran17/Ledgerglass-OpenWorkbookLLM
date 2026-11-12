import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  Ku,
  LU,
  Ls,
  PU,
  Vn,
  Vu,
  _t,
  cn,
  ct,
  jw,
  pf,
  qE,
  rn,
  ut,
  y0,
} from "./03-workbook-model.js";
import { Dr, M, sI } from "./04-cell-input-paste-and-fill.js";
import { ow } from "./07-xlsx-archive-reader.js";
import { Jle, uT } from "./08-xlsx-export-and-preservation.js";
import { yh } from "./09-workbook-links.js";
import { m0 } from "./12-formula-dependency-inspector.js";
import {
  Dxe,
  Ixe,
  Mxe,
  Oxe,
  Rxe,
  XT,
  _xe,
  g_,
  jxe,
  m_,
  sj,
} from "./38-data-analysis-dialog.js";

function Fxe(e) {
  const { text } = Dr(),
    t = ne.useId(),
    n = ne.useRef(null),
    o = yh(),
    a = ne.useRef(e);
  a.current = e;
  const [s, d] = ne.useState("Publish"),
    [u, l] = ne.useState({
      name: "Report output",
      kind: "range",
      sheetId: e.sheetId,
      range: e.range,
      mode: "snapshot",
      population: "full",
    }),
    [c, h] = ne.useState(false),
    [f, p] = ne.useState(false),
    [m, v] = ne.useState(null),
    [g, b] = ne.useState(),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState(""),
    [V, D] = ne.useState(null),
    [B, _] = ne.useState(false),
    C = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      includeHidden: !!e.allowHidden && c,
    },
    T = JSON.stringify([C, e.filter]),
    N = !!m && y === e.workbook && E === T,
    R = e.workbook.sheets.find((ae) => ae.id === u.sheetId),
    L = e.workbook.reportOutputs?.find((ae) => ae.id === j),
    $ = L ? Dxe(e.workbook, L.id, C) : null,
    U =
      s === "Saved outputs"
        ? $?.status !== "unavailable"
          ? L?.snapshot
          : undefined
        : N
          ? m?.snapshot
          : undefined;
  ne.useEffect(() => {
    var de;
    const ae = document.activeElement;
    return (
      (de = n.current?.querySelector("button")) == null || de.focus(),
      () => {
        ae != null && ae.isConnected && ae.focus();
      }
    );
  }, []);
  const G = (ae) => {
      try {
        (F(""), O(""), ae());
      } catch (oe) {
        F(oe instanceof Error ? oe.message : String(oe));
      }
    },
    H = (ae) => {
      (l((oe) => ({ ...oe, ...ae })), b(undefined), v(null), F(""));
    },
    Q = (ae, oe) =>
      G(() => {
        var Se, le;
        const de = Rxe(e.workbook, ae, C, oe);
        (v(de),
          b(oe),
          l(de.request),
          w(e.workbook),
          A(T),
          d("Publish"),
          (le =
            (Se = n.current?.querySelector(".lg-reports-scroll")) == null
              ? undefined
              : Se.scrollTo) == null || le.call(Se, 0, 0));
      }),
    X$1 = async (ae, oe) => {
      (_(true), F(""));
      try {
        const de = a.current,
          he = () =>
            m_(a.current.workbook, oe, {
              readOnly: a.current.readOnly,
              allowedSheetIds: a.current.allowedSheetIds,
              includeHidden: !!a.current.allowHidden && c,
            });
        he();
        const Se =
          ae === "png"
            ? await jxe(oe)
            : (() => {
                const le =
                  ae === "svg" ? g_(oe) : ae === "html" ? Mxe(oe) : Oxe(oe);
                return {
                  filename: le.filename,
                  blob: new Blob([le.content], { type: le.mime }),
                };
              })();
        if ((he(), de.onDownload !== a.current.onDownload))
          throw Error("The report destination changed. Export again.");
        (await (a.current.onDownload ?? sj)(Se),
          O(`${ae.toUpperCase()} snapshot exported.`));
      } catch (de) {
        F(de.message);
      } finally {
        _(false);
      }
    },
    J = async (ae) => {
      (_(true), D(null), F(""));
      try {
        const oe = await ow(ae, o.begin());
        (uT(oe), D(oe));
      } catch (oe) {
        F(oe.message);
      } finally {
        _(false);
      }
    };
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      ref: n,
      className: "lg-reports-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": t,
      onKeyDown: (ae) => {
        var oe, de;
        if (
          (ae.key === "Escape" && (ae.stopPropagation(), e.onClose()),
          ae.key === "Tab")
        ) {
          const he = [
            ...n.current.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
            ),
          ];
          ae.shiftKey && document.activeElement === he[0]
            ? (ae.preventDefault(), (oe = he.at(-1)) == null || oe.focus())
            : !ae.shiftKey &&
              document.activeElement === he.at(-1) &&
              (ae.preventDefault(), (de = he[0]) == null || de.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", { children: "REPORTING" }),
                i.jsx("h2", {
                  id: t,
                  children: i.jsx(M, {
                    children: "Report outputs & embedded objects",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close report outputs"),
              onClick: e.onClose,
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsx("nav", {
          "aria-label": text("Report workflow"),
          children: ["Publish", "Saved outputs", "Object inventory"].map((ae) =>
            i.jsx(
              "button",
              {
                "aria-pressed": s === ae,
                onClick: () => {
                  (d(ae), F(""));
                },
                children: text(ae),
              },
              ae,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "lg-reports-scroll",
          children: [
            s === "Publish" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    className: "lg-reports-intro",
                    children:
                      "Create a fixed image or table for another report. A linked output retains its source definition and refreshes only after review.",
                  }),
                  i.jsxs("div", {
                    className: "lg-reports-fields",
                    children: [
                      i.jsxs("label", {
                        children: [
                          "Report name",
                          i.jsx("input", {
                            value: u.name,
                            onChange: (ae) => H({ name: ae.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Source worksheet",
                          i.jsx("select", {
                            value: u.sheetId,
                            onChange: (ae) => {
                              (H({
                                sheetId: ae.target.value,
                                kind: "range",
                                objectId: undefined,
                              }),
                                p(false));
                            },
                            children: e.workbook.sheets
                              .filter(
                                (ae) =>
                                  (!e.allowedSheetIds ||
                                    e.allowedSheetIds.includes(ae.id)) &&
                                  (C.includeHidden ||
                                    (!ae.hidden && !ae.veryHidden)),
                              )
                              .map((ae) =>
                                i.jsx(
                                  "option",
                                  { value: ae.id, children: ae.name },
                                  ae.id,
                                ),
                              ),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Output content",
                          i.jsxs("select", {
                            value: u.kind === "range" ? "range" : u.objectId,
                            onChange: (ae) => {
                              const oe = R?.objects?.find(
                                (he) => he.id === ae.target.value,
                              );
                              H(
                                oe
                                  ? {
                                      kind: "object",
                                      objectId: oe.id,
                                      range:
                                        oe.kind === "chart"
                                          ? oe.range
                                          : oe.anchor,
                                    }
                                  : { kind: "range", objectId: undefined },
                              );
                            },
                            children: [
                              i.jsx("option", {
                                value: "range",
                                children: "Cell range / table",
                              }),
                              R?.objects?.map((ae) =>
                                i.jsxs(
                                  "option",
                                  {
                                    value: ae.id,
                                    children: [ae.kind, ": ", ae.label],
                                  },
                                  ae.id,
                                ),
                              ),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Report range",
                          i.jsx("input", {
                            value: u.range,
                            disabled: u.kind === "object",
                            onChange: (ae) => H({ range: ae.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Reuse mode",
                          i.jsxs("select", {
                            value: u.mode,
                            onChange: (ae) => H({ mode: ae.target.value }),
                            children: [
                              i.jsx("option", {
                                value: "snapshot",
                                children: "Fixed snapshot",
                              }),
                              i.jsx("option", {
                                value: "linked",
                                children: "Linked output — reviewed refresh",
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Report population",
                          i.jsxs("select", {
                            value: u.population,
                            onChange: (ae) =>
                              H({ population: ae.target.value }),
                            children: [
                              i.jsx("option", {
                                value: "full",
                                children: "Full source population",
                              }),
                              i.jsx("option", {
                                value: "visible",
                                children: "Visible / matching rows",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "lg-reports-options",
                    children: [
                      e.filter &&
                        u.sheetId === e.sheetId &&
                        i.jsxs("label", {
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: f,
                              onChange: (ae) => {
                                (p(ae.target.checked),
                                  H({
                                    filter: ae.target.checked
                                      ? e.filter
                                      : undefined,
                                    ...(ae.target.checked
                                      ? { range: e.filter.range }
                                      : {}),
                                  }));
                              },
                            }),
                            "Use active population filter",
                          ],
                        }),
                      e.allowHidden &&
                        i.jsxs("label", {
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: c,
                              onChange: (ae) => {
                                (h(ae.target.checked), v(null));
                              },
                            }),
                            "Authorize hidden source cells",
                          ],
                        }),
                    ],
                  }),
                  i.jsx("p", {
                    className: "lg-reports-help",
                    children:
                      "Up to 500 rows, 20 columns and 2,000 cells. Hidden sources require explicit authorization, even when a visible-only output excludes their rows. Source formulas must be deterministic and fully traceable.",
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: B,
                    onClick: () =>
                      Q(
                        {
                          ...u,
                          filter: g
                            ? u.filter
                            : f
                              ? (e.filter ?? u.filter)
                              : undefined,
                        },
                        g,
                      ),
                    children: "Preview report output",
                  }),
                  m &&
                    !N &&
                    i.jsx("p", {
                      role: "alert",
                      children:
                        "The workbook or permissions changed. Preview again.",
                    }),
                  N &&
                    i.jsxs("section", {
                      className: "lg-reports-preview",
                      children: [
                        i.jsxs("div", {
                          className: "lg-reports-preview-title",
                          children: [
                            i.jsx("strong", {
                              children: m.refreshId
                                ? "Reviewed refresh"
                                : "Snapshot preview",
                            }),
                            i.jsxs("span", {
                              children: [
                                "Revision ",
                                m.snapshot.revision,
                                " ·",
                                " ",
                                m.snapshot.rows.length,
                                " rows ·",
                                " ",
                                m.snapshot.request.mode,
                              ],
                            }),
                          ],
                        }),
                        i.jsx("div", {
                          className: "lg-reports-visual",
                          children: i.jsx(XT, { snapshot: m.snapshot }),
                        }),
                      ],
                    }),
                ],
              }),
            s === "Saved outputs" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    className: "lg-reports-intro",
                    children:
                      "Snapshots stay fixed. Linked outputs retain the last published snapshot until you review and publish a refresh.",
                  }),
                  i.jsxs("label", {
                    className: "lg-reports-select",
                    children: [
                      "Saved report",
                      i.jsxs("select", {
                        value: j,
                        onChange: (ae) => z(ae.target.value),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Choose a saved output",
                          }),
                          e.workbook.reportOutputs?.map((ae) =>
                            i.jsxs(
                              "option",
                              {
                                value: ae.id,
                                children: [
                                  ae.snapshot.request.name,
                                  " · v",
                                  ae.version,
                                ],
                              },
                              ae.id,
                            ),
                          ),
                        ],
                      }),
                    ],
                  }),
                  L &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("p", {
                          role: "status",
                          className: `lg-reports-status ${$?.status}`,
                          children: [
                            $?.status.toUpperCase(),
                            " · ",
                            $?.message,
                          ],
                        }),
                        $?.status !== "unavailable" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("div", {
                                className: "lg-reports-actions",
                                children: [
                                  L.snapshot.request.mode === "linked" &&
                                    i.jsx("button", {
                                      disabled: B,
                                      onClick: () => {
                                        (p(!!L.snapshot.request.filter),
                                          Q(L.snapshot.request, L.id));
                                      },
                                      children: "Preview linked refresh",
                                    }),
                                  i.jsx("button", {
                                    disabled: e.readOnly || B,
                                    onClick: () =>
                                      G(() => {
                                        (e.onWorkbookChange(
                                          _xe(e.workbook, L.id, C),
                                          "Remove report output",
                                        ),
                                          z(""));
                                      }),
                                    children: "Remove saved output",
                                  }),
                                ],
                              }),
                              i.jsx("div", {
                                className: "lg-reports-visual",
                                children: i.jsx(XT, { snapshot: L.snapshot }),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              }),
            s === "Object inventory" &&
              i.jsxs(i.Fragment, {
                children: [
                  i.jsx("p", {
                    className: "lg-reports-intro",
                    children:
                      "Inspect an authorized Excel container for OLE, ActiveX and embedded-document references. Inventory is inert metadata; no object opens or executes, and external references are never fetched.",
                  }),
                  i.jsxs("label", {
                    className: "lg-reports-file",
                    children: [
                      "Inspect embedded objects",
                      i.jsx("input", {
                        type: "file",
                        disabled: B,
                        accept: ".xlsx,.xlsm,.xlam,.xltm,.xltx,.xls,.bin",
                        onChange: (ae) => {
                          const oe = ae.target.files?.[0];
                          (oe && J(oe), (ae.target.value = ""));
                        },
                      }),
                    ],
                  }),
                  V &&
                    i.jsxs("section", {
                      children: [
                        i.jsx("h3", { children: V.fileName }),
                        i.jsx("p", { children: V.sha256 }),
                        i.jsx(lj, { items: uT(V).items }),
                        i.jsx("button", {
                          disabled: e.readOnly || B,
                          onClick: () =>
                            G(() => {
                              (e.onWorkbookChange(
                                Jle(e.workbook, V, { readOnly: e.readOnly }),
                                "Retain embedded-object inventory",
                              ),
                                O(
                                  "Inert object inventory retained. Keep the original file separately.",
                                ),
                                D(null));
                            }),
                          children: "Retain inventory",
                        }),
                      ],
                    }),
                  e.workbook.embeddedObjectInventories?.map((ae) =>
                    i.jsxs(
                      "section",
                      {
                        children: [
                          i.jsx("h3", { children: ae.fileName }),
                          i.jsxs("p", {
                            children: [
                              "Inspected ",
                              ae.inspectedAt,
                              " · execution: never",
                            ],
                          }),
                          i.jsx(lj, { items: ae.items }),
                          i.jsx("button", {
                            disabled: B,
                            onClick: () =>
                              void (e.onDownload ?? sj)({
                                filename: "Embedded-object-inventory.json",
                                blob: new Blob([JSON.stringify(ae, null, 2)], {
                                  type: "application/json",
                                }),
                              }),
                            children: "Export inventory JSON",
                          }),
                        ],
                      },
                      ae.id,
                    ),
                  ),
                ],
              }),
            k &&
              i.jsx("p", {
                role: "alert",
                className: "lg-reports-error",
                children: k,
              }),
            I && i.jsx("p", { role: "status", children: I }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("div", {
              className: "lg-reports-actions",
              children:
                U &&
                ["svg", "png", "html", "json"].map((ae) =>
                  i.jsxs(
                    "button",
                    {
                      disabled: B,
                      onClick: () => void X$1(ae, U),
                      children: ["Export ", ae.toUpperCase()],
                    },
                    ae,
                  ),
                ),
            }),
            i.jsxs("div", {
              className: "lg-reports-actions",
              children: [
                e.readOnly && i.jsx("span", { children: "Read-only host" }),
                s === "Publish" &&
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: !N || e.readOnly || B,
                    onClick: () =>
                      G(() => {
                        const ae = !!(m != null && m.refreshId),
                          oe = Ixe(e.workbook, m, C);
                        (e.onWorkbookChange(
                          oe,
                          ae
                            ? "Refresh linked report output"
                            : "Publish report output",
                        ),
                          z(m?.refreshId ?? oe.reportOutputs.at(-1).id),
                          b(undefined),
                          v(null),
                          O(
                            ae
                              ? "Linked output refreshed. Previously exported snapshots are unchanged."
                              : "Report output published.",
                          ),
                          d("Saved outputs"));
                      }),
                    children:
                      m != null && m.refreshId
                        ? "Publish reviewed refresh"
                        : "Publish report output",
                  }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}
function lj({ items }) {
  return i.jsx("div", {
    className: "lg-reports-inventory",
    children: i.jsxs("table", {
      children: [
        i.jsx("thead", {
          children: i.jsxs("tr", {
            children: [
              i.jsx("th", { children: "Part / reference" }),
              i.jsx("th", { children: "Kind" }),
              i.jsx("th", { children: "Bytes" }),
              i.jsx("th", { children: "Execution" }),
              i.jsx("th", { children: "Disposition" }),
            ],
          }),
        }),
        i.jsx("tbody", {
          children: items.map((r, t) =>
            i.jsxs(
              "tr",
              {
                children: [
                  i.jsx("td", { children: r.path }),
                  i.jsx("td", { children: r.kind }),
                  i.jsx("td", { children: r.bytes }),
                  i.jsx("td", { children: r.execution }),
                  i.jsx("td", { children: r.note }),
                ],
              },
              t,
            ),
          ),
        }),
      ],
    }),
  });
}
const b2 = (e, r) => {
  const t = ct(e),
    n = _t(r);
  return (
    n.row >= t.firstRow &&
    n.row <= t.lastRow &&
    n.col >= t.firstCol &&
    n.col <= t.lastCol
  );
};
function xW(e, r) {
  if (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  if (
    r.allowedSheetIds !== undefined &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (t) => typeof t != "string" || !e.sheets.some((n) => n.id === t),
      ))
  )
    throw Error("What-if host scope contains an unknown worksheet.");
  return {
    includeHidden: !!r.includeHidden,
    allowedSheetIds: [
      ...new Set(r.allowedSheetIds ?? e.sheets.map((t) => t.id)),
    ].sort(),
  };
}
function GT(e, r, t) {
  const n = e.sheets.find((a) => a.id === r.sheetId);
  if (!n || !t.allowedSheetIds.includes(n.id)) return false;
  const o = _t(r.address);
  return (
    t.includeHidden ||
    (!n.hidden &&
      !n.veryHidden &&
      !cn(n, "row").some((a) => a.start <= o.row && o.row <= a.end) &&
      !cn(n, "column").some((a) => a.start <= o.col && o.col <= a.end))
  );
}
const Bxe = new Set([
  ...y0,
  "ISFORMULA",
  "INDEX",
  "CHOOSE",
  "INDIRECT",
  "OFFSET",
  "LET",
  "LAMBDA",
  "RAND",
  "RANDBETWEEN",
  "RANDARRAY",
  "NOW",
  "TODAY",
  "CELL",
  "INFO",
  "RTD",
  "WEBSERVICE",
  "STOCKHISTORY",
  "EVALUATE",
  "GET.CELL",
]);
function EW(e, r, t = {}) {
  const n = LU(r),
    o = xW(e, t),
    a = PU(n);
  for (const f of [...n.inputs, ...n.outputs])
    if (!GT(e, f, o))
      throw Error(
        "A what-if source is missing or outside the visible, authorized host context.",
      );
  const s = n.inputs.map((f, p) => {
      var y, w;
      const m = e.sheets.find((E) => E.id === f.sheetId),
        v = m.cells[f.address],
        g = _t(f.address);
      if (!v || !qE(v.raw))
        throw Error(
          "Changing cells must contain finite numeric inputs, not formulas, text, blanks or spill outputs.",
        );
      if ((y = m.protection) != null && y.enabled)
        throw Error(
          "A changing worksheet is protected. Unprotect it before analyzing scenarios.",
        );
      if ((w = m.merges) != null && w.some((E) => b2(E, f.address)))
        throw Error("Merged cells cannot be what-if inputs.");
      if (
        [...(m.pivots ?? []), ...(m.queries ?? [])].some((E) =>
          b2(E.outputRange, f.address),
        )
      )
        throw Error("Generated pivot/query results cannot be what-if inputs.");
      for (const E of m.tables ?? []) {
        if (!b2(E.range, f.address)) continue;
        const A = ct(E.range),
          k = E.columns[g.col - A.firstCol];
        if (
          g.row === A.firstRow ||
          (E.showTotals && g.row === A.lastRow) ||
          (k != null && k.calculatedFormula) ||
          Object.entries(m.cells).some(([F, I]) => {
            const O = _t(F);
            return (
              O.col === g.col &&
              O.row > A.firstRow &&
              O.row <= A.lastRow - (E.showTotals ? 1 : 0) &&
              !I.literal &&
              typeof I.raw == "string" &&
              I.raw.startsWith("=")
            );
          })
        )
          throw Error(
            "Table headers, totals and calculated columns cannot be what-if inputs.",
          );
      }
      const b = (m.validations ?? []).filter((E) => b2(E.range, f.address));
      if (
        [v.raw, ...a.map((E) => E.inputs[p])].some((E) =>
          b.some((A) => !jw(A, E)),
        )
      )
        throw Error(
          `An assumption for ${m.name}!${f.address} violates its data validation constraints. Baseline unchanged.`,
        );
      return v.raw;
    }),
    d = new Map();
  for (const f of n.outputs) {
    const p = e.sheets.find((y) => y.id === f.sheetId).cells[f.address];
    if (!p || p.literal || typeof p.raw != "string" || !p.raw.startsWith("="))
      throw Error(
        "Every result cell must contain an actual calculation formula.",
      );
    const m = m0(
        {
          ...e,
          sheets: e.sheets.map((y) =>
            y.veryHidden ? { ...y, hidden: true } : y,
          ),
        },
        {
          ...f,
          direction: "precedents",
          depth: 8,
          ...o,
          limits: {
            maxFormulaCells: 5e3,
            maxNodes: Ls.dependencyCells,
            maxReferences: 1e4,
          },
        },
      ),
      v = [m.target, ...m.nodes],
      g = new Set(v.map((y) => y.key));
    if (m.coverage.truncated || m.coverage.depthLimited)
      throw Error(
        "What-if analysis requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
      );
    const b = m.issues.find((y) => g.has(y.from));
    if (b || m.cycles.length)
      throw Error(
        `What-if analysis requires supported static dependencies: ${b?.message ?? "a cycle was found."}`,
      );
    for (const y of v) d.set(y.key, { sheetId: y.sheetId, address: y.address });
  }
  if (d.size > Ls.dependencyCells)
    throw Error(
      "What-if analysis supports at most 500 dependency cells across all results.",
    );
  if (n.inputs.some((f) => !d.has(Vu(f))))
    throw Error(
      "Each changing input must be a precedent of at least one selected formula result.",
    );
  const u = [...d.values()].map((f) => {
      if (!GT(e, f, o))
        throw Error(
          "A formula dependency is outside the visible, authorized host context.",
        );
      const p = e.sheets.find((v) => v.id === f.sheetId),
        m = p.cells[f.address];
      if (!m || m.raw === null)
        throw Error(
          "What-if analysis requires populated static dependencies; blank or possible spill-child cells are unsupported.",
        );
      if (!m.literal && typeof m.raw == "string" && m.raw.startsWith("=")) {
        const v = pf(m.raw, e, p.id).replace(
          /"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g,
          "",
        );
        if (
          [...v.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)]
            .map((b) =>
              b[1]
                .toUpperCase()
                .replace(/^_XLFN\./, "")
                .replace(/^_XLWS\./, ""),
            )
            .some((b) => Bxe.has(b)) ||
          /[{}#]/.test(v)
        )
          throw Error(
            "What-if analysis excludes volatile, dynamic-reference, array/spill and unsupported reusable-function dependencies.",
          );
      }
      return {
        ...f,
        sheetName: p.name,
        raw: m.raw,
        ...(m.literal === undefined ? {} : { literal: m.literal }),
      };
    }),
    l = d.size * (a.length + 1);
  if (l > Ls.workUnits)
    throw Error(
      "What-if evaluation would exceed 50,000 dependency-cell work units. Reduce the values, scenarios or dependencies.",
    );
  const c = {
    ...e,
    sheets: e.sheets
      .filter((f) => o.allowedSheetIds.includes(f.id))
      .map((f) => ({
        ...f,
        cells: Object.fromEntries(
          Object.entries(f.cells).filter(([p]) =>
            d.has(Vu({ sheetId: f.id, address: p })),
          ),
        ),
      })),
  };
  function h(f, p) {
    try {
      const m = new Map(n.inputs.map((y, w) => [Vu(y), p[w]])),
        v = {
          ...c,
          sheets: c.sheets.map((y) => ({
            ...y,
            cells: Object.fromEntries(
              Object.entries(y.cells).map(([w, E]) => {
                const A = m.get(Vu({ sheetId: y.id, address: w }));
                return [w, A === undefined ? E : { ...E, raw: A }];
              }),
            ),
          })),
        },
        g = Vn(v),
        b = n.outputs.map((y) => g(y.sheetId, y.address));
      if (g.getOperationalErrors().length)
        throw Error(
          `Operational calculation error: ${g.getOperationalErrors().join(", ")}. An error handler cannot conceal this failure.`,
        );
      if (b.some((y) => !qE(y)))
        throw Error(
          `Each formula must return a finite number; received ${b.map(String).join(", ")}.`,
        );
      for (const y of d.values())
        if (g.getSpillRange(y.sheetId, y.address))
          throw Error(
            "A trial produced an array/spill; scalar results are required.",
          );
      if (g.getOperationalErrors().length)
        throw Error(
          "Dependency inspection encountered an operational calculation error.",
        );
      return { name: f, inputs: [...p], outputs: b, error: null };
    } catch (m) {
      return {
        name: f,
        inputs: [...p],
        outputs: null,
        error: (m instanceof Error
          ? m.message
          : "Trial calculation failed."
        ).slice(0, 2e3),
      };
    }
  }
  return {
    snapshot: JSON.stringify(e),
    requestSnapshot: JSON.stringify(n),
    request: n,
    context: o,
    baseline: h("Baseline", s),
    trials: a.map((f) => h(f.name, f.inputs)),
    assumptions: u,
    sourceRevision: e.revision,
    dependencyCells: d.size,
    workUnits: l,
  };
}
function CW(e, r, t) {
  if (JSON.stringify(e) !== r.snapshot)
    throw Error(
      "The workbook changed after preview. Preview and review the analysis again.",
    );
  const n = EW(e, r.request, t);
  if (JSON.stringify(n) !== JSON.stringify(r))
    throw Error(
      "The what-if preview or host context changed. Preview and review again.",
    );
  return n;
}
function AW(e, r, t = {}) {
  const { snapshot, requestSnapshot, context, ...s } = CW(e, r, t);
  return s;
}
function Lxe(e, r, t, n = {}) {
  if (n.readOnly) throw Error("This workbook is read-only.");
  const o = CW(e, r, n);
  if (o.request.kind !== "scenarios")
    throw Error(
      "Sensitivity tables are reports. Only named scenarios can be applied.",
    );
  if (
    t !== "baseline" &&
    (!Number.isInteger(t) || t < 0 || t >= o.trials.length)
  )
    throw Error("Choose one reviewed scenario.");
  const a = t === "baseline" ? o.baseline : o.trials[t];
  if (a.error || !a.outputs)
    throw Error(
      "A failed calculation cannot be applied. The baseline is unchanged.",
    );
  return a.inputs.every((s, d) => s === o.baseline.inputs[d])
    ? e
    : {
        ...e,
        sheets: e.sheets.map((s) => {
          const d = { ...s.cells };
          return (
            o.request.inputs.forEach((u, l) => {
              u.sheetId === s.id &&
                (d[u.address] = { ...d[u.address], raw: a.inputs[l] });
            }),
            { ...s, cells: d }
          );
        }),
      };
}
const hS = (e, r) => {
  return `${e.assumptions.find((n) => Vu(n) === Vu(r))?.sheetName ?? r.sheetId}!${r.address}`;
};
function kW(e) {
  const r = e.request.inputs.map((a) => hS(e, a)),
    t = e.request.outputs.map((a) => hS(e, a)),
    n = [
      [e.request.name],
      ["Historical what-if report", e.request.kind],
      ["Source revision", e.sourceRevision],
      [
        "Refresh explicitly from What-if Analysis; existing reports retain their original assumptions.",
      ],
    ],
    o = (a) => a.outputs ?? e.request.outputs.map(() => `Error: ${a.error}`);
  if (
    (n.push(
      ["Baseline", ...r, ...t],
      [e.baseline.name, ...e.baseline.inputs, ...o(e.baseline)],
      [],
    ),
    e.request.kind === "scenarios")
  ) {
    n.push(["Scenario", ...r, ...t]);
    for (const a of e.trials) n.push([a.name, ...a.inputs, ...o(a)]);
  } else if (e.request.kind === "one-variable")
    if (e.request.orientation === "column") {
      n.push([r[0], ...t]);
      for (const a of e.trials) n.push([a.inputs[0], ...o(a)]);
    } else
      (n.push([r[0], ...e.request.values]),
        t.forEach((a, s) => n.push([a, ...e.trials.map((d) => o(d)[s])])));
  else {
    const { rowValues, columnValues } = e.request;
    (n.push([`Down: ${r[1]} / Across: ${r[0]}`, ...rowValues]),
      columnValues.forEach((d, u) =>
        n.push([
          d,
          ...rowValues.map((l, c) => o(e.trials[u * rowValues.length + c])[0]),
        ]),
      ));
  }
  return (
    n.push(
      [],
      [
        "Source assumptions at preview",
        "Original value or formula",
        "Literal text",
      ],
      ...e.assumptions.map((a) => [hS(e, a), a.raw, a.literal ?? false]),
      [],
      ["Dependency cells", e.dependencyCells],
      ["Dependency-cell work units", e.workUnits],
    ),
    n
  );
}
function Pxe(e, r, t = {}, n) {
  var f;
  if (t.readOnly) throw Error("This workbook is read-only.");
  if ((f = e.structureProtection) != null && f.enabled)
    throw Error(
      "Workbook structure is protected. Unprotect it before saving a new report.",
    );
  const o = AW(e, r, t),
    a = e.whatIfAnalyses ?? [];
  if (n && !a.some((p) => p.id === n))
    throw Error("The saved analysis no longer exists.");
  if (
    n &&
    !v_(
      e,
      a.find((p) => p.id === n),
      t,
    )
  )
    throw Error("The saved analysis is outside the authorized context.");
  if (!n && a.length >= Ls.savedAnalyses)
    throw Error(
      "At most 40 saved analyses are supported. Remove an unused definition first.",
    );
  if (
    a.some(
      (p) =>
        p.id !== n &&
        p.request.name.toLowerCase() === o.request.name.toLowerCase(),
    )
  )
    throw Error(
      "A saved analysis already has this name. Choose another name or update that analysis.",
    );
  const s =
      o.request.name
        .replace(/[\\/\[\]*?:]/g, " ")
        .slice(0, 25)
        .trim() || "What-if",
    d = new Set(e.sheets.map((p) => p.name.toLowerCase()));
  let u = s,
    l = 2;
  for (; d.has(u.toLowerCase());) {
    const p = ` ${l++}`;
    u = `${s.slice(0, 31 - p.length)}${p}`;
  }
  const c = Ku(u);
  for (const [p, m] of kW(o).entries())
    for (const [v, g] of m.entries())
      c.cells[`${ut(v)}${p + 1}`] = {
        raw: g,
        ...(typeof g == "string" ? { literal: true } : {}),
        ...(p === 0 || p === 4 || p === 7
          ? { style: { bold: true, fill: "#e2f0d9" } }
          : {}),
      };
  ((c.columnWidths = { 0: 300, 1: 240, 2: 180, 3: 160 }), (c.frozenRows = 4));
  const h = { id: n ?? crypto.randomUUID(), reportSheetId: c.id, ...o };
  return rn({
    ...e,
    sheets: [...e.sheets, c],
    whatIfAnalyses: n ? a.map((p) => (p.id === n ? h : p)) : [...a, h],
  });
}
function v_(e, r, t = {}) {
  const n = xW(e, t);
  return r.assumptions.every((o) =>
    e.sheets.some((a) => a.id === o.sheetId)
      ? GT(e, o, n)
      : t.allowedSheetIds === undefined,
  );
}
function Uxe(e, r, t = {}) {
  var o;
  if (t.readOnly) throw Error("This workbook is read-only.");
  if ((o = e.structureProtection) != null && o.enabled)
    throw Error("Workbook structure is protected.");
  const n = e.whatIfAnalyses?.find((s) => s.id === r);
  if (!n || !v_(e, n, t))
    throw Error(
      "The saved analysis is missing or outside the authorized context.",
    );
  return { ...e, whatIfAnalyses: e.whatIfAnalyses.filter((s) => s.id !== r) };
}
function y2(e, r) {
  const t = r === "de-DE" || r === "fr-FR",
    n = e.split(t ? /[;\n]/ : /[,;\n]/);
  try {
    return n.map((o) => sI(o, "Assumption", r));
  } catch {
    throw Error(
      t
        ? "Enter finite numbers separated by semicolons or new lines; use a comma for decimals. Empty assumptions are not zero."
        : "Enter comma-separated finite numbers; empty assumptions are not zero.",
    );
  }
}

export { AW, EW, Fxe, Lxe, Pxe, Uxe, kW, v_, y2 };
