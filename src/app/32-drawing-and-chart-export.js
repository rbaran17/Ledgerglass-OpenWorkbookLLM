import i from "react/jsx-runtime";
import ne from "react";
import { X, ChartColumn, ImagePlus, Shapes, Trash2, Move } from "lucide-react";
import c_ from "react-dom/server";
import { toCanvas } from "html-to-image";
import { PDFDocument } from "pdf-lib";
import {
  BE,
  Dw,
  F2,
  Fw,
  I5,
  Ip,
  KR,
  Nr,
  P1,
  Qne,
  RN,
  Rx,
  Vn,
  Ww,
  XE,
  XR,
  ZR,
  aR,
  cn,
  ct,
  ek,
  gg,
  rae,
  sR,
  ul,
  ut,
  vx,
  wg,
  yC,
  zU,
} from "./03-workbook-model.js";
import { Dr, M, ci } from "./04-cell-input-paste-and-fill.js";
import { lle } from "./07-xlsx-archive-reader.js";
import { yh } from "./09-workbook-links.js";
import { jI } from "./22-themes-and-cell-styles.js";
import { Fz } from "./23-format-cells-dialog.js";
import { LI, wfe, xfe, yfe } from "./31-pivot-explore-and-slicers.js";

function Efe({
  report: e,
  error: r,
  revision: t,
  workbookId: n,
  onExport: o = wfe,
}) {
  const { text: a, language: s } = Dr(),
    [d, u] = ne.useState(0),
    [l, c] = ne.useState(""),
    h = ne.useRef(null),
    f = (e == null ? void 0 : e.points) ?? [],
    p = Math.min(d, Math.max(0, Math.ceil(f.length / 20) - 1)),
    m = (v = false) => {
      var g;
      try {
        if ((c(""), !e)) throw Error(r);
        if (v) {
          const b = (g = h.current) == null ? void 0 : g.querySelector("svg");
          if (!b)
            throw Error("Resolve chart errors before exporting the visual.");
          o({
            filename: "Chart.svg",
            mime: "image/svg+xml",
            content: new XMLSerializer().serializeToString(b),
          });
        } else
          o({
            filename: "Chart-inspection.json",
            mime: "application/json",
            content: JSON.stringify(
              { workbookId: n, revision: t, draft: !0, ...e },
              null,
              2,
            ),
          });
      } catch (b) {
        c(b instanceof Error ? b.message : String(b));
      }
    };
  return i.jsxs("section", {
    className: "dg-chart-inspection",
    "aria-label": a("Chart inspection"),
    lang: s("Chart inspection"),
    "data-ui-label": "Chart inspection",
    children: [
      i.jsx("h3", {
        children: i.jsx(M, { children: "Preview and source inspection" }),
      }),
      r && i.jsx("p", { className: "dg-error", children: r }),
      e &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsx("div", {
              ref: h,
              className:
                e.kind === "sparkline"
                  ? "dg-spark-preview"
                  : "dg-chart-preview",
              children: i.jsx(LI, { report: e }),
            }),
            i.jsx("p", { children: e.scope }),
            i.jsxs("p", {
              className: "dg-muted",
              children: [
                e.population.policy,
                " ",
                i.jsx(M, {
                  children:
                    "Exports contain this complete draft configuration and source inspection at revision ",
                }),
                t,
                i.jsx(M, { children: ". Source cells are unchanged." }),
              ],
            }),
            e.diagnostics.length > 0 &&
              i.jsx("ul", {
                className: "dg-chart-diagnostics",
                children: e.diagnostics.map((v, g) =>
                  i.jsxs("li", { children: [v.severity, ": ", v.message] }, g),
                ),
              }),
            i.jsx("button", {
              type: "button",
              onClick: () => m(),
              children: i.jsx(M, { children: "Export chart inspection" }),
            }),
            i.jsx("button", {
              type: "button",
              disabled: e.diagnostics.some((v) => v.severity === "error"),
              onClick: () => m(true),
              children: i.jsx(M, { children: "Export chart SVG" }),
            }),
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(M, {
                    children: "Inspect source data and full series",
                  }),
                }),
                i.jsxs("dl", {
                  children: [
                    i.jsx("dt", {
                      children: i.jsx(M, { children: "Title / alt text" }),
                    }),
                    i.jsxs("dd", { children: [e.title, " / ", e.altText] }),
                    i.jsx("dt", {
                      children: i.jsx(M, { children: "Axes and units" }),
                    }),
                    i.jsxs("dd", {
                      children: [
                        i.jsx(M, { children: "X: " }),
                        e.axes.x.title || "(automatic)",
                        " (",
                        e.axes.x.scale,
                        i.jsx(M, { children: "); Y: " }),
                        e.axes.y.title || "(automatic)",
                        " (",
                        e.axes.y.scale,
                        i.jsx(M, { children: "); units:" }),
                        " ",
                        e.axes.units || "not specified",
                      ],
                    }),
                    i.jsx("dt", { children: i.jsx(M, { children: "Series" }) }),
                    i.jsx("dd", {
                      children: e.series
                        .map(
                          (v) =>
                            `${v.name} (column ${v.column}, ${v.kind}, ${v.included ? "included" : "excluded"})`,
                        )
                        .join("; "),
                    }),
                  ],
                }),
                i.jsx("div", {
                  className: "dg-chart-data",
                  children: i.jsxs("table", {
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Source row" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Scope" }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(M, { children: "Category / X" }),
                            }),
                            e.series.map((v) =>
                              i.jsx("th", { children: v.name }, v.column),
                            ),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: f.slice(p * 20, p * 20 + 20).map((v) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("th", { children: v.row }),
                                i.jsxs("td", {
                                  children: [
                                    v.included ? "Included" : "Excluded",
                                    v.hidden ? " · hidden" : "",
                                  ],
                                }),
                                i.jsxs("td", {
                                  children: [
                                    v.category.address,
                                    ":",
                                    " ",
                                    String(v.category.value ?? "(blank)"),
                                  ],
                                }),
                                v.values.map((g) =>
                                  i.jsxs(
                                    "td",
                                    {
                                      children: [
                                        g.address,
                                        ": ",
                                        String(g.value ?? "(blank)"),
                                        g.raw !== g.value &&
                                          i.jsxs("small", {
                                            children: [
                                              " ",
                                              i.jsx(M, { children: "Raw: " }),
                                              String(g.raw),
                                            ],
                                          }),
                                        g.note &&
                                          i.jsxs("small", {
                                            children: [
                                              " ",
                                              i.jsx(M, {
                                                children: "Note: ",
                                              }),
                                              g.note,
                                            ],
                                          }),
                                      ],
                                    },
                                    g.address,
                                  ),
                                ),
                              ],
                            },
                            v.row,
                          ),
                        ),
                      }),
                    ],
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: !p,
                  onClick: () => u(p - 1),
                  children: i.jsx(M, {
                    children: "Previous chart source rows",
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: (p + 1) * 20 >= f.length,
                  onClick: () => u(p + 1),
                  children: i.jsx(M, { children: "Next chart source rows" }),
                }),
                e.histogram &&
                  i.jsxs("p", {
                    children: [
                      i.jsx(M, { children: "Bins:" }),
                      " ",
                      e.histogram
                        .map(
                          (v) =>
                            `${v.lower}–${v.upper}: ${v.count} rows ${v.upperInclusive ? "(upper inclusive)" : "(upper exclusive)"}`,
                        )
                        .join("; "),
                    ],
                  }),
                e.waterfall &&
                  i.jsxs("p", {
                    children: [
                      i.jsx(M, { children: "Waterfall:" }),
                      " ",
                      e.waterfall
                        .map((v) => `${v.label}: ${v.start} → ${v.end}`)
                        .join("; "),
                    ],
                  }),
              ],
            }),
          ],
        }),
      l && i.jsx("p", { role: "alert", className: "dg-error", children: l }),
    ],
  });
}
function Wz(e) {
  var _, C, T, N, R;
  const { text: r, language: t } = Dr(),
    n = yh(),
    o =
      (C =
        (_ = e.workbook.sheets.find((L) => L.id === e.sheetId)) == null
          ? void 0
          : _.objects) == null
        ? void 0
        : C.find((L) => L.id === e.objectId),
    a =
      (o == null ? void 0 : o.kind) ??
      (/picture|image/i.test(e.mode)
        ? "picture"
        : /chart/i.test(e.mode)
          ? "chart"
          : "shape"),
    s = e.workbook.sheets.find((L) => L.id === e.sheetId),
    [d] = ne.useState(() => ({ current: F2(e.workbook, e.sheetId, e.range) })),
    u = ne.useRef(e.mode),
    l = ne.useRef(e.objectId),
    c = ne.useRef(e),
    h = ne.useRef(null),
    f = ne.useRef(true),
    p = ne.useRef(0);
  (ne.useRef(null), (c.current = e));
  const [m, v] = ne.useState(() =>
      o
        ? structuredClone(o)
        : {
            ...Qne(a, e.range),
            ...(a === "chart"
              ? { chart: "column", range: e.range, headers: true }
              : {}),
          },
    ),
    [g, b] = ne.useState(
      ((T = o == null ? void 0 : o.evidence) == null
        ? void 0
        : T.sourceLocator) ?? "",
    ),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(false),
    k =
      d.current.guard !== JSON.stringify(e.workbook) ||
      d.current.sheetId !== e.sheetId ||
      d.current.selection !== e.range ||
      u.current !== e.mode ||
      l.current !== e.objectId,
    F =
      e.readOnly ||
      k ||
      !s ||
      !!((N = s == null ? void 0 : s.protection) != null && N.enabled),
    I = () => {
      const L = c.current;
      if (u.current !== L.mode || l.current !== L.objectId)
        throw Error("The drawing selection changed. Reopen this dialog.");
      BE(L.workbook, L.sheetId, {
        readOnly: L.readOnly,
        snapshot: d.current,
        selection: L.range,
      });
    },
    O = (L) => {
      !F && !E && v(($) => ({ ...$, ...L }));
    };
  (ne.useEffect(() => {
    var $, U;
    f.current = true;
    const L = document.activeElement;
    return (
      (U = ($ = h.current) == null ? void 0 : $.querySelector("button")) ==
        null || U.focus(),
      () => {
        ((f.current = false),
          p.current++,
          n.cancel(),
          L != null && L.isConnected && L.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      F && (p.current++, n.cancel(), A(false));
    }, [F]));
  async function j(L) {
    if (!L) return;
    const $ = ++p.current;
    try {
      if (
        (I(),
        A(!0),
        w(""),
        L.size > 2 * 1024 * 1024 ||
          L.size === 0 ||
          !["image/png", "image/jpeg"].includes(L.type))
      )
        throw Error("Choose a PNG or JPEG picture no larger than 2 MB.");
      const { dataURL: U } = await lle(L, n.begin());
      if (!f.current || $ !== p.current) return;
      I();
      const G = P1(U);
      if (G.bytes.length !== L.size)
        throw Error("The file bytes changed while reading the picture.");
      const H = ek(U, {
          origin: "local-file",
          filename: L.name,
          ...(g ? { sourceLocator: g } : {}),
        }),
        Q = await new Promise((J, q) => {
          const W = new Image();
          ((W.onload = () =>
            J({ width: W.naturalWidth, height: W.naturalHeight })),
            (W.onerror = () => q(Error("The picture could not be decoded."))),
            (W.src = U));
        });
      if (!f.current || $ !== p.current) return;
      if ((I(), Q.width !== G.width || Q.height !== G.height))
        throw Error(
          "Decoded picture dimensions do not match its validated header.",
        );
      const X = Math.min(1, 400 / G.width, 250 / G.height);
      v((J) => ({
        ...J,
        dataURL: U,
        evidence: H,
        label: L.name.slice(0, 500),
        width: Math.max(24, Math.round(G.width * X)),
        height: Math.max(24, Math.round(G.height * X)),
      }));
    } catch (U) {
      f.current &&
        $ === p.current &&
        w(U instanceof Error ? U.message : "Picture import failed.");
    } finally {
      f.current && $ === p.current && A(false);
    }
  }
  function z() {
    var L;
    try {
      if ((I(), E)) return;
      const $ = c.current,
        U = $.workbook.sheets.find((X) => X.id === $.sheetId);
      if (
        (a === "chart" && (sR(m.range, U), aR(Ww(U, m, Vn($.workbook)))),
        a === "picture" && !m.dataURL)
      )
        throw Error("Choose a picture first.");
      const G = structuredClone(m);
      if (
        (a === "picture" &&
          (G.evidence = {
            ...(m.evidence ?? ek(m.dataURL, { origin: "legacy-native" })),
            sourceLocator: g || void 0,
          }),
        e.objectId &&
          !((L = U.objects) != null && L.some((X) => X.id === e.objectId)))
      )
        throw Error("The drawing no longer exists.");
      const H = e.objectId
          ? (U.objects ?? []).map((X) => (X.id === e.objectId ? G : X))
          : [...(U.objects ?? []), G],
        Q = I5($.workbook, $.sheetId, H, {
          readOnly: $.readOnly,
          snapshot: d.current,
          selection: $.range,
        });
      (Q !== $.workbook &&
        $.onWorkbookChange(
          Q,
          e.objectId ? "Update " + a + " properties" : "Insert " + a,
        ),
        $.onClose());
    } catch ($) {
      w($ instanceof Error ? $.message : "The object could not be saved.");
    }
  }
  const V =
      m.evidence ??
      (m.dataURL ? ek(m.dataURL, { origin: "legacy-native" }) : void 0),
    D = ne.useMemo(() => {
      if (a !== "chart" || !s) return { report: null, error: "" };
      try {
        return { report: Ww(s, m, Vn(e.workbook)), error: "" };
      } catch (L) {
        return {
          report: null,
          error: L instanceof Error ? L.message : String(L),
        };
      }
    }, [a, s, m, e.workbook]),
    B = a[0].toUpperCase() + a.slice(1) + " properties";
  return i.jsx("div", {
    className: "dg-backdrop",
    onKeyDown: (L) => {
      var $;
      if (
        (L.stopPropagation(),
        L.key === "Escape" && (L.preventDefault(), e.onClose()),
        L.key === "Tab")
      ) {
        const U = [
            ...((($ = h.current) == null
              ? void 0
              : $.querySelectorAll(
                  "button:not(:disabled),input:not(:disabled),select:not(:disabled)",
                )) ?? []),
          ].filter((Q) => !Q.closest("fieldset:disabled")),
          G = U[0],
          H = U.at(-1);
        L.shiftKey && document.activeElement === G
          ? (L.preventDefault(), H == null || H.focus())
          : !L.shiftKey &&
            document.activeElement === H &&
            (L.preventDefault(), G == null || G.focus());
      }
    },
    children: i.jsxs("section", {
      className: `dg-dialog ${a === "chart" ? "dg-chart-dialog" : ""}`,
      ref: h,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": e.objectId ? B : `Insert ${a}`,
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("span", {
              role: "heading",
              "aria-level": 2,
              children: [
                a === "picture"
                  ? i.jsx(ImagePlus, { size: 19 })
                  : a === "chart"
                    ? i.jsx(ChartColumn, { size: 19 })
                    : i.jsx(Shapes, { size: 19 }),
                e.objectId ? B : "Insert " + a,
              ],
            }),
            i.jsx("button", {
              "aria-label": r("Close drawing tools"),
              onClick: e.onClose,
              lang: t("Close drawing tools"),
              "data-ui-label": "Close drawing tools",
              children: i.jsx(X, { size: 17 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "dg-form",
          children: [
            k &&
              i.jsx("p", {
                role: "alert",
                className: "dg-error",
                children: i.jsx(M, {
                  children:
                    "The workbook or selection changed. Close and reopen this dialog.",
                }),
              }),
            e.readOnly &&
              i.jsx("p", {
                role: "alert",
                className: "dg-error",
                children: i.jsx(M, {
                  children:
                    "This workbook is read-only. Object properties remain available to view.",
                }),
              }),
            ((R = s == null ? void 0 : s.protection) == null
              ? void 0
              : R.enabled) &&
              i.jsx("p", {
                role: "alert",
                className: "dg-error",
                children: i.jsx(M, {
                  children:
                    "Unprotect this worksheet before changing drawings.",
                }),
              }),
            i.jsxs("fieldset", {
              className: "dg-edit-fields",
              disabled: F || E,
              children: [
                a === "picture" &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsxs("label", {
                        className: "dg-upload",
                        children: [
                          i.jsx(M, { children: "Choose a local picture" }),
                          i.jsx("input", {
                            type: "file",
                            accept: "image/png,image/jpeg",
                            "aria-label": r("Picture file"),
                            disabled: F || E,
                            onChange: (L) => {
                              var $;
                              return void j(
                                ($ = L.target.files) == null ? void 0 : $[0],
                              );
                            },
                            lang: t("Picture file"),
                            "data-ui-label": "Picture file",
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        className: "dg-muted",
                        children: i.jsx(M, {
                          children:
                            "PNG or JPEG · up to 2 MB · stored with the workbook",
                        }),
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Source locator (plain text)" }),
                          i.jsx("input", {
                            "aria-label": r("Picture source locator"),
                            value: g,
                            maxLength: 2e3,
                            onChange: (L) => b(L.target.value),
                            lang: t("Picture source locator"),
                            "data-ui-label": "Picture source locator",
                          }),
                        ],
                      }),
                      m.dataURL &&
                        i.jsx("img", {
                          className: "dg-picture-preview",
                          src: m.dataURL,
                          alt: m.label,
                        }),
                    ],
                  }),
                a === "shape" &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(M, { children: "Shape" }),
                      i.jsxs("select", {
                        "aria-label": r("Shape type"),
                        value: m.shape,
                        onChange: (L) => O({ shape: L.target.value }),
                        lang: t("Shape type"),
                        "data-ui-label": "Shape type",
                        children: [
                          i.jsx("option", {
                            value: "rectangle",
                            children: "Rectangle",
                          }),
                          i.jsx("option", {
                            value: "ellipse",
                            children: "Ellipse",
                          }),
                          i.jsx("option", {
                            value: "arrow",
                            children: "Arrow",
                          }),
                          i.jsx("option", {
                            value: "text",
                            children: "Text box",
                          }),
                        ],
                      }),
                    ],
                  }),
                a === "chart" &&
                  i.jsx(xfe, { object: m, report: D.report, update: O }),
                i.jsxs("label", {
                  children: [
                    a === "chart" ? "Chart title" : "Label",
                    i.jsx("input", {
                      "aria-label": r("Object label"),
                      value: m.label,
                      maxLength: 500,
                      onChange: (L) => O({ label: L.target.value }),
                      lang: t("Object label"),
                      "data-ui-label": "Object label",
                    }),
                  ],
                }),
                m.chart !== "sparkline" &&
                  i.jsxs("div", {
                    className: "dg-two",
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Width (px)" }),
                          i.jsx(ci, {
                            min: 24,
                            max: 4e3,
                            "aria-label": r("Object width"),
                            value: m.width,
                            onValueChange: (L) => O({ width: Number(L) }),
                            lang: t("Object width"),
                            "data-ui-label": "Object width",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Height (px)" }),
                          i.jsx(ci, {
                            min: 24,
                            max: 4e3,
                            "aria-label": r("Object height"),
                            value: m.height,
                            onValueChange: (L) => O({ height: Number(L) }),
                            lang: t("Object height"),
                            "data-ui-label": "Object height",
                          }),
                        ],
                      }),
                    ],
                  }),
                a === "shape" &&
                  i.jsxs("div", {
                    className: "dg-two",
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Fill" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": r("Object fill"),
                            value: m.fill,
                            onChange: (L) => O({ fill: L.target.value }),
                            lang: t("Object fill"),
                            "data-ui-label": "Object fill",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(M, { children: "Outline" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": r("Object outline"),
                            value: m.stroke,
                            onChange: (L) => O({ stroke: L.target.value }),
                            lang: t("Object outline"),
                            "data-ui-label": "Object outline",
                          }),
                        ],
                      }),
                    ],
                  }),
                i.jsxs("p", {
                  className: "dg-muted",
                  children: [
                    i.jsx(M, { children: "Anchor: " }),
                    m.anchor,
                    i.jsx(M, {
                      children:
                        ". Select the object to move, resize, rename, or delete it.",
                    }),
                  ],
                }),
              ],
            }),
            a === "chart" &&
              i.jsx(Efe, {
                ...D,
                revision: e.workbook.revision,
                workbookId: e.workbook.id,
                onExport: e.onChartExport,
              }),
            a === "picture" &&
              V &&
              i.jsxs("dl", {
                className: "dg-evidence-properties",
                "aria-label": r("Picture evidence properties"),
                lang: t("Picture evidence properties"),
                "data-ui-label": "Picture evidence properties",
                children: [
                  i.jsx("dt", { children: i.jsx(M, { children: "Origin" }) }),
                  i.jsx("dd", {
                    children:
                      V.origin === "local-file"
                        ? "Local file bytes"
                        : V.origin === "xlsx-media"
                          ? "Observed XLSX media"
                          : "Observed embedded bytes; original source unknown",
                  }),
                  i.jsx("dt", {
                    children:
                      V.origin === "local-file"
                        ? "Original filename"
                        : "Observed filename",
                  }),
                  i.jsx("dd", { children: V.filename ?? "Not available" }),
                  i.jsx("dt", {
                    children: i.jsx(M, { children: "Byte length" }),
                  }),
                  i.jsx("dd", { children: V.byteLength.toLocaleString() }),
                  i.jsx("dt", { children: i.jsx(M, { children: "SHA-256" }) }),
                  i.jsx("dd", { children: V.sha256 }),
                  i.jsx("dt", {
                    children: i.jsx(M, { children: "Source reference" }),
                  }),
                  i.jsx("dd", { children: V.sourceReference }),
                  V.containerFilename &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(M, { children: "Imported workbook" }),
                        }),
                        i.jsx("dd", { children: V.containerFilename }),
                      ],
                    }),
                  V.mediaReference &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(M, {
                            children: "Observed media reference",
                          }),
                        }),
                        i.jsx("dd", { children: V.mediaReference }),
                      ],
                    }),
                  i.jsx("dt", {
                    children: i.jsx(M, { children: "Source locator" }),
                  }),
                  i.jsx("dd", { children: g || "Not recorded" }),
                ],
              }),
            y &&
              i.jsx("p", { role: "alert", className: "dg-error", children: y }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: e.onClose,
              children: i.jsx(M, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "dg-primary",
              onClick: z,
              disabled: F || E,
              children: E
                ? "Reading picture…"
                : e.objectId
                  ? "Save " + a + " properties"
                  : "Insert",
            }),
          ],
        }),
      ],
    }),
  });
}
function Hz({ object: e, sheet: r, calc: t, textScale: n = 1 }) {
  const o = ne.useMemo(() => {
    if (e.kind !== "picture") return "";
    try {
      return (P1(e.dataURL ?? ""), "");
    } catch (a) {
      return a instanceof Error ? a.message : "Invalid picture.";
    }
  }, [e.kind, e.dataURL]);
  return e.kind === "picture"
    ? o
      ? i.jsxs("span", {
          role: "img",
          "aria-label": "Invalid picture: " + o,
          children: [
            i.jsx(M, { children: "Picture cannot be displayed: " }),
            o,
          ],
        })
      : i.jsx("img", { src: e.dataURL, alt: e.label, draggable: false })
    : e.kind === "shape"
      ? i.jsxs("div", {
          role: "img",
          "aria-label": e.label,
          style: {
            position: "relative",
            display: "block",
            width: "100%",
            height: "100%",
            userSelect: "none",
          },
          children: [
            i.jsx("svg", {
              viewBox: "0 0 240 140",
              preserveAspectRatio: "none",
              "aria-hidden": "true",
              style: {
                position: "absolute",
                inset: 0,
                display: "block",
                width: "100%",
                height: "100%",
              },
              children: i.jsx("g", {
                fill: e.fill,
                stroke: e.stroke,
                strokeWidth: "2",
                children:
                  e.shape === "ellipse"
                    ? i.jsx("ellipse", {
                        cx: "120",
                        cy: "70",
                        rx: "116",
                        ry: "66",
                      })
                    : e.shape === "arrow"
                      ? i.jsx("path", { d: "M4 45H158V8L235 70 158 132V95H4Z" })
                      : e.shape !== "text"
                        ? i.jsx("rect", {
                            x: "2",
                            y: "2",
                            width: "236",
                            height: "136",
                            rx: "2",
                          })
                        : null,
              }),
            }),
            i.jsx("div", {
              className: "dg-shape-label",
              "aria-hidden": "true",
              style: {
                position: "absolute",
                inset: 8 * n,
                width: "auto",
                height: "auto",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                font: `${15 * n}px Arial, sans-serif`,
                color: e.stroke,
                whiteSpace: "pre-wrap",
                textAlign: "center",
                overflow: "hidden",
              },
              children: i.jsx("span", {
                style: {
                  minWidth: 0,
                  maxWidth: "100%",
                  overflowWrap: "anywhere",
                },
                children: e.label,
              }),
            }),
          ],
        })
      : i.jsx(yfe, { object: e, sheet: r, calc: t });
}
function Cfe(e) {
  var y, w;
  const { text: r, language: t } = Dr(),
    n = e.workbook.sheets.find((E) => E.id === e.sheetId),
    o = ne.useRef(e);
  o.current = e;
  const a =
      e.readOnly ||
      !n ||
      !!((y = n == null ? void 0 : n.protection) != null && y.enabled),
    [s, d] = ne.useState(null),
    [u, l] = ne.useState(""),
    [c, h] = ne.useState(null),
    [f, p] = ne.useState(null),
    m = ne.useRef(null);
  m.current = f;
  function v(E, A, k) {
    var F;
    try {
      const I = o.current,
        O = I.workbook.sheets.find((z) => z.id === I.sheetId);
      if (!(
        (F = O == null ? void 0 : O.objects) != null &&
        F.some((z) => z.id === E.id)
      ))
        throw Error("The selected drawing no longer exists.");
      const j = I5(
        I.workbook,
        I.sheetId,
        O.objects.map((z) => (z.id === E.id ? E : z)),
        {
          readOnly: I.readOnly,
          snapshot: k ?? F2(I.workbook, I.sheetId, I.selection),
          selection: I.selection,
        },
      );
      (j !== I.workbook && I.onWorkbookChange(j, A), l(""));
    } catch (I) {
      l(I instanceof Error ? I.message : "Could not update object.");
    }
  }
  function g(E) {
    try {
      const A = o.current,
        k = A.workbook.sheets.find((I) => I.id === A.sheetId),
        F = I5(
          A.workbook,
          A.sheetId,
          ((k == null ? void 0 : k.objects) ?? []).filter((I) => I.id !== E),
          {
            readOnly: A.readOnly,
            snapshot: F2(A.workbook, A.sheetId, A.selection),
            selection: A.selection,
          },
        );
      (F !== A.workbook && A.onWorkbookChange(F, "Delete drawing"), d(null));
    } catch (A) {
      l(A instanceof Error ? A.message : "Could not delete drawing.");
    }
  }
  function b(E, A, k, F) {
    try {
      if (
        (BE(e.workbook, e.sheetId, { readOnly: e.readOnly }),
        !Number.isFinite(e.scale) || e.scale <= 0)
      )
        throw Error("Invalid drawing zoom.");
      (p({
        id: E.id,
        type: A,
        x: k,
        y: F,
        initial: structuredClone(E),
        next: structuredClone(E),
        snapshot: F2(e.workbook, e.sheetId, e.selection),
        scale: e.scale,
      }),
        l(""));
    } catch (I) {
      l(I instanceof Error ? I.message : "Could not start drawing edit.");
    }
  }
  return (
    ne.useEffect(() => {
      (d(null), p(null), h(null));
    }, [e.sheetId]),
    ne.useEffect(() => {
      if (f)
        try {
          if (
            (BE(e.workbook, e.sheetId, {
              readOnly: e.readOnly,
              snapshot: f.snapshot,
              selection: e.selection,
            }),
            e.scale !== f.scale)
          )
            throw Error("The zoom changed. Start the drawing move again.");
        } catch (E) {
          (p(null),
            l(E instanceof Error ? E.message : "Drawing edit was cancelled."));
        }
    }, [e.workbook, e.readOnly, e.selection, e.scale, f]),
    ne.useEffect(() => {
      if (!f) return;
      const E = (I) => {
          const O = m.current;
          if (!O) return;
          const j = (I.clientX - O.x) / O.scale,
            z = (I.clientY - O.y) / O.scale,
            V = {
              ...O,
              next:
                O.type === "move"
                  ? {
                      ...O.initial,
                      offsetX: Math.max(
                        -4e3,
                        Math.min(4e3, O.initial.offsetX + j),
                      ),
                      offsetY: Math.max(
                        -4e3,
                        Math.min(4e3, O.initial.offsetY + z),
                      ),
                    }
                  : {
                      ...O.initial,
                      width: Math.max(24, Math.min(4e3, O.initial.width + j)),
                      height: Math.max(24, Math.min(4e3, O.initial.height + z)),
                    },
            };
          ((m.current = V), p(V));
        },
        A = () => {
          const I = m.current;
          I &&
            (v(
              I.next,
              I.type === "move" ? "Move drawing" : "Resize drawing",
              I.snapshot,
            ),
            (m.current = null),
            p(null));
        },
        k = () => {
          ((m.current = null), p(null));
        },
        F = (I) => {
          I.key === "Escape" && k();
        };
      return (
        window.addEventListener("pointermove", E),
        window.addEventListener("pointerup", A),
        window.addEventListener("pointercancel", k),
        window.addEventListener("keydown", F),
        () => {
          (window.removeEventListener("pointermove", E),
            window.removeEventListener("pointerup", A),
            window.removeEventListener("pointercancel", k),
            window.removeEventListener("keydown", F));
        }
      );
    }, [!!f]),
    i.jsxs(i.Fragment, {
      children: [
        i.jsxs("div", {
          className: "dg-drawing-overlay",
          role: "group",
          "aria-label": r("Worksheet drawings"),
          lang: t("Worksheet drawings"),
          "data-ui-label": "Worksheet drawings",
          children: [
            ((n == null ? void 0 : n.objects) ?? []).map((E) => {
              const A = (f == null ? void 0 : f.id) === E.id ? f.next : E,
                k = e.positionForCell(A.anchor),
                F = A.kind === "chart" && A.chart === "sparkline",
                I = XR(n, A);
              return k
                ? i.jsxs(
                    "div",
                    {
                      className: `dg-drawing ${s === A.id ? "dg-drawing-selected" : ""}`,
                      "data-object-id": A.id,
                      style: {
                        left: k.left + I.offsetX * e.scale,
                        top: k.top + I.offsetY * e.scale,
                        width: I.width * e.scale,
                        height: I.height * e.scale,
                      },
                      tabIndex: 0,
                      "aria-label": A.kind + ": " + A.label,
                      onClick: () => d(A.id),
                      onKeyDown: (O) => {
                        (O.stopPropagation(),
                          O.key === "Escape" && d(null),
                          (O.key === "Delete" || O.key === "Backspace") &&
                            O.target === O.currentTarget &&
                            !a &&
                            (O.preventDefault(), g(A.id)));
                      },
                      children: [
                        i.jsx(Hz, {
                          object: A,
                          sheet: n,
                          calc: e.calc,
                          textScale: e.scale,
                        }),
                        s === A.id &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("div", {
                                className: "dg-object-toolbar",
                                children: [
                                  (A.kind === "picture" ||
                                    A.kind === "chart") &&
                                    i.jsx("button", {
                                      className: "dg-properties-button",
                                      "aria-label":
                                        (A.kind === "chart"
                                          ? "Chart"
                                          : "Picture") +
                                        " properties: " +
                                        A.label,
                                      onClick: () => h(A.id),
                                      children: "ⓘ",
                                    }),
                                  !a &&
                                    i.jsxs(i.Fragment, {
                                      children: [
                                        i.jsx("button", {
                                          "aria-label": "Move " + A.label,
                                          title: F
                                            ? "Set the sparkline anchor in Chart properties"
                                            : "Drag to move",
                                          disabled: F,
                                          onPointerDown: (O) => {
                                            (O.preventDefault(),
                                              b(
                                                A,
                                                "move",
                                                O.clientX,
                                                O.clientY,
                                              ));
                                          },
                                          children: i.jsx(Move, { size: 14 }),
                                        }),
                                        i.jsx("input", {
                                          "aria-label": r(
                                            "Selected object label",
                                          ),
                                          value: A.label,
                                          onChange: (O) =>
                                            v(
                                              { ...A, label: O.target.value },
                                              "Rename drawing",
                                            ),
                                          lang: t("Selected object label"),
                                          "data-ui-label":
                                            "Selected object label",
                                        }),
                                        i.jsx("button", {
                                          "aria-label": "Delete " + A.label,
                                          onClick: () => g(A.id),
                                          children: i.jsx(Trash2, { size: 14 }),
                                        }),
                                      ],
                                    }),
                                ],
                              }),
                              !a &&
                                !F &&
                                i.jsx("button", {
                                  className: "dg-resize-object",
                                  "aria-label": "Resize " + A.label,
                                  onPointerDown: (O) => {
                                    (O.preventDefault(),
                                      b(A, "resize", O.clientX, O.clientY));
                                  },
                                }),
                            ],
                          }),
                      ],
                    },
                    A.id,
                  )
                : null;
            }),
            u &&
              i.jsx("div", {
                className: "dg-object-error",
                role: "alert",
                children: u,
              }),
          ],
        }),
        c &&
          ((w = n == null ? void 0 : n.objects) == null
            ? void 0
            : w.some((E) => E.id === c)) &&
          i.jsx(Wz, {
            mode: "drawing-properties",
            objectId: c,
            workbook: e.workbook,
            sheetId: e.sheetId,
            range: e.selection ?? n.objects.find((E) => E.id === c).anchor,
            readOnly: e.readOnly,
            onClose: () => h(null),
            onWorkbookChange: e.onWorkbookChange,
          }),
      ],
    })
  );
}
const qz =
  ".dg-print-paper{box-sizing:border-box;print-color-adjust:exact;-webkit-print-color-adjust:exact;color:#000;background:white;position:relative;overflow:hidden;break-after:page;page-break-after:always}.dg-print-paper:last-child{break-after:auto;page-break-after:auto}.dg-print-content{position:relative;overflow:hidden}.dg-print-table{border-collapse:collapse;table-layout:fixed;font-family:Calibri,Arial,sans-serif;font-size:14px}.dg-print-table td,.dg-print-table th{box-sizing:border-box;padding:0 3px;overflow:hidden;white-space:nowrap;font-weight:400}.dg-print-table th{font:11px Arial;background:#f3f3f3}.dg-print-table thead tr{height:22px}.dg-print-table td>div{overflow:hidden;line-height:1.35}.dg-print-header,.dg-print-footer{position:absolute;left:0;right:0;text-align:center;font:11px Arial;color:#444;white-space:pre-wrap;line-height:1.3;overflow-wrap:anywhere}.dg-print-object{position:absolute;overflow:hidden}.dg-print-object>img,.dg-print-object>svg{width:100%;height:100%;object-fit:contain}.dg-shape-label{width:100%;height:100%;display:flex;align-items:center;justify-content:center;font:15px Arial;white-space:pre-wrap;text-align:center;overflow:hidden}@media print{.dg-print-paper{margin:0!important;box-shadow:none!important;zoom:1!important}}";
function Xz({
  workbook: e,
  sheet: r,
  layout: t,
  page: n,
  index: o,
  total: a,
  scale: s,
  calc: d,
  ruleStyles: u,
}) {
  const l = ul(r.rowCount, cn(r, "row")),
    c = ul(r.colCount, cn(r, "column")),
    h = KR(t),
    f = ct(t.printArea),
    p =
      n.rows ??
      Array.from(
        { length: n.lastRow - n.firstRow + 1 },
        (E, A) => n.firstRow + A,
      ).filter((E) => !l.isHidden(E)),
    m =
      n.columns ??
      Array.from(
        { length: n.lastCol - n.firstCol + 1 },
        (E, A) => n.firstCol + A,
      ).filter((E) => !c.isHidden(E)),
    v = m.reduce((E, A) => E + Ip(r, A), t.headings ? 32 : 0),
    g = new Map(p.map((E, A) => [E, A])),
    b = [0];
  for (const E of p) b.push(b.at(-1) + gg(r, E));
  const y = new Map();
  for (const E of r.merges ?? []) {
    const A = ct(E),
      k = p.filter((O) => O >= A.firstRow && O <= A.lastRow),
      F = m.filter((O) => O >= A.firstCol && O <= A.lastCol);
    if (!k.length || !F.length) continue;
    const I = {
      firstRow: k[0],
      lastRow: k.at(-1),
      firstCol: F[0],
      lastCol: F.at(-1),
      rowSpan: k.length,
      colSpan: F.length,
      suppressContent: l.isHidden(A.firstRow) || c.isHidden(A.firstCol),
      anchor: `${ut(A.firstCol)}${A.firstRow + 1}`,
    };
    if (!(I.firstRow > I.lastRow || I.firstCol > I.lastCol))
      for (const O of k) for (const j of F) y.set(O * 16384 + j, I);
  }
  const w = (E) => rae(E, e, r, o + 1, a);
  return i.jsxs("article", {
    className: "dg-print-paper",
    "aria-label": `Print page ${o + 1}`,
    style: {
      boxSizing: "border-box",
      width: h.width,
      height: h.height,
      paddingTop: t.margins.top * 96,
      paddingBottom: t.margins.bottom * 96,
      paddingLeft: t.margins.left * 96,
      paddingRight: t.margins.right * 96,
    },
    children: [
      i.jsx("div", {
        className: "dg-print-header",
        style: {
          top: t.margins.header * 96,
          left: t.margins.left * 96,
          right: t.margins.right * 96,
        },
        children: w(t.header),
      }),
      i.jsx("div", {
        className: "dg-print-content",
        style: {
          width: h.width - (t.margins.left + t.margins.right) * 96,
          height: h.height - (t.margins.top + t.margins.bottom) * 96,
        },
        children: i.jsxs("div", {
          style: {
            transform: `scale(${s})`,
            transformOrigin: "top left",
            width: v,
          },
          children: [
            i.jsxs("table", {
              className: "dg-print-table",
              style: { width: v },
              children: [
                i.jsxs("colgroup", {
                  children: [
                    t.headings && i.jsx("col", { style: { width: 32 } }),
                    m.map((E) =>
                      i.jsx("col", { style: { width: Ip(r, E) } }, E),
                    ),
                  ],
                }),
                t.headings &&
                  i.jsx("thead", {
                    children: i.jsxs("tr", {
                      children: [
                        i.jsx("th", {}),
                        m.map((E) => i.jsx("th", { children: ut(E) }, E)),
                      ],
                    }),
                  }),
                i.jsx("tbody", {
                  children: p.map((E) =>
                    i.jsxs(
                      "tr",
                      {
                        style: { height: gg(r, E) },
                        children: [
                          t.headings && i.jsx("th", { children: E + 1 }),
                          m.map((A) => {
                            var V;
                            const k = y.get(E * 16384 + A);
                            if (k && (E !== k.firstRow || A !== k.firstCol))
                              return null;
                            const F =
                                (k == null ? void 0 : k.anchor) ??
                                `${ut(A)}${E + 1}`,
                              I =
                                k != null && k.suppressContent
                                  ? void 0
                                  : r.cells[F],
                              O = {
                                ...I,
                                raw: (I == null ? void 0 : I.raw) ?? null,
                                style:
                                  k != null && k.suppressContent
                                    ? {}
                                    : {
                                        ...(I == null ? void 0 : I.style),
                                        ...u(F).style,
                                      },
                              },
                              j =
                                k != null && k.suppressContent
                                  ? null
                                  : d(r.id, F),
                              z =
                                b[
                                  g.get((k == null ? void 0 : k.lastRow) ?? E) +
                                    1
                                ] - b[g.get(E)];
                            return i.jsx(
                              "td",
                              {
                                "data-print-cell": F,
                                "data-print-in-area":
                                  E >= f.firstRow &&
                                  E <= f.lastRow &&
                                  A >= f.firstCol &&
                                  A <= f.lastCol
                                    ? "true"
                                    : "false",
                                "data-print-number":
                                  typeof j == "number" ? "true" : void 0,
                                colSpan: (k == null ? void 0 : k.colSpan) ?? 1,
                                rowSpan: (k == null ? void 0 : k.rowSpan) ?? 1,
                                style: Dw(O.style, j, {
                                  defaultBorder: t.gridlines
                                    ? "1px solid #d4d4d4"
                                    : "none",
                                }),
                                children: i.jsx("div", {
                                  style: {
                                    maxHeight: z - 1,
                                    ...RN(O.style),
                                    position: "relative",
                                  },
                                  children: i.jsx(Fz, {
                                    decoration:
                                      k != null && k.suppressContent
                                        ? void 0
                                        : u(F),
                                    children: i.jsx(jI, {
                                      value: j,
                                      cell: O,
                                      dateSystem:
                                        (V = e.datePolicy) == null
                                          ? void 0
                                          : V.dateSystem,
                                    }),
                                  }),
                                }),
                              },
                              A,
                            );
                          }),
                        ],
                      },
                      E,
                    ),
                  ),
                }),
              ],
            }),
            (r.objects ?? []).map((E) => {
              const A = ct(E.anchor),
                k = XR(r, E);
              if (
                l.isHidden(A.firstRow) ||
                c.isHidden(A.firstCol) ||
                !p.includes(A.firstRow) ||
                !m.includes(A.firstCol)
              )
                return null;
              let F = t.headings ? 32 : 0,
                I = t.headings ? 22 : 0;
              for (const O of m.slice(0, m.indexOf(A.firstCol))) F += Ip(r, O);
              for (const O of p.slice(0, p.indexOf(A.firstRow))) I += gg(r, O);
              return i.jsx(
                "div",
                {
                  className: "dg-print-object",
                  "data-print-object": E.id,
                  "data-print-in-area":
                    A.firstRow >= f.firstRow &&
                    A.firstRow <= f.lastRow &&
                    A.firstCol >= f.firstCol &&
                    A.firstCol <= f.lastCol
                      ? "true"
                      : "false",
                  "data-print-label": E.label,
                  style: {
                    left: F + k.offsetX,
                    top: I + k.offsetY,
                    width: k.width,
                    height: k.height,
                  },
                  children: i.jsx(Hz, { object: E, sheet: r, calc: d }),
                },
                E.id,
              );
            }),
          ],
        }),
      }),
      i.jsx("div", {
        className: "dg-print-footer",
        style: {
          bottom: t.margins.footer * 96,
          left: t.margins.left * 96,
          right: t.margins.right * 96,
        },
        children: w(t.footer),
      }),
    ],
  });
}

function uW(e, r, t) {
  const n = { ...Nr(e, r) };
  Rx(t, n);
  try {
    yC(n.objects, n);
  } catch (d) {
    throw Error(
      `Unsupported or invalid print object: ${d instanceof Error ? d.message : "check the worksheet drawings."}`,
    );
  }
  const o = zU(n, n.pageLayout),
    a = Vn(e),
    s = Fw(n, a, vx(n, a));
  for (const d of o.pages) {
    const u =
        d.rows ??
        Array.from(
          { length: d.lastRow - d.firstRow + 1 },
          (c, h) => d.firstRow + h,
        ),
      l =
        d.columns ??
        Array.from(
          { length: d.lastCol - d.firstCol + 1 },
          (c, h) => d.firstCol + h,
        );
    for (const c of u) for (const h of l) s(`${ut(h)}${c + 1}`);
  }
  return { ...o, sheet: n, layout: n.pageLayout, calc: a, ruleStyles: s };
}
function cW(e, r, t) {
  var u;
  const n = Nr(e, r),
    o = uW(e, r, t ?? n.pageLayout ?? { ...ZR(n), ...XE }),
    a = o.pages.map((l, c) =>
      i.jsx(
        Xz,
        {
          workbook: e,
          sheet: o.sheet,
          layout: o.layout,
          page: l,
          index: c,
          total: o.pages.length,
          scale: o.scale,
          calc: o.calc,
          ruleStyles: o.ruleStyles,
        },
        c,
      ),
    ),
    s = o.layout.paper === "a4" ? "A4" : o.layout.paper;
  return {
    html:
      "<!doctype html>" +
      c_.renderToStaticMarkup(
        i.jsxs("html", {
          lang: "en",
          children: [
            i.jsxs("head", {
              children: [
                i.jsx("meta", { charSet: "utf-8" }),
                i.jsx("title", {
                  children: `${e.name} — ${n.name} — revision ${e.revision}`,
                }),
                i.jsx("meta", {
                  name: "description",
                  content: `Workbook ${e.name}; revision ${e.revision}; period ${((u = e.metadata) == null ? void 0 : u.period) || "not set"}; ${o.pages.length} page(s).`,
                }),
                i.jsx("style", {
                  children: `*{box-sizing:border-box}html,body{margin:0;padding:0;background:white;font-family:Calibri,Arial,sans-serif}${qz}@page{size:${s} ${o.layout.orientation};margin:0}`,
                }),
              ],
            }),
            i.jsx("body", { children: a }),
          ],
        }),
      ),
    pageCount: o.pages.length,
    notices: [
      ...(o.notices ?? []),
      ...o.ruleStyles
        .errors()
        .map((l) => `Conditional formatting: ${l.message}`),
    ],
  };
}
function zy(e, r, t = 0.8) {
  return (
    e.left < r.left - t ||
    e.right > r.right + t ||
    e.top < r.top - t ||
    e.bottom > r.bottom + t
  );
}
function cS(e) {
  var s;
  if (e.scrollWidth > e.clientWidth + 1 || e.scrollHeight > e.clientHeight + 1)
    return true;
  const r = e.ownerDocument,
    t = e.getBoundingClientRect(),
    n = r.createTreeWalker(e, 4),
    o = r.createRange();
  let a;
  for (; (a = n.nextNode());)
    if ((s = a.textContent) != null && s.trim()) {
      o.selectNodeContents(a);
      for (const d of o.getClientRects()) if (zy(d, t)) return true;
    }
  return false;
}
function mw(e) {
  const r = Array.from(e.querySelectorAll(".dg-print-paper")),
    t = [];
  r.length ||
    t.push({
      severity: "error",
      code: "empty",
      message: "There are no pages to print. Choose a populated print area.",
    });
  let n = false;
  return (
    r.forEach((o, a) => {
      var f, p;
      const s = a + 1,
        d = o.getBoundingClientRect(),
        u = o.querySelector(".dg-print-content"),
        l = o.querySelector(".dg-print-table");
      if (!d.width || !d.height || !u || !l) {
        t.push({
          severity: "error",
          code: "layout",
          page: s,
          message: `Page ${s}: print layout could not be measured. Keep the preview visible and retry in a supported browser.`,
        });
        return;
      }
      const c = u.getBoundingClientRect();
      zy(l.getBoundingClientRect(), c) &&
        t.push({
          severity: "error",
          code: "layout",
          page: s,
          message: `Page ${s}: rows or columns extend beyond the printable area. Reduce scale, increase margins' available space, or adjust row/column sizes.`,
        });
      for (const m of ["header", "footer"]) {
        const v = o.querySelector(`.dg-print-${m}`);
        if (!((f = v == null ? void 0 : v.textContent) != null && f.trim()))
          continue;
        const g = v.getBoundingClientRect();
        (zy(g, d) ||
          cS(v) ||
          (m === "header" ? g.bottom > c.top + 0.8 : g.top < c.bottom - 0.8)) &&
          t.push({
            severity: "error",
            code: "identity",
            page: s,
            message: `Page ${s}: the ${m} is clipped or overlaps the schedule. Shorten its text or increase the ${m === "header" ? "top" : "bottom"} margin.`,
          });
      }
      const h = [];
      for (const m of o.querySelectorAll("[data-print-cell]")) {
        const v = m.firstElementChild;
        if (!((p = v == null ? void 0 : v.textContent) != null && p.trim()))
          continue;
        m.dataset.printInArea !== "false" && (n = true);
        const g = m.dataset.printCell,
          b = m.dataset.printNumber === "true";
        (b && h.push({ address: g, bounds: m.getBoundingClientRect() }),
          cS(v) &&
            t.push({
              severity: b ? "error" : "warning",
              code: b ? "number" : "text",
              page: s,
              address: g,
              message: `Page ${s}, ${g}: ${b ? "the full number does not fit" : "text is clipped"}. AutoFit or widen the column, increase row height, or reduce the cell font size before printing.`,
            }));
      }
      for (const m of o.querySelectorAll(".dg-print-object")) {
        m.dataset.printInArea !== "false" && (n = true);
        const v = m.dataset.printLabel || m.dataset.printObject || "Drawing",
          g = m.getBoundingClientRect();
        zy(g, c) &&
          t.push({
            severity: "error",
            code: "object",
            page: s,
            message: `Page ${s}: “${v}” crosses the printable boundary. Move or resize it; drawings are not continued onto the next page.`,
          });
        const b = m.querySelector(".dg-shape-label"),
          y = [...m.querySelectorAll("svg text")];
        ((b && cS(b)) || y.some((w) => zy(w.getBoundingClientRect(), g))) &&
          t.push({
            severity: "error",
            code: "object",
            page: s,
            message: `Page ${s}: the label in “${v}” is clipped. Enlarge the object or shorten the label.`,
          });
        for (const w of h)
          Math.min(g.right, w.bounds.right) - Math.max(g.left, w.bounds.left) >
            1 &&
            Math.min(g.bottom, w.bounds.bottom) -
              Math.max(g.top, w.bounds.top) >
              1 &&
            t.push({
              severity: "error",
              code: "object",
              page: s,
              address: w.address,
              message: `Page ${s}, ${w.address}: “${v}” overlaps a numeric cell. Move the drawing so the amount remains visible.`,
            });
        for (const w of m.querySelectorAll("img"))
          (!w.complete || w.naturalWidth <= 0) &&
            t.push({
              severity: "error",
              code: "image",
              page: s,
              message: `Page ${s}: picture “${v}” could not be decoded. Replace it before printing.`,
            });
      }
    }),
    r.length &&
      !n &&
      !t.some((o) => o.code === "layout") &&
      t.unshift({
        severity: "error",
        code: "empty",
        message:
          "The print area is empty: no displayed values or visible objects. Choose a populated print area. Formatting, repeated titles outside the area, and page identifiers alone do not count as schedule content.",
      }),
    { pages: r.length, issues: t }
  );
}
function k3(e) {
  const r = e.issues.filter((t) => t.severity === "error");
  if (r.length)
    throw Error(
      `${r.length} print issue(s). ${r
        .slice(0, 3)
        .map((t) => t.message)
        .join(" ")}`,
    );
}
async function gw(e, r = {}) {
  var s;
  const t = e.ownerDocument ?? e,
    { signal: n, timeoutMs: o = 1e4 } = r;
  if (n != null && n.aborted) throw Error("Print preparation was cancelled.");
  const a = [];
  try {
    const d = Promise.all([
      ((s = t.fonts) == null ? void 0 : s.ready) ?? Promise.resolve(),
      ...Array.from(e.querySelectorAll("img")).map(
        (u) =>
          new Promise((l, c) => {
            const h = () =>
              u.naturalWidth > 0
                ? l()
                : c(
                    Error(
                      `Picture “${u.alt || "Untitled"}” could not be decoded. Replace it before printing.`,
                    ),
                  );
            if (u.complete) {
              h();
              return;
            }
            const f = () =>
              c(
                Error(
                  `Picture “${u.alt || "Untitled"}” could not be loaded. Replace it before printing.`,
                ),
              );
            (u.addEventListener("load", h, { once: !0 }),
              u.addEventListener("error", f, { once: !0 }),
              a.push(() => {
                (u.removeEventListener("load", h),
                  u.removeEventListener("error", f));
              }));
          }),
      ),
    ]);
    await Promise.race([
      d,
      new Promise((u, l) => {
        const c = setTimeout(
            () =>
              l(
                Error(
                  "Print assets did not finish loading within 10 seconds. Check the pictures and fonts, then retry.",
                ),
              ),
            o,
          ),
          h = () => l(Error("Print preparation was cancelled."));
        (n == null || n.addEventListener("abort", h, { once: !0 }),
          a.push(() => {
            (clearTimeout(c), n == null || n.removeEventListener("abort", h));
          }));
      }),
    ]);
  } finally {
    a.forEach((d) => d());
  }
}
async function Mme(e, r, t) {
  if (
    (t.assertCurrent(),
    e.document.write(r),
    e.document.close(),
    await gw(e.document, t),
    t.assertCurrent(),
    e.closed)
  )
    throw Error("The print window was closed. Reopen print preview to retry.");
  const n = mw(e.document);
  return (k3(n), e.focus(), e.print(), n);
}
const m2 = { pixelRatio: 2, maxBytes: 50 * 1024 * 1024, pageTimeoutMs: 3e4 };
function jme(e, r) {
  return `${
    `${e.name}-${Nr(e, r).name}-rev-${e.revision}`
      .replace(/[\x00-\x1f\x7f<>:"/\\|?*]+/g, "-")
      .replace(/[. ]+$/g, "")
      .slice(0, 160) || "Schedule"
  }.pdf`;
}
async function Fme(e, r, t) {
  var h, f, p, m, v, g;
  const n = () => {
      var b;
      if ((b = t.signal) != null && b.aborted)
        throw Error("PDF creation was cancelled.");
      t.assertCurrent();
    },
    o = async (b) => {
      var E;
      let y, w;
      try {
        return await Promise.race([
          b,
          new Promise((A, k) => {
            var F, I;
            ((y = setTimeout(
              () =>
                k(
                  Error(
                    "A PDF page did not finish within 30 seconds. Reduce its complexity or use browser printing.",
                  ),
                ),
              m2.pageTimeoutMs,
            )),
              (w = () => k(Error("PDF creation was cancelled."))),
              (F = t.signal) == null ||
                F.addEventListener("abort", w, { once: !0 }),
              (I = t.signal) != null && I.aborted && w());
          }),
        ]);
      } finally {
        (clearTimeout(y),
          w && ((E = t.signal) == null || E.removeEventListener("abort", w)));
      }
    };
  if ((n(), typeof document > "u"))
    throw Error("PDF rendering requires a browser with canvas support.");
  const a = Nr(e, r),
    s = t.layout ?? a.pageLayout ?? { ...ZR(a), ...XE },
    d = cW(e, r, s),
    u = KR(s),
    l = document.createElement("iframe");
  ((l.title = "Preparing schedule PDF"),
    l.setAttribute("aria-hidden", "true"),
    (l.tabIndex = -1),
    l.setAttribute("sandbox", "allow-same-origin"),
    (l.style.cssText = `position:fixed;left:-100000px;top:0;width:${u.width}px;height:${u.height}px;border:0;pointer-events:none;`));
  const c = [];
  try {
    await new Promise((F, I) => {
      var z;
      const O = () => I(Error("PDF creation was cancelled.")),
        j = setTimeout(
          () =>
            I(
              Error(
                "The PDF rendering frame could not load. Check the embedding host's content policy.",
              ),
            ),
          1e4,
        );
      ((l.onload = () => F()),
        (l.onerror = () => I(Error("The PDF rendering frame could not load."))),
        (z = t.signal) == null || z.addEventListener("abort", O, { once: !0 }),
        c.push(() => {
          var V;
          (clearTimeout(j),
            (V = t.signal) == null || V.removeEventListener("abort", O),
            (l.onload = null),
            (l.onerror = null));
        }),
        (l.srcdoc = d.html.replace(
          "<head>",
          `<head><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src 'none'">`,
        )),
        document.body.append(l));
    });
    const b = l.contentDocument;
    if (!b)
      throw Error(
        "The host does not allow the isolated PDF rendering document.",
      );
    (await gw(b, { signal: t.signal }), n());
    const y = mw(b);
    k3(y);
    const w = [...b.querySelectorAll(".dg-print-paper")];
    if (w.length !== d.pageCount || !w.length || w.length > wg.pages)
      throw Error(
        "The PDF page plan changed. Reopen print preview and try again.",
      );
    const E = await PDFDocument.create();
    (E.setTitle(`${e.name} — ${a.name}`),
      E.setSubject(
        `Period: ${((h = e.metadata) == null ? void 0 : h.period) || "not set"}; engagement: ${((f = e.metadata) == null ? void 0 : f.engagement) || "not set"}; revision ${e.revision}. Image-based schedule snapshot.`,
      ),
      E.setCreator("Ledgerglass schedule PDF"),
      E.setProducer("Ledgerglass / pdf-lib"));
    let A = 0;
    (p = t.onProgress) == null || p.call(t, 0, w.length);
    for (let F = 0; F < w.length; F++) {
      if ((m = t.signal) != null && m.aborted)
        throw Error("PDF creation was cancelled.");
      const I = await o(
        toCanvas(w[F], {
          width: u.width,
          height: u.height,
          pixelRatio: m2.pixelRatio,
          backgroundColor: "#ffffff",
          skipFonts: !0,
          fontEmbedCSS: "",
          skipAutoScale: !0,
        }),
      );
      try {
        if ((v = t.signal) != null && v.aborted)
          throw Error("PDF creation was cancelled.");
        if (
          I.width !== Math.floor(u.width * 2) ||
          I.height !== Math.floor(u.height * 2)
        )
          throw Error(
            "The browser changed the PDF page resolution. Use browser printing instead.",
          );
        const O = await o(
          new Promise((V, D) =>
            I.toBlob(
              (B) =>
                B ? V(B) : D(Error("The browser could not render a PDF page.")),
              "image/png",
            ),
          ),
        );
        if (((A += O.size), A > m2.maxBytes))
          throw Error(
            "PDF images exceed 50 MB. Reduce the print area or image sizes.",
          );
        const j = await E.embedPng(await O.arrayBuffer()),
          z = E.addPage([u.width * 0.75, u.height * 0.75]);
        z.drawImage(j, {
          x: 0,
          y: 0,
          width: z.getWidth(),
          height: z.getHeight(),
        });
      } finally {
        ((I.width = 0), (I.height = 0));
      }
      (g = t.onProgress) == null || g.call(t, F + 1, w.length);
    }
    n();
    const k = await E.save();
    if ((n(), k.length > m2.maxBytes))
      throw Error(
        "The PDF exceeds 50 MB. Reduce the print area or image sizes.",
      );
    return {
      bytes: k,
      filename: jme(e, r),
      pageCount: w.length,
      inspection: y,
      notices: [
        "Image-based PDF at 192 dpi: text is not selectable, searchable or accessibility-tagged. Use browser printing when a text PDF is needed.",
        ...d.notices,
      ],
    };
  } finally {
    (c.forEach((b) => b()), l.remove());
  }
}

export { Cfe, Fme, Mme, Wz, Xz, cW, gw, k3, mw, qz, uW };
