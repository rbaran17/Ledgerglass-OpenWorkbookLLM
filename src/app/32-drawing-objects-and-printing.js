import i from "react/jsx-runtime";
import ne from "react";
import { X, ChartColumn, ImagePlus, Shapes, Trash2, Move } from "lucide-react";
import c_ from "react-dom/server";
import { toCanvas } from "html-to-image";
import { PDFDocument } from "pdf-lib";
import {
  assertDrawingPreviewCurrent,
  cellStyleToCss,
  F2,
  evaluateConditionalFormatRules,
  setSheetDrawingObjects,
  columnWidthAt,
  resolvePageDimensions,
  findSheetOrThrow,
  inspectPictureDataUrl,
  createAnchoredDrawingObject,
  textWrapToCss,
  validatePageSetup,
  calculateWorkbookWithSettings,
  inspectChartSource,
  defaultPrintHeaderFooter,
  applyChartDrawingLayout,
  createDefaultPageSetup,
  formatDiagnosticsMessage,
  cn,
  parseRangeAddress,
  createPictureEvidence,
  rowHeightAt,
  expandHeaderFooterTemplate,
  measureChartSelection,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  vx,
  printRenderLimits,
  validateDrawingObjects,
  zU,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { importPictureFile } from "./07-file-import-and-worker-dispatch.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import { CellValuePreview } from "./22-themes-and-cell-styles.js";
import { ConditionalFormatIndicator } from "./23-format-cells-and-conditional-formatting.js";
import {
  ChartReportSvg,
  downloadTextFile,
  ChartSettingsPanel,
  ChartPreviewFrame,
} from "./31-pivot-explorer-and-charts.js";

function ChartInspectionPanel({
  report,
  error,
  revision,
  workbookId,
  onExport = downloadTextFile,
}) {
  const { text, language } = useWorkbookMessages(),
    [d, u] = ne.useState(0),
    [l, c] = ne.useState(""),
    h = ne.useRef(null),
    f = report?.points ?? [],
    p = Math.min(d, Math.max(0, Math.ceil(f.length / 20) - 1)),
    m = (v = false) => {
      try {
        if ((c(""), !report)) throw Error(error);
        if (v) {
          const b = h.current?.querySelector("svg");
          if (!b)
            throw Error("Resolve chart errors before exporting the visual.");
          onExport({
            filename: "Chart.svg",
            mime: "image/svg+xml",
            content: new XMLSerializer().serializeToString(b),
          });
        } else
          onExport({
            filename: "Chart-inspection.json",
            mime: "application/json",
            content: JSON.stringify(
              { workbookId, revision, draft: true, ...report },
              null,
              2,
            ),
          });
      } catch (error) {
        c(error instanceof Error ? error.message : String(error));
      }
    };
  return i.jsxs("section", {
    className: "dg-chart-inspection",
    "aria-label": text("Chart inspection"),
    lang: language("Chart inspection"),
    "data-ui-label": "Chart inspection",
    children: [
      i.jsx("h3", {
        children: i.jsx(LocalizedText, {
          children: "Preview and source inspection",
        }),
      }),
      error && i.jsx("p", { className: "dg-error", children: error }),
      report &&
        i.jsxs(i.Fragment, {
          children: [
            i.jsx("div", {
              ref: h,
              className:
                report.kind === "sparkline"
                  ? "dg-spark-preview"
                  : "dg-chart-preview",
              children: i.jsx(ChartReportSvg, { report }),
            }),
            i.jsx("p", { children: report.scope }),
            i.jsxs("p", {
              className: "dg-muted",
              children: [
                report.population.policy,
                " ",
                i.jsx(LocalizedText, {
                  children:
                    "Exports contain this complete draft configuration and source inspection at revision ",
                }),
                revision,
                i.jsx(LocalizedText, {
                  children: ". Source cells are unchanged.",
                }),
              ],
            }),
            report.diagnostics.length > 0 &&
              i.jsx("ul", {
                className: "dg-chart-diagnostics",
                children: report.diagnostics.map((diagnostic, g) =>
                  i.jsxs(
                    "li",
                    {
                      children: [diagnostic.severity, ": ", diagnostic.message],
                    },
                    g,
                  ),
                ),
              }),
            i.jsx("button", {
              type: "button",
              onClick: () => m(),
              children: i.jsx(LocalizedText, {
                children: "Export chart inspection",
              }),
            }),
            i.jsx("button", {
              type: "button",
              disabled: report.diagnostics.some(
                (diagnostic) => diagnostic.severity === "error",
              ),
              onClick: () => m(true),
              children: i.jsx(LocalizedText, { children: "Export chart SVG" }),
            }),
            i.jsxs("details", {
              children: [
                i.jsx("summary", {
                  children: i.jsx(LocalizedText, {
                    children: "Inspect source data and full series",
                  }),
                }),
                i.jsxs("dl", {
                  children: [
                    i.jsx("dt", {
                      children: i.jsx(LocalizedText, {
                        children: "Title / alt text",
                      }),
                    }),
                    i.jsxs("dd", {
                      children: [report.title, " / ", report.altText],
                    }),
                    i.jsx("dt", {
                      children: i.jsx(LocalizedText, {
                        children: "Axes and units",
                      }),
                    }),
                    i.jsxs("dd", {
                      children: [
                        i.jsx(LocalizedText, { children: "X: " }),
                        report.axes.x.title || "(automatic)",
                        " (",
                        report.axes.x.scale,
                        i.jsx(LocalizedText, { children: "); Y: " }),
                        report.axes.y.title || "(automatic)",
                        " (",
                        report.axes.y.scale,
                        i.jsx(LocalizedText, { children: "); units:" }),
                        " ",
                        report.axes.units || "not specified",
                      ],
                    }),
                    i.jsx("dt", {
                      children: i.jsx(LocalizedText, { children: "Series" }),
                    }),
                    i.jsx("dd", {
                      children: report.series
                        .map(
                          (sery) =>
                            `${sery.name} (column ${sery.column}, ${sery.kind}, ${sery.included ? "included" : "excluded"})`,
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
                              children: i.jsx(LocalizedText, {
                                children: "Source row",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Scope",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Category / X",
                              }),
                            }),
                            report.series.map((sery) =>
                              i.jsx("th", { children: sery.name }, sery.column),
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
                                v.values.map((value) =>
                                  i.jsxs(
                                    "td",
                                    {
                                      children: [
                                        value.address,
                                        ": ",
                                        String(value.value ?? "(blank)"),
                                        value.raw !== value.value &&
                                          i.jsxs("small", {
                                            children: [
                                              " ",
                                              i.jsx(LocalizedText, {
                                                children: "Raw: ",
                                              }),
                                              String(value.raw),
                                            ],
                                          }),
                                        value.note &&
                                          i.jsxs("small", {
                                            children: [
                                              " ",
                                              i.jsx(LocalizedText, {
                                                children: "Note: ",
                                              }),
                                              value.note,
                                            ],
                                          }),
                                      ],
                                    },
                                    value.address,
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
                  children: i.jsx(LocalizedText, {
                    children: "Previous chart source rows",
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: (p + 1) * 20 >= f.length,
                  onClick: () => u(p + 1),
                  children: i.jsx(LocalizedText, {
                    children: "Next chart source rows",
                  }),
                }),
                report.histogram &&
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Bins:" }),
                      " ",
                      report.histogram
                        .map(
                          (v) =>
                            `${v.lower}–${v.upper}: ${v.count} rows ${v.upperInclusive ? "(upper inclusive)" : "(upper exclusive)"}`,
                        )
                        .join("; "),
                    ],
                  }),
                report.waterfall &&
                  i.jsxs("p", {
                    children: [
                      i.jsx(LocalizedText, { children: "Waterfall:" }),
                      " ",
                      report.waterfall
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
function PicturePropertiesDialog(e) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    o = e.workbook.sheets
      .find((sheet) => sheet.id === e.sheetId)
      ?.objects?.find((object) => object.id === e.objectId),
    a =
      o?.kind ??
      (/picture|image/i.test(e.mode)
        ? "picture"
        : /chart/i.test(e.mode)
          ? "chart"
          : "shape"),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    [d] = ne.useState(() => ({ current: F2(e.workbook, e.sheetId, e.range) })),
    u = ne.useRef(e.mode),
    l = ne.useRef(e.objectId),
    c = ne.useRef(e),
    h = ne.useRef(null),
    f = ne.useRef(true),
    p = ne.useRef(0);
  (ne.useRef(null), (c.current = e));
  const [object, v] = ne.useState(() =>
      o
        ? structuredClone(o)
        : {
            ...createAnchoredDrawingObject(a, e.range),
            ...(a === "chart"
              ? { chart: "column", range: e.range, headers: true }
              : {}),
          },
    ),
    [sourceLocator, b] = ne.useState(o?.evidence?.sourceLocator ?? ""),
    [y, w] = ne.useState(""),
    [E, A] = ne.useState(false),
    k =
      d.current.guard !== JSON.stringify(e.workbook) ||
      d.current.sheetId !== e.sheetId ||
      d.current.selection !== e.range ||
      u.current !== e.mode ||
      l.current !== e.objectId,
    F = e.readOnly || k || !sheet || !!sheet?.protection?.enabled,
    I = () => {
      const L = c.current;
      if (u.current !== L.mode || l.current !== L.objectId)
        throw Error("The drawing selection changed. Reopen this dialog.");
      assertDrawingPreviewCurrent(L.workbook, L.sheetId, {
        readOnly: L.readOnly,
        snapshot: d.current,
        selection: L.range,
      });
    },
    update = (L) => {
      !F && !E && v(($) => ({ ...$, ...L }));
    };
  (ne.useEffect(() => {
    f.current = true;
    const activeElement = document.activeElement;
    return (
      h.current?.querySelector("button")?.focus(),
      () => {
        ((f.current = false),
          p.current++,
          cancellableOperation.cancel(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      F && (p.current++, cancellableOperation.cancel(), A(false));
    }, [F]));
  async function j(L) {
    if (!L) return;
    const $ = ++p.current;
    try {
      if (
        (I(),
        A(true),
        w(""),
        L.size > 2 * 1024 * 1024 ||
          L.size === 0 ||
          !["image/png", "image/jpeg"].includes(L.type))
      )
        throw Error("Choose a PNG or JPEG picture no larger than 2 MB.");
      const { dataURL } = await importPictureFile(
        L,
        cancellableOperation.begin(),
      );
      if (!f.current || $ !== p.current) return;
      I();
      const G = inspectPictureDataUrl(dataURL);
      if (G.bytes.length !== L.size)
        throw Error("The file bytes changed while reading the picture.");
      const pictureEvidence = createPictureEvidence(dataURL, {
          origin: "local-file",
          filename: L.name,
          ...(sourceLocator ? { sourceLocator: sourceLocator } : {}),
        }),
        Q = await new Promise((J, q) => {
          const W = new Image();
          ((W.onload = () =>
            J({ width: W.naturalWidth, height: W.naturalHeight })),
            (W.onerror = () => q(Error("The picture could not be decoded."))),
            (W.src = dataURL));
        });
      if (!f.current || $ !== p.current) return;
      if ((I(), Q.width !== G.width || Q.height !== G.height))
        throw Error(
          "Decoded picture dimensions do not match its validated header.",
        );
      const X = Math.min(1, 400 / G.width, 250 / G.height);
      v((J) => ({
        ...J,
        dataURL,
        evidence: pictureEvidence,
        label: L.name.slice(0, 500),
        width: Math.max(24, Math.round(G.width * X)),
        height: Math.max(24, Math.round(G.height * X)),
      }));
    } catch (error) {
      f.current &&
        $ === p.current &&
        w(error instanceof Error ? error.message : "Picture import failed.");
    } finally {
      f.current && $ === p.current && A(false);
    }
  }
  function onClick() {
    try {
      if ((I(), E)) return;
      const $ = c.current,
        sheet = $.workbook.sheets.find((sheet) => sheet.id === $.sheetId);
      if (
        (a === "chart" &&
          (measureChartSelection(object.range, sheet),
          formatDiagnosticsMessage(
            inspectChartSource(
              sheet,
              object,
              calculateWorkbookWithSettings($.workbook),
            ),
          )),
        a === "picture" && !object.dataURL)
      )
        throw Error("Choose a picture first.");
      const G = structuredClone(object);
      if (
        (a === "picture" &&
          (G.evidence = {
            ...(object.evidence ??
              createPictureEvidence(object.dataURL, {
                origin: "legacy-native",
              })),
            sourceLocator: sourceLocator || undefined,
          }),
        e.objectId &&
          !sheet.objects?.some((object) => object.id === e.objectId))
      )
        throw Error("The drawing no longer exists.");
      const H = e.objectId
          ? (sheet.objects ?? []).map((X) => (X.id === e.objectId ? G : X))
          : [...(sheet.objects ?? []), G],
        Q = setSheetDrawingObjects($.workbook, $.sheetId, H, {
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
    } catch (error) {
      w(
        error instanceof Error
          ? error.message
          : "The object could not be saved.",
      );
    }
  }
  const V =
      object.evidence ??
      (object.dataURL
        ? createPictureEvidence(object.dataURL, { origin: "legacy-native" })
        : undefined),
    D = ne.useMemo(() => {
      if (a !== "chart" || !sheet) return { report: null, error: "" };
      try {
        return {
          report: inspectChartSource(
            sheet,
            object,
            calculateWorkbookWithSettings(e.workbook),
          ),
          error: "",
        };
      } catch (error) {
        return {
          report: null,
          error: error instanceof Error ? error.message : String(error),
        };
      }
    }, [a, sheet, object, e.workbook]),
    B = a[0].toUpperCase() + a.slice(1) + " properties";
  return i.jsx("div", {
    className: "dg-backdrop",
    onKeyDown: (event) => {
      if (
        (event.stopPropagation(),
        event.key === "Escape" && (event.preventDefault(), e.onClose()),
        event.key === "Tab")
      ) {
        const U = [
            ...(h.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled)",
            ) ?? []),
          ].filter((Q) => !Q.closest("fieldset:disabled")),
          G = U[0],
          H = U.at(-1);
        event.shiftKey && document.activeElement === G
          ? (event.preventDefault(), H == null || H.focus())
          : !event.shiftKey &&
            document.activeElement === H &&
            (event.preventDefault(), G == null || G.focus());
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
              "aria-label": text("Close drawing tools"),
              onClick: e.onClose,
              lang: language("Close drawing tools"),
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
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook or selection changed. Close and reopen this dialog.",
                }),
              }),
            e.readOnly &&
              i.jsx("p", {
                role: "alert",
                className: "dg-error",
                children: i.jsx(LocalizedText, {
                  children:
                    "This workbook is read-only. Object properties remain available to view.",
                }),
              }),
            sheet?.protection?.enabled &&
              i.jsx("p", {
                role: "alert",
                className: "dg-error",
                children: i.jsx(LocalizedText, {
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
                          i.jsx(LocalizedText, {
                            children: "Choose a local picture",
                          }),
                          i.jsx("input", {
                            type: "file",
                            accept: "image/png,image/jpeg",
                            "aria-label": text("Picture file"),
                            disabled: F || E,
                            onChange: (event) =>
                              void j(event.target.files?.[0]),
                            lang: language("Picture file"),
                            "data-ui-label": "Picture file",
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        className: "dg-muted",
                        children: i.jsx(LocalizedText, {
                          children:
                            "PNG or JPEG · up to 2 MB · stored with the workbook",
                        }),
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Source locator (plain text)",
                          }),
                          i.jsx("input", {
                            "aria-label": text("Picture source locator"),
                            value: sourceLocator,
                            maxLength: 2e3,
                            onChange: (event) => b(event.target.value),
                            lang: language("Picture source locator"),
                            "data-ui-label": "Picture source locator",
                          }),
                        ],
                      }),
                      object.dataURL &&
                        i.jsx("img", {
                          className: "dg-picture-preview",
                          src: object.dataURL,
                          alt: object.label,
                        }),
                    ],
                  }),
                a === "shape" &&
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Shape" }),
                      i.jsxs("select", {
                        "aria-label": text("Shape type"),
                        value: object.shape,
                        onChange: (event) =>
                          update({ shape: event.target.value }),
                        lang: language("Shape type"),
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
                  i.jsx(ChartSettingsPanel, {
                    object: object,
                    report: D.report,
                    update: update,
                  }),
                i.jsxs("label", {
                  children: [
                    a === "chart" ? "Chart title" : "Label",
                    i.jsx("input", {
                      "aria-label": text("Object label"),
                      value: object.label,
                      maxLength: 500,
                      onChange: (event) =>
                        update({ label: event.target.value }),
                      lang: language("Object label"),
                      "data-ui-label": "Object label",
                    }),
                  ],
                }),
                object.chart !== "sparkline" &&
                  i.jsxs("div", {
                    className: "dg-two",
                    children: [
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Width (px)" }),
                          i.jsx(LocalizedNumberField, {
                            min: 24,
                            max: 4e3,
                            "aria-label": text("Object width"),
                            value: object.width,
                            onValueChange: (L) => update({ width: Number(L) }),
                            lang: language("Object width"),
                            "data-ui-label": "Object width",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Height (px)" }),
                          i.jsx(LocalizedNumberField, {
                            min: 24,
                            max: 4e3,
                            "aria-label": text("Object height"),
                            value: object.height,
                            onValueChange: (L) => update({ height: Number(L) }),
                            lang: language("Object height"),
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
                          i.jsx(LocalizedText, { children: "Fill" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": text("Object fill"),
                            value: object.fill,
                            onChange: (event) =>
                              update({ fill: event.target.value }),
                            lang: language("Object fill"),
                            "data-ui-label": "Object fill",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx(LocalizedText, { children: "Outline" }),
                          i.jsx("input", {
                            type: "color",
                            "aria-label": text("Object outline"),
                            value: object.stroke,
                            onChange: (event) =>
                              update({ stroke: event.target.value }),
                            lang: language("Object outline"),
                            "data-ui-label": "Object outline",
                          }),
                        ],
                      }),
                    ],
                  }),
                i.jsxs("p", {
                  className: "dg-muted",
                  children: [
                    i.jsx(LocalizedText, { children: "Anchor: " }),
                    object.anchor,
                    i.jsx(LocalizedText, {
                      children:
                        ". Select the object to move, resize, rename, or delete it.",
                    }),
                  ],
                }),
              ],
            }),
            a === "chart" &&
              i.jsx(ChartInspectionPanel, {
                ...D,
                revision: e.workbook.revision,
                workbookId: e.workbook.id,
                onExport: e.onChartExport,
              }),
            a === "picture" &&
              V &&
              i.jsxs("dl", {
                className: "dg-evidence-properties",
                "aria-label": text("Picture evidence properties"),
                lang: language("Picture evidence properties"),
                "data-ui-label": "Picture evidence properties",
                children: [
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, { children: "Origin" }),
                  }),
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
                    children: i.jsx(LocalizedText, { children: "Byte length" }),
                  }),
                  i.jsx("dd", { children: V.byteLength.toLocaleString() }),
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, { children: "SHA-256" }),
                  }),
                  i.jsx("dd", { children: V.sha256 }),
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, {
                      children: "Source reference",
                    }),
                  }),
                  i.jsx("dd", { children: V.sourceReference }),
                  V.containerFilename &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Imported workbook",
                          }),
                        }),
                        i.jsx("dd", { children: V.containerFilename }),
                      ],
                    }),
                  V.mediaReference &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Observed media reference",
                          }),
                        }),
                        i.jsx("dd", { children: V.mediaReference }),
                      ],
                    }),
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, {
                      children: "Source locator",
                    }),
                  }),
                  i.jsx("dd", { children: sourceLocator || "Not recorded" }),
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
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "dg-primary",
              onClick: onClick,
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
function DrawingObjectSvg({ object, sheet, calc, textScale = 1 }) {
  const o = ne.useMemo(() => {
    if (object.kind !== "picture") return "";
    try {
      return (inspectPictureDataUrl(object.dataURL ?? ""), "");
    } catch (error) {
      return error instanceof Error ? error.message : "Invalid picture.";
    }
  }, [object.kind, object.dataURL]);
  return object.kind === "picture"
    ? o
      ? i.jsxs("span", {
          role: "img",
          "aria-label": "Invalid picture: " + o,
          children: [
            i.jsx(LocalizedText, { children: "Picture cannot be displayed: " }),
            o,
          ],
        })
      : i.jsx("img", {
          src: object.dataURL,
          alt: object.label,
          draggable: false,
        })
    : object.kind === "shape"
      ? i.jsxs("div", {
          role: "img",
          "aria-label": object.label,
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
                fill: object.fill,
                stroke: object.stroke,
                strokeWidth: "2",
                children:
                  object.shape === "ellipse"
                    ? i.jsx("ellipse", {
                        cx: "120",
                        cy: "70",
                        rx: "116",
                        ry: "66",
                      })
                    : object.shape === "arrow"
                      ? i.jsx("path", { d: "M4 45H158V8L235 70 158 132V95H4Z" })
                      : object.shape !== "text"
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
                inset: 8 * textScale,
                width: "auto",
                height: "auto",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                font: `${15 * textScale}px Arial, sans-serif`,
                color: object.stroke,
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
                children: object.label,
              }),
            }),
          ],
        })
      : i.jsx(ChartPreviewFrame, { object, sheet, calc });
}
function DrawingObjectOverlay(e) {
  const { text, language } = useWorkbookMessages(),
    sheet = e.workbook.sheets.find((sheet) => sheet.id === e.sheetId),
    o = ne.useRef(e);
  o.current = e;
  const a = e.readOnly || !sheet || !!sheet?.protection?.enabled,
    [s, d] = ne.useState(null),
    [u, l] = ne.useState(""),
    [objectId, h] = ne.useState(null),
    [f, p] = ne.useState(null),
    m = ne.useRef(null);
  m.current = f;
  function v(E, A, k) {
    try {
      const I = o.current,
        sheet = I.workbook.sheets.find((sheet) => sheet.id === I.sheetId);
      if (!sheet?.objects?.some((object) => object.id === E.id))
        throw Error("The selected drawing no longer exists.");
      const j = setSheetDrawingObjects(
        I.workbook,
        I.sheetId,
        sheet.objects.map((object) => (object.id === E.id ? E : object)),
        {
          readOnly: I.readOnly,
          snapshot: k ?? F2(I.workbook, I.sheetId, I.selection),
          selection: I.selection,
        },
      );
      (j !== I.workbook && I.onWorkbookChange(j, A), l(""));
    } catch (error) {
      l(error instanceof Error ? error.message : "Could not update object.");
    }
  }
  function g(E) {
    try {
      const A = o.current,
        sheet = A.workbook.sheets.find((sheet) => sheet.id === A.sheetId),
        F = setSheetDrawingObjects(
          A.workbook,
          A.sheetId,
          (sheet?.objects ?? []).filter((I) => I.id !== E),
          {
            readOnly: A.readOnly,
            snapshot: F2(A.workbook, A.sheetId, A.selection),
            selection: A.selection,
          },
        );
      (F !== A.workbook && A.onWorkbookChange(F, "Delete drawing"), d(null));
    } catch (error) {
      l(error instanceof Error ? error.message : "Could not delete drawing.");
    }
  }
  function b(E, A, k, F) {
    try {
      if (
        (assertDrawingPreviewCurrent(e.workbook, e.sheetId, {
          readOnly: e.readOnly,
        }),
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
    } catch (error) {
      l(
        error instanceof Error
          ? error.message
          : "Could not start drawing edit.",
      );
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
            (assertDrawingPreviewCurrent(e.workbook, e.sheetId, {
              readOnly: e.readOnly,
              snapshot: f.snapshot,
              selection: e.selection,
            }),
            e.scale !== f.scale)
          )
            throw Error("The zoom changed. Start the drawing move again.");
        } catch (error) {
          (p(null),
            l(
              error instanceof Error
                ? error.message
                : "Drawing edit was cancelled.",
            ));
        }
    }, [e.workbook, e.readOnly, e.selection, e.scale, f]),
    ne.useEffect(() => {
      if (!f) return;
      const E = (event) => {
          const O = m.current;
          if (!O) return;
          const j = (event.clientX - O.x) / O.scale,
            z = (event.clientY - O.y) / O.scale,
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
        F = (event) => {
          event.key === "Escape" && k();
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
          "aria-label": text("Worksheet drawings"),
          lang: language("Worksheet drawings"),
          "data-ui-label": "Worksheet drawings",
          children: [
            (sheet?.objects ?? []).map((E) => {
              const object = f?.id === E.id ? f.next : E,
                k = e.positionForCell(object.anchor),
                disabled =
                  object.kind === "chart" && object.chart === "sparkline",
                I = applyChartDrawingLayout(sheet, object);
              return k
                ? i.jsxs(
                    "div",
                    {
                      className: `dg-drawing ${s === object.id ? "dg-drawing-selected" : ""}`,
                      "data-object-id": object.id,
                      style: {
                        left: k.left + I.offsetX * e.scale,
                        top: k.top + I.offsetY * e.scale,
                        width: I.width * e.scale,
                        height: I.height * e.scale,
                      },
                      tabIndex: 0,
                      "aria-label": object.kind + ": " + object.label,
                      onClick: () => d(object.id),
                      onKeyDown: (event) => {
                        (event.stopPropagation(),
                          event.key === "Escape" && d(null),
                          (event.key === "Delete" ||
                            event.key === "Backspace") &&
                            event.target === event.currentTarget &&
                            !a &&
                            (event.preventDefault(), g(object.id)));
                      },
                      children: [
                        i.jsx(DrawingObjectSvg, {
                          object: object,
                          sheet: sheet,
                          calc: e.calc,
                          textScale: e.scale,
                        }),
                        s === object.id &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("div", {
                                className: "dg-object-toolbar",
                                children: [
                                  (object.kind === "picture" ||
                                    object.kind === "chart") &&
                                    i.jsx("button", {
                                      className: "dg-properties-button",
                                      "aria-label":
                                        (object.kind === "chart"
                                          ? "Chart"
                                          : "Picture") +
                                        " properties: " +
                                        object.label,
                                      onClick: () => h(object.id),
                                      children: "ⓘ",
                                    }),
                                  !a &&
                                    i.jsxs(i.Fragment, {
                                      children: [
                                        i.jsx("button", {
                                          "aria-label": "Move " + object.label,
                                          title: disabled
                                            ? "Set the sparkline anchor in Chart properties"
                                            : "Drag to move",
                                          disabled: disabled,
                                          onPointerDown: (event) => {
                                            (event.preventDefault(),
                                              b(
                                                object,
                                                "move",
                                                event.clientX,
                                                event.clientY,
                                              ));
                                          },
                                          children: i.jsx(Move, { size: 14 }),
                                        }),
                                        i.jsx("input", {
                                          "aria-label": text(
                                            "Selected object label",
                                          ),
                                          value: object.label,
                                          onChange: (event) =>
                                            v(
                                              {
                                                ...object,
                                                label: event.target.value,
                                              },
                                              "Rename drawing",
                                            ),
                                          lang: language(
                                            "Selected object label",
                                          ),
                                          "data-ui-label":
                                            "Selected object label",
                                        }),
                                        i.jsx("button", {
                                          "aria-label":
                                            "Delete " + object.label,
                                          onClick: () => g(object.id),
                                          children: i.jsx(Trash2, { size: 14 }),
                                        }),
                                      ],
                                    }),
                                ],
                              }),
                              !a &&
                                !disabled &&
                                i.jsx("button", {
                                  className: "dg-resize-object",
                                  "aria-label": "Resize " + object.label,
                                  onPointerDown: (event) => {
                                    (event.preventDefault(),
                                      b(
                                        object,
                                        "resize",
                                        event.clientX,
                                        event.clientY,
                                      ));
                                  },
                                }),
                            ],
                          }),
                      ],
                    },
                    object.id,
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
        objectId &&
          sheet?.objects?.some((object) => object.id === objectId) &&
          i.jsx(PicturePropertiesDialog, {
            mode: "drawing-properties",
            objectId: objectId,
            workbook: e.workbook,
            sheetId: e.sheetId,
            range:
              e.selection ??
              sheet.objects.find((object) => object.id === objectId).anchor,
            readOnly: e.readOnly,
            onClose: () => h(null),
            onWorkbookChange: e.onWorkbookChange,
          }),
      ],
    })
  );
}
const printPaperStyles =
  ".dg-print-paper{box-sizing:border-box;print-color-adjust:exact;-webkit-print-color-adjust:exact;color:#000;background:white;position:relative;overflow:hidden;break-after:page;page-break-after:always}.dg-print-paper:last-child{break-after:auto;page-break-after:auto}.dg-print-content{position:relative;overflow:hidden}.dg-print-table{border-collapse:collapse;table-layout:fixed;font-family:Calibri,Arial,sans-serif;font-size:14px}.dg-print-table td,.dg-print-table th{box-sizing:border-box;padding:0 3px;overflow:hidden;white-space:nowrap;font-weight:400}.dg-print-table th{font:11px Arial;background:#f3f3f3}.dg-print-table thead tr{height:22px}.dg-print-table td>div{overflow:hidden;line-height:1.35}.dg-print-header,.dg-print-footer{position:absolute;left:0;right:0;text-align:center;font:11px Arial;color:#444;white-space:pre-wrap;line-height:1.3;overflow-wrap:anywhere}.dg-print-object{position:absolute;overflow:hidden}.dg-print-object>img,.dg-print-object>svg{width:100%;height:100%;object-fit:contain}.dg-shape-label{width:100%;height:100%;display:flex;align-items:center;justify-content:center;font:15px Arial;white-space:pre-wrap;text-align:center;overflow:hidden}@media print{.dg-print-paper{margin:0!important;box-shadow:none!important;zoom:1!important}}";
function PrintPagePreview({
  workbook,
  sheet,
  layout,
  page,
  index,
  total,
  scale,
  calc,
  ruleStyles,
}) {
  const axisVisibilityIndex = createAxisVisibilityIndex(
      sheet.rowCount,
      cn(sheet, "row"),
    ),
    c = createAxisVisibilityIndex(sheet.colCount, cn(sheet, "column")),
    pageDimensions = resolvePageDimensions(layout),
    rangeAddress = parseRangeAddress(layout.printArea),
    p =
      page.rows ??
      Array.from(
        { length: page.lastRow - page.firstRow + 1 },
        (E, A) => page.firstRow + A,
      ).filter((E) => !axisVisibilityIndex.isHidden(E)),
    m =
      page.columns ??
      Array.from(
        { length: page.lastCol - page.firstCol + 1 },
        (E, A) => page.firstCol + A,
      ).filter((E) => !c.isHidden(E)),
    width = m.reduce(
      (E, A) => E + columnWidthAt(sheet, A),
      layout.headings ? 32 : 0,
    ),
    g = new Map(p.map((E, A) => [E, A])),
    b = [0];
  for (const E of p) b.push(b.at(-1) + rowHeightAt(sheet, E));
  const y = new Map();
  for (const E of sheet.merges ?? []) {
    const rangeAddress = parseRangeAddress(E),
      k = p.filter(
        (O) => O >= rangeAddress.firstRow && O <= rangeAddress.lastRow,
      ),
      F = m.filter(
        (O) => O >= rangeAddress.firstCol && O <= rangeAddress.lastCol,
      );
    if (!k.length || !F.length) continue;
    const I = {
      firstRow: k[0],
      lastRow: k.at(-1),
      firstCol: F[0],
      lastCol: F.at(-1),
      rowSpan: k.length,
      colSpan: F.length,
      suppressContent:
        axisVisibilityIndex.isHidden(rangeAddress.firstRow) ||
        c.isHidden(rangeAddress.firstCol),
      anchor: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    };
    if (!(I.firstRow > I.lastRow || I.firstCol > I.lastCol))
      for (const O of k) for (const j of F) y.set(O * 16384 + j, I);
  }
  const w = (E) =>
    expandHeaderFooterTemplate(E, workbook, sheet, index + 1, total);
  return i.jsxs("article", {
    className: "dg-print-paper",
    "aria-label": `Print page ${index + 1}`,
    style: {
      boxSizing: "border-box",
      width: pageDimensions.width,
      height: pageDimensions.height,
      paddingTop: layout.margins.top * 96,
      paddingBottom: layout.margins.bottom * 96,
      paddingLeft: layout.margins.left * 96,
      paddingRight: layout.margins.right * 96,
    },
    children: [
      i.jsx("div", {
        className: "dg-print-header",
        style: {
          top: layout.margins.header * 96,
          left: layout.margins.left * 96,
          right: layout.margins.right * 96,
        },
        children: w(layout.header),
      }),
      i.jsx("div", {
        className: "dg-print-content",
        style: {
          width:
            pageDimensions.width -
            (layout.margins.left + layout.margins.right) * 96,
          height:
            pageDimensions.height -
            (layout.margins.top + layout.margins.bottom) * 96,
        },
        children: i.jsxs("div", {
          style: {
            transform: `scale(${scale})`,
            transformOrigin: "top left",
            width: width,
          },
          children: [
            i.jsxs("table", {
              className: "dg-print-table",
              style: { width: width },
              children: [
                i.jsxs("colgroup", {
                  children: [
                    layout.headings && i.jsx("col", { style: { width: 32 } }),
                    m.map((E) =>
                      i.jsx(
                        "col",
                        { style: { width: columnWidthAt(sheet, E) } },
                        E,
                      ),
                    ),
                  ],
                }),
                layout.headings &&
                  i.jsx("thead", {
                    children: i.jsxs("tr", {
                      children: [
                        i.jsx("th", {}),
                        m.map((E) =>
                          i.jsx("th", { children: columnIndexToLetters(E) }, E),
                        ),
                      ],
                    }),
                  }),
                i.jsx("tbody", {
                  children: p.map((E) =>
                    i.jsxs(
                      "tr",
                      {
                        style: { height: rowHeightAt(sheet, E) },
                        children: [
                          layout.headings && i.jsx("th", { children: E + 1 }),
                          m.map((A) => {
                            const k = y.get(E * 16384 + A);
                            if (k && (E !== k.firstRow || A !== k.firstCol))
                              return null;
                            const F =
                                k?.anchor ??
                                `${columnIndexToLetters(A)}${E + 1}`,
                              cell = k?.suppressContent
                                ? undefined
                                : sheet.cells[F],
                              O = {
                                ...cell,
                                raw: cell?.raw ?? null,
                                style: k?.suppressContent
                                  ? {}
                                  : {
                                      ...cell?.style,
                                      ...ruleStyles(F).style,
                                    },
                              },
                              j = k?.suppressContent ? null : calc(sheet.id, F),
                              z = b[g.get(k?.lastRow ?? E) + 1] - b[g.get(E)];
                            return i.jsx(
                              "td",
                              {
                                "data-print-cell": F,
                                "data-print-in-area":
                                  E >= rangeAddress.firstRow &&
                                  E <= rangeAddress.lastRow &&
                                  A >= rangeAddress.firstCol &&
                                  A <= rangeAddress.lastCol
                                    ? "true"
                                    : "false",
                                "data-print-number":
                                  typeof j == "number" ? "true" : undefined,
                                colSpan: k?.colSpan ?? 1,
                                rowSpan: k?.rowSpan ?? 1,
                                style: cellStyleToCss(O.style, j, {
                                  defaultBorder: layout.gridlines
                                    ? "1px solid #d4d4d4"
                                    : "none",
                                }),
                                children: i.jsx("div", {
                                  style: {
                                    maxHeight: z - 1,
                                    ...textWrapToCss(O.style),
                                    position: "relative",
                                  },
                                  children: i.jsx(ConditionalFormatIndicator, {
                                    decoration: k?.suppressContent
                                      ? undefined
                                      : ruleStyles(F),
                                    children: i.jsx(CellValuePreview, {
                                      value: j,
                                      cell: O,
                                      dateSystem:
                                        workbook.datePolicy?.dateSystem,
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
            (sheet.objects ?? []).map((object) => {
              const A = parseRangeAddress(object.anchor),
                k = applyChartDrawingLayout(sheet, object);
              if (
                axisVisibilityIndex.isHidden(A.firstRow) ||
                c.isHidden(A.firstCol) ||
                !p.includes(A.firstRow) ||
                !m.includes(A.firstCol)
              )
                return null;
              let F = layout.headings ? 32 : 0,
                I = layout.headings ? 22 : 0;
              for (const O of m.slice(0, m.indexOf(A.firstCol)))
                F += columnWidthAt(sheet, O);
              for (const O of p.slice(0, p.indexOf(A.firstRow)))
                I += rowHeightAt(sheet, O);
              return i.jsx(
                "div",
                {
                  className: "dg-print-object",
                  "data-print-object": object.id,
                  "data-print-in-area":
                    A.firstRow >= rangeAddress.firstRow &&
                    A.firstRow <= rangeAddress.lastRow &&
                    A.firstCol >= rangeAddress.firstCol &&
                    A.firstCol <= rangeAddress.lastCol
                      ? "true"
                      : "false",
                  "data-print-label": object.label,
                  style: {
                    left: F + k.offsetX,
                    top: I + k.offsetY,
                    width: k.width,
                    height: k.height,
                  },
                  children: i.jsx(DrawingObjectSvg, {
                    object: object,
                    sheet,
                    calc,
                  }),
                },
                object.id,
              );
            }),
          ],
        }),
      }),
      i.jsx("div", {
        className: "dg-print-footer",
        style: {
          bottom: layout.margins.footer * 96,
          left: layout.margins.left * 96,
          right: layout.margins.right * 96,
        },
        children: w(layout.footer),
      }),
    ],
  });
}

function buildPrintPages(workbook, sheetId, t) {
  const sheet = { ...findSheetOrThrow(workbook, sheetId) };
  validatePageSetup(t, sheet);
  try {
    validateDrawingObjects(sheet.objects, sheet);
  } catch (error) {
    throw Error(
      `Unsupported or invalid print object: ${error instanceof Error ? error.message : "check the worksheet drawings."}`,
    );
  }
  const o = zU(sheet, sheet.pageLayout),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    ruleStyles = evaluateConditionalFormatRules(
      sheet,
      workbookWithSettings,
      vx(sheet, workbookWithSettings),
    );
  for (const page of o.pages) {
    const u =
        page.rows ??
        Array.from(
          { length: page.lastRow - page.firstRow + 1 },
          (c, h) => page.firstRow + h,
        ),
      l =
        page.columns ??
        Array.from(
          { length: page.lastCol - page.firstCol + 1 },
          (c, h) => page.firstCol + h,
        );
    for (const c of u)
      for (const h of l) ruleStyles(`${columnIndexToLetters(h)}${c + 1}`);
  }
  return {
    ...o,
    sheet: sheet,
    layout: sheet.pageLayout,
    calc: workbookWithSettings,
    ruleStyles: ruleStyles,
  };
}
function renderPrintDocument(workbook, sheetId, layout) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    printPages = buildPrintPages(
      workbook,
      sheetId,
      layout ??
        sheetOrThrow.pageLayout ?? {
          ...createDefaultPageSetup(sheetOrThrow),
          ...defaultPrintHeaderFooter,
        },
    ),
    a = printPages.pages.map((page, c) =>
      i.jsx(
        PrintPagePreview,
        {
          workbook: workbook,
          sheet: printPages.sheet,
          layout: printPages.layout,
          page: page,
          index: c,
          total: printPages.pages.length,
          scale: printPages.scale,
          calc: printPages.calc,
          ruleStyles: printPages.ruleStyles,
        },
        c,
      ),
    ),
    s = printPages.layout.paper === "a4" ? "A4" : printPages.layout.paper;
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
                  children: `${workbook.name} — ${sheetOrThrow.name} — revision ${workbook.revision}`,
                }),
                i.jsx("meta", {
                  name: "description",
                  content: `Workbook ${workbook.name}; revision ${workbook.revision}; period ${workbook.metadata?.period || "not set"}; ${printPages.pages.length} page(s).`,
                }),
                i.jsx("style", {
                  children: `*{box-sizing:border-box}html,body{margin:0;padding:0;background:white;font-family:Calibri,Arial,sans-serif}${printPaperStyles}@page{size:${s} ${printPages.layout.orientation};margin:0}`,
                }),
              ],
            }),
            i.jsx("body", { children: a }),
          ],
        }),
      ),
    pageCount: printPages.pages.length,
    notices: [
      ...(printPages.notices ?? []),
      ...printPages.ruleStyles
        .errors()
        .map((l) => `Conditional formatting: ${l.message}`),
    ],
  };
}
function boundsOverlap(e, r, t = 0.8) {
  return (
    e.left < r.left - t ||
    e.right > r.right + t ||
    e.top < r.top - t ||
    e.bottom > r.bottom + t
  );
}
function measureRenderedTextBounds(e) {
  if (e.scrollWidth > e.clientWidth + 1 || e.scrollHeight > e.clientHeight + 1)
    return true;
  const ownerDocument = e.ownerDocument,
    t = e.getBoundingClientRect(),
    n = ownerDocument.createTreeWalker(e, 4),
    o = ownerDocument.createRange();
  let a;
  for (; (a = n.nextNode());)
    if (a.textContent?.trim()) {
      o.selectNodeContents(a);
      for (const d of o.getClientRects()) if (boundsOverlap(d, t)) return true;
    }
  return false;
}
function measurePrintLayoutIssues(e) {
  const r = Array.from(e.querySelectorAll(".dg-print-paper")),
    issues = [];
  r.length ||
    issues.push({
      severity: "error",
      code: "empty",
      message: "There are no pages to print. Choose a populated print area.",
    });
  let n = false;
  return (
    r.forEach((o, a) => {
      const page = a + 1,
        d = o.getBoundingClientRect(),
        u = o.querySelector(".dg-print-content"),
        l = o.querySelector(".dg-print-table");
      if (!d.width || !d.height || !u || !l) {
        issues.push({
          severity: "error",
          code: "layout",
          page: page,
          message: `Page ${page}: print layout could not be measured. Keep the preview visible and retry in a supported browser.`,
        });
        return;
      }
      const c = u.getBoundingClientRect();
      boundsOverlap(l.getBoundingClientRect(), c) &&
        issues.push({
          severity: "error",
          code: "layout",
          page: page,
          message: `Page ${page}: rows or columns extend beyond the printable area. Reduce scale, increase margins' available space, or adjust row/column sizes.`,
        });
      for (const m of ["header", "footer"]) {
        const v = o.querySelector(`.dg-print-${m}`);
        if (!v?.textContent?.trim()) continue;
        const g = v.getBoundingClientRect();
        (boundsOverlap(g, d) ||
          measureRenderedTextBounds(v) ||
          (m === "header" ? g.bottom > c.top + 0.8 : g.top < c.bottom - 0.8)) &&
          issues.push({
            severity: "error",
            code: "identity",
            page: page,
            message: `Page ${page}: the ${m} is clipped or overlaps the schedule. Shorten its text or increase the ${m === "header" ? "top" : "bottom"} margin.`,
          });
      }
      const h = [];
      for (const m of o.querySelectorAll("[data-print-cell]")) {
        const firstElementChild = m.firstElementChild;
        if (!firstElementChild?.textContent?.trim()) continue;
        m.dataset.printInArea !== "false" && (n = true);
        const printCell = m.dataset.printCell,
          b = m.dataset.printNumber === "true";
        (b && h.push({ address: printCell, bounds: m.getBoundingClientRect() }),
          measureRenderedTextBounds(firstElementChild) &&
            issues.push({
              severity: b ? "error" : "warning",
              code: b ? "number" : "text",
              page: page,
              address: printCell,
              message: `Page ${page}, ${printCell}: ${b ? "the full number does not fit" : "text is clipped"}. AutoFit or widen the column, increase row height, or reduce the cell font size before printing.`,
            }));
      }
      for (const m of o.querySelectorAll(".dg-print-object")) {
        m.dataset.printInArea !== "false" && (n = true);
        const v = m.dataset.printLabel || m.dataset.printObject || "Drawing",
          g = m.getBoundingClientRect();
        boundsOverlap(g, c) &&
          issues.push({
            severity: "error",
            code: "object",
            page: page,
            message: `Page ${page}: “${v}” crosses the printable boundary. Move or resize it; drawings are not continued onto the next page.`,
          });
        const b = m.querySelector(".dg-shape-label"),
          y = [...m.querySelectorAll("svg text")];
        ((b && measureRenderedTextBounds(b)) ||
          y.some((w) => boundsOverlap(w.getBoundingClientRect(), g))) &&
          issues.push({
            severity: "error",
            code: "object",
            page: page,
            message: `Page ${page}: the label in “${v}” is clipped. Enlarge the object or shorten the label.`,
          });
        for (const w of h)
          Math.min(g.right, w.bounds.right) - Math.max(g.left, w.bounds.left) >
            1 &&
            Math.min(g.bottom, w.bounds.bottom) -
              Math.max(g.top, w.bounds.top) >
              1 &&
            issues.push({
              severity: "error",
              code: "object",
              page: page,
              address: w.address,
              message: `Page ${page}, ${w.address}: “${v}” overlaps a numeric cell. Move the drawing so the amount remains visible.`,
            });
        for (const w of m.querySelectorAll("img"))
          (!w.complete || w.naturalWidth <= 0) &&
            issues.push({
              severity: "error",
              code: "image",
              page: page,
              message: `Page ${page}: picture “${v}” could not be decoded. Replace it before printing.`,
            });
      }
    }),
    r.length &&
      !n &&
      !issues.some((issue) => issue.code === "layout") &&
      issues.unshift({
        severity: "error",
        code: "empty",
        message:
          "The print area is empty: no displayed values or visible objects. Choose a populated print area. Formatting, repeated titles outside the area, and page identifiers alone do not count as schedule content.",
      }),
    { pages: r.length, issues: issues }
  );
}
function summarizePrintIssues(e) {
  const r = e.issues.filter((issue) => issue.severity === "error");
  if (r.length)
    throw Error(
      `${r.length} print issue(s). ${r
        .slice(0, 3)
        .map((t) => t.message)
        .join(" ")}`,
    );
}
async function waitForPrintAssets(e, r = {}) {
  const t = e.ownerDocument ?? e,
    { signal, timeoutMs = 1e4 } = r;
  if (signal?.aborted) throw Error("Print preparation was cancelled.");
  const a = [];
  try {
    const d = Promise.all([
      t.fonts?.ready ?? Promise.resolve(),
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
            (u.addEventListener("load", h, { once: true }),
              u.addEventListener("error", f, { once: true }),
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
            timeoutMs,
          ),
          h = () => l(Error("Print preparation was cancelled."));
        (signal?.addEventListener("abort", h, { once: true }),
          a.push(() => {
            (clearTimeout(c), signal?.removeEventListener("abort", h));
          }));
      }),
    ]);
  } finally {
    a.forEach((d) => d());
  }
}
async function printDocumentInWindow(e, html, t) {
  if (
    (t.assertCurrent(),
    e.document.write(html),
    e.document.close(),
    await waitForPrintAssets(e.document, t),
    t.assertCurrent(),
    e.closed)
  )
    throw Error("The print window was closed. Reopen print preview to retry.");
  const n = measurePrintLayoutIssues(e.document);
  return (summarizePrintIssues(n), e.focus(), e.print(), n);
}
const pdfRenderLimits = {
  pixelRatio: 2,
  maxBytes: 50 * 1024 * 1024,
  pageTimeoutMs: 3e4,
};
function buildScheduleFileName(workbook, sheetId) {
  return `${
    `${workbook.name}-${findSheetOrThrow(workbook, sheetId).name}-rev-${workbook.revision}`
      .replace(/[\x00-\x1f\x7f<>:"/\\|?*]+/g, "-")
      .replace(/[. ]+$/g, "")
      .slice(0, 160) || "Schedule"
  }.pdf`;
}
async function renderSchedulePdf(workbook, sheetId, t) {
  const n = () => {
      if (t.signal?.aborted) throw Error("PDF creation was cancelled.");
      t.assertCurrent();
    },
    o = async (b) => {
      let y, w;
      try {
        return await Promise.race([
          b,
          new Promise((A, k) => {
            ((y = setTimeout(
              () =>
                k(
                  Error(
                    "A PDF page did not finish within 30 seconds. Reduce its complexity or use browser printing.",
                  ),
                ),
              pdfRenderLimits.pageTimeoutMs,
            )),
              (w = () => k(Error("PDF creation was cancelled."))),
              t.signal?.addEventListener("abort", w, { once: true }),
              t.signal?.aborted && w());
          }),
        ]);
      } finally {
        (clearTimeout(y), w && t.signal?.removeEventListener("abort", w));
      }
    };
  if ((n(), typeof document > "u"))
    throw Error("PDF rendering requires a browser with canvas support.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    layout = t.layout ??
      sheetOrThrow.pageLayout ?? {
        ...createDefaultPageSetup(sheetOrThrow),
        ...defaultPrintHeaderFooter,
      },
    d = renderPrintDocument(workbook, sheetId, layout),
    pageDimensions = resolvePageDimensions(layout),
    l = document.createElement("iframe");
  ((l.title = "Preparing schedule PDF"),
    l.setAttribute("aria-hidden", "true"),
    (l.tabIndex = -1),
    l.setAttribute("sandbox", "allow-same-origin"),
    (l.style.cssText = `position:fixed;left:-100000px;top:0;width:${pageDimensions.width}px;height:${pageDimensions.height}px;border:0;pointer-events:none;`));
  const c = [];
  try {
    await new Promise((F, I) => {
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
        t.signal?.addEventListener("abort", O, { once: true }),
        c.push(() => {
          (clearTimeout(j),
            t.signal?.removeEventListener("abort", O),
            (l.onload = null),
            (l.onerror = null));
        }),
        (l.srcdoc = d.html.replace(
          "<head>",
          `<head><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src 'none'">`,
        )),
        document.body.append(l));
    });
    const contentDocument = l.contentDocument;
    if (!contentDocument)
      throw Error(
        "The host does not allow the isolated PDF rendering document.",
      );
    (await waitForPrintAssets(contentDocument, { signal: t.signal }), n());
    const inspection = measurePrintLayoutIssues(contentDocument);
    summarizePrintIssues(inspection);
    const w = [...contentDocument.querySelectorAll(".dg-print-paper")];
    if (
      w.length !== d.pageCount ||
      !w.length ||
      w.length > printRenderLimits.pages
    )
      throw Error(
        "The PDF page plan changed. Reopen print preview and try again.",
      );
    const E = await PDFDocument.create();
    (E.setTitle(`${workbook.name} — ${sheetOrThrow.name}`),
      E.setSubject(
        `Period: ${workbook.metadata?.period || "not set"}; engagement: ${workbook.metadata?.engagement || "not set"}; revision ${workbook.revision}. Image-based schedule snapshot.`,
      ),
      E.setCreator("Ledgerglass schedule PDF"),
      E.setProducer("Ledgerglass / pdf-lib"));
    let A = 0;
    t.onProgress?.(0, w.length);
    for (let F = 0; F < w.length; F++) {
      if (t.signal?.aborted) throw Error("PDF creation was cancelled.");
      const I = await o(
        toCanvas(w[F], {
          width: pageDimensions.width,
          height: pageDimensions.height,
          pixelRatio: pdfRenderLimits.pixelRatio,
          backgroundColor: "#ffffff",
          skipFonts: true,
          fontEmbedCSS: "",
          skipAutoScale: true,
        }),
      );
      try {
        if (t.signal?.aborted) throw Error("PDF creation was cancelled.");
        if (
          I.width !== Math.floor(pageDimensions.width * 2) ||
          I.height !== Math.floor(pageDimensions.height * 2)
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
        if (((A += O.size), A > pdfRenderLimits.maxBytes))
          throw Error(
            "PDF images exceed 50 MB. Reduce the print area or image sizes.",
          );
        const j = await E.embedPng(await O.arrayBuffer()),
          z = E.addPage([
            pageDimensions.width * 0.75,
            pageDimensions.height * 0.75,
          ]);
        z.drawImage(j, {
          x: 0,
          y: 0,
          width: z.getWidth(),
          height: z.getHeight(),
        });
      } finally {
        ((I.width = 0), (I.height = 0));
      }
      t.onProgress?.(F + 1, w.length);
    }
    n();
    const bytes = await E.save();
    if ((n(), bytes.length > pdfRenderLimits.maxBytes))
      throw Error(
        "The PDF exceeds 50 MB. Reduce the print area or image sizes.",
      );
    return {
      bytes: bytes,
      filename: buildScheduleFileName(workbook, sheetId),
      pageCount: w.length,
      inspection: inspection,
      notices: [
        "Image-based PDF at 192 dpi: text is not selectable, searchable or accessibility-tagged. Use browser printing when a text PDF is needed.",
        ...d.notices,
      ],
    };
  } finally {
    (c.forEach((b) => b()), l.remove());
  }
}

export {
  DrawingObjectOverlay,
  renderSchedulePdf,
  printDocumentInWindow,
  PicturePropertiesDialog,
  PrintPagePreview,
  renderPrintDocument,
  waitForPrintAssets,
  summarizePrintIssues,
  measurePrintLayoutIssues,
  printPaperStyles,
  buildPrintPages,
};
