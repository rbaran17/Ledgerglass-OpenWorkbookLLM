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
    [pageIndex, setPageIndex] = ne.useState(0),
    [exportError, setExportError] = ne.useState(""),
    previewRef = ne.useRef(null),
    points = report?.points ?? [],
    currentPage = Math.min(
      pageIndex,
      Math.max(0, Math.ceil(points.length / 20) - 1),
    ),
    exportChart = (isSvgExport = false) => {
      try {
        if ((setExportError(""), !report)) throw Error(error);
        if (isSvgExport) {
          const svgElement = previewRef.current?.querySelector("svg");
          if (!svgElement)
            throw Error("Resolve chart errors before exporting the visual.");
          onExport({
            filename: "Chart.svg",
            mime: "image/svg+xml",
            content: new XMLSerializer().serializeToString(svgElement),
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
        setExportError(error instanceof Error ? error.message : String(error));
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
              ref: previewRef,
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
                children: report.diagnostics.map((diagnostic, index) =>
                  i.jsxs(
                    "li",
                    {
                      children: [diagnostic.severity, ": ", diagnostic.message],
                    },
                    index,
                  ),
                ),
              }),
            i.jsx("button", {
              type: "button",
              onClick: () => exportChart(),
              children: i.jsx(LocalizedText, {
                children: "Export chart inspection",
              }),
            }),
            i.jsx("button", {
              type: "button",
              disabled: report.diagnostics.some(
                (diagnostic) => diagnostic.severity === "error",
              ),
              onClick: () => exportChart(true),
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
                        children: points
                          .slice(currentPage * 20, currentPage * 20 + 20)
                          .map((point) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", { children: point.row }),
                                  i.jsxs("td", {
                                    children: [
                                      point.included ? "Included" : "Excluded",
                                      point.hidden ? " · hidden" : "",
                                    ],
                                  }),
                                  i.jsxs("td", {
                                    children: [
                                      point.category.address,
                                      ":",
                                      " ",
                                      String(point.category.value ?? "(blank)"),
                                    ],
                                  }),
                                  point.values.map((value) =>
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
                              point.row,
                            ),
                          ),
                      }),
                    ],
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: !currentPage,
                  onClick: () => setPageIndex(currentPage - 1),
                  children: i.jsx(LocalizedText, {
                    children: "Previous chart source rows",
                  }),
                }),
                i.jsx("button", {
                  type: "button",
                  disabled: (currentPage + 1) * 20 >= points.length,
                  onClick: () => setPageIndex(currentPage + 1),
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
                          (bin) =>
                            `${bin.lower}–${bin.upper}: ${bin.count} rows ${bin.upperInclusive ? "(upper inclusive)" : "(upper exclusive)"}`,
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
                        .map(
                          (step) =>
                            `${step.label}: ${step.start} → ${step.end}`,
                        )
                        .join("; "),
                    ],
                  }),
              ],
            }),
          ],
        }),
      exportError &&
        i.jsx("p", {
          role: "alert",
          className: "dg-error",
          children: exportError,
        }),
    ],
  });
}
function PicturePropertiesDialog(props) {
  const { text, language } = useWorkbookMessages(),
    cancellableOperation = useCancellableOperation(),
    existingObject = props.workbook.sheets
      .find((sheet) => sheet.id === props.sheetId)
      ?.objects?.find((object) => object.id === props.objectId),
    objectKind =
      existingObject?.kind ??
      (/picture|image/i.test(props.mode)
        ? "picture"
        : /chart/i.test(props.mode)
          ? "chart"
          : "shape"),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    [snapshotRef] = ne.useState(() => ({
      current: F2(props.workbook, props.sheetId, props.range),
    })),
    modeRef = ne.useRef(props.mode),
    objectIdRef = ne.useRef(props.objectId),
    propsRef = ne.useRef(props),
    dialogRef = ne.useRef(null),
    isMountedRef = ne.useRef(true),
    requestIdRef = ne.useRef(0);
  (ne.useRef(null), (propsRef.current = props));
  const [object, setObject] = ne.useState(() =>
      existingObject
        ? structuredClone(existingObject)
        : {
            ...createAnchoredDrawingObject(objectKind, props.range),
            ...(objectKind === "chart"
              ? { chart: "column", range: props.range, headers: true }
              : {}),
          },
    ),
    [sourceLocator, setSourceLocator] = ne.useState(
      existingObject?.evidence?.sourceLocator ?? "",
    ),
    [pictureError, setPictureError] = ne.useState(""),
    [isImporting, setIsImporting] = ne.useState(false),
    isSnapshotStale =
      snapshotRef.current.guard !== JSON.stringify(props.workbook) ||
      snapshotRef.current.sheetId !== props.sheetId ||
      snapshotRef.current.selection !== props.range ||
      modeRef.current !== props.mode ||
      objectIdRef.current !== props.objectId,
    isEditLocked =
      props.readOnly ||
      isSnapshotStale ||
      !sheet ||
      !!sheet?.protection?.enabled,
    assertDrawingCurrent = () => {
      const latestProps = propsRef.current;
      if (
        modeRef.current !== latestProps.mode ||
        objectIdRef.current !== latestProps.objectId
      )
        throw Error("The drawing selection changed. Reopen this dialog.");
      assertDrawingPreviewCurrent(latestProps.workbook, latestProps.sheetId, {
        readOnly: latestProps.readOnly,
        snapshot: snapshotRef.current,
        selection: latestProps.range,
      });
    },
    update = (patch) => {
      !isEditLocked &&
        !isImporting &&
        setObject((previous) => ({ ...previous, ...patch }));
    };
  (ne.useEffect(() => {
    isMountedRef.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        ((isMountedRef.current = false),
          requestIdRef.current++,
          cancellableOperation.cancel(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      isEditLocked &&
        (requestIdRef.current++,
        cancellableOperation.cancel(),
        setIsImporting(false));
    }, [isEditLocked]));
  async function handlePictureFile(file) {
    if (!file) return;
    const requestId = ++requestIdRef.current;
    try {
      if (
        (assertDrawingCurrent(),
        setIsImporting(true),
        setPictureError(""),
        file.size > 2 * 1024 * 1024 ||
          file.size === 0 ||
          !["image/png", "image/jpeg"].includes(file.type))
      )
        throw Error("Choose a PNG or JPEG picture no larger than 2 MB.");
      const { dataURL } = await importPictureFile(
        file,
        cancellableOperation.begin(),
      );
      if (!isMountedRef.current || requestId !== requestIdRef.current) return;
      assertDrawingCurrent();
      const pictureInspection = inspectPictureDataUrl(dataURL);
      if (pictureInspection.bytes.length !== file.size)
        throw Error("The file bytes changed while reading the picture.");
      const pictureEvidence = createPictureEvidence(dataURL, {
          origin: "local-file",
          filename: file.name,
          ...(sourceLocator ? { sourceLocator: sourceLocator } : {}),
        }),
        decodedSize = await new Promise((resolve, reject) => {
          const image = new Image();
          ((image.onload = () =>
            resolve({
              width: image.naturalWidth,
              height: image.naturalHeight,
            })),
            (image.onerror = () =>
              reject(Error("The picture could not be decoded."))),
            (image.src = dataURL));
        });
      if (!isMountedRef.current || requestId !== requestIdRef.current) return;
      if (
        (assertDrawingCurrent(),
        decodedSize.width !== pictureInspection.width ||
          decodedSize.height !== pictureInspection.height)
      )
        throw Error(
          "Decoded picture dimensions do not match its validated header.",
        );
      const scale = Math.min(
        1,
        400 / pictureInspection.width,
        250 / pictureInspection.height,
      );
      setObject((previousObject) => ({
        ...previousObject,
        dataURL,
        evidence: pictureEvidence,
        label: file.name.slice(0, 500),
        width: Math.max(24, Math.round(pictureInspection.width * scale)),
        height: Math.max(24, Math.round(pictureInspection.height * scale)),
      }));
    } catch (error) {
      isMountedRef.current &&
        requestId === requestIdRef.current &&
        setPictureError(
          error instanceof Error ? error.message : "Picture import failed.",
        );
    } finally {
      isMountedRef.current &&
        requestId === requestIdRef.current &&
        setIsImporting(false);
    }
  }
  function onClick() {
    try {
      if ((assertDrawingCurrent(), isImporting)) return;
      const latestProps = propsRef.current,
        sheet = latestProps.workbook.sheets.find(
          (sheet) => sheet.id === latestProps.sheetId,
        );
      if (
        (objectKind === "chart" &&
          (measureChartSelection(object.range, sheet),
          formatDiagnosticsMessage(
            inspectChartSource(
              sheet,
              object,
              calculateWorkbookWithSettings(latestProps.workbook),
            ),
          )),
        objectKind === "picture" && !object.dataURL)
      )
        throw Error("Choose a picture first.");
      const nextObject = structuredClone(object);
      if (
        (objectKind === "picture" &&
          (nextObject.evidence = {
            ...(object.evidence ??
              createPictureEvidence(object.dataURL, {
                origin: "legacy-native",
              })),
            sourceLocator: sourceLocator || undefined,
          }),
        props.objectId &&
          !sheet.objects?.some((object) => object.id === props.objectId))
      )
        throw Error("The drawing no longer exists.");
      const nextObjects = props.objectId
          ? (sheet.objects ?? []).map((drawingObject) =>
              drawingObject.id === props.objectId ? nextObject : drawingObject,
            )
          : [...(sheet.objects ?? []), nextObject],
        updatedWorkbook = setSheetDrawingObjects(
          latestProps.workbook,
          latestProps.sheetId,
          nextObjects,
          {
            readOnly: latestProps.readOnly,
            snapshot: snapshotRef.current,
            selection: latestProps.range,
          },
        );
      (updatedWorkbook !== latestProps.workbook &&
        latestProps.onWorkbookChange(
          updatedWorkbook,
          props.objectId
            ? "Update " + objectKind + " properties"
            : "Insert " + objectKind,
        ),
        latestProps.onClose());
    } catch (error) {
      setPictureError(
        error instanceof Error
          ? error.message
          : "The object could not be saved.",
      );
    }
  }
  const evidence =
      object.evidence ??
      (object.dataURL
        ? createPictureEvidence(object.dataURL, { origin: "legacy-native" })
        : undefined),
    chartInspection = ne.useMemo(() => {
      if (objectKind !== "chart" || !sheet) return { report: null, error: "" };
      try {
        return {
          report: inspectChartSource(
            sheet,
            object,
            calculateWorkbookWithSettings(props.workbook),
          ),
          error: "",
        };
      } catch (error) {
        return {
          report: null,
          error: error instanceof Error ? error.message : String(error),
        };
      }
    }, [objectKind, sheet, object, props.workbook]),
    dialogTitle =
      objectKind[0].toUpperCase() + objectKind.slice(1) + " properties";
  return i.jsx("div", {
    className: "dg-backdrop",
    onKeyDown: (event) => {
      if (
        (event.stopPropagation(),
        event.key === "Escape" && (event.preventDefault(), props.onClose()),
        event.key === "Tab")
      ) {
        const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled)",
            ) ?? []),
          ].filter((element) => !element.closest("fieldset:disabled")),
          firstFocusable = focusableElements[0],
          lastFocusable = focusableElements.at(-1);
        event.shiftKey && document.activeElement === firstFocusable
          ? (event.preventDefault(),
            lastFocusable == null || lastFocusable.focus())
          : !event.shiftKey &&
            document.activeElement === lastFocusable &&
            (event.preventDefault(),
            firstFocusable == null || firstFocusable.focus());
      }
    },
    children: i.jsxs("section", {
      className: `dg-dialog ${objectKind === "chart" ? "dg-chart-dialog" : ""}`,
      ref: dialogRef,
      role: "dialog",
      "aria-modal": "true",
      "aria-label": props.objectId ? dialogTitle : `Insert ${objectKind}`,
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("span", {
              role: "heading",
              "aria-level": 2,
              children: [
                objectKind === "picture"
                  ? i.jsx(ImagePlus, { size: 19 })
                  : objectKind === "chart"
                    ? i.jsx(ChartColumn, { size: 19 })
                    : i.jsx(Shapes, { size: 19 }),
                props.objectId ? dialogTitle : "Insert " + objectKind,
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close drawing tools"),
              onClick: props.onClose,
              lang: language("Close drawing tools"),
              "data-ui-label": "Close drawing tools",
              children: i.jsx(X, { size: 17 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "dg-form",
          children: [
            isSnapshotStale &&
              i.jsx("p", {
                role: "alert",
                className: "dg-error",
                children: i.jsx(LocalizedText, {
                  children:
                    "The workbook or selection changed. Close and reopen this dialog.",
                }),
              }),
            props.readOnly &&
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
              disabled: isEditLocked || isImporting,
              children: [
                objectKind === "picture" &&
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
                            disabled: isEditLocked || isImporting,
                            onChange: (event) =>
                              void handlePictureFile(event.target.files?.[0]),
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
                            onChange: (event) =>
                              setSourceLocator(event.target.value),
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
                objectKind === "shape" &&
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
                objectKind === "chart" &&
                  i.jsx(ChartSettingsPanel, {
                    object: object,
                    report: chartInspection.report,
                    update: update,
                  }),
                i.jsxs("label", {
                  children: [
                    objectKind === "chart" ? "Chart title" : "Label",
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
                            onValueChange: (widthValue) =>
                              update({ width: Number(widthValue) }),
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
                            onValueChange: (heightValue) =>
                              update({ height: Number(heightValue) }),
                            lang: language("Object height"),
                            "data-ui-label": "Object height",
                          }),
                        ],
                      }),
                    ],
                  }),
                objectKind === "shape" &&
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
            objectKind === "chart" &&
              i.jsx(ChartInspectionPanel, {
                ...chartInspection,
                revision: props.workbook.revision,
                workbookId: props.workbook.id,
                onExport: props.onChartExport,
              }),
            objectKind === "picture" &&
              evidence &&
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
                      evidence.origin === "local-file"
                        ? "Local file bytes"
                        : evidence.origin === "xlsx-media"
                          ? "Observed XLSX media"
                          : "Observed embedded bytes; original source unknown",
                  }),
                  i.jsx("dt", {
                    children:
                      evidence.origin === "local-file"
                        ? "Original filename"
                        : "Observed filename",
                  }),
                  i.jsx("dd", {
                    children: evidence.filename ?? "Not available",
                  }),
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, { children: "Byte length" }),
                  }),
                  i.jsx("dd", {
                    children: evidence.byteLength.toLocaleString(),
                  }),
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, { children: "SHA-256" }),
                  }),
                  i.jsx("dd", { children: evidence.sha256 }),
                  i.jsx("dt", {
                    children: i.jsx(LocalizedText, {
                      children: "Source reference",
                    }),
                  }),
                  i.jsx("dd", { children: evidence.sourceReference }),
                  evidence.containerFilename &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Imported workbook",
                          }),
                        }),
                        i.jsx("dd", { children: evidence.containerFilename }),
                      ],
                    }),
                  evidence.mediaReference &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsx("dt", {
                          children: i.jsx(LocalizedText, {
                            children: "Observed media reference",
                          }),
                        }),
                        i.jsx("dd", { children: evidence.mediaReference }),
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
            pictureError &&
              i.jsx("p", {
                role: "alert",
                className: "dg-error",
                children: pictureError,
              }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              className: "dg-primary",
              onClick: onClick,
              disabled: isEditLocked || isImporting,
              children: isImporting
                ? "Reading picture…"
                : props.objectId
                  ? "Save " + objectKind + " properties"
                  : "Insert",
            }),
          ],
        }),
      ],
    }),
  });
}
function DrawingObjectSvg({ object, sheet, calc, textScale = 1 }) {
  const pictureError = ne.useMemo(() => {
    if (object.kind !== "picture") return "";
    try {
      return (inspectPictureDataUrl(object.dataURL ?? ""), "");
    } catch (error) {
      return error instanceof Error ? error.message : "Invalid picture.";
    }
  }, [object.kind, object.dataURL]);
  return object.kind === "picture"
    ? pictureError
      ? i.jsxs("span", {
          role: "img",
          "aria-label": "Invalid picture: " + pictureError,
          children: [
            i.jsx(LocalizedText, { children: "Picture cannot be displayed: " }),
            pictureError,
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
function DrawingObjectOverlay(props) {
  const { text, language } = useWorkbookMessages(),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    propsRef = ne.useRef(props);
  propsRef.current = props;
  const isEditLocked = props.readOnly || !sheet || !!sheet?.protection?.enabled,
    [selectedDrawingId, setSelectedDrawingId] = ne.useState(null),
    [editError, setEditError] = ne.useState(""),
    [objectId, setObjectId] = ne.useState(null),
    [dragState, setDragState] = ne.useState(null),
    dragStateRef = ne.useRef(null);
  dragStateRef.current = dragState;
  function commitDrawingObject(nextObject, historyLabel, snapshot) {
    try {
      const latestProps = propsRef.current,
        sheet = latestProps.workbook.sheets.find(
          (sheet) => sheet.id === latestProps.sheetId,
        );
      if (!sheet?.objects?.some((object) => object.id === nextObject.id))
        throw Error("The selected drawing no longer exists.");
      const updatedWorkbook = setSheetDrawingObjects(
        latestProps.workbook,
        latestProps.sheetId,
        sheet.objects.map((object) =>
          object.id === nextObject.id ? nextObject : object,
        ),
        {
          readOnly: latestProps.readOnly,
          snapshot:
            snapshot ??
            F2(
              latestProps.workbook,
              latestProps.sheetId,
              latestProps.selection,
            ),
          selection: latestProps.selection,
        },
      );
      (updatedWorkbook !== latestProps.workbook &&
        latestProps.onWorkbookChange(updatedWorkbook, historyLabel),
        setEditError(""));
    } catch (error) {
      setEditError(
        error instanceof Error ? error.message : "Could not update object.",
      );
    }
  }
  function deleteDrawing(drawingId) {
    try {
      const latestProps = propsRef.current,
        sheet = latestProps.workbook.sheets.find(
          (sheet) => sheet.id === latestProps.sheetId,
        ),
        updatedWorkbook = setSheetDrawingObjects(
          latestProps.workbook,
          latestProps.sheetId,
          (sheet?.objects ?? []).filter(
            (drawingObject) => drawingObject.id !== drawingId,
          ),
          {
            readOnly: latestProps.readOnly,
            snapshot: F2(
              latestProps.workbook,
              latestProps.sheetId,
              latestProps.selection,
            ),
            selection: latestProps.selection,
          },
        );
      (updatedWorkbook !== latestProps.workbook &&
        latestProps.onWorkbookChange(updatedWorkbook, "Delete drawing"),
        setSelectedDrawingId(null));
    } catch (error) {
      setEditError(
        error instanceof Error ? error.message : "Could not delete drawing.",
      );
    }
  }
  function beginDragEdit(drawingObject, dragType, pointerX, pointerY) {
    try {
      if (
        (assertDrawingPreviewCurrent(props.workbook, props.sheetId, {
          readOnly: props.readOnly,
        }),
        !Number.isFinite(props.scale) || props.scale <= 0)
      )
        throw Error("Invalid drawing zoom.");
      (setDragState({
        id: drawingObject.id,
        type: dragType,
        x: pointerX,
        y: pointerY,
        initial: structuredClone(drawingObject),
        next: structuredClone(drawingObject),
        snapshot: F2(props.workbook, props.sheetId, props.selection),
        scale: props.scale,
      }),
        setEditError(""));
    } catch (error) {
      setEditError(
        error instanceof Error
          ? error.message
          : "Could not start drawing edit.",
      );
    }
  }
  return (
    ne.useEffect(() => {
      (setSelectedDrawingId(null), setDragState(null), setObjectId(null));
    }, [props.sheetId]),
    ne.useEffect(() => {
      if (dragState)
        try {
          if (
            (assertDrawingPreviewCurrent(props.workbook, props.sheetId, {
              readOnly: props.readOnly,
              snapshot: dragState.snapshot,
              selection: props.selection,
            }),
            props.scale !== dragState.scale)
          )
            throw Error("The zoom changed. Start the drawing move again.");
        } catch (error) {
          (setDragState(null),
            setEditError(
              error instanceof Error
                ? error.message
                : "Drawing edit was cancelled.",
            ));
        }
    }, [
      props.workbook,
      props.readOnly,
      props.selection,
      props.scale,
      dragState,
    ]),
    ne.useEffect(() => {
      if (!dragState) return;
      const handlePointerMove = (event) => {
          const drag = dragStateRef.current;
          if (!drag) return;
          const deltaX = (event.clientX - drag.x) / drag.scale,
            deltaY = (event.clientY - drag.y) / drag.scale,
            nextDrag = {
              ...drag,
              next:
                drag.type === "move"
                  ? {
                      ...drag.initial,
                      offsetX: Math.max(
                        -4e3,
                        Math.min(4e3, drag.initial.offsetX + deltaX),
                      ),
                      offsetY: Math.max(
                        -4e3,
                        Math.min(4e3, drag.initial.offsetY + deltaY),
                      ),
                    }
                  : {
                      ...drag.initial,
                      width: Math.max(
                        24,
                        Math.min(4e3, drag.initial.width + deltaX),
                      ),
                      height: Math.max(
                        24,
                        Math.min(4e3, drag.initial.height + deltaY),
                      ),
                    },
            };
          ((dragStateRef.current = nextDrag), setDragState(nextDrag));
        },
        handlePointerUp = () => {
          const drag = dragStateRef.current;
          drag &&
            (commitDrawingObject(
              drag.next,
              drag.type === "move" ? "Move drawing" : "Resize drawing",
              drag.snapshot,
            ),
            (dragStateRef.current = null),
            setDragState(null));
        },
        handlePointerCancel = () => {
          ((dragStateRef.current = null), setDragState(null));
        },
        handleKeyDown = (event) => {
          event.key === "Escape" && handlePointerCancel();
        };
      return (
        window.addEventListener("pointermove", handlePointerMove),
        window.addEventListener("pointerup", handlePointerUp),
        window.addEventListener("pointercancel", handlePointerCancel),
        window.addEventListener("keydown", handleKeyDown),
        () => {
          (window.removeEventListener("pointermove", handlePointerMove),
            window.removeEventListener("pointerup", handlePointerUp),
            window.removeEventListener("pointercancel", handlePointerCancel),
            window.removeEventListener("keydown", handleKeyDown));
        }
      );
    }, [!!dragState]),
    i.jsxs(i.Fragment, {
      children: [
        i.jsxs("div", {
          className: "dg-drawing-overlay",
          role: "group",
          "aria-label": text("Worksheet drawings"),
          lang: language("Worksheet drawings"),
          "data-ui-label": "Worksheet drawings",
          children: [
            (sheet?.objects ?? []).map((storedObject) => {
              const object =
                  dragState?.id === storedObject.id
                    ? dragState.next
                    : storedObject,
                anchorPosition = props.positionForCell(object.anchor),
                disabled =
                  object.kind === "chart" && object.chart === "sparkline",
                drawingLayout = applyChartDrawingLayout(sheet, object);
              return anchorPosition
                ? i.jsxs(
                    "div",
                    {
                      className: `dg-drawing ${selectedDrawingId === object.id ? "dg-drawing-selected" : ""}`,
                      "data-object-id": object.id,
                      style: {
                        left:
                          anchorPosition.left +
                          drawingLayout.offsetX * props.scale,
                        top:
                          anchorPosition.top +
                          drawingLayout.offsetY * props.scale,
                        width: drawingLayout.width * props.scale,
                        height: drawingLayout.height * props.scale,
                      },
                      tabIndex: 0,
                      "aria-label": object.kind + ": " + object.label,
                      onClick: () => setSelectedDrawingId(object.id),
                      onKeyDown: (event) => {
                        (event.stopPropagation(),
                          event.key === "Escape" && setSelectedDrawingId(null),
                          (event.key === "Delete" ||
                            event.key === "Backspace") &&
                            event.target === event.currentTarget &&
                            !isEditLocked &&
                            (event.preventDefault(), deleteDrawing(object.id)));
                      },
                      children: [
                        i.jsx(DrawingObjectSvg, {
                          object: object,
                          sheet: sheet,
                          calc: props.calc,
                          textScale: props.scale,
                        }),
                        selectedDrawingId === object.id &&
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
                                      onClick: () => setObjectId(object.id),
                                      children: "ⓘ",
                                    }),
                                  !isEditLocked &&
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
                                              beginDragEdit(
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
                                            commitDrawingObject(
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
                                          onClick: () =>
                                            deleteDrawing(object.id),
                                          children: i.jsx(Trash2, { size: 14 }),
                                        }),
                                      ],
                                    }),
                                ],
                              }),
                              !isEditLocked &&
                                !disabled &&
                                i.jsx("button", {
                                  className: "dg-resize-object",
                                  "aria-label": "Resize " + object.label,
                                  onPointerDown: (event) => {
                                    (event.preventDefault(),
                                      beginDragEdit(
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
            editError &&
              i.jsx("div", {
                className: "dg-object-error",
                role: "alert",
                children: editError,
              }),
          ],
        }),
        objectId &&
          sheet?.objects?.some((object) => object.id === objectId) &&
          i.jsx(PicturePropertiesDialog, {
            mode: "drawing-properties",
            objectId: objectId,
            workbook: props.workbook,
            sheetId: props.sheetId,
            range:
              props.selection ??
              sheet.objects.find((object) => object.id === objectId).anchor,
            readOnly: props.readOnly,
            onClose: () => setObjectId(null),
            onWorkbookChange: props.onWorkbookChange,
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
    columnVisibilityIndex = createAxisVisibilityIndex(
      sheet.colCount,
      cn(sheet, "column"),
    ),
    pageDimensions = resolvePageDimensions(layout),
    rangeAddress = parseRangeAddress(layout.printArea),
    visibleRows =
      page.rows ??
      Array.from(
        { length: page.lastRow - page.firstRow + 1 },
        (E, rowOffset) => page.firstRow + rowOffset,
      ).filter((rowIndex) => !axisVisibilityIndex.isHidden(rowIndex)),
    visibleColumns =
      page.columns ??
      Array.from(
        { length: page.lastCol - page.firstCol + 1 },
        (E, columnOffset) => page.firstCol + columnOffset,
      ).filter((columnIndex) => !columnVisibilityIndex.isHidden(columnIndex)),
    width = visibleColumns.reduce(
      (totalWidth, column) => totalWidth + columnWidthAt(sheet, column),
      layout.headings ? 32 : 0,
    ),
    rowPositions = new Map(
      visibleRows.map((row, rowPosition) => [row, rowPosition]),
    ),
    rowTops = [0];
  for (const visibleRow of visibleRows)
    rowTops.push(rowTops.at(-1) + rowHeightAt(sheet, visibleRow));
  const mergeInfoByCell = new Map();
  for (const mergeRange of sheet.merges ?? []) {
    const rangeAddress = parseRangeAddress(mergeRange),
      mergeRows = visibleRows.filter(
        (candidateRow) =>
          candidateRow >= rangeAddress.firstRow &&
          candidateRow <= rangeAddress.lastRow,
      ),
      mergeColumns = visibleColumns.filter(
        (candidateColumn) =>
          candidateColumn >= rangeAddress.firstCol &&
          candidateColumn <= rangeAddress.lastCol,
      );
    if (!mergeRows.length || !mergeColumns.length) continue;
    const mergeInfo = {
      firstRow: mergeRows[0],
      lastRow: mergeRows.at(-1),
      firstCol: mergeColumns[0],
      lastCol: mergeColumns.at(-1),
      rowSpan: mergeRows.length,
      colSpan: mergeColumns.length,
      suppressContent:
        axisVisibilityIndex.isHidden(rangeAddress.firstRow) ||
        columnVisibilityIndex.isHidden(rangeAddress.firstCol),
      anchor: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    };
    if (!(
      mergeInfo.firstRow > mergeInfo.lastRow ||
      mergeInfo.firstCol > mergeInfo.lastCol
    ))
      for (const mergeRow of mergeRows)
        for (const mergeColumn of mergeColumns)
          mergeInfoByCell.set(mergeRow * 16384 + mergeColumn, mergeInfo);
  }
  const expandHeaderFooter = (template) =>
    expandHeaderFooterTemplate(template, workbook, sheet, index + 1, total);
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
        children: expandHeaderFooter(layout.header),
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
                    visibleColumns.map((colIndex) =>
                      i.jsx(
                        "col",
                        { style: { width: columnWidthAt(sheet, colIndex) } },
                        colIndex,
                      ),
                    ),
                  ],
                }),
                layout.headings &&
                  i.jsx("thead", {
                    children: i.jsxs("tr", {
                      children: [
                        i.jsx("th", {}),
                        visibleColumns.map((headerColumn) =>
                          i.jsx(
                            "th",
                            { children: columnIndexToLetters(headerColumn) },
                            headerColumn,
                          ),
                        ),
                      ],
                    }),
                  }),
                i.jsx("tbody", {
                  children: visibleRows.map((bodyRow) =>
                    i.jsxs(
                      "tr",
                      {
                        style: { height: rowHeightAt(sheet, bodyRow) },
                        children: [
                          layout.headings &&
                            i.jsx("th", { children: bodyRow + 1 }),
                          visibleColumns.map((bodyColumn) => {
                            const cellMerge = mergeInfoByCell.get(
                              bodyRow * 16384 + bodyColumn,
                            );
                            if (
                              cellMerge &&
                              (bodyRow !== cellMerge.firstRow ||
                                bodyColumn !== cellMerge.firstCol)
                            )
                              return null;
                            const cellAddress =
                                cellMerge?.anchor ??
                                `${columnIndexToLetters(bodyColumn)}${bodyRow + 1}`,
                              cell = cellMerge?.suppressContent
                                ? undefined
                                : sheet.cells[cellAddress],
                              displayCell = {
                                ...cell,
                                raw: cell?.raw ?? null,
                                style: cellMerge?.suppressContent
                                  ? {}
                                  : {
                                      ...cell?.style,
                                      ...ruleStyles(cellAddress).style,
                                    },
                              },
                              cellValue = cellMerge?.suppressContent
                                ? null
                                : calc(sheet.id, cellAddress),
                              spanHeight =
                                rowTops[
                                  rowPositions.get(
                                    cellMerge?.lastRow ?? bodyRow,
                                  ) + 1
                                ] - rowTops[rowPositions.get(bodyRow)];
                            return i.jsx(
                              "td",
                              {
                                "data-print-cell": cellAddress,
                                "data-print-in-area":
                                  bodyRow >= rangeAddress.firstRow &&
                                  bodyRow <= rangeAddress.lastRow &&
                                  bodyColumn >= rangeAddress.firstCol &&
                                  bodyColumn <= rangeAddress.lastCol
                                    ? "true"
                                    : "false",
                                "data-print-number":
                                  typeof cellValue == "number"
                                    ? "true"
                                    : undefined,
                                colSpan: cellMerge?.colSpan ?? 1,
                                rowSpan: cellMerge?.rowSpan ?? 1,
                                style: cellStyleToCss(
                                  displayCell.style,
                                  cellValue,
                                  {
                                    defaultBorder: layout.gridlines
                                      ? "1px solid #d4d4d4"
                                      : "none",
                                  },
                                ),
                                children: i.jsx("div", {
                                  style: {
                                    maxHeight: spanHeight - 1,
                                    ...textWrapToCss(displayCell.style),
                                    position: "relative",
                                  },
                                  children: i.jsx(ConditionalFormatIndicator, {
                                    decoration: cellMerge?.suppressContent
                                      ? undefined
                                      : ruleStyles(cellAddress),
                                    children: i.jsx(CellValuePreview, {
                                      value: cellValue,
                                      cell: displayCell,
                                      dateSystem:
                                        workbook.datePolicy?.dateSystem,
                                    }),
                                  }),
                                }),
                              },
                              bodyColumn,
                            );
                          }),
                        ],
                      },
                      bodyRow,
                    ),
                  ),
                }),
              ],
            }),
            (sheet.objects ?? []).map((object) => {
              const objectAnchor = parseRangeAddress(object.anchor),
                objectLayout = applyChartDrawingLayout(sheet, object);
              if (
                axisVisibilityIndex.isHidden(objectAnchor.firstRow) ||
                columnVisibilityIndex.isHidden(objectAnchor.firstCol) ||
                !visibleRows.includes(objectAnchor.firstRow) ||
                !visibleColumns.includes(objectAnchor.firstCol)
              )
                return null;
              let leftOffset = layout.headings ? 32 : 0,
                topOffset = layout.headings ? 22 : 0;
              for (const precedingColumn of visibleColumns.slice(
                0,
                visibleColumns.indexOf(objectAnchor.firstCol),
              ))
                leftOffset += columnWidthAt(sheet, precedingColumn);
              for (const precedingRow of visibleRows.slice(
                0,
                visibleRows.indexOf(objectAnchor.firstRow),
              ))
                topOffset += rowHeightAt(sheet, precedingRow);
              return i.jsx(
                "div",
                {
                  className: "dg-print-object",
                  "data-print-object": object.id,
                  "data-print-in-area":
                    objectAnchor.firstRow >= rangeAddress.firstRow &&
                    objectAnchor.firstRow <= rangeAddress.lastRow &&
                    objectAnchor.firstCol >= rangeAddress.firstCol &&
                    objectAnchor.firstCol <= rangeAddress.lastCol
                      ? "true"
                      : "false",
                  "data-print-label": object.label,
                  style: {
                    left: leftOffset + objectLayout.offsetX,
                    top: topOffset + objectLayout.offsetY,
                    width: objectLayout.width,
                    height: objectLayout.height,
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
        children: expandHeaderFooter(layout.footer),
      }),
    ],
  });
}

function buildPrintPages(workbook, sheetId, layout) {
  const sheet = { ...findSheetOrThrow(workbook, sheetId) };
  validatePageSetup(layout, sheet);
  try {
    validateDrawingObjects(sheet.objects, sheet);
  } catch (error) {
    throw Error(
      `Unsupported or invalid print object: ${error instanceof Error ? error.message : "check the worksheet drawings."}`,
    );
  }
  const paginated = zU(sheet, sheet.pageLayout),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    ruleStyles = evaluateConditionalFormatRules(
      sheet,
      workbookWithSettings,
      vx(sheet, workbookWithSettings),
    );
  for (const page of paginated.pages) {
    const pageRows =
        page.rows ??
        Array.from(
          { length: page.lastRow - page.firstRow + 1 },
          (c, rowOffset) => page.firstRow + rowOffset,
        ),
      pageColumns =
        page.columns ??
        Array.from(
          { length: page.lastCol - page.firstCol + 1 },
          (c, columnOffset) => page.firstCol + columnOffset,
        );
    for (const pageRow of pageRows)
      for (const pageColumn of pageColumns)
        ruleStyles(`${columnIndexToLetters(pageColumn)}${pageRow + 1}`);
  }
  return {
    ...paginated,
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
    pageElements = printPages.pages.map((page, pageIndex) =>
      i.jsx(
        PrintPagePreview,
        {
          workbook: workbook,
          sheet: printPages.sheet,
          layout: printPages.layout,
          page: page,
          index: pageIndex,
          total: printPages.pages.length,
          scale: printPages.scale,
          calc: printPages.calc,
          ruleStyles: printPages.ruleStyles,
        },
        pageIndex,
      ),
    ),
    paperSize =
      printPages.layout.paper === "a4" ? "A4" : printPages.layout.paper;
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
                  children: `*{box-sizing:border-box}html,body{margin:0;padding:0;background:white;font-family:Calibri,Arial,sans-serif}${printPaperStyles}@page{size:${paperSize} ${printPages.layout.orientation};margin:0}`,
                }),
              ],
            }),
            i.jsx("body", { children: pageElements }),
          ],
        }),
      ),
    pageCount: printPages.pages.length,
    notices: [
      ...(printPages.notices ?? []),
      ...printPages.ruleStyles
        .errors()
        .map((ruleError) => `Conditional formatting: ${ruleError.message}`),
    ],
  };
}
function boundsOverlap(childBounds, containerBounds, tolerance = 0.8) {
  return (
    childBounds.left < containerBounds.left - tolerance ||
    childBounds.right > containerBounds.right + tolerance ||
    childBounds.top < containerBounds.top - tolerance ||
    childBounds.bottom > containerBounds.bottom + tolerance
  );
}
function measureRenderedTextBounds(element) {
  if (
    element.scrollWidth > element.clientWidth + 1 ||
    element.scrollHeight > element.clientHeight + 1
  )
    return true;
  const ownerDocument = element.ownerDocument,
    elementBounds = element.getBoundingClientRect(),
    treeWalker = ownerDocument.createTreeWalker(element, 4),
    textRange = ownerDocument.createRange();
  let textNode;
  for (; (textNode = treeWalker.nextNode());)
    if (textNode.textContent?.trim()) {
      textRange.selectNodeContents(textNode);
      for (const clientRect of textRange.getClientRects())
        if (boundsOverlap(clientRect, elementBounds)) return true;
    }
  return false;
}
function measurePrintLayoutIssues(printDocument) {
  const paperPages = Array.from(
      printDocument.querySelectorAll(".dg-print-paper"),
    ),
    issues = [];
  paperPages.length ||
    issues.push({
      severity: "error",
      code: "empty",
      message: "There are no pages to print. Choose a populated print area.",
    });
  let hasContentInArea = false;
  return (
    paperPages.forEach((paperElement, pageIndex) => {
      const page = pageIndex + 1,
        paperBounds = paperElement.getBoundingClientRect(),
        contentElement = paperElement.querySelector(".dg-print-content"),
        tableElement = paperElement.querySelector(".dg-print-table");
      if (
        !paperBounds.width ||
        !paperBounds.height ||
        !contentElement ||
        !tableElement
      ) {
        issues.push({
          severity: "error",
          code: "layout",
          page: page,
          message: `Page ${page}: print layout could not be measured. Keep the preview visible and retry in a supported browser.`,
        });
        return;
      }
      const contentBounds = contentElement.getBoundingClientRect();
      boundsOverlap(tableElement.getBoundingClientRect(), contentBounds) &&
        issues.push({
          severity: "error",
          code: "layout",
          page: page,
          message: `Page ${page}: rows or columns extend beyond the printable area. Reduce scale, increase margins' available space, or adjust row/column sizes.`,
        });
      for (const region of ["header", "footer"]) {
        const regionElement = paperElement.querySelector(`.dg-print-${region}`);
        if (!regionElement?.textContent?.trim()) continue;
        const regionBounds = regionElement.getBoundingClientRect();
        (boundsOverlap(regionBounds, paperBounds) ||
          measureRenderedTextBounds(regionElement) ||
          (region === "header"
            ? regionBounds.bottom > contentBounds.top + 0.8
            : regionBounds.top < contentBounds.bottom - 0.8)) &&
          issues.push({
            severity: "error",
            code: "identity",
            page: page,
            message: `Page ${page}: the ${region} is clipped or overlaps the schedule. Shorten its text or increase the ${region === "header" ? "top" : "bottom"} margin.`,
          });
      }
      const numericCells = [];
      for (const cellElement of paperElement.querySelectorAll(
        "[data-print-cell]",
      )) {
        const firstElementChild = cellElement.firstElementChild;
        if (!firstElementChild?.textContent?.trim()) continue;
        cellElement.dataset.printInArea !== "false" &&
          (hasContentInArea = true);
        const printCell = cellElement.dataset.printCell,
          isNumberCell = cellElement.dataset.printNumber === "true";
        (isNumberCell &&
          numericCells.push({
            address: printCell,
            bounds: cellElement.getBoundingClientRect(),
          }),
          measureRenderedTextBounds(firstElementChild) &&
            issues.push({
              severity: isNumberCell ? "error" : "warning",
              code: isNumberCell ? "number" : "text",
              page: page,
              address: printCell,
              message: `Page ${page}, ${printCell}: ${isNumberCell ? "the full number does not fit" : "text is clipped"}. AutoFit or widen the column, increase row height, or reduce the cell font size before printing.`,
            }));
      }
      for (const objectElement of paperElement.querySelectorAll(
        ".dg-print-object",
      )) {
        objectElement.dataset.printInArea !== "false" &&
          (hasContentInArea = true);
        const objectLabel =
            objectElement.dataset.printLabel ||
            objectElement.dataset.printObject ||
            "Drawing",
          objectBounds = objectElement.getBoundingClientRect();
        boundsOverlap(objectBounds, contentBounds) &&
          issues.push({
            severity: "error",
            code: "object",
            page: page,
            message: `Page ${page}: “${objectLabel}” crosses the printable boundary. Move or resize it; drawings are not continued onto the next page.`,
          });
        const labelElement = objectElement.querySelector(".dg-shape-label"),
          svgTextElements = [...objectElement.querySelectorAll("svg text")];
        ((labelElement && measureRenderedTextBounds(labelElement)) ||
          svgTextElements.some((svgText) =>
            boundsOverlap(svgText.getBoundingClientRect(), objectBounds),
          )) &&
          issues.push({
            severity: "error",
            code: "object",
            page: page,
            message: `Page ${page}: the label in “${objectLabel}” is clipped. Enlarge the object or shorten the label.`,
          });
        for (const numericCell of numericCells)
          Math.min(objectBounds.right, numericCell.bounds.right) -
            Math.max(objectBounds.left, numericCell.bounds.left) >
            1 &&
            Math.min(objectBounds.bottom, numericCell.bounds.bottom) -
              Math.max(objectBounds.top, numericCell.bounds.top) >
              1 &&
            issues.push({
              severity: "error",
              code: "object",
              page: page,
              address: numericCell.address,
              message: `Page ${page}, ${numericCell.address}: “${objectLabel}” overlaps a numeric cell. Move the drawing so the amount remains visible.`,
            });
        for (const imageElement of objectElement.querySelectorAll("img"))
          (!imageElement.complete || imageElement.naturalWidth <= 0) &&
            issues.push({
              severity: "error",
              code: "image",
              page: page,
              message: `Page ${page}: picture “${objectLabel}” could not be decoded. Replace it before printing.`,
            });
      }
    }),
    paperPages.length &&
      !hasContentInArea &&
      !issues.some((issue) => issue.code === "layout") &&
      issues.unshift({
        severity: "error",
        code: "empty",
        message:
          "The print area is empty: no displayed values or visible objects. Choose a populated print area. Formatting, repeated titles outside the area, and page identifiers alone do not count as schedule content.",
      }),
    { pages: paperPages.length, issues: issues }
  );
}
function summarizePrintIssues(result) {
  const errorIssues = result.issues.filter(
    (issue) => issue.severity === "error",
  );
  if (errorIssues.length)
    throw Error(
      `${errorIssues.length} print issue(s). ${errorIssues
        .slice(0, 3)
        .map((issue) => issue.message)
        .join(" ")}`,
    );
}
async function waitForPrintAssets(root, options = {}) {
  const targetDocument = root.ownerDocument ?? root,
    { signal, timeoutMs = 1e4 } = options;
  if (signal?.aborted) throw Error("Print preparation was cancelled.");
  const cleanups = [];
  try {
    const assetsReady = Promise.all([
      targetDocument.fonts?.ready ?? Promise.resolve(),
      ...Array.from(root.querySelectorAll("img")).map(
        (imageElement) =>
          new Promise((resolve, reject) => {
            const onLoaded = () =>
              imageElement.naturalWidth > 0
                ? resolve()
                : reject(
                    Error(
                      `Picture “${imageElement.alt || "Untitled"}” could not be decoded. Replace it before printing.`,
                    ),
                  );
            if (imageElement.complete) {
              onLoaded();
              return;
            }
            const onLoadError = () =>
              reject(
                Error(
                  `Picture “${imageElement.alt || "Untitled"}” could not be loaded. Replace it before printing.`,
                ),
              );
            (imageElement.addEventListener("load", onLoaded, { once: true }),
              imageElement.addEventListener("error", onLoadError, {
                once: true,
              }),
              cleanups.push(() => {
                (imageElement.removeEventListener("load", onLoaded),
                  imageElement.removeEventListener("error", onLoadError));
              }));
          }),
      ),
    ]);
    await Promise.race([
      assetsReady,
      new Promise((u, raceReject) => {
        const timeoutId = setTimeout(
            () =>
              raceReject(
                Error(
                  "Print assets did not finish loading within 10 seconds. Check the pictures and fonts, then retry.",
                ),
              ),
            timeoutMs,
          ),
          onAbort = () => raceReject(Error("Print preparation was cancelled."));
        (signal?.addEventListener("abort", onAbort, { once: true }),
          cleanups.push(() => {
            (clearTimeout(timeoutId),
              signal?.removeEventListener("abort", onAbort));
          }));
      }),
    ]);
  } finally {
    cleanups.forEach((cleanup) => cleanup());
  }
}
async function printDocumentInWindow(printWindow, html, printRequest) {
  if (
    (printRequest.assertCurrent(),
    printWindow.document.write(html),
    printWindow.document.close(),
    await waitForPrintAssets(printWindow.document, printRequest),
    printRequest.assertCurrent(),
    printWindow.closed)
  )
    throw Error("The print window was closed. Reopen print preview to retry.");
  const layoutReport = measurePrintLayoutIssues(printWindow.document);
  return (
    summarizePrintIssues(layoutReport),
    printWindow.focus(),
    printWindow.print(),
    layoutReport
  );
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
async function renderSchedulePdf(workbook, sheetId, options) {
  const throwIfCancelled = () => {
      if (options.signal?.aborted) throw Error("PDF creation was cancelled.");
      options.assertCurrent();
    },
    withPageTimeout = async (pagePromise) => {
      let timeoutId, onAbort;
      try {
        return await Promise.race([
          pagePromise,
          new Promise((resolve, reject) => {
            ((timeoutId = setTimeout(
              () =>
                reject(
                  Error(
                    "A PDF page did not finish within 30 seconds. Reduce its complexity or use browser printing.",
                  ),
                ),
              pdfRenderLimits.pageTimeoutMs,
            )),
              (onAbort = () => reject(Error("PDF creation was cancelled."))),
              options.signal?.addEventListener("abort", onAbort, {
                once: true,
              }),
              options.signal?.aborted && onAbort());
          }),
        ]);
      } finally {
        (clearTimeout(timeoutId),
          onAbort && options.signal?.removeEventListener("abort", onAbort));
      }
    };
  if ((throwIfCancelled(), typeof document > "u"))
    throw Error("PDF rendering requires a browser with canvas support.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    layout = options.layout ??
      sheetOrThrow.pageLayout ?? {
        ...createDefaultPageSetup(sheetOrThrow),
        ...defaultPrintHeaderFooter,
      },
    printDocument = renderPrintDocument(workbook, sheetId, layout),
    pageDimensions = resolvePageDimensions(layout),
    renderFrame = document.createElement("iframe");
  ((renderFrame.title = "Preparing schedule PDF"),
    renderFrame.setAttribute("aria-hidden", "true"),
    (renderFrame.tabIndex = -1),
    renderFrame.setAttribute("sandbox", "allow-same-origin"),
    (renderFrame.style.cssText = `position:fixed;left:-100000px;top:0;width:${pageDimensions.width}px;height:${pageDimensions.height}px;border:0;pointer-events:none;`));
  const cleanupCallbacks = [];
  try {
    await new Promise((resolveLoad, rejectLoad) => {
      const onFrameAbort = () =>
          rejectLoad(Error("PDF creation was cancelled.")),
        frameTimeoutId = setTimeout(
          () =>
            rejectLoad(
              Error(
                "The PDF rendering frame could not load. Check the embedding host's content policy.",
              ),
            ),
          1e4,
        );
      ((renderFrame.onload = () => resolveLoad()),
        (renderFrame.onerror = () =>
          rejectLoad(Error("The PDF rendering frame could not load."))),
        options.signal?.addEventListener("abort", onFrameAbort, { once: true }),
        cleanupCallbacks.push(() => {
          (clearTimeout(frameTimeoutId),
            options.signal?.removeEventListener("abort", onFrameAbort),
            (renderFrame.onload = null),
            (renderFrame.onerror = null));
        }),
        (renderFrame.srcdoc = printDocument.html.replace(
          "<head>",
          `<head><meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src 'none'">`,
        )),
        document.body.append(renderFrame));
    });
    const contentDocument = renderFrame.contentDocument;
    if (!contentDocument)
      throw Error(
        "The host does not allow the isolated PDF rendering document.",
      );
    (await waitForPrintAssets(contentDocument, { signal: options.signal }),
      throwIfCancelled());
    const inspection = measurePrintLayoutIssues(contentDocument);
    summarizePrintIssues(inspection);
    const paperElements = [
      ...contentDocument.querySelectorAll(".dg-print-paper"),
    ];
    if (
      paperElements.length !== printDocument.pageCount ||
      !paperElements.length ||
      paperElements.length > printRenderLimits.pages
    )
      throw Error(
        "The PDF page plan changed. Reopen print preview and try again.",
      );
    const pdfDocument = await PDFDocument.create();
    (pdfDocument.setTitle(`${workbook.name} — ${sheetOrThrow.name}`),
      pdfDocument.setSubject(
        `Period: ${workbook.metadata?.period || "not set"}; engagement: ${workbook.metadata?.engagement || "not set"}; revision ${workbook.revision}. Image-based schedule snapshot.`,
      ),
      pdfDocument.setCreator("Ledgerglass schedule PDF"),
      pdfDocument.setProducer("Ledgerglass / pdf-lib"));
    let totalImageBytes = 0;
    options.onProgress?.(0, paperElements.length);
    for (let pageIndex = 0; pageIndex < paperElements.length; pageIndex++) {
      if (options.signal?.aborted) throw Error("PDF creation was cancelled.");
      const pageCanvas = await withPageTimeout(
        toCanvas(paperElements[pageIndex], {
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
        if (options.signal?.aborted) throw Error("PDF creation was cancelled.");
        if (
          pageCanvas.width !== Math.floor(pageDimensions.width * 2) ||
          pageCanvas.height !== Math.floor(pageDimensions.height * 2)
        )
          throw Error(
            "The browser changed the PDF page resolution. Use browser printing instead.",
          );
        const pageBlob = await withPageTimeout(
          new Promise((resolveBlob, rejectBlob) =>
            pageCanvas.toBlob(
              (blob) =>
                blob
                  ? resolveBlob(blob)
                  : rejectBlob(
                      Error("The browser could not render a PDF page."),
                    ),
              "image/png",
            ),
          ),
        );
        if (
          ((totalImageBytes += pageBlob.size),
          totalImageBytes > pdfRenderLimits.maxBytes)
        )
          throw Error(
            "PDF images exceed 50 MB. Reduce the print area or image sizes.",
          );
        const pdfImage = await pdfDocument.embedPng(
            await pageBlob.arrayBuffer(),
          ),
          pdfPage = pdfDocument.addPage([
            pageDimensions.width * 0.75,
            pageDimensions.height * 0.75,
          ]);
        pdfPage.drawImage(pdfImage, {
          x: 0,
          y: 0,
          width: pdfPage.getWidth(),
          height: pdfPage.getHeight(),
        });
      } finally {
        ((pageCanvas.width = 0), (pageCanvas.height = 0));
      }
      options.onProgress?.(pageIndex + 1, paperElements.length);
    }
    throwIfCancelled();
    const bytes = await pdfDocument.save();
    if ((throwIfCancelled(), bytes.length > pdfRenderLimits.maxBytes))
      throw Error(
        "The PDF exceeds 50 MB. Reduce the print area or image sizes.",
      );
    return {
      bytes: bytes,
      filename: buildScheduleFileName(workbook, sheetId),
      pageCount: paperElements.length,
      inspection: inspection,
      notices: [
        "Image-based PDF at 192 dpi: text is not selectable, searchable or accessibility-tagged. Use browser printing when a text PDF is needed.",
        ...printDocument.notices,
      ],
    };
  } finally {
    (cleanupCallbacks.forEach((cleanup) => cleanup()), renderFrame.remove());
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
