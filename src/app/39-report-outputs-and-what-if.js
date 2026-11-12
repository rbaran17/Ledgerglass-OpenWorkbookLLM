import { isFiniteNumber } from "./03-workbook-model.js";
import { downloadBlobAsFile } from "./09-workbook-links.js";
import { isCellInRange } from "./36-watch-window-and-goal-seek.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  createWorksheetRecord,
  validateWhatIfDefinition,
  whatIfLimits,
  expandWhatIfScenarios,
  calculateWorkbookWithSettings,
  cellAddressKey,
  _t,
  cn,
  parseRangeAddress,
  validateDataValidationRule,
  resolveDefinedNameValue,
  normalizeWorkbookInput,
  columnIndexToLetters,
  matrixResultFunctionNames,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  validateLocalizedNumberEntry,
} from "./04-cell-input-paste-and-fill.js";
import { importAutomationFile } from "./07-file-import-and-worker-dispatch.js";
import {
  retainEmbeddedObjectInventory,
  inspectEmbeddedObjects,
} from "./08-xlsx-import-export-and-preservation.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import { inspectFormulaDependencies } from "./12-consolidation-and-dependency-inspection.js";
import {
  describeReportOutputStatus,
  publishReportOutput,
  buildReportHtmlFile,
  buildReportSvgFile,
  refreshReportOutput,
  ReportOutputView,
  removeReportOutput,
  renderReportSvg,
  renderReportPngBlob,
  validateReportSourceScope,
} from "./38-data-analysis-dialog.js";

function ReportOutputsDialog(e) {
  const { text } = useWorkbookMessages(),
    t = ne.useId(),
    n = ne.useRef(null),
    cancellableOperation = useCancellableOperation(),
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
    [checked, h] = ne.useState(false),
    [f, p] = ne.useState(false),
    [m, v] = ne.useState(null),
    [g, b] = ne.useState(),
    [y, w] = ne.useState(null),
    [E, A] = ne.useState(""),
    [k, F] = ne.useState(""),
    [I, O] = ne.useState(""),
    [j, z] = ne.useState(""),
    [inspection, D] = ne.useState(null),
    [disabled, _] = ne.useState(false),
    C = {
      readOnly: e.readOnly,
      allowedSheetIds: e.allowedSheetIds,
      includeHidden: !!e.allowHidden && checked,
    },
    T = JSON.stringify([C, e.filter]),
    N = !!m && y === e.workbook && E === T,
    R = e.workbook.sheets.find((sheet) => sheet.id === u.sheetId),
    L = e.workbook.reportOutputs?.find((reportOutput) => reportOutput.id === j),
    $ = L ? describeReportOutputStatus(e.workbook, L.id, C) : null,
    U =
      s === "Saved outputs"
        ? $?.status !== "unavailable"
          ? L?.snapshot
          : undefined
        : N
          ? m?.snapshot
          : undefined;
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      n.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const G = (ae) => {
      try {
        (F(""), O(""), ae());
      } catch (error) {
        F(error instanceof Error ? error.message : String(error));
      }
    },
    H = (ae) => {
      (l((oe) => ({ ...oe, ...ae })), b(undefined), v(null), F(""));
    },
    Q = (request, refreshId) =>
      G(() => {
        var Se;
        const de = refreshReportOutput(e.workbook, request, C, refreshId);
        (v(de),
          b(refreshId),
          l(de.request),
          w(e.workbook),
          A(T),
          d("Publish"),
          ((Se = n.current?.querySelector(".lg-reports-scroll")) == null
            ? undefined
            : Se.scrollTo
          )?.call(Se, 0, 0));
      }),
    X$1 = async (ae, oe) => {
      (_(true), F(""));
      try {
        const de = a.current,
          he = () =>
            validateReportSourceScope(a.current.workbook, oe, {
              readOnly: a.current.readOnly,
              allowedSheetIds: a.current.allowedSheetIds,
              includeHidden: !!a.current.allowHidden && checked,
            });
        he();
        const Se =
          ae === "png"
            ? await renderReportPngBlob(oe)
            : (() => {
                const le =
                  ae === "svg"
                    ? renderReportSvg(oe)
                    : ae === "html"
                      ? buildReportHtmlFile(oe)
                      : buildReportSvgFile(oe);
                return {
                  filename: le.filename,
                  blob: new Blob([le.content], { type: le.mime }),
                };
              })();
        if ((he(), de.onDownload !== a.current.onDownload))
          throw Error("The report destination changed. Export again.");
        (await (a.current.onDownload ?? downloadBlobAsFile)(Se),
          O(`${ae.toUpperCase()} snapshot exported.`));
      } catch (error) {
        F(error.message);
      } finally {
        _(false);
      }
    },
    J = async (ae) => {
      (_(true), D(null), F(""));
      try {
        const oe = await importAutomationFile(ae, cancellableOperation.begin());
        (inspectEmbeddedObjects(oe), D(oe));
      } catch (error) {
        F(error.message);
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
      onKeyDown: (event) => {
        var oe, de;
        if (
          (event.key === "Escape" && (event.stopPropagation(), e.onClose()),
          event.key === "Tab")
        ) {
          const he = [
            ...n.current.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
            ),
          ];
          event.shiftKey && document.activeElement === he[0]
            ? (event.preventDefault(), (oe = he.at(-1)) == null || oe.focus())
            : !event.shiftKey &&
              document.activeElement === he.at(-1) &&
              (event.preventDefault(), (de = he[0]) == null || de.focus());
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
                  children: i.jsx(LocalizedText, {
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
                            onChange: (event) =>
                              H({ name: event.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Source worksheet",
                          i.jsx("select", {
                            value: u.sheetId,
                            onChange: (event) => {
                              (H({
                                sheetId: event.target.value,
                                kind: "range",
                                objectId: undefined,
                              }),
                                p(false));
                            },
                            children: e.workbook.sheets
                              .filter(
                                (sheet) =>
                                  (!e.allowedSheetIds ||
                                    e.allowedSheetIds.includes(sheet.id)) &&
                                  (C.includeHidden ||
                                    (!sheet.hidden && !sheet.veryHidden)),
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
                            onChange: (event) => {
                              const oe = R?.objects?.find(
                                (object) => object.id === event.target.value,
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
                              R?.objects?.map((object) =>
                                i.jsxs(
                                  "option",
                                  {
                                    value: object.id,
                                    children: [object.kind, ": ", object.label],
                                  },
                                  object.id,
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
                            onChange: (event) =>
                              H({ range: event.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Reuse mode",
                          i.jsxs("select", {
                            value: u.mode,
                            onChange: (event) =>
                              H({ mode: event.target.value }),
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
                            onChange: (event) =>
                              H({ population: event.target.value }),
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
                              onChange: (event) => {
                                (p(event.target.checked),
                                  H({
                                    filter: event.target.checked
                                      ? e.filter
                                      : undefined,
                                    ...(event.target.checked
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
                              checked: checked,
                              onChange: (event) => {
                                (h(event.target.checked), v(null));
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
                    disabled: disabled,
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
                          children: i.jsx(ReportOutputView, {
                            snapshot: m.snapshot,
                          }),
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
                        onChange: (event) => z(event.target.value),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Choose a saved output",
                          }),
                          e.workbook.reportOutputs?.map((reportOutput) =>
                            i.jsxs(
                              "option",
                              {
                                value: reportOutput.id,
                                children: [
                                  reportOutput.snapshot.request.name,
                                  " · v",
                                  reportOutput.version,
                                ],
                              },
                              reportOutput.id,
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
                                      disabled: disabled,
                                      onClick: () => {
                                        (p(!!L.snapshot.request.filter),
                                          Q(L.snapshot.request, L.id));
                                      },
                                      children: "Preview linked refresh",
                                    }),
                                  i.jsx("button", {
                                    disabled: e.readOnly || disabled,
                                    onClick: () =>
                                      G(() => {
                                        (e.onWorkbookChange(
                                          removeReportOutput(
                                            e.workbook,
                                            L.id,
                                            C,
                                          ),
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
                                children: i.jsx(ReportOutputView, {
                                  snapshot: L.snapshot,
                                }),
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
                        disabled: disabled,
                        accept: ".xlsx,.xlsm,.xlam,.xltm,.xltx,.xls,.bin",
                        onChange: (event) => {
                          const oe = event.target.files?.[0];
                          (oe && J(oe), (event.target.value = ""));
                        },
                      }),
                    ],
                  }),
                  inspection &&
                    i.jsxs("section", {
                      children: [
                        i.jsx("h3", { children: inspection.fileName }),
                        i.jsx("p", { children: inspection.sha256 }),
                        i.jsx(ObjectInventoryTable, {
                          items: inspectEmbeddedObjects(inspection).items,
                        }),
                        i.jsx("button", {
                          disabled: e.readOnly || disabled,
                          onClick: () =>
                            G(() => {
                              (e.onWorkbookChange(
                                retainEmbeddedObjectInventory(
                                  e.workbook,
                                  inspection,
                                  {
                                    readOnly: e.readOnly,
                                  },
                                ),
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
                  e.workbook.embeddedObjectInventories?.map(
                    (embeddedObjectInventory) =>
                      i.jsxs(
                        "section",
                        {
                          children: [
                            i.jsx("h3", {
                              children: embeddedObjectInventory.fileName,
                            }),
                            i.jsxs("p", {
                              children: [
                                "Inspected ",
                                embeddedObjectInventory.inspectedAt,
                                " · execution: never",
                              ],
                            }),
                            i.jsx(ObjectInventoryTable, {
                              items: embeddedObjectInventory.items,
                            }),
                            i.jsx("button", {
                              disabled: disabled,
                              onClick: () =>
                                void (e.onDownload ?? downloadBlobAsFile)({
                                  filename: "Embedded-object-inventory.json",
                                  blob: new Blob(
                                    [
                                      JSON.stringify(
                                        embeddedObjectInventory,
                                        null,
                                        2,
                                      ),
                                    ],
                                    {
                                      type: "application/json",
                                    },
                                  ),
                                }),
                              children: "Export inventory JSON",
                            }),
                          ],
                        },
                        embeddedObjectInventory.id,
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
                      disabled: disabled,
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
                    disabled: !N || e.readOnly || disabled,
                    onClick: () =>
                      G(() => {
                        const ae = !!m?.refreshId,
                          oe = publishReportOutput(e.workbook, m, C);
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
                    children: m?.refreshId
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
function ObjectInventoryTable({ items }) {
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
          children: items.map((item, t) =>
            i.jsxs(
              "tr",
              {
                children: [
                  i.jsx("td", { children: item.path }),
                  i.jsx("td", { children: item.kind }),
                  i.jsx("td", { children: item.bytes }),
                  i.jsx("td", { children: item.execution }),
                  i.jsx("td", { children: item.note }),
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

function resolveWhatIfHostScope(workbook, r) {
  if (r.includeHidden !== undefined && typeof r.includeHidden != "boolean")
    throw Error("Include hidden must be boolean.");
  if (
    r.allowedSheetIds !== undefined &&
    (!Array.isArray(r.allowedSheetIds) ||
      r.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("What-if host scope contains an unknown worksheet.");
  return {
    includeHidden: !!r.includeHidden,
    allowedSheetIds: [
      ...new Set(r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id)),
    ].sort(),
  };
}
function isVisibleAuthorizedCell(workbook, r, whatIfHostScope) {
  const n = workbook.sheets.find((sheet) => sheet.id === r.sheetId);
  if (!n || !whatIfHostScope.allowedSheetIds.includes(n.id)) return false;
  const o = _t(r.address);
  return (
    whatIfHostScope.includeHidden ||
    (!n.hidden &&
      !n.veryHidden &&
      !cn(n, "row").some((a) => a.start <= o.row && o.row <= a.end) &&
      !cn(n, "column").some((a) => a.start <= o.col && o.col <= a.end))
  );
}
const whatIfAllowedFunctions = new Set([
  ...matrixResultFunctionNames,
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
function computeWhatIfResult(workbook, request, t = {}) {
  const n = validateWhatIfDefinition(request),
    whatIfHostScope = resolveWhatIfHostScope(workbook, t),
    a = expandWhatIfScenarios(n);
  for (const f of [...n.inputs, ...n.outputs])
    if (!isVisibleAuthorizedCell(workbook, f, whatIfHostScope))
      throw Error(
        "A what-if source is missing or outside the visible, authorized host context.",
      );
  const s = n.inputs.map((input, p) => {
      const m = workbook.sheets.find((sheet) => sheet.id === input.sheetId),
        cell = m.cells[input.address],
        g = _t(input.address);
      if (!cell || !isFiniteNumber(cell.raw))
        throw Error(
          "Changing cells must contain finite numeric inputs, not formulas, text, blanks or spill outputs.",
        );
      if (m.protection?.enabled)
        throw Error(
          "A changing worksheet is protected. Unprotect it before analyzing scenarios.",
        );
      if (m.merges?.some((merge) => isCellInRange(merge, input.address)))
        throw Error("Merged cells cannot be what-if inputs.");
      if (
        [...(m.pivots ?? []), ...(m.queries ?? [])].some((E) =>
          isCellInRange(E.outputRange, input.address),
        )
      )
        throw Error("Generated pivot/query results cannot be what-if inputs.");
      for (const E of m.tables ?? []) {
        if (!isCellInRange(E.range, input.address)) continue;
        const rangeAddress = parseRangeAddress(E.range),
          k = E.columns[g.col - rangeAddress.firstCol];
        if (
          g.row === rangeAddress.firstRow ||
          (E.showTotals && g.row === rangeAddress.lastRow) ||
          k?.calculatedFormula ||
          Object.entries(m.cells).some(([F, I]) => {
            const O = _t(F);
            return (
              O.col === g.col &&
              O.row > rangeAddress.firstRow &&
              O.row <= rangeAddress.lastRow - (E.showTotals ? 1 : 0) &&
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
      const b = (m.validations ?? []).filter((E) =>
        isCellInRange(E.range, input.address),
      );
      if (
        [cell.raw, ...a.map((E) => E.inputs[p])].some((E) =>
          b.some((A) => !validateDataValidationRule(A, E)),
        )
      )
        throw Error(
          `An assumption for ${m.name}!${input.address} violates its data validation constraints. Baseline unchanged.`,
        );
      return cell.raw;
    }),
    d = new Map();
  for (const output of n.outputs) {
    const cell = workbook.sheets.find((sheet) => sheet.id === output.sheetId)
      .cells[output.address];
    if (
      !cell ||
      cell.literal ||
      typeof cell.raw != "string" ||
      !cell.raw.startsWith("=")
    )
      throw Error(
        "Every result cell must contain an actual calculation formula.",
      );
    const m = inspectFormulaDependencies(
        {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
          ),
        },
        {
          ...output,
          direction: "precedents",
          depth: 8,
          ...whatIfHostScope,
          limits: {
            maxFormulaCells: 5e3,
            maxNodes: whatIfLimits.dependencyCells,
            maxReferences: 1e4,
          },
        },
      ),
      v = [m.target, ...m.nodes],
      g = new Set(v.map((event) => event.key));
    if (m.coverage.truncated || m.coverage.depthLimited)
      throw Error(
        "What-if analysis requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
      );
    const b = m.issues.find((issue) => g.has(issue.from));
    if (b || m.cycles.length)
      throw Error(
        `What-if analysis requires supported static dependencies: ${b?.message ?? "a cycle was found."}`,
      );
    for (const y of v) d.set(y.key, { sheetId: y.sheetId, address: y.address });
  }
  if (d.size > whatIfLimits.dependencyCells)
    throw Error(
      "What-if analysis supports at most 500 dependency cells across all results.",
    );
  if (n.inputs.some((input) => !d.has(cellAddressKey(input))))
    throw Error(
      "Each changing input must be a precedent of at least one selected formula result.",
    );
  const assumptions = [...d.values()].map((f) => {
      if (!isVisibleAuthorizedCell(workbook, f, whatIfHostScope))
        throw Error(
          "A formula dependency is outside the visible, authorized host context.",
        );
      const p = workbook.sheets.find((sheet) => sheet.id === f.sheetId),
        cell = p.cells[f.address];
      if (!cell || cell.raw === null)
        throw Error(
          "What-if analysis requires populated static dependencies; blank or possible spill-child cells are unsupported.",
        );
      if (
        !cell.literal &&
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=")
      ) {
        const v = resolveDefinedNameValue(cell.raw, workbook, p.id).replace(
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
            .some((b) => whatIfAllowedFunctions.has(b)) ||
          /[{}#]/.test(v)
        )
          throw Error(
            "What-if analysis excludes volatile, dynamic-reference, array/spill and unsupported reusable-function dependencies.",
          );
      }
      return {
        ...f,
        sheetName: p.name,
        raw: cell.raw,
        ...(cell.literal === undefined ? {} : { literal: cell.literal }),
      };
    }),
    workUnits = d.size * (a.length + 1);
  if (workUnits > whatIfLimits.workUnits)
    throw Error(
      "What-if evaluation would exceed 50,000 dependency-cell work units. Reduce the values, scenarios or dependencies.",
    );
  const c = {
    ...workbook,
    sheets: workbook.sheets
      .filter((sheet) => whatIfHostScope.allowedSheetIds.includes(sheet.id))
      .map((f) => ({
        ...f,
        cells: Object.fromEntries(
          Object.entries(f.cells).filter(([address]) =>
            d.has(cellAddressKey({ sheetId: f.id, address: address })),
          ),
        ),
      })),
  };
  function h(f, p) {
    try {
      const m = new Map(
          n.inputs.map((input, w) => [cellAddressKey(input), p[w]]),
        ),
        v = {
          ...c,
          sheets: c.sheets.map((sheet) => ({
            ...sheet,
            cells: Object.fromEntries(
              Object.entries(sheet.cells).map(([address, E]) => {
                const raw = m.get(
                  cellAddressKey({ sheetId: sheet.id, address: address }),
                );
                return [address, raw === undefined ? E : { ...E, raw: raw }];
              }),
            ),
          })),
        },
        workbookWithSettings = calculateWorkbookWithSettings(v),
        outputs = n.outputs.map((output) =>
          workbookWithSettings(output.sheetId, output.address),
        );
      if (workbookWithSettings.getOperationalErrors().length)
        throw Error(
          `Operational calculation error: ${workbookWithSettings.getOperationalErrors().join(", ")}. An error handler cannot conceal this failure.`,
        );
      if (outputs.some((output) => !isFiniteNumber(output)))
        throw Error(
          `Each formula must return a finite number; received ${outputs.map(String).join(", ")}.`,
        );
      for (const y of d.values())
        if (workbookWithSettings.getSpillRange(y.sheetId, y.address))
          throw Error(
            "A trial produced an array/spill; scalar results are required.",
          );
      if (workbookWithSettings.getOperationalErrors().length)
        throw Error(
          "Dependency inspection encountered an operational calculation error.",
        );
      return { name: f, inputs: [...p], outputs: outputs, error: null };
    } catch (error) {
      return {
        name: f,
        inputs: [...p],
        outputs: null,
        error: (error instanceof Error
          ? error.message
          : "Trial calculation failed."
        ).slice(0, 2e3),
      };
    }
  }
  return {
    snapshot: JSON.stringify(workbook),
    requestSnapshot: JSON.stringify(n),
    request: n,
    context: whatIfHostScope,
    baseline: h("Baseline", s),
    trials: a.map((f) => h(f.name, f.inputs)),
    assumptions: assumptions,
    sourceRevision: workbook.revision,
    dependencyCells: d.size,
    workUnits: workUnits,
  };
}
function createWhatIfPreviewSnapshot(workbook, r, t) {
  if (JSON.stringify(workbook) !== r.snapshot)
    throw Error(
      "The workbook changed after preview. Preview and review the analysis again.",
    );
  const whatIfResult = computeWhatIfResult(workbook, r.request, t);
  if (JSON.stringify(whatIfResult) !== JSON.stringify(r))
    throw Error(
      "The what-if preview or host context changed. Preview and review again.",
    );
  return whatIfResult;
}
function buildWhatIfPreview(workbook, r, t = {}) {
  const { snapshot, requestSnapshot, context, ...s } =
    createWhatIfPreviewSnapshot(workbook, r, t);
  return s;
}
function applyWhatIfScenario(workbook, r, t, n = {}) {
  if (n.readOnly) throw Error("This workbook is read-only.");
  const whatIfPreviewSnapshot = createWhatIfPreviewSnapshot(workbook, r, n);
  if (whatIfPreviewSnapshot.request.kind !== "scenarios")
    throw Error(
      "Sensitivity tables are reports. Only named scenarios can be applied.",
    );
  if (
    t !== "baseline" &&
    (!Number.isInteger(t) || t < 0 || t >= whatIfPreviewSnapshot.trials.length)
  )
    throw Error("Choose one reviewed scenario.");
  const a =
    t === "baseline"
      ? whatIfPreviewSnapshot.baseline
      : whatIfPreviewSnapshot.trials[t];
  if (a.error || !a.outputs)
    throw Error(
      "A failed calculation cannot be applied. The baseline is unchanged.",
    );
  return a.inputs.every(
    (input, d) => input === whatIfPreviewSnapshot.baseline.inputs[d],
  )
    ? workbook
    : {
        ...workbook,
        sheets: workbook.sheets.map((sheet) => {
          const cells = { ...sheet.cells };
          return (
            whatIfPreviewSnapshot.request.inputs.forEach((input, l) => {
              input.sheetId === sheet.id &&
                (cells[input.address] = {
                  ...cells[input.address],
                  raw: a.inputs[l],
                });
            }),
            { ...sheet, cells: cells }
          );
        }),
      };
}
const formatAssumptionReference = (e, r) =>
  `${e.assumptions.find((assumption) => cellAddressKey(assumption) === cellAddressKey(r))?.sheetName ?? r.sheetId}!${r.address}`;
function buildWhatIfReportRows(whatIfPreview) {
  const r = whatIfPreview.request.inputs.map((input) =>
      formatAssumptionReference(whatIfPreview, input),
    ),
    t = whatIfPreview.request.outputs.map((output) =>
      formatAssumptionReference(whatIfPreview, output),
    ),
    n = [
      [whatIfPreview.request.name],
      ["Historical what-if report", whatIfPreview.request.kind],
      ["Source revision", whatIfPreview.sourceRevision],
      [
        "Refresh explicitly from What-if Analysis; existing reports retain their original assumptions.",
      ],
    ],
    o = (a) =>
      a.outputs ?? whatIfPreview.request.outputs.map(() => `Error: ${a.error}`);
  if (
    (n.push(
      ["Baseline", ...r, ...t],
      [
        whatIfPreview.baseline.name,
        ...whatIfPreview.baseline.inputs,
        ...o(whatIfPreview.baseline),
      ],
      [],
    ),
    whatIfPreview.request.kind === "scenarios")
  ) {
    n.push(["Scenario", ...r, ...t]);
    for (const trial of whatIfPreview.trials)
      n.push([trial.name, ...trial.inputs, ...o(trial)]);
  } else if (whatIfPreview.request.kind === "one-variable")
    if (whatIfPreview.request.orientation === "column") {
      n.push([r[0], ...t]);
      for (const trial of whatIfPreview.trials)
        n.push([trial.inputs[0], ...o(trial)]);
    } else
      (n.push([r[0], ...whatIfPreview.request.values]),
        t.forEach((a, s) =>
          n.push([a, ...whatIfPreview.trials.map((trial) => o(trial)[s])]),
        ));
  else {
    const { rowValues, columnValues } = whatIfPreview.request;
    (n.push([`Down: ${r[1]} / Across: ${r[0]}`, ...rowValues]),
      columnValues.forEach((columnValue, u) =>
        n.push([
          columnValue,
          ...rowValues.map(
            (rowValue, c) =>
              o(whatIfPreview.trials[u * rowValues.length + c])[0],
          ),
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
      ...whatIfPreview.assumptions.map((assumption) => [
        formatAssumptionReference(whatIfPreview, assumption),
        assumption.raw,
        assumption.literal ?? false,
      ]),
      [],
      ["Dependency cells", whatIfPreview.dependencyCells],
      ["Dependency-cell work units", whatIfPreview.workUnits],
    ),
    n
  );
}
function saveWhatIfAnalysis(workbook, r, t = {}, n) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error(
      "Workbook structure is protected. Unprotect it before saving a new report.",
    );
  const whatIfPreview = buildWhatIfPreview(workbook, r, t),
    a = workbook.whatIfAnalyses ?? [];
  if (n && !a.some((p) => p.id === n))
    throw Error("The saved analysis no longer exists.");
  if (
    n &&
    !isWhatIfAnalysisInScope(
      workbook,
      a.find((p) => p.id === n),
      t,
    )
  )
    throw Error("The saved analysis is outside the authorized context.");
  if (!n && a.length >= whatIfLimits.savedAnalyses)
    throw Error(
      "At most 40 saved analyses are supported. Remove an unused definition first.",
    );
  if (
    a.some(
      (p) =>
        p.id !== n &&
        p.request.name.toLowerCase() ===
          whatIfPreview.request.name.toLowerCase(),
    )
  )
    throw Error(
      "A saved analysis already has this name. Choose another name or update that analysis.",
    );
  const s =
      whatIfPreview.request.name
        .replace(/[\\/\[\]*?:]/g, " ")
        .slice(0, 25)
        .trim() || "What-if",
    d = new Set(workbook.sheets.map((sheet) => sheet.name.toLowerCase()));
  let u = s,
    l = 2;
  for (; d.has(u.toLowerCase());) {
    const p = ` ${l++}`;
    u = `${s.slice(0, 31 - p.length)}${p}`;
  }
  const worksheetRecord = createWorksheetRecord(u);
  for (const [p, m] of buildWhatIfReportRows(whatIfPreview).entries())
    for (const [v, raw] of m.entries())
      worksheetRecord.cells[`${columnIndexToLetters(v)}${p + 1}`] = {
        raw: raw,
        ...(typeof raw == "string" ? { literal: true } : {}),
        ...(p === 0 || p === 4 || p === 7
          ? { style: { bold: true, fill: "#e2f0d9" } }
          : {}),
      };
  ((worksheetRecord.columnWidths = { 0: 300, 1: 240, 2: 180, 3: 160 }),
    (worksheetRecord.frozenRows = 4));
  const h = {
    id: n ?? crypto.randomUUID(),
    reportSheetId: worksheetRecord.id,
    ...whatIfPreview,
  };
  return normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, worksheetRecord],
    whatIfAnalyses: n ? a.map((p) => (p.id === n ? h : p)) : [...a, h],
  });
}
function isWhatIfAnalysisInScope(workbook, r, t = {}) {
  const whatIfHostScope = resolveWhatIfHostScope(workbook, t);
  return r.assumptions.every((assumption) =>
    workbook.sheets.some((sheet) => sheet.id === assumption.sheetId)
      ? isVisibleAuthorizedCell(workbook, assumption, whatIfHostScope)
      : t.allowedSheetIds === undefined,
  );
}
function findSavedWhatIfAnalysis(workbook, r, t = {}) {
  if (t.readOnly) throw Error("This workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error("Workbook structure is protected.");
  const n = workbook.whatIfAnalyses?.find(
    (whatIfAnalys) => whatIfAnalys.id === r,
  );
  if (!n || !isWhatIfAnalysisInScope(workbook, n, t))
    throw Error(
      "The saved analysis is missing or outside the authorized context.",
    );
  return {
    ...workbook,
    whatIfAnalyses: workbook.whatIfAnalyses.filter(
      (whatIfAnalys) => whatIfAnalys.id !== r,
    ),
  };
}
function parseAssumptionNumbers(e, locale) {
  const t = locale === "de-DE" || locale === "fr-FR",
    n = e.split(t ? /[;\n]/ : /[,;\n]/);
  try {
    return n.map((o) => validateLocalizedNumberEntry(o, "Assumption", locale));
  } catch {
    throw Error(
      t
        ? "Enter finite numbers separated by semicolons or new lines; use a comma for decimals. Empty assumptions are not zero."
        : "Enter comma-separated finite numbers; empty assumptions are not zero.",
    );
  }
}

export {
  buildWhatIfPreview,
  computeWhatIfResult,
  ReportOutputsDialog,
  applyWhatIfScenario,
  saveWhatIfAnalysis,
  findSavedWhatIfAnalysis,
  buildWhatIfReportRows,
  isWhatIfAnalysisInScope,
  parseAssumptionNumbers,
};
