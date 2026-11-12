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

function ReportOutputsDialog(props) {
  const { text } = useWorkbookMessages(),
    dialogTitleId = ne.useId(),
    dialogRef = ne.useRef(null),
    cancellableOperation = useCancellableOperation(),
    latestPropsRef = ne.useRef(props);
  latestPropsRef.current = props;
  const [activeTab, setActiveTab] = ne.useState("Publish"),
    [requestDraft, setRequestDraft] = ne.useState({
      name: "Report output",
      kind: "range",
      sheetId: props.sheetId,
      range: props.range,
      mode: "snapshot",
      population: "full",
    }),
    [checked, setHiddenAuthorized] = ne.useState(false),
    [isFilterEnabled, setIsFilterEnabled] = ne.useState(false),
    [preview, setPreview] = ne.useState(null),
    [refreshTargetId, setRefreshTargetId] = ne.useState(),
    [previewWorkbook, setPreviewWorkbook] = ne.useState(null),
    [previewScopeKey, setPreviewScopeKey] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [selectedOutputId, setSelectedOutputId] = ne.useState(""),
    [inspection, setInspection] = ne.useState(null),
    [disabled, setDisabled] = ne.useState(false),
    reportScope = {
      readOnly: props.readOnly,
      allowedSheetIds: props.allowedSheetIds,
      includeHidden: !!props.allowHidden && checked,
    },
    currentScopeKey = JSON.stringify([reportScope, props.filter]),
    isPreviewCurrent =
      !!preview &&
      previewWorkbook === props.workbook &&
      previewScopeKey === currentScopeKey,
    sheet = props.workbook.sheets.find(
      (sheet) => sheet.id === requestDraft.sheetId,
    ),
    selectedOutput = props.workbook.reportOutputs?.find(
      (reportOutput) => reportOutput.id === selectedOutputId,
    ),
    selectedOutputStatus = selectedOutput
      ? describeReportOutputStatus(
          props.workbook,
          selectedOutput.id,
          reportScope,
        )
      : null,
    exportableSnapshot =
      activeTab === "Saved outputs"
        ? selectedOutputStatus?.status !== "unavailable"
          ? selectedOutput?.snapshot
          : undefined
        : isPreviewCurrent
          ? preview?.snapshot
          : undefined;
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const runAction = (action) => {
      try {
        (setErrorMessage(""), setStatusMessage(""), action());
      } catch (error) {
        setErrorMessage(error instanceof Error ? error.message : String(error));
      }
    },
    updateRequestDraft = (draftPatch) => {
      (setRequestDraft((previousDraft) => ({
        ...previousDraft,
        ...draftPatch,
      })),
        setRefreshTargetId(undefined),
        setPreview(null),
        setErrorMessage(""));
    },
    previewReportOutput = (request, refreshId) =>
      runAction(() => {
        var scrollContainer;
        const refreshedPreview = refreshReportOutput(
          props.workbook,
          request,
          reportScope,
          refreshId,
        );
        (setPreview(refreshedPreview),
          setRefreshTargetId(refreshId),
          setRequestDraft(refreshedPreview.request),
          setPreviewWorkbook(props.workbook),
          setPreviewScopeKey(currentScopeKey),
          setActiveTab("Publish"),
          ((scrollContainer =
            dialogRef.current?.querySelector(".lg-reports-scroll")) == null
            ? undefined
            : scrollContainer.scrollTo
          )?.call(scrollContainer, 0, 0));
      }),
    X$1 = async (format, snapshot) => {
      (setDisabled(true), setErrorMessage(""));
      try {
        const startProps = latestPropsRef.current,
          revalidateSourceScope = () =>
            validateReportSourceScope(
              latestPropsRef.current.workbook,
              snapshot,
              {
                readOnly: latestPropsRef.current.readOnly,
                allowedSheetIds: latestPropsRef.current.allowedSheetIds,
                includeHidden: !!latestPropsRef.current.allowHidden && checked,
              },
            );
        revalidateSourceScope();
        const exportFile =
          format === "png"
            ? await renderReportPngBlob(snapshot)
            : (() => {
                const renderedFile =
                  format === "svg"
                    ? renderReportSvg(snapshot)
                    : format === "html"
                      ? buildReportHtmlFile(snapshot)
                      : buildReportSvgFile(snapshot);
                return {
                  filename: renderedFile.filename,
                  blob: new Blob([renderedFile.content], {
                    type: renderedFile.mime,
                  }),
                };
              })();
        if (
          (revalidateSourceScope(),
          startProps.onDownload !== latestPropsRef.current.onDownload)
        )
          throw Error("The report destination changed. Export again.");
        (await (latestPropsRef.current.onDownload ?? downloadBlobAsFile)(
          exportFile,
        ),
          setStatusMessage(`${format.toUpperCase()} snapshot exported.`));
      } catch (error) {
        setErrorMessage(error.message);
      } finally {
        setDisabled(false);
      }
    },
    inspectObjectFile = async (file) => {
      (setDisabled(true), setInspection(null), setErrorMessage(""));
      try {
        const inspection = await importAutomationFile(
          file,
          cancellableOperation.begin(),
        );
        (inspectEmbeddedObjects(inspection), setInspection(inspection));
      } catch (error) {
        setErrorMessage(error.message);
      } finally {
        setDisabled(false);
      }
    };
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "lg-reports-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": dialogTitleId,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.key === "Escape" && (event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusables = [
            ...dialogRef.current.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),select:not(:disabled),summary,[tabindex="0"]',
            ),
          ];
          event.shiftKey && document.activeElement === focusables[0]
            ? (event.preventDefault(),
              (lastFocusable = focusables.at(-1)) == null ||
                lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === focusables.at(-1) &&
              (event.preventDefault(),
              (firstFocusable = focusables[0]) == null ||
                firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsxs("div", {
              children: [
                i.jsx("p", { children: "REPORTING" }),
                i.jsx("h2", {
                  id: dialogTitleId,
                  children: i.jsx(LocalizedText, {
                    children: "Report outputs & embedded objects",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close report outputs"),
              onClick: props.onClose,
              children: i.jsx(X, { size: 20 }),
            }),
          ],
        }),
        i.jsx("nav", {
          "aria-label": text("Report workflow"),
          children: ["Publish", "Saved outputs", "Object inventory"].map(
            (tabName) =>
              i.jsx(
                "button",
                {
                  "aria-pressed": activeTab === tabName,
                  onClick: () => {
                    (setActiveTab(tabName), setErrorMessage(""));
                  },
                  children: text(tabName),
                },
                tabName,
              ),
          ),
        }),
        i.jsxs("div", {
          className: "lg-reports-scroll",
          children: [
            activeTab === "Publish" &&
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
                            value: requestDraft.name,
                            onChange: (event) =>
                              updateRequestDraft({ name: event.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Source worksheet",
                          i.jsx("select", {
                            value: requestDraft.sheetId,
                            onChange: (event) => {
                              (updateRequestDraft({
                                sheetId: event.target.value,
                                kind: "range",
                                objectId: undefined,
                              }),
                                setIsFilterEnabled(false));
                            },
                            children: props.workbook.sheets
                              .filter(
                                (sheet) =>
                                  (!props.allowedSheetIds ||
                                    props.allowedSheetIds.includes(sheet.id)) &&
                                  (reportScope.includeHidden ||
                                    (!sheet.hidden && !sheet.veryHidden)),
                              )
                              .map((worksheet) =>
                                i.jsx(
                                  "option",
                                  {
                                    value: worksheet.id,
                                    children: worksheet.name,
                                  },
                                  worksheet.id,
                                ),
                              ),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Output content",
                          i.jsxs("select", {
                            value:
                              requestDraft.kind === "range"
                                ? "range"
                                : requestDraft.objectId,
                            onChange: (event) => {
                              const selectedObject = sheet?.objects?.find(
                                (object) => object.id === event.target.value,
                              );
                              updateRequestDraft(
                                selectedObject
                                  ? {
                                      kind: "object",
                                      objectId: selectedObject.id,
                                      range:
                                        selectedObject.kind === "chart"
                                          ? selectedObject.range
                                          : selectedObject.anchor,
                                    }
                                  : { kind: "range", objectId: undefined },
                              );
                            },
                            children: [
                              i.jsx("option", {
                                value: "range",
                                children: "Cell range / table",
                              }),
                              sheet?.objects?.map((object) =>
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
                            value: requestDraft.range,
                            disabled: requestDraft.kind === "object",
                            onChange: (event) =>
                              updateRequestDraft({ range: event.target.value }),
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          "Reuse mode",
                          i.jsxs("select", {
                            value: requestDraft.mode,
                            onChange: (event) =>
                              updateRequestDraft({ mode: event.target.value }),
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
                            value: requestDraft.population,
                            onChange: (event) =>
                              updateRequestDraft({
                                population: event.target.value,
                              }),
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
                      props.filter &&
                        requestDraft.sheetId === props.sheetId &&
                        i.jsxs("label", {
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: isFilterEnabled,
                              onChange: (event) => {
                                (setIsFilterEnabled(event.target.checked),
                                  updateRequestDraft({
                                    filter: event.target.checked
                                      ? props.filter
                                      : undefined,
                                    ...(event.target.checked
                                      ? { range: props.filter.range }
                                      : {}),
                                  }));
                              },
                            }),
                            "Use active population filter",
                          ],
                        }),
                      props.allowHidden &&
                        i.jsxs("label", {
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: checked,
                              onChange: (event) => {
                                (setHiddenAuthorized(event.target.checked),
                                  setPreview(null));
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
                      previewReportOutput(
                        {
                          ...requestDraft,
                          filter: refreshTargetId
                            ? requestDraft.filter
                            : isFilterEnabled
                              ? (props.filter ?? requestDraft.filter)
                              : undefined,
                        },
                        refreshTargetId,
                      ),
                    children: "Preview report output",
                  }),
                  preview &&
                    !isPreviewCurrent &&
                    i.jsx("p", {
                      role: "alert",
                      children:
                        "The workbook or permissions changed. Preview again.",
                    }),
                  isPreviewCurrent &&
                    i.jsxs("section", {
                      className: "lg-reports-preview",
                      children: [
                        i.jsxs("div", {
                          className: "lg-reports-preview-title",
                          children: [
                            i.jsx("strong", {
                              children: preview.refreshId
                                ? "Reviewed refresh"
                                : "Snapshot preview",
                            }),
                            i.jsxs("span", {
                              children: [
                                "Revision ",
                                preview.snapshot.revision,
                                " ·",
                                " ",
                                preview.snapshot.rows.length,
                                " rows ·",
                                " ",
                                preview.snapshot.request.mode,
                              ],
                            }),
                          ],
                        }),
                        i.jsx("div", {
                          className: "lg-reports-visual",
                          children: i.jsx(ReportOutputView, {
                            snapshot: preview.snapshot,
                          }),
                        }),
                      ],
                    }),
                ],
              }),
            activeTab === "Saved outputs" &&
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
                        value: selectedOutputId,
                        onChange: (event) =>
                          setSelectedOutputId(event.target.value),
                        children: [
                          i.jsx("option", {
                            value: "",
                            children: "Choose a saved output",
                          }),
                          props.workbook.reportOutputs?.map((reportOutput) =>
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
                  selectedOutput &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("p", {
                          role: "status",
                          className: `lg-reports-status ${selectedOutputStatus?.status}`,
                          children: [
                            selectedOutputStatus?.status.toUpperCase(),
                            " · ",
                            selectedOutputStatus?.message,
                          ],
                        }),
                        selectedOutputStatus?.status !== "unavailable" &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsxs("div", {
                                className: "lg-reports-actions",
                                children: [
                                  selectedOutput.snapshot.request.mode ===
                                    "linked" &&
                                    i.jsx("button", {
                                      disabled: disabled,
                                      onClick: () => {
                                        (setIsFilterEnabled(
                                          !!selectedOutput.snapshot.request
                                            .filter,
                                        ),
                                          previewReportOutput(
                                            selectedOutput.snapshot.request,
                                            selectedOutput.id,
                                          ));
                                      },
                                      children: "Preview linked refresh",
                                    }),
                                  i.jsx("button", {
                                    disabled: props.readOnly || disabled,
                                    onClick: () =>
                                      runAction(() => {
                                        (props.onWorkbookChange(
                                          removeReportOutput(
                                            props.workbook,
                                            selectedOutput.id,
                                            reportScope,
                                          ),
                                          "Remove report output",
                                        ),
                                          setSelectedOutputId(""));
                                      }),
                                    children: "Remove saved output",
                                  }),
                                ],
                              }),
                              i.jsx("div", {
                                className: "lg-reports-visual",
                                children: i.jsx(ReportOutputView, {
                                  snapshot: selectedOutput.snapshot,
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              }),
            activeTab === "Object inventory" &&
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
                          const selectedFile = event.target.files?.[0];
                          (selectedFile && inspectObjectFile(selectedFile),
                            (event.target.value = ""));
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
                          disabled: props.readOnly || disabled,
                          onClick: () =>
                            runAction(() => {
                              (props.onWorkbookChange(
                                retainEmbeddedObjectInventory(
                                  props.workbook,
                                  inspection,
                                  {
                                    readOnly: props.readOnly,
                                  },
                                ),
                                "Retain embedded-object inventory",
                              ),
                                setStatusMessage(
                                  "Inert object inventory retained. Keep the original file separately.",
                                ),
                                setInspection(null));
                            }),
                          children: "Retain inventory",
                        }),
                      ],
                    }),
                  props.workbook.embeddedObjectInventories?.map(
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
                                void (props.onDownload ?? downloadBlobAsFile)({
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
            errorMessage &&
              i.jsx("p", {
                role: "alert",
                className: "lg-reports-error",
                children: errorMessage,
              }),
            statusMessage &&
              i.jsx("p", { role: "status", children: statusMessage }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("div", {
              className: "lg-reports-actions",
              children:
                exportableSnapshot &&
                ["svg", "png", "html", "json"].map((exportFormat) =>
                  i.jsxs(
                    "button",
                    {
                      disabled: disabled,
                      onClick: () => void X$1(exportFormat, exportableSnapshot),
                      children: ["Export ", exportFormat.toUpperCase()],
                    },
                    exportFormat,
                  ),
                ),
            }),
            i.jsxs("div", {
              className: "lg-reports-actions",
              children: [
                props.readOnly && i.jsx("span", { children: "Read-only host" }),
                activeTab === "Publish" &&
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: !isPreviewCurrent || props.readOnly || disabled,
                    onClick: () =>
                      runAction(() => {
                        const isRefresh = !!preview?.refreshId,
                          publishedWorkbook = publishReportOutput(
                            props.workbook,
                            preview,
                            reportScope,
                          );
                        (props.onWorkbookChange(
                          publishedWorkbook,
                          isRefresh
                            ? "Refresh linked report output"
                            : "Publish report output",
                        ),
                          setSelectedOutputId(
                            preview?.refreshId ??
                              publishedWorkbook.reportOutputs.at(-1).id,
                          ),
                          setRefreshTargetId(undefined),
                          setPreview(null),
                          setStatusMessage(
                            isRefresh
                              ? "Linked output refreshed. Previously exported snapshots are unchanged."
                              : "Report output published.",
                          ),
                          setActiveTab("Saved outputs"));
                      }),
                    children: preview?.refreshId
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
          children: items.map((item, itemIndex) =>
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
              itemIndex,
            ),
          ),
        }),
      ],
    }),
  });
}

function resolveWhatIfHostScope(workbook, scopeRequest) {
  if (
    scopeRequest.includeHidden !== undefined &&
    typeof scopeRequest.includeHidden != "boolean"
  )
    throw Error("Include hidden must be boolean.");
  if (
    scopeRequest.allowedSheetIds !== undefined &&
    (!Array.isArray(scopeRequest.allowedSheetIds) ||
      scopeRequest.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("What-if host scope contains an unknown worksheet.");
  return {
    includeHidden: !!scopeRequest.includeHidden,
    allowedSheetIds: [
      ...new Set(
        scopeRequest.allowedSheetIds ??
          workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
  };
}
function isVisibleAuthorizedCell(workbook, cellRef, whatIfHostScope) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === cellRef.sheetId);
  if (!sheet || !whatIfHostScope.allowedSheetIds.includes(sheet.id))
    return false;
  const cellPosition = _t(cellRef.address);
  return (
    whatIfHostScope.includeHidden ||
    (!sheet.hidden &&
      !sheet.veryHidden &&
      !cn(sheet, "row").some(
        (rowSpan) =>
          rowSpan.start <= cellPosition.row && cellPosition.row <= rowSpan.end,
      ) &&
      !cn(sheet, "column").some(
        (columnSpan) =>
          columnSpan.start <= cellPosition.col &&
          cellPosition.col <= columnSpan.end,
      ))
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
function computeWhatIfResult(workbook, request, hostScopeOptions = {}) {
  const definition = validateWhatIfDefinition(request),
    whatIfHostScope = resolveWhatIfHostScope(workbook, hostScopeOptions),
    scenarios = expandWhatIfScenarios(definition);
  for (const sourceCell of [...definition.inputs, ...definition.outputs])
    if (!isVisibleAuthorizedCell(workbook, sourceCell, whatIfHostScope))
      throw Error(
        "A what-if source is missing or outside the visible, authorized host context.",
      );
  const baselineInputs = definition.inputs.map((input, inputIndex) => {
      const sheet = workbook.sheets.find((sheet) => sheet.id === input.sheetId),
        cell = sheet.cells[input.address],
        inputPosition = _t(input.address);
      if (!cell || !isFiniteNumber(cell.raw))
        throw Error(
          "Changing cells must contain finite numeric inputs, not formulas, text, blanks or spill outputs.",
        );
      if (sheet.protection?.enabled)
        throw Error(
          "A changing worksheet is protected. Unprotect it before analyzing scenarios.",
        );
      if (sheet.merges?.some((merge) => isCellInRange(merge, input.address)))
        throw Error("Merged cells cannot be what-if inputs.");
      if (
        [...(sheet.pivots ?? []), ...(sheet.queries ?? [])].some(
          (generatedResult) =>
            isCellInRange(generatedResult.outputRange, input.address),
        )
      )
        throw Error("Generated pivot/query results cannot be what-if inputs.");
      for (const table of sheet.tables ?? []) {
        if (!isCellInRange(table.range, input.address)) continue;
        const rangeAddress = parseRangeAddress(table.range),
          tableColumn =
            table.columns[inputPosition.col - rangeAddress.firstCol];
        if (
          inputPosition.row === rangeAddress.firstRow ||
          (table.showTotals && inputPosition.row === rangeAddress.lastRow) ||
          tableColumn?.calculatedFormula ||
          Object.entries(sheet.cells).some(([cellAddress, cell]) => {
            const candidatePosition = _t(cellAddress);
            return (
              candidatePosition.col === inputPosition.col &&
              candidatePosition.row > rangeAddress.firstRow &&
              candidatePosition.row <=
                rangeAddress.lastRow - (table.showTotals ? 1 : 0) &&
              !cell.literal &&
              typeof cell.raw == "string" &&
              cell.raw.startsWith("=")
            );
          })
        )
          throw Error(
            "Table headers, totals and calculated columns cannot be what-if inputs.",
          );
      }
      const inputValidations = (sheet.validations ?? []).filter((validation) =>
        isCellInRange(validation.range, input.address),
      );
      if (
        [
          cell.raw,
          ...scenarios.map((scenario) => scenario.inputs[inputIndex]),
        ].some((candidateValue) =>
          inputValidations.some(
            (validationRule) =>
              !validateDataValidationRule(validationRule, candidateValue),
          ),
        )
      )
        throw Error(
          `An assumption for ${sheet.name}!${input.address} violates its data validation constraints. Baseline unchanged.`,
        );
      return cell.raw;
    }),
    dependencyMap = new Map();
  for (const output of definition.outputs) {
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
    const dependencyInspection = inspectFormulaDependencies(
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
      dependencyNodes = [
        dependencyInspection.target,
        ...dependencyInspection.nodes,
      ],
      dependencyKeys = new Set(dependencyNodes.map((event) => event.key));
    if (
      dependencyInspection.coverage.truncated ||
      dependencyInspection.coverage.depthLimited
    )
      throw Error(
        "What-if analysis requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
      );
    const issue = dependencyInspection.issues.find((issue) =>
      dependencyKeys.has(issue.from),
    );
    if (issue || dependencyInspection.cycles.length)
      throw Error(
        `What-if analysis requires supported static dependencies: ${issue?.message ?? "a cycle was found."}`,
      );
    for (const dependencyNode of dependencyNodes)
      dependencyMap.set(dependencyNode.key, {
        sheetId: dependencyNode.sheetId,
        address: dependencyNode.address,
      });
  }
  if (dependencyMap.size > whatIfLimits.dependencyCells)
    throw Error(
      "What-if analysis supports at most 500 dependency cells across all results.",
    );
  if (
    definition.inputs.some((input) => !dependencyMap.has(cellAddressKey(input)))
  )
    throw Error(
      "Each changing input must be a precedent of at least one selected formula result.",
    );
  const assumptions = [...dependencyMap.values()].map((dependency) => {
      if (!isVisibleAuthorizedCell(workbook, dependency, whatIfHostScope))
        throw Error(
          "A formula dependency is outside the visible, authorized host context.",
        );
      const sheet = workbook.sheets.find(
          (sheet) => sheet.id === dependency.sheetId,
        ),
        cell = sheet.cells[dependency.address];
      if (!cell || cell.raw === null)
        throw Error(
          "What-if analysis requires populated static dependencies; blank or possible spill-child cells are unsupported.",
        );
      if (
        !cell.literal &&
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=")
      ) {
        const formulaText = resolveDefinedNameValue(
          cell.raw,
          workbook,
          sheet.id,
        ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
        if (
          [...formulaText.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)]
            .map((match) =>
              match[1]
                .toUpperCase()
                .replace(/^_XLFN\./, "")
                .replace(/^_XLWS\./, ""),
            )
            .some((functionName) => whatIfAllowedFunctions.has(functionName)) ||
          /[{}#]/.test(formulaText)
        )
          throw Error(
            "What-if analysis excludes volatile, dynamic-reference, array/spill and unsupported reusable-function dependencies.",
          );
      }
      return {
        ...dependency,
        sheetName: sheet.name,
        raw: cell.raw,
        ...(cell.literal === undefined ? {} : { literal: cell.literal }),
      };
    }),
    workUnits = dependencyMap.size * (scenarios.length + 1);
  if (workUnits > whatIfLimits.workUnits)
    throw Error(
      "What-if evaluation would exceed 50,000 dependency-cell work units. Reduce the values, scenarios or dependencies.",
    );
  const scopedWorkbook = {
    ...workbook,
    sheets: workbook.sheets
      .filter((sheet) => whatIfHostScope.allowedSheetIds.includes(sheet.id))
      .map((sheet) => ({
        ...sheet,
        cells: Object.fromEntries(
          Object.entries(sheet.cells).filter(([address]) =>
            dependencyMap.has(
              cellAddressKey({ sheetId: sheet.id, address: address }),
            ),
          ),
        ),
      })),
  };
  function runTrial(trialName, trialInputs) {
    try {
      const inputOverrides = new Map(
          definition.inputs.map((input, inputIndex) => [
            cellAddressKey(input),
            trialInputs[inputIndex],
          ]),
        ),
        workbook = {
          ...scopedWorkbook,
          sheets: scopedWorkbook.sheets.map((sheet) => ({
            ...sheet,
            cells: Object.fromEntries(
              Object.entries(sheet.cells).map(([address, cell]) => {
                const raw = inputOverrides.get(
                  cellAddressKey({ sheetId: sheet.id, address: address }),
                );
                return [
                  address,
                  raw === undefined ? cell : { ...cell, raw: raw },
                ];
              }),
            ),
          })),
        },
        workbookWithSettings = calculateWorkbookWithSettings(workbook),
        outputs = definition.outputs.map((output) =>
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
      for (const dependencyCell of dependencyMap.values())
        if (
          workbookWithSettings.getSpillRange(
            dependencyCell.sheetId,
            dependencyCell.address,
          )
        )
          throw Error(
            "A trial produced an array/spill; scalar results are required.",
          );
      if (workbookWithSettings.getOperationalErrors().length)
        throw Error(
          "Dependency inspection encountered an operational calculation error.",
        );
      return {
        name: trialName,
        inputs: [...trialInputs],
        outputs: outputs,
        error: null,
      };
    } catch (error) {
      return {
        name: trialName,
        inputs: [...trialInputs],
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
    requestSnapshot: JSON.stringify(definition),
    request: definition,
    context: whatIfHostScope,
    baseline: runTrial("Baseline", baselineInputs),
    trials: scenarios.map((scenario) =>
      runTrial(scenario.name, scenario.inputs),
    ),
    assumptions: assumptions,
    sourceRevision: workbook.revision,
    dependencyCells: dependencyMap.size,
    workUnits: workUnits,
  };
}
function createWhatIfPreviewSnapshot(workbook, preview, hostScopeOptions) {
  if (JSON.stringify(workbook) !== preview.snapshot)
    throw Error(
      "The workbook changed after preview. Preview and review the analysis again.",
    );
  const whatIfResult = computeWhatIfResult(
    workbook,
    preview.request,
    hostScopeOptions,
  );
  if (JSON.stringify(whatIfResult) !== JSON.stringify(preview))
    throw Error(
      "The what-if preview or host context changed. Preview and review again.",
    );
  return whatIfResult;
}
function buildWhatIfPreview(workbook, preview, hostScopeOptions = {}) {
  const { snapshot, requestSnapshot, context, ...previewDetails } =
    createWhatIfPreviewSnapshot(workbook, preview, hostScopeOptions);
  return previewDetails;
}
function applyWhatIfScenario(workbook, preview, scenarioChoice, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  const whatIfPreviewSnapshot = createWhatIfPreviewSnapshot(
    workbook,
    preview,
    options,
  );
  if (whatIfPreviewSnapshot.request.kind !== "scenarios")
    throw Error(
      "Sensitivity tables are reports. Only named scenarios can be applied.",
    );
  if (
    scenarioChoice !== "baseline" &&
    (!Number.isInteger(scenarioChoice) ||
      scenarioChoice < 0 ||
      scenarioChoice >= whatIfPreviewSnapshot.trials.length)
  )
    throw Error("Choose one reviewed scenario.");
  const selectedScenario =
    scenarioChoice === "baseline"
      ? whatIfPreviewSnapshot.baseline
      : whatIfPreviewSnapshot.trials[scenarioChoice];
  if (selectedScenario.error || !selectedScenario.outputs)
    throw Error(
      "A failed calculation cannot be applied. The baseline is unchanged.",
    );
  return selectedScenario.inputs.every(
    (input, inputIndex) =>
      input === whatIfPreviewSnapshot.baseline.inputs[inputIndex],
  )
    ? workbook
    : {
        ...workbook,
        sheets: workbook.sheets.map((sheet) => {
          const cells = { ...sheet.cells };
          return (
            whatIfPreviewSnapshot.request.inputs.forEach(
              (input, requestInputIndex) => {
                input.sheetId === sheet.id &&
                  (cells[input.address] = {
                    ...cells[input.address],
                    raw: selectedScenario.inputs[requestInputIndex],
                  });
              },
            ),
            { ...sheet, cells: cells }
          );
        }),
      };
}
const formatAssumptionReference = (whatIfPreview, cellReference) =>
  `${whatIfPreview.assumptions.find((assumption) => cellAddressKey(assumption) === cellAddressKey(cellReference))?.sheetName ?? cellReference.sheetId}!${cellReference.address}`;
function buildWhatIfReportRows(whatIfPreview) {
  const inputLabels = whatIfPreview.request.inputs.map((input) =>
      formatAssumptionReference(whatIfPreview, input),
    ),
    outputLabels = whatIfPreview.request.outputs.map((output) =>
      formatAssumptionReference(whatIfPreview, output),
    ),
    reportRows = [
      [whatIfPreview.request.name],
      ["Historical what-if report", whatIfPreview.request.kind],
      ["Source revision", whatIfPreview.sourceRevision],
      [
        "Refresh explicitly from What-if Analysis; existing reports retain their original assumptions.",
      ],
    ],
    formatOutputs = (scenarioResult) =>
      scenarioResult.outputs ??
      whatIfPreview.request.outputs.map(() => `Error: ${scenarioResult.error}`);
  if (
    (reportRows.push(
      ["Baseline", ...inputLabels, ...outputLabels],
      [
        whatIfPreview.baseline.name,
        ...whatIfPreview.baseline.inputs,
        ...formatOutputs(whatIfPreview.baseline),
      ],
      [],
    ),
    whatIfPreview.request.kind === "scenarios")
  ) {
    reportRows.push(["Scenario", ...inputLabels, ...outputLabels]);
    for (const trial of whatIfPreview.trials)
      reportRows.push([trial.name, ...trial.inputs, ...formatOutputs(trial)]);
  } else if (whatIfPreview.request.kind === "one-variable")
    if (whatIfPreview.request.orientation === "column") {
      reportRows.push([inputLabels[0], ...outputLabels]);
      for (const trial of whatIfPreview.trials)
        reportRows.push([trial.inputs[0], ...formatOutputs(trial)]);
    } else
      (reportRows.push([inputLabels[0], ...whatIfPreview.request.values]),
        outputLabels.forEach((outputLabel, outputIndex) =>
          reportRows.push([
            outputLabel,
            ...whatIfPreview.trials.map(
              (trial) => formatOutputs(trial)[outputIndex],
            ),
          ]),
        ));
  else {
    const { rowValues, columnValues } = whatIfPreview.request;
    (reportRows.push([
      `Down: ${inputLabels[1]} / Across: ${inputLabels[0]}`,
      ...rowValues,
    ]),
      columnValues.forEach((columnValue, columnValueIndex) =>
        reportRows.push([
          columnValue,
          ...rowValues.map(
            (rowValue, rowValueIndex) =>
              formatOutputs(
                whatIfPreview.trials[
                  columnValueIndex * rowValues.length + rowValueIndex
                ],
              )[0],
          ),
        ]),
      ));
  }
  return (
    reportRows.push(
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
    reportRows
  );
}
function saveWhatIfAnalysis(workbook, preview, options = {}, savedAnalysisId) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error(
      "Workbook structure is protected. Unprotect it before saving a new report.",
    );
  const whatIfPreview = buildWhatIfPreview(workbook, preview, options),
    savedAnalyses = workbook.whatIfAnalyses ?? [];
  if (
    savedAnalysisId &&
    !savedAnalyses.some((analysis) => analysis.id === savedAnalysisId)
  )
    throw Error("The saved analysis no longer exists.");
  if (
    savedAnalysisId &&
    !isWhatIfAnalysisInScope(
      workbook,
      savedAnalyses.find(
        (savedAnalysis) => savedAnalysis.id === savedAnalysisId,
      ),
      options,
    )
  )
    throw Error("The saved analysis is outside the authorized context.");
  if (!savedAnalysisId && savedAnalyses.length >= whatIfLimits.savedAnalyses)
    throw Error(
      "At most 40 saved analyses are supported. Remove an unused definition first.",
    );
  if (
    savedAnalyses.some(
      (otherAnalysis) =>
        otherAnalysis.id !== savedAnalysisId &&
        otherAnalysis.request.name.toLowerCase() ===
          whatIfPreview.request.name.toLowerCase(),
    )
  )
    throw Error(
      "A saved analysis already has this name. Choose another name or update that analysis.",
    );
  const baseSheetName =
      whatIfPreview.request.name
        .replace(/[\\/\[\]*?:]/g, " ")
        .slice(0, 25)
        .trim() || "What-if",
    existingSheetNames = new Set(
      workbook.sheets.map((sheet) => sheet.name.toLowerCase()),
    );
  let sheetName = baseSheetName,
    suffixNumber = 2;
  for (; existingSheetNames.has(sheetName.toLowerCase());) {
    const suffix = ` ${suffixNumber++}`;
    sheetName = `${baseSheetName.slice(0, 31 - suffix.length)}${suffix}`;
  }
  const worksheetRecord = createWorksheetRecord(sheetName);
  for (const [rowIndex, rowCells] of buildWhatIfReportRows(
    whatIfPreview,
  ).entries())
    for (const [columnIndex, raw] of rowCells.entries())
      worksheetRecord.cells[
        `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`
      ] = {
        raw: raw,
        ...(typeof raw == "string" ? { literal: true } : {}),
        ...(rowIndex === 0 || rowIndex === 4 || rowIndex === 7
          ? { style: { bold: true, fill: "#e2f0d9" } }
          : {}),
      };
  ((worksheetRecord.columnWidths = { 0: 300, 1: 240, 2: 180, 3: 160 }),
    (worksheetRecord.frozenRows = 4));
  const analysisRecord = {
    id: savedAnalysisId ?? crypto.randomUUID(),
    reportSheetId: worksheetRecord.id,
    ...whatIfPreview,
  };
  return normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, worksheetRecord],
    whatIfAnalyses: savedAnalysisId
      ? savedAnalyses.map((existingAnalysis) =>
          existingAnalysis.id === savedAnalysisId
            ? analysisRecord
            : existingAnalysis,
        )
      : [...savedAnalyses, analysisRecord],
  });
}
function isWhatIfAnalysisInScope(workbook, analysis, options = {}) {
  const whatIfHostScope = resolveWhatIfHostScope(workbook, options);
  return analysis.assumptions.every((assumption) =>
    workbook.sheets.some((sheet) => sheet.id === assumption.sheetId)
      ? isVisibleAuthorizedCell(workbook, assumption, whatIfHostScope)
      : options.allowedSheetIds === undefined,
  );
}
function findSavedWhatIfAnalysis(workbook, analysisId, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error("Workbook structure is protected.");
  const savedAnalysis = workbook.whatIfAnalyses?.find(
    (whatIfAnalys) => whatIfAnalys.id === analysisId,
  );
  if (
    !savedAnalysis ||
    !isWhatIfAnalysisInScope(workbook, savedAnalysis, options)
  )
    throw Error(
      "The saved analysis is missing or outside the authorized context.",
    );
  return {
    ...workbook,
    whatIfAnalyses: workbook.whatIfAnalyses.filter(
      (whatIfAnalys) => whatIfAnalys.id !== analysisId,
    ),
  };
}
function parseAssumptionNumbers(assumptionText, locale) {
  const usesDecimalComma = locale === "de-DE" || locale === "fr-FR",
    entries = assumptionText.split(usesDecimalComma ? /[;\n]/ : /[,;\n]/);
  try {
    return entries.map((entry) =>
      validateLocalizedNumberEntry(entry, "Assumption", locale),
    );
  } catch {
    throw Error(
      usesDecimalComma
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
