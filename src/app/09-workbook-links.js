import { unzipSync } from "fflate";
import ne from "react";
import i from "react/jsx-runtime";
import { X } from "lucide-react";
import { string, number, array, literal, unknown, object } from "zod";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  validateRetainedSourceFile,
  assertWorkbookStructureEditable,
  sourceImportLimits,
  requireRetainedXlsxSource,
  validateExtractedSourceTables,
  normalizeWorkbookInput,
  workbookLinkLimits,
  columnIndexToLetters,
  resolveExtractedCellValue,
  buildRetainedSourceRecord,
} from "./03-workbook-model.js";
import {
  assertSourceInspectionCurrent,
  applyWorkbookLinkPreview,
  recordSourceInspection,
  previewWorkbookLinkOperation,
  buildWorkbookLinkReview,
} from "./05-add-ins-and-custom-functions.js";
import {
  importWorkbookFile,
  runFileImport,
} from "./07-file-import-and-worker-dispatch.js";
import {
  F$,
  buildExchangeSchemaReport,
  importXlsxWorkbook,
  createEmptyExportLedger,
} from "./08-xlsx-import-export-and-preservation.js";

function buildSourceImportRecord(activeWorkbook) {
  const retainedSource = requireRetainedXlsxSource(activeWorkbook),
    xlsxWorkbookPackage = importXlsxWorkbook(
      retainedSource.bytes,
      createEmptyExportLedger(),
    ),
    exchangeSchemaReport = buildExchangeSchemaReport(
      "import",
      retainedSource.fileName,
      retainedSource.bytes.length,
      activeWorkbook,
      xlsxWorkbookPackage,
    );
  return (
    (exchangeSchemaReport.partInventory = F$(
      unzipSync(retainedSource.bytes),
      retainedSource.sha256,
    )),
    exchangeSchemaReport
  );
}
const operationScopeContext = ne.createContext({}),
  useOperationScope = () => ne.useContext(operationScopeContext);
function useCancellableOperation() {
  const operationScope = useOperationScope(),
    controllerRef = ne.useRef(null),
    cancel = ne.useCallback(() => controllerRef.current?.abort(), []);
  return (
    ne.useEffect(() => () => cancel(), [cancel]),
    {
      cancel: cancel,
      begin: () => {
        cancel();
        const controller = new AbortController();
        return (
          (controllerRef.current = controller),
          {
            ...operationScope,
            signal: operationScope.signal
              ? AbortSignal.any([controller.signal, operationScope.signal])
              : controller.signal,
          }
        );
      },
    }
  );
}
function WorkbookLinkStatusBadge({ workbook, sources, onReview }) {
  const statusMessage = ne.useMemo(() => {
    try {
      const workbookLinkReview = buildWorkbookLinkReview(workbook, { sources });
      return !workbookLinkReview.links.length &&
        !workbookLinkReview.unsupported.length
        ? null
        : `${workbookLinkReview.links.length} workbook link targets · ${workbookLinkReview.links.filter((link) => link.status !== "matches-approved-snapshot").length} need source review · ${workbookLinkReview.unsupported.length} unsupported. Displayed linked values are cached snapshots.`;
    } catch {
      return "Workbook link inventory needs review; links exceed the supported review limit.";
    }
  }, [workbook, sources]);
  return statusMessage
    ? i.jsxs("div", {
        className: "lg-workbook-link-status",
        role: "status",
        children: [
          i.jsx("span", { children: statusMessage }),
          i.jsx("button", {
            onClick: onReview,
            children: "Review workbook links",
          }),
        ],
      })
    : null;
}
function WorkbookLinksDialog(props) {
  const latestProps = ne.useRef(props);
  latestProps.current = props;
  const dialogRef = ne.useRef(null),
    cancellableOperation = useCancellableOperation(),
    requestCounter = ne.useRef(0),
    isMountedRef = ne.useRef(true),
    [localSources, setLocalSources] = ne.useState([]),
    [chosenTarget, setChosenTarget] = ne.useState(""),
    [replacementTarget, setReplacementTarget] = ne.useState(""),
    [action, setAction] = ne.useState("refresh"),
    [isLiteralActivated, setIsLiteralActivated] = ne.useState(false),
    [isCacheAcknowledged, setIsCacheAcknowledged] = ne.useState(false),
    [workbookLinkPreview, setWorkbookLinkPreview] = ne.useState(null),
    [operationError, setOperationError] = ne.useState(""),
    [operationNotice, setOperationNotice] = ne.useState(""),
    [cellPageIndex, setCellPageIndex] = ne.useState(0),
    [previewPageIndex, setPreviewPageIndex] = ne.useState(0),
    [disabled, setDisabled] = ne.useState(false),
    [workbook, setWorkbook] = ne.useState(null),
    [target, setTarget] = ne.useState(""),
    [allowedSheetIds, setAllowedSheetIds] = ne.useState([]),
    [importWarnings, setImportWarnings] = ne.useState([]),
    [checked, setChecked] = ne.useState(false),
    sources = props.approvedSources ?? localSources,
    inventoryResult = ne.useMemo(() => {
      try {
        return {
          inventory: buildWorkbookLinkReview(props.workbook, {
            sources: sources,
          }),
          error: "",
        };
      } catch (error) {
        return { inventory: null, error: error.message };
      }
    }, [props.workbook, sources]),
    X$1 = inventoryResult.inventory,
    selectedLink = X$1?.links.find(
      (link) => link.target === (chosenTarget || X$1.links[0]?.target),
    ),
    selectedTarget = selectedLink?.target ?? "",
    linkedCells =
      X$1?.cells.filter((cell) =>
        cell.references.some(
          (reference) => reference.target === selectedTarget,
        ),
      ) ?? [],
    updateApprovedSources = (nextSources) => {
      if (props.approvedSources !== undefined && !props.onApprovedSourcesChange)
        throw Error("Source approvals are controlled by your host.");
      (props.onApprovedSourcesChange?.(nextSources),
        props.approvedSources === undefined && setLocalSources(nextSources),
        setWorkbookLinkPreview(null));
    },
    getOperationOptions = () => ({
      sources: latestProps.current.approvedSources ?? localSources,
      readOnly: latestProps.current.readOnly,
      allowedSheetIds: latestProps.current.allowedSheetIds,
    });
  (ne.useEffect(() => {
    isMountedRef.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.focus(),
      () => {
        ((isMountedRef.current = false),
          requestCounter.current++,
          cancellableOperation.cancel(),
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      setWorkbookLinkPreview(null);
    }, [props.workbook, props.approvedSources]));
  const runOperation = (operation) => {
      (setOperationError(""), setOperationNotice(""));
      try {
        operation();
      } catch (error) {
        setOperationError(
          error instanceof Error
            ? error.message
            : "Workbook link operation failed.",
        );
      }
    },
    onClick = () =>
      runOperation(() => {
        (setWorkbookLinkPreview(
          previewWorkbookLinkOperation(
            props.workbook,
            {
              action: action,
              target: selectedTarget,
              replacementTarget: replacementTarget,
              activateLiteral: isLiteralActivated,
              acknowledgeCached: isCacheAcknowledged,
            },
            getOperationOptions(),
          ),
        ),
          setCellPageIndex(0),
          setPreviewPageIndex(0));
      }),
    applyLinkPreview = () =>
      runOperation(() => {
        if (!workbookLinkPreview) return;
        const nextWorkbook = applyWorkbookLinkPreview(
          latestProps.current.workbook,
          workbookLinkPreview,
          getOperationOptions(),
        );
        (latestProps.current.onWorkbookChange(
          nextWorkbook,
          `${workbookLinkPreview.request.action} workbook links`,
        ),
          setWorkbookLinkPreview(null),
          setOperationNotice(
            "Reviewed workbook link changes applied. Undo restores the prior workbook.",
          ));
      }),
    attachSourceFile = async (sourceFile) => {
      const requestId = ++requestCounter.current;
      (setDisabled(true),
        setOperationError(""),
        setWorkbook(null),
        setWorkbookLinkPreview(null));
      try {
        const importResult = await importWorkbookFile(
          sourceFile,
          cancellableOperation.begin(),
        );
        if (!isMountedRef.current || requestId !== requestCounter.current)
          return;
        (setWorkbook(importResult.workbook),
          setTarget(selectedTarget || `[${sourceFile.name}]`),
          setAllowedSheetIds(
            importResult.workbook.sheets
              .filter((sheet) => !sheet.hidden && !sheet.veryHidden)
              .map((visibleSheet) => visibleSheet.id),
          ),
          setImportWarnings(importResult.warnings),
          setChecked(false));
      } catch (error) {
        isMountedRef.current &&
          requestId === requestCounter.current &&
          setOperationError(error.message);
      } finally {
        isMountedRef.current &&
          requestId === requestCounter.current &&
          setDisabled(false);
      }
    },
    canAttachFile =
      props.allowFileSources !== false &&
      (props.approvedSources === undefined || !!props.onApprovedSourcesChange),
    resetReviewState = () => {
      (setWorkbookLinkPreview(null),
        setOperationError(""),
        setOperationNotice(""),
        setCellPageIndex(0));
    };
  return i.jsx("div", {
    className: "lg-reports-overlay",
    children: i.jsxs("section", {
      className: "lg-reports-dialog lg-workbook-links",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": "Workbook links",
      tabIndex: -1,
      ref: dialogRef,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" && (event.preventDefault(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ).filter((element) => !element.hidden),
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1);
          event.shiftKey &&
          (document.activeElement === firstFocusable ||
            document.activeElement === dialogRef.current)
            ? (event.preventDefault(),
              lastFocusable == null || lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === lastFocusable &&
              (event.preventDefault(),
              firstFocusable == null || firstFocusable.focus());
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
              onClick: props.onClose,
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
            (operationError || inventoryResult.error) &&
              i.jsx("p", {
                role: "alert",
                className: "lg-links-error",
                children: operationError || inventoryResult.error,
              }),
            operationNotice &&
              i.jsx("p", { role: "status", children: operationNotice }),
            i.jsxs("div", {
              className: "lg-links-summary",
              children: [
                i.jsxs("strong", {
                  children: [X$1?.links.length ?? 0, " linked sources"],
                }),
                i.jsx("span", { children: "Manual refresh only" }),
                i.jsx("span", {
                  children:
                    "Native JSON retains cached values and revision evidence",
                }),
              ],
            }),
            X$1?.links.length
              ? i.jsxs(i.Fragment, {
                  children: [
                    i.jsxs("label", {
                      children: [
                        "Linked workbook",
                        i.jsx("select", {
                          "aria-label": "Linked workbook",
                          value: selectedTarget,
                          onChange: (event) => {
                            (setChosenTarget(event.target.value),
                              setIsLiteralActivated(false),
                              setIsCacheAcknowledged(false),
                              resetReviewState());
                          },
                          children: X$1.links.map((link) =>
                            i.jsxs(
                              "option",
                              {
                                value: link.target,
                                children: [link.target, " · ", link.status],
                              },
                              link.target,
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
                            i.jsx("dd", { children: selectedTarget }),
                            selectedLink?.packageTarget &&
                              i.jsxs(i.Fragment, {
                                children: [
                                  i.jsx("dt", {
                                    children: "Original XLSX target",
                                  }),
                                  i.jsx("dd", {
                                    children: selectedLink.packageTarget,
                                  }),
                                ],
                              }),
                            i.jsx("dt", { children: "Status" }),
                            i.jsx("dd", {
                              children: selectedLink?.status,
                            }),
                            i.jsx("dt", { children: "Last resolved revision" }),
                            i.jsx("dd", {
                              children:
                                selectedLink?.lastResolvedRevision ??
                                "Never resolved",
                            }),
                            i.jsx("dt", { children: "Resolved at" }),
                            i.jsx("dd", {
                              children:
                                selectedLink?.lastResolvedAt ?? "Not available",
                            }),
                            i.jsx("dt", { children: "Source snapshot" }),
                            i.jsx("dd", {
                              children:
                                selectedLink?.sourceName ??
                                "No approved result",
                            }),
                          ],
                        }),
                        i.jsx("p", { children: X$1.cachePolicy }),
                        selectedLink?.cached &&
                          selectedLink.status !== "matches-approved-snapshot" &&
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
                                children: linkedCells
                                  .slice(
                                    cellPageIndex * 20,
                                    cellPageIndex * 20 + 20,
                                  )
                                  .map((linkedCell) =>
                                    i.jsxs(
                                      "tr",
                                      {
                                        children: [
                                          i.jsxs("td", {
                                            children: [
                                              linkedCell.sheetName,
                                              "!",
                                              linkedCell.address,
                                            ],
                                          }),
                                          i.jsx("td", {
                                            children: linkedCell.references
                                              .filter(
                                                (reference) =>
                                                  reference.target ===
                                                  selectedTarget,
                                              )
                                              .map(
                                                (matchedReference) =>
                                                  `${matchedReference.sheet}!${matchedReference.range}`,
                                              )
                                              .join(", "),
                                          }),
                                          i.jsx("td", {
                                            children: linkedCell.formula,
                                          }),
                                          i.jsx("td", {
                                            children: linkedCell.literal
                                              ? "Inert text — activation required"
                                              : "Formula",
                                          }),
                                        ],
                                      },
                                      linkedCell.sheetId + linkedCell.address,
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
                          disabled: cellPageIndex === 0,
                          onClick: () => setCellPageIndex(cellPageIndex - 1),
                          children: "Previous linked cells",
                        }),
                        i.jsxs("span", {
                          children: [
                            "Page ",
                            cellPageIndex + 1,
                            " of ",
                            Math.max(1, Math.ceil(linkedCells.length / 20)),
                          ],
                        }),
                        i.jsx("button", {
                          disabled:
                            (cellPageIndex + 1) * 20 >= linkedCells.length,
                          onClick: () => setCellPageIndex(cellPageIndex + 1),
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
                              value: action,
                              onChange: (event) => {
                                (setAction(event.target.value),
                                  resetReviewState());
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
                        action === "replace" &&
                          i.jsxs("label", {
                            children: [
                              "Replacement source",
                              i.jsxs("select", {
                                "aria-label": "Replacement source",
                                value: replacementTarget,
                                onChange: (event) => {
                                  (setReplacementTarget(event.target.value),
                                    resetReviewState());
                                },
                                children: [
                                  i.jsx("option", {
                                    value: "",
                                    children: "Choose an approved source",
                                  }),
                                  sources
                                    .filter(
                                      (event) =>
                                        event.target !== selectedTarget,
                                    )
                                    .map((approvedSource) =>
                                      i.jsxs(
                                        "option",
                                        {
                                          value: approvedSource.target,
                                          children: [
                                            approvedSource.target,
                                            " · ",
                                            approvedSource.workbook.name,
                                          ],
                                        },
                                        approvedSource.target,
                                      ),
                                    ),
                                ],
                              }),
                            ],
                          }),
                        linkedCells.some(
                          (literalCell) => literalCell.literal,
                        ) &&
                          i.jsxs("label", {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: isLiteralActivated,
                                onChange: (event) => {
                                  (setIsLiteralActivated(event.target.checked),
                                    resetReviewState());
                                },
                              }),
                              "Activate these inert formulas using only reviewed cached source values",
                            ],
                          }),
                        action === "break" &&
                          i.jsxs("label", {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: isCacheAcknowledged,
                                onChange: (event) => {
                                  (setIsCacheAcknowledged(event.target.checked),
                                    resetReviewState());
                                },
                              }),
                              "I understand that complete formulas will be replaced with the previewed cached values",
                            ],
                          }),
                        i.jsx("button", {
                          className: "lg-reports-primary",
                          disabled: props.readOnly || disabled,
                          onClick: onClick,
                          children:
                            action === "refresh"
                              ? "Preview refresh"
                              : action === "replace"
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
            !!X$1?.unsupported.length &&
              i.jsxs("section", {
                className: "lg-links-warning",
                children: [
                  i.jsx("h3", { children: "Unsupported external syntax" }),
                  X$1.unsupported.slice(0, 20).map((unsupportedEntry) =>
                    i.jsxs(
                      "p",
                      {
                        children: [
                          unsupportedEntry.sheetName,
                          "!",
                          unsupportedEntry.address,
                          ": ",
                          unsupportedEntry.unsupported,
                        ],
                      },
                      unsupportedEntry.sheetId + unsupportedEntry.address,
                    ),
                  ),
                ],
              }),
            !!props.workbook.workbookLinks?.packageTargets?.length &&
              i.jsxs("details", {
                children: [
                  i.jsx("summary", {
                    children: "Retained XLSX target inventory",
                  }),
                  props.workbook.workbookLinks.packageTargets.map(
                    (packageTarget) =>
                      i.jsxs(
                        "p",
                        {
                          children: [
                            packageTarget.key,
                            " → ",
                            packageTarget.target,
                            " (",
                            packageTarget.part,
                            "). Retained original metadata; never fetched.",
                          ],
                        },
                        packageTarget.key,
                      ),
                  ),
                ],
              }),
            i.jsxs("section", {
              className: "lg-links-card",
              children: [
                i.jsx("h3", { children: "Approved source snapshots" }),
                sources.length
                  ? i.jsx("ul", {
                      children: sources.map((source) =>
                        i.jsxs(
                          "li",
                          {
                            children: [
                              source.target,
                              " · ",
                              source.workbook.name,
                              " · revision",
                              " ",
                              source.workbook.revision,
                              " · ",
                              source.allowedSheetIds.length,
                              " approved worksheets",
                              " ",
                              (!props.approvedSources ||
                                props.onApprovedSourcesChange) &&
                                i.jsxs("button", {
                                  onClick: () =>
                                    runOperation(() =>
                                      updateApprovedSources(
                                        sources.filter(
                                          (candidateSource) =>
                                            candidateSource !== source,
                                        ),
                                      ),
                                    ),
                                  children: [
                                    "Remove approval for ",
                                    source.target,
                                  ],
                                }),
                            ],
                          },
                          source.target,
                        ),
                      ),
                    })
                  : i.jsx("p", {
                      children:
                        "No source workbook is approved in this session.",
                    }),
                canAttachFile
                  ? i.jsxs("label", {
                      children: [
                        "Attach source workbook",
                        i.jsx("input", {
                          "aria-label": "Attach source workbook",
                          type: "file",
                          accept: ".xlsx,.xltx,.json,.csv,.tsv",
                          disabled: disabled,
                          onChange: (event) => {
                            const selectedFile = event.target.files?.[0];
                            ((event.target.value = ""),
                              selectedFile && attachSourceFile(selectedFile));
                          },
                        }),
                      ],
                    })
                  : i.jsx("p", {
                      children:
                        "Your host supplies source approvals. Ask the host application to attach an authorized snapshot.",
                    }),
                disabled &&
                  i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("p", {
                        role: "status",
                        children: "Inspecting selected source…",
                      }),
                      i.jsx("button", {
                        onClick: () => {
                          (requestCounter.current++,
                            cancellableOperation.cancel(),
                            setDisabled(false),
                            setWorkbook(null));
                        },
                        children: "Cancel source inspection",
                      }),
                    ],
                  }),
                workbook &&
                  i.jsxs("div", {
                    className: "lg-links-pending",
                    children: [
                      i.jsxs("h4", { children: ["Approve ", workbook.name] }),
                      i.jsxs("label", {
                        children: [
                          "Exact formula qualifier",
                          i.jsx("input", {
                            "aria-label": "Source qualifier",
                            value: target,
                            onChange: (event) => setTarget(event.target.value),
                          }),
                        ],
                      }),
                      i.jsx("p", {
                        children:
                          "Match the selected reference exactly. A matching filename alone does not grant access.",
                      }),
                      workbook.sheets.map((sheet) =>
                        i.jsxs(
                          "label",
                          {
                            className: "lg-links-check",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: allowedSheetIds.includes(sheet.id),
                                disabled: !!(sheet.hidden || sheet.veryHidden),
                                onChange: (event) =>
                                  setAllowedSheetIds(
                                    event.target.checked
                                      ? [...allowedSheetIds, sheet.id]
                                      : allowedSheetIds.filter(
                                          (allowedSheetId) =>
                                            allowedSheetId !== sheet.id,
                                        ),
                                  ),
                              }),
                              sheet.name,
                              sheet.hidden || sheet.veryHidden
                                ? " (hidden; host approval required)"
                                : "",
                            ],
                          },
                          sheet.id,
                        ),
                      ),
                      importWarnings.length > 0 &&
                        i.jsxs("details", {
                          children: [
                            i.jsxs("summary", {
                              children: [
                                importWarnings.length,
                                " import notices",
                              ],
                            }),
                            importWarnings.map((notice, noticeIndex) =>
                              i.jsx("p", { children: notice }, noticeIndex),
                            ),
                          ],
                        }),
                      i.jsxs("label", {
                        className: "lg-links-check",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: checked,
                            onChange: (event) =>
                              setChecked(event.target.checked),
                          }),
                          "I approve this source and the selected worksheets for explicit refresh",
                        ],
                      }),
                      i.jsx("button", {
                        disabled: !checked || !allowedSheetIds.length,
                        onClick: () =>
                          runOperation(() => {
                            const updatedSources = [
                              ...sources.filter(
                                (event) => event.target !== target,
                              ),
                              {
                                target: target,
                                workbook: workbook,
                                allowedSheetIds: allowedSheetIds,
                              },
                            ];
                            (buildWorkbookLinkReview(props.workbook, {
                              sources: updatedSources,
                            }),
                              updateApprovedSources(updatedSources),
                              setWorkbook(null));
                          }),
                        children: "Approve selected source",
                      }),
                      i.jsx("button", {
                        onClick: () => setWorkbook(null),
                        children: "Discard source",
                      }),
                    ],
                  }),
              ],
            }),
            workbookLinkPreview &&
              i.jsxs("section", {
                className: "lg-links-confirm",
                role: "region",
                "aria-label": "Workbook link preview",
                children: [
                  i.jsx("h3", {
                    children:
                      workbookLinkPreview.request.action === "break"
                        ? "Values that will replace formulas"
                        : "Reviewed results",
                  }),
                  i.jsx("p", { children: workbookLinkPreview.warning }),
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
                          children: workbookLinkPreview.cells
                            .slice(
                              previewPageIndex * 20,
                              previewPageIndex * 20 + 20,
                            )
                            .map((previewCell) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("td", {
                                      children: [
                                        previewCell.sheetName,
                                        "!",
                                        previewCell.address,
                                      ],
                                    }),
                                    i.jsx("td", {
                                      children: previewCell.before,
                                    }),
                                    i.jsx("td", {
                                      children: String(
                                        previewCell.after ?? "(blank)",
                                      ),
                                    }),
                                    i.jsxs("td", {
                                      children: [
                                        previewCell.result.type,
                                        " ·",
                                        " ",
                                        String(
                                          previewCell.result.value ?? "(blank)",
                                        ),
                                      ],
                                    }),
                                  ],
                                },
                                previewCell.sheetId + previewCell.address,
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
                        disabled: previewPageIndex === 0,
                        onClick: () =>
                          setPreviewPageIndex(previewPageIndex - 1),
                        children: "Previous preview cells",
                      }),
                      i.jsxs("span", {
                        children: [
                          workbookLinkPreview.cells.length,
                          " affected cells · preview page",
                          " ",
                          previewPageIndex + 1,
                          " of ",
                          Math.ceil(workbookLinkPreview.cells.length / 20),
                        ],
                      }),
                      i.jsx("button", {
                        disabled:
                          (previewPageIndex + 1) * 20 >=
                          workbookLinkPreview.cells.length,
                        onClick: () =>
                          setPreviewPageIndex(previewPageIndex + 1),
                        children: "Next preview cells",
                      }),
                    ],
                  }),
                  i.jsx("button", {
                    className: "lg-reports-primary",
                    disabled: props.readOnly || disabled,
                    onClick: applyLinkPreview,
                    children: "Apply reviewed link changes",
                  }),
                  i.jsx("button", {
                    onClick: () => setWorkbookLinkPreview(null),
                    children: "Cancel link preview",
                  }),
                ],
              }),
            i.jsxs("p", {
              children: [
                "Limits: ",
                workbookLinkLimits.sources,
                " sources,",
                " ",
                workbookLinkLimits.cells.toLocaleString(),
                " cached source cells and ",
                workbookLinkLimits.formulas,
                " destination formulas. Static A1 cells/ranges only. Linked arrays, external names, DDE and automatic/network refresh are unsupported. XLSX exports link formulas as inert text; native JSON retains the reviewed cache. Breaking links does not alter an attached original file.",
              ],
            }),
          ],
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("span", {
              children: props.readOnly
                ? "Read-only: inspection remains available."
                : "Every mutation requires a preview and confirmation.",
            }),
            i.jsx("button", { onClick: props.onClose, children: "Close" }),
          ],
        }),
      ],
    }),
  });
}
function Rue(file, options) {
  return runFileImport("source", file, options);
}
function toSourceFileDescriptor(inspection) {
  return {
    fileName: inspection.fileName,
    bytes: assertSourceInspectionCurrent(inspection),
    sha256: inspection.sha256,
  };
}
function parseExtractionResult(inspection, extractionPayload) {
  const sourceBytes = assertSourceInspectionCurrent(inspection);
  if (inspection.method !== "document-adapter")
    throw Error("JSON/XML use the native structured parser.");
  const extraction = object({
    schemaVersion: literal(1),
    sourceSha256: string(),
    extractor: string().min(1).max(200),
    pageCount: number().int().min(1).max(1e4),
    tables: array(unknown()).min(1).max(sourceImportLimits.tables),
  })
    .strict()
    .parse(extractionPayload);
  if (extraction.sourceSha256 !== inspection.sha256)
    throw Error("Extraction does not match the selected source hash.");
  if (["png", "jpeg"].includes(inspection.format) && extraction.pageCount !== 1)
    throw Error("An image has one source page.");
  const validatedInspection = validateExtractedSourceTables({
    ...inspection,
    extractor: extraction.extractor,
    pageCount: extraction.pageCount,
    tables: extraction.tables,
    warnings: [
      ...inspection.warnings,
      "Extraction values, confidence and regions are adapter assertions. Review the original; hashes do not prove extraction accuracy.",
    ],
  });
  for (const table of validatedInspection.tables)
    for (const row of table.rows)
      for (const tableCell of row)
        if (tableCell.region.page > extraction.pageCount)
          throw Error("A source region exceeds the declared page count.");
  return (
    validatedInspection.tables.length > 1 &&
      validatedInspection.tables.forEach((table) => (table.ambiguous = true)),
    recordSourceInspection(
      validatedInspection,
      sourceBytes,
      inspection.fileName,
    ),
    validatedInspection
  );
}
async function runDocumentExtraction(
  inspection,
  extractor,
  checked,
  options = {},
) {
  if (checked !== true)
    throw Error(
      "Approve the host extractor before sharing the selected source.",
    );
  if (inspection.method !== "document-adapter")
    throw Error("Choose a document source for extraction.");
  const sourceFileDescriptor = toSourceFileDescriptor(inspection);
  if (options.signal?.aborted) throw Error("Document extraction cancelled.");
  const timeoutMs = options.timeoutMs ?? 3e4;
  if (!Number.isInteger(timeoutMs) || timeoutMs < 1 || timeoutMs > 6e4)
    throw Error("Invalid document extraction time budget.");
  const extractionController = new AbortController(),
    abortExtraction = () => extractionController.abort();
  options.signal?.addEventListener("abort", abortExtraction, { once: true });
  let timeoutId;
  try {
    return await new Promise((resolve, reject) => {
      let isSettled = false;
      const settle = (failure, value) => {
          isSettled ||
            ((isSettled = true),
            extractionController.signal.removeEventListener("abort", onAborted),
            extractionController.abort(),
            failure ? reject(failure) : resolve(value));
        },
        onAborted = () => settle(Error("Document extraction cancelled."));
      (extractionController.signal.addEventListener("abort", onAborted, {
        once: true,
      }),
        (timeoutId = setTimeout(
          () => settle(Error("Document extraction exceeded its time budget.")),
          timeoutMs,
        )),
        Promise.resolve()
          .then(() =>
            extractor.extract({
              file: new File(
                [new Uint8Array(sourceFileDescriptor.bytes)],
                sourceFileDescriptor.fileName,
              ),
              sha256: inspection.sha256,
              signal: extractionController.signal,
            }),
          )
          .then(
            (extractorResponse) => {
              if (!isSettled)
                try {
                  settle(
                    undefined,
                    parseExtractionResult(inspection, extractorResponse),
                  );
                } catch {
                  settle(
                    Error(
                      "Extraction has invalid types, confidence, regions or source identity.",
                    ),
                  );
                }
            },
            () =>
              settle(
                Error("Document extraction failed. The workbook is unchanged."),
              ),
          ));
    });
  } finally {
    (clearTimeout(timeoutId),
      options.signal?.removeEventListener("abort", abortExtraction),
      extractionController.abort());
  }
}
const importPreviewRegistry = new WeakMap();
function buildSourceReview(workbook, inspection, review) {
  const sourceBytes = assertSourceInspectionCurrent(inspection);
  if (
    (validateExtractedSourceTables(inspection),
    assertWorkbookStructureEditable(workbook),
    typeof review.sheetName != "string" ||
      !review.sheetName.trim() ||
      review.sheetName.length > 31 ||
      /[\\/?*\[\]:\x00-\x1f]/.test(review.sheetName) ||
      workbook.sheets.some(
        (sheet) =>
          sheet.name.toLowerCase() === review.sheetName.trim().toLowerCase(),
      ))
  )
    throw Error("Choose a unique worksheet name of 1–31 characters.");
  const table = inspection.tables.find((table) => table.id === review.tableId);
  if (!table) throw Error("Select a source table.");
  if (
    (table.ambiguous || inspection.tables.length > 1) &&
    review.ambiguityReviewed !== true
  )
    throw Error("Review and confirm the ambiguous table selection.");
  if (!review.resolutions || typeof review.resolutions != "object")
    throw Error("Invalid source review.");
  const cells = {},
    reviewedCellKeys = new Set();
  let accepted = 0,
    excluded = 0,
    corrected = 0;
  if (
    (table.columns.forEach((column, headerColumnIndex) => {
      cells[columnIndexToLetters(headerColumnIndex) + "1"] = {
        raw: column.name,
        literal: true,
        style: { bold: true },
      };
    }),
    table.rows.forEach((row, rowIndex) =>
      row.forEach((sourceCell, columnIndex) => {
        const resolutionKey = `${rowIndex}:${columnIndex}`;
        reviewedCellKeys.add(resolutionKey);
        const cellResolution = review.resolutions[resolutionKey],
          extractedCellValue = resolveExtractedCellValue(
            sourceCell,
            cellResolution,
          );
        (cellResolution?.decision === "exclude"
          ? excluded++
          : cellResolution?.decision === "correct"
            ? corrected++
            : accepted++,
          (cells[columnIndexToLetters(columnIndex) + (rowIndex + 2)] = {
            raw: extractedCellValue,
            ...(typeof extractedCellValue == "string" ? { literal: true } : {}),
          }));
      }),
    ),
    Object.keys(review.resolutions).some(
      (resolution) => !reviewedCellKeys.has(resolution),
    ))
  )
    throw Error("A review decision refers to an unknown cell.");
  const sheet = {
      id: crypto.randomUUID(),
      name: review.sheetName.trim(),
      cells: cells,
      rowCount: workbookLimits.maxRows,
      colCount: workbookLimits.maxCols,
    },
    record = {
      id: crypto.randomUUID(),
      sheetId: sheet.id,
      sheetName: sheet.name,
      importedAt: new Date().toISOString(),
      original: buildRetainedSourceRecord(
        sourceBytes,
        inspection.fileName,
        inspection.format,
      ),
      method: inspection.method,
      extractor: inspection.extractor,
      ...(inspection.pageCount ? { pageCount: inspection.pageCount } : {}),
      table: structuredClone(table),
      resolutions: structuredClone(review.resolutions),
      ambiguityReviewed: review.ambiguityReviewed,
    };
  normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, sheet],
    sourceImports: [...(workbook.sourceImports ?? []), record],
  });
  const sourceReview = {
    workbookHash: fingerprintCanonicalValue(workbook),
    inspection: inspection,
    review: structuredClone(review),
    record: record,
    sheet: sheet,
    accepted: accepted,
    excluded: excluded,
    corrected: corrected,
  };
  return (
    importPreviewRegistry.set(
      sourceReview,
      fingerprintCanonicalValue(sourceReview),
    ),
    sourceReview
  );
}
function commitSourceImport(workbook, preview, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (
    (assertWorkbookStructureEditable(workbook),
    importPreviewRegistry.get(preview) !== fingerprintCanonicalValue(preview) ||
      preview.workbookHash !== fingerprintCanonicalValue(workbook))
  )
    throw Error("The workbook or review changed. Preview the import again.");
  assertSourceInspectionCurrent(preview.inspection);
  const workbookInput = normalizeWorkbookInput({
    ...workbook,
    sheets: [...workbook.sheets, preview.sheet],
    sourceImports: [...(workbook.sourceImports ?? []), preview.record],
  });
  return (importPreviewRegistry.delete(preview), workbookInput);
}
function findSourceImport(workbook, sourceImportId) {
  const foundImport = workbook.sourceImports?.find(
    (sourceImport) => sourceImport.id === sourceImportId,
  );
  if (!foundImport) throw Error("Source import is unavailable.");
  return {
    bytes: validateRetainedSourceFile(foundImport.original),
    fileName: foundImport.original.fileName,
    sha256: foundImport.original.sha256,
  };
}
function resolveImportedSourceCell(
  workbook,
  sourceImportId,
  rowIndex,
  columnIndex,
) {
  const importRecord = workbook.sourceImports?.find(
      (sourceImport) => sourceImport.id === sourceImportId,
    ),
    source = importRecord?.table.rows[rowIndex]?.[columnIndex];
  if (!importRecord || !source)
    throw Error("Imported source cell is unavailable.");
  const address = columnIndexToLetters(columnIndex) + (rowIndex + 2),
    sheet = workbook.sheets.find((sheet) => sheet.id === importRecord.sheetId),
    cell = sheet?.cells[address],
    resolution = importRecord.resolutions[`${rowIndex}:${columnIndex}`],
    extractedCellValue = resolveExtractedCellValue(source, resolution);
  return {
    address: address,
    sheetId: importRecord.sheetId,
    sourceHash: importRecord.original.sha256,
    source: source,
    decision: resolution?.decision ?? "accepted",
    importedValue: extractedCellValue,
    currentValue: cell?.raw ?? null,
    status: sheet
      ? cell?.raw === extractedCellValue &&
        (typeof extractedCellValue != "string" || cell.literal === true)
        ? "matches-import"
        : "changed"
      : "missing-sheet",
  };
}
const downloadBlobAsFile = ({ filename, blob }) => {
  const objectUrl = URL.createObjectURL(blob),
    link = document.createElement("a");
  ((link.href = objectUrl),
    (link.download = filename),
    link.click(),
    setTimeout(() => URL.revokeObjectURL(objectUrl), 1e3));
};

export {
  commitSourceImport,
  useOperationScope,
  toSourceFileDescriptor,
  runDocumentExtraction,
  resolveImportedSourceCell,
  WorkbookLinksDialog,
  findSourceImport,
  Rue,
  buildSourceImportRecord,
  WorkbookLinkStatusBadge,
  operationScopeContext,
  buildSourceReview,
  downloadBlobAsFile,
  parseExtractionResult,
  useCancellableOperation,
};
