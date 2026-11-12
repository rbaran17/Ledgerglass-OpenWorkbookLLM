import { isPlainRecord } from "./19b-assistant-protocol.js";

import { downloadBlobAsFile } from "./09-workbook-links.js";
import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  deepFreezeValue,
  measureSerializedBytes,
  defineDataPropertyByKey,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  builtinFunctionNames,
  requireAlternateFormatOriginal,
  parseRangeAddress,
  touchWorkbookModifiedAt,
  buildCanonicalWorkbookSnapshot,
  normalizeWorkbookInput,
  columnIndexToLetters,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  importWorkbookFile,
  describeImportSource,
  templateConversionRoutes,
  detectSupportedFileFormats,
} from "./07-file-import-and-worker-dispatch.js";
import {
  exportXltxTemplate,
  exportWorkbookToXlsx,
} from "./08-xlsx-import-export-and-preservation.js";
import { useCancellableOperation } from "./09-workbook-links.js";
import {
  validateConvertedSource,
  describeSourceFile,
  XlsxConversionReview,
} from "./10-source-import-dialog.js";

function FormatConversionDialog(props) {
  const { text } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    cancellableOperation = useCancellableOperation(),
    latestPropsRef = ne.useRef(props),
    requestCounterRef = ne.useRef(0),
    isMountedRef = ne.useRef(true);
  latestPropsRef.current = props;
  const [sourceFile, setSourceFile] = ne.useState(null),
    [inspectedFile, setInspectedFile] = ne.useState(null),
    [checked, setChecked] = ne.useState(false),
    [disabled, setDisabled] = ne.useState(false),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [pendingReview, setPendingReview] = ne.useState(null),
    [successMessage, setSuccessMessage] = ne.useState("");
  (ne.useEffect(() => {
    isMountedRef.current = true;
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.focus(),
      () => {
        ((isMountedRef.current = false),
          requestCounterRef.current++,
          activeElement?.isConnected && activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      pendingReview || dialogRef.current?.focus();
    }, [pendingReview]));
  const runOperation = async (operation) => {
      (setErrorMessage(""), setSuccessMessage(""), setDisabled(true));
      try {
        await operation();
      } catch (error) {
        isMountedRef.current && setErrorMessage(error.message);
      } finally {
        isMountedRef.current && setDisabled(false);
      }
    },
    inspectFile = (file) => {
      const requestId = ++requestCounterRef.current;
      return (
        setSourceFile(null),
        setInspectedFile(null),
        setChecked(false),
        setPendingReview(null),
        runOperation(async () => {
          const detectedSource = await detectSupportedFileFormats(
            file,
            cancellableOperation.begin(),
          );
          isMountedRef.current &&
            requestId === requestCounterRef.current &&
            (setSourceFile(detectedSource), setInspectedFile(file));
        })
      );
    };
  ne.useEffect(() => {
    props.initialFile && inspectFile(props.initialFile);
  }, [props.initialFile]);
  const reviewImportFile = (workbookFile, isTemplatePreview) =>
      runOperation(async () => {
        if (latestPropsRef.current.readOnly)
          throw Error("This workbook is read-only.");
        const workbook = latestPropsRef.current.workbook,
          importRequestId = ++requestCounterRef.current,
          imported = isTemplatePreview
            ? await importWorkbookFile(
                workbookFile,
                cancellableOperation.begin(),
              )
            : await validateConvertedSource(
                sourceFile,
                workbookFile,
                checked,
                cancellableOperation.begin(),
              );
        if (!(
          !isMountedRef.current || importRequestId !== requestCounterRef.current
        )) {
          if (
            latestPropsRef.current.workbook !== workbook ||
            latestPropsRef.current.readOnly
          )
            throw Error(
              "The workbook or permissions changed. Review the conversion again.",
            );
          if (!imported.report)
            throw Error(
              "A compatibility report is required for format conversion.",
            );
          setPendingReview({
            snapshot: workbook,
            imported: imported,
            report: imported.report,
          });
        }
      }),
    reviewExport = (isTemplateExport) =>
      runOperation(async () => {
        const workbook = latestPropsRef.current.workbook,
          exportRequestId = ++requestCounterRef.current,
          exported = isTemplateExport
            ? await exportXltxTemplate(workbook)
            : await exportWorkbookToXlsx(workbook);
        if (!(
          !isMountedRef.current || exportRequestId !== requestCounterRef.current
        )) {
          if (latestPropsRef.current.workbook !== workbook)
            throw Error("The workbook changed. Review the export again.");
          setPendingReview({
            snapshot: workbook,
            blob: exported.blob,
            report: {
              ...exported.report,
              fileName: describeImportSource(
                workbook,
                isTemplateExport ? "xltx" : "xlsx",
              ),
            },
          });
        }
      }),
    downloadFile = (filename, blob) =>
      (latestPropsRef.current.onDownload ?? downloadBlobAsFile)({
        filename: filename,
        blob: blob,
      }),
    downloadOriginal = (isWorkbookOriginal) =>
      runOperation(async () => {
        const original = isWorkbookOriginal
          ? requireAlternateFormatOriginal(latestPropsRef.current.workbook)
          : describeSourceFile(sourceFile);
        await downloadFile(
          original.fileName,
          new Blob([new Uint8Array(original.bytes)], {
            type: "application/octet-stream",
          }),
        );
      });
  return pendingReview
    ? i.jsx(XlsxConversionReview, {
        report: pendingReview.report,
        originalLabel: sourceFile
          ? `Download original ${sourceFile.format.toUpperCase()}`
          : undefined,
        blocked:
          errorMessage ||
          (pendingReview.snapshot !== props.workbook
            ? "The workbook changed. Cancel and review again."
            : pendingReview.imported && props.readOnly
              ? "This workbook is read-only."
              : disabled
                ? "Completing the reviewed operation."
                : undefined),
        onClose: () => {
          (setPendingReview(null), setErrorMessage(""));
        },
        onOriginal:
          sourceFile && pendingReview.imported
            ? () => void downloadOriginal(false)
            : undefined,
        onReport: () =>
          void runOperation(async () => {
            await downloadFile(
              pendingReview.report.fileName + ".compatibility.json",
              new Blob([JSON.stringify(pendingReview.report, null, 2)], {
                type: "application/json",
              }),
            );
          }),
        onConfirm: () =>
          void runOperation(async () => {
            if (latestPropsRef.current.workbook !== pendingReview.snapshot)
              throw Error("The workbook changed. Review the conversion again.");
            if (pendingReview.imported) {
              if (latestPropsRef.current.readOnly)
                throw Error("This workbook is read-only.");
              (latestPropsRef.current.onWorkbookChange(
                pendingReview.imported.workbook,
                "Open reviewed workbook conversion",
              ),
                setSuccessMessage(
                  "Converted workbook opened. Keep a native JSON backup to retain both source files and conversion provenance.",
                ));
            } else
              pendingReview.blob &&
                (await downloadFile(
                  pendingReview.report.fileName,
                  pendingReview.blob,
                ),
                setSuccessMessage(
                  "Derived workbook downloaded. Verify fidelity in the source application.",
                ));
            setPendingReview(null);
          }),
      })
    : i.jsx("div", {
        className: "lg-reports-overlay",
        children: i.jsxs("section", {
          role: "dialog",
          "aria-modal": "true",
          "aria-labelledby": titleId,
          lang: "en",
          tabIndex: -1,
          ref: dialogRef,
          className: "lg-reports-dialog lg-formats-dialog",
          onKeyDown: (event) => {
            var lastFocusable, firstFocusable;
            if (
              (event.key === "Escape" &&
                (event.stopPropagation(), props.onClose()),
              event.key === "Tab")
            ) {
              const focusableElements = [
                ...dialogRef.current.querySelectorAll(
                  'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
                ),
              ];
              event.shiftKey &&
              (document.activeElement === focusableElements[0] ||
                document.activeElement === dialogRef.current)
                ? (event.preventDefault(),
                  (lastFocusable = focusableElements.at(-1)) == null ||
                    lastFocusable.focus())
                : !event.shiftKey &&
                  document.activeElement === focusableElements.at(-1) &&
                  (event.preventDefault(),
                  (firstFocusable = focusableElements[0]) == null ||
                    firstFocusable.focus());
            }
          },
          children: [
            i.jsxs("header", {
              children: [
                i.jsxs("div", {
                  children: [
                    i.jsx("p", { children: "WORKBOOK EXCHANGE" }),
                    i.jsx("h2", {
                      id: titleId,
                      children: i.jsx(LocalizedText, {
                        children: "Workbook formats",
                      }),
                    }),
                  ],
                }),
                i.jsx("button", {
                  "aria-label": text("Close workbook formats"),
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
                    "Open a template or bring an older workbook into this editor. Review conversion losses and keep the untouched original with your work.",
                }),
                errorMessage &&
                  i.jsx("p", {
                    role: "alert",
                    className: "lg-formats-error",
                    children: errorMessage,
                  }),
                successMessage &&
                  i.jsx("p", { role: "status", children: successMessage }),
                disabled &&
                  i.jsx("p", {
                    role: "status",
                    children: "Processing workbook…",
                  }),
                i.jsxs("div", {
                  className: "lg-formats-actions",
                  children: [
                    i.jsxs("label", {
                      children: [
                        "Inspect source file",
                        i.jsx("input", {
                          type: "file",
                          accept: ".xls,.xlsb,.xlsm,.xltx,.xltm,.ods",
                          "aria-label": text("Inspect alternate workbook"),
                          disabled: disabled,
                          onChange: (event) => {
                            const pickedInspectFile = event.target.files?.[0];
                            ((event.target.value = ""),
                              pickedInspectFile &&
                                inspectFile(pickedInspectFile));
                          },
                        }),
                      ],
                    }),
                    i.jsx("button", {
                      disabled: disabled,
                      onClick: () => void reviewExport(false),
                      children: "Review XLSX export",
                    }),
                    i.jsx("button", {
                      disabled: disabled,
                      onClick: () => void reviewExport(true),
                      children: "Review XLTX export",
                    }),
                  ],
                }),
                props.workbook.alternateSource &&
                  i.jsxs("section", {
                    className: "lg-formats-source",
                    children: [
                      i.jsx("h3", { children: "Retained original" }),
                      i.jsxs("p", {
                        children: [
                          props.workbook.alternateSource.fileName,
                          " ·",
                          " ",
                          props.workbook.alternateSource.conversion.method,
                          " · equivalence not verified",
                        ],
                      }),
                      i.jsxs("p", {
                        className: "lg-formats-hash",
                        children: [
                          "SHA-256: ",
                          props.workbook.alternateSource.sha256,
                        ],
                      }),
                      i.jsx("button", {
                        disabled: disabled,
                        onClick: () => void downloadOriginal(true),
                        children: "Download retained original",
                      }),
                    ],
                  }),
                sourceFile &&
                  i.jsxs("section", {
                    className: "lg-formats-source",
                    children: [
                      i.jsx("h3", { children: sourceFile.fileName }),
                      i.jsxs("p", {
                        children: [
                          sourceFile.format.toUpperCase(),
                          " identified by content ·",
                          " ",
                          (sourceFile.fileBytes / 1024).toFixed(1),
                          " KiB ·",
                          " ",
                          sourceFile.parts.length,
                          " inventory entries",
                        ],
                      }),
                      i.jsx("p", { children: sourceFile.identity }),
                      i.jsxs("p", {
                        className: "lg-formats-hash",
                        children: ["SHA-256: ", sourceFile.sha256],
                      }),
                      i.jsxs("div", {
                        className: "lg-formats-actions",
                        children: [
                          i.jsx("button", {
                            disabled: disabled,
                            onClick: () => void downloadOriginal(false),
                            children: "Download inspected original",
                          }),
                          i.jsx("button", {
                            disabled: disabled,
                            onClick: () =>
                              void runOperation(async () => {
                                await downloadFile(
                                  sourceFile.fileName + ".format-review.json",
                                  new Blob(
                                    [JSON.stringify(sourceFile, null, 2)],
                                    {
                                      type: "application/json",
                                    },
                                  ),
                                );
                              }),
                            children: "Download format report",
                          }),
                          props.onAutomation &&
                            ["xls", "xlsb", "xlsm", "xltx", "xltm"].includes(
                              sourceFile.format,
                            ) &&
                            i.jsx("button", {
                              disabled: disabled,
                              onClick: () =>
                                void runOperation(async () => {
                                  await props.onAutomation(inspectedFile);
                                }),
                              children: "Inspect automation",
                            }),
                        ],
                      }),
                      i.jsx("ul", {
                        children: sourceFile.warnings.map((warning) =>
                          i.jsx("li", { children: warning }, warning),
                        ),
                      }),
                      sourceFile.format === "xltx" &&
                        i.jsx("p", {
                          children: i.jsx("button", {
                            className: "lg-reports-primary",
                            disabled: disabled || props.readOnly,
                            onClick: () =>
                              void reviewImportFile(inspectedFile, true),
                            children: "Preview template workbook",
                          }),
                        }),
                      i.jsxs("details", {
                        open: sourceFile.format !== "xltx",
                        children: [
                          i.jsx("summary", {
                            children: "Convert with the source application",
                          }),
                          i.jsx("ol", {
                            children: sourceFile.steps.map((step) =>
                              i.jsx("li", { children: step }, step),
                            ),
                          }),
                          i.jsxs("label", {
                            className: "lg-formats-ack",
                            children: [
                              i.jsx("input", {
                                type: "checkbox",
                                checked: checked,
                                disabled: disabled || props.readOnly,
                                onChange: (event) =>
                                  setChecked(event.target.checked),
                              }),
                              "I acknowledge possible conversion losses and will compare formulas, dates, styles and special parts with the original.",
                            ],
                          }),
                          i.jsxs("label", {
                            children: [
                              "Choose converted XLSX",
                              i.jsx("input", {
                                type: "file",
                                accept: ".xlsx",
                                "aria-label": text("Choose converted XLSX"),
                                disabled:
                                  disabled || props.readOnly || !checked,
                                onChange: (event) => {
                                  const pickedConvertedFile =
                                    event.target.files?.[0];
                                  ((event.target.value = ""),
                                    pickedConvertedFile &&
                                      reviewImportFile(
                                        pickedConvertedFile,
                                        false,
                                      ));
                                },
                              }),
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("details", {
                        children: [
                          i.jsxs("summary", {
                            children: [
                              "Source inventory (",
                              sourceFile.parts.length,
                              ")",
                            ],
                          }),
                          i.jsx("p", {
                            children:
                              "Original-only inventory; no content is executed. This is not a lossless conversion claim.",
                          }),
                          i.jsx("ul", {
                            className: "lg-formats-inventory",
                            children: sourceFile.parts.map((part, partIndex) =>
                              i.jsxs(
                                "li",
                                {
                                  children: [
                                    part.path,
                                    " · ",
                                    part.bytes,
                                    " known bytes ·",
                                    " ",
                                    part.disposition,
                                  ],
                                },
                                partIndex,
                              ),
                            ),
                          }),
                        ],
                      }),
                    ],
                  }),
                i.jsx("h3", {
                  children: "Read, write and preservation support",
                }),
                i.jsx("div", {
                  className: "lg-formats-table",
                  children: i.jsxs("table", {
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", { scope: "col", children: "Format" }),
                            i.jsx("th", { scope: "col", children: "Read" }),
                            i.jsx("th", { scope: "col", children: "Write" }),
                            i.jsx("th", { scope: "col", children: "Preserve" }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: templateConversionRoutes.map(
                          (templateConversionRoute) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsxs("th", {
                                    scope: "row",
                                    children: [
                                      ".",
                                      templateConversionRoute.format,
                                    ],
                                  }),
                                  i.jsx("td", {
                                    children: templateConversionRoute.read,
                                  }),
                                  i.jsx("td", {
                                    children: templateConversionRoute.write,
                                  }),
                                  i.jsx("td", {
                                    children: templateConversionRoute.preserve,
                                  }),
                                ],
                              },
                              templateConversionRoute.format,
                            ),
                        ),
                      }),
                    ],
                  }),
                }),
                i.jsx("p", {
                  children:
                    "Files: up to 10 MiB. ZIP packages: up to 2,000 entries and 50 MiB expanded. Password-protected containers require an unencrypted copy. External conversions and native Office fidelity must be validated in the source application.",
                }),
              ],
            }),
            i.jsxs("footer", {
              children: [
                i.jsx("span", {
                  children: props.readOnly
                    ? "Read-only: inspection and export available; import disabled."
                    : "Imports replace the current workbook after review. Undo restores the previous workbook.",
                }),
                i.jsx("button", { onClick: props.onClose, children: "Close" }),
              ],
            }),
          ],
        }),
      });
}
class WorkbookRevisionConflictError extends Error {
  constructor(code, message, currentWorkbook) {
    super(
      currentWorkbook
        ? `${message} Current saved workbook ${currentWorkbook.workbookId}, revision ${currentWorkbook.revision}.`
        : message,
    );
    defineDataPropertyByKey(this, "currentRevision");
    defineDataPropertyByKey(this, "currentWorkbookId");
    ((this.code = code),
      (this.name = "PersistenceError"),
      (this.currentRevision = currentWorkbook?.revision),
      (this.currentWorkbookId = currentWorkbook?.workbookId));
  }
}
const localWorkbookNamespace = "ledgerglass-local-workbook",
  maxLocalCheckpoints = 20,
  cloneStructuredValue = (value) => structuredClone(value),
  measureUtf8Bytes = (text) => new TextEncoder().encode(text).byteLength;
function throwPersistenceError(code, message, guard) {
  throw new WorkbookRevisionConflictError(code, message, guard);
}
function stringifyCanonicalJson(value) {
  return value === undefined
    ? "undefined"
    : value === null || typeof value != "object"
      ? JSON.stringify(value)
      : Array.isArray(value)
        ? `[${value.map(stringifyCanonicalJson).join(",")}]`
        : `{${Object.keys(value)
            .sort()
            .filter((key) => value[key] !== undefined)
            .map(
              (propertyKey) =>
                `${JSON.stringify(propertyKey)}:${stringifyCanonicalJson(value[propertyKey])}`,
            )
            .join(",")}}`;
}
function computeTextDigest(text) {
  let hashA = 2654435769,
    hashB = 2246822507,
    hashC = 3266489909,
    hashD = 668265263;
  for (let charIndex = 0; charIndex < text.length; charIndex++) {
    const charCode = text.charCodeAt(charIndex);
    ((hashA = Math.imul(hashA ^ charCode, 2246822507)),
      (hashB = Math.imul(hashB ^ charCode, 3266489909)),
      (hashC = Math.imul(hashC ^ charCode, 668265263)),
      (hashD = Math.imul(hashD ^ charCode, 374761393)));
  }
  return [text.length, hashA, hashB, hashC, hashD]
    .map((hashValue) => (hashValue >>> 0).toString(16).padStart(8, "0"))
    .join("");
}
function normalizeWorkbookSnapshot(workbook) {
  return normalizeWorkbookInput(workbook);
}
function normalizeSavedSelection(workbookSnapshot, view) {
  const sheet = workbookSnapshot.sheets.find((sheet) => !sheet.hidden);
  if (!sheet)
    return throwPersistenceError(
      "invalid",
      "A saved workbook needs a visible worksheet.",
    );
  if (view === undefined)
    return {
      activeSheetId: sheet.id,
      selection: { sheetId: sheet.id, range: "A1" },
    };
  if (
    !isPlainRecord(view) ||
    typeof view.activeSheetId != "string" ||
    !isPlainRecord(view.selection)
  )
    return throwPersistenceError(
      "invalid",
      "Invalid saved worksheet selection.",
    );
  const activeSheet = workbookSnapshot.sheets.find(
    (sheet) => sheet.id === view.activeSheetId && !sheet.hidden,
  );
  if (
    !activeSheet ||
    view.selection.sheetId !== activeSheet.id ||
    typeof view.selection.range != "string"
  )
    return throwPersistenceError(
      "invalid",
      "Saved selection must belong to the active visible worksheet.",
    );
  let parsedRange;
  try {
    parsedRange = parseRangeAddress(view.selection.range);
  } catch {
    return throwPersistenceError("invalid", "Invalid saved cell range.");
  }
  if (
    parsedRange.lastRow >= activeSheet.rowCount ||
    parsedRange.lastCol >= activeSheet.colCount
  )
    return throwPersistenceError(
      "invalid",
      "Saved selection exceeds worksheet dimensions.",
    );
  const startAddress = `${columnIndexToLetters(parsedRange.firstCol)}${parsedRange.firstRow + 1}`,
    endAddress = `${columnIndexToLetters(parsedRange.lastCol)}${parsedRange.lastRow + 1}`;
  return {
    activeSheetId: activeSheet.id,
    selection: {
      sheetId: activeSheet.id,
      range:
        startAddress === endAddress
          ? startAddress
          : `${startAddress}:${endAddress}`,
    },
  };
}
function createWorkbookSnapshot(
  workbook,
  view,
  savedAt = new Date().toISOString(),
) {
  const workbookSnapshot = normalizeWorkbookSnapshot(workbook),
    savedSelection = normalizeSavedSelection(workbookSnapshot, view),
    unsignedSnapshot = {
      workbook: workbookSnapshot,
      view: savedSelection,
      savedAt: savedAt,
    };
  return {
    ...unsignedSnapshot,
    digest: computeTextDigest(stringifyCanonicalJson(unsignedSnapshot)),
  };
}
function parseWorkbookSnapshot(storedSnapshot) {
  if (
    !isPlainRecord(storedSnapshot) ||
    !isPlainRecord(storedSnapshot.workbook) ||
    typeof storedSnapshot.savedAt != "string" ||
    !Number.isFinite(Date.parse(storedSnapshot.savedAt)) ||
    typeof storedSnapshot.digest != "string"
  )
    return throwPersistenceError(
      "corrupt",
      "Stored workbook snapshot is incomplete.",
    );
  if (
    computeTextDigest(
      stringifyCanonicalJson({
        workbook: storedSnapshot.workbook,
        view: storedSnapshot.view,
        savedAt: storedSnapshot.savedAt,
      }),
    ) !== storedSnapshot.digest
  )
    return throwPersistenceError(
      "corrupt",
      "Stored workbook snapshot failed its integrity check.",
    );
  try {
    return createWorkbookSnapshot(
      normalizeWorkbookSnapshot(storedSnapshot.workbook),
      storedSnapshot.view,
      storedSnapshot.savedAt,
    );
  } catch {
    return throwPersistenceError(
      "corrupt",
      "Stored workbook snapshot does not satisfy workbook validation.",
    );
  }
}
function createCheckpointRecord(checkpoint) {
  return {
    id: checkpoint.id,
    label: checkpoint.label,
    createdAt: checkpoint.createdAt,
    workbookId: checkpoint.snapshot.workbook.id,
    revision: checkpoint.snapshot.workbook.revision,
  };
}
function compareWorkbookVersions(fromWorkbook, workbook) {
  const workbookSnapshot = normalizeWorkbookSnapshot(fromWorkbook),
    toWorkbookSnapshot = normalizeWorkbookSnapshot(workbook);
  if (workbookSnapshot.id !== toWorkbookSnapshot.id)
    return throwPersistenceError(
      "invalid",
      "Version comparison requires the same workbook identity.",
    );
  const comparison = {
      fromRevision: workbookSnapshot.revision,
      toRevision: toWorkbookSnapshot.revision,
      cellChanges: [],
      metadataChanges: [],
      structuralChanges: [],
    },
    valuesDiffer = (leftValue, rightValue) =>
      stringifyCanonicalJson(leftValue) !== stringifyCanonicalJson(rightValue),
    fromEventIds = new Set(
      workbookSnapshot.events.map((fromEvent) => fromEvent.id),
    ),
    newEvents = toWorkbookSnapshot.events.filter(
      (toEvent) => !fromEventIds.has(toEvent.id),
    ),
    changeOrigins = new Map();
  for (const newEvent of newEvents)
    for (const change of newEvent.changes)
      changeOrigins.set(`${change.sheetId}:${change.address}`, {
        event: newEvent,
        change: change,
      });
  const makeAttribution = (sourceEvent, basis) =>
      sourceEvent
        ? {
            basis: basis,
            actor: sourceEvent.actor,
            timestamp: sourceEvent.timestamp,
            eventId: sourceEvent.id,
          }
        : { basis: "unavailable" },
    operationAttribution = (actionPattern) =>
      makeAttribution(
        [...newEvents]
          .reverse()
          .find((historyEvent) => actionPattern.test(historyEvent.action)),
        "operation-history",
      ),
    structuralAttribution = () =>
      operationAttribution(
        /sheet|structure|tab color|visibility|restore|row|column|merge|drawing|picture|shape|chart|validation|conditional|page|table|pivot|query/i,
      );
  for (const path of new Set([
    ...Object.keys(workbookSnapshot),
    ...Object.keys(toWorkbookSnapshot),
  ])) {
    if (["sheets", "events", "revision", "id", "schemaVersion"].includes(path))
      continue;
    const fromValue = workbookSnapshot[path],
      toValue = toWorkbookSnapshot[path];
    valuesDiffer(fromValue, toValue) &&
      comparison.metadataChanges.push({
        path: path,
        before: cloneStructuredValue(fromValue),
        after: cloneStructuredValue(toValue),
        attribution: operationAttribution(
          path === "definedNames"
            ? /name|formula|restore/i
            : path === "formulaWatches"
              ? /watch|row|column|sheet|move|cut|restore/i
              : path === "styleCatalog"
                ? /theme|style|restore/i
                : /workbook|properties|engagement|period|template|restore/i,
        ),
      });
  }
  valuesDiffer(
    workbookSnapshot.sheets.map((sheet) => sheet.id),
    toWorkbookSnapshot.sheets.map((sheet) => sheet.id),
  ) &&
    comparison.structuralChanges.push({
      path: "sheets.order",
      before: workbookSnapshot.sheets.map((sheet) => sheet.id),
      after: toWorkbookSnapshot.sheets.map((sheet) => sheet.id),
      attribution: structuralAttribution(),
    });
  for (const sheetId of new Set([
    ...workbookSnapshot.sheets.map((sheet) => sheet.id),
    ...toWorkbookSnapshot.sheets.map((sheet) => sheet.id),
  ])) {
    const sheet = workbookSnapshot.sheets.find((sheet) => sheet.id === sheetId),
      toSheet = toWorkbookSnapshot.sheets.find((sheet) => sheet.id === sheetId),
      omitCells = (sheetRecord) =>
        sheetRecord
          ? Object.fromEntries(
              Object.entries(sheetRecord).filter(
                ([entryKey]) => entryKey !== "cells",
              ),
            )
          : undefined,
      before = omitCells(sheet),
      after = omitCells(toSheet);
    if (!sheet || !toSheet)
      comparison.structuralChanges.push({
        path: `sheets.${sheetId}`,
        before: before,
        after: after,
        attribution: structuralAttribution(),
      });
    else
      for (const sheetProperty of new Set([
        ...Object.keys(before),
        ...Object.keys(after),
      ]))
        valuesDiffer(before[sheetProperty], after[sheetProperty]) &&
          comparison.structuralChanges.push({
            path: `sheets.${sheetId}.${sheetProperty}`,
            before: cloneStructuredValue(before[sheetProperty]),
            after: cloneStructuredValue(after[sheetProperty]),
            attribution: structuralAttribution(),
          });
    for (const address of new Set([
      ...Object.keys(sheet?.cells ?? {}),
      ...Object.keys(toSheet?.cells ?? {}),
    ])) {
      const fromCell = sheet?.cells[address],
        toCell = toSheet?.cells[address];
      valuesDiffer(fromCell, toCell) &&
        comparison.cellChanges.push({
          sheetId: sheetId,
          address: address,
          ...(fromCell === undefined
            ? {}
            : { before: cloneStructuredValue(fromCell) }),
          ...(toCell === undefined
            ? {}
            : { after: cloneStructuredValue(toCell) }),
          attribution: makeAttribution(
            stringifyCanonicalJson(
              changeOrigins.get(`${sheetId}:${address}`)?.change.after,
            ) === stringifyCanonicalJson(toCell)
              ? changeOrigins.get(`${sheetId}:${address}`)?.event
              : undefined,
            "cell-history",
          ),
        });
    }
  }
  return comparison;
}
function createLocalWorkbookStore(storage, namespace, options = {}) {
  if (
    typeof namespace != "string" ||
    !namespace.trim() ||
    namespace.length > 300
  )
    return throwPersistenceError(
      "invalid",
      "A local workbook namespace must contain 1–300 characters.",
    );
  const maxCheckpoints = options.maxCheckpoints ?? maxLocalCheckpoints,
    maxBytes = options.maxBytes ?? 10 * 1024 * 1024;
  if (
    !Number.isInteger(maxCheckpoints) ||
    maxCheckpoints < 1 ||
    maxCheckpoints > maxLocalCheckpoints ||
    !Number.isInteger(maxBytes) ||
    maxBytes < 1024 ||
    maxBytes > 50 * 1024 * 1024
  )
    return throwPersistenceError(
      "invalid",
      "Use 1–20 checkpoints and a local envelope size between 1 KiB and 50 MiB.",
    );
  function readStoredText() {
    try {
      return storage.getItem(namespace);
    } catch {
      return throwPersistenceError(
        "unavailable",
        "Browser storage is unavailable. Keep the workbook open and export your unsaved changes.",
      );
    }
  }
  function loadStoredEnvelope() {
    const storedText = readStoredText();
    if (storedText === null) return null;
    if (storedText.length > maxBytes || measureUtf8Bytes(storedText) > maxBytes)
      return throwPersistenceError(
        "limit",
        "The saved local workbook exceeds this store's size limit. Preserve it for export or increase the configured limit.",
      );
    let rawEnvelope;
    try {
      rawEnvelope = JSON.parse(storedText);
    } catch {
      return throwPersistenceError(
        "corrupt",
        "The saved workbook cannot be read. Its original storage has been preserved; export in-memory work before choosing another namespace.",
      );
    }
    if (!isPlainRecord(rawEnvelope))
      return throwPersistenceError(
        "corrupt",
        "Invalid local workbook envelope.",
      );
    let envelope,
      recovered = false,
      warning;
    if (rawEnvelope.schemaVersion === 1) {
      let legacySnapshot;
      try {
        legacySnapshot = normalizeWorkbookSnapshot(rawEnvelope);
      } catch {
        return throwPersistenceError(
          "corrupt",
          "The legacy saved workbook is invalid; original storage has been preserved.",
        );
      }
      const workbookSnapshot = createWorkbookSnapshot(
        legacySnapshot,
        normalizeSavedSelection(legacySnapshot),
      );
      ((envelope = {
        format: localWorkbookNamespace,
        version: 1,
        generation: 0,
        current: workbookSnapshot,
        checkpoints: [],
      }),
        (warning =
          "Legacy workbook loaded. The next successful save will add local recovery and version metadata."));
    } else {
      if (
        rawEnvelope.format !== localWorkbookNamespace ||
        rawEnvelope.version !== 1 ||
        !Number.isSafeInteger(rawEnvelope.generation) ||
        Number(rawEnvelope.generation) < 1 ||
        !Array.isArray(rawEnvelope.checkpoints) ||
        rawEnvelope.checkpoints.length > maxLocalCheckpoints
      )
        return throwPersistenceError(
          "corrupt",
          "Unsupported or corrupt local workbook envelope; stored data was preserved.",
        );
      let currentSnapshot, previous;
      if (rawEnvelope.previous !== undefined)
        try {
          previous = parseWorkbookSnapshot(rawEnvelope.previous);
        } catch {
          warning =
            "The previous recovery snapshot is invalid; the current workbook remains available.";
        }
      try {
        currentSnapshot = parseWorkbookSnapshot(rawEnvelope.current);
      } catch {
        if (!previous)
          return throwPersistenceError(
            "corrupt",
            "Both saved workbook snapshots are invalid. Original storage has been preserved for recovery.",
          );
        ((currentSnapshot = previous),
          (recovered = true),
          (warning =
            "Recovered the previous complete save because the latest snapshot is invalid. Review it before saving; the original storage is unchanged."));
      }
      const checkpointIds = new Set(),
        checkpoints = [];
      for (const checkpoint of rawEnvelope.checkpoints) {
        if (
          !isPlainRecord(checkpoint) ||
          typeof checkpoint.id != "string" ||
          !checkpoint.id ||
          checkpoint.id.length > 100 ||
          checkpointIds.has(checkpoint.id) ||
          typeof checkpoint.label != "string" ||
          !checkpoint.label.trim() ||
          checkpoint.label.length > 120 ||
          typeof checkpoint.createdAt != "string" ||
          !Number.isFinite(Date.parse(checkpoint.createdAt))
        )
          return throwPersistenceError(
            "corrupt",
            "Saved checkpoint metadata is invalid; original storage was preserved.",
          );
        let snapshot;
        try {
          snapshot = parseWorkbookSnapshot(checkpoint.snapshot);
        } catch {
          return throwPersistenceError(
            "corrupt",
            "A saved checkpoint is corrupt; original storage was preserved.",
          );
        }
        if (snapshot.workbook.id !== currentSnapshot.workbook.id)
          return throwPersistenceError(
            "corrupt",
            "A checkpoint belongs to another workbook; original storage was preserved.",
          );
        (checkpointIds.add(checkpoint.id),
          checkpoints.push({
            id: checkpoint.id,
            label: checkpoint.label,
            createdAt: checkpoint.createdAt,
            snapshot: snapshot,
          }));
      }
      envelope = {
        format: localWorkbookNamespace,
        version: 1,
        generation: Number(rawEnvelope.generation),
        current: currentSnapshot,
        ...(previous ? { previous: previous } : {}),
        checkpoints: checkpoints,
      };
    }
    const { workbook, view } = envelope.current;
    return {
      envelope: envelope,
      result: {
        workbook: cloneStructuredValue(workbook),
        view: cloneStructuredValue(view),
        guard: {
          namespace: namespace,
          generation: envelope.generation,
          workbookId: workbook.id,
          revision: workbook.revision,
          digest: computeTextDigest(storedText),
        },
        recovered: recovered,
        ...(warning ? { warning: warning } : {}),
        checkpoints: envelope.checkpoints.map(createCheckpointRecord),
      },
    };
  }
  function checkExpectedGuard(expectedGuard) {
    const storedState = loadStoredEnvelope();
    return expectedGuard === null
      ? storedState
        ? throwPersistenceError(
            "conflict",
            "This local namespace already contains a workbook. Reload it before saving; unsaved work remains in memory.",
            storedState.result.guard,
          )
        : null
      : !storedState ||
          expectedGuard.namespace !== namespace ||
          stringifyCanonicalJson(expectedGuard) !==
            stringifyCanonicalJson(storedState.result.guard)
        ? throwPersistenceError(
            "conflict",
            "Another workbook window saved a different version. Reload or export your unsaved changes before resolving the conflict.",
            storedState?.result.guard,
          )
        : storedState;
  }
  function writeEnvelope(envelope, expectedGuard) {
    const serializedText = JSON.stringify(envelope);
    if (
      serializedText.length > maxBytes ||
      measureUtf8Bytes(serializedText) > maxBytes
    )
      return throwPersistenceError(
        "limit",
        "Local version storage exceeds its configured size limit. Export the workbook or use a larger host store; the previous save is unchanged.",
      );
    checkExpectedGuard(expectedGuard);
    try {
      storage.setItem(namespace, serializedText);
    } catch (error) {
      const isQuotaError =
        error instanceof Error &&
        /quota|space|full/i.test(`${error.name} ${error.message}`);
      return throwPersistenceError(
        isQuotaError ? "quota" : "unavailable",
        "Local save failed. The previous complete save is unchanged; keep this workbook open and export your unsaved changes.",
      );
    }
    return {
      workbook: cloneStructuredValue(envelope.current.workbook),
      view: cloneStructuredValue(envelope.current.view),
      guard: {
        namespace: namespace,
        generation: envelope.generation,
        workbookId: envelope.current.workbook.id,
        revision: envelope.current.workbook.revision,
        digest: computeTextDigest(serializedText),
      },
      recovered: false,
      checkpoints: envelope.checkpoints.map(createCheckpointRecord),
    };
  }
  function nextGeneration(storedState) {
    const generation = (storedState?.envelope.generation ?? 0) + 1;
    return Number.isSafeInteger(generation)
      ? generation
      : throwPersistenceError(
          "limit",
          "Local version generation limit reached.",
        );
  }
  function findCheckpoint(storedState, checkpointId) {
    const foundCheckpoint = storedState?.envelope.checkpoints.find(
      (checkpoint) => checkpoint.id === checkpointId,
    );
    return (
      foundCheckpoint ||
      throwPersistenceError(
        "invalid",
        "This checkpoint is unavailable or has expired from the bounded local history.",
      )
    );
  }
  return {
    load: () => loadStoredEnvelope()?.result ?? null,
    save(workbook, view, expectedGuard, m = {}) {
      const storedState = checkExpectedGuard(expectedGuard),
        workbookSnapshot = createWorkbookSnapshot(workbook, view),
        isSameWorkbook =
          !storedState ||
          storedState.result.workbook.id === workbookSnapshot.workbook.id;
      return !isSameWorkbook && !m.allowIdentityChange
        ? throwPersistenceError(
            "conflict",
            "Opening another workbook requires an explicit namespace replacement or a separate namespace.",
            storedState?.result.guard,
          )
        : isSameWorkbook &&
            storedState &&
            workbookSnapshot.workbook.revision <
              storedState.result.workbook.revision
          ? throwPersistenceError(
              "conflict",
              "An older workbook revision cannot overwrite a newer saved revision. Restore a checkpoint as a new revision instead.",
              storedState.result.guard,
            )
          : isSameWorkbook &&
              storedState &&
              workbookSnapshot.workbook.revision ===
                storedState.result.workbook.revision &&
              stringifyCanonicalJson(workbookSnapshot.workbook) !==
                stringifyCanonicalJson(storedState.result.workbook)
            ? throwPersistenceError(
                "conflict",
                "Workbook content changed without a new revision. Commit the edit before saving.",
                storedState.result.guard,
              )
            : writeEnvelope(
                {
                  format: localWorkbookNamespace,
                  version: 1,
                  generation: nextGeneration(storedState),
                  current: workbookSnapshot,
                  ...(storedState
                    ? { previous: storedState.envelope.current }
                    : {}),
                  checkpoints: isSameWorkbook
                    ? (storedState?.envelope.checkpoints ?? [])
                    : [],
                },
                expectedGuard,
              );
    },
    checkpoint(label, expectedGuard) {
      if (typeof label != "string" || !label.trim() || label.length > 120)
        return throwPersistenceError(
          "invalid",
          "Checkpoint label must contain 1–120 characters.",
        );
      const storedState = checkExpectedGuard(expectedGuard),
        newCheckpoint = {
          id: crypto.randomUUID(),
          label: label.trim(),
          createdAt: new Date().toISOString(),
          snapshot: cloneStructuredValue(storedState.envelope.current),
        },
        checkpoints = [
          ...storedState.envelope.checkpoints,
          newCheckpoint,
        ].slice(-maxCheckpoints),
        writeResult = writeEnvelope(
          {
            ...storedState.envelope,
            generation: nextGeneration(storedState),
            checkpoints: checkpoints,
          },
          expectedGuard,
        );
      return (
        storedState.envelope.checkpoints.length >= maxCheckpoints &&
          (writeResult.warning = `The oldest checkpoint expired. This local store retains the newest ${maxCheckpoints} checkpoints.`),
        writeResult
      );
    },
    listCheckpoints: () => loadStoredEnvelope()?.result.checkpoints ?? [],
    compareCheckpoint: (checkpointId, workbook) =>
      compareWorkbookVersions(
        findCheckpoint(loadStoredEnvelope(), checkpointId).snapshot.workbook,
        workbook,
      ),
    restoreCheckpoint(
      checkpointId,
      workbook,
      expectedGuard,
      actor = "Auditor",
    ) {
      const storedState = checkExpectedGuard(expectedGuard),
        workbookSnapshot = normalizeWorkbookSnapshot(workbook);
      if (
        workbookSnapshot.id !== storedState.result.workbook.id ||
        workbookSnapshot.revision !== storedState.result.workbook.revision ||
        stringifyCanonicalJson(workbookSnapshot) !==
          stringifyCanonicalJson(storedState.result.workbook)
      )
        return throwPersistenceError(
          "conflict",
          "Save or export your current unsaved changes before restoring a checkpoint.",
          storedState.result.guard,
        );
      if (
        workbookSnapshot.structureProtection?.enabled ||
        workbookSnapshot.sheets.some((sheet) => sheet.protection?.enabled)
      )
        return throwPersistenceError(
          "invalid",
          "Unlock the workbook structure and protected worksheets before restoring a version.",
        );
      if (typeof actor != "string" || !actor.trim() || actor.length > 100)
        return throwPersistenceError(
          "invalid",
          "A restore actor must contain 1–100 characters.",
        );
      const checkpoint = findCheckpoint(storedState, checkpointId),
        comparison = compareWorkbookVersions(
          workbookSnapshot,
          checkpoint.snapshot.workbook,
        ),
        timestamp = new Date().toISOString();
      let restoredWorkbook = cloneStructuredValue(checkpoint.snapshot.workbook);
      if (
        ((restoredWorkbook.revision = workbookSnapshot.revision + 1),
        !Number.isSafeInteger(restoredWorkbook.revision) ||
          restoredWorkbook.revision >= Number.MAX_SAFE_INTEGER)
      )
        return throwPersistenceError(
          "limit",
          "Workbook revision limit reached.",
        );
      (workbookSnapshot.metadata &&
        (restoredWorkbook.metadata = {
          ...(restoredWorkbook.metadata ?? workbookSnapshot.metadata),
          createdAt: workbookSnapshot.metadata.createdAt,
          modifiedAt: [
            workbookSnapshot.metadata.modifiedAt,
            restoredWorkbook.metadata?.modifiedAt ?? timestamp,
            timestamp,
          ]
            .sort()
            .at(-1),
        }),
        (restoredWorkbook = touchWorkbookModifiedAt(
          restoredWorkbook,
          timestamp,
        )));
      const chunkCount = Math.max(
        1,
        Math.ceil(comparison.cellChanges.length / 500),
      );
      return (
        (restoredWorkbook.events = [
          ...workbookSnapshot.events,
          ...Array.from({ length: chunkCount }, (I, chunkIndex) => ({
            id: crypto.randomUUID(),
            timestamp: timestamp,
            actor: actor,
            action: `Restored checkpoint “${checkpoint.label}” from revision ${checkpoint.snapshot.workbook.revision}${chunkCount > 1 ? ` (${chunkIndex + 1}/${chunkCount})` : ""}`,
            changes: comparison.cellChanges
              .slice(chunkIndex * 500, (chunkIndex + 1) * 500)
              .map(({ attribution, ...cellChange }) => cellChange),
          })),
        ]),
        (restoredWorkbook = normalizeWorkbookSnapshot(restoredWorkbook)),
        writeEnvelope(
          {
            ...storedState.envelope,
            generation: nextGeneration(storedState),
            previous: storedState.envelope.current,
            current: createWorkbookSnapshot(
              restoredWorkbook,
              checkpoint.snapshot.view,
              timestamp,
            ),
          },
          expectedGuard,
        )
      );
    },
  };
}
function createWorkbookSession(options = {}) {
  let inMemoryText = null,
    sessionState = null;
  const namespace = `session-${crypto.randomUUID()}`,
    localWorkbookStore = createLocalWorkbookStore(
      {
        getItem: () => inMemoryText,
        setItem: (f, serializedText) => {
          inMemoryText = serializedText;
        },
      },
      namespace,
      options,
    );
  let savedGuard = null;
  const checkpointSizes = new Map(),
    maxBytes = options.maxBytes ?? 10 * 1024 * 1024,
    isSameJson = (leftValue, rightValue) =>
      leftValue === rightValue ||
      JSON.stringify(leftValue) === JSON.stringify(rightValue);
  function assertSessionGuard(expectedGuard) {
    if (!isSameJson(expectedGuard, sessionState?.guard ?? null))
      throw new WorkbookRevisionConflictError(
        "conflict",
        "The saved session changed. Reload before replacing it.",
        sessionState?.guard,
      );
  }
  function commitSession(sessionInput) {
    const generation = (sessionState?.guard.generation ?? 0) + 1;
    if (!Number.isSafeInteger(generation))
      throw new WorkbookRevisionConflictError(
        "limit",
        "Session generation limit reached.",
      );
    return (
      (sessionState = deepFreezeValue({
        ...sessionInput,
        workbook: buildCanonicalWorkbookSnapshot(sessionInput.workbook),
        view: structuredClone(sessionInput.view),
        guard: {
          namespace: namespace,
          generation: generation,
          workbookId: sessionInput.workbook.id,
          revision: sessionInput.workbook.revision,
          digest: crypto.randomUUID(),
        },
      })),
      sessionState
    );
  }
  function persistCurrentSession() {
    if (!sessionState)
      throw new WorkbookRevisionConflictError(
        "invalid",
        "Save this session before creating a version.",
      );
    const savedResult = localWorkbookStore.save(
      sessionState.workbook,
      sessionState.view,
      savedGuard,
      {
        allowIdentityChange: true,
      },
    );
    return ((savedGuard = savedResult.guard), savedResult);
  }
  return {
    load: () => sessionState,
    save(workbook, workbookView, guard, v = {}) {
      assertSessionGuard(guard);
      const canonicalWorkbookSnapshot =
          buildCanonicalWorkbookSnapshot(workbook),
        sheet = canonicalWorkbookSnapshot.sheets.find(
          (sheet) => sheet.id === workbookView.activeSheetId && !sheet.hidden,
        );
      if (!sheet || workbookView.selection.sheetId !== sheet.id)
        throw new WorkbookRevisionConflictError(
          "invalid",
          "Saved selection must belong to the active visible worksheet.",
        );
      const isSameWorkbook =
        !sessionState ||
        sessionState.workbook.id === canonicalWorkbookSnapshot.id;
      if (!isSameWorkbook && !v.allowIdentityChange)
        throw new WorkbookRevisionConflictError(
          "conflict",
          "Opening another workbook requires an explicit session replacement.",
          sessionState?.guard,
        );
      if (
        isSameWorkbook &&
        sessionState &&
        (canonicalWorkbookSnapshot.revision < sessionState.workbook.revision ||
          (canonicalWorkbookSnapshot.revision ===
            sessionState.workbook.revision &&
            !isSameJson(canonicalWorkbookSnapshot, sessionState.workbook)))
      )
        throw new WorkbookRevisionConflictError(
          "conflict",
          "Commit a new revision before replacing saved content.",
          sessionState.guard,
        );
      const view = validateSavedSelection(
          canonicalWorkbookSnapshot,
          workbookView,
        ),
        toSnapshotRecord = (workbook, view) => ({
          workbook: workbook,
          view: view,
          savedAt: "1970-01-01T00:00:00.000Z",
          digest: "0".repeat(40),
        }),
        retainedCheckpoints = isSameWorkbook
          ? (sessionState?.checkpoints ?? [])
          : [],
        checkpointBytes = retainedCheckpoints.reduce(
          (total, checkpoint) =>
            total + (checkpointSizes.get(checkpoint.id) ?? 0),
          0,
        );
      if (
        measureSerializedBytes({
          format: "ledgerglass-local-workbook",
          version: 1,
          generation: (sessionState?.guard.generation ?? 0) + 1,
          current: toSnapshotRecord(canonicalWorkbookSnapshot, view),
          ...(sessionState
            ? {
                previous: toSnapshotRecord(
                  sessionState.workbook,
                  sessionState.view,
                ),
              }
            : {}),
          checkpoints: [],
        }) +
          checkpointBytes +
          Math.max(0, retainedCheckpoints.length - 1) >
        maxBytes
      )
        throw new WorkbookRevisionConflictError(
          "limit",
          "Session version storage exceeds its configured size limit. Export the workbook or use a larger host store; the previous save is unchanged.",
        );
      return (
        isSameWorkbook || checkpointSizes.clear(),
        commitSession({
          workbook: canonicalWorkbookSnapshot,
          view: view,
          guard: guard,
          recovered: false,
          checkpoints: isSameWorkbook ? (sessionState?.checkpoints ?? []) : [],
        })
      );
    },
    checkpoint(label, expectedGuard) {
      assertSessionGuard(expectedGuard);
      const persistedResult = persistCurrentSession(),
        checkpointResult = localWorkbookStore.checkpoint(
          label,
          persistedResult.guard,
        );
      savedGuard = checkpointResult.guard;
      const checkpointIds = new Set(
        checkpointResult.checkpoints.map((checkpoint) => checkpoint.id),
      );
      for (const cachedCheckpointId of checkpointSizes.keys())
        checkpointIds.has(cachedCheckpointId) ||
          checkpointSizes.delete(cachedCheckpointId);
      for (const checkpointId of checkpointIds)
        checkpointSizes.has(checkpointId) ||
          checkpointSizes.set(
            checkpointId,
            measureSerializedBytes({
              id: checkpointId,
              label: checkpointResult.checkpoints.find(
                (checkpoint) => checkpoint.id === checkpointId,
              ).label,
              createdAt: "1970-01-01T00:00:00.000Z",
              snapshot: {
                workbook: checkpointResult.workbook,
                view: checkpointResult.view,
                savedAt: "1970-01-01T00:00:00.000Z",
                digest: "0".repeat(40),
              },
            }),
          );
      return commitSession(checkpointResult);
    },
    listCheckpoints: () => structuredClone(sessionState?.checkpoints ?? []),
    compareCheckpoint: (checkpointId, workbook) =>
      localWorkbookStore.compareCheckpoint(checkpointId, workbook),
    restoreCheckpoint(checkpointId, workbook, expectedGuard, actor) {
      if (
        (assertSessionGuard(expectedGuard),
        !sessionState || !isSameJson(workbook, sessionState.workbook))
      )
        throw new WorkbookRevisionConflictError(
          "conflict",
          "Save the current revision before restoring a checkpoint.",
          sessionState?.guard,
        );
      const persistedResult = persistCurrentSession(),
        restoredResult = localWorkbookStore.restoreCheckpoint(
          checkpointId,
          workbook,
          persistedResult.guard,
          actor,
        );
      return (
        (savedGuard = restoredResult.guard),
        commitSession(restoredResult)
      );
    },
  };
}
function validateSavedSelection(canonicalWorkbookSnapshot, workbookView) {
  const sheet = canonicalWorkbookSnapshot.sheets.find(
      (sheet) => sheet.id === workbookView.activeSheetId,
    ),
    rangeAddress = parseRangeAddress(workbookView.selection.range);
  if (
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount
  )
    throw new WorkbookRevisionConflictError(
      "invalid",
      "Saved selection exceeds worksheet dimensions.",
    );
  const firstCellAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    lastCellAddress = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  return {
    activeSheetId: sheet.id,
    selection: {
      sheetId: sheet.id,
      range:
        firstCellAddress === lastCellAddress
          ? firstCellAddress
          : `${firstCellAddress}:${lastCellAddress}`,
    },
  };
}
const buildCompositeKey = (sheetId, address) =>
    JSON.stringify([sheetId, address]),
  isFormulaInputValue = (cell) =>
    !!cell &&
    !cell.literal &&
    typeof cell.raw == "string" &&
    cell.raw.startsWith("="),
  bareSheetNamePattern = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/,
  sheetNamePattern = String.raw`(?:'(?:[^']|'')+'|[A-Za-z_\\][A-Za-z0-9_.\\]*)`,
  cellRangePattern = String.raw`(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?:\s*:\s*\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}\s*:\s*\$?[A-Za-z]{1,3}|\$?[1-9]\d*\s*:\s*\$?[1-9]\d*)`,
  singleRangeReferencePattern = new RegExp(
    `^(${sheetNamePattern}\\s*!\\s*)?(${cellRangePattern})(#)?`,
  ),
  sheetRangeReferencePattern = new RegExp(
    `^${sheetNamePattern}\\s*:\\s*${sheetNamePattern}\\s*!\\s*${cellRangePattern}`,
  ),
  knownFunctionNameSet = new Set(
    builtinFunctionNames.map((builtinFunctionName) =>
      builtinFunctionName.toUpperCase(),
    ),
  ),
  requireIntegerInRange = (
    providedValue,
    defaultValue,
    maximum,
    fieldLabel,
  ) => {
    const resolvedValue = providedValue ?? defaultValue;
    if (
      !Number.isInteger(resolvedValue) ||
      resolvedValue < 1 ||
      resolvedValue > maximum
    )
      throw Error(`${fieldLabel} must be an integer from 1 to ${maximum}.`);
    return resolvedValue;
  };
function parseCellReference(address) {
  const normalizedAddress = address.replace(/[\s$]/g, "").toUpperCase(),
    columnRangeMatch = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(normalizedAddress),
    rowRangeMatch = /^([1-9]\d*):([1-9]\d*)$/.exec(normalizedAddress),
    rangeAddress = parseRangeAddress(
      columnRangeMatch
        ? `${columnRangeMatch[1]}1:${columnRangeMatch[2]}${workbookLimits.maxRows}`
        : rowRangeMatch
          ? `A${rowRangeMatch[1]}:XFD${rowRangeMatch[2]}`
          : normalizedAddress,
    ),
    cellAddress = formatCellAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
    lastCellAddress = formatCellAddress(
      rangeAddress.lastRow,
      rangeAddress.lastCol,
    );
  return cellAddress === lastCellAddress
    ? cellAddress
    : `${cellAddress}:${lastCellAddress}`;
}

export {
  FormatConversionDialog,
  createLocalWorkbookStore,
  requireIntegerInRange,
  singleRangeReferencePattern,
  WorkbookRevisionConflictError,
  buildCompositeKey,
  parseCellReference,
  bareSheetNamePattern,
  knownFunctionNameSet,
  isFormulaInputValue,
  compareWorkbookVersions,
  sheetRangeReferencePattern,
  createWorkbookSession,
};
export { cloneStructuredValue, isPlainRecord };
