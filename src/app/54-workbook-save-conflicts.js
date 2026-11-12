import { clampNumber } from "./53-find-replace-go-to.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  string,
  strictObject,
  enum as _enum,
  discriminatedUnion,
  literal,
} from "zod";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  assertWorkbookStructureEditable,
  isSameContentHash,
  createWorksheetRecord,
  isExternalSourceError,
  externalSourceFormatVersion,
  ExternalSourceError,
  validateExternalSourcePayload,
  parseRangeAddress,
  externalSourceLimits,
  parseActiveSourceApproval,
  buildCanonicalWorkbookSnapshot,
  normalizeWorkbookInput,
  columnIndexToLetters,
  computeSourceControlTotals,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
} from "./04-cell-input-paste-and-fill.js";
import {
  createLocalWorkbookStore,
  WorkbookRevisionConflictError,
  compareWorkbookVersions,
} from "./11-local-workbook-store-and-versions.js";
import { inspectOutputNameSafety } from "./12-consolidation-and-dependency-inspection.js";
import {
  resolveAxisScale,
  physicalToLogicalPixels,
} from "./53-find-replace-go-to.js";

function computeVirtualizedRange(
  axis,
  scrollOffset,
  t,
  zoomScale,
  overscan = 5,
) {
  const axisScale = resolveAxisScale(axis.total, t, zoomScale),
    clampedOffset = clampNumber(scrollOffset, 0, axisScale.physicalMax),
    logicalOffset = physicalToLogicalPixels(
      clampedOffset,
      axis.total,
      t,
      zoomScale,
    ),
    first = axis.count
      ? Math.max(0, axis.indexAt(logicalOffset) - overscan)
      : 0,
    last = axis.count
      ? Math.min(
          axis.count,
          axis.indexAt(logicalOffset + axisScale.viewport / zoomScale) +
            overscan +
            1,
        )
      : 0,
    leading = Math.max(
      0,
      clampedOffset / zoomScale + axis.offset(first) - logicalOffset,
    ),
    trailing = Math.max(
      0,
      axisScale.physicalPixels / zoomScale -
        leading -
        (axis.offset(last) - axis.offset(first)),
    );
  return {
    first: first,
    last: last,
    leading: leading,
    trailing: trailing,
    logicalOffset: logicalOffset,
    total: axisScale.physicalPixels / zoomScale,
  };
}
function workbookLockName(embeddedDemoKey) {
  return `ledgerglass:workbook:${embeddedDemoKey}`;
}
function getBrowserLockManager() {
  try {
    return typeof navigator < "u" ? (navigator.locks ?? null) : null;
  } catch {
    return null;
  }
}
function cloneSaveInputs(saveInputs) {
  try {
    return structuredClone(saveInputs);
  } catch {
    throw new WorkbookRevisionConflictError(
      "invalid",
      "The workbook save inputs could not be captured. Keep your unsaved work open and correct the invalid input before retrying.",
    );
  }
}
function saveWorkbookWithLock(e, workbookKey, options = {}) {
  const localWorkbookStore = createLocalWorkbookStore(e, workbookKey, {
      maxBytes: options.maxBytes,
      maxCheckpoints: options.maxCheckpoints,
    }),
    lockManager =
      options.locks === undefined ? getBrowserLockManager() : options.locks,
    lockTimeoutMs = options.lockTimeoutMs ?? 5e3,
    lockName = workbookLockName(workbookKey);
  if (
    !Number.isInteger(lockTimeoutMs) ||
    lockTimeoutMs < 1 ||
    lockTimeoutMs > 6e4
  )
    throw new WorkbookRevisionConflictError(
      "invalid",
      "The browser save lock timeout must be between 1 and 60,000 milliseconds.",
    );
  function runWithLock(operation, signal) {
    return new Promise((resolve, reject) => {
      const createUnavailableError = (message) =>
          new WorkbookRevisionConflictError("unavailable", message),
        abortController = new AbortController(),
        deadline = performance.now() + lockTimeoutMs;
      let isPending = true,
        hasLockGranted = false,
        hasOperationFailed = false,
        operationError,
        abortReason,
        timeoutId;
      const cleanup = () => {
          (timeoutId !== undefined && clearTimeout(timeoutId),
            signal?.removeEventListener("abort", handleAbort));
        },
        abortWith = (reason) => {
          !isPending ||
            hasLockGranted ||
            ((isPending = false),
            (abortReason = reason),
            cleanup(),
            abortController.abort(reason),
            reject(reason));
        },
        handleAbort = () =>
          abortWith(
            createUnavailableError(
              "The queued browser save was cancelled. No changes were written; your unsaved work remains in memory.",
            ),
          ),
        createTimeoutError = () =>
          createUnavailableError(
            "Timed out waiting for another workbook window to finish saving. No changes were written; retry or export your unsaved work.",
          ),
        handleFailure = (error) => {
          ((isPending = false),
            cleanup(),
            reject(
              hasOperationFailed
                ? operationError
                : error instanceof WorkbookRevisionConflictError
                  ? error
                  : createUnavailableError(
                      "Browser save locking is unavailable. No changes were written; keep the workbook open and export your unsaved work.",
                    ),
            ));
        };
      if (signal?.aborted) {
        handleAbort();
        return;
      }
      try {
        if (!lockManager || typeof lockManager.request != "function") {
          handleFailure(
            createUnavailableError(
              "Browser save locking is unavailable. No changes were written; keep the workbook open and export your unsaved work.",
            ),
          );
          return;
        }
        (signal?.addEventListener("abort", handleAbort, { once: true }),
          (timeoutId = setTimeout(
            () => abortWith(createTimeoutError()),
            lockTimeoutMs,
          )));
        const lockRequest = lockManager.request(
          lockName,
          { mode: "exclusive", signal: abortController.signal },
          (lock) => {
            if (!isPending || abortController.signal.aborted)
              throw (
                abortReason ??
                createUnavailableError("The queued browser save was cancelled.")
              );
            if (performance.now() >= deadline) {
              const timeoutError = createTimeoutError();
              throw (abortWith(timeoutError), timeoutError);
            }
            if (!lock || lock.name !== lockName || lock.mode !== "exclusive")
              throw createUnavailableError(
                "The browser did not grant the required exclusive save lock.",
              );
            ((hasLockGranted = true), cleanup());
            try {
              return operation();
            } catch (error) {
              throw (
                (hasOperationFailed = true),
                (operationError = error),
                error
              );
            }
          },
        );
        Promise.resolve(lockRequest).then((result) => {
          ((isPending = false), cleanup(), resolve(result));
        }, handleFailure);
      } catch (error) {
        handleFailure(error);
      }
    });
  }
  return {
    coordination: "web-locks",
    load: localWorkbookStore.load,
    listCheckpoints: localWorkbookStore.listCheckpoints,
    compareCheckpoint: localWorkbookStore.compareCheckpoint,
    async save(workbook, view, expected, saveOptions = {}) {
      const clonedInputs = cloneSaveInputs({
          workbook: workbook,
          view: view,
          expected: expected,
          options: { allowIdentityChange: saveOptions.allowIdentityChange },
        }),
        signal = saveOptions.signal;
      return runWithLock(
        () =>
          localWorkbookStore.save(
            clonedInputs.workbook,
            clonedInputs.view,
            clonedInputs.expected,
            clonedInputs.options,
          ),
        signal,
      );
    },
    async checkpoint(label, expected, checkpointOptions = {}) {
      const clonedInputs = cloneSaveInputs({
          label: label,
          expected: expected,
        }),
        signal = checkpointOptions.signal;
      return runWithLock(
        () =>
          localWorkbookStore.checkpoint(
            clonedInputs.label,
            clonedInputs.expected,
          ),
        signal,
      );
    },
    async restoreCheckpoint(
      checkpointId,
      workbook,
      expected,
      actor,
      restoreOptions = {},
    ) {
      const clonedInputs = cloneSaveInputs({
          id: checkpointId,
          workbook: workbook,
          expected: expected,
          actor: actor,
        }),
        signal = restoreOptions.signal;
      return runWithLock(
        () =>
          localWorkbookStore.restoreCheckpoint(
            clonedInputs.id,
            clonedInputs.workbook,
            clonedInputs.expected,
            clonedInputs.actor,
          ),
        signal,
      );
    },
  };
}
const workbookSectionLabels = {
    name: "Name",
    metadata: "Workbook properties",
    createdAt: "Created",
    modifiedAt: "Last modified",
    engagement: "Engagement",
    period: "Period",
    template: "Template",
    sourceRevision: "Original revision",
    sourceWorkbookId: "Original workbook",
    definedNames: "Named ranges",
    formulaWatches: "Watched formulas",
    styleCatalog: "Styles and theme",
    reviewThreads: "Review comments",
    structureProtection: "Workbook protection",
    protection: "Protection",
    enabled: "Enabled",
    hidden: "Hidden",
    veryHidden: "Hidden from worksheet menu",
    tabColor: "Tab color",
    rowCount: "Worksheet rows",
    colCount: "Worksheet columns",
    columnWidths: "Column widths",
    rowHeights: "Row heights",
    frozenRows: "Frozen rows",
    frozenCols: "Frozen columns",
    hiddenRows: "Hidden rows",
    hiddenCols: "Hidden columns",
    merges: "Merged cells",
    drawings: "Pictures, shapes and charts",
    validations: "Cell validation rules",
    conditionalFormats: "Conditional formatting",
    pageLayout: "Print layout",
    tables: "Tables",
    pivots: "Pivot tables",
    queries: "Queries",
    literal: "Entered as text",
    numFmt: "Number format",
    format: "Number format category",
    style: "Formatting",
    note: "Note",
    tickmark: "Review mark",
    bold: "Bold",
    italic: "Italic",
    underline: "Underline",
    fontFamily: "Font",
    fontSize: "Font size",
    wrap: "Wrap text",
    border: "Borders",
    fill: "Fill color",
    color: "Text color",
    align: "Alignment",
  },
  humanizeFieldName = (fieldName) =>
    workbookSectionLabels[fieldName] ??
    fieldName
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .replace(/[_-]/g, " ")
      .replace(/^./, (character) => character.toUpperCase()),
  wp = (value) =>
    value !== null && typeof value == "object" && !Array.isArray(value),
  stableSerialize = (value) =>
    Array.isArray(value)
      ? `[${value.map(stableSerialize).join(",")}]`
      : wp(value)
        ? `{${Object.keys(value)
            .sort()
            .map(
              (key) => `${JSON.stringify(key)}:${stableSerialize(value[key])}`,
            )
            .join(",")}}`
        : (JSON.stringify(value) ?? "undefined"),
  truncateDisplayText = (text, maxLength = 220) =>
    text.length > maxLength ? `${text.slice(0, maxLength)}…` : text;
function summarizeConflictValue(value, depth = 0) {
  if (value === undefined) return "Not set";
  if (value === null || value === "") return "Blank";
  if (typeof value == "boolean") return value ? "Yes" : "No";
  if (typeof value == "string")
    return value.startsWith("data:image/")
      ? "Embedded image"
      : /^\d{4}-\d{2}-\d{2}T/.test(value) && Number.isFinite(Date.parse(value))
        ? new Date(value).toLocaleString(undefined, {
            year: "numeric",
            month: "short",
            day: "numeric",
            hour: "numeric",
            minute: "2-digit",
            second: "2-digit",
            fractionalSecondDigits: 3,
            timeZoneName: "short",
          })
        : truncateDisplayText(value);
  if (Array.isArray(value))
    return value.length
      ? depth >= 2
        ? `${value.length} items`
        : truncateDisplayText(
            value
              .slice(0, 4)
              .map((item) => summarizeConflictValue(item, depth + 1))
              .join("; ") +
              (value.length > 4 ? `; ${value.length - 4} more items` : ""),
          )
      : "None";
  if (wp(value)) {
    const entries = Object.entries(value);
    return entries.length
      ? depth >= 2
        ? `${entries.length} settings`
        : truncateDisplayText(
            entries
              .slice(0, 4)
              .map(
                ([entryKey, entryValue]) =>
                  `${humanizeFieldName(entryKey)}: ${summarizeConflictValue(entryValue, depth + 1)}`,
              )
              .join(" · ") +
              (entries.length > 4
                ? ` · ${entries.length - 4} more settings`
                : ""),
          )
      : "No settings";
  }
  return String(value);
}
function describeCellContent(cell) {
  return cell
    ? cell.raw === null || cell.raw === ""
      ? "Blank"
      : typeof cell.raw == "boolean"
        ? cell.raw
          ? "TRUE"
          : "FALSE"
        : typeof cell.raw == "string"
          ? (cell.literal
              ? "Text: "
              : cell.raw.startsWith("=")
                ? "Formula: "
                : "") + truncateDisplayText(cell.raw)
          : String(cell.raw)
    : "Empty cell";
}
function collectConflictChanges(conflict) {
  const changes = [],
    collectChange = (path, before, after, depth) => {
      if (stableSerialize(before) !== stableSerialize(after)) {
        if (
          depth < 3 &&
          (wp(before) || wp(after)) &&
          (before === undefined || wp(before)) &&
          (after === undefined || wp(after))
        ) {
          const beforeObject = wp(before) ? before : {},
            afterObject = wp(after) ? after : {},
            allKeys = new Set([
              ...Object.keys(beforeObject),
              ...Object.keys(afterObject),
            ]);
          if (allKeys.size) {
            for (const key of allKeys)
              collectChange(
                [...path, key],
                beforeObject[key],
                afterObject[key],
                depth + 1,
              );
            return;
          }
        }
        changes.push({ path: path, before: before, after: after });
      }
    };
  return (
    collectChange(conflict.path.split("."), conflict.before, conflict.after, 0),
    changes
  );
}
function WorkbookSaveConflictDialog(props) {
  const { text, language } = useWorkbookMessages(),
    [savedWorkbook, setSavedWorkbook] = ne.useState(null),
    [comparison, setComparison] = ne.useState(null),
    [loadError, setLoadError] = ne.useState(""),
    [checked, setChecked] = ne.useState(false),
    [open, setOpen] = ne.useState(false),
    dialogRef = ne.useRef(null),
    onClick = () => {
      (setChecked(false), setOpen(false));
      try {
        const loaded = props.loadSaved();
        (setSavedWorkbook(loaded),
          setComparison(
            loaded?.workbook.id === props.workbook.id
              ? compareWorkbookVersions(loaded.workbook, props.workbook)
              : null,
          ),
          setLoadError(
            loaded
              ? ""
              : "The saved workbook is no longer available. Export your local changes before continuing.",
          ));
      } catch (error) {
        (setSavedWorkbook(null),
          setComparison(null),
          setLoadError(
            error instanceof Error
              ? error.message
              : "The saved workbook cannot be read.",
          ));
      }
    };
  (ne.useEffect(onClick, [props.workbook]),
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current?.querySelector("button")?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []));
  const sheetNameFor = (sheetId) =>
      props.workbook.sheets.find((sheet) => sheet.id === sheetId)?.name ??
      savedWorkbook?.workbook.sheets.find((sheet) => sheet.id === sheetId)
        ?.name ??
      "Removed worksheet",
    pathLabel = (path) =>
      path[0] === "sheets" && path[1] === "order"
        ? "Worksheet order"
        : path[0] === "sheets"
          ? [
              sheetNameFor(path[1]),
              ...path.slice(2).map(humanizeFieldName),
            ].join(" · ")
          : path.map(humanizeFieldName).join(" · "),
    propertyChanges = comparison
      ? [
          ...comparison.metadataChanges,
          ...comparison.structuralChanges,
        ].flatMap(collectConflictChanges)
      : [],
    formatChangeValue = (path, value) =>
      path[0] === "sheets" && path[1] === "order" && Array.isArray(value)
        ? value.map((sheetId) => sheetNameFor(String(sheetId))).join(" → ") ||
          "No worksheets"
        : summarizeConflictValue(value);
  return i.jsx("div", {
    className: "wm-backdrop",
    children: i.jsxs("div", {
      ref: dialogRef,
      className: "wm-dialog pc-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-label": text("Workbook save conflict"),
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(), event.stopPropagation(), props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
              ...(dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
              ) ?? []),
            ],
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
      lang: language("Workbook save conflict"),
      "data-ui-label": "Workbook save conflict",
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              children: i.jsx(LocalizedText, {
                children: "Workbook save conflict",
              }),
            }),
            i.jsx("button", {
              "aria-label": text("Close save conflict"),
              onClick: props.onClose,
              lang: language("Close save conflict"),
              "data-ui-label": "Close save conflict",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "wm-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Your edits are still in this editor. Export them to keep a portable copy, or compare them with the saved version before reloading.",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsx("strong", {
                  children: i.jsx(LocalizedText, {
                    children: "My local workbook:",
                  }),
                }),
                " ",
                props.workbook.name,
                " ",
                i.jsx(LocalizedText, { children: "· revision" }),
                " ",
                props.workbook.revision,
              ],
            }),
            savedWorkbook &&
              i.jsxs("p", {
                children: [
                  i.jsx("strong", {
                    children: i.jsx(LocalizedText, {
                      children: "Current saved workbook:",
                    }),
                  }),
                  " ",
                  savedWorkbook.workbook.name,
                  " ",
                  i.jsx(LocalizedText, { children: "· revision " }),
                  savedWorkbook.workbook.revision,
                ],
              }),
            i.jsxs("div", {
              className: "pc-actions",
              children: [
                i.jsx("button", {
                  onClick: props.onExport,
                  children: i.jsx(LocalizedText, {
                    children: "Export my changes",
                  }),
                }),
                i.jsx("button", {
                  onClick: onClick,
                  children: i.jsx(LocalizedText, {
                    children: "Refresh comparison",
                  }),
                }),
              ],
            }),
            savedWorkbook &&
              savedWorkbook.workbook.id !== props.workbook.id &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "The saved location now contains a different workbook. Export your local workbook before deciding whether to load it.",
                }),
              }),
            comparison &&
              i.jsxs("section", {
                "aria-label": text("Conflicting workbook changes"),
                lang: language("Conflicting workbook changes"),
                "data-ui-label": "Conflicting workbook changes",
                children: [
                  i.jsx("h3", {
                    children: i.jsx(LocalizedText, {
                      children: "Changes from saved to my local workbook",
                    }),
                  }),
                  i.jsxs("p", {
                    children: [
                      comparison.cellChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "cell changes ·" }),
                      " ",
                      comparison.metadataChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "workbook changes ·" }),
                      " ",
                      comparison.structuralChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "worksheet changes" }),
                    ],
                  }),
                  comparison.cellChanges.length > 0 &&
                    i.jsxs("table", {
                      className: "pc-table",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Cell",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Saved",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "My local value",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: comparison.cellChanges
                            .slice(0, 100)
                            .map((cellChange) => {
                              const changedFields = [
                                ...new Set([
                                  ...Object.keys(cellChange.before ?? {}),
                                  ...Object.keys(cellChange.after ?? {}),
                                ]),
                              ].filter(
                                (field) =>
                                  field !== "raw" &&
                                  stableSerialize(
                                    cellChange.before?.[field],
                                  ) !==
                                    stableSerialize(cellChange.after?.[field]),
                              );
                              return i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsxs("th", {
                                      scope: "row",
                                      children: [
                                        sheetNameFor(cellChange.sheetId),
                                        "!",
                                        cellChange.address,
                                      ],
                                    }),
                                    [cellChange.before, cellChange.after].map(
                                      (cell, sideIndex) =>
                                        i.jsxs(
                                          "td",
                                          {
                                            children: [
                                              i.jsx("span", {
                                                className: "pc-cell-value",
                                                children:
                                                  describeCellContent(cell),
                                              }),
                                              changedFields.length > 0 &&
                                                i.jsx("ul", {
                                                  className: "pc-cell-fields",
                                                  children: changedFields.map(
                                                    (changedField) =>
                                                      i.jsxs(
                                                        "li",
                                                        {
                                                          children: [
                                                            i.jsxs("strong", {
                                                              children: [
                                                                humanizeFieldName(
                                                                  changedField,
                                                                ),
                                                                ":",
                                                              ],
                                                            }),
                                                            " ",
                                                            summarizeConflictValue(
                                                              cell?.[
                                                                changedField
                                                              ],
                                                            ),
                                                          ],
                                                        },
                                                        changedField,
                                                      ),
                                                  ),
                                                }),
                                            ],
                                          },
                                          sideIndex,
                                        ),
                                    ),
                                  ],
                                },
                                `${cellChange.sheetId}:${cellChange.address}`,
                              );
                            }),
                        }),
                      ],
                    }),
                  comparison.cellChanges.length > 100 &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Showing the first 100 changed cells. Full comparison details include every change.",
                      }),
                    }),
                  propertyChanges.length > 0 &&
                    i.jsxs("table", {
                      className: "pc-table",
                      children: [
                        i.jsx("thead", {
                          children: i.jsxs("tr", {
                            children: [
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Property",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "Saved",
                                }),
                              }),
                              i.jsx("th", {
                                scope: "col",
                                children: i.jsx(LocalizedText, {
                                  children: "My local setting",
                                }),
                              }),
                            ],
                          }),
                        }),
                        i.jsx("tbody", {
                          children: propertyChanges
                            .slice(0, 100)
                            .map((propertyChange) =>
                              i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      scope: "row",
                                      children: pathLabel(propertyChange.path),
                                    }),
                                    i.jsx("td", {
                                      children: formatChangeValue(
                                        propertyChange.path,
                                        propertyChange.before,
                                      ),
                                    }),
                                    i.jsx("td", {
                                      children: formatChangeValue(
                                        propertyChange.path,
                                        propertyChange.after,
                                      ),
                                    }),
                                  ],
                                },
                                JSON.stringify(propertyChange.path),
                              ),
                            ),
                        }),
                      ],
                    }),
                  propertyChanges.length > 100 &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Showing the first 100 changed properties. Full comparison details include every change.",
                      }),
                    }),
                  (comparison.cellChanges.length > 0 ||
                    propertyChanges.length > 0) &&
                    i.jsxs("details", {
                      className: "pc-details",
                      open: open,
                      onToggle: (event) => setOpen(event.currentTarget.open),
                      children: [
                        i.jsx("summary", {
                          children: i.jsx(LocalizedText, {
                            children: "Full comparison details",
                          }),
                        }),
                        open &&
                          i.jsxs(i.Fragment, {
                            children: [
                              i.jsx("p", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Complete recorded values, formulas, notes, formatting and properties, including long values and additional items summarized above.",
                                }),
                              }),
                              i.jsx("pre", {
                                tabIndex: 0,
                                children: JSON.stringify(comparison, null, 2),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "pc-footer",
          children: [
            loadError && i.jsx("p", { role: "alert", children: loadError }),
            i.jsxs("label", {
              className: "wm-confirm",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: checked,
                  disabled: !savedWorkbook,
                  onChange: (event) => setChecked(event.target.checked),
                }),
                i.jsx(LocalizedText, {
                  children:
                    "Discard my local changes and load the saved workbook.",
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Reloading replaces this editor's data and clears its local Undo/Redo history. Other open workbooks remain unchanged.",
              }),
            }),
            i.jsxs("div", {
              className: "pc-footer-actions",
              children: [
                i.jsx("button", {
                  disabled: !savedWorkbook || !checked,
                  onClick: () => {
                    try {
                      props.onReload(savedWorkbook);
                    } catch (error) {
                      (setChecked(false),
                        setLoadError(
                          error instanceof Error
                            ? error.message
                            : "The saved workbook could not be loaded.",
                        ));
                    }
                  },
                  children: i.jsx(LocalizedText, {
                    children: "Reload saved workbook",
                  }),
                }),
                i.jsx("button", {
                  onClick: props.onClose,
                  children: i.jsx(LocalizedText, {
                    children: "Keep editing my local copy",
                  }),
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}
const diagnosticGroupDefinitions = {
    persistence: ["load", "save", "checkpoint", "list", "compare", "restore"],
    calculation: ["run"],
    file: [
      "workbook",
      "delimited",
      "automation",
      "json",
      "picture",
      "formats",
      "source",
    ],
    provider: ["request"],
    source: ["catalog", "authorize", "refresh"],
    host: ["change", "selection", "loaded", "save-state", "proposal"],
  },
  wH = new Set([
    "unavailable",
    "timeout",
    "cancelled",
    "conflict",
    "corrupt",
    "quota",
    "limit",
    "invalid",
    "refused",
    "unknown",
  ]),
  workbookErrorCodes = {
    expired: "refused",
    schema: "invalid",
    stale: "conflict",
    FILE_TIMEOUT: "timeout",
    FILE_CANCELLED: "cancelled",
    FILE_LIMIT: "limit",
    FILE_BUSY: "limit",
    FILE_INVALID: "invalid",
    FILE_ACTIVE: "refused",
    FILE_TYPE: "invalid",
    FILE_READ: "unavailable",
    FILE_WORKER: "unavailable",
    malformed_response: "invalid",
    invalid_request: "invalid",
    "#TIMEOUT!": "timeout",
    "#LIMIT!": "limit",
    "#BLOCKED!": "unavailable",
    "#PENDING!": "cancelled",
  };
function isAbortOrQuotaError(error) {
  try {
    if (error && typeof error == "object") {
      const { code, name } = error;
      if (typeof code == "string") {
        if (wH.has(code)) return code;
        if (Object.hasOwn(workbookErrorCodes, code))
          return workbookErrorCodes[code];
      }
      if (name === "AbortError") return "cancelled";
      if (name === "TimeoutError") return "timeout";
      if (name === "QuotaExceededError") return "quota";
    }
  } catch {}
  return "unknown";
}
function createWorkbookDiagnosticsMonitor(options = {}) {
  const maxRetainedEvents = options.maxEvents ?? 200;
  if (
    !Number.isInteger(maxRetainedEvents) ||
    maxRetainedEvents < 2 ||
    maxRetainedEvents > 1e3
  )
    throw new Error(
      "Diagnostic history must retain between 2 and 1000 events.",
    );
  const listeners = new Set(),
    latestAttempts = new Map();
  let eventSequence = 0,
    attemptCounter = 0;
  const initialHealth = Object.fromEntries(
    Object.keys(diagnosticGroupDefinitions).map((diagnosticGroupDefinition) => [
      diagnosticGroupDefinition,
      Object.freeze({
        active: 0,
        completed: 0,
        failures: 0,
        cancellations: 0,
        maxElapsedMs: 0,
        totalElapsedMs: 0,
        issues: Object.freeze({}),
      }),
    ]),
  );
  let snapshot = Object.freeze({
      events: Object.freeze([]),
      health: Object.freeze(initialHealth),
      droppedEvents: 0,
      sinkFailures: 0,
      droppedDeliveries: 0,
    }),
    isNotifyScheduled = false;
  const notifySubscribers = () => {
      isNotifyScheduled ||
        ((isNotifyScheduled = true),
        queueMicrotask(() => {
          isNotifyScheduled = false;
          for (const listener of listeners)
            try {
              listener();
            } catch {}
        }));
    },
    recordSinkFailure = () => {
      ((snapshot = Object.freeze({
        ...snapshot,
        sinkFailures: snapshot.sinkFailures + 1,
      })),
        notifySubscribers());
    };
  let pendingDeliveries = 0;
  const recordEvent = (diagnosticEvent, adapterHealth) => {
      const nextEvents = [...snapshot.events, Object.freeze(diagnosticEvent)],
        droppedCount = Math.max(0, nextEvents.length - maxRetainedEvents);
      if (
        ((snapshot = Object.freeze({
          ...snapshot,
          events: Object.freeze(nextEvents.slice(-maxRetainedEvents)),
          health: Object.freeze({
            ...snapshot.health,
            [diagnosticEvent.adapter]: Object.freeze(adapterHealth),
          }),
          droppedEvents: snapshot.droppedEvents + droppedCount,
        })),
        notifySubscribers(),
        pendingDeliveries >= 16)
      ) {
        snapshot = Object.freeze({
          ...snapshot,
          droppedDeliveries: snapshot.droppedDeliveries + 1,
        });
        return;
      }
      try {
        const deliveryPromise = options.onEvent?.(diagnosticEvent);
        deliveryPromise &&
          (pendingDeliveries++,
          Promise.resolve(deliveryPromise).then(
            () => {
              pendingDeliveries--;
            },
            () => {
              (pendingDeliveries--, recordSinkFailure());
            },
          ));
      } catch {
        recordSinkFailure();
      }
    },
    start = (adapter, operation) => {
      if (
        !Object.hasOwn(diagnosticGroupDefinitions, adapter) ||
        !diagnosticGroupDefinitions[adapter].includes(operation)
      )
        throw new Error("Unknown diagnostic adapter or operation.");
      const attempt = ++attemptCounter,
        startedAt = performance.now(),
        attemptKey = `${adapter}:${operation}`;
      let isFinished = false;
      recordEvent(
        {
          version: 1,
          sequence: ++eventSequence,
          attempt: attempt,
          adapter: adapter,
          operation: operation,
          phase: "started",
        },
        {
          ...snapshot.health[adapter],
          active: snapshot.health[adapter].active + 1,
        },
      );
      const finish = (outcome = "succeeded", errorCode) => {
        if (isFinished) return;
        ((isFinished = true),
          ["succeeded", "failed", "cancelled", "recovered"].includes(outcome) ||
            (outcome = "failed"));
        const code = errorCode && wH.has(errorCode) ? errorCode : "unknown",
          elapsedMs = Math.max(
            0,
            Math.round((performance.now() - startedAt) * 10) / 10,
          ),
          previousHealth = snapshot.health[adapter],
          issuesByOperation = { ...previousHealth.issues };
        (attempt >= (latestAttempts.get(attemptKey) ?? 0) &&
          outcome !== "cancelled" &&
          (latestAttempts.set(attemptKey, attempt),
          outcome === "failed"
            ? (issuesByOperation[operation] = code)
            : delete issuesByOperation[operation]),
          recordEvent(
            {
              version: 1,
              sequence: ++eventSequence,
              attempt: attempt,
              adapter: adapter,
              operation: operation,
              phase: "completed",
              outcome: outcome,
              ...(outcome === "failed"
                ? { code: code }
                : outcome === "cancelled"
                  ? { code: "cancelled" }
                  : {}),
              elapsedMs: elapsedMs,
            },
            {
              active: Math.max(0, previousHealth.active - 1),
              completed: previousHealth.completed + 1,
              failures: previousHealth.failures + +(outcome === "failed"),
              cancellations:
                previousHealth.cancellations + +(outcome === "cancelled"),
              maxElapsedMs: Math.max(previousHealth.maxElapsedMs, elapsedMs),
              totalElapsedMs: previousHealth.totalElapsedMs + elapsedMs,
              issues: Object.freeze(issuesByOperation),
            },
          ));
      };
      return {
        finish: finish,
        fail(error) {
          const failureCode = isAbortOrQuotaError(error);
          finish(
            failureCode === "cancelled" ? "cancelled" : "failed",
            failureCode,
          );
        },
      };
    };
  return {
    start: start,
    run(adapter, operation, callback) {
      const attemptHandle = start(adapter, operation);
      try {
        const result = callback();
        return result && typeof result.then == "function"
          ? Promise.resolve(result).then(
              (value) => (attemptHandle.finish(), value),
              (reason) => {
                throw (attemptHandle.fail(reason), reason);
              },
            )
          : (attemptHandle.finish(), result);
      } catch (error) {
        throw (attemptHandle.fail(error), error);
      }
    },
    getSnapshot: () => snapshot,
    subscribe(listener) {
      return (
        listeners.add(listener),
        () => {
          listeners.delete(listener);
        }
      );
    },
    clearEvents() {
      ((snapshot = Object.freeze({ ...snapshot, events: Object.freeze([]) })),
        notifySubscribers());
    },
  };
}
const invalidStorageAckError = () =>
    new WorkbookRevisionConflictError(
      "invalid",
      "The storage adapter returned an invalid acknowledgement. Your local changes are retained; check the saved copy before retrying.",
    ),
  workbookSnapshotsEqual = (leftValue, rightValue) =>
    leftValue === rightValue ||
    JSON.stringify(leftValue) === JSON.stringify(rightValue);
function captureWorkbookView(workbook, view) {
  const sheet = workbook.sheets.find(
    (sheet) => !sheet.hidden && sheet.id === view.activeSheetId,
  );
  if (!sheet || view.selection.sheetId !== sheet.id)
    throw invalidStorageAckError();
  const rangeAddress = parseRangeAddress(view.selection.range);
  if (
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount
  )
    throw invalidStorageAckError();
  const firstAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    lastAddress = `${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`;
  return {
    activeSheetId: sheet.id,
    selection: {
      sheetId: sheet.id,
      range:
        firstAddress === lastAddress
          ? firstAddress
          : `${firstAddress}:${lastAddress}`,
    },
  };
}
function normalizeCheckpointRecord(record, savedRecord) {
  try {
    const canonicalWorkbookSnapshot = buildCanonicalWorkbookSnapshot(
        record.workbook,
      ),
      view = captureWorkbookView(canonicalWorkbookSnapshot, record.view),
      guard = record.guard;
    if (
      !guard ||
      typeof guard.namespace != "string" ||
      !guard.namespace ||
      !Number.isSafeInteger(guard.generation) ||
      guard.generation < 0 ||
      typeof guard.digest != "string" ||
      !guard.digest ||
      guard.workbookId !== canonicalWorkbookSnapshot.id ||
      guard.revision !== canonicalWorkbookSnapshot.revision ||
      typeof record.recovered != "boolean" ||
      !Array.isArray(record.checkpoints) ||
      record.checkpoints.length > 20 ||
      record.checkpoints.some(
        (checkpoint) =>
          !checkpoint ||
          typeof checkpoint.id != "string" ||
          typeof checkpoint.label != "string" ||
          typeof checkpoint.workbookId != "string" ||
          !Number.isSafeInteger(checkpoint.revision) ||
          checkpoint.revision < 0 ||
          typeof checkpoint.createdAt != "string" ||
          !Number.isFinite(Date.parse(checkpoint.createdAt)),
      ) ||
      (savedRecord &&
        (!workbookSnapshotsEqual(
          canonicalWorkbookSnapshot,
          buildCanonicalWorkbookSnapshot(savedRecord.workbook),
        ) ||
          !workbookSnapshotsEqual(
            view,
            captureWorkbookView(savedRecord.workbook, savedRecord.view),
          )))
    )
      throw invalidStorageAckError();
    return {
      ...record,
      workbook: canonicalWorkbookSnapshot,
      view: view,
      guard: { ...guard },
      checkpoints: record.checkpoints.map((checkpoint) => ({ ...checkpoint })),
    };
  } catch {
    throw invalidStorageAckError();
  }
}
function instrumentWorkbookPersistence(
  storageAdapter,
  activeReliabilityMonitor,
) {
  const normalizeRecordResult = (result, savedRecord) =>
    result && typeof result.then == "function"
      ? Promise.resolve(result).then((storedRecord) =>
          normalizeCheckpointRecord(storedRecord, savedRecord),
        )
      : normalizeCheckpointRecord(result, savedRecord);
  return {
    ...("coordination" in storageAdapter
      ? { coordination: storageAdapter.coordination }
      : {}),
    load: () => {
      const attempt = activeReliabilityMonitor.start("persistence", "load");
      try {
        const loadedRecord = storageAdapter.load(),
          normalizedRecord = loadedRecord
            ? normalizeCheckpointRecord(loadedRecord)
            : null;
        return (
          attempt.finish(
            normalizedRecord?.recovered ? "recovered" : "succeeded",
          ),
          normalizedRecord
        );
      } catch (error) {
        throw (attempt.fail(error), error);
      }
    },
    save: (workbook, view, guard, s) =>
      activeReliabilityMonitor.run("persistence", "save", () => {
        const canonicalRecord = {
          workbook: buildCanonicalWorkbookSnapshot(workbook),
          view: captureWorkbookView(workbook, view),
        };
        return normalizeRecordResult(
          storageAdapter.save(
            canonicalRecord.workbook,
            structuredClone(canonicalRecord.view),
            guard ? { ...guard } : null,
            s,
          ),
          canonicalRecord,
        );
      }),
    checkpoint: (n, o, a) =>
      activeReliabilityMonitor.run("persistence", "checkpoint", () =>
        normalizeRecordResult(storageAdapter.checkpoint(n, o, a)),
      ),
    listCheckpoints: () =>
      activeReliabilityMonitor.run("persistence", "list", () =>
        storageAdapter.listCheckpoints(),
      ),
    compareCheckpoint: (n, o) =>
      activeReliabilityMonitor.run("persistence", "compare", () =>
        storageAdapter.compareCheckpoint(n, o),
      ),
    restoreCheckpoint: (n, o, a, s, d) =>
      activeReliabilityMonitor.run("persistence", "restore", () =>
        normalizeRecordResult(storageAdapter.restoreCheckpoint(n, o, a, s, d)),
      ),
  };
}
function findRecoverableCheckpoint(
  workbookPersistence,
  book,
  hasUnsavedChanges,
) {
  const loadedRecord = workbookPersistence.load();
  if (!loadedRecord) return null;
  const checkpointRecord = normalizeCheckpointRecord(loadedRecord);
  return checkpointRecord.recovered ||
    !workbookSnapshotsEqual(
      checkpointRecord.workbook,
      buildCanonicalWorkbookSnapshot(book),
    ) ||
    hasUnsavedChanges
    ? null
    : checkpointRecord;
}
const diagnosticGroupLabels = {
  persistence: "Saving and recovery",
  calculation: "Calculation",
  file: "File processing",
  provider: "AI provider",
  source: "External sources",
  host: "Application callbacks",
};
function WorkbookHealthDialog({ monitor, onClose }) {
  const snapshot = ne.useSyncExternalStore(
      monitor.subscribe,
      monitor.getSnapshot,
      monitor.getSnapshot,
    ),
    titleId = ne.useId(),
    closeButtonRef = ne.useRef(null),
    workbookMessages = useWorkbookMessages();
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      closeButtonRef.current?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const onClick = () => {
    const downloadUrl = URL.createObjectURL(
        new Blob(
          [
            JSON.stringify(
              {
                format: "ledgerglass-reliability-v1",
                ...monitor.getSnapshot(),
              },
              null,
              2,
            ),
          ],
          { type: "application/json" },
        ),
      ),
      downloadLink = document.createElement("a");
    ((downloadLink.href = downloadUrl),
      (downloadLink.download = "Workbook-diagnostics.json"),
      downloadLink.click(),
      setTimeout(() => URL.revokeObjectURL(downloadUrl), 1e3));
  };
  return i.jsx("div", {
    className: "lg-modal-backdrop",
    children: i.jsxs("section", {
      className: "lg-reliability",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.key === "Escape" && (event.stopPropagation(), onClose()),
          event.key === "Tab")
        ) {
          const focusableButtons = [
            ...event.currentTarget.querySelectorAll("button:not(:disabled)"),
          ];
          event.shiftKey && document.activeElement === focusableButtons[0]
            ? (event.preventDefault(),
              (lastFocusable = focusableButtons.at(-1)) == null ||
                lastFocusable.focus())
            : !event.shiftKey &&
              document.activeElement === focusableButtons.at(-1) &&
              (event.preventDefault(),
              (firstFocusable = focusableButtons[0]) == null ||
                firstFocusable.focus());
        }
      },
      children: [
        i.jsxs("header", {
          children: [
            i.jsx("h2", {
              id: titleId,
              children: i.jsx(LocalizedText, { children: "Workbook health" }),
            }),
            i.jsx("button", {
              ref: closeButtonRef,
              type: "button",
              "aria-label": workbookMessages.text("Close workbook health"),
              onClick: onClose,
              children: i.jsx(LocalizedText, { children: "Close" }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Diagnostics contain operation categories, outcomes and timings. Workbook contents, names, prompts and error details are excluded.",
          }),
        }),
        i.jsx("div", {
          className: "lg-reliability-table",
          children: i.jsxs("table", {
            children: [
              i.jsx("thead", {
                children: i.jsx("tr", {
                  children: [
                    "Service",
                    "State",
                    "Completed",
                    "Failures",
                    "Longest operation",
                  ].map((headerLabel) =>
                    i.jsx(
                      "th",
                      {
                        scope: "col",
                        children: i.jsx(LocalizedText, {
                          children: headerLabel,
                        }),
                      },
                      headerLabel,
                    ),
                  ),
                }),
              }),
              i.jsx("tbody", {
                children: Object.keys(diagnosticGroupLabels).map(
                  (diagnosticGroupLabel) => {
                    const groupHealth = snapshot.health[diagnosticGroupLabel],
                      statusLabel = Object.keys(groupHealth.issues).length
                        ? "Needs attention"
                        : groupHealth.active
                          ? "Working…"
                          : groupHealth.completed
                            ? "Ready"
                            : "Not used";
                    return i.jsxs(
                      "tr",
                      {
                        children: [
                          i.jsx("th", {
                            scope: "row",
                            children: i.jsx(LocalizedText, {
                              children:
                                diagnosticGroupLabels[diagnosticGroupLabel],
                            }),
                          }),
                          i.jsx("td", {
                            children: i.jsx(LocalizedText, {
                              children: statusLabel,
                            }),
                          }),
                          i.jsx("td", { children: groupHealth.completed }),
                          i.jsx("td", { children: groupHealth.failures }),
                          i.jsxs("td", {
                            children: [
                              workbookMessages.locale
                                ? groupHealth.maxElapsedMs.toLocaleString(
                                    workbookMessages.locale,
                                  )
                                : groupHealth.maxElapsedMs,
                              " ",
                              "ms",
                            ],
                          }),
                        ],
                      },
                      diagnosticGroupLabel,
                    );
                  },
                ),
              }),
            ],
          }),
        }),
        i.jsxs("p", {
          role: "status",
          children: [
            i.jsx(LocalizedText, { children: "Retained events" }),
            ": ",
            snapshot.events.length,
            " ",
            "· ",
            i.jsx(LocalizedText, { children: "Discarded older events" }),
            ":",
            " ",
            snapshot.droppedEvents,
            " ·",
            " ",
            i.jsx(LocalizedText, { children: "Diagnostic delivery failures" }),
            ":",
            " ",
            snapshot.sinkFailures,
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Save failures remain visible in the editor. Export your local changes before reloading a saved copy. A pending save must finish before another retry.",
          }),
        }),
        i.jsxs("footer", {
          children: [
            i.jsx("button", {
              type: "button",
              onClick: onClick,
              children: i.jsx(LocalizedText, {
                children: "Download diagnostics",
              }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: () => monitor.clearEvents(),
              children: i.jsx(LocalizedText, {
                children: "Clear diagnostic history",
              }),
            }),
          ],
        }),
      ],
    }),
  });
}
const cellSnapshotCache = new WeakMap();
function snapshotSheetContent(sheet) {
  return fingerprintCanonicalValue({
    cells: Object.fromEntries(
      Object.entries(sheet.cells).map(([cellAddress, cell]) => [
        cellAddress,
        { ...cell, style: undefined, numFmt: undefined, format: undefined },
      ]),
    ),
    tables: sheet.tables,
    merges: sheet.merges,
    rowOrigins: sheet.rowOriginIds,
  });
}
function checkSourceSheetAccess(workbook, sheetId, options) {
  if (options.readOnly) throw new ExternalSourceError("refused");
  if (!sheetId) {
    assertWorkbookStructureEditable(workbook);
    return;
  }
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (
    !sheet ||
    sheet.hidden ||
    sheet.veryHidden ||
    sheet.protection?.enabled ||
    (options.allowedSheetIds && !options.allowedSheetIds.includes(sheetId))
  )
    throw new ExternalSourceError("refused");
  if (!sheet.externalConnection) throw new ExternalSourceError("invalid");
  if (snapshotSheetContent(sheet) !== sheet.externalConnection.outputHash)
    throw new ExternalSourceError("stale");
}
function EH(workbook, snapshot, grant, plan, options = {}) {
  return evaluateExternalSourceQuery(workbook, snapshot, grant, plan, options);
}
function evaluateExternalSourceQuery(
  workbook,
  snapshot,
  grant,
  plan,
  options = {},
  batchSheetIds = [],
) {
  checkSourceSheetAccess(workbook, plan.sheetId, options);
  const activeSourceApproval = parseActiveSourceApproval(grant),
    payload = validateExternalSourcePayload(
      snapshot,
      activeSourceApproval.source,
    ),
    existingConnection = plan.sheetId
      ? workbook.sheets.find((sheet) => sheet.id === plan.sheetId)
          .externalConnection
      : undefined;
  if (
    existingConnection &&
    !isSameContentHash(existingConnection.source, payload.source)
  )
    throw new ExternalSourceError("schema");
  if (
    !plan.sheetId &&
    (workbook.sheets.filter((sheet) => sheet.externalConnection).length >=
      externalSourceLimits.sources ||
      !plan.outputName.trim() ||
      plan.outputName !== plan.outputName.trim() ||
      plan.outputName.length > 31 ||
      /[\\/\[\]*?:\u0000-\u001f]/.test(plan.outputName) ||
      /^'|'$/.test(plan.outputName) ||
      workbook.sheets.some(
        (sheet) => sheet.name.toLowerCase() === plan.outputName.toLowerCase(),
      ))
  )
    throw new ExternalSourceError("invalid");
  if (
    (plan.sheetId || inspectOutputNameSafety(workbook, plan.outputName),
    workbook.sheets
      .filter(
        (sheet) =>
          sheet.id !== plan.sheetId && !batchSheetIds.includes(sheet.id),
      )
      .reduce((total, sheet) => total + Object.keys(sheet.cells).length, 0) +
      (payload.rows.length + 1) * payload.source.columns.length >
      workbookLimits.maxCells)
  )
    throw new ExternalSourceError("limit");
  const queryPlan = {
    workbookHash: fingerprintCanonicalValue(workbook),
    outputName: plan.outputName,
    ...(plan.sheetId ? { sheetId: plan.sheetId } : {}),
    snapshot: payload,
    grant: activeSourceApproval,
    ...(existingConnection
      ? {
          previous: {
            rows: existingConnection.lastSuccess.rows,
            controls: structuredClone(existingConnection.lastSuccess.controls),
            at: existingConnection.lastSuccess.at,
          },
        }
      : {}),
    controls: computeSourceControlTotals(payload),
  };
  if (
    queryPlan.controls.some(
      (control) =>
        !Number.isFinite(
          control.sum -
            (queryPlan.previous?.controls.find(
              (previousControl) => previousControl.field === control.field,
            )?.sum ?? 0),
        ),
    )
  )
    throw new ExternalSourceError("limit");
  return (
    cellSnapshotCache.set(queryPlan, fingerprintCanonicalValue(queryPlan)),
    queryPlan
  );
}
function grantSheetSnapshotSource(workbook, plan, grant, options = {}) {
  if (
    !cellSnapshotCache.has(plan) ||
    cellSnapshotCache.get(plan) !== fingerprintCanonicalValue(plan) ||
    fingerprintCanonicalValue(workbook) !== plan.workbookHash
  )
    throw new ExternalSourceError("stale");
  checkSourceSheetAccess(workbook, plan.sheetId, options);
  const activeSourceApproval = parseActiveSourceApproval(grant);
  if (
    (parseActiveSourceApproval(plan.grant),
    !isSameContentHash(activeSourceApproval.source, plan.snapshot.source))
  )
    throw new ExternalSourceError("schema");
  EH(workbook, plan.snapshot, activeSourceApproval, plan, options);
  const sourceSheetSnapshot = buildSourceSheetSnapshot(workbook, plan),
    workbookInput = normalizeWorkbookInput({
      ...workbook,
      sheets: plan.sheetId
        ? workbook.sheets.map((sheet) =>
            sheet.id === sourceSheetSnapshot.id ? sourceSheetSnapshot : sheet,
          )
        : [...workbook.sheets, sourceSheetSnapshot],
    });
  return (cellSnapshotCache.delete(plan), workbookInput);
}
function buildSourceSheetSnapshot(workbook, plan) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === plan.sheetId),
    targetSheet = sheet ? { ...sheet } : createWorksheetRecord(plan.outputName),
    cells = {};
  ([
    plan.snapshot.source.columns.map((column) => column.name),
    ...plan.snapshot.rows,
  ].forEach((row, rowIndex) =>
    row.forEach((raw, columnIndex) => {
      const cellAddress = `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`,
        existingCell = sheet?.cells[cellAddress];
      cells[cellAddress] = {
        raw: raw,
        ...(typeof raw == "string" && /^[=+@-]/.test(raw)
          ? { literal: true }
          : {}),
        ...(existingCell?.style
          ? { style: structuredClone(existingCell.style) }
          : rowIndex === 0
            ? { style: { bold: true, fill: "#eef2ff" } }
            : {}),
        ...(existingCell?.numFmt ? { numFmt: existingCell.numFmt } : {}),
        ...(existingCell?.format ? { format: existingCell.format } : {}),
        note: `Source ${plan.snapshot.source.id} · version ${plan.snapshot.version} · ${rowIndex === 0 ? "header" : `record ${rowIndex}`} · as of ${plan.snapshot.sourceTimestamp}`,
      };
    }),
  ),
    (targetSheet.cells = cells));
  const syncedAt = new Date().toISOString();
  return (
    (targetSheet.externalConnection = {
      protocol: externalSourceFormatVersion,
      source: structuredClone(plan.snapshot.source),
      lastSuccess: {
        at: syncedAt,
        sourceTimestamp: plan.snapshot.sourceTimestamp,
        version: plan.snapshot.version,
        rows: plan.snapshot.rows.length,
        controls: structuredClone(plan.controls),
        dataHash: fingerprintCanonicalValue(plan.snapshot),
      },
      lastAttempt: { at: syncedAt, outcome: "succeeded" },
      outputHash: snapshotSheetContent(targetSheet),
    }),
    targetSheet
  );
}
const sourceGrantCache = new WeakMap();
function assertSourceGrantsValid(workbook, sheetIds, options = {}) {
  if (
    !sheetIds.length ||
    sheetIds.length > externalSourceLimits.sources ||
    new Set(sheetIds).size !== sheetIds.length
  )
    throw new ExternalSourceError("invalid");
  for (const sheetId of sheetIds)
    checkSourceSheetAccess(workbook, sheetId, options);
}
function planSourceGrants(workbook, grants, options = {}) {
  const sheetIds = grants.map((entry) => entry.sheetId);
  if (
    (assertSourceGrantsValid(workbook, sheetIds, options),
    new TextEncoder().encode(JSON.stringify(grants)).length >
      externalSourceLimits.responseBytes)
  )
    throw new ExternalSourceError("limit");
  const plans = grants.map((grantEntry) =>
    evaluateExternalSourceQuery(
      workbook,
      grantEntry.snapshot,
      grantEntry.grant,
      {
        sheetId: grantEntry.sheetId,
        outputName: workbook.sheets.find(
          (sheet) => sheet.id === grantEntry.sheetId,
        ).name,
      },
      options,
      sheetIds,
    ),
  );
  if (
    workbook.sheets
      .filter((sheet) => !sheetIds.includes(sheet.id))
      .reduce((total, sheet) => total + Object.keys(sheet.cells).length, 0) +
      plans.reduce(
        (planCells, plan) =>
          planCells +
          (plan.snapshot.rows.length + 1) * plan.snapshot.source.columns.length,
        0,
      ) >
    workbookLimits.maxCells
  )
    throw new ExternalSourceError("limit");
  const planBundle = {
    workbookHash: fingerprintCanonicalValue(workbook),
    plans: plans,
  };
  return (
    sourceGrantCache.set(planBundle, fingerprintCanonicalValue(planBundle)),
    planBundle
  );
}
function applySourceGrantPlans(workbook, planBundle, approvals, options = {}) {
  if (
    !sourceGrantCache.has(planBundle) ||
    sourceGrantCache.get(planBundle) !==
      fingerprintCanonicalValue(planBundle) ||
    planBundle.workbookHash !== fingerprintCanonicalValue(workbook)
  )
    throw new ExternalSourceError("stale");
  const parsedApprovals = approvals.map((approval) =>
      parseActiveSourceApproval(approval),
    ),
    planSourceIds = new Set(
      planBundle.plans.map((plan) => plan.snapshot.source.id),
    );
  if (
    parsedApprovals.length !== planSourceIds.size ||
    new Set(parsedApprovals.map((parsedApproval) => parsedApproval.source.id))
      .size !== parsedApprovals.length ||
    parsedApprovals.some(
      (unplannedApproval) => !planSourceIds.has(unplannedApproval.source.id),
    )
  )
    throw new ExternalSourceError("refused");
  for (const plan of planBundle.plans) {
    parseActiveSourceApproval(plan.grant);
    const matchingApproval = parsedApprovals.find(
      (approvalCandidate) =>
        approvalCandidate.source.id === plan.snapshot.source.id,
    );
    if (
      !matchingApproval ||
      !isSameContentHash(matchingApproval.source, plan.snapshot.source)
    )
      throw new ExternalSourceError("schema");
  }
  planSourceGrants(
    workbook,
    planBundle.plans.map((plan) => ({
      sheetId: plan.sheetId,
      snapshot: plan.snapshot,
      grant: parsedApprovals.find(
        (grantApproval) => grantApproval.source.id === plan.snapshot.source.id,
      ),
    })),
    options,
  );
  const snapshotsBySheetId = new Map(
      planBundle.plans.map((plan) => [
        plan.sheetId,
        buildSourceSheetSnapshot(workbook, plan),
      ]),
    ),
    workbookInput = normalizeWorkbookInput({
      ...workbook,
      sheets: workbook.sheets.map(
        (sheet) => snapshotsBySheetId.get(sheet.id) ?? sheet,
      ),
    });
  sourceGrantCache.delete(planBundle);
  for (const plan of planBundle.plans) cellSnapshotCache.delete(plan);
  return workbookInput;
}
function normalizeExternalSourceState(
  workbook,
  sheetId,
  sourceId,
  code,
  options = {},
) {
  if (options.readOnly) return workbook;
  const targetSheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  if (
    !targetSheet?.externalConnection ||
    targetSheet.externalConnection.source.id !== sourceId ||
    targetSheet.hidden ||
    targetSheet.veryHidden ||
    targetSheet.protection?.enabled ||
    (options.allowedSheetIds && !options.allowedSheetIds.includes(sheetId))
  )
    return workbook;
  const externalConnection = targetSheet.externalConnection,
    attemptAt = new Date(
      Math.max(Date.now(), Date.parse(externalConnection.lastSuccess.at)),
    ).toISOString();
  return normalizeWorkbookInput({
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet === targetSheet
        ? {
            ...sheet,
            externalConnection: {
              ...externalConnection,
              lastAttempt: {
                at: attemptAt,
                outcome: code === "cancelled" ? "cancelled" : "failed",
                code: code,
              },
            },
          }
        : sheet,
    ),
  });
}
async function runSourceWithTimeout(operationName, operation, options = {}) {
  const timeoutMs = options.timeoutMs ?? externalSourceLimits.timeoutMs;
  if (
    !Number.isInteger(timeoutMs) ||
    timeoutMs < 1 ||
    timeoutMs > externalSourceLimits.timeoutMs
  )
    throw new ExternalSourceError("invalid");
  const reliabilityRecord = options.reliability?.start("source", operationName),
    abortController = new AbortController();
  let timeoutId,
    onAbort = () => {};
  try {
    if (options.signal?.aborted) throw new ExternalSourceError("cancelled");
    const timeoutPromise = new Promise((resolve, reject) => {
        ((onAbort = () => {
          (abortController.abort(),
            reject(new ExternalSourceError("cancelled")));
        }),
          options.signal?.addEventListener("abort", onAbort, { once: true }),
          (timeoutId = setTimeout(() => {
            (abortController.abort(),
              reject(new ExternalSourceError("timeout")));
          }, timeoutMs)));
      }),
      operationResult = await Promise.race([
        Promise.resolve().then(() => operation(abortController.signal)),
        timeoutPromise,
      ]);
    if (abortController.signal.aborted)
      throw new ExternalSourceError("cancelled");
    return (reliabilityRecord?.finish(), operationResult);
  } catch (error) {
    const sourceError = isExternalSourceError(error);
    throw (reliabilityRecord?.fail(sourceError), sourceError);
  } finally {
    (clearTimeout(timeoutId),
      options.signal?.removeEventListener("abort", onAbort));
  }
}
async function fetchSourceCatalog(adapter, options = {}) {
  return runSourceWithTimeout(
    "catalog",
    async (signal) => {
      const rawCatalog = await adapter.catalog({ signal: signal });
      if (
        !Array.isArray(rawCatalog) ||
        rawCatalog.length > externalSourceLimits.sources
      )
        throw new ExternalSourceError("invalid");
      const parsedCatalog = rawCatalog.map((approval) =>
        parseActiveSourceApproval(approval),
      );
      if (
        new Set(parsedCatalog.map((parsedApproval) => parsedApproval.source.id))
          .size !== parsedCatalog.length
      )
        throw new ExternalSourceError("invalid");
      return parsedCatalog;
    },
    options,
  );
}
async function authorizeSourceAccess(adapter, sourceId, options = {}) {
  return runSourceWithTimeout(
    "authorize",
    async (signal) => {
      const activeSourceApproval = parseActiveSourceApproval(
        await adapter.authorize(sourceId, { signal: signal }),
      );
      if (activeSourceApproval.source.id !== sourceId)
        throw new ExternalSourceError("refused");
      return activeSourceApproval;
    },
    options,
  );
}
async function refreshSourceSnapshot(adapter, sourceId, options = {}) {
  const grant = await authorizeSourceAccess(adapter, sourceId, options),
    snapshot = await runSourceWithTimeout(
      "refresh",
      async (signal) => {
        const validatedSnapshot = validateExternalSourcePayload(
          await adapter.refresh(sourceId, { signal: signal }),
          grant.source,
        );
        return (parseActiveSourceApproval(grant), validatedSnapshot);
      },
      options,
    );
  return { grant: grant, snapshot: snapshot };
}
discriminatedUnion("operation", [
  strictObject({ operation: literal("catalog") }),
  strictObject({
    operation: _enum(["authorize", "refresh"]),
    sourceId: string().regex(/^[A-Za-z0-9][A-Za-z0-9._-]{0,99}$/),
  }),
]);
async function runSourceRefreshBatch(
  workbook,
  adapter,
  sheetIds,
  options = {},
) {
  assertSourceGrantsValid(workbook, sheetIds, options);
  const entries = [],
    resultsBySourceId = new Map();
  let failureCode;
  try {
    await runSourceWithTimeout(
      "refresh",
      async (signal) => {
        for (const sheetId of sheetIds) {
          if (signal.aborted) throw new ExternalSourceError("cancelled");
          const sourceId = workbook.sheets.find((sheet) => sheet.id === sheetId)
            .externalConnection.source.id;
          let refreshResult = resultsBySourceId.get(sourceId);
          if (!refreshResult) {
            try {
              refreshResult = await refreshSourceSnapshot(adapter, sourceId, {
                ...options,
                signal: signal,
              });
            } catch (error) {
              refreshResult = isExternalSourceError(error).code;
            }
            if (signal.aborted) throw new ExternalSourceError("cancelled");
            resultsBySourceId.set(sourceId, refreshResult);
          }
          if (
            (entries.push(
              typeof refreshResult == "string"
                ? { sheetId: sheetId, sourceId: sourceId, error: refreshResult }
                : { sheetId: sheetId, sourceId: sourceId, ...refreshResult },
            ),
            new TextEncoder().encode(JSON.stringify(entries)).length >
              externalSourceLimits.responseBytes)
          )
            throw new ExternalSourceError("limit");
          options.onProgress?.(entries.length, sheetIds.length);
        }
      },
      options,
    );
  } catch (error) {
    failureCode = isExternalSourceError(error).code;
  }
  if (failureCode) {
    for (const sheetId of sheetIds)
      entries.some((entry) => entry.sheetId === sheetId) ||
        entries.push({
          sheetId: sheetId,
          sourceId: workbook.sheets.find((sheet) => sheet.id === sheetId)
            .externalConnection.source.id,
          error: failureCode,
        });
    return failureCode === "limit"
      ? {
          entries: entries.map(({ sheetId, sourceId }) => ({
            sheetId,
            sourceId,
            error: failureCode,
          })),
          error: failureCode,
        }
      : { entries: entries, error: failureCode };
  }
  if (entries.some((entry) => entry.error)) return { entries: entries };
  try {
    const plan = planSourceGrants(
      workbook,
      entries.map((entry) => ({
        sheetId: entry.sheetId,
        snapshot: entry.snapshot,
        grant: entry.grant,
      })),
      options,
    );
    return { entries: entries, plan: plan };
  } catch (error) {
    return { entries: entries, error: isExternalSourceError(error).code };
  }
}
async function G2e(adapter, plan, options = {}) {
  return runSourceWithTimeout(
    "authorize",
    async (signal) => {
      const authorizedApprovals = [];
      for (const sourceId of new Set(
        plan.plans.map((plan) => plan.snapshot.source.id),
      )) {
        if (signal.aborted) throw new ExternalSourceError("cancelled");
        authorizedApprovals.push(
          await authorizeSourceAccess(adapter, sourceId, {
            ...options,
            signal: signal,
          }),
        );
      }
      return authorizedApprovals.map((approval) =>
        parseActiveSourceApproval(approval),
      );
    },
    options,
  );
}

export {
  instrumentWorkbookPersistence,
  computeVirtualizedRange,
  snapshotSheetContent,
  EH,
  G2e,
  applySourceGrantPlans,
  WorkbookSaveConflictDialog,
  authorizeSourceAccess,
  refreshSourceSnapshot,
  WorkbookHealthDialog,
  grantSheetSnapshotSource,
  runSourceRefreshBatch,
  normalizeExternalSourceState,
  workbookLockName,
  fetchSourceCatalog,
  createWorkbookDiagnosticsMonitor,
  saveWorkbookWithLock,
  findRecoverableCheckpoint,
};
