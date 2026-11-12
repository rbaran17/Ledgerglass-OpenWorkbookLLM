import { isFiniteNumber } from "./03-workbook-model.js";
import { buildCompositeKey } from "./11-local-workbook-store-and-versions.js";
import { X } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { fingerprintCanonicalValue } from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  B2,
  enumerateSelectionCells,
  validateFormulaWatches,
  calculateWorkbookWithSettings,
  _t,
  cn,
  parseRangeAddress,
  validateDataValidationRule,
  resolveDefinedNameValue,
  columnIndexToLetters,
  matrixResultFunctionNames,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  formatCellEntryForEdit,
  validateLocalizedNumberEntry,
} from "./04-cell-input-paste-and-fill.js";
import { inspectFormulaDependencies } from "./12-consolidation-and-dependency-inspection.js";
import {
  normalizeFormulaWatchContext,
  classifyWatchedCell,
} from "./35-formula-inspection-and-compatibility.js";

function applyFormulaWatchAction(workbook, request, options = {}) {
  if (options.readOnly)
    throw Error(
      "This workbook is read-only. Watch inspection remains available.",
    );
  const formulaWatchContext = normalizeFormulaWatchContext(workbook, options),
    existingWatches = validateFormulaWatches(workbook.formulaWatches, workbook);
  let nextWatches = existingWatches,
    added = [],
    removed = [];
  if (request.kind === "add") {
    const sheet = workbook.sheets.find((sheet) => sheet.id === request.sheetId);
    if (!sheet)
      throw Error("The worksheet selected for watching no longer exists.");
    if (!Array.isArray(request.ranges))
      throw Error("Provide a list of selected ranges to watch.");
    const selectedCells = enumerateSelectionCells(request.ranges, B2),
      usedWatchIds = new Set(existingWatches.map((watch) => watch.id));
    for (const address of selectedCells) {
      if (
        existingWatches.some(
          (anchoredWatch) =>
            !anchoredWatch.orphaned &&
            anchoredWatch.anchor.sheetId === sheet.id &&
            anchoredWatch.anchor.address === address,
        )
      )
        continue;
      let watchNumber = 1;
      for (; usedWatchIds.has(`watch-${watchNumber}`);) watchNumber++;
      const newWatch = {
        id: `watch-${watchNumber}`,
        anchor: { sheetId: sheet.id, sheetName: sheet.name, address: address },
      };
      if (
        classifyWatchedCell(workbook, newWatch, formulaWatchContext) !== "ready"
      )
        throw Error(
          "A selected watch source is outside the visible, authorized context. Review visibility or host access before adding it.",
        );
      (usedWatchIds.add(newWatch.id), added.push(newWatch));
    }
    if (existingWatches.length + added.length > B2)
      throw Error(
        "A workbook may contain at most 100 formula watches, including orphaned entries. Remove a watch before adding another.",
      );
    nextWatches = [...existingWatches, ...added];
  } else if (request.kind === "remove") {
    if (
      !Array.isArray(request.ids) ||
      request.ids.length > B2 ||
      request.ids.some(
        (watchId) =>
          typeof watchId != "string" ||
          !existingWatches.some((knownWatch) => knownWatch.id === watchId),
      )
    )
      throw Error(
        "A selected watch no longer exists. Review the current watch list.",
      );
    const removedIds = new Set(request.ids);
    ((removed = existingWatches.filter((removedWatch) =>
      removedIds.has(removedWatch.id),
    )),
      (nextWatches = existingWatches.filter(
        (keptWatch) => !removedIds.has(keptWatch.id),
      )));
  } else if (request.kind === "clear")
    ((removed = existingWatches), (nextWatches = []));
  else throw Error("Unknown formula watch action.");
  const next =
    added.length || removed.length
      ? { ...workbook, formulaWatches: structuredClone(nextWatches) }
      : workbook;
  return {
    snapshot: JSON.stringify(workbook),
    request: structuredClone(request),
    context: formulaWatchContext,
    added: added,
    removed: removed,
    next: next,
  };
}
function assertWatchPreviewCurrent(workbook, preview, options = {}) {
  if (JSON.stringify(workbook) !== preview.snapshot)
    throw Error(
      "The workbook changed after the watch preview. Review the action again.",
    );
  const currentAction = applyFormulaWatchAction(
    workbook,
    preview.request,
    options,
  );
  if (JSON.stringify(currentAction) !== JSON.stringify(preview))
    throw Error(
      "The watch preview or host context changed. Review the action again.",
    );
  return currentAction.next;
}
function inspectFormulaWatches(workbook, options = {}) {
  const formulaWatchContext = normalizeFormulaWatchContext(workbook, options),
    watches = validateFormulaWatches(workbook.formulaWatches, workbook),
    readyWatches = [],
    rows = watches.map((watch) => {
      const classification = classifyWatchedCell(
        workbook,
        watch,
        formulaWatchContext,
      );
      return classification === "ready"
        ? (readyWatches.push(watch),
          { id: watch.id, status: "unavailable", anchor: watch.anchor })
        : classification === "orphaned"
          ? options.allowedSheetIds !== undefined &&
            !formulaWatchContext.allowedSheetIds.includes(watch.anchor.sheetId)
            ? {
                id: watch.id,
                status: "excluded",
                note: "This watch is outside the current host context.",
              }
            : {
                id: watch.id,
                status: "orphaned",
                anchor: watch.anchor,
                note: `Source unavailable (${watch.orphaned ?? "missing-sheet"}). This watch will not reconnect automatically.`,
              }
          : {
              id: watch.id,
              status: classification,
              note:
                classification === "hidden"
                  ? "This source is hidden. Use Include hidden only within the host-authorized context."
                  : "This watch is outside the current host context.",
            };
    }),
    inspection = {
      workbookId: workbook.id,
      revision: workbook.revision,
      rows: rows,
      formulaCellsScanned: 0,
      inspectionPasses: 0,
      notes: [
        "Watch results are current local calculations, not Excel step-by-step evaluation. Unresolved, hidden or out-of-context dependencies withhold the result.",
        "The embedding host must supply authorized workbook data. Include hidden is a local visibility choice.",
      ],
    };
  if (!readyWatches.length) return inspection;
  let tempSheetName = "",
    tempSheetId = "";
  for (let attempt = 0; attempt < 4; attempt++) {
    ((tempSheetName = `_watch_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`),
      (tempSheetId = `watch-inspection-${crypto.randomUUID()}`));
    const lowerName = tempSheetName.toLowerCase();
    if (
      !workbook.sheets.some(
        (sheet) =>
          sheet.id === tempSheetId ||
          sheet.name.toLowerCase() === lowerName ||
          Object.values(sheet.cells).some(
            (cell) =>
              typeof cell.raw == "string" &&
              cell.raw.toLowerCase().includes(lowerName),
          ),
      ) &&
      !workbook.definedNames?.some((definedName) =>
        definedName.refersTo.toLowerCase().includes(lowerName),
      )
    )
      break;
    if (attempt === 3)
      throw Error(
        "A temporary watch inspection context could not be isolated.",
      );
  }
  const inspectionSheet = {
    id: tempSheetId,
    name: tempSheetName,
    rowCount: workbookLimits.maxRows,
    colCount: workbookLimits.maxCols,
    cells: Object.fromEntries(
      readyWatches.map((readyWatch, index) => [
        `A${index + 1}`,
        {
          raw: `='${readyWatch.anchor.sheetName.replace(/'/g, "''")}'!${readyWatch.anchor.address}`,
        },
      ]),
    ),
  };
  inspectionSheet.cells.B1 = { raw: `=SUM(A1:A${readyWatches.length})` };
  const dependencyInspection = inspectFormulaDependencies(
    {
      ...workbook,
      sheets: [
        inspectionSheet,
        ...workbook.sheets.map((sheet) =>
          sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
        ),
      ],
    },
    {
      sheetId: inspectionSheet.id,
      address: "B1",
      direction: "precedents",
      depth: 8,
      includeHidden: formulaWatchContext.includeHidden,
      allowedSheetIds: [
        ...formulaWatchContext.allowedSheetIds,
        inspectionSheet.id,
      ],
      limits: {
        maxFormulaCells: 5e3 + readyWatches.length + 1,
        maxNodes: 1e3,
        maxReferences: 1e4,
      },
    },
  );
  ((inspection.inspectionPasses = 1),
    (inspection.formulaCellsScanned = Math.max(
      0,
      dependencyInspection.coverage.formulaCellsScanned -
        readyWatches.length -
        1,
    )));
  const nodesByKey = new Map(
      dependencyInspection.nodes.map((node) => [
        JSON.stringify([node.sheetId, node.address]),
        node,
      ]),
    ),
    formulaCellsBySheet = new Map(
      workbook.sheets.map((sheet) => [
        sheet.id,
        Object.entries(sheet.cells)
          .filter(
            ([, cell]) =>
              !cell.literal &&
              typeof cell.raw == "string" &&
              cell.raw.startsWith("="),
          )
          .map(([address]) => _t(address)),
      ]),
    ),
    footprintBlankKeys = new Set();
  let formulaScanCount = 0,
    scanLimitReached = false;
  for (const [cellKey, cell] of nodesByKey) {
    if (cell.raw !== null) continue;
    const parsedAddress = _t(cell.address);
    for (const formulaPosition of formulaCellsBySheet.get(cell.sheetId) ?? []) {
      if (++formulaScanCount > 1e5) {
        scanLimitReached = true;
        break;
      }
      if (
        formulaPosition.row <= parsedAddress.row &&
        formulaPosition.col <= parsedAddress.col &&
        (parsedAddress.row - formulaPosition.row + 1) *
          (parsedAddress.col - formulaPosition.col + 1) <=
          1e4
      ) {
        footprintBlankKeys.add(cellKey);
        break;
      }
    }
    if (scanLimitReached) break;
  }
  const edgesByFrom = new Map();
  for (const edge of dependencyInspection.edges)
    edgesByFrom.set(edge.from, [
      ...(edgesByFrom.get(edge.from) ?? []),
      edge.to,
    ]);
  const dependsOnWithheld = (nodeKey, visited = new Set()) =>
    footprintBlankKeys.has(nodeKey)
      ? true
      : visited.has(nodeKey)
        ? false
        : (visited.add(nodeKey),
          (edgesByFrom.get(nodeKey) ?? []).some((childKey) =>
            dependsOnWithheld(childKey, visited),
          ));
  for (const row of inspection.rows) {
    if (row.status !== "unavailable" || !row.anchor) continue;
    const cell = nodesByKey.get(
      JSON.stringify([row.anchor.sheetId, row.anchor.address]),
    );
    if (!cell) {
      row.note =
        "Result withheld: this source is beyond the safe inspection budget.";
      continue;
    }
    if (
      ((row.raw = cell.raw),
      cell.resultStatus !== "unavailable" &&
        (dependsOnWithheld(
          JSON.stringify([row.anchor.sheetId, row.anchor.address]),
        ) ||
          ((scanLimitReached ||
            dependencyInspection.coverage.truncated ||
            dependencyInspection.coverage.depthLimited) &&
            (cell.raw === null ||
              (typeof cell.raw == "string" &&
                cell.raw.startsWith("=") &&
                cell.resultStatus === "calculated")))))
    ) {
      ((row.result = null),
        (row.resultStatus = "unavailable"),
        (row.status = "unavailable"),
        (row.note =
          "Result withheld: possible runtime spill output or an incomplete dependency inspection. Watch the stored source formula or narrow the workbook context."));
      continue;
    }
    ((row.result = cell.result),
      (row.resultStatus = cell.resultStatus),
      (row.status =
        cell.resultStatus === "unavailable" ? "unavailable" : "available"),
      cell.resultNote && (row.note = cell.resultNote));
  }
  return (
    (footprintBlankKeys.size || scanLimitReached) &&
      inspection.notes.push(
        "Blank cells inside a possible formula output footprint are conservatively withheld, as are formulas depending on them. Runtime spill-child identification remains outside the supported watch scope.",
      ),
    dependencyInspection.coverage.truncated &&
      inspection.notes.push(
        "The shared dependency scan reached a limit. Some results may be withheld; no incomplete result is presented as complete.",
      ),
    inspection
  );
}
function navigateToFormulaWatch(workbook, watchId, options = {}) {
  const watch = validateFormulaWatches(workbook.formulaWatches, workbook).find(
    (entry) => entry.id === watchId,
  );
  if (
    !watch ||
    classifyWatchedCell(
      workbook,
      watch,
      normalizeFormulaWatchContext(workbook, options),
    ) !== "ready"
  )
    throw Error(
      "This watch cannot be navigated in the current visible, authorized context.",
    );
  return structuredClone(watch.anchor);
}
const rj = (value) =>
  value === undefined
    ? "Unavailable"
    : value === null
      ? "(blank)"
      : typeof value == "boolean"
        ? String(value).toUpperCase()
        : String(value);
function FormulaWatchWindow(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [checked, setChecked] = ne.useState(false),
    [ids, setSelectedIds] = ne.useState([]),
    [preview, setPreview] = ne.useState(null),
    [actionError, setActionError] = ne.useState(""),
    [previewSelectionKey, setPreviewSelectionKey] = ne.useState(""),
    watchContext = {
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: props.allowedSheetIds,
    },
    selectionKey = JSON.stringify([
      props.sheetId,
      props.ranges ?? [props.range],
    ]),
    inspectionState = ne.useMemo(() => {
      try {
        return {
          result: inspectFormulaWatches(props.workbook, {
            includeHidden: !!props.allowHidden && checked,
            allowedSheetIds: props.allowedSheetIds,
          }),
        };
      } catch (error) {
        return {
          error:
            error instanceof Error
              ? error.message
              : "The watch list could not be read.",
        };
      }
    }, [props.workbook, props.allowHidden, props.allowedSheetIds, checked]);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('[data-ui-label="Close Watch Window"]')
        ?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const reportError = (failure) =>
      setActionError(
        failure instanceof Error
          ? failure.message
          : "The watch action could not be completed.",
      ),
    runWatchAction = (request) => {
      try {
        const actionPreview = applyFormulaWatchAction(props.workbook, request, {
          ...watchContext,
          readOnly: props.readOnly,
        });
        (setPreview(actionPreview),
          setPreviewSelectionKey(selectionKey),
          setActionError(""));
      } catch (error) {
        (setPreview(null), reportError(error));
      }
    },
    onClick = () => {
      if (preview)
        try {
          if (
            preview.request.kind === "add" &&
            previewSelectionKey !== selectionKey
          )
            throw Error(
              "The worksheet selection changed after the watch preview. Review the selection again.",
            );
          const nextWorkbook = assertWatchPreviewCurrent(
            props.workbook,
            preview,
            {
              ...watchContext,
              readOnly: props.readOnly,
            },
          );
          (nextWorkbook !== props.workbook &&
            props.onWorkbookChange(
              nextWorkbook,
              preview.request.kind === "add"
                ? `Add ${preview.added.length} formula watch${preview.added.length === 1 ? "" : "es"}`
                : preview.request.kind === "clear"
                  ? "Clear formula watches"
                  : `Remove ${preview.removed.length} formula watch${preview.removed.length === 1 ? "" : "es"}`,
            ),
            setPreview(null),
            setSelectedIds([]),
            setActionError(""));
        } catch (error) {
          reportError(error);
        }
    },
    navigateToWatch = (watchId) => {
      try {
        const anchor = navigateToFormulaWatch(
          props.workbook,
          watchId,
          watchContext,
        );
        (props.onNavigate(anchor.sheetId, anchor.address), setActionError(""));
      } catch (error) {
        reportError(error);
      }
    },
    watchRows = inspectionState.result?.rows ?? [],
    isPreviewCurrent =
      !!preview &&
      preview.snapshot === JSON.stringify(props.workbook) &&
      preview.context.includeHidden === watchContext.includeHidden &&
      JSON.stringify(preview.context.allowedSheetIds) ===
        JSON.stringify(
          [
            ...new Set(
              watchContext.allowedSheetIds ??
                props.workbook.sheets.map((sheet) => sheet.id),
            ),
          ].sort(),
        ) &&
      (preview.request.kind !== "add" || previewSelectionKey === selectionKey);
  return i.jsx("div", {
    className: "fw-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "fw-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${titleId}-title`,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              'button:not(:disabled),input:not(:disabled),summary,[tabindex="0"]',
            ) ?? []),
          ].filter((element) => !element.hidden);
          event.shiftKey && document.activeElement === focusableElements[0]
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
                i.jsx("h2", {
                  id: `${titleId}-title`,
                  children: i.jsx(LocalizedText, { children: "Watch Window" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Keep key cells and calculations in view across worksheets.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Watch Window"),
              onClick: props.onClose,
              lang: language("Close Watch Window"),
              "data-ui-label": "Close Watch Window",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "fw-toolbar",
          children: [
            i.jsx("button", {
              disabled:
                props.readOnly ||
                !props.workbook.sheets.some(
                  (sheet) => sheet.id === props.sheetId,
                ),
              onClick: () =>
                runWatchAction({
                  kind: "add",
                  sheetId: props.sheetId,
                  ranges: props.ranges ?? [props.range],
                }),
              children: i.jsx(LocalizedText, { children: "Add selection" }),
            }),
            i.jsx("button", {
              disabled: props.readOnly || !ids.length,
              onClick: () => runWatchAction({ kind: "remove", ids: ids }),
              children: i.jsx(LocalizedText, { children: "Remove selected" }),
            }),
            i.jsx("button", {
              disabled: props.readOnly || !watchRows.length,
              onClick: () => runWatchAction({ kind: "clear" }),
              children: i.jsx(LocalizedText, { children: "Clear watches" }),
            }),
            props.allowHidden &&
              i.jsxs("label", {
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (setChecked(event.target.checked),
                        setPreview(null),
                        setActionError(""));
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden" }),
                ],
              }),
          ],
        }),
        i.jsxs("p", {
          className: "fw-context",
          children: [
            props.workbook.name,
            " ",
            i.jsx(LocalizedText, { children: "· Revision " }),
            props.workbook.revision,
            " · ",
            watchRows.length,
            i.jsx(LocalizedText, { children: "/100 watches" }),
            props.readOnly ? " · Read-only inspection" : "",
          ],
        }),
        (inspectionState.error || actionError) &&
          i.jsx("p", {
            role: "alert",
            className: "fw-error",
            children: actionError || inspectionState.error,
          }),
        i.jsxs("div", {
          className: "fw-table-wrap",
          children: [
            i.jsxs("table", {
              children: [
                i.jsx("thead", {
                  children: i.jsxs("tr", {
                    children: [
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, { children: "Select" }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, {
                          children: "Worksheet / Cell",
                        }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, {
                          children: "Formula / Raw value",
                        }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, {
                          children: "Current result",
                        }),
                      }),
                      i.jsx("th", {
                        children: i.jsx(LocalizedText, { children: "Status" }),
                      }),
                    ],
                  }),
                }),
                i.jsx("tbody", {
                  children: watchRows.map((cell, rowIndex) =>
                    i.jsxs(
                      "tr",
                      {
                        children: [
                          i.jsx("td", {
                            children: i.jsx("input", {
                              type: "checkbox",
                              "aria-label": `Select watch ${rowIndex + 1}`,
                              checked: ids.includes(cell.id),
                              disabled: props.readOnly,
                              onChange: (event) =>
                                setSelectedIds((previousIds) =>
                                  event.target.checked
                                    ? [...previousIds, cell.id]
                                    : previousIds.filter(
                                        (selectedId) => selectedId !== cell.id,
                                      ),
                                ),
                            }),
                          }),
                          i.jsx("td", {
                            children: cell.anchor
                              ? i.jsxs("button", {
                                  className: "fw-link",
                                  disabled: cell.status === "orphaned",
                                  onClick: () => navigateToWatch(cell.id),
                                  "aria-label": `Go to ${cell.anchor.sheetName}!${cell.anchor.address}`,
                                  children: [
                                    cell.anchor.sheetName,
                                    "!",
                                    cell.anchor.address,
                                  ],
                                })
                              : i.jsxs("span", {
                                  children: [
                                    i.jsx(LocalizedText, {
                                      children: "Watch ",
                                    }),
                                    rowIndex + 1,
                                  ],
                                }),
                          }),
                          i.jsx("td", {
                            children: i.jsx("code", { children: rj(cell.raw) }),
                          }),
                          i.jsx("td", {
                            children:
                              cell.status === "available"
                                ? rj(cell.result)
                                : "Unavailable",
                          }),
                          i.jsxs("td", {
                            children: [
                              i.jsx("strong", {
                                children:
                                  cell.status === "available"
                                    ? cell.resultStatus === "literal"
                                      ? "Value"
                                      : "Calculated"
                                    : cell.status === "unavailable"
                                      ? "Result withheld"
                                      : cell.status[0].toUpperCase() +
                                        cell.status.slice(1),
                              }),
                              cell.note &&
                                i.jsx("small", { children: cell.note }),
                            ],
                          }),
                        ],
                      },
                      cell.id,
                    ),
                  ),
                }),
              ],
            }),
            !watchRows.length &&
              !inspectionState.error &&
              i.jsx("p", {
                children: i.jsx(LocalizedText, {
                  children:
                    "No watched cells. Select cells in the worksheet and choose Add selection.",
                }),
              }),
          ],
        }),
        preview &&
          i.jsxs("section", {
            className: "fw-preview",
            "aria-label": text("Watch action preview"),
            lang: language("Watch action preview"),
            "data-ui-label": "Watch action preview",
            children: [
              i.jsx("h3", {
                children:
                  preview.request.kind === "add"
                    ? "Add watches"
                    : preview.request.kind === "clear"
                      ? "Clear watch list"
                      : "Remove watches",
              }),
              i.jsxs("p", {
                children: [
                  preview.added.length,
                  " ",
                  i.jsx(LocalizedText, { children: "added · " }),
                  preview.removed.length,
                  " ",
                  i.jsx(LocalizedText, {
                    children:
                      "removed. Cell content and calculation formulas remain unchanged.",
                  }),
                ],
              }),
              !isPreviewCurrent &&
                i.jsx("p", {
                  role: "status",
                  children: i.jsx(LocalizedText, {
                    children:
                      "This preview is stale. Locations are withheld until you review the current selection and access context again.",
                  }),
                }),
              i.jsx("ul", {
                children: [
                  ...preview.added.map((watch) => ({
                    watch: watch,
                    action: "Add",
                  })),
                  ...preview.removed.map((watch) => ({
                    watch: watch,
                    action: "Remove",
                  })),
                ].map(({ watch, action }) => {
                  const currentRow = watchRows.find(
                      (watchRow) => watchRow.id === watch.id,
                    ),
                    watchLocation =
                      isPreviewCurrent &&
                      (action === "Add" || currentRow?.anchor)
                        ? `${watch.anchor.sheetName}!${watch.anchor.address}`
                        : "Watch outside current context";
                  return i.jsxs(
                    "li",
                    {
                      children: [
                        action,
                        ": ",
                        watchLocation,
                        watch.orphaned ? ` (${watch.orphaned})` : "",
                      ],
                    },
                    watch.id,
                  );
                }),
              }),
              !preview.added.length &&
                !preview.removed.length &&
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "The selected cells are already watched. No change is needed.",
                  }),
                }),
              i.jsx("button", {
                onClick: () => setPreview(null),
                children: i.jsx(LocalizedText, {
                  children: "Cancel watch action",
                }),
              }),
              i.jsx("button", {
                className: "fw-primary",
                disabled: props.readOnly,
                onClick: onClick,
                children: i.jsx(LocalizedText, {
                  children: "Apply watch action",
                }),
              }),
            ],
          }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Calculation scope and limits",
              }),
            }),
            inspectionState.result?.notes.map((note) =>
              i.jsx("p", { children: note }, note),
            ),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Results share one bounded inspection pass across the watch list. Unsupported references or sources beyond the safe dependency budget display Result withheld. Orphaned entries retain their former source and never reconnect automatically.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: props.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
const goalSeekLimits = {
    maxIterations: 100,
    maxEvaluations: 103,
    maxDependencyCells: 500,
    maxWorkUnits: 5e4,
    maxDepth: 8,
    maxFormulaCells: 5e3,
  },
  isCellInRange = (rangeText, cellAddress) => {
    const rangeAddress = parseRangeAddress(rangeText),
      position = _t(cellAddress);
    return (
      position.row >= rangeAddress.firstRow &&
      position.row <= rangeAddress.lastRow &&
      position.col >= rangeAddress.firstCol &&
      position.col <= rangeAddress.lastCol
    );
  };
function parseGoalSeekAddress(address) {
  if (
    typeof address != "string" ||
    !/^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}$/.test(address)
  )
    throw Error(
      "Goal Seek needs one valid Excel cell address for each source.",
    );
  const position = _t(address.replace(/\$/g, "").toUpperCase());
  if (
    position.row >= workbookLimits.maxRows ||
    position.col >= workbookLimits.maxCols
  )
    throw Error("A Goal Seek cell is outside Excel's address bounds.");
  return `${columnIndexToLetters(position.col)}${position.row + 1}`;
}
function normalizeGoalSeekRequest(request) {
  if (
    !request ||
    typeof request != "object" ||
    Object.keys(request).some(
      (key) =>
        ![
          "formulaSheetId",
          "formulaAddress",
          "variableSheetId",
          "variableAddress",
          "target",
          "lowerBound",
          "upperBound",
          "tolerance",
          "maxIterations",
        ].includes(key),
    )
  )
    throw Error("Invalid Goal Seek request.");
  if (
    typeof request.formulaSheetId != "string" ||
    typeof request.variableSheetId != "string"
  )
    throw Error("Choose a formula worksheet and variable worksheet.");
  const normalizedRequest = {
    ...request,
    formulaAddress: parseGoalSeekAddress(request.formulaAddress),
    variableAddress: parseGoalSeekAddress(request.variableAddress),
    tolerance: request.tolerance ?? 1e-6,
    maxIterations: request.maxIterations ?? 64,
  };
  if (
    !isFiniteNumber(normalizedRequest.target) ||
    !isFiniteNumber(normalizedRequest.lowerBound) ||
    !isFiniteNumber(normalizedRequest.upperBound) ||
    normalizedRequest.lowerBound >= normalizedRequest.upperBound
  )
    throw Error(
      "Target and bounds must be finite numbers, with lower bound smaller than upper bound.",
    );
  if (
    !isFiniteNumber(normalizedRequest.tolerance) ||
    normalizedRequest.tolerance <= 0
  )
    throw Error("Tolerance must be a finite positive absolute error.");
  if (
    !Number.isInteger(normalizedRequest.maxIterations) ||
    normalizedRequest.maxIterations < 1 ||
    normalizedRequest.maxIterations > goalSeekLimits.maxIterations
  )
    throw Error("Goal Seek supports from 1 to 100 iterations.");
  return normalizedRequest;
}
function solveGoalSeek(workbook, request, options = {}) {
  if (options.readOnly)
    throw Error(
      "This workbook is read-only. Goal Seek cannot propose a writable input.",
    );
  const goalSeekRequest = normalizeGoalSeekRequest(request),
    sheet = workbook.sheets.find(
      (sheet) => sheet.id === goalSeekRequest.formulaSheetId,
    ),
    variableSheet = workbook.sheets.find(
      (sheet) => sheet.id === goalSeekRequest.variableSheetId,
    );
  if (!sheet || !variableSheet)
    throw Error("A selected Goal Seek worksheet no longer exists.");
  if (
    options.includeHidden !== undefined &&
    typeof options.includeHidden != "boolean"
  )
    throw Error("Include hidden must be boolean.");
  if (
    options.allowedSheetIds !== undefined &&
    (!Array.isArray(options.allowedSheetIds) ||
      options.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("Goal Seek host scope contains an unknown worksheet.");
  const context = {
    includeHidden: !!options.includeHidden,
    allowedSheetIds: [
      ...new Set(
        options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
  };
  for (const [sourceSheet, sourceAddress] of [
    [sheet, goalSeekRequest.formulaAddress],
    [variableSheet, goalSeekRequest.variableAddress],
  ]) {
    const sourcePosition = _t(sourceAddress);
    if (
      !context.allowedSheetIds.includes(sourceSheet.id) ||
      (!context.includeHidden &&
        (sourceSheet.hidden ||
          sourceSheet.veryHidden ||
          cn(sourceSheet, "row").some(
            (rowRange) =>
              rowRange.start <= sourcePosition.row &&
              rowRange.end >= sourcePosition.row,
          ) ||
          cn(sourceSheet, "column").some(
            (columnRange) =>
              columnRange.start <= sourcePosition.col &&
              columnRange.end >= sourcePosition.col,
          )))
    )
      throw Error(
        "A Goal Seek source is outside the visible, authorized host context.",
      );
  }
  const cell = sheet.cells[goalSeekRequest.formulaAddress],
    variableCell = variableSheet.cells[goalSeekRequest.variableAddress];
  if (
    !cell ||
    cell.literal ||
    typeof cell.raw != "string" ||
    !cell.raw.startsWith("=")
  )
    throw Error("Set cell must contain an actual calculation formula.");
  if (!variableCell || !isFiniteNumber(variableCell.raw))
    throw Error(
      "Changing cell must contain a finite numeric input, not a formula, text, blank or spill output.",
    );
  if (variableSheet.protection?.enabled)
    throw Error(
      "The changing worksheet is protected. Unprotect it before Goal Seek.",
    );
  if (
    variableSheet.merges?.some((merge) =>
      isCellInRange(merge, goalSeekRequest.variableAddress),
    )
  )
    throw Error("A merged cell cannot be the Goal Seek variable.");
  for (const table of variableSheet.tables ?? []) {
    if (!isCellInRange(table.range, goalSeekRequest.variableAddress)) continue;
    const rangeAddress = parseRangeAddress(table.range),
      variablePosition = _t(goalSeekRequest.variableAddress),
      tableColumn = table.columns[variablePosition.col - rangeAddress.firstCol],
      calculatedColumnFormula = tableColumn?.calculatedFormula;
    if (
      variablePosition.row === rangeAddress.firstRow ||
      (table.showTotals && variablePosition.row === rangeAddress.lastRow) ||
      calculatedColumnFormula ||
      Object.entries(variableSheet.cells).some(([cellAddress, cell]) => {
        const cellPosition = _t(cellAddress);
        return (
          cellPosition.col === variablePosition.col &&
          cellPosition.row > rangeAddress.firstRow &&
          cellPosition.row <=
            rangeAddress.lastRow - (table.showTotals ? 1 : 0) &&
          !cell.literal &&
          typeof cell.raw == "string" &&
          cell.raw.startsWith("=")
        );
      })
    )
      throw Error(
        "A table header, totals row or calculated column cannot be the Goal Seek variable.",
      );
  }
  if (
    [...(variableSheet.pivots ?? []), ...(variableSheet.queries ?? [])].some(
      (pivotOrQuery) =>
        isCellInRange(
          pivotOrQuery.outputRange,
          goalSeekRequest.variableAddress,
        ),
    )
  )
    throw Error(
      "A generated pivot/query output cannot be the Goal Seek variable.",
    );
  if (
    variableCell.raw < goalSeekRequest.lowerBound ||
    variableCell.raw > goalSeekRequest.upperBound
  )
    throw Error("The starting input must be inside the supplied bounds.");
  const variableValidations = (variableSheet.validations ?? []).filter(
    (validationRule) =>
      isCellInRange(validationRule.range, goalSeekRequest.variableAddress),
  );
  if (
    variableValidations.some(
      (unsupportedRule) => !["number", "date"].includes(unsupportedRule.kind),
    )
  )
    throw Error(
      "Goal Seek supports continuous numeric/date validation only; discrete list, integer or text-length constraints need a different input.",
    );
  if (
    variableValidations.some(
      (rule) =>
        ![
          variableCell.raw,
          goalSeekRequest.lowerBound,
          goalSeekRequest.upperBound,
        ].every((value) => validateDataValidationRule(rule, value)),
    )
  )
    throw Error(
      "The starting input and both bounds must satisfy the variable's data validation rules.",
    );
  const dependencyReport = inspectFormulaDependencies(
      {
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          sheet.veryHidden ? { ...sheet, hidden: true } : sheet,
        ),
      },
      {
        sheetId: sheet.id,
        address: goalSeekRequest.formulaAddress,
        direction: "precedents",
        depth: goalSeekLimits.maxDepth,
        includeHidden: context.includeHidden,
        allowedSheetIds: context.allowedSheetIds,
        limits: {
          maxFormulaCells: goalSeekLimits.maxFormulaCells,
          maxNodes: goalSeekLimits.maxDependencyCells,
          maxReferences: 1e4,
        },
      },
    ),
    dependencyNodes = [dependencyReport.target, ...dependencyReport.nodes],
    dependencyKeys = new Set(dependencyNodes.map((event) => event.key));
  if (
    dependencyNodes.length > goalSeekLimits.maxDependencyCells ||
    dependencyReport.coverage.truncated ||
    dependencyReport.coverage.depthLimited
  )
    throw Error(
      "Goal Seek requires a complete dependency closure within 500 cells, 5,000 scanned formulas and depth 8.",
    );
  const issue = dependencyReport.issues.find((issue) =>
    dependencyKeys.has(issue.from),
  );
  if (issue || dependencyReport.cycles.length)
    throw Error(
      `Goal Seek requires supported static dependencies: ${issue?.message ?? "the dependency graph contains a cycle."}`,
    );
  if (
    !dependencyKeys.has(
      buildCompositeKey(variableSheet.id, goalSeekRequest.variableAddress),
    )
  )
    throw Error(
      "The Set cell formula does not depend on the selected changing cell.",
    );
  const unsupportedFunctionNames = new Set([
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
  for (const dependencyNode of dependencyNodes) {
    const sheet = workbook.sheets.find(
        (sheet) => sheet.id === dependencyNode.sheetId,
      ),
      cell = sheet.cells[dependencyNode.address];
    if (cell?.raw === null || !cell)
      throw Error(
        "Goal Seek requires populated static dependency cells; blank or possible spill-child dependencies are not supported.",
      );
    if (
      !cell.literal &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=")
    ) {
      const resolvedFormula = resolveDefinedNameValue(
        cell.raw,
        workbook,
        sheet.id,
      ).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'(?=\s*!)/g, "");
      if (
        [...resolvedFormula.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)]
          .map((match) =>
            match[1]
              .toUpperCase()
              .replace(/^_XLFN\./, "")
              .replace(/^_XLWS\./, ""),
          )
          .some((functionName) => unsupportedFunctionNames.has(functionName)) ||
        /[{}#]/.test(resolvedFormula)
      )
        throw Error(
          "Goal Seek does not run volatile, dynamic-reference, array/spill or unsupported reusable-function dependencies.",
        );
    }
  }
  const scopedWorkbook = {
    ...workbook,
    sheets: workbook.sheets
      .filter((sheet) => context.allowedSheetIds.includes(sheet.id))
      .map((allowedSheet) => ({
        ...allowedSheet,
        cells: Object.fromEntries(
          Object.entries(allowedSheet.cells).filter(([address]) =>
            dependencyKeys.has(buildCompositeKey(allowedSheet.id, address)),
          ),
        ),
      })),
  };
  let evaluations = 0,
    iterations = 0,
    workUnits = 0;
  const baseline = { input: variableCell.raw, output: null };
  let bestTrial = { input: variableCell.raw, value: null, residual: null };
  const trialCache = new Map();
  class k extends Error {
    constructor(failureStatus, message) {
      (super(message), (this.status = failureStatus));
    }
  }
  const evaluateTrial = (trialInput) => {
      const cachedTrial = trialCache.get(trialInput);
      if (cachedTrial) return cachedTrial;
      if (
        evaluations >= goalSeekLimits.maxEvaluations ||
        workUnits + dependencyNodes.length > goalSeekLimits.maxWorkUnits
      )
        throw new k(
          "work-limit",
          "The Goal Seek evaluation/work budget was reached.",
        );
      (evaluations++, (workUnits += dependencyNodes.length));
      const workbook = {
          ...scopedWorkbook,
          sheets: scopedWorkbook.sheets.map((sheet) =>
            sheet.id === variableSheet.id
              ? {
                  ...sheet,
                  cells: {
                    ...sheet.cells,
                    [goalSeekRequest.variableAddress]: {
                      ...variableCell,
                      raw: trialInput,
                    },
                  },
                }
              : sheet,
          ),
        },
        workbookWithSettings = calculateWorkbookWithSettings(workbook),
        formulaResult = workbookWithSettings(
          sheet.id,
          goalSeekRequest.formulaAddress,
        );
      if (workbookWithSettings.getOperationalErrors().length)
        throw new k(
          "calculation-error",
          `A trial encountered an operational calculation error (${workbookWithSettings.getOperationalErrors().join(", ")}), even if a formula error handler concealed it. The baseline is unchanged.`,
        );
      if (!isFiniteNumber(formulaResult))
        throw new k(
          "calculation-error",
          `The formula returned ${String(formulaResult)} at input ${trialInput}; a finite numeric result is required.`,
        );
      for (const dependency of dependencyNodes)
        if (
          workbookWithSettings.getSpillRange(
            dependency.sheetId,
            dependency.address,
          )
        )
          throw new k(
            "calculation-error",
            "A trial produced an array/spill result. Goal Seek requires scalar dependencies.",
          );
      if (workbookWithSettings.getOperationalErrors().length)
        throw new k(
          "calculation-error",
          "Dependency inspection encountered an operational calculation error; the baseline is unchanged.",
        );
      const residual = formulaResult - goalSeekRequest.target;
      if (!isFiniteNumber(residual))
        throw new k(
          "calculation-error",
          "The residual exceeded finite numeric precision.",
        );
      const trialResult = {
        input: trialInput,
        value: formulaResult,
        residual: residual,
      };
      return (
        trialCache.set(trialInput, trialResult),
        (bestTrial.residual === null ||
          Math.abs(residual) < Math.abs(bestTrial.residual)) &&
          (bestTrial = trialResult),
        trialResult
      );
    },
    buildOutcome = (outcomeStatus, outcomeMessage) => {
      const converged =
        outcomeStatus === "converged" || outcomeStatus === "already-satisfied";
      if (
        converged &&
        variableValidations.some(
          (validation) =>
            !validateDataValidationRule(validation, bestTrial.input),
        )
      )
        return buildOutcome(
          "validation-error",
          "The candidate violates the variable's data validation rule.",
        );
      const patches =
          converged && bestTrial.input !== variableCell.raw
            ? [
                {
                  sheetId: variableSheet.id,
                  address: goalSeekRequest.variableAddress,
                  cell: { ...variableCell, raw: bestTrial.input },
                },
              ]
            : [],
        next = patches.length
          ? {
              ...workbook,
              sheets: workbook.sheets.map((sheet) =>
                sheet.id === variableSheet.id
                  ? {
                      ...sheet,
                      cells: {
                        ...sheet.cells,
                        [goalSeekRequest.variableAddress]: patches[0].cell,
                      },
                    }
                  : sheet,
              ),
            }
          : workbook;
      return {
        snapshot: JSON.stringify(workbook),
        request: goalSeekRequest,
        context: context,
        baseline: baseline,
        dependencyCells: dependencyNodes.length,
        result: {
          status: outcomeStatus,
          converged: converged,
          input: bestTrial.input,
          achievedValue: bestTrial.value,
          residual: bestTrial.residual,
          absoluteResidual:
            bestTrial.residual === null ? null : Math.abs(bestTrial.residual),
          iterations: iterations,
          evaluations: evaluations,
          workUnits: workUnits,
          algorithm: "Bracketed bisection",
          message: outcomeMessage,
        },
        patches: patches,
        next: next,
      };
    };
  try {
    const baselineTrial = evaluateTrial(variableCell.raw);
    if (
      ((baseline.output = baselineTrial.value),
      Math.abs(baselineTrial.residual) <= goalSeekRequest.tolerance)
    )
      return buildOutcome(
        "already-satisfied",
        "The baseline already meets the requested tolerance. No input change is needed.",
      );
    let lowerTrial = evaluateTrial(goalSeekRequest.lowerBound);
    if (Math.abs(lowerTrial.residual) <= goalSeekRequest.tolerance)
      return buildOutcome(
        "converged",
        "The lower bound meets the requested tolerance.",
      );
    let upperTrial = evaluateTrial(goalSeekRequest.upperBound);
    if (Math.abs(upperTrial.residual) <= goalSeekRequest.tolerance)
      return buildOutcome(
        "converged",
        "The upper bound meets the requested tolerance.",
      );
    if (Math.sign(lowerTrial.residual) === Math.sign(upperTrial.residual))
      return buildOutcome(
        "not-bracketed",
        "The supplied bounds do not bracket a target crossing. A solution may exist elsewhere or inside a non-monotonic interval; choose a narrower valid bracket.",
      );
    for (
      iterations = 1;
      iterations <= goalSeekRequest.maxIterations;
      iterations++
    ) {
      const midpointInput = lowerTrial.input / 2 + upperTrial.input / 2;
      if (
        midpointInput === lowerTrial.input ||
        midpointInput === upperTrial.input
      )
        return buildOutcome(
          "precision-limit",
          "The bracket cannot be refined further at numeric precision; the residual still exceeds tolerance.",
        );
      const midpointTrial = evaluateTrial(midpointInput);
      if (Math.abs(midpointTrial.residual) <= goalSeekRequest.tolerance)
        return buildOutcome(
          "converged",
          "The candidate meets the requested absolute residual tolerance.",
        );
      Math.sign(midpointTrial.residual) === Math.sign(lowerTrial.residual)
        ? (lowerTrial = midpointTrial)
        : (upperTrial = midpointTrial);
    }
    return (
      (iterations = goalSeekRequest.maxIterations),
      buildOutcome(
        "iteration-limit",
        "The iteration limit was reached without meeting the residual tolerance. The baseline is unchanged.",
      )
    );
  } catch (error) {
    return buildOutcome(
      error instanceof k ? error.status : "calculation-error",
      error instanceof Error
        ? error.message
        : "The trial calculation failed; the baseline is unchanged.",
    );
  }
}
function applyGoalSeekResult(workbook, preview, options = {}) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (JSON.stringify(workbook) !== preview.snapshot)
    throw Error(
      "The workbook changed after the Goal Seek preview. Solve and review again.",
    );
  if (!preview.result.converged)
    throw Error(
      "Only a candidate that meets the residual tolerance can be applied.",
    );
  const freshPreview = solveGoalSeek(workbook, preview.request, options);
  if (JSON.stringify(freshPreview) !== JSON.stringify(preview))
    throw Error(
      "The Goal Seek preview or host context changed. Solve and review again.",
    );
  return freshPreview.next;
}
function GoalSeekDialog(props) {
  const workbookMessages = useWorkbookMessages(),
    [entryLocale] = ne.useState(workbookMessages.entryLocale),
    parseNumberEntry = (entryValue, fieldLabel) =>
      validateLocalizedNumberEntry(entryValue, fieldLabel, entryLocale),
    { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const [formulaSheetId, setFormulaSheetId] = ne.useState(props.sheetId),
    [formulaAddress, setFormulaAddress] = ne.useState(props.address),
    [variableSheetId, setVariableSheetId] = ne.useState(props.sheetId),
    [variableAddress, setVariableAddress] = ne.useState(""),
    [targetValue, setTargetValue] = ne.useState(""),
    [lowerBound, setLowerBound] = ne.useState("0"),
    [upperBound, setUpperBound] = ne.useState("100"),
    [tolerance, setTolerance] = ne.useState(() =>
      formatCellEntryForEdit({ raw: 1e-6 }, entryLocale),
    ),
    [maxIterations, setMaxIterations] = ne.useState("64"),
    [checked, setIncludeHidden] = ne.useState(false),
    [preview, setPreview] = ne.useState(null),
    [errorMessage, setErrorMessage] = ne.useState(""),
    hostScope = {
      includeHidden: !!props.allowHidden && checked,
      allowedSheetIds: props.allowedSheetIds,
    },
    allowedSheetIds = [
      ...new Set(
        props.allowedSheetIds ?? props.workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
    selectableSheets = props.workbook.sheets.filter(
      (sheet) =>
        allowedSheetIds.includes(sheet.id) &&
        (hostScope.includeHidden || (!sheet.hidden && !sheet.veryHidden)),
    ),
    hasFormulaSheet = selectableSheets.some(
      (formulaSheetOption) => formulaSheetOption.id === formulaSheetId,
    ),
    hasVariableSheet = selectableSheets.some(
      (variableSheetOption) => variableSheetOption.id === variableSheetId,
    ),
    isPreviewCurrent =
      !!preview &&
      preview.snapshot === JSON.stringify(props.workbook) &&
      preview.context.includeHidden === hostScope.includeHidden &&
      JSON.stringify(preview.context.allowedSheetIds) ===
        JSON.stringify(allowedSheetIds);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector('[data-ui-label="Set cell"]')?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  const updateField = (setValue, nextValue) => {
      (setValue(nextValue), setPreview(null), setErrorMessage(""));
    },
    showError = (failure) =>
      setErrorMessage(
        failure instanceof Error
          ? failure.message
          : "Goal Seek could not complete the operation.",
      ),
    onClick = () => {
      try {
        const request = {
          formulaSheetId: formulaSheetId,
          formulaAddress: formulaAddress,
          variableSheetId: variableSheetId,
          variableAddress: variableAddress,
          target: parseNumberEntry(targetValue, "Target value"),
          lowerBound: parseNumberEntry(lowerBound, "Lower bound"),
          upperBound: parseNumberEntry(upperBound, "Upper bound"),
          tolerance: parseNumberEntry(tolerance, "Tolerance"),
          maxIterations: parseNumberEntry(maxIterations, "Maximum iterations"),
        };
        (setPreview(
          solveGoalSeek(props.workbook, request, {
            ...hostScope,
            readOnly: props.readOnly,
          }),
        ),
          setErrorMessage(""));
      } catch (error) {
        (setPreview(null), showError(error));
      }
    },
    X$1 = () => {
      if (preview)
        try {
          const appliedWorkbook = applyGoalSeekResult(props.workbook, preview, {
            ...hostScope,
            readOnly: props.readOnly,
          });
          (appliedWorkbook !== props.workbook &&
            props.onWorkbookChange(
              appliedWorkbook,
              `Goal Seek: ${preview.request.variableAddress} → ${preview.result.input}`,
            ),
            setPreview(null),
            setErrorMessage(""));
        } catch (error) {
          showError(error);
        }
    };
  return i.jsx("div", {
    className: "gs-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "gs-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), onCloseRef.current()),
          event.key === "Tab")
        ) {
          const focusables = [
            ...(dialogRef.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),summary",
            ) ?? []),
          ].filter((element) => !element.hidden);
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
                i.jsx("h2", {
                  id: `${dialogId}-title`,
                  children: i.jsx(LocalizedText, { children: "Goal Seek" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Find one numeric input that brings a formula to a target value.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close Goal Seek"),
              onClick: props.onClose,
              lang: language("Close Goal Seek"),
              "data-ui-label": "Close Goal Seek",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsx("p", {
          children: i.jsx(LocalizedText, {
            children:
              "Enter a lower and upper bound around the starting input. Solve previews a candidate; Apply changes the workbook only after review.",
          }),
        }),
        i.jsxs("div", {
          className: "gs-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Formula worksheet" }),
                i.jsxs("select", {
                  "aria-label": text("Formula worksheet"),
                  value: hasFormulaSheet ? formulaSheetId : "",
                  disabled: props.readOnly,
                  onChange: (event) =>
                    updateField(setFormulaSheetId, event.target.value),
                  lang: language("Formula worksheet"),
                  "data-ui-label": "Formula worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    selectableSheets.map((formulaSheetChoice) =>
                      i.jsx(
                        "option",
                        {
                          value: formulaSheetChoice.id,
                          children: formulaSheetChoice.name,
                        },
                        formulaSheetChoice.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Set cell" }),
                i.jsx("input", {
                  "aria-label": text("Set cell"),
                  value: hasFormulaSheet ? formulaAddress : "",
                  placeholder: text("B1"),
                  readOnly: props.readOnly,
                  onChange: (event) =>
                    updateField(setFormulaAddress, event.target.value),
                  lang: language("Set cell"),
                  "data-ui-label": "Set cell",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Changing worksheet" }),
                i.jsxs("select", {
                  "aria-label": text("Changing worksheet"),
                  value: hasVariableSheet ? variableSheetId : "",
                  disabled: props.readOnly,
                  onChange: (event) =>
                    updateField(setVariableSheetId, event.target.value),
                  lang: language("Changing worksheet"),
                  "data-ui-label": "Changing worksheet",
                  children: [
                    i.jsx("option", {
                      value: "",
                      disabled: true,
                      children: "Choose worksheet",
                    }),
                    selectableSheets.map((variableSheetChoice) =>
                      i.jsx(
                        "option",
                        {
                          value: variableSheetChoice.id,
                          children: variableSheetChoice.name,
                        },
                        variableSheetChoice.id,
                      ),
                    ),
                  ],
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "By changing cell" }),
                i.jsx("input", {
                  "aria-label": text("By changing cell"),
                  value: hasVariableSheet ? variableAddress : "",
                  placeholder: text("A1"),
                  readOnly: props.readOnly,
                  onChange: (event) =>
                    updateField(setVariableAddress, event.target.value),
                  lang: language("By changing cell"),
                  "data-ui-label": "By changing cell",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "To value" }),
                i.jsx("input", {
                  "aria-label": text("Target value"),
                  value: targetValue,
                  inputMode: "decimal",
                  readOnly: props.readOnly,
                  onChange: (event) =>
                    updateField(setTargetValue, event.target.value),
                  lang: language("Target value"),
                  "data-ui-label": "Target value",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Absolute tolerance" }),
                i.jsx("input", {
                  "aria-label": text("Absolute tolerance"),
                  value: tolerance,
                  inputMode: "decimal",
                  readOnly: props.readOnly,
                  onChange: (event) =>
                    updateField(setTolerance, event.target.value),
                  lang: language("Absolute tolerance"),
                  "data-ui-label": "Absolute tolerance",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Lower bound" }),
                i.jsx("input", {
                  "aria-label": text("Lower bound"),
                  value: lowerBound,
                  inputMode: "decimal",
                  readOnly: props.readOnly,
                  onChange: (event) =>
                    updateField(setLowerBound, event.target.value),
                  lang: language("Lower bound"),
                  "data-ui-label": "Lower bound",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Upper bound" }),
                i.jsx("input", {
                  "aria-label": text("Upper bound"),
                  value: upperBound,
                  inputMode: "decimal",
                  readOnly: props.readOnly,
                  onChange: (event) =>
                    updateField(setUpperBound, event.target.value),
                  lang: language("Upper bound"),
                  "data-ui-label": "Upper bound",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Maximum iterations" }),
                i.jsx("input", {
                  "aria-label": text("Maximum iterations"),
                  value: maxIterations,
                  inputMode: "numeric",
                  readOnly: props.readOnly,
                  onChange: (event) =>
                    updateField(setMaxIterations, event.target.value),
                  lang: language("Maximum iterations"),
                  "data-ui-label": "Maximum iterations",
                }),
              ],
            }),
            props.allowHidden &&
              i.jsxs("label", {
                className: "gs-checkbox",
                children: [
                  i.jsx("input", {
                    type: "checkbox",
                    checked: checked,
                    onChange: (event) => {
                      (setIncludeHidden(event.target.checked),
                        setPreview(null),
                        setErrorMessage(""));
                    },
                  }),
                  i.jsx(LocalizedText, { children: "Include hidden sources" }),
                ],
              }),
          ],
        }),
        errorMessage &&
          i.jsx("p", {
            role: "alert",
            className: "gs-error",
            children: errorMessage,
          }),
        i.jsx("button", {
          className: "gs-primary",
          disabled: props.readOnly || !hasFormulaSheet || !hasVariableSheet,
          onClick: onClick,
          children: i.jsx(LocalizedText, { children: "Solve" }),
        }),
        preview &&
          i.jsxs("section", {
            className: "gs-preview",
            "aria-label": text("Goal Seek result"),
            lang: language("Goal Seek result"),
            "data-ui-label": "Goal Seek result",
            children: [
              isPreviewCurrent
                ? i.jsxs(i.Fragment, {
                    children: [
                      i.jsx("h3", {
                        children: preview.result.converged
                          ? preview.result.status === "already-satisfied"
                            ? "Target already satisfied"
                            : "Candidate found"
                          : "No applicable solution",
                      }),
                      i.jsx("p", {
                        role: "status",
                        children: preview.result.message,
                      }),
                      i.jsxs("dl", {
                        children: [
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Formula",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              props.workbook.sheets.find(
                                (sheet) =>
                                  sheet.id === preview.request.formulaSheetId,
                              )?.name,
                              "!",
                              preview.request.formulaAddress,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Changing input",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              props.workbook.sheets.find(
                                (sheet) =>
                                  sheet.id === preview.request.variableSheetId,
                              )?.name,
                              "!",
                              preview.request.variableAddress,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Baseline input",
                            }),
                          }),
                          i.jsx("dd", { children: preview.baseline.input }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Baseline result",
                            }),
                          }),
                          i.jsx("dd", {
                            children: preview.baseline.output ?? "Unavailable",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Target",
                            }),
                          }),
                          i.jsx("dd", { children: preview.request.target }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Candidate input",
                            }),
                          }),
                          i.jsx("dd", { children: preview.result.input }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Achieved value",
                            }),
                          }),
                          i.jsx("dd", {
                            children:
                              preview.result.achievedValue ?? "Unavailable",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Residual (result − target)",
                            }),
                          }),
                          i.jsx("dd", {
                            children: preview.result.residual ?? "Unavailable",
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Absolute tolerance",
                            }),
                          }),
                          i.jsx("dd", { children: preview.request.tolerance }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Iterations / Evaluations",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              preview.result.iterations,
                              " / ",
                              preview.result.evaluations,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Method",
                            }),
                          }),
                          i.jsx("dd", { children: preview.result.algorithm }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Search bounds",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              preview.request.lowerBound,
                              " ",
                              i.jsx(LocalizedText, { children: "to " }),
                              preview.request.upperBound,
                            ],
                          }),
                          i.jsx("dt", {
                            children: i.jsx(LocalizedText, {
                              children: "Dependency cells / Work units",
                            }),
                          }),
                          i.jsxs("dd", {
                            children: [
                              preview.dependencyCells,
                              " / ",
                              preview.result.workUnits,
                            ],
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "The baseline workbook has not changed.",
                          }),
                          !preview.result.converged &&
                            " Choose different bounds or settings and solve again; a failed candidate cannot be applied.",
                        ],
                      }),
                    ],
                  })
                : i.jsx("p", {
                    role: "status",
                    children: i.jsx(LocalizedText, {
                      children:
                        "This preview is stale. Its source values and locations are withheld. Solve again using the current workbook and access context.",
                    }),
                  }),
              i.jsx("button", {
                onClick: () => setPreview(null),
                children: i.jsx(LocalizedText, { children: "Discard result" }),
              }),
              i.jsx("button", {
                className: "gs-primary",
                disabled: props.readOnly || !preview.result.converged,
                onClick: X$1,
                children: i.jsx(LocalizedText, { children: "Apply result" }),
              }),
            ],
          }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Supported scope and limits",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "One continuous numeric input and one scalar formula, using deterministic supported dependencies. Volatile/dynamic references, array/spill outputs, blank dependency cells, discrete validation constraints and calculated table inputs are excluded. A valid bracket is required; this search does not find every possible root or prove a unique solution.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "At most 100 iterations, 103 trial evaluations, 500 dependency cells through depth 8, and 50,000 dependency-cell work units. Absolute residual tolerance determines success; a narrow bracket alone does not.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Include hidden is a local visibility choice within the host-authorized workbook. The embedding host owns approval, permissions, Undo and persistence.",
              }),
            }),
          ],
        }),
        i.jsx("footer", {
          children: i.jsx("button", {
            onClick: props.onClose,
            children: i.jsx(LocalizedText, { children: "Close" }),
          }),
        }),
      ],
    }),
  });
}
function normalizeAnalysisPermissions(workbook, options = {}) {
  if (
    (options.readOnly !== undefined && typeof options.readOnly != "boolean") ||
    (options.includeHidden !== undefined &&
      typeof options.includeHidden != "boolean")
  )
    throw Error("Analysis permissions must be boolean.");
  if (
    options.allowedSheetIds !== undefined &&
    (!Array.isArray(options.allowedSheetIds) ||
      options.allowedSheetIds.some(
        (allowedSheetId) =>
          typeof allowedSheetId != "string" ||
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ))
  )
    throw Error("The host scope contains an unknown worksheet.");
  return {
    readOnly: !!options.readOnly,
    includeHidden: !!options.includeHidden,
    allowedSheetIds: [
      ...new Set(
        options.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id),
      ),
    ].sort(),
  };
}
function validateAnalysisSource(workbook, source, options) {
  const analysisPermissions = normalizeAnalysisPermissions(workbook, options),
    sheet = workbook.sheets.find((sheet) => sheet.id === source.sheetId);
  if (!sheet || !analysisPermissions.allowedSheetIds.includes(sheet.id))
    throw Error("A source is outside the authorized host context.");
  if (!/^[A-Z]{1,3}[1-9]\d{0,6}$/.test(source.address))
    throw Error("Use a canonical cell address.");
  const cellPosition = _t(source.address);
  if (cellPosition.row >= 1048576 || cellPosition.col >= 16384)
    throw Error("The cell address exceeds Excel dimensions.");
  if (
    !analysisPermissions.includeHidden &&
    (sheet.hidden ||
      sheet.veryHidden ||
      cn(sheet, "row").some(
        (rowRange) =>
          rowRange.start <= cellPosition.row &&
          cellPosition.row <= rowRange.end,
      ) ||
      cn(sheet, "column").some(
        (columnRange) =>
          columnRange.start <= cellPosition.col &&
          cellPosition.col <= columnRange.end,
      ))
  )
    throw Error(
      "A source is hidden. Explicitly include hidden cells within the authorized scope.",
    );
  if (sheet.analysisReport || sheet.dataModelReport)
    throw Error(
      "Analysis and model report snapshots cannot be analysis inputs. Use the original source cells.",
    );
  if (sheet.merges?.some((merge) => parseRangeBounds(merge, source.address)))
    throw Error("Merged analysis sources are ambiguous.");
  return sheet;
}
function parseRangeBounds(rangeText, address) {
  const rangeAddress = parseRangeAddress(rangeText),
    cellPosition = _t(address);
  return (
    cellPosition.row >= rangeAddress.firstRow &&
    cellPosition.row <= rangeAddress.lastRow &&
    cellPosition.col >= rangeAddress.firstCol &&
    cellPosition.col <= rangeAddress.lastCol
  );
}
function p_(workbook) {
  return fingerprintCanonicalValue({
    sheets: workbook.sheets
      .filter((sheet) => !sheet.analysisReport)
      .map((sheetCopy) => ({ ...sheetCopy })),
    calculation: workbook.calculation ?? null,
    datePolicy: workbook.datePolicy ?? null,
    definedNames: workbook.definedNames ?? null,
  });
}
function assertAutomaticCalculation(workbook) {
  if (
    workbook.calculation?.mode === "manual" ||
    workbook.calculation?.iterative
  )
    throw Error(
      "Analysis formulas require automatic, non-iterative calculation.",
    );
}
const rejectRangeAsScalar = (value) => {
    if (Array.isArray(value))
      throw Error("Use SUM, AVERAGE or SUMPRODUCT to reduce a range.");
    return value;
  },
  assertSolverMagnitude = (solverValue) => {
    if (!Number.isFinite(solverValue) || Math.abs(solverValue) > 1e100)
      throw Error("Solver arithmetic exceeds the supported finite magnitude.");
    return solverValue;
  };

export {
  assertSolverMagnitude,
  parseRangeBounds,
  FormulaWatchWindow,
  assertAutomaticCalculation,
  normalizeAnalysisPermissions,
  rejectRangeAsScalar,
  GoalSeekDialog,
  p_,
  validateAnalysisSource,
};
export { isCellInRange };
