import { formatCellAddress } from "./04-cell-input-paste-and-fill.js";
import {
  X,
  Plus,
  ArrowDown,
  ArrowUp,
  FileSpreadsheet,
  Copy,
  EyeOff,
  Eye as Eye$1,
  LockKeyhole,
  Trash2,
  LockKeyholeOpen,
} from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import {
  describeSheetDeletionImpacts,
  addWorksheet,
  setWorksheetTabColor,
  unprotectWorkbookStructure,
  findSheetOrThrow,
  renameWorksheet,
  setWorksheetVisibility,
  deleteWorksheet,
  protectWorkbookStructure,
  _t,
  cn,
  parseRangeAddress,
  assertSheetUnprotected,
  summarizeWorksheets,
  normalizeWorkbookInput,
  columnIndexToLetters,
  checkDataValidationBatch,
  moveWorksheet,
  duplicateWorksheet,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import { applyStructuralEdit } from "./27-validation-and-structural-edits.js";
import {
  assertSortWorkbookAuthorized,
  buildAllowedSheetKey,
  buildSortPlan,
  MW,
  N3,
  stableSortPolicy,
  resolveSortSourceRange,
  assertSortSafeFormulas,
  isNonBlankCellEntry,
  normalizeSortSpecification,
  sortEngineLimits,
} from "./44-clean-and-reshape.js";

function buildSortPreview(workbook, request, options = {}) {
  const targetSheet = assertSortWorkbookAuthorized(
      workbook,
      request.sheetId,
      options,
    ),
    sortSpecification = normalizeSortSpecification(workbook, request),
    sortSourceRange = resolveSortSourceRange(
      workbook,
      sortSpecification.sheetId,
      sortSpecification.range,
      sortSpecification.headers,
    );
  if (sortSourceRange.expansionRequired && !sortSpecification.expandSelection)
    throw Error(
      `This selection is part of populated records. Expand the selection to ${sortSourceRange.expandedRange}, or cancel; partial-record sorting is blocked.`,
    );
  const rangeAddress = parseRangeAddress(sortSourceRange.expandedRange),
    rowCount = sortSourceRange.lastRow - sortSourceRange.firstDataRow + 1;
  if (rowCount < 2)
    throw Error("Choose at least two data rows. Check the header setting.");
  if (
    rowCount > sortEngineLimits.rows ||
    rowCount * (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
      sortEngineLimits.cells
  )
    throw Error(
      "Sort supports at most 10,000 rows and 100,000 record cells per reviewed operation.",
    );
  const hiddenRowRanges = cn(targetSheet, "row"),
    hiddenRows = Array.from(
      { length: rowCount },
      (B, rowOffset) => sortSourceRange.firstDataRow + rowOffset,
    ).filter((rowNumber) =>
      hiddenRowRanges.some(
        (hiddenRange) =>
          rowNumber >= hiddenRange.start && rowNumber <= hiddenRange.end,
      ),
    ).length;
  if (hiddenRows && !sortSpecification.includeHidden)
    throw Error(
      `The source contains ${hiddenRows} hidden rows. Include hidden rows explicitly or cancel so records stay together.`,
    );
  assertSortSafeFormulas(workbook, rowCount, options);
  const sortPlan = buildSortPlan(workbook, sortSpecification, sortSourceRange),
    keys = JSON.stringify(sortPlan),
    collator = new Intl.Collator("en-US", {
      usage: "sort",
      sensitivity: sortSpecification.caseSensitive ? "variant" : "accent",
      numeric: false,
    }),
    comparableText = (text) =>
      sortSpecification.caseSensitive
        ? text.normalize("NFC")
        : text.normalize("NFC").toLocaleLowerCase("en-US"),
    customOrderMaps = sortSpecification.keys.map(
      (key) =>
        new Map(
          key.customOrder?.map((customOrderValue, customOrderIndex) => [
            comparableText(customOrderValue),
            customOrderIndex,
          ]),
        ),
    ),
    valueRank = (keyEntry) =>
      keyEntry.compare === null || keyEntry.compare === ""
        ? 4
        : keyEntry.error
          ? 3
          : typeof keyEntry.compare == "number"
            ? 0
            : typeof keyEntry.compare == "string"
              ? 1
              : 2,
    sortedRows = [...sortPlan].sort((leftRow, rightRow) => {
      for (
        let keyIndex = 0;
        keyIndex < sortSpecification.keys.length;
        keyIndex++
      ) {
        const leftKey = leftRow.keys[keyIndex],
          rightKey = rightRow.keys[keyIndex],
          leftRank = valueRank(leftKey),
          rightRank = valueRank(rightKey),
          directionSign =
            sortSpecification.keys[keyIndex].direction === "ascending" ? 1 : -1;
        if (leftRank === 4 || rightRank === 4) {
          if (leftRank !== rightRank) return leftRank === 4 ? 1 : -1;
          continue;
        }
        if (customOrderMaps[keyIndex].size) {
          const leftCustomIndex = customOrderMaps[keyIndex].get(
              comparableText(String(leftKey.compare)),
            ),
            rightCustomIndex = customOrderMaps[keyIndex].get(
              comparableText(String(rightKey.compare)),
            );
          if (leftCustomIndex !== undefined || rightCustomIndex !== undefined) {
            if (leftCustomIndex === undefined) return 1;
            if (rightCustomIndex === undefined) return -1;
            if (leftCustomIndex !== rightCustomIndex)
              return (leftCustomIndex - rightCustomIndex) * directionSign;
          }
        }
        const comparison =
          leftRank !== rightRank
            ? leftRank - rightRank
            : typeof leftKey.compare == "number" &&
                typeof rightKey.compare == "number"
              ? leftKey.compare - rightKey.compare
              : typeof leftKey.compare == "boolean" &&
                  typeof rightKey.compare == "boolean"
                ? Number(leftKey.compare) - Number(rightKey.compare)
                : collator.compare(
                    String(leftKey.compare),
                    String(rightKey.compare),
                  );
        if (comparison) return comparison * directionSign;
      }
      return leftRow.row - rightRow.row;
    }),
    permutation = Array(rowCount).fill(0);
  sortedRows.forEach(
    (sortedRow, sortedIndex) =>
      (permutation[sortedRow.row - sortSourceRange.firstDataRow] =
        sortSourceRange.firstDataRow + sortedIndex),
  );
  const movedRows = sortedRows.filter(
      (movedRow, positionIndex) =>
        movedRow.row !== sortSourceRange.firstDataRow + positionIndex,
    ).length,
    structuralEdit = {
      axis: "row",
      index: sortSourceRange.firstDataRow,
      count: rowCount,
      remove: false,
      moveBefore: sortSourceRange.firstDataRow,
      permutation: permutation,
    };
  let candidate = movedRows
    ? applyStructuralEdit(workbook, targetSheet.id, structuralEdit)
    : structuredClone(workbook);
  const changedCells = [],
    formulaChanges = [],
    metadataChanges = [],
    destinationRowOf = (rowIndex) =>
      rowIndex >= sortSourceRange.firstDataRow &&
      rowIndex <= sortSourceRange.lastRow
        ? permutation[rowIndex - sortSourceRange.firstDataRow]
        : rowIndex;
  for (const sheet of workbook.sheets) {
    const sheetOrThrow = findSheetOrThrow(candidate, sheet.id);
    for (const [source, cell] of Object.entries(sheet.cells))
      if (N3(cell)) {
        const sourceAddress = _t(source),
          destination =
            sheet.id === targetSheet.id
              ? formatCellAddress(
                  destinationRowOf(sourceAddress.row),
                  sourceAddress.col,
                )
              : source,
          candidateRaw = sheetOrThrow.cells[destination]?.raw;
        candidateRaw !== cell.raw &&
          formulaChanges.push({
            sheet: sheet.name,
            source: source,
            destination: destination,
            before: String(cell.raw),
            after: String(candidateRaw),
          });
      }
    for (const address of new Set([
      ...Object.keys(sheet.cells),
      ...Object.keys(sheetOrThrow.cells),
    ]))
      JSON.stringify(sheet.cells[address]) !==
        JSON.stringify(sheetOrThrow.cells[address]) &&
        changedCells.push({
          sheetId: sheet.id,
          address: address,
          cell: sheetOrThrow.cells[address] ?? { raw: null },
        });
    for (const field of Object.keys(sheetOrThrow).filter(
      (sheetFieldName) => sheetFieldName !== "cells",
    ))
      JSON.stringify(sheet[field]) !== JSON.stringify(sheetOrThrow[field]) &&
        metadataChanges.push({ sheet: sheet.name, field: field });
  }
  if (
    (JSON.stringify(workbook.definedNames) !==
      JSON.stringify(candidate.definedNames) &&
      metadataChanges.push({ sheet: "Workbook", field: "definedNames" }),
    JSON.stringify(workbook.formulaWatches) !==
      JSON.stringify(candidate.formulaWatches) &&
      metadataChanges.push({ sheet: "Workbook", field: "formulaWatches" }),
    changedCells.length > sortEngineLimits.cells)
  )
    throw Error(
      "The sort would change more than 100,000 cells including dependent references.",
    );
  JSON.stringify(workbook.reviewThreads) !==
    JSON.stringify(candidate.reviewThreads) &&
    metadataChanges.push({ sheet: "Workbook", field: "reviewThreads" });
  for (const sheet of workbook.sheets)
    sheet.id !== targetSheet.id &&
      JSON.stringify(sheet) !==
        JSON.stringify(findSheetOrThrow(candidate, sheet.id)) &&
      assertSheetUnprotected(sheet);
  if ((checkDataValidationBatch(candidate, changedCells), movedRows)) {
    const sheetOrThrow = findSheetOrThrow(candidate, targetSheet.id);
    ((candidate.rowOrigins = [...(candidate.rowOrigins ?? [])]),
      (sheetOrThrow.rowOriginIds = { ...sheetOrThrow.rowOriginIds }));
    const populatedRows = new Set(
      Object.entries(targetSheet.cells)
        .filter(([, cellEntry]) => isNonBlankCellEntry(cellEntry))
        .map(([cellAddress]) => _t(cellAddress).row),
    );
    for (const sortedRowEntry of sortedRows)
      if (
        populatedRows.has(sortedRowEntry.row) &&
        !targetSheet.rowOriginIds?.[sortedRowEntry.row]
      ) {
        if (candidate.rowOrigins.length >= 1e5)
          throw Error("Source row identity capacity has been reached.");
        const rowOriginId = crypto.randomUUID();
        ((sheetOrThrow.rowOriginIds[destinationRowOf(sortedRowEntry.row)] =
          rowOriginId),
          candidate.rowOrigins.push({
            id: rowOriginId,
            sourceId: workbook.id,
            sourceName: workbook.name,
            sourceSheet: targetSheet.name,
            originalRow: sortedRowEntry.row + 1,
            kind: "snapshot",
          }));
      }
    candidate = normalizeWorkbookInput(candidate);
  }
  const rowOriginsById = new Map(
      candidate.rowOrigins?.map((rowOrigin) => [rowOrigin.id, rowOrigin]),
    ),
    preview = {
      request: sortSpecification,
      selection: sortSourceRange,
      snapshot: JSON.stringify(workbook),
      rows: sortedRows.map((planRow, rowPosition) => {
        const sourceId = findSheetOrThrow(candidate, targetSheet.id)
            .rowOriginIds?.[sortSourceRange.firstDataRow + rowPosition],
          sourceOrigin = sourceId ? rowOriginsById.get(sourceId) : undefined;
        return {
          sourceRow: planRow.row + 1,
          destinationRow: sortSourceRange.firstDataRow + rowPosition + 1,
          ...(sourceId
            ? {
                sourceId: sourceId,
                sourceName: sourceOrigin?.sourceName,
                originalRow: sourceOrigin?.originalRow,
              }
            : {}),
          keys: planRow.keys.map((key) => key.value),
        };
      }),
      movedRows: movedRows,
      changedCells: changedCells.length,
      formulaChanges: formulaChanges,
      metadataChanges: metadataChanges,
      hiddenRows: hiddenRows,
      policy: stableSortPolicy,
    };
  return (
    MW.set(preview, {
      signature: JSON.stringify(preview),
      candidate: candidate,
      keys: keys,
      allowed: buildAllowedSheetKey(workbook, options),
    }),
    preview
  );
}
function assertSortPreviewCurrent(workbook, preview, options = {}) {
  const previewState = MW.get(preview);
  if (!previewState || previewState.signature !== JSON.stringify(preview))
    throw Error(
      "This sort preview changed or was restored from another session. Preview again.",
    );
  if (
    (assertSortWorkbookAuthorized(workbook, preview.request.sheetId, options),
    previewState.allowed !== buildAllowedSheetKey(workbook, options) ||
      JSON.stringify(workbook) !== preview.snapshot)
  )
    throw Error(
      "The workbook or host permissions changed after the sort preview. Preview again.",
    );
  if (
    previewState.keys !==
    JSON.stringify(buildSortPlan(workbook, preview.request, preview.selection))
  )
    throw Error(
      "Calculated sort keys changed after the preview. Preview and review the new order.",
    );
  return preview.movedRows ? structuredClone(previewState.candidate) : workbook;
}
const buildSortKeySpec = (column) => ({
    column: column,
    direction: "ascending",
    type: "auto",
    customText: "",
    useCustom: false,
  }),
  expandSortSelectionToRecords = (selection) => {
    try {
      const firstRow = parseRangeAddress(
        resolveSortSourceRange(
          selection.workbook,
          selection.sheetId,
          selection.range,
        ).expandedRange,
      ).firstRow;
      return (
        firstRow === 0 ||
        !!selection.workbook.sheets
          .find((sheet) => sheet.id === selection.sheetId)
          ?.tables?.some(
            (table) => parseRangeAddress(table.range).firstRow === firstRow,
          )
      );
    } catch {
      return false;
    }
  };
function CustomSortDialog(dialogProps) {
  const { text, language } = useWorkbookMessages(),
    dialogId = ne.useId(),
    dialogRef = ne.useRef(null),
    previewRef = ne.useRef(null),
    [range, setRange] = ne.useState(dialogProps.range),
    [headers, setHeaders] = ne.useState(() =>
      expandSortSelectionToRecords(dialogProps),
    ),
    [sortKeys, setSortKeys] = ne.useState([
      buildSortKeySpec(parseRangeAddress(dialogProps.range).firstCol),
    ]),
    [isSelectionExpanded, setIsSelectionExpanded] = ne.useState(false),
    [shouldIncludeHidden, setShouldIncludeHidden] = ne.useState(false),
    [isCaseSensitive, setIsCaseSensitive] = ne.useState(false),
    [sortPreview, setSortPreview] = ne.useState(null),
    [previewSnapshot, setPreviewSnapshot] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [pageIndex, setPageIndex] = ne.useState(0),
    sheet = dialogProps.workbook.sheets.find(
      (sheet) => sheet.id === dialogProps.sheetId,
    ),
    allowedSheetIds =
      dialogProps.allowedSheetIds ??
      dialogProps.workbook.sheets.map((sheet) => sheet.id),
    isSheetAllowed = allowedSheetIds.includes(dialogProps.sheetId),
    currentSnapshot = JSON.stringify([
      dialogProps.workbook,
      dialogProps.sheetId,
      dialogProps.range,
      dialogProps.readOnly,
      [...allowedSheetIds].sort(),
    ]),
    isPreviewCurrent =
      !!sortPreview &&
      previewSnapshot === currentSnapshot &&
      !dialogProps.readOnly;
  let sortSource,
    sourceError = "";
  try {
    if (!isSheetAllowed)
      throw Error("The source is outside the host-approved worksheets.");
    sortSource = resolveSortSourceRange(
      dialogProps.workbook,
      dialogProps.sheetId,
      range,
      headers,
    );
  } catch (error) {
    sourceError = error.message;
  }
  const expandedAddress = sortSource
      ? parseRangeAddress(sortSource.expandedRange)
      : undefined,
    columnIndexes = expandedAddress
      ? Array.from(
          { length: expandedAddress.lastCol - expandedAddress.firstCol + 1 },
          (X, columnOffset) => expandedAddress.firstCol + columnOffset,
        )
      : [],
    resetPreview = () => {
      (setSortPreview(null), setErrorMessage(""), setPageIndex(0));
    },
    updateSortKey = (keyIndex, changes) => {
      (resetPreview(),
        setSortKeys(
          sortKeys.map((sortKey, itemIndex) =>
            itemIndex === keyIndex ? { ...sortKey, ...changes } : sortKey,
          ),
        ));
    },
    moveSortKey = (keyIndex, direction) => {
      resetPreview();
      const nextSortKeys = [...sortKeys];
      (([nextSortKeys[keyIndex], nextSortKeys[keyIndex + direction]] = [
        nextSortKeys[keyIndex + direction],
        nextSortKeys[keyIndex],
      ]),
        setSortKeys(nextSortKeys));
    },
    onClick = () => {
      try {
        const sortPreview = buildSortPreview(
          dialogProps.workbook,
          {
            sheetId: dialogProps.sheetId,
            range: range,
            headers: headers,
            keys: sortKeys.map(({ customText, useCustom, ...keySpec }) => ({
              ...keySpec,
              ...(useCustom ? { customOrder: customText.split(/\r?\n/) } : {}),
            })),
            expandSelection: isSelectionExpanded,
            includeHidden: shouldIncludeHidden,
            caseSensitive: isCaseSensitive,
          },
          {
            readOnly: dialogProps.readOnly,
            allowedSheetIds: dialogProps.allowedSheetIds,
          },
        );
        (setSortPreview(sortPreview),
          setPreviewSnapshot(currentSnapshot),
          setPageIndex(0),
          setErrorMessage(""));
      } catch (error) {
        (setSortPreview(null), setErrorMessage(error.message));
      }
    },
    applySortPreview = () => {
      try {
        if (!sortPreview || !isPreviewCurrent)
          throw Error("The workbook or selection changed. Preview again.");
        const sortedWorkbook = assertSortPreviewCurrent(
          dialogProps.workbook,
          sortPreview,
          {
            readOnly: dialogProps.readOnly,
            allowedSheetIds: dialogProps.allowedSheetIds,
          },
        );
        (dialogProps.onWorkbookChange(
          sortedWorkbook,
          `Sort ${sheet?.name}: ${sortPreview.movedRows} records moved`,
        ),
          dialogProps.onClose());
      } catch (error) {
        (setSortPreview(null), setErrorMessage(error.message));
      }
    };
  (ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current
        ?.querySelector('input[data-ui-label="Source range"]')
        ?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []),
    ne.useEffect(() => {
      var previewElement;
      sortPreview &&
        (((previewElement = previewRef.current) == null
          ? undefined
          : previewElement.scrollIntoView
        )?.call(previewElement, {
          block: "start",
        }),
        previewRef.current?.focus({ preventScroll: true }));
    }, [sortPreview]));
  const pageCount = Math.max(
      1,
      Math.ceil((sortPreview?.rows.length ?? 0) / sortEngineLimits.pageSize),
    ),
    currentPage = Math.min(pageIndex, pageCount - 1);
  return i.jsx("div", {
    className: "dup-backdrop",
    onMouseDown: (event) => {
      event.target === event.currentTarget && dialogProps.onClose();
    },
    children: i.jsxs("section", {
      className: "dup-dialog sort-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": `${dialogId}-title`,
      ref: dialogRef,
      onKeyDown: (event) => {
        var lastFocusable, firstFocusable;
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(), dialogProps.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = [
            ...(dialogRef.current?.querySelectorAll(
              "button:not(:disabled),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),summary",
            ) ?? []),
          ].filter(
            (element) =>
              !element.closest("details:not([open])") ||
              element.tagName === "SUMMARY",
          );
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
                  id: `${dialogId}-title`,
                  children: i.jsx(LocalizedText, { children: "Custom sort" }),
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children: "Review the new record order before applying it.",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              "aria-label": text("Close custom sort"),
              onClick: dialogProps.onClose,
              lang: language("Close custom sort"),
              "data-ui-label": "Close custom sort",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "dup-fields",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Worksheet" }),
                i.jsx("strong", {
                  children: isSheetAllowed
                    ? sheet?.name
                    : "Outside current context",
                }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx(LocalizedText, { children: "Source range" }),
                i.jsx("input", {
                  "aria-label": text("Source range"),
                  value: range,
                  onChange: (event) => {
                    (resetPreview(),
                      setIsSelectionExpanded(false),
                      setRange(event.target.value));
                  },
                  lang: language("Source range"),
                  "data-ui-label": "Source range",
                }),
              ],
            }),
            i.jsxs("label", {
              className: "dup-check",
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: headers,
                  onChange: (event) => {
                    (resetPreview(), setHeaders(event.target.checked));
                  },
                }),
                i.jsx(LocalizedText, { children: "My data has headers" }),
              ],
            }),
          ],
        }),
        sourceError
          ? i.jsx("p", { role: "alert", children: sourceError })
          : i.jsxs("div", {
              className: "sort-scope",
              children: [
                i.jsxs("strong", {
                  children: [
                    i.jsx(LocalizedText, { children: "Complete records: " }),
                    sortSource?.expandedRange,
                  ],
                }),
                i.jsx("p", {
                  children: i.jsx(LocalizedText, {
                    children:
                      "Every populated column in these worksheet rows moves together, including fields beside the selection. Headers stay in place. Applying also clears temporary view sorting and search.",
                  }),
                }),
                sortSource?.expansionRequired &&
                  i.jsxs("fieldset", {
                    children: [
                      i.jsx("legend", {
                        children: i.jsx(LocalizedText, {
                          children: "Expand selection",
                        }),
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "radio",
                            name: `${dialogId}-expand`,
                            checked: !isSelectionExpanded,
                            onChange: () => {
                              (resetPreview(), setIsSelectionExpanded(false));
                            },
                          }),
                          i.jsx(LocalizedText, {
                            children:
                              "Keep current selection (sorting blocked to protect records)",
                          }),
                        ],
                      }),
                      i.jsxs("label", {
                        children: [
                          i.jsx("input", {
                            type: "radio",
                            name: `${dialogId}-expand`,
                            checked: isSelectionExpanded,
                            onChange: () => {
                              (resetPreview(), setIsSelectionExpanded(true));
                            },
                          }),
                          i.jsx(LocalizedText, { children: "Expand to " }),
                          sortSource.expandedRange,
                        ],
                      }),
                    ],
                  }),
              ],
            }),
        i.jsxs("div", {
          className: "sort-level-heading",
          children: [
            i.jsx("h3", {
              children: i.jsx(LocalizedText, { children: "Sort levels" }),
            }),
            i.jsxs("button", {
              disabled:
                sortKeys.length >= sortEngineLimits.keys ||
                sortKeys.length >= columnIndexes.length,
              onClick: () => {
                (resetPreview(),
                  setSortKeys([
                    ...sortKeys,
                    buildSortKeySpec(
                      columnIndexes.find(
                        (columnIndex) =>
                          !sortKeys.some(
                            (sortKey) => sortKey.column === columnIndex,
                          ),
                      ),
                    ),
                  ]));
              },
              children: [
                i.jsx(Plus, { size: 15 }),
                i.jsx(LocalizedText, { children: "Add level" }),
              ],
            }),
          ],
        }),
        i.jsx("ol", {
          className: "sort-levels",
          children: sortKeys.map((X$1, levelIndex) =>
            i.jsxs(
              "li",
              {
                children: [
                  i.jsx("span", {
                    className: "sort-level-number",
                    children: levelIndex + 1,
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Column" }),
                      columnIndexes.length > 500
                        ? i.jsx(LocalizedNumberField, {
                            "aria-label": `Column ${levelIndex + 1}`,
                            min: expandedAddress.firstCol + 1,
                            max: expandedAddress.lastCol + 1,
                            value: X$1.column + 1,
                            onValueChange: (columnNumber) =>
                              updateSortKey(levelIndex, {
                                column: Number(columnNumber) - 1,
                              }),
                          })
                        : i.jsx("select", {
                            "aria-label": `Column ${levelIndex + 1}`,
                            value: X$1.column,
                            onChange: (event) =>
                              updateSortKey(levelIndex, {
                                column: Number(event.target.value),
                              }),
                            children: columnIndexes.map((columnIndex) =>
                              i.jsxs(
                                "option",
                                {
                                  value: columnIndex,
                                  children: [
                                    columnIndexToLetters(columnIndex),
                                    headers && expandedAddress
                                      ? ` · ${String(sheet?.cells[`${columnIndexToLetters(columnIndex)}${expandedAddress.firstRow + 1}`]?.raw ?? "(blank header)")}`
                                      : "",
                                  ],
                                },
                                columnIndex,
                              ),
                            ),
                          }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Order" }),
                      i.jsxs("select", {
                        "aria-label": `Order ${levelIndex + 1}`,
                        value: X$1.direction,
                        onChange: (event) =>
                          updateSortKey(levelIndex, {
                            direction: event.target.value,
                          }),
                        children: [
                          i.jsx("option", {
                            value: "ascending",
                            children: "Ascending",
                          }),
                          i.jsx("option", {
                            value: "descending",
                            children: "Descending",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Interpret as" }),
                      i.jsxs("select", {
                        "aria-label": `Interpretation ${levelIndex + 1}`,
                        value: X$1.type,
                        onChange: (event) =>
                          updateSortKey(levelIndex, {
                            type: event.target.value,
                            useCustom: false,
                          }),
                        children: [
                          i.jsx("option", {
                            value: "auto",
                            children: "Automatic",
                          }),
                          i.jsx("option", {
                            value: "number",
                            children: "Number",
                          }),
                          i.jsx("option", {
                            value: "date",
                            children: "Date / time",
                          }),
                          i.jsx("option", { value: "text", children: "Text" }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("div", {
                    className: "sort-level-actions",
                    children: [
                      i.jsx("button", {
                        "aria-label": `Move level ${levelIndex + 1} up`,
                        disabled: levelIndex === 0,
                        onClick: () => moveSortKey(levelIndex, -1),
                        children: i.jsx(ArrowUp, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": `Move level ${levelIndex + 1} down`,
                        disabled: levelIndex === sortKeys.length - 1,
                        onClick: () => moveSortKey(levelIndex, 1),
                        children: i.jsx(ArrowDown, { size: 15 }),
                      }),
                      i.jsx("button", {
                        "aria-label": `Remove level ${levelIndex + 1}`,
                        disabled: sortKeys.length === 1,
                        onClick: () => {
                          (resetPreview(),
                            setSortKeys(
                              sortKeys.filter(
                                (q, keyIndex) => keyIndex !== levelIndex,
                              ),
                            ));
                        },
                        children: i.jsx(X, { size: 15 }),
                      }),
                    ],
                  }),
                  X$1.type === "text" &&
                    i.jsxs("div", {
                      className: "sort-custom",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              checked: X$1.useCustom,
                              onChange: (event) =>
                                updateSortKey(levelIndex, {
                                  useCustom: event.target.checked,
                                }),
                            }),
                            i.jsx(LocalizedText, {
                              children: "Use a custom order for level ",
                            }),
                            levelIndex + 1,
                          ],
                        }),
                        X$1.useCustom &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "One value per line",
                              }),
                              i.jsx("textarea", {
                                "aria-label": `Custom order ${levelIndex + 1}`,
                                rows: 3,
                                placeholder: `High
Medium
Low`,
                                value: X$1.customText,
                                onChange: (event) =>
                                  updateSortKey(levelIndex, {
                                    customText: event.target.value,
                                  }),
                              }),
                              i.jsx("small", {
                                children: i.jsx(LocalizedText, {
                                  children:
                                    "Unlisted values follow the list; blanks stay last. Descending reverses the list.",
                                }),
                              }),
                            ],
                          }),
                      ],
                    }),
                ],
              },
              levelIndex,
            ),
          ),
        }),
        i.jsxs("div", {
          className: "sort-options",
          children: [
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: isCaseSensitive,
                  onChange: (event) => {
                    (resetPreview(), setIsCaseSensitive(event.target.checked));
                  },
                }),
                i.jsx(LocalizedText, { children: "Case-sensitive text" }),
              ],
            }),
            i.jsxs("label", {
              children: [
                i.jsx("input", {
                  type: "checkbox",
                  checked: shouldIncludeHidden,
                  onChange: (event) => {
                    (resetPreview(),
                      setShouldIncludeHidden(event.target.checked));
                  },
                }),
                i.jsx(LocalizedText, { children: "Include hidden rows" }),
              ],
            }),
          ],
        }),
        i.jsx("p", {
          className: "sort-policy",
          children: i.jsx(LocalizedText, {
            children:
              "Ties keep their original row order. Blanks stay last. Automatic keeps numbers and numeric text distinct. Limit: 10,000 rows and 100,000 cells.",
          }),
        }),
        i.jsxs("details", {
          children: [
            i.jsx("summary", {
              children: i.jsx(LocalizedText, {
                children: "Comparison and reference policy",
              }),
            }),
            Object.values(stableSortPolicy).map((policyText) =>
              i.jsx("p", { children: policyText }, policyText),
            ),
          ],
        }),
        dialogProps.readOnly &&
          i.jsx("p", {
            role: "status",
            children: i.jsx(LocalizedText, {
              children:
                "This host is read-only. Sorting cannot change records.",
            }),
          }),
        errorMessage &&
          i.jsx("p", {
            className: "dup-error",
            role: "alert",
            children: errorMessage,
          }),
        sortPreview &&
          !isPreviewCurrent &&
          i.jsx("p", {
            role: "alert",
            children: i.jsx(LocalizedText, {
              children:
                "The workbook, selection or permissions changed. Preview again.",
            }),
          }),
        sortPreview &&
          isPreviewCurrent &&
          i.jsxs("div", {
            className: "dup-preview",
            "aria-label": text("Sort preview"),
            ref: previewRef,
            tabIndex: -1,
            lang: language("Sort preview"),
            "data-ui-label": "Sort preview",
            children: [
              i.jsxs("h3", {
                children: [
                  sortPreview.movedRows,
                  " ",
                  i.jsx(LocalizedText, { children: "records move · " }),
                  sortPreview.changedCells,
                  " ",
                  i.jsx(LocalizedText, { children: "cells change" }),
                ],
              }),
              i.jsxs("p", {
                children: [
                  sortPreview.rows.length,
                  " ",
                  i.jsx(LocalizedText, { children: "records reviewed · " }),
                  sortPreview.hiddenRows,
                  " ",
                  i.jsx(LocalizedText, { children: "hidden rows · " }),
                  sortPreview.formulaChanges.length,
                  " ",
                  i.jsx(LocalizedText, { children: "formula rewrites ·" }),
                  " ",
                  sortPreview.metadataChanges.length,
                  " ",
                  i.jsx(LocalizedText, { children: "metadata updates" }),
                ],
              }),
              i.jsx("div", {
                className: "dup-scroll",
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
                              children: "Destination row",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Sort values",
                            }),
                          }),
                          i.jsx("th", {
                            children: i.jsx(LocalizedText, {
                              children: "Source identity",
                            }),
                          }),
                        ],
                      }),
                    }),
                    i.jsx("tbody", {
                      children: sortPreview.rows
                        .slice(
                          currentPage * sortEngineLimits.pageSize,
                          (currentPage + 1) * sortEngineLimits.pageSize,
                        )
                        .map((previewRow) =>
                          i.jsxs(
                            "tr",
                            {
                              children: [
                                i.jsx("td", { children: previewRow.sourceRow }),
                                i.jsx("td", {
                                  children: previewRow.destinationRow,
                                }),
                                i.jsx("td", {
                                  children: previewRow.keys
                                    .map((key) =>
                                      key === null || key === ""
                                        ? "(blank)"
                                        : String(key),
                                    )
                                    .join(" · "),
                                }),
                                i.jsxs("td", {
                                  children: [
                                    i.jsxs("span", {
                                      children: [
                                        previewRow.sourceName ??
                                          "Unchanged source",
                                        previewRow.originalRow
                                          ? ` · row ${previewRow.originalRow}`
                                          : "",
                                      ],
                                    }),
                                    previewRow.sourceId &&
                                      i.jsx("small", {
                                        title: previewRow.sourceId,
                                        children: previewRow.sourceId,
                                      }),
                                  ],
                                }),
                              ],
                            },
                            previewRow.sourceRow,
                          ),
                        ),
                    }),
                  ],
                }),
              }),
              i.jsxs("nav", {
                "aria-label": text("Sort preview pages"),
                lang: language("Sort preview pages"),
                "data-ui-label": "Sort preview pages",
                children: [
                  i.jsx("button", {
                    disabled: currentPage === 0,
                    onClick: () => setPageIndex(currentPage - 1),
                    children: i.jsx(LocalizedText, { children: "Previous" }),
                  }),
                  i.jsxs("span", {
                    children: [
                      i.jsx(LocalizedText, { children: "Page " }),
                      currentPage + 1,
                      " ",
                      i.jsx(LocalizedText, { children: "of " }),
                      pageCount,
                    ],
                  }),
                  i.jsx("button", {
                    disabled: currentPage === pageCount - 1,
                    onClick: () => setPageIndex(currentPage + 1),
                    children: i.jsx(LocalizedText, { children: "Next" }),
                  }),
                ],
              }),
              sortPreview.formulaChanges.length > 0 &&
                i.jsxs("details", {
                  children: [
                    i.jsxs("summary", {
                      children: [
                        i.jsx(LocalizedText, { children: "Review " }),
                        sortPreview.formulaChanges.length,
                        " ",
                        i.jsx(LocalizedText, { children: "formula changes" }),
                      ],
                    }),
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "First 50 changes shown; the SDK preview contains the complete list.",
                      }),
                    }),
                    sortPreview.formulaChanges
                      .slice(0, 50)
                      .map((formulaChange, formulaChangeIndex) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              i.jsxs("strong", {
                                children: [
                                  formulaChange.sheet,
                                  "!",
                                  formulaChange.source,
                                  " → ",
                                  formulaChange.destination,
                                ],
                              }),
                              i.jsx("br", {}),
                              i.jsxs("code", {
                                children: [
                                  formulaChange.before,
                                  " → ",
                                  formulaChange.after,
                                ],
                              }),
                            ],
                          },
                          formulaChangeIndex,
                        ),
                      ),
                  ],
                }),
              sortPreview.metadataChanges.length > 0 &&
                i.jsxs("details", {
                  children: [
                    i.jsx("summary", {
                      children: i.jsx(LocalizedText, {
                        children: "Review metadata updates",
                      }),
                    }),
                    sortPreview.metadataChanges.map(
                      (metadataChange, metadataChangeIndex) =>
                        i.jsxs(
                          "p",
                          {
                            children: [
                              metadataChange.sheet,
                              ": ",
                              metadataChange.field,
                            ],
                          },
                          metadataChangeIndex,
                        ),
                    ),
                  ],
                }),
            ],
          }),
        i.jsxs("footer", {
          className: "dup-actions",
          children: [
            i.jsx("button", {
              onClick: dialogProps.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsxs("div", {
              children: [
                i.jsx("button", {
                  disabled: dialogProps.readOnly || !!sourceError,
                  onClick: onClick,
                  children: i.jsx(LocalizedText, { children: "Preview sort" }),
                }),
                i.jsx("button", {
                  className: "dup-primary",
                  disabled: !isPreviewCurrent || !sortPreview?.movedRows,
                  onClick: applySortPreview,
                  children: i.jsx(LocalizedText, { children: "Apply sort" }),
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  });
}
const worksheetVisibilityState = (sheet) =>
    sheet.veryHidden ? "veryHidden" : sheet.hidden ? "hidden" : "visible",
  worksheetVisibilityLabel = (visibilityState) =>
    visibilityState === "veryHidden"
      ? "Very hidden"
      : visibilityState === "hidden"
        ? "Hidden"
        : "Visible",
  describeWorksheetError = (error) =>
    error instanceof Error
      ? error.message
      : "The worksheet operation could not be completed.";
function WorksheetTabBar(props) {
  const { text, language } = useWorkbookMessages(),
    [selectedSheetId, setSelectedSheetId] = ne.useState(props.sheetId),
    activeSheet =
      props.workbook.sheets.find((sheet) => sheet.id === selectedSheetId) ??
      props.workbook.sheets[0],
    [draftName, setDraftName] = ne.useState(activeSheet.name),
    [draftVisibility, setDraftVisibility] = ne.useState(
      worksheetVisibilityState(activeSheet),
    ),
    [tabColor, setTabColor] = ne.useState(activeSheet.tabColor ?? "#107c41"),
    [structurePassword, setStructurePassword] = ne.useState(""),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [statusMessage, setStatusMessage] = ne.useState(""),
    [isCheckingPassword, setIsCheckingPassword] = ne.useState(false),
    [pendingDeletion, setPendingDeletion] = ne.useState(null),
    [checked, setChecked] = ne.useState(false),
    [dependencyPage, setDependencyPage] = ne.useState(0),
    dialogRef = ne.useRef(null),
    selectedSheetRef = ne.useRef(null),
    latestPropsRef = ne.useRef(props),
    isMountedRef = ne.useRef(true);
  latestPropsRef.current = props;
  const titleId = ne.useId(),
    isStructureProtected = !!props.workbook.structureProtection?.enabled,
    sheetSummary = ne
      .useMemo(() => summarizeWorksheets(props.workbook), [props.workbook])
      .find((summary) => summary.sheetId === activeSheet.id),
    disabled = props.readOnly || isStructureProtected || isCheckingPassword,
    activeSheetIndex = props.workbook.sheets.findIndex(
      (sheet) => sheet.id === activeSheet.id,
    ),
    isOnlyVisibleSheet =
      !activeSheet.hidden &&
      props.workbook.sheets.filter((sheet) => !sheet.hidden).length === 1,
    canDeleteSheet =
      !disabled && props.workbook.sheets.length > 1 && !isOnlyVisibleSheet;
  (ne.useEffect(() => {
    isMountedRef.current = true;
    const activeElement = document.activeElement;
    return (
      (selectedSheetRef.current ?? dialogRef.current)?.focus(),
      () => {
        ((isMountedRef.current = false),
          activeElement instanceof HTMLElement &&
            activeElement.isConnected &&
            activeElement.focus());
      }
    );
  }, []),
    ne.useEffect(() => {
      (setSelectedSheetId(activeSheet.id),
        setDraftName(activeSheet.name),
        setDraftVisibility(worksheetVisibilityState(activeSheet)),
        setTabColor(activeSheet.tabColor ?? "#107c41"));
    }, [
      activeSheet.id,
      activeSheet.name,
      activeSheet.hidden,
      activeSheet.veryHidden,
      activeSheet.tabColor,
    ]));
  function selectSheetTab(sheetId) {
    (setSelectedSheetId(sheetId),
      setPendingDeletion(null),
      setChecked(false),
      setErrorMessage(""),
      setStatusMessage(""));
  }
  function runWorkbookOperation(
    operationLabel,
    applyOperation,
    isCreatingSheet = false,
  ) {
    (setErrorMessage(""), setStatusMessage(""));
    try {
      const currentProps = latestPropsRef.current;
      if (currentProps.readOnly) throw new Error("This workbook is read-only.");
      const workbook = applyOperation(currentProps.workbook);
      if (workbook === currentProps.workbook) return;
      (currentProps.onWorkbookChange(workbook, operationLabel),
        setPendingDeletion(null),
        setChecked(false),
        setDependencyPage(0),
        setStatusMessage(operationLabel));
      const newSheet = isCreatingSheet
        ? workbook.sheets.find(
            (sheet) =>
              !currentProps.workbook.sheets.some(
                (existingSheet) => existingSheet.id === sheet.id,
              ),
          )
        : undefined;
      (newSheet
        ? (setSelectedSheetId(newSheet.id),
          newSheet.hidden || currentProps.onSelectSheet(newSheet.id))
        : workbook.sheets.some((sheet) => sheet.id === activeSheet.id) ||
          setSelectedSheetId(workbook.sheets.find((sheet) => !sheet.hidden).id),
        !newSheet &&
          !workbook.sheets.some(
            (sheet) => sheet.id === currentProps.sheetId && !sheet.hidden,
          ) &&
          currentProps.onSelectSheet(
            workbook.sheets.find((sheet) => !sheet.hidden).id,
          ));
    } catch (error) {
      setErrorMessage(describeWorksheetError(error));
    }
  }
  function onClick() {
    const existingNames = new Set(
      props.workbook.sheets.map((sheet) => sheet.name.toLowerCase()),
    );
    let sheetNumber = 1;
    for (; existingNames.has(`sheet${sheetNumber}`);) sheetNumber++;
    runWorkbookOperation(
      "Add worksheet",
      (activeWorkbook) => addWorksheet(activeWorkbook, `Sheet${sheetNumber}`),
      true,
    );
  }
  function startDeletion() {
    (setErrorMessage(""), setStatusMessage(""), setDependencyPage(0));
    try {
      (setPendingDeletion({
        workbook: props.workbook,
        sheetId: activeSheet.id,
        name: activeSheet.name,
        dependencies: describeSheetDeletionImpacts(
          props.workbook,
          activeSheet.id,
        ),
      }),
        setChecked(false));
    } catch (error) {
      setErrorMessage(describeWorksheetError(error));
    }
  }
  function X$1() {
    if (!(!pendingDeletion || !checked)) {
      if (pendingDeletion.workbook !== latestPropsRef.current.workbook) {
        setErrorMessage(
          "The workbook changed. Review current dependencies before deleting.",
        );
        return;
      }
      runWorkbookOperation(
        `Delete worksheet: ${pendingDeletion.name}`,
        (workbook) => deleteWorksheet(workbook, pendingDeletion.sheetId, true),
      );
    }
  }
  async function toggleStructureProtection() {
    const startingProps = latestPropsRef.current,
      wasStructureProtected =
        !!startingProps.workbook.structureProtection?.enabled;
    (setErrorMessage(""), setStatusMessage(""), setIsCheckingPassword(true));
    try {
      if (startingProps.readOnly)
        throw new Error("This workbook is read-only.");
      const updatedWorkbook = await (wasStructureProtected
        ? unprotectWorkbookStructure(startingProps.workbook, structurePassword)
        : protectWorkbookStructure(startingProps.workbook, structurePassword));
      if (!isMountedRef.current) return;
      const currentProps = latestPropsRef.current;
      if (currentProps.readOnly)
        throw new Error(
          "The host changed this workbook to read-only. No structure change was applied.",
        );
      if (currentProps.workbook !== startingProps.workbook)
        throw new Error(
          "The workbook changed while checking the password. Try again with the current workbook.",
        );
      (currentProps.onWorkbookChange(
        updatedWorkbook,
        wasStructureProtected
          ? "Unprotect workbook structure"
          : "Protect workbook structure",
      ),
        setStructurePassword(""),
        setPendingDeletion(null),
        setStatusMessage(
          wasStructureProtected
            ? "Workbook structure is unlocked."
            : "Workbook structure is protected.",
        ));
    } catch (error) {
      isMountedRef.current && setErrorMessage(describeWorksheetError(error));
    } finally {
      isMountedRef.current && setIsCheckingPassword(false);
    }
  }
  const isDeletionStale = !!(
      pendingDeletion && pendingDeletion.workbook !== props.workbook
    ),
    hasUnsupportedDependencies = pendingDeletion?.dependencies.some(
      (dependency) => /unsupported/i.test(dependency.kind),
    );
  return i.jsx("div", {
    className: "sm-backdrop",
    onPointerDown: (event) => {
      event.target === event.currentTarget && props.onClose();
    },
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "sm-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      onKeyDown: (event) => {
        var dialogElement;
        if (
          (event.stopPropagation(),
          event.key === "Escape" &&
            (event.preventDefault(),
            pendingDeletion
              ? (setPendingDeletion(null), setChecked(false))
              : props.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                'button,input,select,textarea,[tabindex="0"]',
              ) ?? [],
            ).filter(
              (element) =>
                !element.matches(":disabled") &&
                !element.closest("[hidden]") &&
                element.tabIndex >= 0,
            ),
            firstFocusable = focusableElements[0],
            lastFocusable = focusableElements.at(-1),
            activeElement = document.activeElement;
          focusableElements.length
            ? event.shiftKey &&
              (activeElement === firstFocusable ||
                !focusableElements.includes(activeElement))
              ? (event.preventDefault(),
                lastFocusable == null || lastFocusable.focus())
              : !event.shiftKey &&
                (activeElement === lastFocusable ||
                  !focusableElements.includes(activeElement)) &&
                (event.preventDefault(),
                firstFocusable == null || firstFocusable.focus())
            : (event.preventDefault(),
              (dialogElement = dialogRef.current) == null ||
                dialogElement.focus());
        }
      },
      children: [
        i.jsxs("header", {
          className: "sm-header",
          children: [
            i.jsxs("div", {
              children: [
                i.jsx(FileSpreadsheet, { size: 19 }),
                i.jsx("h2", {
                  id: titleId,
                  children: i.jsx(LocalizedText, {
                    children: "Manage Worksheets",
                  }),
                }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close worksheet manager"),
              onClick: props.onClose,
              lang: language("Close worksheet manager"),
              "data-ui-label": "Close worksheet manager",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "sm-body",
          children: [
            i.jsxs("aside", {
              className: "sm-sheet-list",
              "aria-label": text("Workbook worksheets"),
              lang: language("Workbook worksheets"),
              "data-ui-label": "Workbook worksheets",
              children: [
                i.jsxs("div", {
                  className: "sm-list-heading",
                  children: [
                    i.jsxs("span", {
                      children: [
                        props.workbook.sheets.length,
                        " ",
                        i.jsx(LocalizedText, { children: "worksheets" }),
                      ],
                    }),
                    i.jsx("button", {
                      type: "button",
                      "aria-label": text("Add worksheet"),
                      title: isStructureProtected
                        ? "Unprotect workbook structure to add worksheets"
                        : "Add worksheet",
                      disabled: disabled,
                      onClick: onClick,
                      lang: language("Add worksheet"),
                      "data-ui-label": "Add worksheet",
                      children: i.jsx(Plus, { size: 16 }),
                    }),
                  ],
                }),
                i.jsx("ul", {
                  children: props.workbook.sheets.map((sheet) =>
                    i.jsx(
                      "li",
                      {
                        children: i.jsxs("button", {
                          type: "button",
                          ref:
                            sheet.id === activeSheet.id
                              ? selectedSheetRef
                              : undefined,
                          "aria-label": `Manage ${sheet.name}`,
                          "aria-pressed": sheet.id === activeSheet.id,
                          onClick: () => selectSheetTab(sheet.id),
                          children: [
                            i.jsx("span", {
                              className: "sm-tab-color",
                              style: {
                                background: sheet.tabColor ?? "#bcc7c0",
                              },
                            }),
                            i.jsxs("span", {
                              className: "sm-sheet-name",
                              title: sheet.name,
                              children: [
                                sheet.name,
                                i.jsxs("small", {
                                  children: [
                                    worksheetVisibilityLabel(
                                      worksheetVisibilityState(sheet),
                                    ),
                                    sheet.id === props.sheetId
                                      ? " · Current"
                                      : "",
                                  ],
                                }),
                              ],
                            }),
                            sheet.hidden && i.jsx(EyeOff, { size: 14 }),
                          ],
                        }),
                      },
                      sheet.id,
                    ),
                  ),
                }),
              ],
            }),
            i.jsxs("div", {
              className: "sm-details",
              children: [
                i.jsxs("div", {
                  className: "sm-detail-title",
                  children: [
                    i.jsx("h3", {
                      title: activeSheet.name,
                      children: activeSheet.name,
                    }),
                    i.jsxs("button", {
                      type: "button",
                      className: "sm-view",
                      disabled: activeSheet.hidden,
                      title: activeSheet.hidden
                        ? "Make the worksheet visible before opening it"
                        : "Open this worksheet",
                      onClick: () => {
                        (props.onSelectSheet(activeSheet.id), props.onClose());
                      },
                      children: [
                        i.jsx(Eye$1, { size: 14 }),
                        i.jsx(LocalizedText, { children: "Go to sheet" }),
                      ],
                    }),
                  ],
                }),
                props.readOnly &&
                  i.jsx("p", {
                    className: "sm-hint",
                    children: i.jsx(LocalizedText, {
                      children:
                        "The host opened this workbook read-only. You can inspect its worksheets.",
                    }),
                  }),
                isStructureProtected &&
                  i.jsxs("p", {
                    className: "sm-lock-note",
                    children: [
                      i.jsx(LockKeyhole, { size: 14 }),
                      i.jsx(LocalizedText, {
                        children:
                          "Workbook structure is protected. Unlock it below to make changes.",
                      }),
                    ],
                  }),
                i.jsxs("section", {
                  className: "sm-inspection",
                  "aria-label": text("Worksheet inspection metadata"),
                  lang: language("Worksheet inspection metadata"),
                  "data-ui-label": "Worksheet inspection metadata",
                  children: [
                    i.jsx("h4", {
                      children: i.jsx(LocalizedText, {
                        children: "Worksheet inspection",
                      }),
                    }),
                    i.jsxs("dl", {
                      children: [
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Visibility",
                              }),
                            }),
                            i.jsx("dd", {
                              children: worksheetVisibilityLabel(
                                sheetSummary.visibility,
                              ),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Stored cells",
                              }),
                            }),
                            i.jsx("dd", {
                              children:
                                sheetSummary.storedCellCount.toLocaleString(),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Formula cells",
                              }),
                            }),
                            i.jsx("dd", {
                              children:
                                sheetSummary.formulaCellCount.toLocaleString(),
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Tab color",
                              }),
                            }),
                            i.jsx("dd", {
                              children: sheetSummary.tabColor ?? "None",
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          children: [
                            i.jsx("dt", {
                              children: i.jsx(LocalizedText, {
                                children: "Calculation",
                              }),
                            }),
                            i.jsx("dd", {
                              children: i.jsx(LocalizedText, {
                                children: "Included when referenced",
                              }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children:
                          "Hidden and very hidden worksheets still participate in calculation. This inspection does not evaluate formulas.",
                      }),
                    }),
                  ],
                }),
                i.jsxs("fieldset", {
                  className: "sm-edit-fields",
                  disabled: disabled,
                  children: [
                    i.jsx("legend", {
                      className: "sm-sr-only",
                      children: i.jsx(LocalizedText, {
                        children: "Worksheet settings",
                      }),
                    }),
                    i.jsxs("form", {
                      className: "sm-name-form",
                      onSubmit: (event) => {
                        (event.preventDefault(),
                          runWorkbookOperation(
                            `Rename worksheet: ${activeSheet.name}`,
                            (workbook) =>
                              renameWorksheet(
                                workbook,
                                activeSheet.id,
                                draftName,
                              ),
                          ));
                      },
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Sheet name" }),
                            i.jsx("input", {
                              "aria-label": text("Sheet name"),
                              value: draftName,
                              maxLength: 31,
                              onChange: (event) =>
                                setDraftName(event.target.value),
                              lang: language("Sheet name"),
                              "data-ui-label": "Sheet name",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "submit",
                          disabled:
                            !draftName.trim() ||
                            draftName.trim() === activeSheet.name,
                          children: i.jsx(LocalizedText, {
                            children: "Rename sheet",
                          }),
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "sm-row-actions",
                      children: [
                        i.jsxs("button", {
                          type: "button",
                          onClick: () =>
                            runWorkbookOperation(
                              `Duplicate worksheet: ${activeSheet.name}`,
                              (workbook) =>
                                duplicateWorksheet(workbook, activeSheet.id),
                              true,
                            ),
                          children: [
                            i.jsx(Copy, { size: 14 }),
                            i.jsx(LocalizedText, {
                              children: "Duplicate sheet",
                            }),
                          ],
                        }),
                        i.jsxs("div", {
                          className: "sm-order",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Move worksheet up"),
                              title: text("Move worksheet earlier"),
                              disabled: activeSheetIndex === 0,
                              onClick: () =>
                                runWorkbookOperation(
                                  `Move worksheet: ${activeSheet.name}`,
                                  (workbook) =>
                                    moveWorksheet(
                                      workbook,
                                      activeSheet.id,
                                      activeSheetIndex - 1,
                                    ),
                                ),
                              lang: language("Move worksheet up"),
                              "data-ui-label": "Move worksheet up",
                              children: i.jsx(ArrowUp, { size: 15 }),
                            }),
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Move worksheet down"),
                              title: text("Move worksheet later"),
                              disabled:
                                activeSheetIndex ===
                                props.workbook.sheets.length - 1,
                              onClick: () =>
                                runWorkbookOperation(
                                  `Move worksheet: ${activeSheet.name}`,
                                  (workbook) =>
                                    moveWorksheet(
                                      workbook,
                                      activeSheet.id,
                                      activeSheetIndex + 1,
                                    ),
                                ),
                              lang: language("Move worksheet down"),
                              "data-ui-label": "Move worksheet down",
                              children: i.jsx(ArrowDown, { size: 15 }),
                            }),
                          ],
                        }),
                      ],
                    }),
                    i.jsxs("div", {
                      className: "sm-setting-row",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Visibility" }),
                            i.jsxs("select", {
                              "aria-label": text("Worksheet visibility"),
                              value: draftVisibility,
                              onChange: (event) =>
                                setDraftVisibility(event.target.value),
                              lang: language("Worksheet visibility"),
                              "data-ui-label": "Worksheet visibility",
                              children: [
                                i.jsx("option", {
                                  value: "visible",
                                  children: "Visible",
                                }),
                                i.jsx("option", {
                                  value: "hidden",
                                  children: "Hidden",
                                }),
                                i.jsx("option", {
                                  value: "veryHidden",
                                  children: "Very hidden",
                                }),
                              ],
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled:
                            draftVisibility ===
                            worksheetVisibilityState(activeSheet),
                          onClick: () =>
                            runWorkbookOperation(
                              `Change worksheet visibility: ${activeSheet.name}`,
                              (workbook) =>
                                setWorksheetVisibility(
                                  workbook,
                                  activeSheet.id,
                                  draftVisibility,
                                ),
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "Apply visibility",
                          }),
                        }),
                      ],
                    }),
                    i.jsx("p", {
                      className: "sm-hint",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Hidden worksheets remain available here. Very hidden records a separate Excel visibility state. At least one worksheet must remain visible.",
                      }),
                    }),
                    i.jsxs("div", {
                      className: "sm-setting-row sm-color-row",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Tab color" }),
                            i.jsx("input", {
                              "aria-label": text("Worksheet tab color"),
                              type: "color",
                              value: tabColor,
                              onChange: (event) =>
                                setTabColor(event.target.value),
                              lang: language("Worksheet tab color"),
                              "data-ui-label": "Worksheet tab color",
                            }),
                          ],
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled:
                            tabColor.toLowerCase() ===
                            activeSheet.tabColor?.toLowerCase(),
                          onClick: () =>
                            runWorkbookOperation(
                              `Change tab color: ${activeSheet.name}`,
                              (workbook) =>
                                setWorksheetTabColor(
                                  workbook,
                                  activeSheet.id,
                                  tabColor,
                                ),
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "Apply tab color",
                          }),
                        }),
                        i.jsx("button", {
                          type: "button",
                          disabled: !activeSheet.tabColor,
                          onClick: () =>
                            runWorkbookOperation(
                              `Remove tab color: ${activeSheet.name}`,
                              (workbook) =>
                                setWorksheetTabColor(workbook, activeSheet.id),
                            ),
                          children: i.jsx(LocalizedText, {
                            children: "No color",
                          }),
                        }),
                      ],
                    }),
                  ],
                }),
                i.jsxs("div", {
                  className: "sm-delete-row",
                  children: [
                    i.jsxs("button", {
                      type: "button",
                      className: "sm-danger",
                      disabled: !canDeleteSheet,
                      title: isOnlyVisibleSheet
                        ? "Keep at least one visible worksheet"
                        : undefined,
                      onClick: startDeletion,
                      children: [
                        i.jsx(Trash2, { size: 14 }),
                        i.jsx(LocalizedText, { children: "Delete worksheet…" }),
                      ],
                    }),
                    i.jsx("small", {
                      children: i.jsx(LocalizedText, {
                        children: "Review affected references before deleting.",
                      }),
                    }),
                  ],
                }),
                pendingDeletion &&
                  i.jsxs("section", {
                    className: "sm-deletion",
                    "aria-label": text("Delete worksheet confirmation"),
                    lang: language("Delete worksheet confirmation"),
                    "data-ui-label": "Delete worksheet confirmation",
                    children: [
                      i.jsxs("h4", {
                        children: [
                          i.jsx(LocalizedText, { children: "Delete “" }),
                          pendingDeletion.name,
                          "”?",
                        ],
                      }),
                      i.jsx("p", {
                        children: i.jsx(LocalizedText, {
                          children:
                            "The worksheet and its contents will be removed. Dependent formulas can become #REF!; dependent refresh definitions are removed while their result cells remain.",
                        }),
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx("strong", {
                            children: pendingDeletion.dependencies.length,
                          }),
                          " ",
                          i.jsx(LocalizedText, { children: "affected" }),
                          " ",
                          pendingDeletion.dependencies.length === 1
                            ? "reference"
                            : "references",
                        ],
                      }),
                      pendingDeletion.dependencies.length
                        ? i.jsx("ul", {
                            children: pendingDeletion.dependencies
                              .slice(
                                dependencyPage * 200,
                                (dependencyPage + 1) * 200,
                              )
                              .map((dependency, dependencyIndex) =>
                                i.jsxs(
                                  "li",
                                  {
                                    children: [
                                      i.jsx("span", {
                                        children: dependency.label,
                                      }),
                                      dependency.address &&
                                        i.jsx("code", {
                                          children: dependency.address,
                                        }),
                                    ],
                                  },
                                  `${dependency.kind}:${dependencyIndex}`,
                                ),
                              ),
                          })
                        : i.jsx("p", {
                            className: "sm-hint",
                            children: i.jsx(LocalizedText, {
                              children: "No dependent references were found.",
                            }),
                          }),
                      pendingDeletion.dependencies.length > 200 &&
                        i.jsxs("div", {
                          className: "sm-dependency-pages",
                          children: [
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Previous references"),
                              disabled: dependencyPage === 0,
                              onClick: () =>
                                setDependencyPage(
                                  (currentPage) => currentPage - 1,
                                ),
                              lang: language("Previous references"),
                              "data-ui-label": "Previous references",
                              children: i.jsx(LocalizedText, {
                                children: "Previous",
                              }),
                            }),
                            i.jsxs("span", {
                              "aria-live": "polite",
                              children: [
                                dependencyPage * 200 + 1,
                                "–",
                                Math.min(
                                  pendingDeletion.dependencies.length,
                                  (dependencyPage + 1) * 200,
                                ),
                                " ",
                                i.jsx(LocalizedText, { children: "of " }),
                                pendingDeletion.dependencies.length,
                                " ",
                                i.jsx(LocalizedText, {
                                  children: "references",
                                }),
                              ],
                            }),
                            i.jsx("button", {
                              type: "button",
                              "aria-label": text("Next references"),
                              disabled:
                                (dependencyPage + 1) * 200 >=
                                pendingDeletion.dependencies.length,
                              onClick: () =>
                                setDependencyPage(
                                  (currentPage) => currentPage + 1,
                                ),
                              lang: language("Next references"),
                              "data-ui-label": "Next references",
                              children: i.jsx(LocalizedText, {
                                children: "Next",
                              }),
                            }),
                          ],
                        }),
                      hasUnsupportedDependencies &&
                        i.jsx("p", {
                          className: "sm-error",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Some references cannot be updated safely. Resolve the listed unsupported references first.",
                          }),
                        }),
                      isDeletionStale &&
                        i.jsx("p", {
                          className: "sm-error",
                          children: i.jsx(LocalizedText, {
                            children:
                              "The workbook changed. Review its current dependencies before deleting.",
                          }),
                        }),
                      i.jsxs("label", {
                        className: "sm-confirm",
                        children: [
                          i.jsx("input", {
                            type: "checkbox",
                            checked: checked,
                            disabled:
                              disabled ||
                              isDeletionStale ||
                              hasUnsupportedDependencies,
                            onChange: (event) =>
                              setChecked(event.target.checked),
                          }),
                          i.jsx(LocalizedText, {
                            children:
                              "I understand that this removes the worksheet and may leave reference errors.",
                          }),
                        ],
                      }),
                      i.jsxs("div", {
                        className: "sm-confirm-actions",
                        children: [
                          i.jsx("button", {
                            type: "button",
                            onClick: () => {
                              (setPendingDeletion(null), setChecked(false));
                            },
                            children: i.jsx(LocalizedText, {
                              children: "Keep worksheet",
                            }),
                          }),
                          isDeletionStale
                            ? i.jsx("button", {
                                type: "button",
                                disabled: disabled,
                                onClick: startDeletion,
                                children: i.jsx(LocalizedText, {
                                  children: "Review current dependencies",
                                }),
                              })
                            : i.jsx("button", {
                                type: "button",
                                className: "sm-danger sm-delete-final",
                                disabled:
                                  !canDeleteSheet ||
                                  !checked ||
                                  hasUnsupportedDependencies,
                                onClick: X$1,
                                children: i.jsx(LocalizedText, {
                                  children: "Delete worksheet",
                                }),
                              }),
                        ],
                      }),
                    ],
                  }),
              ],
            }),
          ],
        }),
        i.jsxs("section", {
          className: "sm-protection",
          "aria-label": text("Workbook structure protection"),
          lang: language("Workbook structure protection"),
          "data-ui-label": "Workbook structure protection",
          children: [
            i.jsxs("div", {
              className: "sm-protection-title",
              children: [
                isStructureProtected
                  ? i.jsx(LockKeyhole, { size: 16 })
                  : i.jsx(LockKeyholeOpen, { size: 16 }),
                i.jsx("strong", {
                  children: i.jsx(LocalizedText, {
                    children: "Workbook structure",
                  }),
                }),
                i.jsx("span", {
                  children: isStructureProtected ? "Protected" : "Unlocked",
                }),
              ],
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Protect worksheet names, order, visibility, tab colors, and adding or deleting sheets.",
              }),
            }),
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "Visibility and structure locks control editing. Access to this workbook and its hidden content is granted separately by the host application.",
              }),
            }),
            i.jsxs("form", {
              onSubmit: (event) => {
                (event.preventDefault(),
                  isCheckingPassword || toggleStructureProtection());
              },
              children: [
                i.jsxs("label", {
                  children: [
                    isStructureProtected
                      ? "Password to unlock"
                      : "Password (optional)",
                    i.jsx("input", {
                      "aria-label": text("Structure password"),
                      type: "password",
                      autoComplete: "off",
                      value: structurePassword,
                      maxLength: 200,
                      disabled: props.readOnly || isCheckingPassword,
                      onChange: (event) =>
                        setStructurePassword(event.target.value),
                      lang: language("Structure password"),
                      "data-ui-label": "Structure password",
                    }),
                  ],
                }),
                i.jsx("button", {
                  type: "submit",
                  disabled: props.readOnly || isCheckingPassword,
                  children: isCheckingPassword
                    ? "Checking…"
                    : isStructureProtected
                      ? "Unprotect workbook structure"
                      : "Protect workbook structure",
                }),
              ],
            }),
          ],
        }),
        (errorMessage || statusMessage) &&
          i.jsx("div", {
            className: "sm-feedback",
            children: errorMessage
              ? i.jsx("p", {
                  role: "alert",
                  className: "sm-error",
                  children: errorMessage,
                })
              : i.jsx("p", { role: "status", children: statusMessage }),
          }),
        i.jsxs("footer", {
          className: "sm-footer",
          children: [
            i.jsx("span", {
              children: i.jsx(LocalizedText, {
                children: "Changes use the workbook’s undo history.",
              }),
            }),
            i.jsx("button", {
              type: "button",
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Done" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const worksheetPropertyLabels = {
    raw: "Value",
    literal: "Entered as text",
    style: "Formatting",
    numFmt: "Number format",
    format: "Format",
    note: "Note",
    tickmark: "Tickmark",
    createdAt: "Created",
    modifiedAt: "Modified",
    engagement: "Engagement",
    period: "Period",
    template: "Template",
    name: "Name",
    hidden: "Hidden",
    veryHidden: "Very hidden",
    tabColor: "Tab color",
    columnWidths: "Column widths",
    rowHeights: "Row heights",
    frozenRows: "Frozen rows",
    structureProtection: "Structure protection",
    protection: "Protection",
    definedNames: "Defined names",
    metadata: "Workbook properties",
    enabled: "Enabled",
    sourceRevision: "Source revision",
    version: "Version",
  },
  formatAuditDetailValue = (detailValue) =>
    detailValue === undefined
      ? "Not set"
      : detailValue === null || detailValue === ""
        ? "Blank"
        : typeof detailValue == "boolean"
          ? detailValue
            ? "Yes"
            : "No"
          : typeof detailValue == "string"
            ? detailValue.startsWith("data:image/")
              ? "Embedded picture"
              : /^\d{4}-\d{2}-\d{2}T.*Z$/.test(detailValue) &&
                  Number.isFinite(Date.parse(detailValue))
                ? new Date(detailValue).toLocaleString()
                : detailValue
            : typeof detailValue != "object"
              ? String(detailValue)
              : Array.isArray(detailValue)
                ? detailValue.length
                  ? detailValue.map(formatAuditDetailValue).join("; ")
                  : "None"
                : Object.entries(detailValue)
                    .filter(
                      ([propertyKey]) =>
                        ![
                          "id",
                          "sheetId",
                          "scopeSheetId",
                          "sourceWorkbookId",
                          "passwordHash",
                          "salt",
                        ].includes(propertyKey),
                    )
                    .map(
                      ([propertyKey, propertyValue]) =>
                        `${worksheetPropertyLabels[propertyKey] ?? propertyKey}: ${formatAuditDetailValue(propertyValue)}`,
                    )
                    .join(" · ") || "No details",
  formatAuditCellValue = (cellValue) =>
    cellValue ? formatAuditDetailValue(cellValue) : "Empty cell",
  formatChangeAttribution = (attribution) =>
    attribution?.actor && attribution.timestamp
      ? `${attribution.actor} · ${new Date(attribution.timestamp).toLocaleString()}${attribution.basis === "operation-history" ? " · from operation history" : ""}`
      : "Author and time not recorded for this change";

export {
  worksheetPropertyLabels,
  formatChangeAttribution,
  formatAuditCellValue,
  WorksheetTabBar,
  formatAuditDetailValue,
  CustomSortDialog,
};
