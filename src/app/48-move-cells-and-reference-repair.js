import {
  assertWorkbookWritable,
  eI,
  formatCellAddress,
} from "./04-cell-input-paste-and-fill.js";
import { isFormulaInputValue } from "./11-local-workbook-store-and-versions.js";
import { formatSortRangeAddress } from "./44-clean-and-reshape.js";
import {
  rangesIntersect,
  isCellInsideRange,
  quoteSheetName,
  throwWorkbookError,
} from "./03-workbook-model.js";
import { X, ArrowRight } from "lucide-react";
import i from "react/jsx-runtime";
import ne from "react";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  summarizeNumberFormat,
  formatNumberWithPattern,
  builtInDateTimeFormats,
  columnWidthAt,
  findSheetOrThrow,
  supportedNumberLocales,
  buildNumberFormatCode,
  getImportedNumberFormatWarning,
  supportedCurrencyCodes,
  Uv,
  calculateWorkbookWithSettings,
  isCellAddressInRanges,
  _t,
  cn,
  parseRangeAddress,
  parseSelectionAreas,
  rowHeightAt,
  normalizeCellRecordInput,
  assertSheetUnprotected,
  validateNumberFormatSettings,
  applyNumberFormatTransaction,
  pQ,
  normalizeWorkbookInput,
  createAxisVisibilityIndex,
  columnIndexToLetters,
  checkDataValidationBatch,
  formatCellValueWithNumberFormat,
  countSelectionCells,
} from "./03-workbook-model.js";
import {
  useWorkbookMessages,
  LocalizedText,
  LocalizedNumberField,
} from "./04-cell-input-paste-and-fill.js";
import {
  pasteOperationLimits,
  hasCellContent,
  parseWholeColumnReference,
  assertCutSourceWithinLimit,
  createCutSourceSnapshot,
  rangeContainsRange,
  parseMovableA1Rectangle,
  assertFormulaReferencesMovable,
  verifyCutSourceSnapshot,
  isUnsafeStructuredReference,
  assertRangeWithinBoundary,
} from "./47-fill-and-paste-special.js";

function remapReferenceForMove(move, sheetId, reference) {
  if (sheetId !== move.source.id)
    return { sheetId: sheetId, reference: reference, changed: false };
  const wholeColumnReference = parseWholeColumnReference(reference);
  if (!rangesIntersect(wholeColumnReference, move.from))
    return { sheetId: sheetId, reference: reference, changed: false };
  if (rangeContainsRange(move.from, wholeColumnReference))
    return {
      sheetId: move.target.id,
      reference: parseMovableA1Rectangle(reference, move.dr, move.dc),
      changed: true,
    };
  const intersection = {
      firstRow: Math.max(wholeColumnReference.firstRow, move.from.firstRow),
      lastRow: Math.min(wholeColumnReference.lastRow, move.from.lastRow),
      firstCol: Math.max(wholeColumnReference.firstCol, move.from.firstCol),
      lastCol: Math.min(wholeColumnReference.lastCol, move.from.lastCol),
    },
    shiftedIntersection = {
      firstRow: intersection.firstRow + move.dr,
      lastRow: intersection.lastRow + move.dr,
      firstCol: intersection.firstCol + move.dc,
      lastCol: intersection.lastCol + move.dc,
    };
  if (
    move.source.id === move.target.id &&
    rangeContainsRange(wholeColumnReference, shiftedIntersection)
  )
    return { sheetId: sheetId, reference: reference, changed: false };
  throwWorkbookError(
    `Reference ${move.source.name}!${reference} includes cells outside the moved block. This move would require a split or cross-sheet range; use explicit references first.`,
  );
}
function qW(book, definedName, visitedNames = new Set()) {
  const nameKey = `${definedName.scopeSheetId ?? ""}:${definedName.name.toLowerCase()}`;
  return visitedNames.has(nameKey) || visitedNames.size > 32
    ? false
    : (visitedNames.add(nameKey),
      assertFormulaReferencesMovable(definedName.refersTo).every(
        (reference) => {
          if (
            reference.kind === "dynamic" ||
            isUnsafeStructuredReference(reference)
          )
            return false;
          if (reference.kind === "reference") return !!reference.sheet;
          if (reference.kind === "name" && !reference.sheet) {
            const workbookName = (book.definedNames ?? []).find(
              (entry) =>
                !entry.scopeSheetId &&
                entry.name.toLowerCase() === reference.name.toLowerCase(),
            );
            return (
              !workbookName || qW(book, workbookName, new Set(visitedNames))
            );
          }
          return true;
        },
      ));
}
function XW(book, definedName, visitedNames = new Set()) {
  const nameKey = `${definedName.scopeSheetId ?? ""}:${definedName.name.toLowerCase()}`;
  return visitedNames.has(nameKey) || visitedNames.size > 32
    ? true
    : (visitedNames.add(nameKey),
      assertFormulaReferencesMovable(definedName.refersTo).some((reference) => {
        if (isUnsafeStructuredReference(reference)) return true;
        if (reference.kind !== "name" || reference.sheet) return false;
        const resolvedName =
          (book.definedNames ?? []).find(
            (entry) =>
              entry.scopeSheetId === definedName.scopeSheetId &&
              entry.name.toLowerCase() === reference.name.toLowerCase(),
          ) ??
          (book.definedNames ?? []).find(
            (workbookEntry) =>
              !workbookEntry.scopeSheetId &&
              workbookEntry.name.toLowerCase() === reference.name.toLowerCase(),
          );
        return !!resolvedName && XW(book, resolvedName, new Set(visitedNames));
      }));
}
function repairReferencesForMove(
  move,
  refersTo,
  originSheetId,
  currentSheetId,
  isMovedCell = false,
) {
  ((move.characters += refersTo.length),
    move.characters > pasteOperationLimits.formulaCharacters &&
      throwWorkbookError(
        "Reference repair exceeds 10 million formula characters. Simplify the workbook before moving cells.",
      ));
  const references = assertFormulaReferencesMovable(refersTo);
  let rebuiltText = "",
    copiedUpTo = 0;
  for (const reference of references) {
    let text = reference.text;
    if (
      (reference.kind === "dynamic" &&
        throwWorkbookError(
          `${reference.name} can refer to cells dynamically. Replace dynamic references with explicit A1 references before moving cells.`,
        ),
      isUnsafeStructuredReference(reference) &&
        isMovedCell &&
        throwWorkbookError(
          "A current-row structured reference cannot be safely moved out of its table context.",
        ),
      reference.kind === "name" && isMovedCell && !reference.sheet)
    ) {
      const definedNames = move.book.definedNames ?? [],
        originScopedName = definedNames.find(
          (originScopedEntry) =>
            originScopedEntry.scopeSheetId === originSheetId &&
            originScopedEntry.name.toLowerCase() ===
              reference.name.toLowerCase(),
        ),
        workbookScopedName = definedNames.find(
          (globalEntry) =>
            !globalEntry.scopeSheetId &&
            globalEntry.name.toLowerCase() === reference.name.toLowerCase(),
        ),
        currentScopedName = definedNames.find(
          (currentScopedEntry) =>
            currentScopedEntry.scopeSheetId === currentSheetId &&
            currentScopedEntry.name.toLowerCase() ===
              reference.name.toLowerCase(),
        ),
        resolvedName = originScopedName ?? workbookScopedName;
      (resolvedName &&
        XW(move.book, resolvedName) &&
        throwWorkbookError(
          `Name ${reference.name} contains a current-row structured reference. Use an explicit qualified cell reference before moving this formula.`,
        ),
        originSheetId !== currentSheetId &&
          (originScopedName ||
            currentScopedName ||
            (workbookScopedName && !qW(move.book, workbookScopedName))) &&
          throwWorkbookError(
            `Name ${reference.name} depends on worksheet scope. Keep the formula on its sheet or replace that name with explicit qualified references before moving it.`,
          ));
    }
    if (reference.kind === "reference") {
      const sheetId = reference.sheet
        ? move.book.sheets.find(
            (sheet) =>
              sheet.name.toLowerCase() === reference.sheet.toLowerCase(),
          )?.id
        : originSheetId;
      if (!sheetId)
        !reference.sheet &&
          rangesIntersect(
            parseWholeColumnReference(reference.reference),
            move.from,
          ) &&
          throwWorkbookError(
            "A workbook-level name has a context-dependent unqualified reference. Qualify its worksheet before moving cells.",
          );
      else {
        const remapResult = remapReferenceForMove(
          move,
          sheetId,
          reference.reference,
        );
        remapResult.changed
          ? (text = `${reference.sheet || remapResult.sheetId !== currentSheetId ? quoteSheetName(findSheetOrThrow(move.book, remapResult.sheetId).name) + "!" : ""}${remapResult.reference}`)
          : !reference.sheet &&
            originSheetId !== currentSheetId &&
            originSheetId &&
            (text = `${quoteSheetName(findSheetOrThrow(move.book, originSheetId).name)}!${reference.reference}`);
      }
    }
    text !== reference.text &&
      ((rebuiltText += refersTo.slice(copiedUpTo, reference.start) + text),
      (copiedUpTo = reference.end));
  }
  return rebuiltText + refersTo.slice(copiedUpTo);
}
function computeDrawingFootprint(sheet, drawing) {
  const anchorCell = _t(drawing.anchor),
    axisFootprint = (axis, anchorIndex, offset, size) => {
      const axisVisibilityIndex = createAxisVisibilityIndex(
          axis === "row" ? sheet.rowCount : sheet.colCount,
          cn(sheet, axis),
        ),
        sizeAt = (axisIndex) =>
          axis === "row"
            ? rowHeightAt(sheet, axisIndex)
            : columnWidthAt(sheet, axisIndex);
      if (axisVisibilityIndex.isHidden(anchorIndex))
        return [anchorIndex, anchorIndex];
      const indexAtOffset = (remaining) => {
        let currentIndex = anchorIndex,
          stepCount = 0;
        if (remaining >= 0)
          for (; remaining >= sizeAt(currentIndex);) {
            remaining -= sizeAt(currentIndex);
            const nextIndex = axisVisibilityIndex.next(currentIndex, 1);
            if (nextIndex === null) return currentIndex;
            ((currentIndex = nextIndex),
              ++stepCount > 1e3 &&
                throwWorkbookError(
                  "Drawing footprint exceeds the move inspection budget.",
                ));
          }
        else
          for (; remaining < 0;) {
            const previousIndex = axisVisibilityIndex.next(currentIndex, -1);
            if (previousIndex === null) return 0;
            ((currentIndex = previousIndex),
              (remaining += sizeAt(currentIndex)),
              ++stepCount > 1e3 &&
                throwWorkbookError(
                  "Drawing footprint exceeds the move inspection budget.",
                ));
          }
        return currentIndex;
      };
      return [indexAtOffset(offset), indexAtOffset(offset + size - 1e-4)];
    },
    rowSpan = axisFootprint(
      "row",
      anchorCell.row,
      drawing.offsetY,
      drawing.height,
    ),
    columnSpan = axisFootprint(
      "column",
      anchorCell.col,
      drawing.offsetX,
      drawing.width,
    );
  return {
    firstRow: rowSpan[0],
    lastRow: rowSpan[1],
    firstCol: columnSpan[0],
    lastCol: columnSpan[1],
  };
}
function findMoveBlockers(move) {
  const formulaEntries = collectFormulaEntries(move.book);
  for (const [sheet, rectangle] of [
    [move.source, move.from],
    [move.target, move.to],
  ]) {
    for (const mergeAddress of sheet.merges ?? [])
      rangesIntersect(rectangle, parseRangeAddress(mergeAddress)) &&
        throwWorkbookError(
          "Unmerge the source and destination before moving cells.",
        );
    for (const table of sheet.tables ?? [])
      rangesIntersect(rectangle, parseRangeAddress(table.range)) &&
        throwWorkbookError(
          `The move intersects table ${table.name}. Table-cell moves are not supported; use ordinary cells outside tables.`,
        );
    for (const pivotOrQuery of [
      ...(sheet.pivots ?? []),
      ...(sheet.queries ?? []),
    ])
      rangesIntersect(rectangle, parseRangeAddress(pivotOrQuery.outputRange)) &&
        throwWorkbookError(
          "The move intersects a generated PivotTable or query output. Move ordinary cells instead.",
        );
    for (const rule of [
      ...(sheet.validations ?? []),
      ...(sheet.conditionalFormats ?? []),
    ])
      rangesIntersect(rectangle, parseRangeAddress(rule.range)) &&
        throwWorkbookError(
          "The move intersects a validation or conditional-formatting rule. Remove or relocate the rule first.",
        );
    for (const drawingObject of sheet.objects ?? [])
      (isCellInsideRange(rectangle, _t(drawingObject.anchor)) ||
        rangesIntersect(
          rectangle,
          computeDrawingFootprint(sheet, drawingObject),
        )) &&
        throwWorkbookError(
          `The move intersects drawing ${drawingObject.label || drawingObject.id}. Drawing footprints cannot be moved or partially overwritten by this command.`,
        );
    for (const formulaEntry of formulaEntries)
      formulaEntry.sheetId === sheet.id &&
        rangesIntersect(rectangle, parseRangeAddress(formulaEntry.range)) &&
        throwWorkbookError(
          "The move intersects a dynamic-array footprint. Convert it to ordinary values or move cells outside the array.",
        );
  }
}
function collectFormulaEntries(book) {
  const workbookWithSettings = calculateWorkbookWithSettings(book),
    entries = [];
  let formulaCount = 0,
    formulaCharacterCount = 0;
  for (const sheet of book.sheets)
    for (const [anchor, cell] of Object.entries(sheet.cells))
      if (isFormulaInputValue(cell)) {
        (++formulaCount > pasteOperationLimits.formulaCells ||
          (formulaCharacterCount += cell.raw.length) >
            pasteOperationLimits.formulaCharacters) &&
          throwWorkbookError(
            "Array inspection supports at most 10,000 formulas and 10 million formula characters. Simplify the workbook before moving cells.",
          );
        const range = workbookWithSettings.getSpillRange(sheet.id, anchor);
        if (range) {
          const rangeAddress = parseRangeAddress(range);
          (rangeAddress.firstRow !== rangeAddress.lastRow ||
            rangeAddress.firstCol !== rangeAddress.lastCol) &&
            entries.push({ sheetId: sheet.id, anchor: anchor, range: range });
        }
      }
  return entries;
}
function formulaEntriesSignature(book) {
  return collectFormulaEntries(book)
    .map((entry) => `${entry.sheetId}:${entry.anchor}:${entry.range}`)
    .sort().join(`
`);
}
function computeRangeMoveResult(
  book,
  sourceSheetId,
  sourceRange,
  targetSheetId,
  targetRange,
  options = {},
) {
  assertWorkbookWritable(options.readOnly);
  const sheetOrThrow = findSheetOrThrow(book, sourceSheetId),
    targetSheet = findSheetOrThrow(book, targetSheetId);
  (assertSheetUnprotected(sheetOrThrow), assertSheetUnprotected(targetSheet));
  const from = assertCutSourceWithinLimit(sheetOrThrow, sourceRange),
    destination = assertRangeWithinBoundary(targetSheet, targetRange),
    sourceRowCount = from.lastRow - from.firstRow + 1,
    sourceColumnCount = from.lastCol - from.firstCol + 1;
  (destination.lastRow !== destination.firstRow ||
    destination.lastCol !== destination.firstCol) &&
    (destination.lastRow - destination.firstRow + 1 !== sourceRowCount ||
      destination.lastCol - destination.firstCol + 1 !== sourceColumnCount) &&
    throwWorkbookError(
      "Choose one destination cell or a rectangle with exactly the source dimensions.",
    );
  const targetRect = {
    firstRow: destination.firstRow,
    lastRow: destination.firstRow + sourceRowCount - 1,
    firstCol: destination.firstCol,
    lastCol: destination.firstCol + sourceColumnCount - 1,
  };
  (targetRect.lastRow >= targetSheet.rowCount ||
    targetRect.lastCol >= targetSheet.colCount) &&
    throwWorkbookError("The moved cells would exceed the worksheet boundary.");
  const sourceSnapshot =
    options.sourceSnapshot ??
    createCutSourceSnapshot(book, sourceSheetId, sourceRange);
  verifyCutSourceSnapshot(book, sourceSnapshot, sourceSheetId, sourceRange);
  const move = {
      book: book,
      source: sheetOrThrow,
      target: targetSheet,
      from: from,
      to: targetRect,
      dr: targetRect.firstRow - from.firstRow,
      dc: targetRect.firstCol - from.firstCol,
      characters: 0,
    },
    sortRangeAddress = formatSortRangeAddress(sourceRange),
    targetRangeAddress = formatSortRangeAddress(
      `${formatCellAddress(targetRect.firstRow, targetRect.firstCol)}:${formatCellAddress(targetRect.lastRow, targetRect.lastCol)}`,
    ),
    metadataChanges = [],
    isNoOpMove =
      sourceSheetId === targetSheetId && move.dr === 0 && move.dc === 0;
  isNoOpMove || findMoveBlockers(move);
  const workbook = isNoOpMove ? book : structuredClone(book),
    sourceAddressByTarget = new Map();
  let overwrittenCells = 0,
    referenceChanges = 0;
  const recordMetadataChange = (metadataChange) => {
    (metadataChanges.push(metadataChange),
      metadataChanges.length > pasteOperationLimits.metadataChanges &&
        throwWorkbookError(
          "The move would change more than 500 metadata references. Narrow the source.",
        ));
  };
  if (
    (!isNoOpMove &&
      book.sheets.some((sheet) =>
        sheet.conditionalFormats?.some(
          (conditionalFormat) => conditionalFormat.kind === "formula",
        ),
      ) &&
      throwWorkbookError(
        "Range moves cannot repair formula-based conditional rules yet. Remove those rules before moving cells.",
      ),
    !isNoOpMove)
  ) {
    const repairResult = pQ(
      workbook,
      sourceSheetId,
      sortRangeAddress,
      targetSheetId,
      targetRangeAddress,
    );
    repairResult.watches && (workbook.formulaWatches = repairResult.watches);
    for (const change of repairResult.changes) recordMetadataChange(change);
    for (const reviewThread of book.reviewThreads ?? [])
      !reviewThread.orphaned &&
        reviewThread.anchor.sheetId === targetSheetId &&
        isCellInsideRange(targetRect, _t(reviewThread.anchor.address)) &&
        !(
          reviewThread.anchor.sheetId === sourceSheetId &&
          isCellInsideRange(from, _t(reviewThread.anchor.address))
        ) &&
        throwWorkbookError(
          "The destination contains an existing review discussion. Choose a destination without anchored discussions to preserve its evidence context.",
        );
    const workingSourceSheet = findSheetOrThrow(workbook, sourceSheetId),
      workingTargetSheet = findSheetOrThrow(workbook, targetSheetId);
    for (let firstRow = from.firstRow; firstRow <= from.lastRow; firstRow++)
      for (let firstCol = from.firstCol; firstCol <= from.lastCol; firstCol++) {
        const cellAddress = formatCellAddress(firstRow, firstCol);
        hasCellContent(sheetOrThrow.cells[cellAddress]) &&
          delete workingSourceSheet.cells[cellAddress];
      }
    for (let firstRow = from.firstRow; firstRow <= from.lastRow; firstRow++)
      for (let firstCol = from.firstCol; firstCol <= from.lastCol; firstCol++) {
        const cellAddress = formatCellAddress(firstRow, firstCol),
          targetAddress = formatCellAddress(
            firstRow + move.dr,
            firstCol + move.dc,
          ),
          cell = sheetOrThrow.cells[cellAddress],
          existingTargetCell = targetSheet.cells[targetAddress];
        (!(
          sourceSheetId === targetSheetId &&
          isCellInsideRange(from, _t(targetAddress))
        ) &&
          hasCellContent(existingTargetCell) &&
          JSON.stringify(existingTargetCell) !== JSON.stringify(cell) &&
          overwrittenCells++,
          hasCellContent(cell)
            ? (workingTargetSheet.cells[targetAddress] =
                normalizeCellRecordInput(cell))
            : hasCellContent(workingTargetSheet.cells[targetAddress]) &&
              delete workingTargetSheet.cells[targetAddress],
          sourceAddressByTarget.set(
            `${targetSheetId}:${targetAddress}`,
            cellAddress,
          ));
      }
    let populatedCellCount = 0;
    for (const sheet of workbook.sheets)
      for (const [populatedAddress, cell] of Object.entries(sheet.cells)) {
        ++populatedCellCount > workbookLimits.maxCells &&
          throwWorkbookError(
            "Range moves inspect at most 100,000 populated cells.",
          );
        const isMovedCell = sourceAddressByTarget.has(
          `${sheet.id}:${populatedAddress}`,
        );
        if (isFormulaInputValue(cell)) {
          const repairedFormula = repairReferencesForMove(
            move,
            cell.raw,
            isMovedCell ? sourceSheetId : sheet.id,
            sheet.id,
            isMovedCell,
          );
          repairedFormula !== cell.raw &&
            ((cell.raw = repairedFormula), referenceChanges++);
        }
        const hyperlink = cell.hyperlink;
        if (
          hyperlink?.kind === "internal" &&
          hyperlink.sheetId === sourceSheetId &&
          isCellInsideRange(from, _t(hyperlink.address))
        ) {
          const hyperlinkPosition = _t(hyperlink.address),
            before = `${sheetOrThrow.name}!${hyperlink.address}`;
          ((cell.hyperlink = {
            ...hyperlink,
            sheetId: targetSheetId,
            address: formatCellAddress(
              hyperlinkPosition.row + move.dr,
              hyperlinkPosition.col + move.dc,
            ),
          }),
            JSON.stringify(hyperlink) !== JSON.stringify(cell.hyperlink) &&
              (referenceChanges++,
              recordMetadataChange({
                kind: "internal-link",
                label: `${sheet.name}!${populatedAddress}`,
                before: before,
                after: `${targetSheet.name}!${cell.hyperlink.address}`,
              })));
        }
      }
    for (const sheet of workbook.sheets)
      for (const table of sheet.tables ?? [])
        for (const column of table.columns)
          if (column.calculatedFormula) {
            const after = repairReferencesForMove(
              move,
              column.calculatedFormula,
              sheet.id,
              sheet.id,
            );
            after !== column.calculatedFormula &&
              (assertSheetUnprotected(findSheetOrThrow(book, sheet.id)),
              recordMetadataChange({
                kind: "table-column",
                label: `${sheet.name}: ${table.name}[${column.name}] calculated rule`,
                before: column.calculatedFormula,
                after: after,
              }),
              (column.calculatedFormula = after),
              referenceChanges++);
          }
    for (const definedName of workbook.definedNames ?? []) {
      const after = repairReferencesForMove(
        move,
        definedName.refersTo,
        definedName.scopeSheetId,
        definedName.scopeSheetId,
      );
      after !== definedName.refersTo &&
        (definedName.scopeSheetId &&
          assertSheetUnprotected(
            findSheetOrThrow(book, definedName.scopeSheetId),
          ),
        recordMetadataChange({
          kind: "defined-name",
          label: definedName.name,
          before: definedName.refersTo,
          after: after,
        }),
        (definedName.refersTo = after),
        referenceChanges++);
    }
    for (const workingReviewThread of workbook.reviewThreads ?? [])
      if (
        !workingReviewThread.orphaned &&
        workingReviewThread.anchor.sheetId === sourceSheetId &&
        isCellInsideRange(from, _t(workingReviewThread.anchor.address))
      ) {
        const anchorPosition = _t(workingReviewThread.anchor.address),
          before = `${workingReviewThread.anchor.sheetName}!${workingReviewThread.anchor.address}`;
        ((workingReviewThread.anchor = {
          sheetId: targetSheetId,
          sheetName: targetSheet.name,
          address: formatCellAddress(
            anchorPosition.row + move.dr,
            anchorPosition.col + move.dc,
          ),
        }),
          recordMetadataChange({
            kind: "review-anchor",
            label: workingReviewThread.id,
            before: before,
            after: `${targetSheet.name}!${workingReviewThread.anchor.address}`,
          }));
      }
    for (const sheet of workbook.sheets) {
      for (const pivotOrQuery of [
        ...(sheet.pivots ?? []),
        ...(sheet.queries ?? []),
      ]) {
        const remappedSource = remapReferenceForMove(
          move,
          pivotOrQuery.source.sheetId,
          pivotOrQuery.source.range,
        );
        remappedSource.changed &&
          (assertSheetUnprotected(findSheetOrThrow(book, sheet.id)),
          recordMetadataChange({
            kind: "analysis-source",
            label: `${sheet.name}: ${pivotOrQuery.name}`,
            before: `${findSheetOrThrow(book, pivotOrQuery.source.sheetId).name}!${pivotOrQuery.source.range}`,
            after: `${findSheetOrThrow(book, remappedSource.sheetId).name}!${remappedSource.reference}`,
          }),
          (pivotOrQuery.source = {
            ...pivotOrQuery.source,
            sheetId: remappedSource.sheetId,
            range: remappedSource.reference,
          }),
          referenceChanges++);
      }
      for (const drawingObject of sheet.objects ?? [])
        if (drawingObject.kind === "chart" && drawingObject.range) {
          const remappedChartRange = remapReferenceForMove(
            move,
            sheet.id,
            drawingObject.range,
          );
          remappedChartRange.changed &&
            (remappedChartRange.sheetId !== sheet.id &&
              throwWorkbookError(
                "A chart source cannot follow data to another worksheet in this implementation. Move the chart data on its current sheet.",
              ),
            assertSheetUnprotected(findSheetOrThrow(book, sheet.id)),
            recordMetadataChange({
              kind: "chart-source",
              label: drawingObject.label || drawingObject.id,
              before: drawingObject.range,
              after: remappedChartRange.reference,
            }),
            (drawingObject.range = remappedChartRange.reference),
            referenceChanges++);
        }
    }
  }
  const previewChanges = [];
  let dependentCount = 0;
  if (!isNoOpMove)
    for (const sheet of workbook.sheets) {
      const sheetOrThrow = findSheetOrThrow(book, sheet.id),
        addressSet = new Set([
          ...Object.keys(sheetOrThrow.cells),
          ...Object.keys(sheet.cells),
        ]);
      for (const address of addressSet) {
        const cell = sheetOrThrow.cells[address],
          afterCell = sheet.cells[address];
        JSON.stringify(cell) !== JSON.stringify(afterCell) &&
          (assertSheetUnprotected(sheetOrThrow),
          previewChanges.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            address: address,
            ...(cell ? { before: normalizeCellRecordInput(cell) } : {}),
            ...(afterCell
              ? { after: normalizeCellRecordInput(afterCell) }
              : {}),
          }),
          previewChanges.length > pasteOperationLimits.changedCells &&
            throwWorkbookError(
              "A move may change at most 500 cells in total, including source clears, destination cells, and dependent references. Move a smaller block.",
            ),
          !(
            sheet.id === sourceSheetId && isCellInsideRange(from, _t(address))
          ) &&
            !(
              sheet.id === targetSheetId &&
              isCellInsideRange(targetRect, _t(address))
            ) &&
            dependentCount++);
      }
    }
  const hasChanges = previewChanges.length > 0 || metadataChanges.length > 0;
  hasChanges &&
    (formulaEntriesSignature(book) !== formulaEntriesSignature(workbook) &&
      throwWorkbookError(
        "The move would create, remove, or resize a dynamic-array footprint. Convert the affected array to ordinary values before moving cells.",
      ),
    (book.events.length >= 2e3 ||
      book.events.reduce(
        (total, historyEvent) => total + historyEvent.changes.length,
        0,
      ) +
        previewChanges.length >
        workbookLimits.maxCells) &&
      throwWorkbookError(
        "The move exceeds workbook history limits. Export the workbook before starting a new workpaper.",
      ),
    book.revision >= Number.MAX_SAFE_INTEGER - 1 &&
      throwWorkbookError("Workbook revision limit reached."),
    checkDataValidationBatch(
      workbook,
      previewChanges.filter(
        (previewChange) =>
          previewChange.before?.raw !== previewChange.after?.raw ||
          previewChange.before?.literal !== previewChange.after?.literal,
      ),
    ));
  const nextWorkbook = hasChanges ? normalizeWorkbookInput(workbook) : book,
    patches = previewChanges.map((previewChange) => ({
      sheetId: previewChange.sheetId,
      address: previewChange.address,
      cell: normalizeCellRecordInput(previewChange.after ?? { raw: null }),
    })),
    planDraft = {
      workbookId: book.id,
      revision: book.revision,
      snapshotGuard: sourceSnapshot.snapshotGuard,
      sourceSheetId: sourceSheetId,
      sourceRange: sortRangeAddress,
      sourceSnapshot: { ...sourceSnapshot },
      targetSheetId: targetSheetId,
      targetRange: targetRangeAddress,
      targetCount: sourceRowCount * sourceColumnCount,
      patches: patches,
      previewPatches: patches,
      count: previewChanges.length,
      dependentCount: dependentCount,
      overwrittenCells: overwrittenCells,
      referenceChanges: referenceChanges,
      metadataChanges: metadataChanges,
      previewChanges: previewChanges,
      summary: `Move ${sheetOrThrow.name}!${sortRangeAddress} to ${targetSheet.name}!${targetRangeAddress}: ${previewChanges.length} cell changes, ${overwrittenCells} destination cells overwritten, ${metadataChanges.length} metadata changes.`,
    };
  return {
    plan: eI(planDraft),
    next: hasChanges
      ? { ...nextWorkbook, revision: book.revision, events: book.events }
      : book,
  };
}
function buildRangeMovePlan(
  workbook,
  sheetId,
  range,
  targetSheetId,
  targetRange,
  options = {},
) {
  return computeRangeMoveResult(
    workbook,
    sheetId,
    range,
    targetSheetId,
    targetRange,
    options,
  ).plan;
}
function applyRangeMovePreview(workbook, plan, options = {}) {
  (assertWorkbookWritable(options.readOnly),
    (!plan ||
      plan.workbookId !== workbook.id ||
      plan.revision !== workbook.revision ||
      plan.snapshotGuard !== JSON.stringify(workbook)) &&
      throwWorkbookError(
        "The workbook changed after this move preview. Cut the source again before moving it.",
      ));
  const rangeMoveResult = computeRangeMoveResult(
    workbook,
    plan.sourceSheetId,
    plan.sourceRange,
    plan.targetSheetId,
    plan.targetRange,
    { sourceSnapshot: plan.sourceSnapshot, readOnly: options.readOnly },
  );
  return (
    (JSON.stringify(rangeMoveResult.plan.previewChanges) !==
      JSON.stringify(plan.previewChanges) ||
      JSON.stringify(rangeMoveResult.plan.metadataChanges) !==
        JSON.stringify(plan.metadataChanges)) &&
      throwWorkbookError(
        "The move no longer matches its preview. Cut the source again.",
      ),
    rangeMoveResult.next
  );
}
function describeCellContents(cell) {
  if (!cell) return "(empty)";
  const parts = [String(cell.raw ?? "(empty)")];
  return (
    cell.literal && parts.push("Literal text"),
    cell.note !== undefined && parts.push("Note: " + cell.note),
    cell.tickmark !== undefined && parts.push("Tickmark: " + cell.tickmark),
    cell.hyperlink &&
      parts.push(
        "Link: " +
          (cell.hyperlink.kind === "external"
            ? cell.hyperlink.url
            : cell.hyperlink.sheetId + "!" + cell.hyperlink.address) +
          (cell.hyperlink.label ? " (" + cell.hyperlink.label + ")" : ""),
      ),
    cell.numberFormat
      ? parts.push(
          "Number display: " +
            Object.entries(cell.numberFormat)
              .map(
                ([formatKey, formatValue]) => formatKey + " = " + formatValue,
              )
              .join(", "),
        )
      : (cell.numFmt || cell.format) &&
        parts.push("Number display: " + (cell.numFmt ?? cell.format)),
    cell.style &&
      Object.keys(cell.style).length &&
      parts.push(
        "Cell style: " +
          Object.entries(cell.style)
            .map(([styleKey, styleValue]) => styleKey + " = " + styleValue)
            .join(", "),
      ),
    parts.join(" · ")
  );
}
const staleMovePreviewMessage =
  "The workbook or destination changed. Cancel and cut the current cells again.";
function MoveCellsDialog(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    propsRef = ne.useRef(props);
  propsRef.current = props;
  const snapshotRef = ne.useRef({
      sheetId: props.sheetId,
      range: props.range,
      source: props.source,
      sourceName:
        props.workbook.sheets.find((sheet) => sheet.id === props.source.sheetId)
          ?.name ?? "Deleted worksheet",
      targetName:
        props.workbook.sheets.find((sheet) => sheet.id === props.sheetId)
          ?.name ?? "Deleted worksheet",
      book: JSON.stringify(props.workbook),
    }),
    [errorMessage, setErrorMessage] = ne.useState(""),
    [pageIndex, setPageIndex] = ne.useState(0),
    isStale =
      snapshotRef.current.book !== JSON.stringify(props.workbook) ||
      snapshotRef.current.sheetId !== props.sheetId ||
      snapshotRef.current.range !== props.range ||
      snapshotRef.current.source !== props.source,
    [planResult] = ne.useState(() => {
      try {
        return {
          plan: buildRangeMovePlan(
            props.workbook,
            props.source.sheetId,
            props.source.range,
            props.sheetId,
            props.range,
            { sourceSnapshot: props.source, readOnly: props.readOnly },
          ),
          error: "",
        };
      } catch (error) {
        return {
          plan: null,
          error:
            error instanceof Error
              ? error.message
              : "The cells cannot be moved.",
        };
      }
    }),
    previewRows = [
      ...(planResult.plan?.previewChanges ?? []).map((previewChange) => ({
        label: `${previewChange.sheetName}!${previewChange.address}`,
        before: describeCellContents(previewChange.before),
        after: describeCellContents(previewChange.after),
      })),
      ...(planResult.plan?.metadataChanges ?? []).map((metadataChange) => ({
        label: metadataChange.label,
        before: metadataChange.before,
        after: metadataChange.after,
      })),
    ],
    pageCount = Math.max(1, Math.ceil(previewRows.length / 15)),
    currentPage = Math.min(pageIndex, pageCount - 1);
  ne.useEffect(() => {
    const activeElement = document.activeElement;
    return (
      dialogRef.current?.querySelector("button:not(:disabled)")?.focus(),
      () => {
        activeElement?.isConnected && activeElement.focus();
      }
    );
  }, []);
  function onClick() {
    try {
      const currentProps = propsRef.current;
      if (currentProps.readOnly) throw Error("This workbook is read-only.");
      if (
        snapshotRef.current.book !== JSON.stringify(currentProps.workbook) ||
        snapshotRef.current.sheetId !== currentProps.sheetId ||
        snapshotRef.current.range !== currentProps.range ||
        snapshotRef.current.source !== currentProps.source
      )
        throw Error(staleMovePreviewMessage);
      if (!planResult.plan) throw Error(planResult.error);
      const movedWorkbook = applyRangeMovePreview(
        currentProps.workbook,
        planResult.plan,
        {
          readOnly: currentProps.readOnly,
        },
      );
      (currentProps.onWorkbookChange(
        movedWorkbook,
        "Move cells and repair references",
      ),
        currentProps.onSelectionChange(planResult.plan.targetRange),
        currentProps.onApplied(),
        currentProps.onClose());
    } catch (error) {
      setErrorMessage(
        error instanceof Error
          ? error.message
          : "The cells could not be moved.",
      );
    }
  }
  return i.jsx("div", {
    className: "et-backdrop",
    children: i.jsxs("section", {
      ref: dialogRef,
      className: "et-dialog",
      role: "dialog",
      "aria-modal": "true",
      "aria-labelledby": titleId,
      tabIndex: -1,
      onKeyDown: (event) => {
        if (
          (event.key === "Escape" &&
            (event.preventDefault(),
            event.stopPropagation(),
            propsRef.current.onClose()),
          event.key === "Tab")
        ) {
          const focusableElements = Array.from(
              dialogRef.current?.querySelectorAll(
                'button:not(:disabled),input:not(:disabled),select:not(:disabled),[tabindex="0"]',
              ) ?? [],
            ),
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
      children: [
        i.jsxs("header", {
          className: "et-header",
          children: [
            i.jsxs("h2", {
              id: titleId,
              children: [
                i.jsx(ArrowRight, { size: 19 }),
                i.jsx(LocalizedText, { children: "Move cells" }),
              ],
            }),
            i.jsx("button", {
              type: "button",
              "aria-label": text("Close move preview"),
              onClick: props.onClose,
              lang: language("Close move preview"),
              "data-ui-label": "Close move preview",
              children: i.jsx(X, { size: 18 }),
            }),
          ],
        }),
        i.jsxs("div", {
          className: "et-body",
          children: [
            i.jsx("p", {
              children: i.jsx(LocalizedText, {
                children:
                  "The source stays in place until you apply this move. The preview includes overwritten cells, formula repairs and review metadata.",
              }),
            }),
            i.jsxs("p", {
              children: [
                i.jsxs("strong", {
                  children: [
                    snapshotRef.current.sourceName,
                    "!",
                    snapshotRef.current.source.range,
                  ],
                }),
                " ",
                "→",
                " ",
                i.jsxs("strong", {
                  children: [
                    snapshotRef.current.targetName,
                    "!",
                    planResult.plan?.targetRange ?? snapshotRef.current.range,
                  ],
                }),
              ],
            }),
            planResult.plan &&
              i.jsxs("section", {
                className: "et-preview",
                "aria-label": text("Move preview"),
                lang: language("Move preview"),
                "data-ui-label": "Move preview",
                children: [
                  i.jsxs("p", {
                    children: [
                      planResult.plan.count,
                      " ",
                      i.jsx(LocalizedText, { children: "cell changes ·" }),
                      " ",
                      planResult.plan.overwrittenCells,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "occupied destinations overwritten · ",
                      }),
                      planResult.plan.referenceChanges,
                      " ",
                      i.jsx(LocalizedText, {
                        children: "dependent references updated · ",
                      }),
                      planResult.plan.metadataChanges.length,
                      " ",
                      i.jsx(LocalizedText, { children: "metadata changes" }),
                    ],
                  }),
                  i.jsxs("table", {
                    className: "et-preview-table",
                    children: [
                      i.jsx("thead", {
                        children: i.jsxs("tr", {
                          children: [
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Location",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "Before",
                              }),
                            }),
                            i.jsx("th", {
                              children: i.jsx(LocalizedText, {
                                children: "After",
                              }),
                            }),
                          ],
                        }),
                      }),
                      i.jsx("tbody", {
                        children: previewRows
                          .slice(currentPage * 15, currentPage * 15 + 15)
                          .map((previewRow, rowIndex) =>
                            i.jsxs(
                              "tr",
                              {
                                children: [
                                  i.jsx("th", { children: previewRow.label }),
                                  i.jsx("td", {
                                    children: String(previewRow.before),
                                  }),
                                  i.jsx("td", {
                                    children: String(previewRow.after),
                                  }),
                                ],
                              },
                              rowIndex,
                            ),
                          ),
                      }),
                    ],
                  }),
                  pageCount > 1 &&
                    i.jsxs("div", {
                      className: "et-pagination",
                      children: [
                        i.jsx("button", {
                          type: "button",
                          disabled: !currentPage,
                          onClick: () => setPageIndex(currentPage - 1),
                          children: i.jsx(LocalizedText, {
                            children: "Previous changes",
                          }),
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
                          type: "button",
                          disabled: currentPage + 1 >= pageCount,
                          onClick: () => setPageIndex(currentPage + 1),
                          children: i.jsx(LocalizedText, {
                            children: "Next changes",
                          }),
                        }),
                      ],
                    }),
                  !previewRows.length &&
                    i.jsx("p", {
                      children: i.jsx(LocalizedText, {
                        children: "Source and destination already match.",
                      }),
                    }),
                ],
              }),
            (errorMessage || isStale || planResult.error) &&
              i.jsx("p", {
                className: "et-error",
                role: "alert",
                children:
                  errorMessage ||
                  (isStale ? staleMovePreviewMessage : planResult.error),
              }),
            props.readOnly &&
              i.jsx("p", {
                role: "status",
                children: i.jsx(LocalizedText, {
                  children: "This workbook is read-only.",
                }),
              }),
          ],
        }),
        i.jsxs("footer", {
          className: "et-footer",
          children: [
            i.jsx("button", {
              type: "button",
              onClick: props.onClose,
              children: i.jsx(LocalizedText, { children: "Cancel" }),
            }),
            i.jsx("button", {
              type: "button",
              className: "et-primary",
              disabled:
                props.readOnly ||
                isStale ||
                !planResult.plan ||
                !previewRows.length,
              onClick: onClick,
              children: i.jsx(LocalizedText, { children: "Apply move" }),
            }),
          ],
        }),
      ],
    }),
  });
}
const numberFormatCategoryLabels = {
    general: "General",
    number: "Number",
    currency: "Currency",
    accounting: "Accounting",
    percent: "Percentage",
    date: "Date",
    time: "Time",
    datetime: "Date and time",
    duration: "Elapsed time",
    custom: "Custom",
  },
  localeDisplayNames = {
    "en-US": "English (United States)",
    "en-GB": "English (United Kingdom)",
    "de-DE": "German (Germany)",
    "fr-FR": "French (France)",
    "ja-JP": "Japanese (Japan)",
    "en-IN": "English (India)",
  };
function resolveSelectionNumberFormat(options) {
  const rangeAddress = parseRangeAddress(options.range),
    firstCellAddress = `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    targetAddress =
      options.active &&
      isCellAddressInRanges(options.ranges ?? [options.range], options.active)
        ? options.active
        : firstCellAddress,
    targetCell = options.workbook.sheets.find(
      (sheet) => sheet.id === options.sheetId,
    )?.cells[targetAddress];
  if (targetCell?.numberFormat)
    try {
      return validateNumberFormatSettings(targetCell.numberFormat);
    } catch {}
  const category = targetCell?.format ?? "general";
  return validateNumberFormatSettings({
    category: category,
    locale: "en-US",
    ...(["number", "currency", "percent"].includes(category)
      ? {
          decimals: Math.min(
            12,
            summarizeNumberFormat(targetCell?.numFmt) ?? 2,
          ),
          useGrouping: targetCell?.numFmt
            ? targetCell.numFmt.includes(",")
            : category !== "percent",
        }
      : {}),
  });
}
function M4e(props) {
  const { text, language } = useWorkbookMessages(),
    titleId = ne.useId(),
    dialogRef = ne.useRef(null),
    onCloseRef = ne.useRef(props.onClose);
  onCloseRef.current = props.onClose;
  const snapshotRef = ne.useRef({
      book: props.workbook.id,
      revision: props.workbook.revision,
      content: JSON.stringify(props.workbook),
      sheet: props.sheetId,
      selection: JSON.stringify(props.ranges ?? [props.range]),
      active: props.active,
    }),
    [initialFormat] = ne.useState(() => resolveSelectionNumberFormat(props)),
    [category, setCategory] = ne.useState(initialFormat.category),
    [locale, setLocale] = ne.useState(initialFormat.locale),
    [currency, setCurrency] = ne.useState(initialFormat.currency ?? "USD"),
    [decimalsText, setDecimalsText] = ne.useState(
      String(initialFormat.decimals ?? 2),
    ),
    [shouldUseGrouping, setShouldUseGrouping] = ne.useState(
      initialFormat.useGrouping ?? true,
    ),
    [scaleText, setScaleText] = ne.useState(String(initialFormat.scale ?? 1)),
    [negativeStyle, setNegativeStyle] = ne.useState(
      initialFormat.negative ?? "minus",
    ),
    [pattern, setPattern] = ne.useState(initialFormat.pattern ?? "#,##0.00"),
    [submitError, setSubmitError] = ne.useState(""),
    sheet = props.workbook.sheets.find((sheet) => sheet.id === props.sheetId),
    selectionAreas = parseSelectionAreas(props.ranges ?? [props.range]),
    rangeAddress = parseRangeAddress(props.range),
    activeAddress =
      props.active && isCellAddressInRanges(selectionAreas, props.active)
        ? props.active
        : `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}`,
    cell = sheet?.cells[activeAddress],
    selectedCellCount = countSelectionCells(selectionAreas),
    importedNumberFormatWarning = getImportedNumberFormatWarning(cell),
    isStale =
      snapshotRef.current.book !== props.workbook.id ||
      snapshotRef.current.revision !== props.workbook.revision ||
      snapshotRef.current.content !== JSON.stringify(props.workbook) ||
      snapshotRef.current.sheet !== props.sheetId ||
      snapshotRef.current.selection !==
        JSON.stringify(props.ranges ?? [props.range]) ||
      snapshotRef.current.active !== props.active,
    isApplyDisabled =
      props.readOnly ||
      isStale ||
      !!sheet?.protection?.enabled ||
      !sheet ||
      selectedCellCount > 500,
    getCalculatedValue = ne.useMemo(
      () => calculateWorkbookWithSettings(props.workbook),
      [props.workbook],
    ),
    hasNumericOptions = [
      "number",
      "currency",
      "accounting",
      "percent",
    ].includes(category),
    hasCurrencyOption = ["currency", "accounting"].includes(category),
    isDateTimeCategory = ["date", "time", "datetime", "duration"].includes(
      category,
    );
  let formatSettings,
    formatError = "",
    currentValue = null;
  try {
    ((currentValue = getCalculatedValue(props.sheetId, activeAddress)),
      (formatSettings = validateNumberFormatSettings({
        category: category,
        locale: locale,
        ...(hasNumericOptions
          ? {
              decimals: decimalsText.trim() === "" ? NaN : Number(decimalsText),
              useGrouping: shouldUseGrouping,
              scale: category === "percent" ? 1 : Number(scaleText),
              negative:
                category === "accounting" ? "parentheses" : negativeStyle,
              ...(hasCurrencyOption ? { currency: currency } : {}),
            }
          : isDateTimeCategory || category === "custom"
            ? { pattern: pattern }
            : {}),
      })));
  } catch (error) {
    formatError =
      error instanceof Error
        ? error.message
        : "The format could not be previewed.";
  }
  const X$1 = [
      {
        name: `${sheet?.name ?? "Worksheet"}!${activeAddress}`,
        value: currentValue,
      },
      ...(isDateTimeCategory ||
      (category === "custom" &&
        /[ymdhs]/i.test(pattern.replace(/"[^"]*"/g, "")))
        ? [
            { name: "Date / time example", value: 46273.625 },
            { name: "Elapsed-time example", value: 1.5 },
          ]
        : [
            { name: "Positive example", value: 1234.567 },
            { name: "Negative example", value: -1234.567 },
            { name: "Zero example", value: 0 },
          ]),
    ],
    selectCategory = (nextCategory) => {
      (setCategory(nextCategory),
        setSubmitError(""),
        nextCategory in builtInDateTimeFormats &&
          setPattern(builtInDateTimeFormats[nextCategory][0]),
        nextCategory === "custom" &&
          setPattern(
            cell?.numberFormat?.category === "custom"
              ? cell.numberFormat.pattern
              : '#,##0.00;(#,##0.00);"-"',
          ));
    };
  return (
    ne.useEffect(() => {
      const activeElement = document.activeElement;
      return (
        dialogRef.current
          ?.querySelector('[role="option"][aria-selected="true"]')
          ?.focus(),
        () => {
          activeElement?.isConnected && activeElement.focus();
        }
      );
    }, []),
    i.jsx("div", {
      className: "nf-backdrop",
      children: i.jsxs("div", {
        className: "nf-dialog",
        role: "dialog",
        "aria-modal": "true",
        "aria-labelledby": `${titleId}-title`,
        ref: dialogRef,
        onKeyDown: (event) => {
          if (
            (event.stopPropagation(),
            event.key === "Escape" &&
              (event.preventDefault(), onCloseRef.current()),
            event.key === "Tab")
          ) {
            const focusableElements = [
                ...(dialogRef.current?.querySelectorAll(
                  'button:not(:disabled):not([tabindex="-1"]),input:not(:disabled),select:not(:disabled),textarea:not(:disabled),[tabindex="0"]',
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
        children: [
          i.jsxs("header", {
            className: "nf-header",
            children: [
              i.jsx("h2", {
                id: `${titleId}-title`,
                children: i.jsx(LocalizedText, { children: "Format Cells" }),
              }),
              i.jsx("button", {
                "aria-label": text("Close number formats"),
                onClick: props.onClose,
                lang: language("Close number formats"),
                "data-ui-label": "Close number formats",
                children: i.jsx(X, { size: 18 }),
              }),
            ],
          }),
          i.jsxs("div", {
            className: "nf-subtitle",
            children: [
              i.jsx(LocalizedText, { children: "Number · " }),
              selectedCellCount.toLocaleString(),
              " ",
              i.jsx(LocalizedText, { children: "selected cells · " }),
              selectionAreas.length,
              " ",
              selectionAreas.length === 1 ? "area" : "areas",
            ],
          }),
          i.jsxs("div", {
            className: "nf-content",
            children: [
              i.jsx("div", {
                className: "nf-category",
                role: "listbox",
                "aria-label": text("Number format category"),
                lang: language("Number format category"),
                "data-ui-label": "Number format category",
                children: Uv.map((categoryKey, categoryIndex) =>
                  i.jsx(
                    "button",
                    {
                      role: "option",
                      "aria-selected": category === categoryKey,
                      tabIndex: category === categoryKey ? 0 : -1,
                      onClick: () => selectCategory(categoryKey),
                      onKeyDown: (event) => {
                        if (
                          ["ArrowUp", "ArrowDown", "Home", "End"].includes(
                            event.key,
                          )
                        ) {
                          event.preventDefault();
                          const nextIndex =
                            event.key === "Home"
                              ? 0
                              : event.key === "End"
                                ? Uv.length - 1
                                : (categoryIndex +
                                    (event.key === "ArrowDown"
                                      ? 1
                                      : Uv.length - 1)) %
                                  Uv.length;
                          (selectCategory(Uv[nextIndex]),
                            event.currentTarget.parentElement?.children[
                              nextIndex
                            ]?.focus());
                        }
                      },
                      children: numberFormatCategoryLabels[categoryKey],
                    },
                    categoryKey,
                  ),
                ),
              }),
              i.jsxs("form", {
                className: "nf-settings",
                onSubmit: (event) => {
                  (event.preventDefault(), setSubmitError(""));
                  try {
                    if (props.readOnly)
                      throw new Error("This workbook is read-only.");
                    if (isStale)
                      throw new Error(
                        "The workbook or selection changed. Close and reopen Format Cells.",
                      );
                    if (!formatSettings)
                      throw new Error(
                        formatError || "Choose a valid number format.",
                      );
                    const nextWorkbook = applyNumberFormatTransaction(
                      props.workbook,
                      props.sheetId,
                      selectionAreas,
                      formatSettings,
                      {
                        readOnly: props.readOnly,
                      },
                    );
                    (nextWorkbook !== props.workbook &&
                      props.onWorkbookChange(nextWorkbook, "Format numbers"),
                      props.onClose());
                  } catch (error) {
                    setSubmitError(
                      error instanceof Error
                        ? error.message
                        : "The format could not be applied.",
                    );
                  }
                },
                children: [
                  isStale &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: i.jsx(LocalizedText, {
                        children:
                          "The workbook or selection changed. Close and reopen Format Cells.",
                      }),
                    }),
                  props.readOnly &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "This workbook is read-only. You can inspect format previews.",
                      }),
                    }),
                  sheet?.protection?.enabled &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: i.jsx(LocalizedText, {
                        children:
                          "This worksheet is protected. Unprotect it before changing cell formats.",
                      }),
                    }),
                  selectedCellCount > 500 &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Select at most 500 unique cells to apply number formats.",
                      }),
                    }),
                  importedNumberFormatWarning &&
                    i.jsx("p", {
                      className: "nf-notice",
                      children: importedNumberFormatWarning,
                    }),
                  i.jsxs("section", {
                    className: "nf-preview",
                    "aria-label": text("Number format preview"),
                    lang: language("Number format preview"),
                    "data-ui-label": "Number format preview",
                    children: [
                      i.jsx("h3", {
                        children: i.jsx(LocalizedText, { children: "Sample" }),
                      }),
                      i.jsxs("table", {
                        children: [
                          i.jsx("thead", {
                            children: i.jsxs("tr", {
                              children: [
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Example",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Stored value",
                                  }),
                                }),
                                i.jsx("th", {
                                  children: i.jsx(LocalizedText, {
                                    children: "Display",
                                  }),
                                }),
                              ],
                            }),
                          }),
                          i.jsx("tbody", {
                            children: X$1.map((sample) => {
                              const formattedSample = formatSettings
                                ? formatCellValueWithNumberFormat(
                                    sample.value,
                                    formatSettings,
                                    props.workbook.datePolicy?.dateSystem,
                                  )
                                : undefined;
                              return i.jsxs(
                                "tr",
                                {
                                  children: [
                                    i.jsx("th", {
                                      scope: "row",
                                      children: sample.name,
                                    }),
                                    i.jsx("td", {
                                      children:
                                        sample.value === null
                                          ? "Blank"
                                          : String(sample.value),
                                    }),
                                    i.jsx("td", {
                                      children: formattedSample?.accounting
                                        ? i.jsxs("span", {
                                            className: `nf-accounting nf-symbol-${formattedSample.accounting.symbolPosition}`,
                                            children: [
                                              i.jsx("span", {
                                                children:
                                                  formattedSample.accounting
                                                    .symbol,
                                              }),
                                              i.jsx("span", {
                                                children:
                                                  formattedSample.accounting
                                                    .amount,
                                              }),
                                            ],
                                          })
                                        : (formattedSample?.text ??
                                          "Unsupported pattern"),
                                    }),
                                  ],
                                },
                                sample.name,
                              );
                            }),
                          }),
                        ],
                      }),
                      i.jsxs("p", {
                        children: [
                          i.jsx(LocalizedText, {
                            children: "Current display:",
                          }),
                          " ",
                          i.jsx("strong", {
                            children:
                              formatNumberWithPattern(
                                currentValue,
                                cell,
                                props.workbook.datePolicy?.dateSystem,
                              ) || "Blank",
                          }),
                          i.jsx(LocalizedText, {
                            children:
                              ". Formatting changes appearance only; stored values and formulas stay unchanged.",
                          }),
                        ],
                      }),
                    ],
                  }),
                  i.jsxs("label", {
                    children: [
                      i.jsx(LocalizedText, { children: "Locale" }),
                      i.jsx("select", {
                        "aria-label": text("Number format locale"),
                        value: locale,
                        onChange: (event) => setLocale(event.target.value),
                        lang: language("Number format locale"),
                        "data-ui-label": "Number format locale",
                        children: supportedNumberLocales.map(
                          (supportedNumberLocale) =>
                            i.jsx(
                              "option",
                              {
                                value: supportedNumberLocale,
                                children:
                                  localeDisplayNames[supportedNumberLocale],
                              },
                              supportedNumberLocale,
                            ),
                        ),
                      }),
                    ],
                  }),
                  hasNumericOptions &&
                    i.jsxs("div", {
                      className: "nf-fields",
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, {
                              children: "Decimal places",
                            }),
                            i.jsx(LocalizedNumberField, {
                              "aria-label": text("Decimal places"),
                              min: 0,
                              max: 12,
                              step: 1,
                              value: decimalsText,
                              onValueChange: (nextValue) =>
                                setDecimalsText(nextValue),
                              lang: language("Decimal places"),
                              "data-ui-label": "Decimal places",
                            }),
                          ],
                        }),
                        hasCurrencyOption &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, { children: "Currency" }),
                              i.jsx("select", {
                                "aria-label": text("Currency"),
                                value: currency,
                                onChange: (event) => {
                                  setCurrency(event.target.value);
                                },
                                lang: language("Currency"),
                                "data-ui-label": "Currency",
                                children: supportedCurrencyCodes.map(
                                  (supportedCurrencyCode) =>
                                    i.jsx(
                                      "option",
                                      {
                                        value: supportedCurrencyCode,
                                        children: supportedCurrencyCode,
                                      },
                                      supportedCurrencyCode,
                                    ),
                                ),
                              }),
                            ],
                          }),
                        category !== "percent" &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Display scale",
                              }),
                              i.jsxs("select", {
                                "aria-label": text("Display scale"),
                                value: scaleText,
                                onChange: (event) =>
                                  setScaleText(event.target.value),
                                lang: language("Display scale"),
                                "data-ui-label": "Display scale",
                                children: [
                                  i.jsx("option", {
                                    value: "1",
                                    children: "Units",
                                  }),
                                  i.jsx("option", {
                                    value: "1000",
                                    children: "Thousands (K)",
                                  }),
                                  i.jsx("option", {
                                    value: "1000000",
                                    children: "Millions (M)",
                                  }),
                                ],
                              }),
                            ],
                          }),
                        category !== "accounting" &&
                          i.jsxs("label", {
                            children: [
                              i.jsx(LocalizedText, {
                                children: "Negative numbers",
                              }),
                              i.jsxs("select", {
                                "aria-label": text("Negative numbers"),
                                value: negativeStyle,
                                onChange: (event) =>
                                  setNegativeStyle(event.target.value),
                                lang: language("Negative numbers"),
                                "data-ui-label": "Negative numbers",
                                children: [
                                  i.jsx("option", {
                                    value: "minus",
                                    children: "Minus sign",
                                  }),
                                  i.jsx("option", {
                                    value: "parentheses",
                                    children: "Parentheses",
                                  }),
                                ],
                              }),
                            ],
                          }),
                        i.jsxs("label", {
                          className: "nf-check",
                          children: [
                            i.jsx("input", {
                              type: "checkbox",
                              "aria-label": text("Use thousands separator"),
                              checked: shouldUseGrouping,
                              onChange: (event) =>
                                setShouldUseGrouping(event.target.checked),
                              lang: language("Use thousands separator"),
                              "data-ui-label": "Use thousands separator",
                            }),
                            i.jsx(LocalizedText, {
                              children: "Use thousands separator",
                            }),
                          ],
                        }),
                      ],
                    }),
                  isDateTimeCategory &&
                    i.jsxs("label", {
                      children: [
                        i.jsx(LocalizedText, { children: "Type" }),
                        i.jsxs("select", {
                          "aria-label": text("Date or time format"),
                          value: pattern,
                          onChange: (event) => setPattern(event.target.value),
                          lang: language("Date or time format"),
                          "data-ui-label": "Date or time format",
                          children: [
                            !builtInDateTimeFormats[category].some(
                              (formatOption) => formatOption === pattern,
                            ) &&
                              i.jsx("option", {
                                value: pattern,
                                children: pattern,
                              }),
                            builtInDateTimeFormats[category].map(
                              (formatOption) =>
                                i.jsx(
                                  "option",
                                  {
                                    value: formatOption,
                                    children: formatOption,
                                  },
                                  formatOption,
                                ),
                            ),
                          ],
                        }),
                      ],
                    }),
                  category === "custom" &&
                    i.jsxs(i.Fragment, {
                      children: [
                        i.jsxs("label", {
                          children: [
                            i.jsx(LocalizedText, { children: "Type" }),
                            i.jsx("input", {
                              "aria-label": text("Custom number format"),
                              value: pattern,
                              maxLength: 200,
                              onChange: (event) =>
                                setPattern(event.target.value),
                              lang: language("Custom number format"),
                              "data-ui-label": "Custom number format",
                            }),
                          ],
                        }),
                        i.jsx("p", {
                          className: "nf-help",
                          children: i.jsx(LocalizedText, {
                            children:
                              "Use 0/# digits, decimals, grouping, percent, quoted labels and positive;negative;zero;text sections. Date/time fields and [h], [m], [s] elapsed time are supported. Colors, conditions, fractions, exponents and spacing-fill codes require other tools or remain unsupported.",
                          }),
                        }),
                      ],
                    }),
                  category === "accounting" &&
                    i.jsx("p", {
                      className: "nf-help",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Accounting separates the currency symbol and amount, uses a dash for zero and parentheses for negative values.",
                      }),
                    }),
                  isDateTimeCategory &&
                    i.jsx("p", {
                      className: "nf-help",
                      children: i.jsx(LocalizedText, {
                        children:
                          "Dates use Excel's 1900 serial system. Pattern order is explicit; the locale changes month names and fractional separators. Elapsed time can exceed 24 hours. Negative calendar dates/times are outside this subset.",
                      }),
                    }),
                  hasNumericOptions &&
                    scaleText !== "1" &&
                    category !== "percent" &&
                    i.jsxs("p", {
                      className: "nf-notice",
                      children: [
                        i.jsx(LocalizedText, {
                          children: "Displayed amounts are divided by",
                        }),
                        " ",
                        Number(scaleText).toLocaleString(),
                        " ",
                        i.jsx(LocalizedText, { children: "and labelled" }),
                        " ",
                        scaleText === "1000" ? "K" : "M",
                        i.jsx(LocalizedText, {
                          children:
                            ". Calculations continue to use the full stored amounts.",
                        }),
                      ],
                    }),
                  formatSettings &&
                    i.jsxs("details", {
                      className: "nf-exchange",
                      children: [
                        i.jsx("summary", {
                          children: i.jsx(LocalizedText, {
                            children: "Saved format and Excel exchange",
                          }),
                        }),
                        i.jsx("code", {
                          children: buildNumberFormatCode(formatSettings),
                        }),
                        i.jsx("p", {
                          children: i.jsx(LocalizedText, {
                            children:
                              "Native JSON retains these settings. Excel exchange retains the generated number mask; locale, accounting alignment and reopening in this editor can differ. Keep native JSON for the complete format settings.",
                          }),
                        }),
                      ],
                    }),
                  (submitError || formatError) &&
                    i.jsx("p", {
                      className: "nf-error",
                      role: "alert",
                      children: submitError || formatError,
                    }),
                  i.jsxs("footer", {
                    className: "nf-actions",
                    children: [
                      i.jsx("button", {
                        type: "button",
                        onClick: props.onClose,
                        children: i.jsx(LocalizedText, { children: "Cancel" }),
                      }),
                      i.jsx("button", {
                        className: "nf-primary",
                        disabled: isApplyDisabled || !formatSettings,
                        children: i.jsx(LocalizedText, {
                          children: "Apply format",
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
    })
  );
}
const wS = 20,
  staleOutlineMessage =
    "The workbook or selection changed. Close this dialog and reopen it to review the current outline.",
  formatRowColumnSpan = (kind, startIndex, endIndex) =>
    kind === "row"
      ? "Rows " +
        (startIndex + 1) +
        (startIndex === endIndex ? "" : "–" + (endIndex + 1))
      : "Columns " +
        columnIndexToLetters(startIndex) +
        (startIndex === endIndex ? "" : "–" + columnIndexToLetters(endIndex));

export { staleOutlineMessage, MoveCellsDialog, M4e, formatRowColumnSpan, wS };
