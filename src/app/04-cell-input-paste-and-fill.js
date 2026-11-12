import {
  DateSeriesRangeError,
  dateToExcelSerial,
  excelEpochUtc,
  parseSeriesSeedDate,
  s$,
  throwDateSeriesRangeError,
} from "./06c-package-and-worksheet-parsing.js";

import {
  formatCellAddress,
  maxExcelDateSerial,
  utcDateFromParts,
} from "./06c-package-and-worksheet-parsing.js";

import { isPlainObjectValue, throwWorkbookError } from "./03-workbook-model.js";
import ne from "react";
import i from "react/jsx-runtime";
import {
  fingerprintExtensionCall,
  deepFreezeValue,
  extensionRuntimeLimits,
  tokenizeFormulaText,
  validateExtensionManifest,
  setExtensionCacheEntries,
  hashExtensionManifest,
  isDeepFrozenValue,
  getExtensionCacheEntries,
  fingerprintCanonicalValue,
  validateExtensionValue,
  dateSystemSerialOffset,
  dateSystemMaxSerial,
  listInstalledExtensionFunctions,
} from "./01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "./02-text-and-financial-functions.js";
import {
  formatNumberWithPattern,
  evaluateWorkbookFormulas,
  builtinFunctionNames,
  findSheetOrThrow,
  calculateWorkbookWithSettings,
  assertNamedStyleBindingCurrent,
  _t,
  parseRangeAddress,
  normalizeCellRecordInput,
  assertNoStopValidationIssues,
  DataValidationError,
  assertSheetUnprotected,
  applyCellPatchTransaction,
  normalizeWorkbookInput,
  columnIndexToLetters,
  shiftFormulaReferences,
} from "./03-workbook-model.js";

const formatRangeAddress = (firstRow, firstCol, rowCount, colCount) => {
  const cellAddress = formatCellAddress(firstRow, firstCol),
    endCellAddress = formatCellAddress(
      firstRow + rowCount - 1,
      firstCol + colCount - 1,
    );
  return cellAddress === endCellAddress
    ? cellAddress
    : `${cellAddress}:${endCellAddress}`;
};
function eI(value) {
  return (
    value &&
      typeof value == "object" &&
      (Object.values(value).forEach(eI), Object.freeze(value)),
    value
  );
}
function stripCellDecorations(cell) {
  const cellRecordInput = normalizeCellRecordInput(cell ?? { raw: null });
  return (
    delete cellRecordInput.note,
    delete cellRecordInput.tickmark,
    delete cellRecordInput.hyperlink,
    cellRecordInput
  );
}
function withRawCellValue(raw, cell) {
  return {
    ...cell,
    raw: raw,
    literal: typeof raw == "string" && raw.startsWith("=") ? true : undefined,
  };
}
function mergeCellDecorations(fillCellPatches, decorations) {
  return {
    ...fillCellPatches,
    ...(decorations?.note !== undefined ? { note: decorations.note } : {}),
    ...(decorations?.hyperlink !== undefined
      ? { hyperlink: structuredClone(decorations.hyperlink) }
      : {}),
    ...(decorations?.tickmark !== undefined
      ? { tickmark: decorations.tickmark }
      : {}),
  };
}
function cellContentsMatch(existingCell, cell) {
  return (
    JSON.stringify(normalizeCellRecordInput(existingCell ?? { raw: null })) ===
    JSON.stringify(normalizeCellRecordInput(cell))
  );
}
function assertWorkbookWritable(readOnly) {
  (readOnly !== undefined &&
    typeof readOnly != "boolean" &&
    throwWorkbookError("Read-only must be boolean."),
    readOnly && throwWorkbookError("This workbook is read-only."));
}
function validateClipboardSnapshot(sheet) {
  (!isPlainObjectValue(sheet) ||
    sheet.version !== 1 ||
    typeof sheet.sourceWorkbookId != "string" ||
    !sheet.sourceWorkbookId ||
    sheet.sourceWorkbookId.length > 100 ||
    typeof sheet.sourceSheetId != "string" ||
    !sheet.sourceSheetId ||
    sheet.sourceSheetId.length > 100 ||
    !Number.isSafeInteger(sheet.sourceRevision) ||
    sheet.sourceRevision < 0 ||
    !Number.isSafeInteger(sheet.rowCount) ||
    !Number.isSafeInteger(sheet.colCount) ||
    sheet.rowCount < 1 ||
    sheet.colCount < 1 ||
    sheet.rowCount * sheet.colCount > workbookLimits.maxPatches ||
    typeof sheet.sourceRange != "string" ||
    !Array.isArray(sheet.cells) ||
    sheet.cells.length !== sheet.rowCount * sheet.colCount) &&
    throwWorkbookError("Invalid or oversized internal clipboard snapshot.");
  const rangeAddress = parseRangeAddress(sheet.sourceRange);
  (rangeAddress.lastRow - rangeAddress.firstRow + 1 !== sheet.rowCount ||
    rangeAddress.lastCol - rangeAddress.firstCol + 1 !== sheet.colCount) &&
    throwWorkbookError("Clipboard dimensions do not match its source range.");
  const cells = sheet.cells.map((cell, cellIndex) => {
    (!isPlainObjectValue(cell) ||
      cell.rowOffset !== Math.floor(cellIndex / sheet.colCount) ||
      cell.colOffset !== cellIndex % sheet.colCount) &&
      throwWorkbookError(
        "Clipboard cells must form a complete row-major rectangle.",
      );
    const strippedCell = stripCellDecorations(
        normalizeCellRecordInput(cell.cell),
      ),
      raw = normalizeCellRecordInput({ raw: cell.value }).raw;
    return {
      rowOffset: cell.rowOffset,
      colOffset: cell.colOffset,
      cell: strippedCell,
      value: raw,
    };
  });
  return {
    version: 1,
    sourceWorkbookId: sheet.sourceWorkbookId,
    sourceSheetId: sheet.sourceSheetId,
    sourceRevision: sheet.sourceRevision,
    sourceRange: formatRangeAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
      sheet.rowCount,
      sheet.colCount,
    ),
    rowCount: sheet.rowCount,
    colCount: sheet.colCount,
    cells: cells,
  };
}
function buildClipboardSnapshot(
  activeWorkbook,
  sourceSheetId,
  activeSelectionRange,
) {
  const sheetOrThrow = findSheetOrThrow(activeWorkbook, sourceSheetId),
    rangeAddress = parseRangeAddress(activeSelectionRange),
    rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1,
    colCount = rangeAddress.lastCol - rangeAddress.firstCol + 1;
  (rowCount * colCount > workbookLimits.maxPatches &&
    throwWorkbookError("Copy at most 500 cells into the internal clipboard."),
    (rangeAddress.lastRow >= sheetOrThrow.rowCount ||
      rangeAddress.lastCol >= sheetOrThrow.colCount) &&
      throwWorkbookError("The copied range exceeds the worksheet."));
  const workbookWithSettings = calculateWorkbookWithSettings(activeWorkbook),
    cells = [];
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset++)
    for (let colOffset = 0; colOffset < colCount; colOffset++) {
      const cellAddress = formatCellAddress(
          rangeAddress.firstRow + rowOffset,
          rangeAddress.firstCol + colOffset,
        ),
        raw = normalizeCellRecordInput({
          raw: workbookWithSettings(sourceSheetId, cellAddress),
        }).raw;
      let cell = stripCellDecorations(sheetOrThrow.cells[cellAddress]);
      if (cell.raw === null && raw !== null) {
        const spillAnchor = workbookWithSettings.getSpillAnchor(
          sourceSheetId,
          cellAddress,
        );
        if (spillAnchor && spillAnchor !== cellAddress) {
          const anchorPosition = _t(spillAnchor);
          (anchorPosition.row < rangeAddress.firstRow ||
            anchorPosition.row > rangeAddress.lastRow ||
            anchorPosition.col < rangeAddress.firstCol ||
            anchorPosition.col > rangeAddress.lastCol) &&
            (cell = withRawCellValue(raw, cell));
        }
      }
      cells.push({
        rowOffset: rowOffset,
        colOffset: colOffset,
        cell: cell,
        value: raw,
      });
    }
  return eI({
    version: 1,
    sourceWorkbookId: activeWorkbook.id,
    sourceSheetId: sourceSheetId,
    sourceRevision: activeWorkbook.revision,
    sourceRange: formatRangeAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
      rowCount,
      colCount,
    ),
    rowCount: rowCount,
    colCount: colCount,
    cells: cells,
  });
}
function buildCellChangePreview(
  workbook,
  requestedPatches,
  previewDetails,
  readOnly,
  deferValidationWarnings = false,
) {
  (assertWorkbookWritable(readOnly),
    assertSheetUnprotected(
      findSheetOrThrow(workbook, previewDetails.targetSheetId),
    ),
    requestedPatches.length > workbookLimits.maxPatches &&
      throwWorkbookError(
        "A fill or paste operation may change at most 500 cells.",
      ));
  const patches = requestedPatches.map((patch) => ({
    sheetId: patch.sheetId,
    address: patch.address,
    cell: normalizeCellRecordInput(patch.cell),
  }));
  let previewWorkbook;
  try {
    previewWorkbook = applyCellPatchTransaction(
      workbook,
      patches,
      "Preview",
      previewDetails.summary.slice(0, 200),
    );
  } catch (error) {
    if (
      !deferValidationWarnings ||
      !(error instanceof DataValidationError) ||
      error.issues.some((issue) => issue.style === "stop")
    )
      throw error;
    previewWorkbook = applyCellPatchTransaction(
      workbook,
      patches,
      "Preview",
      previewDetails.summary.slice(0, 200),
      {
        validationApproval: assertNoStopValidationIssues(error),
      },
    );
  }
  const previewPatches =
      previewWorkbook === workbook
        ? []
        : previewWorkbook.events.at(-1).changes.map((change) => ({
            sheetId: change.sheetId,
            address: change.address,
            cell: normalizeCellRecordInput(change.after ?? { raw: null }),
          })),
    count = previewPatches.length,
    dependentCount = Math.max(0, count - patches.length);
  return eI({
    ...previewDetails,
    workbookId: workbook.id,
    revision: workbook.revision,
    patches: patches,
    previewPatches: previewPatches,
    count: count,
    dependentCount: dependentCount,
    summary:
      previewDetails.summary +
      (dependentCount
        ? ` Includes ${dependentCount} dependent cell changes.`
        : ""),
    snapshotGuard: JSON.stringify(workbook),
  });
}
function applyCellChangePreview(workbook, plan, options = {}) {
  (assertWorkbookWritable(options.readOnly),
    (!plan ||
      plan.workbookId !== workbook.id ||
      plan.revision !== workbook.revision ||
      plan.snapshotGuard !== JSON.stringify(workbook)) &&
      throwWorkbookError(
        "The workbook changed after this preview. Create a new preview before applying it.",
      ),
    assertSheetUnprotected(findSheetOrThrow(workbook, plan.targetSheetId)));
  const appliedWorkbook = applyCellPatchTransaction(
    workbook,
    plan.patches,
    options.actor ?? "Preparer",
    plan.summary.slice(0, 200),
  );
  return appliedWorkbook === workbook
    ? workbook
    : {
        ...appliedWorkbook,
        revision: workbook.revision,
        events: workbook.events,
      };
}
function buildPastePreview(
  workbook,
  sheetId,
  targetRange,
  clipboardInput,
  pasteOptions = {},
) {
  const sheet = validateClipboardSnapshot(clipboardInput),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(targetRange),
    sourceRangeAddress = parseRangeAddress(sheet.sourceRange),
    pasteMode = pasteOptions.mode ?? "all",
    pasteOperation = pasteOptions.operation ?? "none";
  ((!["all", "values", "formulas", "formats"].includes(pasteMode) ||
    !["none", "add", "subtract", "multiply", "divide"].includes(
      pasteOperation,
    )) &&
    throwWorkbookError(
      "Choose a supported paste mode and arithmetic operation.",
    ),
    [pasteOptions.transpose, pasteOptions.skipBlanks].some(
      (optionValue) =>
        optionValue !== undefined && typeof optionValue != "boolean",
    ) && throwWorkbookError("Transpose and skip blanks must be boolean."),
    pasteOperation !== "none" &&
      (pasteMode === "formats" || pasteMode === "formulas") &&
      throwWorkbookError(
        "Arithmetic paste requires All or Values mode; its results are numeric values.",
      ),
    sheet.sourceWorkbookId !== workbook.id &&
      (pasteMode === "all" || pasteMode === "formulas") &&
      sheet.cells.some(
        (cell) =>
          typeof cell.cell.raw == "string" &&
          cell.cell.raw.startsWith("=") &&
          !cell.cell.literal,
      ) &&
      throwWorkbookError(
        "Use Values paste for formulas copied from a different workbook.",
      ));
  const tileRowCount = pasteOptions.transpose ? sheet.colCount : sheet.rowCount,
    tileColCount = pasteOptions.transpose ? sheet.rowCount : sheet.colCount,
    isSingleCellTarget =
      rangeAddress.firstRow === rangeAddress.lastRow &&
      rangeAddress.firstCol === rangeAddress.lastCol,
    targetRowCount = isSingleCellTarget
      ? tileRowCount
      : rangeAddress.lastRow - rangeAddress.firstRow + 1,
    targetColCount = isSingleCellTarget
      ? tileColCount
      : rangeAddress.lastCol - rangeAddress.firstCol + 1;
  ((targetRowCount % tileRowCount || targetColCount % tileColCount) &&
    throwWorkbookError(
      "The target rectangle must be a whole multiple of the copied rows and columns.",
    ),
    targetRowCount * targetColCount > workbookLimits.maxPatches &&
      throwWorkbookError("Paste into at most 500 target cells."),
    (rangeAddress.firstRow + targetRowCount > sheetOrThrow.rowCount ||
      rangeAddress.firstCol + targetColCount > sheetOrThrow.colCount) &&
      throwWorkbookError(
        "The pasted cells would exceed the worksheet boundary.",
      ));
  const normalizedTargetRange = formatRangeAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
      targetRowCount,
      targetColCount,
    ),
    targetLastRow = rangeAddress.firstRow + targetRowCount - 1,
    targetLastCol = rangeAddress.firstCol + targetColCount - 1;
  for (const mergeRange of sheetOrThrow.merges ?? []) {
    const mergeRangeAddress = parseRangeAddress(mergeRange);
    rangeAddress.firstRow <= mergeRangeAddress.lastRow &&
      targetLastRow >= mergeRangeAddress.firstRow &&
      rangeAddress.firstCol <= mergeRangeAddress.lastCol &&
      targetLastCol >= mergeRangeAddress.firstCol &&
      !(
        targetRowCount * targetColCount === 1 &&
        rangeAddress.firstRow === mergeRangeAddress.firstRow &&
        rangeAddress.firstCol === mergeRangeAddress.firstCol
      ) &&
      throwWorkbookError(
        "Unmerge the target area before pasting a block of cells or editing a covered cell.",
      );
  }
  const patches = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  let skippedBlanks = 0;
  for (let rowOffset = 0; rowOffset < targetRowCount; rowOffset++)
    for (let colOffset = 0; colOffset < targetColCount; colOffset++) {
      const sourceRowOffset = pasteOptions.transpose
          ? colOffset % tileColCount
          : rowOffset % tileRowCount,
        sourceColOffset = pasteOptions.transpose
          ? rowOffset % tileRowCount
          : colOffset % tileColCount,
        sourceEntry =
          sheet.cells[sourceRowOffset * sheet.colCount + sourceColOffset],
        isBlankSource =
          (sourceEntry.cell.raw === null && sourceEntry.value === null) ||
          sourceEntry.cell.raw === "";
      if (pasteOptions.skipBlanks && isBlankSource) {
        skippedBlanks++;
        continue;
      }
      const cellAddress = formatCellAddress(
          rangeAddress.firstRow + rowOffset,
          rangeAddress.firstCol + colOffset,
        ),
        existingCell = sheetOrThrow.cells[cellAddress];
      let cell = stripCellDecorations(sourceEntry.cell);
      typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        (cell.raw = shiftFormulaReferences(
          cell.raw,
          rangeAddress.firstRow +
            rowOffset -
            sourceRangeAddress.firstRow -
            sourceRowOffset,
          rangeAddress.firstCol +
            colOffset -
            sourceRangeAddress.firstCol -
            sourceColOffset,
        ));
      let pastedCell;
      if (pasteMode === "all")
        pastedCell = mergeCellDecorations(cell, existingCell);
      else if (pasteMode === "formulas")
        pastedCell = {
          ...(existingCell ?? { raw: null }),
          raw: cell.raw,
          literal: cell.literal,
        };
      else if (pasteMode === "values")
        pastedCell = withRawCellValue(sourceEntry.value, {
          ...(existingCell ?? { raw: null }),
        });
      else {
        pastedCell = { ...(existingCell ?? { raw: null }) };
        for (const formatKey of ["style", "format", "numFmt", "numberFormat"])
          delete pastedCell[formatKey];
        (cell.style !== undefined && (pastedCell.style = cell.style),
          cell.format !== undefined && (pastedCell.format = cell.format),
          cell.numFmt !== undefined && (pastedCell.numFmt = cell.numFmt),
          cell.numberFormat !== undefined &&
            (pastedCell.numberFormat = structuredClone(cell.numberFormat)));
      }
      if (
        (assertNamedStyleBindingCurrent(workbook, pastedCell.style),
        pasteOperation !== "none")
      ) {
        const sourceOperand =
            sourceEntry.value === null || sourceEntry.value === ""
              ? 0
              : sourceEntry.value,
          targetValue = workbookWithSettings(sheetId, cellAddress),
          targetOperand =
            targetValue === null || targetValue === "" ? 0 : targetValue;
        ((typeof sourceOperand != "number" ||
          !Number.isFinite(sourceOperand) ||
          typeof targetOperand != "number" ||
          !Number.isFinite(targetOperand)) &&
          throwWorkbookError(
            `Arithmetic paste requires numbers or blanks at ${cellAddress}.`,
          ),
          pasteOperation === "divide" &&
            sourceOperand === 0 &&
            throwWorkbookError(
              `Arithmetic paste would divide by zero at ${cellAddress}.`,
            ));
        const raw =
          pasteOperation === "add"
            ? targetOperand + sourceOperand
            : pasteOperation === "subtract"
              ? targetOperand - sourceOperand
              : pasteOperation === "multiply"
                ? targetOperand * sourceOperand
                : targetOperand / sourceOperand;
        (Number.isFinite(raw) ||
          throwWorkbookError(
            `Arithmetic paste would overflow at ${cellAddress}.`,
          ),
          (pastedCell = { ...pastedCell, raw: raw }),
          delete pastedCell.literal);
      }
      ((pastedCell = normalizeCellRecordInput(pastedCell)),
        cellContentsMatch(existingCell, pastedCell) ||
          patches.push({
            sheetId: sheetId,
            address: cellAddress,
            cell: pastedCell,
          }));
    }
  return buildCellChangePreview(
    workbook,
    patches,
    {
      targetSheetId: sheetId,
      targetRange: normalizedTargetRange,
      targetCount: targetRowCount * targetColCount,
      sourceRange: sheet.sourceRange,
      skippedBlanks: skippedBlanks,
      summary: `Paste ${pasteMode}${pasteOptions.transpose ? " transposed" : ""}${pasteOperation !== "none" ? ` (${pasteOperation})` : ""} into ${normalizedTargetRange}.`,
    },
    pasteOptions.readOnly,
    pasteOptions.deferValidationWarnings,
  );
}
function applyCellPreview(workbook, plan, options = {}) {
  return applyCellChangePreview(workbook, plan, options);
}

function validateFillDirection(direction) {
  return (
    ["down", "right", "up", "left"].includes(direction) ||
      throwWorkbookError("Choose a supported fill direction."),
    direction
  );
}
function formatRangeBoundsAddress(rangeText) {
  const rangeAddress = parseRangeAddress(rangeText),
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

function cellContentsDiffer(existingCell, cellRecordInput) {
  return (
    JSON.stringify(normalizeCellRecordInput(existingCell ?? { raw: null })) !==
    JSON.stringify(normalizeCellRecordInput(cellRecordInput))
  );
}
function buildFillCellPatches(
  sheetOrThrow,
  cellAddress,
  workbookWithSettings,
  rangeAddress,
) {
  const cell = stripCellDecorations(sheetOrThrow.cells[cellAddress]);
  if (cell.raw !== null) return cell;
  const calculatedValue = workbookWithSettings(sheetOrThrow.id, cellAddress);
  if (calculatedValue === null) return cell;
  const spillAnchor = workbookWithSettings.getSpillAnchor(
    sheetOrThrow.id,
    cellAddress,
  );
  if (spillAnchor && spillAnchor !== cellAddress) {
    const anchorPosition = _t(spillAnchor);
    (anchorPosition.row < rangeAddress.firstRow ||
      anchorPosition.row > rangeAddress.lastRow ||
      anchorPosition.col < rangeAddress.firstCol ||
      anchorPosition.col > rangeAddress.lastCol) &&
      ((cell.raw = calculatedValue),
      typeof calculatedValue == "string" &&
        calculatedValue.startsWith("=") &&
        (cell.literal = true));
  }
  return cell;
}
function assertFillTargetUnmerged(workbook, sheetId, targetAddresses) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  for (const mergeRange of sheetOrThrow.merges ?? []) {
    const rangeAddress = parseRangeAddress(mergeRange);
    targetAddresses.some((targetAddress) => {
      const targetPosition = _t(targetAddress);
      return (
        targetPosition.row >= rangeAddress.firstRow &&
        targetPosition.row <= rangeAddress.lastRow &&
        targetPosition.col >= rangeAddress.firstCol &&
        targetPosition.col <= rangeAddress.lastCol &&
        (targetPosition.row !== rangeAddress.firstRow ||
          targetPosition.col !== rangeAddress.firstCol)
      );
    }) &&
      throwWorkbookError("Unmerge covered target cells before filling them.");
  }
}
function buildFillPreview(
  workbook,
  sheetId,
  targetRange,
  directionInput,
  options = {},
) {
  const direction = validateFillDirection(directionInput),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(targetRange);
  (rangeAddress.lastRow >= sheetOrThrow.rowCount ||
    rangeAddress.lastCol >= sheetOrThrow.colCount) &&
    throwWorkbookError("The fill range exceeds the worksheet.");
  const isVertical = direction === "down" || direction === "up",
    rangeRowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1,
    rangeColCount = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    destinationCount = isVertical
      ? (rangeRowCount - 1) * rangeColCount
      : (rangeColCount - 1) * rangeRowCount;
  destinationCount > workbookLimits.maxPatches &&
    throwWorkbookError("Fill at most 500 destination cells.");
  const sourceRow =
      direction === "up" ? rangeAddress.lastRow : rangeAddress.firstRow,
    sourceCol =
      direction === "left" ? rangeAddress.lastCol : rangeAddress.firstCol,
    patches = [],
    targetAddresses = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    sourceRangeAddress = isVertical
      ? { ...rangeAddress, firstRow: sourceRow, lastRow: sourceRow }
      : { ...rangeAddress, firstCol: sourceCol, lastCol: sourceCol };
  if (destinationCount)
    for (
      let firstRow = rangeAddress.firstRow;
      firstRow <= rangeAddress.lastRow;
      firstRow++
    )
      for (
        let firstCol = rangeAddress.firstCol;
        firstCol <= rangeAddress.lastCol;
        firstCol++
      ) {
        if (isVertical ? firstRow === sourceRow : firstCol === sourceCol)
          continue;
        const sourceRowIndex = isVertical ? sourceRow : firstRow,
          sourceColIndex = isVertical ? firstCol : sourceCol,
          cellAddress = formatCellAddress(firstRow, firstCol);
        targetAddresses.push(cellAddress);
        const fillCellPatches = buildFillCellPatches(
          sheetOrThrow,
          formatCellAddress(sourceRowIndex, sourceColIndex),
          workbookWithSettings,
          sourceRangeAddress,
        );
        typeof fillCellPatches.raw == "string" &&
          fillCellPatches.raw.startsWith("=") &&
          !fillCellPatches.literal &&
          (fillCellPatches.raw = shiftFormulaReferences(
            fillCellPatches.raw,
            firstRow - sourceRowIndex,
            firstCol - sourceColIndex,
          ));
        const cellRecordInput = normalizeCellRecordInput(
          mergeCellDecorations(
            fillCellPatches,
            sheetOrThrow.cells[cellAddress],
          ),
        );
        cellContentsDiffer(sheetOrThrow.cells[cellAddress], cellRecordInput) &&
          patches.push({
            sheetId: sheetId,
            address: cellAddress,
            cell: cellRecordInput,
          });
      }
  return (
    assertFillTargetUnmerged(workbook, sheetId, targetAddresses),
    buildCellChangePreview(
      workbook,
      patches,
      {
        targetSheetId: sheetId,
        targetRange: formatRangeBoundsAddress(targetRange),
        targetCount: destinationCount,
        generatedCount: destinationCount,
        direction: direction,
        summary: `Fill ${direction} within ${formatRangeBoundsAddress(targetRange)}.`,
      },
      options.readOnly,
    )
  );
}

function excelSerialToDate(serial) {
  return new Date(excelEpochUtc + (serial > 60 ? serial - 1 : serial) * s$);
}

function stepSeriesDateSerial(seedSerial, step, dateUnit, monthPolicy) {
  if (
    (Number.isSafeInteger(step) ||
      throwWorkbookError("The date step exceeds safe integer limits."),
    !step)
  )
    return seedSerial;
  if (dateUnit === "day") {
    const shiftedSerial = seedSerial + step;
    return (
      (shiftedSerial < 0 || shiftedSerial > maxExcelDateSerial) &&
        throwDateSeriesRangeError(),
      parseSeriesSeedDate(shiftedSerial)
    );
  }
  seedSerial === 60 &&
    throwWorkbookError(
      "Calendar and weekday series require a real calendar seed; use Days for the fictitious 1900-02-29 serial.",
    );
  const seedDate = excelSerialToDate(seedSerial);
  if (dateUnit === "weekday") {
    const stepSign = Math.sign(step);
    let remainingSteps = Math.abs(step);
    if (seedDate.getUTCDay() === 0 || seedDate.getUTCDay() === 6) {
      do seedDate.setUTCDate(seedDate.getUTCDate() + stepSign);
      while ([0, 6].includes(seedDate.getUTCDay()));
      remainingSteps--;
    }
    for (
      seedDate.setUTCDate(
        seedDate.getUTCDate() + stepSign * Math.floor(remainingSteps / 5) * 7,
      ),
        remainingSteps %= 5;
      remainingSteps;
    )
      (seedDate.setUTCDate(seedDate.getUTCDate() + stepSign),
        [0, 6].includes(seedDate.getUTCDay()) || remainingSteps--);
    return dateToExcelSerial(seedDate);
  }
  const totalMonths =
      seedDate.getUTCMonth() + step * (dateUnit === "year" ? 12 : 1),
    targetYear = seedDate.getUTCFullYear() + Math.floor(totalMonths / 12),
    targetMonthIndex = ((totalMonths % 12) + 12) % 12;
  (!Number.isSafeInteger(totalMonths) ||
    targetYear < 1899 ||
    targetYear > 9999) &&
    throwDateSeriesRangeError();
  const daysInMonth = utcDateFromParts(
      targetYear,
      targetMonthIndex + 1,
      0,
    ).getUTCDate(),
    targetDay =
      monthPolicy === "endOfMonth"
        ? daysInMonth
        : Math.min(seedDate.getUTCDate(), daysInMonth);
  return dateToExcelSerial(
    utcDateFromParts(targetYear, targetMonthIndex, targetDay),
  );
}
function Nae(seedSerial, stepAmount, dateUnit, monthPolicy, dateSystem) {
  const resultSerial =
    stepSeriesDateSerial(
      seedSerial + dateSystemSerialOffset(dateSystem),
      stepAmount,
      dateUnit,
      monthPolicy,
    ) - dateSystemSerialOffset(dateSystem);
  return (
    (resultSerial < 0 || resultSerial > dateSystemMaxSerial(dateSystem)) &&
      throwDateSeriesRangeError(),
    resultSerial
  );
}
function buildSeriesPreview(workbook, sheetId, range, seriesOptions) {
  const direction = validateFillDirection(seriesOptions.direction),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(range),
    isVertical = direction === "down" || direction === "up";
  (["number", "date"].includes(seriesOptions.kind) ||
    throwWorkbookError("Choose a numeric or date series."),
    ((isVertical && rangeAddress.firstCol !== rangeAddress.lastCol) ||
      (!isVertical && rangeAddress.firstRow !== rangeAddress.lastRow)) &&
      throwWorkbookError(
        "Select one column for vertical series or one row for horizontal series.",
      ),
    (rangeAddress.lastRow >= sheetOrThrow.rowCount ||
      rangeAddress.lastCol >= sheetOrThrow.colCount) &&
      throwWorkbookError("The series range exceeds the worksheet."));
  const targetCount = isVertical
    ? rangeAddress.lastRow - rangeAddress.firstRow + 1
    : rangeAddress.lastCol - rangeAddress.firstCol + 1;
  targetCount > workbookLimits.maxPatches &&
    throwWorkbookError("A series may contain at most 500 target cells.");
  const seedRow =
      direction === "up" ? rangeAddress.lastRow : rangeAddress.firstRow,
    seedCol =
      direction === "left" ? rangeAddress.lastCol : rangeAddress.firstCol,
    cellAddress = formatCellAddress(seedRow, seedCol),
    seedCell = stripCellDecorations(sheetOrThrow.cells[cellAddress]),
    step = seriesOptions.step ?? 1;
  (typeof step != "number" ||
    !Number.isFinite(step) ||
    step === 0 ||
    (seriesOptions.kind === "date" && !Number.isSafeInteger(step))) &&
    throwWorkbookError(
      "Series step must be a nonzero finite number; date steps must be whole numbers.",
    );
  const dateUnit = seriesOptions.dateUnit ?? "day",
    monthPolicy = seriesOptions.monthPolicy ?? "preserveDay";
  (!["day", "weekday", "month", "year"].includes(dateUnit) ||
    !["preserveDay", "endOfMonth"].includes(monthPolicy)) &&
    throwWorkbookError("Choose a supported date unit and month policy.");
  const rawStart =
    seriesOptions.start ??
    calculateWorkbookWithSettings(workbook)(sheetId, cellAddress);
  typeof rawStart != "number" &&
    typeof rawStart != "string" &&
    throwWorkbookError(
      "The directional seed must contain a number or date, or supply a start value.",
    );
  const startValue =
      seriesOptions.kind === "date"
        ? parseSeriesSeedDate(rawStart, workbook.datePolicy?.dateSystem)
        : typeof rawStart == "number" && Number.isFinite(rawStart)
          ? rawStart
          : throwWorkbookError(
              "Numeric series require a finite numeric start value.",
            ),
    stopValue =
      seriesOptions.stop === undefined
        ? undefined
        : seriesOptions.kind === "date"
          ? parseSeriesSeedDate(
              seriesOptions.stop,
              workbook.datePolicy?.dateSystem,
            )
          : typeof seriesOptions.stop == "number" &&
              Number.isFinite(seriesOptions.stop)
            ? seriesOptions.stop
            : throwWorkbookError(
                "Numeric series require a finite numeric stop value.",
              );
  stopValue !== undefined &&
    ((step > 0 && stopValue < startValue) ||
      (step < 0 && stopValue > startValue)) &&
    throwWorkbookError("The stop value is opposite the series step direction.");
  const patches = [],
    targetAddresses = [];
  let generatedCount = 0;
  for (let valueIndex = 0; valueIndex < targetCount; valueIndex++) {
    let raw;
    try {
      raw =
        seriesOptions.kind === "date"
          ? Nae(
              startValue,
              valueIndex * step,
              dateUnit,
              monthPolicy,
              workbook.datePolicy?.dateSystem ?? "1900",
            )
          : startValue + valueIndex * step;
    } catch (error) {
      if (
        stopValue !== undefined &&
        valueIndex > 0 &&
        error instanceof DateSeriesRangeError
      )
        break;
      throw error;
    }
    if (
      stopValue !== undefined &&
      (step > 0 ? raw > stopValue : raw < stopValue)
    )
      break;
    Number.isFinite(raw) ||
      throwWorkbookError("The numeric series would overflow.");
    const cellAddress = formatCellAddress(
      seedRow + (isVertical ? valueIndex * (direction === "up" ? -1 : 1) : 0),
      seedCol + (isVertical ? 0 : valueIndex * (direction === "left" ? -1 : 1)),
    );
    (targetAddresses.push(cellAddress), generatedCount++);
    const cellRecordInput = normalizeCellRecordInput(
      mergeCellDecorations(
        {
          ...seedCell,
          raw: raw,
          literal: undefined,
          ...(seriesOptions.kind === "date"
            ? { format: "date", numFmt: "yyyy-mm-dd", numberFormat: undefined }
            : {}),
        },
        sheetOrThrow.cells[cellAddress],
      ),
    );
    cellContentsDiffer(sheetOrThrow.cells[cellAddress], cellRecordInput) &&
      patches.push({
        sheetId: sheetId,
        address: cellAddress,
        cell: cellRecordInput,
      });
  }
  return (
    assertFillTargetUnmerged(workbook, sheetId, targetAddresses),
    buildCellChangePreview(
      workbook,
      patches,
      {
        targetSheetId: sheetId,
        targetRange: formatRangeBoundsAddress(range),
        targetCount: targetCount,
        generatedCount: generatedCount,
        direction: direction,
        summary: `Fill ${seriesOptions.kind === "date" ? `${dateUnit} date` : "numeric"} series ${direction} in ${formatRangeBoundsAddress(range)} (${generatedCount} values).`,
      },
      seriesOptions.readOnly,
    )
  );
}

function buildFillHandlePreview(
  workbook,
  sheetId,
  sourceRange,
  targetRange,
  options = {},
) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(sourceRange),
    targetAddress = parseRangeAddress(targetRange),
    fillMode = options.mode ?? "copy";
  (["copy", "series"].includes(fillMode) ||
    throwWorkbookError("Choose Copy or Series for the fill handle."),
    (rangeAddress.firstRow < targetAddress.firstRow ||
      rangeAddress.lastRow > targetAddress.lastRow ||
      rangeAddress.firstCol < targetAddress.firstCol ||
      rangeAddress.lastCol > targetAddress.lastCol) &&
      throwWorkbookError(
        "The fill-handle target must contain the complete source range.",
      ),
    (targetAddress.lastRow >= sheetOrThrow.rowCount ||
      targetAddress.lastCol >= sheetOrThrow.colCount) &&
      throwWorkbookError("The fill-handle target exceeds the worksheet."));
  const sourceRowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1,
    sourceColCount = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    isVerticalFill =
      rangeAddress.firstCol === targetAddress.firstCol &&
      rangeAddress.lastCol === targetAddress.lastCol,
    isHorizontalFill =
      rangeAddress.firstRow === targetAddress.firstRow &&
      rangeAddress.lastRow === targetAddress.lastRow;
  !isVerticalFill &&
    !isHorizontalFill &&
    throwWorkbookError(
      "Extend the fill handle along one axis while keeping the other source dimension unchanged.",
    );
  const extensionCellCount =
    (targetAddress.lastRow - targetAddress.firstRow + 1) *
      (targetAddress.lastCol - targetAddress.firstCol + 1) -
    sourceRowCount * sourceColCount;
  extensionCellCount > workbookLimits.maxPatches &&
    throwWorkbookError("The fill handle may extend into at most 500 cells.");
  const direction = isVerticalFill
      ? targetAddress.firstRow < rangeAddress.firstRow
        ? "up"
        : "down"
      : targetAddress.firstCol < rangeAddress.firstCol
        ? "left"
        : "right",
    cellChanges = [],
    targetCellAddresses = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  let seriesStartValue = 0,
    seriesStep = 0;
  if (fillMode === "series") {
    (isVerticalFill && sourceRowCount === 2 && sourceColCount === 1) ||
      (isHorizontalFill && sourceRowCount === 1 && sourceColCount === 2) ||
      throwWorkbookError(
        "Series handle fill requires exactly two seed cells in the extension's row or column.",
      );
    const firstSeedValue = workbookWithSettings(
        sheetId,
        formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol),
      ),
      lastSeedValue = workbookWithSettings(
        sheetId,
        formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol),
      );
    ((typeof firstSeedValue != "number" ||
      !Number.isFinite(firstSeedValue) ||
      typeof lastSeedValue != "number" ||
      !Number.isFinite(lastSeedValue) ||
      !Number.isFinite(lastSeedValue - firstSeedValue)) &&
      throwWorkbookError(
        "Series handle seeds must evaluate to finite numbers.",
      ),
      (seriesStartValue = firstSeedValue),
      (seriesStep = lastSeedValue - firstSeedValue));
  }
  const wrapIndex = (index, count) => ((index % count) + count) % count;
  function* extensionCells() {
    if (isVerticalFill)
      for (const [rowSegmentStart, rowSegmentEnd] of [
        [targetAddress.firstRow, rangeAddress.firstRow - 1],
        [rangeAddress.lastRow + 1, targetAddress.lastRow],
      ])
        for (let row = rowSegmentStart; row <= rowSegmentEnd; row++)
          for (
            let firstCol = rangeAddress.firstCol;
            firstCol <= rangeAddress.lastCol;
            firstCol++
          )
            yield { row: row, col: firstCol };
    else
      for (
        let firstRow = rangeAddress.firstRow;
        firstRow <= rangeAddress.lastRow;
        firstRow++
      )
        for (const [colSegmentStart, colSegmentEnd] of [
          [targetAddress.firstCol, rangeAddress.firstCol - 1],
          [rangeAddress.lastCol + 1, targetAddress.lastCol],
        ])
          for (let col = colSegmentStart; col <= colSegmentEnd; col++)
            yield { row: firstRow, col: col };
  }
  if (extensionCellCount)
    for (const { row, col } of extensionCells()) {
      const sourceRow =
          fillMode === "series"
            ? rangeAddress.firstRow
            : rangeAddress.firstRow +
              wrapIndex(row - rangeAddress.firstRow, sourceRowCount),
        sourceCol =
          fillMode === "series"
            ? rangeAddress.firstCol
            : rangeAddress.firstCol +
              wrapIndex(col - rangeAddress.firstCol, sourceColCount),
        fillCellPatches = buildFillCellPatches(
          sheetOrThrow,
          formatCellAddress(sourceRow, sourceCol),
          workbookWithSettings,
          rangeAddress,
        ),
        cellAddress = formatCellAddress(row, col);
      if (fillMode === "series") {
        const seriesValue =
          seriesStartValue +
          (isVerticalFill
            ? row - rangeAddress.firstRow
            : col - rangeAddress.firstCol) *
            seriesStep;
        (Number.isFinite(seriesValue) ||
          throwWorkbookError("The handle series would overflow."),
          (fillCellPatches.raw = seriesValue),
          delete fillCellPatches.literal);
      } else
        typeof fillCellPatches.raw == "string" &&
          fillCellPatches.raw.startsWith("=") &&
          !fillCellPatches.literal &&
          (fillCellPatches.raw = shiftFormulaReferences(
            fillCellPatches.raw,
            row - sourceRow,
            col - sourceCol,
          ));
      const cellRecordInput = normalizeCellRecordInput(
        mergeCellDecorations(fillCellPatches, sheetOrThrow.cells[cellAddress]),
      );
      (targetCellAddresses.push(cellAddress),
        cellContentsDiffer(sheetOrThrow.cells[cellAddress], cellRecordInput) &&
          cellChanges.push({
            sheetId: sheetId,
            address: cellAddress,
            cell: cellRecordInput,
          }));
    }
  return (
    assertFillTargetUnmerged(workbook, sheetId, targetCellAddresses),
    buildCellChangePreview(
      workbook,
      cellChanges,
      {
        targetSheetId: sheetId,
        targetRange: formatRangeBoundsAddress(targetRange),
        targetCount: extensionCellCount,
        generatedCount: extensionCellCount,
        direction: direction,
        summary: `Fill handle ${fillMode} from ${formatRangeBoundsAddress(sourceRange)} into ${formatRangeBoundsAddress(targetRange)} (${extensionCellCount} extension cells).`,
      },
      options.readOnly,
    )
  );
}
function parseDateCellEntry(entryText, existingCell, dateSystem = "1900") {
  if (typeof entryText != "string") throw new Error("Cell entry must be text.");
  const cell = existingCell ? structuredClone(existingCell) : { raw: null };
  if (((cell.literal = entryText.startsWith("'")), cell.literal))
    cell.raw = entryText.slice(1);
  else if (/^\d{4}-\d{2}-\d{2}$/.test(entryText)) {
    try {
      cell.raw = parseSeriesSeedDate(entryText, dateSystem);
    } catch {
      throw new Error(
        `Enter a valid YYYY-MM-DD date between ${dateSystem === "1904" ? "1904-01-01" : "1899-12-31"} and 9999-12-31, or begin with an apostrophe to keep the entry as text.${dateSystem === "1900" ? " Excel's compatibility date 1900-02-29 represents serial 60." : ""}`,
      );
    }
    ((cell.format = "date"),
      (cell.numFmt = "yyyy-mm-dd"),
      (cell.numberFormat = {
        category: "date",
        locale: "en-US",
        pattern: "yyyy-mm-dd",
      }));
  } else cell.raw = parseRawCellEntry(entryText);
  return normalizeCellRecordInput(cell);
}
function parseRawCellEntry(entryText) {
  return entryText === ""
    ? null
    : /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(entryText) &&
        entryText.replace(/[-.]/g, "").length <= 15
      ? Number(entryText)
      : /^(TRUE|FALSE)$/i.test(entryText)
        ? entryText.toUpperCase() === "TRUE"
        : entryText;
}
const supportedWorkbookLocales = ["en-US", "en-GB", "de-DE", "fr-FR"];
function resolveWorkbookLocale(locale) {
  if (locale === undefined) return "en-US";
  if (!supportedWorkbookLocales.includes(locale))
    throw new Error("Choose a supported workbook interface locale.");
  return locale;
}
function formatNumberForLocale(numberValue, entryLocale, useGrouping = true) {
  return new Intl.NumberFormat(entryLocale, {
    useGrouping: useGrouping,
    maximumSignificantDigits: 15,
  }).format(numberValue);
}
function parseLocalizedCellEntry(entryText, cell, dateSystem = "1900", locale) {
  if (
    (resolveWorkbookLocale(locale),
    typeof cell?.raw == "number" &&
      entryText === localizeDecimalPoint(cell.raw, locale))
  )
    return normalizeCellRecordInput(structuredClone(cell));
  if (
    !locale ||
    /^[=']/.test(entryText) ||
    /^\d{4}-\d{2}-\d{2}$/.test(entryText)
  )
    return parseDateCellEntry(entryText, cell, dateSystem);
  const decimalSeparator = locale === "de-DE" || locale === "fr-FR" ? "," : ".",
    groupSeparator =
      locale === "de-DE" ? "." : locale === "fr-FR" ? "[ \\u00a0\\u202f]" : ",",
    escapeRegExp = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
    thousandsGroupPattern = `(?:[1-9]\\d{0,2}(?:${locale === "fr-FR" ? groupSeparator : escapeRegExp(groupSeparator)}\\d{3})+)`;
  if (
    new RegExp(
      `^-?(?:0|[1-9]\\d*|${thousandsGroupPattern})(?:${escapeRegExp(decimalSeparator)}\\d+)?$`,
    ).test(entryText)
  ) {
    const normalizedNumberText = entryText
      .replace(
        locale === "fr-FR"
          ? /[ \u00a0\u202f]/g
          : new RegExp(escapeRegExp(groupSeparator), "g"),
        "",
      )
      .replace(decimalSeparator, ".");
    return parseDateCellEntry(normalizedNumberText, cell, dateSystem);
  }
  return locale === "de-DE" && /^(WAHR|FALSCH)$/i.test(entryText)
    ? parseDateCellEntry(
        entryText.toUpperCase() === "WAHR" ? "TRUE" : "FALSE",
        cell,
        dateSystem,
      )
    : locale === "fr-FR" && /^(VRAI|FAUX)$/i.test(entryText)
      ? parseDateCellEntry(
          entryText.toUpperCase() === "VRAI" ? "TRUE" : "FALSE",
          cell,
          dateSystem,
        )
      : /^-?[\d., \u00a0\u202f]+$/.test(entryText) && /[.,]/.test(entryText)
        ? parseDateCellEntry("'" + entryText, cell, dateSystem)
        : parseDateCellEntry(entryText, cell, dateSystem);
}
function formatCellEntryForEdit(cell, entryLocale) {
  return typeof cell?.raw == "number" && entryLocale
    ? localizeDecimalPoint(cell.raw, entryLocale)
    : `${cell?.literal && typeof cell.raw == "string" ? "'" : ""}${String(cell?.raw ?? "")}`;
}
function localizeDecimalPoint(numberValue, entryLocale) {
  return Number.isFinite(numberValue)
    ? entryLocale === "de-DE" || entryLocale === "fr-FR"
      ? String(numberValue).replace(".", ",")
      : String(numberValue)
    : "";
}
function validateLocalizedNumberEntry(entryText, fieldLabel, locale) {
  let numberText = entryText.trim();
  {
    resolveWorkbookLocale(locale);
    const decimalSeparator =
        locale === "de-DE" || locale === "fr-FR" ? "," : ".",
      groupSeparatorPattern =
        locale === "de-DE"
          ? "\\."
          : locale === "fr-FR"
            ? "[ \\u00a0\\u202f]"
            : ",";
    if (
      !new RegExp(
        `^[+-]?(?:(?:\\d+|[1-9]\\d{0,2}(?:${groupSeparatorPattern}\\d{3})+)(?:${decimalSeparator === "." ? "\\." : ","}\\d*)?|${decimalSeparator === "." ? "\\." : ","}\\d+)(?:[eE][+-]?\\d+)?$`,
      ).test(numberText)
    )
      throw new Error(
        `${fieldLabel} must be a finite number in the selected numeric locale.`,
      );
    numberText = numberText
      .replace(new RegExp(groupSeparatorPattern, "g"), "")
      .replace(decimalSeparator, ".");
  }
  if (!numberText || !Number.isFinite(Number(numberText)))
    throw new Error(`${fieldLabel} must be a finite number.`);
  return Number(numberText);
}
function formatLocalizedNumberDisplay(
  numberValue,
  cell,
  dateSystem = "1900",
  entryLocale,
) {
  return !entryLocale ||
    typeof numberValue != "number" ||
    cell?.numberFormat ||
    cell?.numFmt ||
    cell?.format === "date"
    ? formatNumberWithPattern(numberValue, cell, dateSystem)
    : Number.isFinite(numberValue)
      ? cell?.format === "currency"
        ? new Intl.NumberFormat(entryLocale, {
            style: "currency",
            currency: "USD",
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          }).format(numberValue)
        : cell?.format === "percent"
          ? new Intl.NumberFormat(entryLocale, {
              style: "percent",
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            }).format(numberValue)
          : cell?.format === "number"
            ? new Intl.NumberFormat(entryLocale, {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2,
              }).format(numberValue)
            : formatNumberForLocale(numberValue, entryLocale, false)
      : "#NUM!";
}
const germanWorkbookMessages = {
    "Enter a valid number using the selected locale.":
      "Geben Sie eine gültige Zahl im gewählten Zahlenformat ein.",
    "Enter a number from {min} to {max}.":
      "Geben Sie eine Zahl zwischen {min} und {max} ein.",
    "Calculation Options": "Berechnungsoptionen",
    "Close Calculation Options": "Berechnungsoptionen schließen",
    "Calculation mode": "Berechnungsmodus",
    "Maximum iterations": "Maximale Iterationen",
    "Maximum change": "Maximale Änderung",
    "Time budget (ms)": "Zeitlimit (ms)",
    "Memory budget (MiB)": "Speicherlimit (MiB)",
    "Work budget": "Arbeitslimit",
    "Apply calculation settings": "Berechnungseinstellungen anwenden",
    "Calculate Now": "Jetzt berechnen",
    "Automatic calculation": "Automatische Berechnung",
    "Manual calculation": "Manuelle Berechnung",
    "Enable iterative calculation": "Iterative Berechnung aktivieren",
    "Validation type": "Prüfungstyp",
    "Validation operator": "Prüfungsoperator",
    "Validation range": "Prüfungsbereich",
    "Validation values": "Prüfungswerte",
    "Validation maximum": "Maximalwert der Prüfung",
    "Validation formula": "Prüfungsformel",
    "Close data validation": "Datenüberprüfung schließen",
    "Validation input title": "Titel des Eingabehinweises",
    "Validation input message": "Eingabehinweis",
    "Validation error style": "Art der Fehlermeldung",
    "Validation error title": "Titel der Fehlermeldung",
    "Validation message": "Prüfungsmeldung",
    "Minimum / comparison value": "Minimum / Vergleichswert",
    Maximum: "Maximum",
    Minimum: "Minimum",
    Allow: "Zulassen",
    Condition: "Bedingung",
    "Whole number": "Ganze Zahl",
    Decimal: "Dezimalzahl",
    "Text length": "Textlänge",
    "Custom formula": "Benutzerdefinierte Formel",
    List: "Liste",
    Time: "Uhrzeit",
    "Input message": "Eingabehinweis",
    "Error alert": "Fehlermeldung",
    "Allow blank cells": "Leere Zellen zulassen",
    "Show in-cell dropdown": "Dropdown in der Zelle anzeigen",
    "Save rule": "Regel speichern",
    "Step value": "Schrittweite",
    "Structure count": "Anzahl der Strukturelemente",
    "Rows to append": "Anzufügende Zeilen",
    "Decimal places": "Dezimalstellen",
    "Pages wide": "Seiten breit",
    "Pages tall": "Seiten hoch",
    "Print scale": "Druckskalierung",
    "Object width": "Objektbreite",
    "Object height": "Objekthöhe",
    "Histogram bins": "Histogrammklassen",
    "Fuzzy similarity threshold": "Schwellenwert für unscharfe Ähnlichkeit",
    "Recipe fuzzy threshold": "Ähnlichkeitsschwelle des Ablaufs",
    "Materiality threshold": "Wesentlichkeitsgrenze",
    "Sample size": "Stichprobenumfang",
    Threshold: "Schwellenwert",
    "Indent level": "Einzugsebene",
    "Default (14 px)": "Standard (14 px)",
    "Open-ended": "Ohne Obergrenze",
    Value: "Wert",
    Constant: "Konstante",
    "Separate numeric assumptions with semicolons or new lines. Use a comma for decimals.":
      "Trennen Sie Zahlenwerte durch Semikolons oder Zeilenumbrüche. Verwenden Sie ein Komma als Dezimaltrennzeichen.",
    "Separate numeric assumptions with commas, semicolons or new lines.":
      "Trennen Sie Zahlenwerte durch Kommas, Semikolons oder Zeilenumbrüche.",
    "Enter finite numbers separated by semicolons or new lines; use a comma for decimals. Empty assumptions are not zero.":
      "Geben Sie endliche Zahlen getrennt durch Semikolons oder Zeilenumbrüche ein; verwenden Sie ein Dezimalkomma. Leere Annahmen sind keine Nullwerte.",
    "Workbook chat": "Arbeitsmappen-Chat",
    "New chat": "Neuer Chat",
    "Local demo": "Lokale Demo",
    "Shared workbook context": "Freigegebener Arbeitsmappenkontext",
    "{count} selected cells · shared when you send":
      "{count} ausgewählte Zellen · werden beim Senden geteilt",
    Conversation: "Unterhaltung",
    "How can I help?": "Wie kann ich helfen?",
    "Ask about your selection, attach a file, or work through a formula together.":
      "Stellen Sie eine Frage zur Auswahl, hängen Sie eine Datei an oder besprechen Sie eine Formel.",
    "Explain this selection": "Diese Auswahl erklären",
    "Highlight values above 25000": "Werte über 25000 hervorheben",
    "Add a total formula": "Summenformel hinzufügen",
    "Your message": "Ihre Nachricht",
    "Assistant message": "Nachricht des Assistenten",
    You: "Sie",
    Assistant: "Assistent",
    "File ready": "Datei bereit",
    "Preview of {name}": "Vorschau von {name}",
    "Remove attachment {name}": "Anhang {name} entfernen",
    "Remove attachment": "Anhang entfernen",
    "The original file will accompany your message. Preview is unavailable for this format.":
      "Die Originaldatei wird mit Ihrer Nachricht gesendet. Für dieses Format ist keine Vorschau verfügbar.",
    "This message belongs to an earlier selection or workbook revision and is excluded from the current context.":
      "Diese Nachricht gehört zu einer früheren Auswahl oder Arbeitsmappenversion und wird im aktuellen Kontext nicht verwendet.",
    "Copy assistant response": "Antwort des Assistenten kopieren",
    "Copy response": "Antwort kopieren",
    "Regenerate response": "Antwort erneut erstellen",
    "Retry message": "Nachricht erneut senden",
    "Preparing a response…": "Antwort wird vorbereitet…",
    "Preparing attachments…": "Anhänge werden vorbereitet…",
    "Scroll to latest message": "Zur neuesten Nachricht scrollen",
    "Attached files": "Angehängte Dateien",
    "Upload chat attachments": "Chat-Anhänge hochladen",
    "Attach files": "Dateien anhängen",
    "Attach files · up to 4 files, 5 MB each":
      "Dateien anhängen · bis zu 4 Dateien mit je 5 MB",
    "File support is not enabled by this host adapter":
      "Die einbettende Anwendung hat die Dateifunktion nicht aktiviert",
    "Message the workbook assistant…":
      "Nachricht an den Arbeitsmappenassistenten…",
    "Drop files here": "Dateien hier ablegen",
    "Enter to send": "Eingabetaste zum Senden",
    "Cancel request": "Anfrage abbrechen",
    "Stop response": "Antwort stoppen",
    "Send selection": "Auswahl senden",
    "Send message and selected context":
      "Nachricht und ausgewählten Kontext senden",
    "Local demo · files stay in this browser":
      "Lokale Demo · Dateien bleiben in diesem Browser",
    "Selection, same-context chat and files go to your host":
      "Auswahl, zugehöriger Chat und Dateien werden an Ihre Anwendung gesendet",
    "Selection only · file and chat-history support unavailable":
      "Nur die Auswahl · Dateien und Chatverlauf werden nicht unterstützt",
    "Review proposed changes before applying. Shift+Enter for a new line.":
      "Prüfen Sie vorgeschlagene Änderungen vor dem Anwenden. Umschalt+Eingabe für eine neue Zeile.",
    "Message copied.": "Nachricht kopiert.",
    "Copy is unavailable. Select the message text and use your browser’s Copy command.":
      "Kopieren ist nicht verfügbar. Markieren Sie den Nachrichtentext und verwenden Sie den Kopierbefehl Ihres Browsers.",
    "This provider does not support attachments.":
      "Dieser Anbieter unterstützt keine Anhänge.",
  },
  germanWorkbookStatusMessages = {
    "Workbook health": "Arbeitsmappenstatus",
    "Close workbook health": "Arbeitsmappenstatus schließen",
    "Saving and recovery": "Speichern und Wiederherstellen",
    Calculation: "Berechnung",
    "File processing": "Dateiverarbeitung",
    "AI provider": "KI-Anbieter",
    "Application callbacks": "Anwendungsrückmeldungen",
    Service: "Dienst",
    State: "Status",
    Completed: "Abgeschlossen",
    Failures: "Fehler",
    "Longest operation": "Längster Vorgang",
    "Needs attention": "Handlungsbedarf",
    "Not used": "Nicht verwendet",
    "Retained events": "Gespeicherte Ereignisse",
    "Discarded older events": "Verworfene ältere Ereignisse",
    "Diagnostic delivery failures": "Fehler bei der Diagnoseübermittlung",
    "Download diagnostics": "Diagnose herunterladen",
    "Clear diagnostic history": "Diagnoseverlauf löschen",
    "Waiting for save acknowledgement": "Warten auf Speicherbestätigung",
    "Diagnostics contain operation categories, outcomes and timings. Workbook contents, names, prompts and error details are excluded.":
      "Diagnosen enthalten Vorgangsarten, Ergebnisse und Laufzeiten. Arbeitsmappeninhalte, Namen, Eingaben und Fehlerdetails werden nicht aufgenommen.",
    "Save failures remain visible in the editor. Export your local changes before reloading a saved copy. A pending save must finish before another retry.":
      "Speicherfehler bleiben im Editor sichtbar. Exportieren Sie lokale Änderungen, bevor Sie eine gespeicherte Kopie neu laden. Ein laufender Speichervorgang muss vor einem erneuten Versuch beendet werden.",
    "Undo (Ctrl+Z)": "Rückgängig (Ctrl+Z)",
    "Redo (Ctrl+Y)": "Wiederholen (Ctrl+Y)",
    "Goal Seek": "Zielwertsuche",
    "Scenario Manager": "Szenario-Manager",
    "The current edit keeps its original numeric locale until you apply or cancel it.":
      "Die aktuelle Eingabe behält ihr ursprüngliches Zahlenformat, bis Sie sie übernehmen oder abbrechen.",
    File: "Datei",
    Home: "Start",
    Insert: "Einfügen",
    "Page Layout": "Seitenlayout",
    Formulas: "Formeln",
    Data: "Daten",
    Review: "Überprüfen",
    View: "Ansicht",
    Open: "Öffnen",
    Save: "Speichern",
    Cancel: "Abbrechen",
    Close: "Schließen",
    Apply: "Anwenden",
    Preview: "Vorschau",
    Refresh: "Aktualisieren",
    Delete: "Löschen",
    Name: "Name",
    Reset: "Zurücksetzen",
    Undo: "Rückgängig",
    Redo: "Wiederholen",
    Ready: "Bereit",
    "Working…": "Wird verarbeitet…",
    "Read only": "Schreibgeschützt",
    Average: "Mittelwert",
    Sum: "Summe",
    "Count:": "Anzahl:",
    "Numbers:": "Zahlen:",
    "Workbook ribbon": "Menüband der Arbeitsmappe",
    "Ribbon tabs": "Registerkarten",
    "Formula bar": "Bearbeitungsleiste",
    "Selected range": "Ausgewählter Bereich",
    "Editable worksheet grid": "Bearbeitbares Tabellenraster",
    "Spreadsheet workbook": "Arbeitsmappe",
    "Selection summary": "Zusammenfassung der Auswahl",
    "Skip to worksheet": "Zum Tabellenblatt",
    "Skip to formula bar": "Zur Bearbeitungsleiste",
    "Cell position": "Zellposition",
    "Worksheet navigation": "Tabellennavigation",
    "{sheet}, cell {address}, row {row}, column {column}. {value}{formula}{access}":
      "{sheet}, Zelle {address}, Zeile {row}, Spalte {column}. {value}{formula}{access}",
    Blank: "Leer",
    " Formula: {formula}.": " Formel: {formula}.",
    " Read only.": " Schreibgeschützt.",
    " Error: {error}.": " Fehler: {error}.",
    "Arrow keys move between cells. F2 edits. Escape returns to the ribbon. Use the selected range field to jump to a cell. Formula names, references and decimal points inside formulas use invariant syntax; enter dates as YYYY-MM-DD.":
      "Pfeiltasten bewegen die Auswahl. F2 bearbeitet die Zelle. Escape führt zum Menüband. Das Bereichsfeld springt zu einer Zelle. Formelnamen, Bezüge und Dezimalpunkte in Formeln bleiben unverändert; Datumswerte als JJJJ-MM-TT eingeben.",
    "Cancel cell edit": "Zellbearbeitung abbrechen",
    "Apply cell value": "Zellwert übernehmen",
    "Add worksheet": "Tabellenblatt hinzufügen",
    "Manage worksheets": "Tabellenblätter verwalten",
    "Select all cells": "Alle Zellen auswählen",
    Zoom: "Zoom",
    "Zoom in": "Vergrößern",
    "Zoom out": "Verkleinern",
    "Collapse ribbon": "Menüband reduzieren",
    "Expand ribbon": "Menüband erweitern",
    "Commands & shortcuts": "Befehle und Tastenkombinationen",
    "Close command reference": "Befehlsübersicht schließen",
    "Close dialog": "Dialog schließen",
    "Close side panel": "Seitenbereich schließen",
    "Dismiss notification": "Meldung schließen",
    "Import workbook file": "Arbeitsmappendatei importieren",
    "Cancel file processing": "Dateiverarbeitung abbrechen",
    "Number format": "Zahlenformat",
    "Font family": "Schriftart",
    "Font size": "Schriftgröße",
    Bold: "Fett",
    Italic: "Kursiv",
    Underline: "Unterstrichen",
    Copy: "Kopieren",
    Cut: "Ausschneiden",
    Paste: "Einfügen",
    "Paste Special": "Inhalte einfügen",
    Alignment: "Ausrichtung",
    Font: "Schrift",
    Clipboard: "Zwischenablage",
    Editing: "Bearbeiten",
    Cells: "Zellen",
    Styles: "Formatvorlagen",
    Number: "Zahl",
    "Format Cells": "Zellen formatieren",
    "Cell Styles": "Zellformatvorlagen",
    "Increase Decimal": "Dezimalstellen erhöhen",
    "Decrease Decimal": "Dezimalstellen verringern",
    General: "Standard",
    Currency: "Währung",
    Percent: "Prozent",
    Date: "Datum",
    Text: "Text",
    Locale: "Gebietsschema",
    "New workbook": "Neue Arbeitsmappe",
    "Open workbook": "Arbeitsmappe öffnen",
    "Save workbook data": "Arbeitsmappendaten speichern",
    "Save as Excel (.xlsx)": "Als Excel speichern (.xlsx)",
    "Export CSV": "CSV exportieren",
    "Export TSV": "TSV exportieren",
    "Workbook properties": "Eigenschaften der Arbeitsmappe",
    "Version history": "Versionsverlauf",
    "Review delimited import": "Import mit Trennzeichen prüfen",
    "Review delimited export": "Export mit Trennzeichen prüfen",
    "Import reviewed population": "Geprüfte Daten importieren",
    "Export reviewed data": "Geprüfte Daten exportieren",
    "Save review report": "Prüfbericht speichern",
    Delimiter: "Trennzeichen",
    Encoding: "Zeichenkodierung",
    "Decimal separator": "Dezimaltrennzeichen",
    "Thousands separator": "Tausendertrennzeichen",
    "Header row": "Kopfzeile",
    "Column types": "Spaltentypen",
    Source: "Quelle",
    Issues: "Probleme",
    Type: "Typ",
    Automatic: "Automatisch",
    Comma: "Komma",
    Semicolon: "Semikolon",
    Tab: "Tabulator",
    Range: "Bereich",
    Import: "Importieren",
    Export: "Exportieren",
    "Inspect calculations": "Berechnungen prüfen",
    "Close calculation inspection": "Berechnungsprüfung schließen",
    "Trace Precedents": "Spur zum Vorgänger",
    "Trace Dependents": "Spur zum Nachfolger",
    "Evaluate Formula": "Formel auswerten",
    "Watch Window": "Überwachungsfenster",
    "Show Formulas": "Formeln anzeigen",
    "Insert Function": "Funktion einfügen",
    Precedents: "Vorgänger",
    Dependents: "Nachfolger",
    Depth: "Tiefe",
    Formula: "Formel",
    Value: "Wert",
    Address: "Adresse",
    Sheet: "Blatt",
    Previous: "Zurück",
    Next: "Weiter",
    Sort: "Sortieren",
    Filter: "Filtern",
    Clear: "Löschen",
    Find: "Suchen",
    Replace: "Ersetzen",
    "Find text": "Suchtext",
    "Replace text": "Ersetzungstext",
    "Close find": "Suche schließen",
    "Filter rows": "Zeilen filtern",
    "Close filter": "Filter schließen",
    "Data Validation": "Datenüberprüfung",
    "Create Table": "Tabelle erstellen",
    PivotTable: "PivotTable",
    Charts: "Diagramme",
    Chart: "Diagramm",
    Picture: "Bild",
    Shapes: "Formen",
    "Review XLSX import": "XLSX-Import prüfen",
    "Review XLSX export": "XLSX-Export prüfen",
    "Open derived workbook": "Abgeleitete Arbeitsmappe öffnen",
    "Download derived XLSX": "Abgeleitete XLSX herunterladen",
    "Save compatibility report": "Kompatibilitätsbericht speichern",
    "Close XLSX review": "XLSX-Prüfung schließen",
    "This workbook is read-only.": "Diese Arbeitsmappe ist schreibgeschützt.",
    "The cell could not be saved.":
      "Die Zelle konnte nicht gespeichert werden.",
    "The workbook changed. Review the file again.":
      "Die Arbeitsmappe wurde geändert. Prüfen Sie die Datei erneut.",
    "Only https://, http:// and mailto: links are supported.":
      "Es werden nur Links mit https://, http:// und mailto: unterstützt.",
    "The file contains invalid or unsupported records. Repair it in the source application and retry; the open workbook is unchanged.":
      "Die Datei enthält ungültige oder nicht unterstützte Datensätze. Korrigieren Sie die Quelldatei und versuchen Sie es erneut. Die geöffnete Arbeitsmappe bleibt unverändert.",
    ...germanWorkbookMessages,
  };
function resolveWorkbookMessages(options = {}) {
  const workbookLocale = resolveWorkbookLocale(options.locale),
    messages = {
      ...(workbookLocale === "de-DE" ? germanWorkbookStatusMessages : {}),
      ...options.messages,
    },
    resolveMessageKey = (key) =>
      Object.hasOwn(messages, key) ? key : key.trim(),
    placeholderSignature = (messageText) =>
      [...new Set(messageText.match(/\{\w+\}/g) ?? [])].sort().join("|"),
    hasTranslation = (candidateText) =>
      Object.hasOwn(messages, resolveMessageKey(candidateText)) &&
      typeof messages[resolveMessageKey(candidateText)] == "string" &&
      messages[resolveMessageKey(candidateText)].trim().length > 0 &&
      placeholderSignature(candidateText) ===
        placeholderSignature(messages[resolveMessageKey(candidateText)]),
    translateText = (sourceText) =>
      hasTranslation(sourceText)
        ? resolveMessageKey(sourceText) === sourceText
          ? messages[sourceText]
          : `${sourceText.match(/^\s*/)?.[0] ?? ""}${messages[resolveMessageKey(sourceText)]}${sourceText.match(/\s*$/)?.[0] ?? ""}`
        : sourceText;
  return {
    locale: workbookLocale,
    entryLocale: options.locale,
    text: translateText,
    language: (text) => (hasTranslation(text) ? workbookLocale : "en"),
    message: (messageTemplate, placeholderValues) =>
      translateText(messageTemplate).replace(
        /\{(\w+)\}/g,
        (placeholder, placeholderName) =>
          Object.hasOwn(placeholderValues, placeholderName)
            ? String(placeholderValues[placeholderName])
            : placeholder,
      ),
  };
}
const WorkbookMessagesContext = ne.createContext(resolveWorkbookMessages());
function WorkbookMessagesProvider({ options, children }) {
  const workbookMessages = ne.useMemo(
    () => resolveWorkbookMessages(options),
    [options?.locale, options?.messages],
  );
  return i.jsx(WorkbookMessagesContext.Provider, {
    value: workbookMessages,
    children,
  });
}
const useWorkbookMessages = () => ne.useContext(WorkbookMessagesContext);
function LocalizedText({ children }) {
  const workbookMessages = useWorkbookMessages(),
    childItems = ne.Children.toArray(children);
  if (!childItems.every((child) => typeof child == "string"))
    return i.jsx(i.Fragment, { children });
  const text = childItems.join("");
  return workbookMessages.locale.startsWith("en") &&
    workbookMessages.text(text) === text
    ? i.jsx(i.Fragment, { children })
    : i.jsx("span", {
        lang: workbookMessages.language(text),
        children: workbookMessages.text(text),
      });
}
function LocalizedNumberField({
  value,
  onValueChange,
  allowEmpty = false,
  min,
  max,
  step,
  onFocus,
  onBlur,
  onKeyDown,
  className,
  ...restProps
}) {
  const workbookMessages = useWorkbookMessages(),
    inputId = ne.useId(),
    [displayText, setDisplayText] = ne.useState(() =>
      value === "" || value === undefined
        ? ""
        : localizeDecimalPoint(Number(value), workbookMessages.entryLocale),
    ),
    [isFocused, setIsFocused] = ne.useState(false),
    [hasBlurred, setHasBlurred] = ne.useState(false),
    entryLocaleRef = ne.useRef(workbookMessages.entryLocale),
    lastEmittedValueRef = ne.useRef(undefined),
    minNumber = min === undefined ? undefined : Number(min),
    maxNumber = max === undefined ? undefined : Number(max),
    parseInputValue = (inputText) => {
      if (!inputText.trim() && allowEmpty) return;
      const parsedValue = validateLocalizedNumberEntry(
        inputText,
        "Value",
        entryLocaleRef.current,
      );
      if (
        (minNumber !== undefined && parsedValue < minNumber) ||
        (maxNumber !== undefined && parsedValue > maxNumber)
      )
        throw Error("bounds");
      return parsedValue;
    };
  let currentNumber,
    errorMessage = "";
  try {
    currentNumber = parseInputValue(displayText);
  } catch (error) {
    errorMessage =
      error instanceof Error && error.message === "bounds"
        ? workbookMessages.message("Enter a number from {min} to {max}.", {
            min:
              minNumber === undefined
                ? "−∞"
                : localizeDecimalPoint(minNumber, workbookMessages.entryLocale),
            max:
              maxNumber === undefined
                ? "∞"
                : localizeDecimalPoint(maxNumber, workbookMessages.entryLocale),
          })
        : workbookMessages.text(
            "Enter a valid number using the selected locale.",
          );
  }
  (ne.useLayoutEffect(() => {
    const valueString =
      value === "" || value === undefined ? "" : String(Number(value));
    (lastEmittedValueRef.current !== undefined &&
      (valueString === lastEmittedValueRef.current ||
        (valueString === "" && lastEmittedValueRef.current === "NaN"))) ||
      ((entryLocaleRef.current = workbookMessages.entryLocale),
      setDisplayText(
        value === "" || value === undefined
          ? ""
          : localizeDecimalPoint(Number(value), workbookMessages.entryLocale),
      ),
      setHasBlurred(false));
  }, [value]),
    ne.useLayoutEffect(() => {
      isFocused ||
        errorMessage ||
        ((entryLocaleRef.current = workbookMessages.entryLocale),
        setDisplayText(
          currentNumber === undefined
            ? ""
            : localizeDecimalPoint(currentNumber, workbookMessages.entryLocale),
        ));
    }, [workbookMessages.entryLocale, isFocused]));
  const handleTextChange = (nextText) => {
      setDisplayText(nextText);
      let emittedValue;
      try {
        const parsedNumber = parseInputValue(nextText);
        emittedValue = parsedNumber === undefined ? "" : String(parsedNumber);
      } catch {
        emittedValue = "NaN";
      }
      ((lastEmittedValueRef.current = emittedValue),
        onValueChange(emittedValue));
    },
    describedByIds =
      [
        restProps["aria-describedby"],
        errorMessage ? `${inputId}-error` : null,
        isFocused && entryLocaleRef.current !== workbookMessages.entryLocale
          ? `${inputId}-locale`
          : null,
      ]
        .filter(Boolean)
        .join(" ") || undefined;
  return i.jsxs("span", {
    className: "lg-number-field",
    children: [
      i.jsx("input", {
        ...restProps,
        className,
        type: "text",
        inputMode: "decimal",
        role: "spinbutton",
        value: displayText,
        "aria-valuenow": currentNumber,
        "aria-valuemin": minNumber,
        "aria-valuemax": maxNumber,
        "aria-valuetext": displayText || undefined,
        "aria-invalid": !!errorMessage || undefined,
        "aria-describedby": describedByIds,
        onChange: (event) => handleTextChange(event.target.value),
        onFocus: (focusEvent) => {
          (errorMessage ||
            (entryLocaleRef.current = workbookMessages.entryLocale),
            setIsFocused(true),
            onFocus?.(focusEvent));
        },
        onBlur: (blurEvent) => {
          (setHasBlurred(true), setIsFocused(false), onBlur?.(blurEvent));
        },
        onKeyDown: (keyEvent) => {
          if (
            (onKeyDown?.(keyEvent),
            !(
              keyEvent.defaultPrevented ||
              restProps.readOnly ||
              restProps.disabled ||
              keyEvent.nativeEvent.isComposing ||
              keyEvent.altKey ||
              keyEvent.metaKey ||
              keyEvent.ctrlKey
            ) &&
              (keyEvent.key === "ArrowUp" || keyEvent.key === "ArrowDown"))
          ) {
            keyEvent.preventDefault();
            const stepSize =
              step === undefined || step === "any" ? 1 : Number(step);
            if (!Number.isFinite(stepSize) || stepSize <= 0) return;
            const nextValue = Math.min(
              maxNumber ?? 1 / 0,
              Math.max(
                minNumber ?? -1 / 0,
                (currentNumber ?? minNumber ?? 0) +
                  (keyEvent.key === "ArrowUp" ? stepSize : -stepSize),
              ),
            );
            Number.isFinite(nextValue) &&
              handleTextChange(
                localizeDecimalPoint(
                  Number(nextValue.toPrecision(15)),
                  entryLocaleRef.current,
                ),
              );
          }
        },
      }),
      errorMessage &&
        i.jsx("span", {
          id: `${inputId}-error`,
          className: hasBlurred ? "lg-number-error" : "lg-sr-only",
          role: hasBlurred ? "status" : undefined,
          children: errorMessage,
        }),
      isFocused &&
        entryLocaleRef.current !== workbookMessages.entryLocale &&
        i.jsx("span", {
          id: `${inputId}-locale`,
          className: "lg-number-error",
          children: workbookMessages.text(
            "The current edit keeps its original numeric locale until you apply or cancel it.",
          ),
        }),
    ],
  });
}
const createEmptyAssistantConversation = () => ({
  id: crypto.randomUUID(),
  draft: "",
  messages: [],
  attachments: [],
});
function findOpenFormulaQuote(formulaText, index) {
  let openQuote = "";
  for (let charIndex = 1; charIndex < index; charIndex++) {
    const currentChar = formulaText[charIndex];
    if (
      openQuote &&
      currentChar === openQuote &&
      formulaText[charIndex + 1] === openQuote
    ) {
      charIndex++;
      continue;
    }
    currentChar === openQuote
      ? (openQuote = "")
      : !openQuote &&
        (currentChar === '"' || currentChar === "'") &&
        (openQuote = currentChar);
  }
  return openQuote;
}
const formulaReferencePattern =
  /(?:'(?:[^']|'')+'!|[A-Za-z_][A-Za-z0-9_.]*!)?(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)/g;
function findFormulaReferenceSpan(formulaText, start, end = start) {
  if (!(
    !formulaText.startsWith("=") ||
    start < 1 ||
    end < start ||
    end > formulaText.length
  )) {
    for (const match of formulaText.matchAll(formulaReferencePattern)) {
      const index = match.index,
        matchEnd = index + match[0].length;
      if (
        !(
          findOpenFormulaQuote(formulaText, index) ||
          /[\w.\[!]/.test(formulaText[index - 1] ?? "") ||
          /[\w.\](]/.test(formulaText[matchEnd] ?? "")
        ) &&
        start >= index &&
        end <= matchEnd
      )
        return { start: index, end: matchEnd };
    }
    if (!findOpenFormulaQuote(formulaText, start)) {
      if (end > start) return { start: start, end: end };
      if (/[=(:,;+\-*/^&<>]$/.test(formulaText.slice(0, start).trimEnd()))
        return { start: start, end: end };
    }
  }
}
function replaceFormulaReferenceSpan(formulaText, span, range, sheetName) {
  const referenceText =
    sheetName === undefined
      ? range
      : `'${sheetName.replaceAll("'", "''")}'!${range}`;
  return {
    value:
      formulaText.slice(0, span.start) +
      referenceText +
      formulaText.slice(span.end),
    span: { start: span.start, end: span.start + referenceText.length },
  };
}
function $ae(formulaText, start, end = start) {
  const formulaReferenceSpan = findFormulaReferenceSpan(
    formulaText,
    start,
    end,
  );
  if (
    !formulaReferenceSpan ||
    formulaReferenceSpan.start === formulaReferenceSpan.end
  )
    return;
  const referenceText = formulaText.slice(
      formulaReferenceSpan.start,
      formulaReferenceSpan.end,
    ),
    bangIndex = referenceText.lastIndexOf("!"),
    sheetPrefix = referenceText.slice(0, bangIndex + 1),
    localAddress = referenceText.slice(bangIndex + 1),
    cellMatch = /^(\$?)[A-Za-z]+(\$?)\d+/.exec(localAddress);
  if (!cellMatch) {
    if (
      !/^(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)$/.test(
        localAddress,
      )
    )
      return;
    const isAbsoluteStart = localAddress.startsWith("$"),
      wholeAxisReference =
        sheetPrefix +
        localAddress
          .split(":")
          .map((rangePart) =>
            isAbsoluteStart
              ? rangePart.replace(/^\$/, "")
              : "$" + rangePart.replace(/^\$/, ""),
          )
          .join(":");
    return {
      value:
        formulaText.slice(0, formulaReferenceSpan.start) +
        wholeAxisReference +
        formulaText.slice(formulaReferenceSpan.end),
      span: {
        start: formulaReferenceSpan.start,
        end: formulaReferenceSpan.start + wholeAxisReference.length,
      },
    };
  }
  const anchorMode =
      !cellMatch[1] && !cellMatch[2]
        ? 1
        : cellMatch[1] && cellMatch[2]
          ? 2
          : !cellMatch[1] && cellMatch[2]
            ? 3
            : 0,
    toggledAddress = localAddress.replace(
      /\$?([A-Za-z]+)\$?(\d+)/g,
      (f, columnLetters, rowDigits) =>
        `${anchorMode === 1 || anchorMode === 3 ? "$" : ""}${columnLetters}${anchorMode === 1 || anchorMode === 2 ? "$" : ""}${rowDigits}`,
    ),
    toggledReference = sheetPrefix + toggledAddress;
  return {
    value:
      formulaText.slice(0, formulaReferenceSpan.start) +
      toggledReference +
      formulaText.slice(formulaReferenceSpan.end),
    span: {
      start: formulaReferenceSpan.start,
      end: formulaReferenceSpan.start + toggledReference.length,
    },
  };
}
function autoCloseFormulaParens(formulaText) {
  if (!formulaText.startsWith("=")) return formulaText;
  let openQuote = "",
    parenDepth = 0,
    bracketDepth = 0;
  for (let charIndex = 1; charIndex < formulaText.length; charIndex++) {
    const character = formulaText[charIndex];
    if (openQuote) {
      character === openQuote &&
        (formulaText[charIndex + 1] === openQuote
          ? charIndex++
          : (openQuote = ""));
      continue;
    }
    if (
      (character === '"' || character === "'"
        ? (openQuote = character)
        : character === "["
          ? bracketDepth++
          : character === "]"
            ? bracketDepth--
            : !bracketDepth && character === "("
              ? parenDepth++
              : !bracketDepth && character === ")" && parenDepth--,
      parenDepth < 0 || bracketDepth < 0)
    )
      return formulaText;
  }
  return !openQuote &&
    !bracketDepth &&
    parenDepth > 0 &&
    parenDepth <= 64 &&
    /[\w\])"%]$/.test(formulaText.trimEnd())
    ? formulaText + ")".repeat(parenDepth)
    : formulaText;
}
const formulaScanCache = new WeakMap();
function getFormulaScanList(cells, listKind) {
  const isFrozenCells = isDeepFrozenValue(cells);
  let scanList = isFrozenCells ? formulaScanCache.get(cells) : undefined;
  if (!scanList) {
    scanList = { calls: [], external: [] };
    for (const address in cells) {
      const cell = cells[address];
      typeof cell.raw != "string" ||
        !cell.raw.startsWith("=") ||
        (cell.raw.includes("(") && scanList.calls.push([address, cell]),
        /[\[|\\]|https?:|file:/i.test(cell.raw) &&
          scanList.external.push([address, cell]));
    }
    isFrozenCells && formulaScanCache.set(cells, scanList);
  }
  return scanList[listKind];
}
const objectIdRegistry = new WeakMap();
let objectIdCounter = 0;
function getStableObjectId(invoke) {
  let objectId = objectIdRegistry.get(invoke);
  return (
    objectId ||
      ((objectId = ++objectIdCounter), objectIdRegistry.set(invoke, objectId)),
    objectId
  );
}
function validateExtensionHostPolicy(workbook, policy) {
  const extensions = policy.extensions ?? [],
    allowed =
      policy.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id);
  if (policy.readOnly !== undefined && typeof policy.readOnly != "boolean")
    throw Error("Read-only policy must be boolean.");
  if (
    !Array.isArray(extensions) ||
    extensions.length > extensionRuntimeLimits.extensions ||
    !Array.isArray(allowed) ||
    allowed.some(
      (allowedId) => !workbook.sheets.some((sheet) => sheet.id === allowedId),
    )
  )
    throw Error("Invalid extension host scope.");
  const seenExtensionIds = new Set();
  for (const extension of extensions) {
    const validatedManifest = validateExtensionManifest(extension.manifest);
    if (
      seenExtensionIds.has(validatedManifest.id) ||
      typeof extension.invoke != "function" ||
      !Array.isArray(extension.allowedSheetIds) ||
      extension.allowedSheetIds.some(
        (allowedSheetId) =>
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ) ||
      !Array.isArray(extension.grantedPermissions) ||
      extension.grantedPermissions.some(
        (grantedPermission) =>
          !validatedManifest.permissions.includes(grantedPermission),
      ) ||
      (extension.includeHidden !== undefined &&
        typeof extension.includeHidden != "boolean")
    )
      throw Error("Invalid host extension registration.");
    seenExtensionIds.add(validatedManifest.id);
  }
  return {
    extensions: extensions,
    allowed: allowed,
    hash: () =>
      fingerprintCanonicalValue([
        !!policy.readOnly,
        allowed,
        extensions.map((extension) => [
          hashExtensionManifest(extension.manifest),
          extension.grantedPermissions,
          extension.allowedSheetIds,
          !!extension.includeHidden,
          getStableObjectId(extension.invoke),
        ]),
      ]),
  };
}
const resolveExtensionSheetScope = (workbook, extensionEntry) =>
  fingerprintCanonicalValue([
    extensionEntry.grantedPermissions,
    [...extensionEntry.allowedSheetIds].sort().map((sheetId) => {
      const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
      return [sheetId, !!sheet.hidden, !!sheet.veryHidden];
    }),
    !!extensionEntry.includeHidden,
  ]);
function buildExtensionHostScope(workbook, hostPolicy) {
  return validateExtensionHostPolicy(workbook, hostPolicy)
    .extensions.filter((extension) =>
      extension.manifest.permissions.every((permission) =>
        extension.grantedPermissions.includes(permission),
      ),
    )
    .map((hostExtension) => ({
      id: hostExtension.manifest.id,
      manifestHash: hashExtensionManifest(hostExtension.manifest),
      scopeHash: resolveExtensionSheetScope(workbook, hostExtension),
      sheetIds: hostExtension.allowedSheetIds.filter((allowedSheetId) => {
        const sheet = workbook.sheets.find(
          (sheet) => sheet.id === allowedSheetId,
        );
        return (
          !(sheet.hidden || sheet.veryHidden) || !!hostExtension.includeHidden
        );
      }),
    }));
}
function withExtensionHostScope(workbook, extensions = []) {
  const scopedWorkbook = { ...workbook };
  return (
    setExtensionCacheEntries(
      scopedWorkbook,
      buildExtensionHostScope(workbook, { extensions: extensions }),
    ),
    isDeepFrozenValue(workbook)
      ? deepFreezeValue(scopedWorkbook)
      : scopedWorkbook
  );
}
function assertExtensionWriteScope(workbook, sheetIds, hostPolicy) {
  if (hostPolicy.readOnly) throw Error("The workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error("Workbook structure is protected.");
  const validatedPolicy = validateExtensionHostPolicy(workbook, hostPolicy);
  for (const sheetId of sheetIds) {
    if (!validatedPolicy.allowed.includes(sheetId))
      throw Error("An extension destination is outside the host write scope.");
    assertSheetUnprotected(
      workbook.sheets.find((sheet) => sheet.id === sheetId),
    );
  }
}
function reviewExtensionFormulas(workbook) {
  const builtinNames = new Set(builtinFunctionNames),
    definedNameSet = new Set(
      (workbook.definedNames ?? []).map((definedName) =>
        definedName.name.toUpperCase(),
      ),
    ),
    reviews = [];
  for (const sheet of workbook.sheets)
    for (const [address, cell] of getFormulaScanList(sheet.cells, "calls")) {
      if (typeof cell.raw != "string" || !cell.raw.startsWith("=")) continue;
      const calls = [],
        collectCalls = (tokens, nested = false) => {
          for (const token of tokens) {
            if (
              (token.kind === "group" && collectCalls(token.body, nested),
              token.kind !== "call")
            )
              continue;
            const callee = token.callee;
            let isCustomCall = false;
            if (callee.kind === "symbol" && !callee.qualifier) {
              const functionName = callee.name
                .replace(/^_xlfn\./i, "")
                .toUpperCase();
              ((isCustomCall =
                !builtinNames.has(functionName) &&
                !definedNameSet.has(functionName) &&
                !["LET", "LAMBDA"].includes(functionName)),
                isCustomCall &&
                  calls.push({
                    name: functionName,
                    start: token.start,
                    end: token.end,
                    arguments: token.args.map((arg) =>
                      arg.length
                        ? cell.raw
                            .toString()
                            .slice(arg[0].start, arg.at(-1).end)
                        : "",
                    ),
                    nested: nested,
                  }));
            }
            token.args.forEach((arg) =>
              collectCalls(arg, nested || isCustomCall),
            );
          }
        };
      try {
        collectCalls(tokenizeFormulaText(cell.raw));
      } catch {
        continue;
      }
      if (
        (calls.length &&
          reviews.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            address: address,
            formula: cell.raw,
            literal: !!cell.literal,
            calls: calls,
          }),
        reviews.length > extensionRuntimeLimits.calls)
      )
        throw Error("Extension review exceeds 100 formula cells.");
    }
  return reviews;
}
function requireApprovedExtension(workbook, extensionId, hostPolicy) {
  const installedEntry = workbook.extensions?.installed.find(
    (entry) => entry.manifest.id === extensionId,
  );
  if (!installedEntry || !installedEntry.enabled)
    throw Error("The extension is missing or disabled.");
  const extension = validateExtensionHostPolicy(
    workbook,
    hostPolicy,
  ).extensions.find((extension) => extension.manifest.id === extensionId);
  if (
    !extension ||
    hashExtensionManifest(extension.manifest) !==
      hashExtensionManifest(installedEntry.manifest)
  )
    throw Error(
      "An approved adapter with the exact extension version and manifest is unavailable.",
    );
  if (
    !installedEntry.manifest.permissions.every((permission) =>
      extension.grantedPermissions.includes(permission),
    )
  )
    throw Error("The host has not granted all declared extension permissions.");
  return extension;
}
function evaluateExtensionFunctionCall(workbook, formulaCell, call, binding) {
  const functionSpec = binding.manifest.functions.find(
    (declaredFunction) =>
      `${binding.manifest.namespace}.${declaredFunction.id}` === call.name,
  );
  if (!functionSpec || functionSpec.parameters.length !== call.arguments.length)
    throw Error(
      "Custom function arguments do not match the declared signature.",
    );
  if (call.nested || /\b(?:LET|LAMBDA)\s*\(/i.test(formulaCell.formula))
    throw Error(
      "Nested custom calls and custom calls inside LET/LAMBDA require separate worksheet cells.",
    );
  const allowRead = (sheetId) => {
    const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
    return (
      !!sheet &&
      binding.allowedSheetIds.includes(sheetId) &&
      (!(sheet.hidden || sheet.veryHidden) || !!binding.includeHidden)
    );
  };
  if (!allowRead(formulaCell.sheetId))
    throw Error(
      "The extension formula is outside its approved worksheet scope.",
    );
  return call.arguments.map((argument, parameterIndex) => {
    if (!argument.trim()) throw Error("Extension arguments cannot be omitted.");
    const schema = functionSpec.parameters[parameterIndex].schema,
      referenceMatch =
        /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?$/.exec(
          argument.trim(),
        );
    if (referenceMatch) {
      const sheetName =
          referenceMatch[1]?.replace(/''/g, "'") ?? referenceMatch[2],
        targetSheet = sheetName
          ? workbook.sheets.find(
              (sheet) => sheet.name.toLowerCase() === sheetName.toLowerCase(),
            )
          : workbook.sheets.find((sheet) => sheet.id === formulaCell.sheetId);
      if (!targetSheet || !allowRead(targetSheet.id))
        throw Error(
          "An extension input worksheet is unavailable or unauthorized.",
        );
      const rangeAddress = parseRangeAddress(
          referenceMatch[3] +
            (referenceMatch[4] ? ":" + referenceMatch[4] : ""),
        ),
        rowCount = rangeAddress.lastRow - rangeAddress.firstRow + 1,
        columnCount = rangeAddress.lastCol - rangeAddress.firstCol + 1;
      if (
        rowCount * columnCount > extensionRuntimeLimits.inputCells ||
        (schema.type !== "matrix" && rowCount * columnCount !== 1)
      )
        throw Error(
          "The extension input range exceeds its declared scalar/cell budget.",
        );
      const inputEvaluation = evaluateWorkbookFormulas(
          workbook,
          {},
          { allowRead: allowRead, allowName: () => true, step: () => {} },
        ),
        inputMatrix = Array.from({ length: rowCount }, (F, rowOffset) =>
          Array.from({ length: columnCount }, (O, columnOffset) => {
            const inputTypedValue = inputEvaluation.getTypedValue(
              targetSheet.id,
              `${columnIndexToLetters(rangeAddress.firstCol + columnOffset)}${rangeAddress.firstRow + rowOffset + 1}`,
            );
            if (inputTypedValue.type === "error")
              throw Error(
                "An extension input contains an unavailable calculation.",
              );
            return inputTypedValue.value;
          }),
        );
      if (inputEvaluation.getOperationalErrors().length)
        throw Error("An extension input attempted an unauthorized dependency.");
      return validateExtensionValue(
        schema.type === "matrix" ? inputMatrix : inputMatrix[0][0],
        schema,
      );
    }
    if (schema.type === "matrix")
      throw Error("Matrix parameters require a static worksheet range.");
    const argumentWorkbook = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === formulaCell.sheetId
          ? {
              ...sheet,
              cells: {
                ...sheet.cells,
                [formulaCell.address]: { raw: "=" + argument },
              },
            }
          : sheet,
      ),
    };
    setExtensionCacheEntries(
      argumentWorkbook,
      getExtensionCacheEntries(workbook),
    );
    const argumentEvaluation = evaluateWorkbookFormulas(
        argumentWorkbook,
        {},
        { allowRead: allowRead, allowName: () => true, step: () => {} },
      ),
      argumentTypedValue = argumentEvaluation.getTypedValue(
        formulaCell.sheetId,
        formulaCell.address,
      );
    if (
      argumentTypedValue.type === "error" ||
      argumentEvaluation.getOperationalErrors().length
    )
      throw Error(
        "An extension input is erroneous, stale, unauthorized or over its calculation limit.",
      );
    let value = argumentTypedValue.value;
    if (
      argumentEvaluation.getSpillRange(formulaCell.sheetId, formulaCell.address)
    )
      throw Error("Scalar extension arguments cannot contain arrays.");
    return validateExtensionValue(value, schema);
  });
}
function resolveExtensionCachedResult(workbook, hostPolicyInput = {}) {
  const validatedPolicy = validateExtensionHostPolicy(
      workbook,
      hostPolicyInput,
    ),
    formulaReviews = reviewExtensionFormulas(workbook),
    evaluationWorkbook = formulaReviews.length
      ? withExtensionHostScope(workbook, validatedPolicy.extensions)
      : workbook,
    evaluation = formulaReviews.length
      ? evaluateWorkbookFormulas(evaluationWorkbook)
      : undefined,
    cells = formulaReviews.map((formulaReview) => {
      const typedResult = evaluation.getTypedValue(
          formulaReview.sheetId,
          formulaReview.address,
        ),
        caches = (workbook.extensions?.caches ?? []).filter(
          (cacheEntry) =>
            cacheEntry.sheetId === formulaReview.sheetId &&
            cacheEntry.address === formulaReview.address,
        ),
        functions = formulaReview.calls.map((call) => {
          const installedFunction = listInstalledExtensionFunctions(
            workbook,
          ).find(
            (installedFunctionItem) => installedFunctionItem.name === call.name,
          );
          let functionState = "unavailable";
          if (
            installedFunction &&
            ((functionState = installedFunction.entry.enabled
              ? buildExtensionHostScope(workbook, hostPolicyInput).some(
                  (hostScopeEntry) =>
                    hostScopeEntry.id === installedFunction.entry.manifest.id &&
                    hostScopeEntry.manifestHash ===
                      hashExtensionManifest(installedFunction.entry.manifest) &&
                    hostScopeEntry.sheetIds.includes(formulaReview.sheetId),
                )
                ? "needs-refresh"
                : "unavailable"
              : "disabled"),
            functionState === "needs-refresh")
          )
            try {
              const binding = requireApprovedExtension(
                  workbook,
                  installedFunction.entry.manifest.id,
                  hostPolicyInput,
                ),
                argumentValues = evaluateExtensionFunctionCall(
                  evaluationWorkbook,
                  formulaReview,
                  call,
                  binding,
                ),
                cacheKey = fingerprintExtensionCall(
                  formulaReview.sheetId,
                  formulaReview.address,
                  formulaReview.formula,
                  call.name,
                  argumentValues,
                  hashExtensionManifest(installedFunction.entry.manifest),
                );
              caches.some(
                (cach) =>
                  cach.key === cacheKey &&
                  cach.scopeHash ===
                    resolveExtensionSheetScope(workbook, binding),
              ) && (functionState = "reviewed-cache");
            } catch {}
          return { name: call.name, state: functionState };
        });
      return {
        ...formulaReview,
        importedCache: workbook.importedCalculationCaches?.entries.find(
          (entry) =>
            entry.sheetId === formulaReview.sheetId &&
            entry.address === formulaReview.address &&
            entry.formula === formulaReview.formula,
        )?.cache,
        functions: functions,
        result: typedResult,
        caches: caches,
        cacheLabel: caches.length
          ? typedResult.type === "error" ||
            functions.some(
              (functionEntry) => functionEntry.state !== "reviewed-cache",
            )
            ? "Historical cached result — stale or unavailable"
            : "Reviewed cached result — manual refresh"
          : "No retained result",
      };
    });
  return {
    installed: workbook.extensions?.installed ?? [],
    cells: cells,
    available: validatedPolicy.extensions.map((extension) => ({
      manifest: extension.manifest,
      fullyGranted: extension.manifest.permissions.every((permission) =>
        extension.grantedPermissions.includes(permission),
      ),
    })),
    policy:
      "Only reviewed argument values reach a host adapter. No code, add-in URL, COM, XLL or ActiveX is loaded from workbook content.",
  };
}
const tT = new WeakMap();
function runExtensionOperation(
  workbook,
  extensionId,
  action,
  hostPolicyInput = {},
) {
  const validatedPolicy = validateExtensionHostPolicy(
      workbook,
      hostPolicyInput,
    ),
    installedEntry = workbook.extensions?.installed.find(
      (entry) => entry.manifest.id === extensionId,
    ),
    extension = validatedPolicy.extensions.find(
      (extension) => extension.manifest.id === extensionId,
    );
  if (!["install", "enable", "disable", "remove"].includes(action))
    throw Error("Unknown extension operation.");
  const targetManifest =
    action === "install" ? extension?.manifest : installedEntry?.manifest;
  if (!targetManifest) throw Error("The requested extension is unavailable.");
  if (
    (validateExtensionManifest(targetManifest),
    ["install", "enable"].includes(action) &&
      (!extension ||
        hashExtensionManifest(extension.manifest) !==
          hashExtensionManifest(targetManifest) ||
        !targetManifest.permissions.every((permission) =>
          extension.grantedPermissions.includes(permission),
        )))
  )
    throw Error(
      "The host has not approved this exact manifest and all of its permissions.",
    );
  if (
    targetManifest.functions.some(
      (functionSpec) =>
        builtinFunctionNames.includes(
          `${targetManifest.namespace}.${functionSpec.id}`,
        ) ||
        (workbook.definedNames ?? []).some(
          (definedName) =>
            definedName.name.toUpperCase() ===
            `${targetManifest.namespace}.${functionSpec.id}`,
        ),
    )
  )
    throw Error(
      "Extension functions cannot shadow built-ins or defined names.",
    );
  const affectedReviews = reviewExtensionFormulas(workbook).filter(
    (formulaReview) =>
      formulaReview.calls.some((call) =>
        call.name.startsWith(targetManifest.namespace + "."),
      ),
  );
  assertExtensionWriteScope(
    workbook,
    affectedReviews.map((reviewItem) => reviewItem.sheetId),
    hostPolicyInput,
  );
  const workbookInput = normalizeWorkbookInput(workbook),
    extensionsState = structuredClone(
      workbookInput.extensions ?? {
        schemaVersion: 1,
        installed: [],
        caches: [],
      },
    );
  ((extensionsState.installed = extensionsState.installed.filter(
    (installedItem) => installedItem.manifest.id !== extensionId,
  )),
    action !== "remove" &&
      extensionsState.installed.push({
        manifest: validateExtensionManifest(targetManifest),
        enabled: action !== "disable",
      }),
    (action === "remove" || action === "install") &&
      (extensionsState.caches = extensionsState.caches.filter(
        (cach) => !cach.functionName.startsWith(targetManifest.namespace + "."),
      )),
    (workbookInput.extensions = extensionsState));
  const normalizedWorkbook = normalizeWorkbookInput(workbookInput),
    operationPlan = {
      extensionId: extensionId,
      action: action,
      manifest: structuredClone(targetManifest),
      warning:
        "Raw formulas are preserved. Disabling retains labeled historical results; installing/upgrading or removing clears this extension’s caches. Execution requires a separate argument review.",
    };
  return (
    tT.set(operationPlan, {
      bookHash: fingerprintCanonicalValue(workbook),
      policyHash: validatedPolicy.hash(),
      planHash: fingerprintCanonicalValue(operationPlan),
      candidate: normalizedWorkbook,
    }),
    operationPlan
  );
}
function applyExtensionPreview(workbook, operationPlan, hostPolicyInput = {}) {
  const previewRecord = tT.get(operationPlan);
  if (!previewRecord)
    throw Error("Unknown or already applied extension preview.");
  if (
    hostPolicyInput.readOnly ||
    previewRecord.bookHash !== fingerprintCanonicalValue(workbook) ||
    previewRecord.policyHash !==
      validateExtensionHostPolicy(workbook, hostPolicyInput).hash() ||
    previewRecord.planHash !== fingerprintCanonicalValue(operationPlan)
  )
    throw Error("Workbook, extension or permissions changed. Preview again.");
  return (
    tT.delete(operationPlan),
    normalizeWorkbookInput(previewRecord.candidate)
  );
}
const k$ = new WeakMap();

export {
  buildFillPreview,
  $ae,
  resolveExtensionCachedResult,
  buildClipboardSnapshot,
  WorkbookMessagesProvider,
  evaluateExtensionFunctionCall,
  useWorkbookMessages,
  buildCellChangePreview,
  requireApprovedExtension,
  parseDateCellEntry,
  withExtensionHostScope,
  buildFillHandlePreview,
  findFormulaReferenceSpan,
  createEmptyAssistantConversation,
  LocalizedText,
  parseLocalizedCellEntry,
  runExtensionOperation,
  buildSeriesPreview,
  parseSeriesSeedDate,
  replaceFormulaReferenceSpan,
  applyExtensionPreview,
  localizeDecimalPoint,
  formatCellEntryForEdit,
  LocalizedNumberField,
  resolveExtensionSheetScope,
  buildPastePreview,
  k$,
  applyCellPreview,
  formatLocalizedNumberDisplay,
  validateExtensionHostPolicy,
  reviewExtensionFormulas,
  validateLocalizedNumberEntry,
  applyCellChangePreview,
  assertExtensionWriteScope,
  getFormulaScanList,
  autoCloseFormulaParens,
};
export { assertWorkbookWritable, eI, formatCellAddress, mergeCellDecorations };
