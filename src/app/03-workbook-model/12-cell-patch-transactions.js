import {
  isNonArrayObject,
  throwWithMessage,
} from "./01-cell-styles-and-themes.js";
import { isDeepFrozenValue } from "../01-schemas-date-functions-and-validation.js";
import { workbookLimits } from "../02-text-and-financial-functions.js";
import {
  throwWorkbookError,
  assertNamedStyleBindingCurrent,
  touchWorkbookModifiedAt,
  assertSheetUnprotected,
  validateCellStyle,
  failWithMessage,
} from "./01-cell-styles-and-themes.js";
import {
  buildNumberFormatCode,
  validateNumberFormatSettings,
} from "./02-formula-function-evaluators.js";
import { carryCalculatorRecordForward } from "./05-formula-engine-and-name-operations.js";
import {
  requireBoundedText,
  NC,
  QU,
  columnIndexToLetters,
} from "./06-data-model-and-analysis.js";
import {
  enumerateSelectionCells,
  _t,
  parseRangeAddress,
  buildBorderPresetStyles,
} from "./07-conditional-rules-worksheet-features.js";
import {
  calculateWorkbookWithSettings,
  checkDataValidationBatch,
} from "./09-calculation-engine.js";
import {
  findSheetOrThrow,
  renameTableHeaderField,
} from "./10-table-filters-and-pivot-views.js";
import {
  normalizeWorkbookModel,
  normalizeCellRecordInput,
  t$,
} from "./11-table-model.js";

function applyCellPatchTransaction(
  workbook,
  patches,
  actor = "Preparer",
  action = "Edit cells",
  options = {},
) {
  if (!Array.isArray(patches) || patches.length > workbookLimits.maxPatches)
    return throwWithMessage(
      `A transaction may contain at most ${workbookLimits.maxPatches} cell changes.`,
    );
  if (patches.length === 0) return workbook;
  (requireBoundedText(actor, "Actor", 100),
    requireBoundedText(action, "Action", 200));
  const changedCellKeys = new Set();
  let calculator;
  const changes = patches.map((patch) => {
      if (!isNonArrayObject(patch))
        return throwWithMessage("Invalid cell patch.");
      const sheetOrThrow = findSheetOrThrow(workbook, patch.sheetId);
      assertSheetUnprotected(sheetOrThrow);
      const { row, col } = _t(patch.address);
      if (row >= sheetOrThrow.rowCount || col >= sheetOrThrow.colCount)
        return throwWithMessage(
          `Cell ${patch.address} is outside sheet dimensions.`,
        );
      const address = `${columnIndexToLetters(col)}${row + 1}`;
      assertNamedStyleBindingCurrent(
        workbook,
        normalizeCellRecordInput(patch.cell).style,
      );
      const cell = sheetOrThrow.cells[address];
      if (
        (cell?.raw ?? null) !== patch.cell?.raw ||
        !!cell?.literal != !!patch.cell?.literal
      ) {
        calculator ??= calculateWorkbookWithSettings(workbook);
        const spillAnchorAddress = calculator.getSpillAnchor(
          sheetOrThrow.id,
          address,
        );
        if (spillAnchorAddress && spillAnchorAddress !== address) {
          const patch = patches.find(
            (patch) =>
              patch.sheetId === sheetOrThrow.id &&
              patch.address === spillAnchorAddress,
          );
          if (
            !patch ||
            (!patch.cell.literal &&
              typeof patch.cell.raw == "string" &&
              patch.cell.raw.startsWith("="))
          )
            return throwWithMessage(
              `Cell ${address} belongs to the array formula in ${spillAnchorAddress}. Edit the array formula instead.`,
            );
        }
      }
      const cellKey = `${sheetOrThrow.id}:${address}`;
      return changedCellKeys.has(cellKey)
        ? throwWithMessage("A transaction cannot change the same cell twice.")
        : (changedCellKeys.add(cellKey),
          {
            sheetId: sheetOrThrow.id,
            address: address,
            before: sheetOrThrow.cells[address]
              ? normalizeCellRecordInput(sheetOrThrow.cells[address])
              : undefined,
            after: normalizeCellRecordInput(patch.cell),
          });
    }),
    newCellCount = changes.filter((change) => !change.before).length;
  if (
    workbook.sheets.reduce(
      (populatedCellCount, sheet) =>
        populatedCellCount +
        (isDeepFrozenValue(sheet.cells)
          ? (QU.get(sheet.cells) ?? Object.keys(sheet.cells).length)
          : Object.keys(sheet.cells).length),
      0,
    ) +
      newCellCount >
    workbookLimits.maxCells
  )
    return throwWithMessage(
      "Workbook exceeds the 100,000 populated cell limit.",
    );
  const sheets = workbook.sheets.map((sheet) => {
      const sheetChanges = changes.filter(
        (change) => change.sheetId === sheet.id,
      );
      if (!sheetChanges.length) return sheet;
      const cells = { ...sheet.cells };
      return (
        sheetChanges.forEach((change) => {
          cells[change.address] = change.after;
        }),
        { ...sheet, cells: cells }
      );
    }),
    book = renameTableHeaderField(workbook, { ...workbook, sheets: sheets });
  if (book.sheets !== sheets)
    for (const sheet of book.sheets) {
      const sheetOrThrow = findSheetOrThrow(workbook, sheet.id);
      for (const [address, cell] of Object.entries(sheet.cells)) {
        const change = changes.find(
          (change) => change.sheetId === sheet.id && change.address === address,
        );
        if (change) change.after = normalizeCellRecordInput(cell);
        else if (
          JSON.stringify(sheetOrThrow.cells[address]) !==
            JSON.stringify(cell) &&
          (changes.push({
            sheetId: sheet.id,
            address: address,
            before: sheetOrThrow.cells[address]
              ? normalizeCellRecordInput(sheetOrThrow.cells[address])
              : undefined,
            after: normalizeCellRecordInput(cell),
          }),
          changes.length > workbookLimits.maxPatches)
        )
          return throwWithMessage(
            "Renaming this table field affects more than 500 cells. Reduce the dependent formulas before renaming it.",
          );
      }
    }
  const validationWarningCount = checkDataValidationBatch(
    book,
    changes.filter(
      (change) =>
        change.before?.raw !== change.after.raw ||
        change.before?.literal !== change.after.literal,
    ),
    options.validationApproval,
  );
  if (
    (validationWarningCount &&
      (action = `${action.slice(0, 140)} (accepted ${validationWarningCount} validation warnings)`),
    workbook.events.length >= 2e3)
  )
    return throwWithMessage(
      "Change history has reached its 2,000 transaction limit. Export the workbook before starting a new workpaper.",
    );
  if (
    workbook.events.reduce(
      (changeCount, entry) => changeCount + entry.changes.length,
      0,
    ) +
      changes.length >
    workbookLimits.maxCells
  )
    return throwWithMessage(
      "Change history has reached its 100,000 cell-change limit. Export the workbook before starting a new workpaper.",
    );
  if (workbook.revision >= Number.MAX_SAFE_INTEGER - 1)
    return throwWithMessage("Workbook revision limit reached.");
  const workbookModel = {
    ...touchWorkbookModifiedAt(book),
    revision: workbook.revision + 1,
    sheets: book.sheets,
    events: [
      ...workbook.events,
      {
        id: crypto.randomUUID(),
        timestamp: new Date().toISOString(),
        actor: actor,
        action: action,
        changes: changes,
      },
    ],
  };
  return (
    isDeepFrozenValue(workbook) &&
      NC.has(workbook) &&
      book.sheets === sheets &&
      changes.every(
        (change) => !change.after.hyperlink && !change.after.style?.binding,
      ) &&
      t$(workbookModel),
    carryCalculatorRecordForward(workbook, workbookModel),
    workbookModel
  );
}
function applyBorderEditTransaction(
  workbook,
  sheetId,
  selectionAreas,
  stylePatch,
  options = {},
) {
  options.readOnly && failWithMessage("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const validatedStyle = validateCellStyle(stylePatch),
    selectionAddresses = enumerateSelectionCells(
      typeof selectionAreas == "string" ? [selectionAreas] : selectionAreas,
    ),
    sampleGridAddresses = new Set(selectionAddresses);
  if (
    options.borders ||
    Object.hasOwn(stylePatch, "borders") ||
    Object.hasOwn(stylePatch, "border")
  ) {
    const selectionPositions = selectionAddresses.map(_t);
    (sheetOrThrow.merges ?? []).some((mergeRange) => {
      const rangeAddress = parseRangeAddress(mergeRange);
      return selectionPositions.some(
        (position) =>
          position.row >= rangeAddress.firstRow &&
          position.row <= rangeAddress.lastRow &&
          position.col >= rangeAddress.firstCol &&
          position.col <= rangeAddress.lastCol,
      );
    }) &&
      failWithMessage(
        "Unmerge the intersecting cells before applying border edges or presets.",
      );
  }
  const patches = selectionAddresses.flatMap((address) => {
    const existingCell = sheetOrThrow.cells[address],
      mergedStyle = {
        ...existingCell?.style,
        ...stylePatch,
        ...validatedStyle,
      };
    (Object.hasOwn(stylePatch, "border") && delete mergedStyle.borders,
      options.borders &&
        (mergedStyle.borders = {
          ...existingCell?.style?.borders,
          ...buildBorderPresetStyles(
            address,
            sampleGridAddresses,
            options.borders.preset,
            options.borders.line,
          ),
        }),
      stylePatch.borders !== undefined &&
        (mergedStyle.borders = {
          ...(options.borders
            ? mergedStyle.borders
            : existingCell?.style?.borders),
          ...validatedStyle.borders,
        }));
    const nextStyle = validateCellStyle(mergedStyle),
      cell = {
        ...(existingCell ?? { raw: null }),
        ...(Object.keys(nextStyle).length ? { style: nextStyle } : {}),
      };
    return (
      Object.keys(nextStyle).length || delete cell.style,
      !existingCell && !Object.keys(nextStyle).length
        ? []
        : JSON.stringify(validateCellStyle(existingCell?.style ?? {})) ===
            JSON.stringify(nextStyle)
          ? []
          : JSON.stringify(existingCell) === JSON.stringify(cell)
            ? []
            : [{ sheetId: sheetId, address: address, cell: cell }]
    );
  });
  if (!patches.length) return workbook;
  const patchedWorkbook = applyCellPatchTransaction(
    workbook,
    patches,
    "Preparer",
    "Format cell styles",
  );
  return { ...workbook, sheets: patchedWorkbook.sheets };
}
function applyNumberFormatTransaction(
  workbook,
  sheetId,
  selectionAreas,
  numberFormatInput,
  options = {},
) {
  options.readOnly && throwWorkbookError("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const numberFormat = validateNumberFormatSettings(numberFormatInput),
    selectionAddresses = enumerateSelectionCells(
      typeof selectionAreas == "string" ? [selectionAreas] : selectionAreas,
    ),
    numberFormatCode = buildNumberFormatCode(numberFormat),
    patches = selectionAddresses.flatMap((address) => {
      const existingCell = sheetOrThrow.cells[address],
        cell = {
          ...(existingCell ?? { raw: null }),
          numberFormat: numberFormat,
          numFmt: numberFormatCode,
          format:
            numberFormat.category === "percent"
              ? "percent"
              : numberFormat.category === "currency" ||
                  numberFormat.category === "accounting"
                ? "currency"
                : numberFormat.category === "date" ||
                    numberFormat.category === "datetime"
                  ? "date"
                  : numberFormat.category === "number"
                    ? "number"
                    : "general",
        };
      return JSON.stringify(existingCell) === JSON.stringify(cell)
        ? []
        : [{ sheetId: sheetId, address: address, cell: cell }];
    });
  if (!patches.length) return workbook;
  const patchedWorkbook = applyCellPatchTransaction(
    workbook,
    patches,
    "Preparer",
    "Format numbers",
  );
  return { ...workbook, sheets: patchedWorkbook.sheets };
}
function buildCanonicalWorkbookSnapshot(workbook) {
  const workbookModel = normalizeWorkbookModel(workbook, true);
  return NC.has(workbookModel) && isDeepFrozenValue(workbookModel)
    ? workbookModel
    : t$(workbookModel);
}

export {
  applyBorderEditTransaction,
  applyNumberFormatTransaction,
  applyCellPatchTransaction,
  buildCanonicalWorkbookSnapshot,
};
