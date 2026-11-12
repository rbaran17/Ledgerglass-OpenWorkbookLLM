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
  o = {},
) {
  if (!Array.isArray(patches) || patches.length > workbookLimits.maxPatches)
    return throwWithMessage(
      `A transaction may contain at most ${workbookLimits.maxPatches} cell changes.`,
    );
  if (patches.length === 0) return workbook;
  (requireBoundedText(actor, "Actor", 100),
    requireBoundedText(action, "Action", 200));
  const a = new Set();
  let s;
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
        s ??= calculateWorkbookWithSettings(workbook);
        const k = s.getSpillAnchor(sheetOrThrow.id, address);
        if (k && k !== address) {
          const F = patches.find(
            (I) => I.sheetId === sheetOrThrow.id && I.address === k,
          );
          if (
            !F ||
            (!F.cell.literal &&
              typeof F.cell.raw == "string" &&
              F.cell.raw.startsWith("="))
          )
            return throwWithMessage(
              `Cell ${address} belongs to the array formula in ${k}. Edit the array formula instead.`,
            );
        }
      }
      const w = `${sheetOrThrow.id}:${address}`;
      return a.has(w)
        ? throwWithMessage("A transaction cannot change the same cell twice.")
        : (a.add(w),
          {
            sheetId: sheetOrThrow.id,
            address: address,
            before: sheetOrThrow.cells[address]
              ? normalizeCellRecordInput(sheetOrThrow.cells[address])
              : undefined,
            after: normalizeCellRecordInput(patch.cell),
          });
    }),
    u = changes.filter((change) => !change.before).length;
  if (
    workbook.sheets.reduce(
      (p, m) =>
        p +
        (isDeepFrozenValue(m.cells)
          ? (QU.get(m.cells) ?? Object.keys(m.cells).length)
          : Object.keys(m.cells).length),
      0,
    ) +
      u >
    workbookLimits.maxCells
  )
    return throwWithMessage(
      "Workbook exceeds the 100,000 populated cell limit.",
    );
  const sheets = workbook.sheets.map((sheet) => {
      const m = changes.filter((change) => change.sheetId === sheet.id);
      if (!m.length) return sheet;
      const cells = { ...sheet.cells };
      return (
        m.forEach((g) => {
          cells[g.address] = g.after;
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
          (y) => y.sheetId === sheet.id && y.address === address,
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
  const h = checkDataValidationBatch(
    book,
    changes.filter(
      (change) =>
        change.before?.raw !== change.after.raw ||
        change.before?.literal !== change.after.literal,
    ),
    o.validationApproval,
  );
  if (
    (h &&
      (action = `${action.slice(0, 140)} (accepted ${h} validation warnings)`),
    workbook.events.length >= 2e3)
  )
    return throwWithMessage(
      "Change history has reached its 2,000 transaction limit. Export the workbook before starting a new workpaper.",
    );
  if (
    workbook.events.reduce((p, m) => p + m.changes.length, 0) + changes.length >
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
  n,
  o = {},
) {
  o.readOnly && failWithMessage("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const s = validateCellStyle(n),
    d = enumerateSelectionCells(
      typeof selectionAreas == "string" ? [selectionAreas] : selectionAreas,
    ),
    sampleGridAddresses = new Set(d);
  if (o.borders || Object.hasOwn(n, "borders") || Object.hasOwn(n, "border")) {
    const h = d.map(_t);
    (sheetOrThrow.merges ?? []).some((f) => {
      const rangeAddress = parseRangeAddress(f);
      return h.some(
        (m) =>
          m.row >= rangeAddress.firstRow &&
          m.row <= rangeAddress.lastRow &&
          m.col >= rangeAddress.firstCol &&
          m.col <= rangeAddress.lastCol,
      );
    }) &&
      failWithMessage(
        "Unmerge the intersecting cells before applying border edges or presets.",
      );
  }
  const patches = d.flatMap((address) => {
    const f = sheetOrThrow.cells[address],
      p = { ...f?.style, ...n, ...s };
    (Object.hasOwn(n, "border") && delete p.borders,
      o.borders &&
        (p.borders = {
          ...f?.style?.borders,
          ...buildBorderPresetStyles(
            address,
            sampleGridAddresses,
            o.borders.preset,
            o.borders.line,
          ),
        }),
      n.borders !== undefined &&
        (p.borders = {
          ...(o.borders ? p.borders : f?.style?.borders),
          ...s.borders,
        }));
    const m = validateCellStyle(p),
      cell = {
        ...(f ?? { raw: null }),
        ...(Object.keys(m).length ? { style: m } : {}),
      };
    return (
      Object.keys(m).length || delete cell.style,
      !f && !Object.keys(m).length
        ? []
        : JSON.stringify(validateCellStyle(f?.style ?? {})) ===
            JSON.stringify(m)
          ? []
          : JSON.stringify(f) === JSON.stringify(cell)
            ? []
            : [{ sheetId: sheetId, address: address, cell: cell }]
    );
  });
  if (!patches.length) return workbook;
  const c = applyCellPatchTransaction(
    workbook,
    patches,
    "Preparer",
    "Format cell styles",
  );
  return { ...workbook, sheets: c.sheets };
}
function applyNumberFormatTransaction(
  workbook,
  sheetId,
  selectionAreas,
  n,
  o = {},
) {
  o.readOnly && throwWorkbookError("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const numberFormat = validateNumberFormatSettings(n),
    d = enumerateSelectionCells(
      typeof selectionAreas == "string" ? [selectionAreas] : selectionAreas,
    ),
    numberFormatCode = buildNumberFormatCode(numberFormat),
    patches = d.flatMap((address) => {
      const f = sheetOrThrow.cells[address],
        cell = {
          ...(f ?? { raw: null }),
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
      return JSON.stringify(f) === JSON.stringify(cell)
        ? []
        : [{ sheetId: sheetId, address: address, cell: cell }];
    });
  if (!patches.length) return workbook;
  const c = applyCellPatchTransaction(
    workbook,
    patches,
    "Preparer",
    "Format numbers",
  );
  return { ...workbook, sheets: c.sheets };
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
