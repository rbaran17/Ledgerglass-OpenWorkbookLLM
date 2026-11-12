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

const formatRangeAddress = (e, r, t, n) => {
  const cellAddress = formatCellAddress(e, r),
    a = formatCellAddress(e + t - 1, r + n - 1);
  return cellAddress === a ? cellAddress : `${cellAddress}:${a}`;
};
function eI(e) {
  return (
    e &&
      typeof e == "object" &&
      (Object.values(e).forEach(eI), Object.freeze(e)),
    e
  );
}
function stripCellDecorations(e) {
  const cellRecordInput = normalizeCellRecordInput(e ?? { raw: null });
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
function mergeCellDecorations(fillCellPatches, r) {
  return {
    ...fillCellPatches,
    ...(r?.note !== undefined ? { note: r.note } : {}),
    ...(r?.hyperlink !== undefined
      ? { hyperlink: structuredClone(r.hyperlink) }
      : {}),
    ...(r?.tickmark !== undefined ? { tickmark: r.tickmark } : {}),
  };
}
function cellContentsMatch(e, cell) {
  return (
    JSON.stringify(normalizeCellRecordInput(e ?? { raw: null })) ===
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
  const cells = sheet.cells.map((cell, o) => {
    (!isPlainObjectValue(cell) ||
      cell.rowOffset !== Math.floor(o / sheet.colCount) ||
      cell.colOffset !== o % sheet.colCount) &&
      throwWorkbookError(
        "Clipboard cells must form a complete row-major rectangle.",
      );
    const a = stripCellDecorations(normalizeCellRecordInput(cell.cell)),
      raw = normalizeCellRecordInput({ raw: cell.value }).raw;
    return {
      rowOffset: cell.rowOffset,
      colOffset: cell.colOffset,
      cell: a,
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
        const m = workbookWithSettings.getSpillAnchor(
          sourceSheetId,
          cellAddress,
        );
        if (m && m !== cellAddress) {
          const v = _t(m);
          (v.row < rangeAddress.firstRow ||
            v.row > rangeAddress.lastRow ||
            v.col < rangeAddress.firstCol ||
            v.col > rangeAddress.lastCol) &&
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
  r,
  t,
  readOnly,
  deferValidationWarnings = false,
) {
  (assertWorkbookWritable(readOnly),
    assertSheetUnprotected(findSheetOrThrow(workbook, t.targetSheetId)),
    r.length > workbookLimits.maxPatches &&
      throwWorkbookError(
        "A fill or paste operation may change at most 500 cells.",
      ));
  const patches = r.map((patch) => ({
    sheetId: patch.sheetId,
    address: patch.address,
    cell: normalizeCellRecordInput(patch.cell),
  }));
  let s;
  try {
    s = applyCellPatchTransaction(
      workbook,
      patches,
      "Preview",
      t.summary.slice(0, 200),
    );
  } catch (error) {
    if (
      !deferValidationWarnings ||
      !(error instanceof DataValidationError) ||
      error.issues.some((issue) => issue.style === "stop")
    )
      throw error;
    s = applyCellPatchTransaction(
      workbook,
      patches,
      "Preview",
      t.summary.slice(0, 200),
      {
        validationApproval: assertNoStopValidationIssues(error),
      },
    );
  }
  const previewPatches =
      s === workbook
        ? []
        : s.events.at(-1).changes.map((change) => ({
            sheetId: change.sheetId,
            address: change.address,
            cell: normalizeCellRecordInput(change.after ?? { raw: null }),
          })),
    count = previewPatches.length,
    dependentCount = Math.max(0, count - patches.length);
  return eI({
    ...t,
    workbookId: workbook.id,
    revision: workbook.revision,
    patches: patches,
    previewPatches: previewPatches,
    count: count,
    dependentCount: dependentCount,
    summary:
      t.summary +
      (dependentCount
        ? ` Includes ${dependentCount} dependent cell changes.`
        : ""),
    snapshotGuard: JSON.stringify(workbook),
  });
}
function applyCellChangePreview(workbook, plan, t = {}) {
  (assertWorkbookWritable(t.readOnly),
    (!plan ||
      plan.workbookId !== workbook.id ||
      plan.revision !== workbook.revision ||
      plan.snapshotGuard !== JSON.stringify(workbook)) &&
      throwWorkbookError(
        "The workbook changed after this preview. Create a new preview before applying it.",
      ),
    assertSheetUnprotected(findSheetOrThrow(workbook, plan.targetSheetId)));
  const n = applyCellPatchTransaction(
    workbook,
    plan.patches,
    t.actor ?? "Preparer",
    plan.summary.slice(0, 200),
  );
  return n === workbook
    ? workbook
    : { ...n, revision: workbook.revision, events: workbook.events };
}
function buildPastePreview(workbook, sheetId, t, n, o = {}) {
  const sheet = validateClipboardSnapshot(n),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(t),
    u = parseRangeAddress(sheet.sourceRange),
    l = o.mode ?? "all",
    c = o.operation ?? "none";
  ((!["all", "values", "formulas", "formats"].includes(l) ||
    !["none", "add", "subtract", "multiply", "divide"].includes(c)) &&
    throwWorkbookError(
      "Choose a supported paste mode and arithmetic operation.",
    ),
    [o.transpose, o.skipBlanks].some(
      (k) => k !== undefined && typeof k != "boolean",
    ) && throwWorkbookError("Transpose and skip blanks must be boolean."),
    c !== "none" &&
      (l === "formats" || l === "formulas") &&
      throwWorkbookError(
        "Arithmetic paste requires All or Values mode; its results are numeric values.",
      ),
    sheet.sourceWorkbookId !== workbook.id &&
      (l === "all" || l === "formulas") &&
      sheet.cells.some(
        (cell) =>
          typeof cell.cell.raw == "string" &&
          cell.cell.raw.startsWith("=") &&
          !cell.cell.literal,
      ) &&
      throwWorkbookError(
        "Use Values paste for formulas copied from a different workbook.",
      ));
  const h = o.transpose ? sheet.colCount : sheet.rowCount,
    f = o.transpose ? sheet.rowCount : sheet.colCount,
    p =
      rangeAddress.firstRow === rangeAddress.lastRow &&
      rangeAddress.firstCol === rangeAddress.lastCol,
    m = p ? h : rangeAddress.lastRow - rangeAddress.firstRow + 1,
    v = p ? f : rangeAddress.lastCol - rangeAddress.firstCol + 1;
  ((m % h || v % f) &&
    throwWorkbookError(
      "The target rectangle must be a whole multiple of the copied rows and columns.",
    ),
    m * v > workbookLimits.maxPatches &&
      throwWorkbookError("Paste into at most 500 target cells."),
    (rangeAddress.firstRow + m > sheetOrThrow.rowCount ||
      rangeAddress.firstCol + v > sheetOrThrow.colCount) &&
      throwWorkbookError(
        "The pasted cells would exceed the worksheet boundary.",
      ));
  const g = formatRangeAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
      m,
      v,
    ),
    b = rangeAddress.firstRow + m - 1,
    y = rangeAddress.firstCol + v - 1;
  for (const k of sheetOrThrow.merges ?? []) {
    const F = parseRangeAddress(k);
    rangeAddress.firstRow <= F.lastRow &&
      b >= F.firstRow &&
      rangeAddress.firstCol <= F.lastCol &&
      y >= F.firstCol &&
      !(
        m * v === 1 &&
        rangeAddress.firstRow === F.firstRow &&
        rangeAddress.firstCol === F.firstCol
      ) &&
      throwWorkbookError(
        "Unmerge the target area before pasting a block of cells or editing a covered cell.",
      );
  }
  const w = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  let skippedBlanks = 0;
  for (let k = 0; k < m; k++)
    for (let F = 0; F < v; F++) {
      const I = o.transpose ? F % f : k % h,
        O = o.transpose ? k % h : F % f,
        j = sheet.cells[I * sheet.colCount + O],
        z = (j.cell.raw === null && j.value === null) || j.cell.raw === "";
      if (o.skipBlanks && z) {
        skippedBlanks++;
        continue;
      }
      const cellAddress = formatCellAddress(
          rangeAddress.firstRow + k,
          rangeAddress.firstCol + F,
        ),
        D = sheetOrThrow.cells[cellAddress];
      let cell = stripCellDecorations(j.cell);
      typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        !cell.literal &&
        (cell.raw = shiftFormulaReferences(
          cell.raw,
          rangeAddress.firstRow + k - u.firstRow - I,
          rangeAddress.firstCol + F - u.firstCol - O,
        ));
      let _;
      if (l === "all") _ = mergeCellDecorations(cell, D);
      else if (l === "formulas")
        _ = { ...(D ?? { raw: null }), raw: cell.raw, literal: cell.literal };
      else if (l === "values")
        _ = withRawCellValue(j.value, { ...(D ?? { raw: null }) });
      else {
        _ = { ...(D ?? { raw: null }) };
        for (const C of ["style", "format", "numFmt", "numberFormat"])
          delete _[C];
        (cell.style !== undefined && (_.style = cell.style),
          cell.format !== undefined && (_.format = cell.format),
          cell.numFmt !== undefined && (_.numFmt = cell.numFmt),
          cell.numberFormat !== undefined &&
            (_.numberFormat = structuredClone(cell.numberFormat)));
      }
      if ((assertNamedStyleBindingCurrent(workbook, _.style), c !== "none")) {
        const C = j.value === null || j.value === "" ? 0 : j.value,
          T = workbookWithSettings(sheetId, cellAddress),
          N = T === null || T === "" ? 0 : T;
        ((typeof C != "number" ||
          !Number.isFinite(C) ||
          typeof N != "number" ||
          !Number.isFinite(N)) &&
          throwWorkbookError(
            `Arithmetic paste requires numbers or blanks at ${cellAddress}.`,
          ),
          c === "divide" &&
            C === 0 &&
            throwWorkbookError(
              `Arithmetic paste would divide by zero at ${cellAddress}.`,
            ));
        const raw =
          c === "add"
            ? N + C
            : c === "subtract"
              ? N - C
              : c === "multiply"
                ? N * C
                : N / C;
        (Number.isFinite(raw) ||
          throwWorkbookError(
            `Arithmetic paste would overflow at ${cellAddress}.`,
          ),
          (_ = { ..._, raw: raw }),
          delete _.literal);
      }
      ((_ = normalizeCellRecordInput(_)),
        cellContentsMatch(D, _) ||
          w.push({ sheetId: sheetId, address: cellAddress, cell: _ }));
    }
  return buildCellChangePreview(
    workbook,
    w,
    {
      targetSheetId: sheetId,
      targetRange: g,
      targetCount: m * v,
      sourceRange: sheet.sourceRange,
      skippedBlanks: skippedBlanks,
      summary: `Paste ${l}${o.transpose ? " transposed" : ""}${c !== "none" ? ` (${c})` : ""} into ${g}.`,
    },
    o.readOnly,
    o.deferValidationWarnings,
  );
}
function applyCellPreview(workbook, plan, t = {}) {
  return applyCellChangePreview(workbook, plan, t);
}

function validateFillDirection(direction) {
  return (
    ["down", "right", "up", "left"].includes(direction) ||
      throwWorkbookError("Choose a supported fill direction."),
    direction
  );
}
function formatRangeBoundsAddress(e) {
  const rangeAddress = parseRangeAddress(e),
    cellAddress = formatCellAddress(
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
    n = formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol);
  return cellAddress === n ? cellAddress : `${cellAddress}:${n}`;
}

function cellContentsDiffer(e, cellRecordInput) {
  return (
    JSON.stringify(normalizeCellRecordInput(e ?? { raw: null })) !==
    JSON.stringify(normalizeCellRecordInput(cellRecordInput))
  );
}
function buildFillCellPatches(
  sheetOrThrow,
  r,
  workbookWithSettings,
  rangeAddress,
) {
  const cell = stripCellDecorations(sheetOrThrow.cells[r]);
  if (cell.raw !== null) return cell;
  const a = workbookWithSettings(sheetOrThrow.id, r);
  if (a === null) return cell;
  const s = workbookWithSettings.getSpillAnchor(sheetOrThrow.id, r);
  if (s && s !== r) {
    const d = _t(s);
    (d.row < rangeAddress.firstRow ||
      d.row > rangeAddress.lastRow ||
      d.col < rangeAddress.firstCol ||
      d.col > rangeAddress.lastCol) &&
      ((cell.raw = a),
      typeof a == "string" && a.startsWith("=") && (cell.literal = true));
  }
  return cell;
}
function assertFillTargetUnmerged(workbook, sheetId, t) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  for (const o of sheetOrThrow.merges ?? []) {
    const rangeAddress = parseRangeAddress(o);
    t.some((s) => {
      const d = _t(s);
      return (
        d.row >= rangeAddress.firstRow &&
        d.row <= rangeAddress.lastRow &&
        d.col >= rangeAddress.firstCol &&
        d.col <= rangeAddress.lastCol &&
        (d.row !== rangeAddress.firstRow || d.col !== rangeAddress.firstCol)
      );
    }) &&
      throwWorkbookError("Unmerge covered target cells before filling them.");
  }
}
function buildFillPreview(workbook, sheetId, t, n, o = {}) {
  const direction = validateFillDirection(n),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(t);
  (rangeAddress.lastRow >= sheetOrThrow.rowCount ||
    rangeAddress.lastCol >= sheetOrThrow.colCount) &&
    throwWorkbookError("The fill range exceeds the worksheet.");
  const u = direction === "down" || direction === "up",
    l = rangeAddress.lastRow - rangeAddress.firstRow + 1,
    c = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    h = u ? (l - 1) * c : (c - 1) * l;
  h > workbookLimits.maxPatches &&
    throwWorkbookError("Fill at most 500 destination cells.");
  const f = direction === "up" ? rangeAddress.lastRow : rangeAddress.firstRow,
    p = direction === "left" ? rangeAddress.lastCol : rangeAddress.firstCol,
    m = [],
    v = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    b = u
      ? { ...rangeAddress, firstRow: f, lastRow: f }
      : { ...rangeAddress, firstCol: p, lastCol: p };
  if (h)
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
        if (u ? firstRow === f : firstCol === p) continue;
        const E = u ? f : firstRow,
          A = u ? firstCol : p,
          cellAddress = formatCellAddress(firstRow, firstCol);
        v.push(cellAddress);
        const fillCellPatches = buildFillCellPatches(
          sheetOrThrow,
          formatCellAddress(E, A),
          workbookWithSettings,
          b,
        );
        typeof fillCellPatches.raw == "string" &&
          fillCellPatches.raw.startsWith("=") &&
          !fillCellPatches.literal &&
          (fillCellPatches.raw = shiftFormulaReferences(
            fillCellPatches.raw,
            firstRow - E,
            firstCol - A,
          ));
        const cellRecordInput = normalizeCellRecordInput(
          mergeCellDecorations(
            fillCellPatches,
            sheetOrThrow.cells[cellAddress],
          ),
        );
        cellContentsDiffer(sheetOrThrow.cells[cellAddress], cellRecordInput) &&
          m.push({
            sheetId: sheetId,
            address: cellAddress,
            cell: cellRecordInput,
          });
      }
  return (
    assertFillTargetUnmerged(workbook, sheetId, v),
    buildCellChangePreview(
      workbook,
      m,
      {
        targetSheetId: sheetId,
        targetRange: formatRangeBoundsAddress(t),
        targetCount: h,
        generatedCount: h,
        direction: direction,
        summary: `Fill ${direction} within ${formatRangeBoundsAddress(t)}.`,
      },
      o.readOnly,
    )
  );
}

function excelSerialToDate(e) {
  return new Date(excelEpochUtc + (e > 60 ? e - 1 : e) * s$);
}

function stepSeriesDateSerial(e, r, t, n) {
  if (
    (Number.isSafeInteger(r) ||
      throwWorkbookError("The date step exceeds safe integer limits."),
    !r)
  )
    return e;
  if (t === "day") {
    const c = e + r;
    return (
      (c < 0 || c > maxExcelDateSerial) && throwDateSeriesRangeError(),
      parseSeriesSeedDate(c)
    );
  }
  e === 60 &&
    throwWorkbookError(
      "Calendar and weekday series require a real calendar seed; use Days for the fictitious 1900-02-29 serial.",
    );
  const o = excelSerialToDate(e);
  if (t === "weekday") {
    const c = Math.sign(r);
    let h = Math.abs(r);
    if (o.getUTCDay() === 0 || o.getUTCDay() === 6) {
      do o.setUTCDate(o.getUTCDate() + c);
      while ([0, 6].includes(o.getUTCDay()));
      h--;
    }
    for (o.setUTCDate(o.getUTCDate() + c * Math.floor(h / 5) * 7), h %= 5; h;)
      (o.setUTCDate(o.getUTCDate() + c), [0, 6].includes(o.getUTCDay()) || h--);
    return dateToExcelSerial(o);
  }
  const a = o.getUTCMonth() + r * (t === "year" ? 12 : 1),
    s = o.getUTCFullYear() + Math.floor(a / 12),
    d = ((a % 12) + 12) % 12;
  (!Number.isSafeInteger(a) || s < 1899 || s > 9999) &&
    throwDateSeriesRangeError();
  const u = utcDateFromParts(s, d + 1, 0).getUTCDate(),
    l = n === "endOfMonth" ? u : Math.min(o.getUTCDate(), u);
  return dateToExcelSerial(utcDateFromParts(s, d, l));
}
function Nae(e, r, t, n, o) {
  const a =
    stepSeriesDateSerial(e + dateSystemSerialOffset(o), r, t, n) -
    dateSystemSerialOffset(o);
  return (
    (a < 0 || a > dateSystemMaxSerial(o)) && throwDateSeriesRangeError(),
    a
  );
}
function buildSeriesPreview(workbook, sheetId, range, n) {
  const direction = validateFillDirection(n.direction),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(range),
    d = direction === "down" || direction === "up";
  (["number", "date"].includes(n.kind) ||
    throwWorkbookError("Choose a numeric or date series."),
    ((d && rangeAddress.firstCol !== rangeAddress.lastCol) ||
      (!d && rangeAddress.firstRow !== rangeAddress.lastRow)) &&
      throwWorkbookError(
        "Select one column for vertical series or one row for horizontal series.",
      ),
    (rangeAddress.lastRow >= sheetOrThrow.rowCount ||
      rangeAddress.lastCol >= sheetOrThrow.colCount) &&
      throwWorkbookError("The series range exceeds the worksheet."));
  const targetCount = d
    ? rangeAddress.lastRow - rangeAddress.firstRow + 1
    : rangeAddress.lastCol - rangeAddress.firstCol + 1;
  targetCount > workbookLimits.maxPatches &&
    throwWorkbookError("A series may contain at most 500 target cells.");
  const l = direction === "up" ? rangeAddress.lastRow : rangeAddress.firstRow,
    c = direction === "left" ? rangeAddress.lastCol : rangeAddress.firstCol,
    cellAddress = formatCellAddress(l, c),
    f = stripCellDecorations(sheetOrThrow.cells[cellAddress]),
    p = n.step ?? 1;
  (typeof p != "number" ||
    !Number.isFinite(p) ||
    p === 0 ||
    (n.kind === "date" && !Number.isSafeInteger(p))) &&
    throwWorkbookError(
      "Series step must be a nonzero finite number; date steps must be whole numbers.",
    );
  const m = n.dateUnit ?? "day",
    v = n.monthPolicy ?? "preserveDay";
  (!["day", "weekday", "month", "year"].includes(m) ||
    !["preserveDay", "endOfMonth"].includes(v)) &&
    throwWorkbookError("Choose a supported date unit and month policy.");
  const g =
    n.start ?? calculateWorkbookWithSettings(workbook)(sheetId, cellAddress);
  typeof g != "number" &&
    typeof g != "string" &&
    throwWorkbookError(
      "The directional seed must contain a number or date, or supply a start value.",
    );
  const b =
      n.kind === "date"
        ? parseSeriesSeedDate(g, workbook.datePolicy?.dateSystem)
        : typeof g == "number" && Number.isFinite(g)
          ? g
          : throwWorkbookError(
              "Numeric series require a finite numeric start value.",
            ),
    y =
      n.stop === undefined
        ? undefined
        : n.kind === "date"
          ? parseSeriesSeedDate(n.stop, workbook.datePolicy?.dateSystem)
          : typeof n.stop == "number" && Number.isFinite(n.stop)
            ? n.stop
            : throwWorkbookError(
                "Numeric series require a finite numeric stop value.",
              );
  y !== undefined &&
    ((p > 0 && y < b) || (p < 0 && y > b)) &&
    throwWorkbookError("The stop value is opposite the series step direction.");
  const w = [],
    E = [];
  let generatedCount = 0;
  for (let O = 0; O < targetCount; O++) {
    let raw;
    try {
      raw =
        n.kind === "date"
          ? Nae(b, O * p, m, v, workbook.datePolicy?.dateSystem ?? "1900")
          : b + O * p;
    } catch (error) {
      if (y !== undefined && O > 0 && error instanceof DateSeriesRangeError)
        break;
      throw error;
    }
    if (y !== undefined && (p > 0 ? raw > y : raw < y)) break;
    Number.isFinite(raw) ||
      throwWorkbookError("The numeric series would overflow.");
    const z = formatCellAddress(
      l + (d ? O * (direction === "up" ? -1 : 1) : 0),
      c + (d ? 0 : O * (direction === "left" ? -1 : 1)),
    );
    (E.push(z), generatedCount++);
    const cellRecordInput = normalizeCellRecordInput(
      mergeCellDecorations(
        {
          ...f,
          raw: raw,
          literal: undefined,
          ...(n.kind === "date"
            ? { format: "date", numFmt: "yyyy-mm-dd", numberFormat: undefined }
            : {}),
        },
        sheetOrThrow.cells[z],
      ),
    );
    cellContentsDiffer(sheetOrThrow.cells[z], cellRecordInput) &&
      w.push({ sheetId: sheetId, address: z, cell: cellRecordInput });
  }
  return (
    assertFillTargetUnmerged(workbook, sheetId, E),
    buildCellChangePreview(
      workbook,
      w,
      {
        targetSheetId: sheetId,
        targetRange: formatRangeBoundsAddress(range),
        targetCount: targetCount,
        generatedCount: generatedCount,
        direction: direction,
        summary: `Fill ${n.kind === "date" ? `${m} date` : "numeric"} series ${direction} in ${formatRangeBoundsAddress(range)} (${generatedCount} values).`,
      },
      n.readOnly,
    )
  );
}

function buildFillHandlePreview(
  workbook,
  sheetId,
  sourceRange,
  targetRange,
  o = {},
) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(sourceRange),
    d = parseRangeAddress(targetRange),
    u = o.mode ?? "copy";
  (["copy", "series"].includes(u) ||
    throwWorkbookError("Choose Copy or Series for the fill handle."),
    (rangeAddress.firstRow < d.firstRow ||
      rangeAddress.lastRow > d.lastRow ||
      rangeAddress.firstCol < d.firstCol ||
      rangeAddress.lastCol > d.lastCol) &&
      throwWorkbookError(
        "The fill-handle target must contain the complete source range.",
      ),
    (d.lastRow >= sheetOrThrow.rowCount ||
      d.lastCol >= sheetOrThrow.colCount) &&
      throwWorkbookError("The fill-handle target exceeds the worksheet."));
  const l = rangeAddress.lastRow - rangeAddress.firstRow + 1,
    c = rangeAddress.lastCol - rangeAddress.firstCol + 1,
    h =
      rangeAddress.firstCol === d.firstCol &&
      rangeAddress.lastCol === d.lastCol,
    f =
      rangeAddress.firstRow === d.firstRow &&
      rangeAddress.lastRow === d.lastRow;
  !h &&
    !f &&
    throwWorkbookError(
      "Extend the fill handle along one axis while keeping the other source dimension unchanged.",
    );
  const p = (d.lastRow - d.firstRow + 1) * (d.lastCol - d.firstCol + 1) - l * c;
  p > workbookLimits.maxPatches &&
    throwWorkbookError("The fill handle may extend into at most 500 cells.");
  const direction = h
      ? d.firstRow < rangeAddress.firstRow
        ? "up"
        : "down"
      : d.firstCol < rangeAddress.firstCol
        ? "left"
        : "right",
    v = [],
    g = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  let y = 0,
    w = 0;
  if (u === "series") {
    (h && l === 2 && c === 1) ||
      (f && l === 1 && c === 2) ||
      throwWorkbookError(
        "Series handle fill requires exactly two seed cells in the extension's row or column.",
      );
    const k = workbookWithSettings(
        sheetId,
        formatCellAddress(rangeAddress.firstRow, rangeAddress.firstCol),
      ),
      F = workbookWithSettings(
        sheetId,
        formatCellAddress(rangeAddress.lastRow, rangeAddress.lastCol),
      );
    ((typeof k != "number" ||
      !Number.isFinite(k) ||
      typeof F != "number" ||
      !Number.isFinite(F) ||
      !Number.isFinite(F - k)) &&
      throwWorkbookError(
        "Series handle seeds must evaluate to finite numbers.",
      ),
      (y = k),
      (w = F - k));
  }
  const E = (k, F) => ((k % F) + F) % F;
  function* A() {
    if (h)
      for (const [k, F] of [
        [d.firstRow, rangeAddress.firstRow - 1],
        [rangeAddress.lastRow + 1, d.lastRow],
      ])
        for (let row = k; row <= F; row++)
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
        for (const [F, I] of [
          [d.firstCol, rangeAddress.firstCol - 1],
          [rangeAddress.lastCol + 1, d.lastCol],
        ])
          for (let col = F; col <= I; col++) yield { row: firstRow, col: col };
  }
  if (p)
    for (const { row, col } of A()) {
      const I =
          u === "series"
            ? rangeAddress.firstRow
            : rangeAddress.firstRow + E(row - rangeAddress.firstRow, l),
        O =
          u === "series"
            ? rangeAddress.firstCol
            : rangeAddress.firstCol + E(col - rangeAddress.firstCol, c),
        fillCellPatches = buildFillCellPatches(
          sheetOrThrow,
          formatCellAddress(I, O),
          workbookWithSettings,
          rangeAddress,
        ),
        cellAddress = formatCellAddress(row, col);
      if (u === "series") {
        const D =
          y +
          (h ? row - rangeAddress.firstRow : col - rangeAddress.firstCol) * w;
        (Number.isFinite(D) ||
          throwWorkbookError("The handle series would overflow."),
          (fillCellPatches.raw = D),
          delete fillCellPatches.literal);
      } else
        typeof fillCellPatches.raw == "string" &&
          fillCellPatches.raw.startsWith("=") &&
          !fillCellPatches.literal &&
          (fillCellPatches.raw = shiftFormulaReferences(
            fillCellPatches.raw,
            row - I,
            col - O,
          ));
      const cellRecordInput = normalizeCellRecordInput(
        mergeCellDecorations(fillCellPatches, sheetOrThrow.cells[cellAddress]),
      );
      (g.push(cellAddress),
        cellContentsDiffer(sheetOrThrow.cells[cellAddress], cellRecordInput) &&
          v.push({
            sheetId: sheetId,
            address: cellAddress,
            cell: cellRecordInput,
          }));
    }
  return (
    assertFillTargetUnmerged(workbook, sheetId, g),
    buildCellChangePreview(
      workbook,
      v,
      {
        targetSheetId: sheetId,
        targetRange: formatRangeBoundsAddress(targetRange),
        targetCount: p,
        generatedCount: p,
        direction: direction,
        summary: `Fill handle ${u} from ${formatRangeBoundsAddress(sourceRange)} into ${formatRangeBoundsAddress(targetRange)} (${p} extension cells).`,
      },
      o.readOnly,
    )
  );
}
function parseDateCellEntry(e, r, t = "1900") {
  if (typeof e != "string") throw new Error("Cell entry must be text.");
  const cell = r ? structuredClone(r) : { raw: null };
  if (((cell.literal = e.startsWith("'")), cell.literal)) cell.raw = e.slice(1);
  else if (/^\d{4}-\d{2}-\d{2}$/.test(e)) {
    try {
      cell.raw = parseSeriesSeedDate(e, t);
    } catch {
      throw new Error(
        `Enter a valid YYYY-MM-DD date between ${t === "1904" ? "1904-01-01" : "1899-12-31"} and 9999-12-31, or begin with an apostrophe to keep the entry as text.${t === "1900" ? " Excel's compatibility date 1900-02-29 represents serial 60." : ""}`,
      );
    }
    ((cell.format = "date"),
      (cell.numFmt = "yyyy-mm-dd"),
      (cell.numberFormat = {
        category: "date",
        locale: "en-US",
        pattern: "yyyy-mm-dd",
      }));
  } else cell.raw = parseRawCellEntry(e);
  return normalizeCellRecordInput(cell);
}
function parseRawCellEntry(e) {
  return e === ""
    ? null
    : /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(e) &&
        e.replace(/[-.]/g, "").length <= 15
      ? Number(e)
      : /^(TRUE|FALSE)$/i.test(e)
        ? e.toUpperCase() === "TRUE"
        : e;
}
const supportedWorkbookLocales = ["en-US", "en-GB", "de-DE", "fr-FR"];
function resolveWorkbookLocale(locale) {
  if (locale === undefined) return "en-US";
  if (!supportedWorkbookLocales.includes(locale))
    throw new Error("Choose a supported workbook interface locale.");
  return locale;
}
function formatNumberForLocale(e, entryLocale, useGrouping = true) {
  return new Intl.NumberFormat(entryLocale, {
    useGrouping: useGrouping,
    maximumSignificantDigits: 15,
  }).format(e);
}
function parseLocalizedCellEntry(e, cell, t = "1900", locale) {
  if (
    (resolveWorkbookLocale(locale),
    typeof cell?.raw == "number" &&
      e === localizeDecimalPoint(cell.raw, locale))
  )
    return normalizeCellRecordInput(structuredClone(cell));
  if (!locale || /^[=']/.test(e) || /^\d{4}-\d{2}-\d{2}$/.test(e))
    return parseDateCellEntry(e, cell, t);
  const o = locale === "de-DE" || locale === "fr-FR" ? "," : ".",
    a =
      locale === "de-DE" ? "." : locale === "fr-FR" ? "[ \\u00a0\\u202f]" : ",",
    s = (l) => l.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
    d = `(?:[1-9]\\d{0,2}(?:${locale === "fr-FR" ? a : s(a)}\\d{3})+)`;
  if (new RegExp(`^-?(?:0|[1-9]\\d*|${d})(?:${s(o)}\\d+)?$`).test(e)) {
    const l = e
      .replace(
        locale === "fr-FR" ? /[ \u00a0\u202f]/g : new RegExp(s(a), "g"),
        "",
      )
      .replace(o, ".");
    return parseDateCellEntry(l, cell, t);
  }
  return locale === "de-DE" && /^(WAHR|FALSCH)$/i.test(e)
    ? parseDateCellEntry(e.toUpperCase() === "WAHR" ? "TRUE" : "FALSE", cell, t)
    : locale === "fr-FR" && /^(VRAI|FAUX)$/i.test(e)
      ? parseDateCellEntry(
          e.toUpperCase() === "VRAI" ? "TRUE" : "FALSE",
          cell,
          t,
        )
      : /^-?[\d., \u00a0\u202f]+$/.test(e) && /[.,]/.test(e)
        ? parseDateCellEntry("'" + e, cell, t)
        : parseDateCellEntry(e, cell, t);
}
function formatCellEntryForEdit(cell, entryLocale) {
  return typeof cell?.raw == "number" && entryLocale
    ? localizeDecimalPoint(cell.raw, entryLocale)
    : `${cell?.literal && typeof cell.raw == "string" ? "'" : ""}${String(cell?.raw ?? "")}`;
}
function localizeDecimalPoint(e, entryLocale) {
  return Number.isFinite(e)
    ? entryLocale === "de-DE" || entryLocale === "fr-FR"
      ? String(e).replace(".", ",")
      : String(e)
    : "";
}
function validateLocalizedNumberEntry(e, r, locale) {
  let n = e.trim();
  {
    resolveWorkbookLocale(locale);
    const o = locale === "de-DE" || locale === "fr-FR" ? "," : ".",
      a =
        locale === "de-DE"
          ? "\\."
          : locale === "fr-FR"
            ? "[ \\u00a0\\u202f]"
            : ",";
    if (
      !new RegExp(
        `^[+-]?(?:(?:\\d+|[1-9]\\d{0,2}(?:${a}\\d{3})+)(?:${o === "." ? "\\." : ","}\\d*)?|${o === "." ? "\\." : ","}\\d+)(?:[eE][+-]?\\d+)?$`,
      ).test(n)
    )
      throw new Error(
        `${r} must be a finite number in the selected numeric locale.`,
      );
    n = n.replace(new RegExp(a, "g"), "").replace(o, ".");
  }
  if (!n || !Number.isFinite(Number(n)))
    throw new Error(`${r} must be a finite number.`);
  return Number(n);
}
function formatLocalizedNumberDisplay(e, cell, t = "1900", entryLocale) {
  return !entryLocale ||
    typeof e != "number" ||
    cell?.numberFormat ||
    cell?.numFmt ||
    cell?.format === "date"
    ? formatNumberWithPattern(e, cell, t)
    : Number.isFinite(e)
      ? cell?.format === "currency"
        ? new Intl.NumberFormat(entryLocale, {
            style: "currency",
            currency: "USD",
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          }).format(e)
        : cell?.format === "percent"
          ? new Intl.NumberFormat(entryLocale, {
              style: "percent",
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            }).format(e)
          : cell?.format === "number"
            ? new Intl.NumberFormat(entryLocale, {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2,
              }).format(e)
            : formatNumberForLocale(e, entryLocale, false)
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
    t = {
      ...(workbookLocale === "de-DE" ? germanWorkbookStatusMessages : {}),
      ...options.messages,
    },
    n = (d) => (Object.hasOwn(t, d) ? d : d.trim()),
    o = (d) => [...new Set(d.match(/\{\w+\}/g) ?? [])].sort().join("|"),
    a = (d) =>
      Object.hasOwn(t, n(d)) &&
      typeof t[n(d)] == "string" &&
      t[n(d)].trim().length > 0 &&
      o(d) === o(t[n(d)]),
    s = (d) =>
      a(d)
        ? n(d) === d
          ? t[d]
          : `${d.match(/^\s*/)?.[0] ?? ""}${t[n(d)]}${d.match(/\s*$/)?.[0] ?? ""}`
        : d;
  return {
    locale: workbookLocale,
    entryLocale: options.locale,
    text: s,
    language: (d) => (a(d) ? workbookLocale : "en"),
    message: (d, u) =>
      s(d).replace(/\{(\w+)\}/g, (l, c) =>
        Object.hasOwn(u, c) ? String(u[c]) : l,
      ),
  };
}
const WorkbookMessagesContext = ne.createContext(resolveWorkbookMessages());
function WorkbookMessagesProvider({ options, children }) {
  const t = ne.useMemo(
    () => resolveWorkbookMessages(options),
    [options?.locale, options?.messages],
  );
  return i.jsx(WorkbookMessagesContext.Provider, { value: t, children });
}
const useWorkbookMessages = () => ne.useContext(WorkbookMessagesContext);
function LocalizedText({ children }) {
  const workbookMessages = useWorkbookMessages(),
    t = ne.Children.toArray(children);
  if (!t.every((o) => typeof o == "string"))
    return i.jsx(i.Fragment, { children });
  const n = t.join("");
  return workbookMessages.locale.startsWith("en") &&
    workbookMessages.text(n) === n
    ? i.jsx(i.Fragment, { children })
    : i.jsx("span", {
        lang: workbookMessages.language(n),
        children: workbookMessages.text(n),
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
  ...c
}) {
  const workbookMessages = useWorkbookMessages(),
    f = ne.useId(),
    [p, m] = ne.useState(() =>
      value === "" || value === undefined
        ? ""
        : localizeDecimalPoint(Number(value), workbookMessages.entryLocale),
    ),
    [v, g] = ne.useState(false),
    [b, y] = ne.useState(false),
    w = ne.useRef(workbookMessages.entryLocale),
    E = ne.useRef(undefined),
    A = min === undefined ? undefined : Number(min),
    k = max === undefined ? undefined : Number(max),
    F = (V) => {
      if (!V.trim() && allowEmpty) return;
      const D = validateLocalizedNumberEntry(V, "Value", w.current);
      if ((A !== undefined && D < A) || (k !== undefined && D > k))
        throw Error("bounds");
      return D;
    };
  let I,
    O = "";
  try {
    I = F(p);
  } catch (error) {
    O =
      error instanceof Error && error.message === "bounds"
        ? workbookMessages.message("Enter a number from {min} to {max}.", {
            min:
              A === undefined
                ? "−∞"
                : localizeDecimalPoint(A, workbookMessages.entryLocale),
            max:
              k === undefined
                ? "∞"
                : localizeDecimalPoint(k, workbookMessages.entryLocale),
          })
        : workbookMessages.text(
            "Enter a valid number using the selected locale.",
          );
  }
  (ne.useLayoutEffect(() => {
    const V = value === "" || value === undefined ? "" : String(Number(value));
    (E.current !== undefined &&
      (V === E.current || (V === "" && E.current === "NaN"))) ||
      ((w.current = workbookMessages.entryLocale),
      m(
        value === "" || value === undefined
          ? ""
          : localizeDecimalPoint(Number(value), workbookMessages.entryLocale),
      ),
      y(false));
  }, [value]),
    ne.useLayoutEffect(() => {
      v ||
        O ||
        ((w.current = workbookMessages.entryLocale),
        m(
          I === undefined
            ? ""
            : localizeDecimalPoint(I, workbookMessages.entryLocale),
        ));
    }, [workbookMessages.entryLocale, v]));
  const j = (V) => {
      m(V);
      let D;
      try {
        const B = F(V);
        D = B === undefined ? "" : String(B);
      } catch {
        D = "NaN";
      }
      ((E.current = D), onValueChange(D));
    },
    z =
      [
        c["aria-describedby"],
        O ? `${f}-error` : null,
        v && w.current !== workbookMessages.entryLocale ? `${f}-locale` : null,
      ]
        .filter(Boolean)
        .join(" ") || undefined;
  return i.jsxs("span", {
    className: "lg-number-field",
    children: [
      i.jsx("input", {
        ...c,
        className,
        type: "text",
        inputMode: "decimal",
        role: "spinbutton",
        value: p,
        "aria-valuenow": I,
        "aria-valuemin": A,
        "aria-valuemax": k,
        "aria-valuetext": p || undefined,
        "aria-invalid": !!O || undefined,
        "aria-describedby": z,
        onChange: (event) => j(event.target.value),
        onFocus: (V) => {
          (O || (w.current = workbookMessages.entryLocale),
            g(true),
            onFocus?.(V));
        },
        onBlur: (V) => {
          (y(true), g(false), onBlur?.(V));
        },
        onKeyDown: (V) => {
          if (
            (onKeyDown?.(V),
            !(
              V.defaultPrevented ||
              c.readOnly ||
              c.disabled ||
              V.nativeEvent.isComposing ||
              V.altKey ||
              V.metaKey ||
              V.ctrlKey
            ) &&
              (V.key === "ArrowUp" || V.key === "ArrowDown"))
          ) {
            V.preventDefault();
            const D = step === undefined || step === "any" ? 1 : Number(step);
            if (!Number.isFinite(D) || D <= 0) return;
            const B = Math.min(
              k ?? 1 / 0,
              Math.max(
                A ?? -1 / 0,
                (I ?? A ?? 0) + (V.key === "ArrowUp" ? D : -D),
              ),
            );
            Number.isFinite(B) &&
              j(localizeDecimalPoint(Number(B.toPrecision(15)), w.current));
          }
        },
      }),
      O &&
        i.jsx("span", {
          id: `${f}-error`,
          className: b ? "lg-number-error" : "lg-sr-only",
          role: b ? "status" : undefined,
          children: O,
        }),
      v &&
        w.current !== workbookMessages.entryLocale &&
        i.jsx("span", {
          id: `${f}-locale`,
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
function findOpenFormulaQuote(e, index) {
  let t = "";
  for (let n = 1; n < index; n++) {
    const o = e[n];
    if (t && o === t && e[n + 1] === t) {
      n++;
      continue;
    }
    o === t ? (t = "") : !t && (o === '"' || o === "'") && (t = o);
  }
  return t;
}
const formulaReferencePattern =
  /(?:'(?:[^']|'')+'!|[A-Za-z_][A-Za-z0-9_.]*!)?(?:\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)/g;
function findFormulaReferenceSpan(e, start, end = start) {
  if (!(!e.startsWith("=") || start < 1 || end < start || end > e.length)) {
    for (const n of e.matchAll(formulaReferencePattern)) {
      const index = n.index,
        a = index + n[0].length;
      if (
        !(
          findOpenFormulaQuote(e, index) ||
          /[\w.\[!]/.test(e[index - 1] ?? "") ||
          /[\w.\](]/.test(e[a] ?? "")
        ) &&
        start >= index &&
        end <= a
      )
        return { start: index, end: a };
    }
    if (!findOpenFormulaQuote(e, start)) {
      if (end > start) return { start: start, end: end };
      if (/[=(:,;+\-*/^&<>]$/.test(e.slice(0, start).trimEnd()))
        return { start: start, end: end };
    }
  }
}
function replaceFormulaReferenceSpan(e, r, range, n) {
  const o = n === undefined ? range : `'${n.replaceAll("'", "''")}'!${range}`;
  return {
    value: e.slice(0, r.start) + o + e.slice(r.end),
    span: { start: r.start, end: r.start + o.length },
  };
}
function $ae(e, start, end = start) {
  const formulaReferenceSpan = findFormulaReferenceSpan(e, start, end);
  if (
    !formulaReferenceSpan ||
    formulaReferenceSpan.start === formulaReferenceSpan.end
  )
    return;
  const o = e.slice(formulaReferenceSpan.start, formulaReferenceSpan.end),
    a = o.lastIndexOf("!"),
    s = o.slice(0, a + 1),
    d = o.slice(a + 1),
    u = /^(\$?)[A-Za-z]+(\$?)\d+/.exec(d);
  if (!u) {
    if (!/^(?:\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)$/.test(d))
      return;
    const f = d.startsWith("$"),
      p =
        s +
        d
          .split(":")
          .map((m) => (f ? m.replace(/^\$/, "") : "$" + m.replace(/^\$/, "")))
          .join(":");
    return {
      value:
        e.slice(0, formulaReferenceSpan.start) +
        p +
        e.slice(formulaReferenceSpan.end),
      span: {
        start: formulaReferenceSpan.start,
        end: formulaReferenceSpan.start + p.length,
      },
    };
  }
  const l = !u[1] && !u[2] ? 1 : u[1] && u[2] ? 2 : !u[1] && u[2] ? 3 : 0,
    c = d.replace(
      /\$?([A-Za-z]+)\$?(\d+)/g,
      (f, p, m) =>
        `${l === 1 || l === 3 ? "$" : ""}${p}${l === 1 || l === 2 ? "$" : ""}${m}`,
    ),
    h = s + c;
  return {
    value:
      e.slice(0, formulaReferenceSpan.start) +
      h +
      e.slice(formulaReferenceSpan.end),
    span: {
      start: formulaReferenceSpan.start,
      end: formulaReferenceSpan.start + h.length,
    },
  };
}
function autoCloseFormulaParens(e) {
  if (!e.startsWith("=")) return e;
  let r = "",
    t = 0,
    n = 0;
  for (let o = 1; o < e.length; o++) {
    const a = e[o];
    if (r) {
      a === r && (e[o + 1] === r ? o++ : (r = ""));
      continue;
    }
    if (
      (a === '"' || a === "'"
        ? (r = a)
        : a === "["
          ? n++
          : a === "]"
            ? n--
            : !n && a === "("
              ? t++
              : !n && a === ")" && t--,
      t < 0 || n < 0)
    )
      return e;
  }
  return !r && !n && t > 0 && t <= 64 && /[\w\])"%]$/.test(e.trimEnd())
    ? e + ")".repeat(t)
    : e;
}
const formulaScanCache = new WeakMap();
function getFormulaScanList(cells, r) {
  const t = isDeepFrozenValue(cells);
  let n = t ? formulaScanCache.get(cells) : undefined;
  if (!n) {
    n = { calls: [], external: [] };
    for (const o in cells) {
      const cell = cells[o];
      typeof cell.raw != "string" ||
        !cell.raw.startsWith("=") ||
        (cell.raw.includes("(") && n.calls.push([o, cell]),
        /[\[|\\]|https?:|file:/i.test(cell.raw) && n.external.push([o, cell]));
    }
    t && formulaScanCache.set(cells, n);
  }
  return n[r];
}
const objectIdRegistry = new WeakMap();
let objectIdCounter = 0;
function getStableObjectId(invoke) {
  let r = objectIdRegistry.get(invoke);
  return (r || ((r = ++objectIdCounter), objectIdRegistry.set(invoke, r)), r);
}
function validateExtensionHostPolicy(workbook, r) {
  const extensions = r.extensions ?? [],
    allowed = r.allowedSheetIds ?? workbook.sheets.map((sheet) => sheet.id);
  if (r.readOnly !== undefined && typeof r.readOnly != "boolean")
    throw Error("Read-only policy must be boolean.");
  if (
    !Array.isArray(extensions) ||
    extensions.length > extensionRuntimeLimits.extensions ||
    !Array.isArray(allowed) ||
    allowed.some((a) => !workbook.sheets.some((sheet) => sheet.id === a))
  )
    throw Error("Invalid extension host scope.");
  const o = new Set();
  for (const extension of extensions) {
    const s = validateExtensionManifest(extension.manifest);
    if (
      o.has(s.id) ||
      typeof extension.invoke != "function" ||
      !Array.isArray(extension.allowedSheetIds) ||
      extension.allowedSheetIds.some(
        (allowedSheetId) =>
          !workbook.sheets.some((sheet) => sheet.id === allowedSheetId),
      ) ||
      !Array.isArray(extension.grantedPermissions) ||
      extension.grantedPermissions.some(
        (grantedPermission) => !s.permissions.includes(grantedPermission),
      ) ||
      (extension.includeHidden !== undefined &&
        typeof extension.includeHidden != "boolean")
    )
      throw Error("Invalid host extension registration.");
    o.add(s.id);
  }
  return {
    extensions: extensions,
    allowed: allowed,
    hash: () =>
      fingerprintCanonicalValue([
        !!r.readOnly,
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
const resolveExtensionSheetScope = (workbook, r) =>
  fingerprintCanonicalValue([
    r.grantedPermissions,
    [...r.allowedSheetIds].sort().map((t) => {
      const n = workbook.sheets.find((sheet) => sheet.id === t);
      return [t, !!n.hidden, !!n.veryHidden];
    }),
    !!r.includeHidden,
  ]);
function buildExtensionHostScope(workbook, r) {
  return validateExtensionHostPolicy(workbook, r)
    .extensions.filter((extension) =>
      extension.manifest.permissions.every((permission) =>
        extension.grantedPermissions.includes(permission),
      ),
    )
    .map((t) => ({
      id: t.manifest.id,
      manifestHash: hashExtensionManifest(t.manifest),
      scopeHash: resolveExtensionSheetScope(workbook, t),
      sheetIds: t.allowedSheetIds.filter((allowedSheetId) => {
        const o = workbook.sheets.find((sheet) => sheet.id === allowedSheetId);
        return !(o.hidden || o.veryHidden) || !!t.includeHidden;
      }),
    }));
}
function withExtensionHostScope(workbook, extensions = []) {
  const t = { ...workbook };
  return (
    setExtensionCacheEntries(
      t,
      buildExtensionHostScope(workbook, { extensions: extensions }),
    ),
    isDeepFrozenValue(workbook) ? deepFreezeValue(t) : t
  );
}
function assertExtensionWriteScope(workbook, r, t) {
  if (t.readOnly) throw Error("The workbook is read-only.");
  if (workbook.structureProtection?.enabled)
    throw Error("Workbook structure is protected.");
  const n = validateExtensionHostPolicy(workbook, t);
  for (const a of r) {
    if (!n.allowed.includes(a))
      throw Error("An extension destination is outside the host write scope.");
    assertSheetUnprotected(workbook.sheets.find((sheet) => sheet.id === a));
  }
}
function reviewExtensionFormulas(workbook) {
  const r = new Set(builtinFunctionNames),
    t = new Set((workbook.definedNames ?? []).map((o) => o.name.toUpperCase())),
    n = [];
  for (const sheet of workbook.sheets)
    for (const [address, cell] of getFormulaScanList(sheet.cells, "calls")) {
      if (typeof cell.raw != "string" || !cell.raw.startsWith("=")) continue;
      const calls = [],
        u = (l, nested = false) => {
          for (const h of l) {
            if ((h.kind === "group" && u(h.body, nested), h.kind !== "call"))
              continue;
            const callee = h.callee;
            let p = false;
            if (callee.kind === "symbol" && !callee.qualifier) {
              const m = callee.name.replace(/^_xlfn\./i, "").toUpperCase();
              ((p = !r.has(m) && !t.has(m) && !["LET", "LAMBDA"].includes(m)),
                p &&
                  calls.push({
                    name: m,
                    start: h.start,
                    end: h.end,
                    arguments: h.args.map((arg) =>
                      arg.length
                        ? cell.raw
                            .toString()
                            .slice(arg[0].start, arg.at(-1).end)
                        : "",
                    ),
                    nested: nested,
                  }));
            }
            h.args.forEach((arg) => u(arg, nested || p));
          }
        };
      try {
        u(tokenizeFormulaText(cell.raw));
      } catch {
        continue;
      }
      if (
        (calls.length &&
          n.push({
            sheetId: sheet.id,
            sheetName: sheet.name,
            address: address,
            formula: cell.raw,
            literal: !!cell.literal,
            calls: calls,
          }),
        n.length > extensionRuntimeLimits.calls)
      )
        throw Error("Extension review exceeds 100 formula cells.");
    }
  return n;
}
function requireApprovedExtension(workbook, extensionId, t) {
  const n = workbook.extensions?.installed.find(
    (s) => s.manifest.id === extensionId,
  );
  if (!n || !n.enabled) throw Error("The extension is missing or disabled.");
  const o = validateExtensionHostPolicy(workbook, t).extensions.find(
    (extension) => extension.manifest.id === extensionId,
  );
  if (
    !o ||
    hashExtensionManifest(o.manifest) !== hashExtensionManifest(n.manifest)
  )
    throw Error(
      "An approved adapter with the exact extension version and manifest is unavailable.",
    );
  if (
    !n.manifest.permissions.every((permission) =>
      o.grantedPermissions.includes(permission),
    )
  )
    throw Error("The host has not granted all declared extension permissions.");
  return o;
}
function evaluateExtensionFunctionCall(workbook, r, call, binding) {
  const o = binding.manifest.functions.find(
    (s) => `${binding.manifest.namespace}.${s.id}` === call.name,
  );
  if (!o || o.parameters.length !== call.arguments.length)
    throw Error(
      "Custom function arguments do not match the declared signature.",
    );
  if (call.nested || /\b(?:LET|LAMBDA)\s*\(/i.test(r.formula))
    throw Error(
      "Nested custom calls and custom calls inside LET/LAMBDA require separate worksheet cells.",
    );
  const allowRead = (s) => {
    const d = workbook.sheets.find((sheet) => sheet.id === s);
    return (
      !!d &&
      binding.allowedSheetIds.includes(s) &&
      (!(d.hidden || d.veryHidden) || !!binding.includeHidden)
    );
  };
  if (!allowRead(r.sheetId))
    throw Error(
      "The extension formula is outside its approved worksheet scope.",
    );
  return call.arguments.map((argument, d) => {
    if (!argument.trim()) throw Error("Extension arguments cannot be omitted.");
    const schema = o.parameters[d].schema,
      l =
        /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?$/.exec(
          argument.trim(),
        );
    if (l) {
      const g = l[1]?.replace(/''/g, "'") ?? l[2],
        b = g
          ? workbook.sheets.find(
              (sheet) => sheet.name.toLowerCase() === g.toLowerCase(),
            )
          : workbook.sheets.find((sheet) => sheet.id === r.sheetId);
      if (!b || !allowRead(b.id))
        throw Error(
          "An extension input worksheet is unavailable or unauthorized.",
        );
      const rangeAddress = parseRangeAddress(l[3] + (l[4] ? ":" + l[4] : "")),
        w = rangeAddress.lastRow - rangeAddress.firstRow + 1,
        E = rangeAddress.lastCol - rangeAddress.firstCol + 1;
      if (
        w * E > extensionRuntimeLimits.inputCells ||
        (schema.type !== "matrix" && w * E !== 1)
      )
        throw Error(
          "The extension input range exceeds its declared scalar/cell budget.",
        );
      const A = evaluateWorkbookFormulas(
          workbook,
          {},
          { allowRead: allowRead, allowName: () => true, step: () => {} },
        ),
        k = Array.from({ length: w }, (F, I) =>
          Array.from({ length: E }, (O, j) => {
            const z = A.getTypedValue(
              b.id,
              `${columnIndexToLetters(rangeAddress.firstCol + j)}${rangeAddress.firstRow + I + 1}`,
            );
            if (z.type === "error")
              throw Error(
                "An extension input contains an unavailable calculation.",
              );
            return z.value;
          }),
        );
      if (A.getOperationalErrors().length)
        throw Error("An extension input attempted an unauthorized dependency.");
      return validateExtensionValue(
        schema.type === "matrix" ? k : k[0][0],
        schema,
      );
    }
    if (schema.type === "matrix")
      throw Error("Matrix parameters require a static worksheet range.");
    const c = {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === r.sheetId
          ? {
              ...sheet,
              cells: { ...sheet.cells, [r.address]: { raw: "=" + argument } },
            }
          : sheet,
      ),
    };
    setExtensionCacheEntries(c, getExtensionCacheEntries(workbook));
    const h = evaluateWorkbookFormulas(
        c,
        {},
        { allowRead: allowRead, allowName: () => true, step: () => {} },
      ),
      f = h.getTypedValue(r.sheetId, r.address);
    if (f.type === "error" || h.getOperationalErrors().length)
      throw Error(
        "An extension input is erroneous, stale, unauthorized or over its calculation limit.",
      );
    let value = f.value;
    if (h.getSpillRange(r.sheetId, r.address))
      throw Error("Scalar extension arguments cannot contain arrays.");
    return validateExtensionValue(value, schema);
  });
}
function resolveExtensionCachedResult(workbook, r = {}) {
  const t = validateExtensionHostPolicy(workbook, r),
    n = reviewExtensionFormulas(workbook),
    o = n.length ? withExtensionHostScope(workbook, t.extensions) : workbook,
    a = n.length ? evaluateWorkbookFormulas(o) : undefined,
    cells = n.map((u) => {
      const l = a.getTypedValue(u.sheetId, u.address),
        caches = (workbook.extensions?.caches ?? []).filter(
          (v) => v.sheetId === u.sheetId && v.address === u.address,
        ),
        functions = u.calls.map((call) => {
          const g = listInstalledExtensionFunctions(workbook).find(
            (y) => y.name === call.name,
          );
          let b = "unavailable";
          if (
            g &&
            ((b = g.entry.enabled
              ? buildExtensionHostScope(workbook, r).some(
                  (y) =>
                    y.id === g.entry.manifest.id &&
                    y.manifestHash ===
                      hashExtensionManifest(g.entry.manifest) &&
                    y.sheetIds.includes(u.sheetId),
                )
                ? "needs-refresh"
                : "unavailable"
              : "disabled"),
            b === "needs-refresh")
          )
            try {
              const binding = requireApprovedExtension(
                  workbook,
                  g.entry.manifest.id,
                  r,
                ),
                w = evaluateExtensionFunctionCall(o, u, call, binding),
                E = fingerprintExtensionCall(
                  u.sheetId,
                  u.address,
                  u.formula,
                  call.name,
                  w,
                  hashExtensionManifest(g.entry.manifest),
                );
              caches.some(
                (cach) =>
                  cach.key === E &&
                  cach.scopeHash ===
                    resolveExtensionSheetScope(workbook, binding),
              ) && (b = "reviewed-cache");
            } catch {}
          return { name: call.name, state: b };
        });
      return {
        ...u,
        importedCache: workbook.importedCalculationCaches?.entries.find(
          (entry) =>
            entry.sheetId === u.sheetId &&
            entry.address === u.address &&
            entry.formula === u.formula,
        )?.cache,
        functions: functions,
        result: l,
        caches: caches,
        cacheLabel: caches.length
          ? l.type === "error" ||
            functions.some((v) => v.state !== "reviewed-cache")
            ? "Historical cached result — stale or unavailable"
            : "Reviewed cached result — manual refresh"
          : "No retained result",
      };
    });
  return {
    installed: workbook.extensions?.installed ?? [],
    cells: cells,
    available: t.extensions.map((extension) => ({
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
function runExtensionOperation(workbook, extensionId, action, n = {}) {
  const o = validateExtensionHostPolicy(workbook, n),
    a = workbook.extensions?.installed.find(
      (m) => m.manifest.id === extensionId,
    ),
    s = o.extensions.find((extension) => extension.manifest.id === extensionId);
  if (!["install", "enable", "disable", "remove"].includes(action))
    throw Error("Unknown extension operation.");
  const d = action === "install" ? s?.manifest : a?.manifest;
  if (!d) throw Error("The requested extension is unavailable.");
  if (
    (validateExtensionManifest(d),
    ["install", "enable"].includes(action) &&
      (!s ||
        hashExtensionManifest(s.manifest) !== hashExtensionManifest(d) ||
        !d.permissions.every((permission) =>
          s.grantedPermissions.includes(permission),
        )))
  )
    throw Error(
      "The host has not approved this exact manifest and all of its permissions.",
    );
  if (
    d.functions.some(
      (m) =>
        builtinFunctionNames.includes(`${d.namespace}.${m.id}`) ||
        (workbook.definedNames ?? []).some(
          (v) => v.name.toUpperCase() === `${d.namespace}.${m.id}`,
        ),
    )
  )
    throw Error(
      "Extension functions cannot shadow built-ins or defined names.",
    );
  const u = reviewExtensionFormulas(workbook).filter((m) =>
    m.calls.some((call) => call.name.startsWith(d.namespace + ".")),
  );
  assertExtensionWriteScope(
    workbook,
    u.map((m) => m.sheetId),
    n,
  );
  const workbookInput = normalizeWorkbookInput(workbook),
    c = structuredClone(
      workbookInput.extensions ?? {
        schemaVersion: 1,
        installed: [],
        caches: [],
      },
    );
  ((c.installed = c.installed.filter((m) => m.manifest.id !== extensionId)),
    action !== "remove" &&
      c.installed.push({
        manifest: validateExtensionManifest(d),
        enabled: action !== "disable",
      }),
    (action === "remove" || action === "install") &&
      (c.caches = c.caches.filter(
        (cach) => !cach.functionName.startsWith(d.namespace + "."),
      )),
    (workbookInput.extensions = c));
  const h = normalizeWorkbookInput(workbookInput),
    f = {
      extensionId: extensionId,
      action: action,
      manifest: structuredClone(d),
      warning:
        "Raw formulas are preserved. Disabling retains labeled historical results; installing/upgrading or removing clears this extension’s caches. Execution requires a separate argument review.",
    };
  return (
    tT.set(f, {
      bookHash: fingerprintCanonicalValue(workbook),
      policyHash: o.hash(),
      planHash: fingerprintCanonicalValue(f),
      candidate: h,
    }),
    f
  );
}
function applyExtensionPreview(workbook, r, t = {}) {
  const n = tT.get(r);
  if (!n) throw Error("Unknown or already applied extension preview.");
  if (
    t.readOnly ||
    n.bookHash !== fingerprintCanonicalValue(workbook) ||
    n.policyHash !== validateExtensionHostPolicy(workbook, t).hash() ||
    n.planHash !== fingerprintCanonicalValue(r)
  )
    throw Error("Workbook, extension or permissions changed. Preview again.");
  return (tT.delete(r), normalizeWorkbookInput(n.candidate));
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
