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
} from "./01-core-schemas-and-hashing.js";
import { workbookLimits } from "./02-financial-functions.js";
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

function throwCellEditError(e) {
  throw new Error(e);
}
const formatCellAddress = (e, r) => `${columnIndexToLetters(r)}${e + 1}`,
  formatRangeAddress = (e, r, t, n) => {
    const o = formatCellAddress(e, r),
      a = formatCellAddress(e + t - 1, r + n - 1);
    return o === a ? o : `${o}:${a}`;
  },
  isPlainObjectValue = (e) =>
    !!e &&
    typeof e == "object" &&
    !Array.isArray(e) &&
    [Object.prototype, null].includes(Object.getPrototypeOf(e));
function eI(e) {
  return (
    e &&
      typeof e == "object" &&
      (Object.values(e).forEach(eI), Object.freeze(e)),
    e
  );
}
function stripCellDecorations(e) {
  const r = normalizeCellRecordInput(e ?? { raw: null });
  return (delete r.note, delete r.tickmark, delete r.hyperlink, r);
}
function withRawCellValue(e, r) {
  return {
    ...r,
    raw: e,
    literal: typeof e == "string" && e.startsWith("=") ? true : undefined,
  };
}
function mergeCellDecorations(e, r) {
  return {
    ...e,
    ...(r?.note !== undefined ? { note: r.note } : {}),
    ...(r?.hyperlink !== undefined
      ? { hyperlink: structuredClone(r.hyperlink) }
      : {}),
    ...(r?.tickmark !== undefined ? { tickmark: r.tickmark } : {}),
  };
}
function cellContentsMatch(e, r) {
  return (
    JSON.stringify(normalizeCellRecordInput(e ?? { raw: null })) ===
    JSON.stringify(normalizeCellRecordInput(r))
  );
}
function assertWorkbookWritable(e) {
  (e !== undefined &&
    typeof e != "boolean" &&
    throwCellEditError("Read-only must be boolean."),
    e && throwCellEditError("This workbook is read-only."));
}
function validateClipboardSnapshot(e) {
  (!isPlainObjectValue(e) ||
    e.version !== 1 ||
    typeof e.sourceWorkbookId != "string" ||
    !e.sourceWorkbookId ||
    e.sourceWorkbookId.length > 100 ||
    typeof e.sourceSheetId != "string" ||
    !e.sourceSheetId ||
    e.sourceSheetId.length > 100 ||
    !Number.isSafeInteger(e.sourceRevision) ||
    e.sourceRevision < 0 ||
    !Number.isSafeInteger(e.rowCount) ||
    !Number.isSafeInteger(e.colCount) ||
    e.rowCount < 1 ||
    e.colCount < 1 ||
    e.rowCount * e.colCount > workbookLimits.maxPatches ||
    typeof e.sourceRange != "string" ||
    !Array.isArray(e.cells) ||
    e.cells.length !== e.rowCount * e.colCount) &&
    throwCellEditError("Invalid or oversized internal clipboard snapshot.");
  const r = parseRangeAddress(e.sourceRange);
  (r.lastRow - r.firstRow + 1 !== e.rowCount ||
    r.lastCol - r.firstCol + 1 !== e.colCount) &&
    throwCellEditError("Clipboard dimensions do not match its source range.");
  const t = e.cells.map((n, o) => {
    (!isPlainObjectValue(n) ||
      n.rowOffset !== Math.floor(o / e.colCount) ||
      n.colOffset !== o % e.colCount) &&
      throwCellEditError(
        "Clipboard cells must form a complete row-major rectangle.",
      );
    const a = stripCellDecorations(normalizeCellRecordInput(n.cell)),
      s = normalizeCellRecordInput({ raw: n.value }).raw;
    return {
      rowOffset: n.rowOffset,
      colOffset: n.colOffset,
      cell: a,
      value: s,
    };
  });
  return {
    version: 1,
    sourceWorkbookId: e.sourceWorkbookId,
    sourceSheetId: e.sourceSheetId,
    sourceRevision: e.sourceRevision,
    sourceRange: formatRangeAddress(
      r.firstRow,
      r.firstCol,
      e.rowCount,
      e.colCount,
    ),
    rowCount: e.rowCount,
    colCount: e.colCount,
    cells: t,
  };
}
function buildClipboardSnapshot(e, r, t) {
  const n = findSheetOrThrow(e, r),
    o = parseRangeAddress(t),
    a = o.lastRow - o.firstRow + 1,
    s = o.lastCol - o.firstCol + 1;
  (a * s > workbookLimits.maxPatches &&
    throwCellEditError("Copy at most 500 cells into the internal clipboard."),
    (o.lastRow >= n.rowCount || o.lastCol >= n.colCount) &&
      throwCellEditError("The copied range exceeds the worksheet."));
  const d = calculateWorkbookWithSettings(e),
    u = [];
  for (let l = 0; l < a; l++)
    for (let c = 0; c < s; c++) {
      const h = formatCellAddress(o.firstRow + l, o.firstCol + c),
        f = normalizeCellRecordInput({ raw: d(r, h) }).raw;
      let p = stripCellDecorations(n.cells[h]);
      if (p.raw === null && f !== null) {
        const m = d.getSpillAnchor(r, h);
        if (m && m !== h) {
          const v = _t(m);
          (v.row < o.firstRow ||
            v.row > o.lastRow ||
            v.col < o.firstCol ||
            v.col > o.lastCol) &&
            (p = withRawCellValue(f, p));
        }
      }
      u.push({ rowOffset: l, colOffset: c, cell: p, value: f });
    }
  return eI({
    version: 1,
    sourceWorkbookId: e.id,
    sourceSheetId: r,
    sourceRevision: e.revision,
    sourceRange: formatRangeAddress(o.firstRow, o.firstCol, a, s),
    rowCount: a,
    colCount: s,
    cells: u,
  });
}
function buildCellChangePreview(e, r, t, n, o = false) {
  (assertWorkbookWritable(n),
    assertSheetUnprotected(findSheetOrThrow(e, t.targetSheetId)),
    r.length > workbookLimits.maxPatches &&
      throwCellEditError(
        "A fill or paste operation may change at most 500 cells.",
      ));
  const a = r.map((c) => ({
    sheetId: c.sheetId,
    address: c.address,
    cell: normalizeCellRecordInput(c.cell),
  }));
  let s;
  try {
    s = applyCellPatchTransaction(e, a, "Preview", t.summary.slice(0, 200));
  } catch (c) {
    if (
      !o ||
      !(c instanceof DataValidationError) ||
      c.issues.some((h) => h.style === "stop")
    )
      throw c;
    s = applyCellPatchTransaction(e, a, "Preview", t.summary.slice(0, 200), {
      validationApproval: assertNoStopValidationIssues(c),
    });
  }
  const d =
      s === e
        ? []
        : s.events.at(-1).changes.map((c) => ({
            sheetId: c.sheetId,
            address: c.address,
            cell: normalizeCellRecordInput(c.after ?? { raw: null }),
          })),
    u = d.length,
    l = Math.max(0, u - a.length);
  return eI({
    ...t,
    workbookId: e.id,
    revision: e.revision,
    patches: a,
    previewPatches: d,
    count: u,
    dependentCount: l,
    summary: t.summary + (l ? ` Includes ${l} dependent cell changes.` : ""),
    snapshotGuard: JSON.stringify(e),
  });
}
function applyCellChangePreview(e, r, t = {}) {
  (assertWorkbookWritable(t.readOnly),
    (!r ||
      r.workbookId !== e.id ||
      r.revision !== e.revision ||
      r.snapshotGuard !== JSON.stringify(e)) &&
      throwCellEditError(
        "The workbook changed after this preview. Create a new preview before applying it.",
      ),
    assertSheetUnprotected(findSheetOrThrow(e, r.targetSheetId)));
  const n = applyCellPatchTransaction(
    e,
    r.patches,
    t.actor ?? "Preparer",
    r.summary.slice(0, 200),
  );
  return n === e ? e : { ...n, revision: e.revision, events: e.events };
}
function buildPastePreview(e, r, t, n, o = {}) {
  const a = validateClipboardSnapshot(n),
    s = findSheetOrThrow(e, r),
    d = parseRangeAddress(t),
    u = parseRangeAddress(a.sourceRange),
    l = o.mode ?? "all",
    c = o.operation ?? "none";
  ((!["all", "values", "formulas", "formats"].includes(l) ||
    !["none", "add", "subtract", "multiply", "divide"].includes(c)) &&
    throwCellEditError(
      "Choose a supported paste mode and arithmetic operation.",
    ),
    [o.transpose, o.skipBlanks].some(
      (k) => k !== undefined && typeof k != "boolean",
    ) && throwCellEditError("Transpose and skip blanks must be boolean."),
    c !== "none" &&
      (l === "formats" || l === "formulas") &&
      throwCellEditError(
        "Arithmetic paste requires All or Values mode; its results are numeric values.",
      ),
    a.sourceWorkbookId !== e.id &&
      (l === "all" || l === "formulas") &&
      a.cells.some(
        (k) =>
          typeof k.cell.raw == "string" &&
          k.cell.raw.startsWith("=") &&
          !k.cell.literal,
      ) &&
      throwCellEditError(
        "Use Values paste for formulas copied from a different workbook.",
      ));
  const h = o.transpose ? a.colCount : a.rowCount,
    f = o.transpose ? a.rowCount : a.colCount,
    p = d.firstRow === d.lastRow && d.firstCol === d.lastCol,
    m = p ? h : d.lastRow - d.firstRow + 1,
    v = p ? f : d.lastCol - d.firstCol + 1;
  ((m % h || v % f) &&
    throwCellEditError(
      "The target rectangle must be a whole multiple of the copied rows and columns.",
    ),
    m * v > workbookLimits.maxPatches &&
      throwCellEditError("Paste into at most 500 target cells."),
    (d.firstRow + m > s.rowCount || d.firstCol + v > s.colCount) &&
      throwCellEditError(
        "The pasted cells would exceed the worksheet boundary.",
      ));
  const g = formatRangeAddress(d.firstRow, d.firstCol, m, v),
    b = d.firstRow + m - 1,
    y = d.firstCol + v - 1;
  for (const k of s.merges ?? []) {
    const F = parseRangeAddress(k);
    d.firstRow <= F.lastRow &&
      b >= F.firstRow &&
      d.firstCol <= F.lastCol &&
      y >= F.firstCol &&
      !(
        m * v === 1 &&
        d.firstRow === F.firstRow &&
        d.firstCol === F.firstCol
      ) &&
      throwCellEditError(
        "Unmerge the target area before pasting a block of cells or editing a covered cell.",
      );
  }
  const w = [],
    E = calculateWorkbookWithSettings(e);
  let A = 0;
  for (let k = 0; k < m; k++)
    for (let F = 0; F < v; F++) {
      const I = o.transpose ? F % f : k % h,
        O = o.transpose ? k % h : F % f,
        j = a.cells[I * a.colCount + O],
        z = (j.cell.raw === null && j.value === null) || j.cell.raw === "";
      if (o.skipBlanks && z) {
        A++;
        continue;
      }
      const V = formatCellAddress(d.firstRow + k, d.firstCol + F),
        D = s.cells[V];
      let B = stripCellDecorations(j.cell);
      typeof B.raw == "string" &&
        B.raw.startsWith("=") &&
        !B.literal &&
        (B.raw = shiftFormulaReferences(
          B.raw,
          d.firstRow + k - u.firstRow - I,
          d.firstCol + F - u.firstCol - O,
        ));
      let _;
      if (l === "all") _ = mergeCellDecorations(B, D);
      else if (l === "formulas")
        _ = { ...(D ?? { raw: null }), raw: B.raw, literal: B.literal };
      else if (l === "values")
        _ = withRawCellValue(j.value, { ...(D ?? { raw: null }) });
      else {
        _ = { ...(D ?? { raw: null }) };
        for (const C of ["style", "format", "numFmt", "numberFormat"])
          delete _[C];
        (B.style !== undefined && (_.style = B.style),
          B.format !== undefined && (_.format = B.format),
          B.numFmt !== undefined && (_.numFmt = B.numFmt),
          B.numberFormat !== undefined &&
            (_.numberFormat = structuredClone(B.numberFormat)));
      }
      if ((assertNamedStyleBindingCurrent(e, _.style), c !== "none")) {
        const C = j.value === null || j.value === "" ? 0 : j.value,
          T = E(r, V),
          N = T === null || T === "" ? 0 : T;
        ((typeof C != "number" ||
          !Number.isFinite(C) ||
          typeof N != "number" ||
          !Number.isFinite(N)) &&
          throwCellEditError(
            `Arithmetic paste requires numbers or blanks at ${V}.`,
          ),
          c === "divide" &&
            C === 0 &&
            throwCellEditError(
              `Arithmetic paste would divide by zero at ${V}.`,
            ));
        const R =
          c === "add"
            ? N + C
            : c === "subtract"
              ? N - C
              : c === "multiply"
                ? N * C
                : N / C;
        (Number.isFinite(R) ||
          throwCellEditError(`Arithmetic paste would overflow at ${V}.`),
          (_ = { ..._, raw: R }),
          delete _.literal);
      }
      ((_ = normalizeCellRecordInput(_)),
        cellContentsMatch(D, _) || w.push({ sheetId: r, address: V, cell: _ }));
    }
  return buildCellChangePreview(
    e,
    w,
    {
      targetSheetId: r,
      targetRange: g,
      targetCount: m * v,
      sourceRange: a.sourceRange,
      skippedBlanks: A,
      summary: `Paste ${l}${o.transpose ? " transposed" : ""}${c !== "none" ? ` (${c})` : ""} into ${g}.`,
    },
    o.readOnly,
    o.deferValidationWarnings,
  );
}
function applyPastePreview(e, r, t = {}) {
  return applyCellChangePreview(e, r, t);
}
function throwFillError(e) {
  throw new Error(e);
}
const formatCellReference = (e, r) => `${columnIndexToLetters(r)}${e + 1}`;
function validateFillDirection(e) {
  return (
    ["down", "right", "up", "left"].includes(e) ||
      throwFillError("Choose a supported fill direction."),
    e
  );
}
function formatRangeBoundsAddress(e) {
  const r = parseRangeAddress(e),
    t = formatCellReference(r.firstRow, r.firstCol),
    n = formatCellReference(r.lastRow, r.lastCol);
  return t === n ? t : `${t}:${n}`;
}
function withCellDecorations(e, r) {
  return {
    ...e,
    ...(r?.note !== undefined ? { note: r.note } : {}),
    ...(r?.hyperlink !== undefined
      ? { hyperlink: structuredClone(r.hyperlink) }
      : {}),
    ...(r?.tickmark !== undefined ? { tickmark: r.tickmark } : {}),
  };
}
function cellContentsDiffer(e, r) {
  return (
    JSON.stringify(normalizeCellRecordInput(e ?? { raw: null })) !==
    JSON.stringify(normalizeCellRecordInput(r))
  );
}
function buildFillCellPatches(e, r, t, n) {
  const o = stripCellDecorations(e.cells[r]);
  if (o.raw !== null) return o;
  const a = t(e.id, r);
  if (a === null) return o;
  const s = t.getSpillAnchor(e.id, r);
  if (s && s !== r) {
    const d = _t(s);
    (d.row < n.firstRow ||
      d.row > n.lastRow ||
      d.col < n.firstCol ||
      d.col > n.lastCol) &&
      ((o.raw = a),
      typeof a == "string" && a.startsWith("=") && (o.literal = true));
  }
  return o;
}
function assertFillTargetUnmerged(e, r, t) {
  const n = findSheetOrThrow(e, r);
  for (const o of n.merges ?? []) {
    const a = parseRangeAddress(o);
    t.some((s) => {
      const d = _t(s);
      return (
        d.row >= a.firstRow &&
        d.row <= a.lastRow &&
        d.col >= a.firstCol &&
        d.col <= a.lastCol &&
        (d.row !== a.firstRow || d.col !== a.firstCol)
      );
    }) && throwFillError("Unmerge covered target cells before filling them.");
  }
}
function buildFillPreview(e, r, t, n, o = {}) {
  const a = validateFillDirection(n),
    s = findSheetOrThrow(e, r),
    d = parseRangeAddress(t);
  (d.lastRow >= s.rowCount || d.lastCol >= s.colCount) &&
    throwFillError("The fill range exceeds the worksheet.");
  const u = a === "down" || a === "up",
    l = d.lastRow - d.firstRow + 1,
    c = d.lastCol - d.firstCol + 1,
    h = u ? (l - 1) * c : (c - 1) * l;
  h > workbookLimits.maxPatches &&
    throwFillError("Fill at most 500 destination cells.");
  const f = a === "up" ? d.lastRow : d.firstRow,
    p = a === "left" ? d.lastCol : d.firstCol,
    m = [],
    v = [],
    g = calculateWorkbookWithSettings(e),
    b = u
      ? { ...d, firstRow: f, lastRow: f }
      : { ...d, firstCol: p, lastCol: p };
  if (h)
    for (let y = d.firstRow; y <= d.lastRow; y++)
      for (let w = d.firstCol; w <= d.lastCol; w++) {
        if (u ? y === f : w === p) continue;
        const E = u ? f : y,
          A = u ? w : p,
          k = formatCellReference(y, w);
        v.push(k);
        const F = buildFillCellPatches(s, formatCellReference(E, A), g, b);
        typeof F.raw == "string" &&
          F.raw.startsWith("=") &&
          !F.literal &&
          (F.raw = shiftFormulaReferences(F.raw, y - E, w - A));
        const I = normalizeCellRecordInput(withCellDecorations(F, s.cells[k]));
        cellContentsDiffer(s.cells[k], I) &&
          m.push({ sheetId: r, address: k, cell: I });
      }
  return (
    assertFillTargetUnmerged(e, r, v),
    buildCellChangePreview(
      e,
      m,
      {
        targetSheetId: r,
        targetRange: formatRangeBoundsAddress(t),
        targetCount: h,
        generatedCount: h,
        direction: a,
        summary: `Fill ${a} within ${formatRangeBoundsAddress(t)}.`,
      },
      o.readOnly,
    )
  );
}
const s$ = 864e5,
  excelEpochUtc = Date.UTC(1899, 11, 31),
  maxExcelDateSerial = 2958465;
class DateSeriesRangeError extends Error {}
function throwDateSeriesRangeError() {
  throw new DateSeriesRangeError(
    "Date series must stay between 1899-12-31 and 9999-12-31.",
  );
}
function utcDateFromParts(e, r, t) {
  const n = new Date(0);
  return (n.setUTCHours(0, 0, 0, 0), n.setUTCFullYear(e, r, t), n);
}
function excelSerialToDate(e) {
  return new Date(excelEpochUtc + (e > 60 ? e - 1 : e) * s$);
}
function dateToExcelSerial(e) {
  const r = (e.getTime() - excelEpochUtc) / s$,
    t = r >= 60 ? r + 1 : r;
  return (
    (!Number.isSafeInteger(t) || t < 0 || t > maxExcelDateSerial) &&
      throwDateSeriesRangeError(),
    t
  );
}
function parseSeriesSeedDate(e, r = "1900") {
  if (typeof e == "number")
    return (
      (!Number.isSafeInteger(e) || e < 0 || e > dateSystemMaxSerial(r)) &&
        throwFillError(
          `Date values must be whole Excel ${r} serials between 0 and ${dateSystemMaxSerial(r)}.`,
        ),
      e
    );
  if (
    ((typeof e != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(e)) &&
      throwFillError(
        "Enter a date as YYYY-MM-DD or a whole Excel date serial.",
      ),
    e === "1900-02-29" && r === "1900")
  )
    return 60;
  const [t, n, o] = e.split("-").map(Number),
    a = utcDateFromParts(t, n - 1, o);
  return (
    (a.getUTCFullYear() !== t ||
      a.getUTCMonth() !== n - 1 ||
      a.getUTCDate() !== o) &&
      throwFillError("The series date is not a valid calendar date."),
    parseSeriesSeedDate(dateToExcelSerial(a) - dateSystemSerialOffset(r), r)
  );
}
function stepSeriesDateSerial(e, r, t, n) {
  if (
    (Number.isSafeInteger(r) ||
      throwFillError("The date step exceeds safe integer limits."),
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
    throwFillError(
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
function buildSeriesPreview(e, r, t, n) {
  const o = validateFillDirection(n.direction),
    a = findSheetOrThrow(e, r),
    s = parseRangeAddress(t),
    d = o === "down" || o === "up";
  (["number", "date"].includes(n.kind) ||
    throwFillError("Choose a numeric or date series."),
    ((d && s.firstCol !== s.lastCol) || (!d && s.firstRow !== s.lastRow)) &&
      throwFillError(
        "Select one column for vertical series or one row for horizontal series.",
      ),
    (s.lastRow >= a.rowCount || s.lastCol >= a.colCount) &&
      throwFillError("The series range exceeds the worksheet."));
  const u = d ? s.lastRow - s.firstRow + 1 : s.lastCol - s.firstCol + 1;
  u > workbookLimits.maxPatches &&
    throwFillError("A series may contain at most 500 target cells.");
  const l = o === "up" ? s.lastRow : s.firstRow,
    c = o === "left" ? s.lastCol : s.firstCol,
    h = formatCellReference(l, c),
    f = stripCellDecorations(a.cells[h]),
    p = n.step ?? 1;
  (typeof p != "number" ||
    !Number.isFinite(p) ||
    p === 0 ||
    (n.kind === "date" && !Number.isSafeInteger(p))) &&
    throwFillError(
      "Series step must be a nonzero finite number; date steps must be whole numbers.",
    );
  const m = n.dateUnit ?? "day",
    v = n.monthPolicy ?? "preserveDay";
  (!["day", "weekday", "month", "year"].includes(m) ||
    !["preserveDay", "endOfMonth"].includes(v)) &&
    throwFillError("Choose a supported date unit and month policy.");
  const g = n.start ?? calculateWorkbookWithSettings(e)(r, h);
  typeof g != "number" &&
    typeof g != "string" &&
    throwFillError(
      "The directional seed must contain a number or date, or supply a start value.",
    );
  const b =
      n.kind === "date"
        ? parseSeriesSeedDate(g, e.datePolicy?.dateSystem)
        : typeof g == "number" && Number.isFinite(g)
          ? g
          : throwFillError(
              "Numeric series require a finite numeric start value.",
            ),
    y =
      n.stop === undefined
        ? undefined
        : n.kind === "date"
          ? parseSeriesSeedDate(n.stop, e.datePolicy?.dateSystem)
          : typeof n.stop == "number" && Number.isFinite(n.stop)
            ? n.stop
            : throwFillError(
                "Numeric series require a finite numeric stop value.",
              );
  y !== undefined &&
    ((p > 0 && y < b) || (p < 0 && y > b)) &&
    throwFillError("The stop value is opposite the series step direction.");
  const w = [],
    E = [];
  let A = 0;
  for (let O = 0; O < u; O++) {
    let j;
    try {
      j =
        n.kind === "date"
          ? Nae(b, O * p, m, v, e.datePolicy?.dateSystem ?? "1900")
          : b + O * p;
    } catch (D) {
      if (y !== undefined && O > 0 && D instanceof DateSeriesRangeError) break;
      throw D;
    }
    if (y !== undefined && (p > 0 ? j > y : j < y)) break;
    Number.isFinite(j) || throwFillError("The numeric series would overflow.");
    const z = formatCellReference(
      l + (d ? O * (o === "up" ? -1 : 1) : 0),
      c + (d ? 0 : O * (o === "left" ? -1 : 1)),
    );
    (E.push(z), A++);
    const V = normalizeCellRecordInput(
      withCellDecorations(
        {
          ...f,
          raw: j,
          literal: undefined,
          ...(n.kind === "date"
            ? { format: "date", numFmt: "yyyy-mm-dd", numberFormat: undefined }
            : {}),
        },
        a.cells[z],
      ),
    );
    cellContentsDiffer(a.cells[z], V) &&
      w.push({ sheetId: r, address: z, cell: V });
  }
  return (
    assertFillTargetUnmerged(e, r, E),
    buildCellChangePreview(
      e,
      w,
      {
        targetSheetId: r,
        targetRange: formatRangeBoundsAddress(t),
        targetCount: u,
        generatedCount: A,
        direction: o,
        summary: `Fill ${n.kind === "date" ? `${m} date` : "numeric"} series ${o} in ${formatRangeBoundsAddress(t)} (${A} values).`,
      },
      n.readOnly,
    )
  );
}
function applyFillPreview(e, r, t = {}) {
  return applyCellChangePreview(e, r, t);
}
function buildFillHandlePreview(e, r, t, n, o = {}) {
  const a = findSheetOrThrow(e, r),
    s = parseRangeAddress(t),
    d = parseRangeAddress(n),
    u = o.mode ?? "copy";
  (["copy", "series"].includes(u) ||
    throwFillError("Choose Copy or Series for the fill handle."),
    (s.firstRow < d.firstRow ||
      s.lastRow > d.lastRow ||
      s.firstCol < d.firstCol ||
      s.lastCol > d.lastCol) &&
      throwFillError(
        "The fill-handle target must contain the complete source range.",
      ),
    (d.lastRow >= a.rowCount || d.lastCol >= a.colCount) &&
      throwFillError("The fill-handle target exceeds the worksheet."));
  const l = s.lastRow - s.firstRow + 1,
    c = s.lastCol - s.firstCol + 1,
    h = s.firstCol === d.firstCol && s.lastCol === d.lastCol,
    f = s.firstRow === d.firstRow && s.lastRow === d.lastRow;
  !h &&
    !f &&
    throwFillError(
      "Extend the fill handle along one axis while keeping the other source dimension unchanged.",
    );
  const p = (d.lastRow - d.firstRow + 1) * (d.lastCol - d.firstCol + 1) - l * c;
  p > workbookLimits.maxPatches &&
    throwFillError("The fill handle may extend into at most 500 cells.");
  const m = h
      ? d.firstRow < s.firstRow
        ? "up"
        : "down"
      : d.firstCol < s.firstCol
        ? "left"
        : "right",
    v = [],
    g = [],
    b = calculateWorkbookWithSettings(e);
  let y = 0,
    w = 0;
  if (u === "series") {
    (h && l === 2 && c === 1) ||
      (f && l === 1 && c === 2) ||
      throwFillError(
        "Series handle fill requires exactly two seed cells in the extension's row or column.",
      );
    const k = b(r, formatCellReference(s.firstRow, s.firstCol)),
      F = b(r, formatCellReference(s.lastRow, s.lastCol));
    ((typeof k != "number" ||
      !Number.isFinite(k) ||
      typeof F != "number" ||
      !Number.isFinite(F) ||
      !Number.isFinite(F - k)) &&
      throwFillError("Series handle seeds must evaluate to finite numbers."),
      (y = k),
      (w = F - k));
  }
  const E = (k, F) => ((k % F) + F) % F;
  function* A() {
    if (h)
      for (const [k, F] of [
        [d.firstRow, s.firstRow - 1],
        [s.lastRow + 1, d.lastRow],
      ])
        for (let I = k; I <= F; I++)
          for (let O = s.firstCol; O <= s.lastCol; O++)
            yield { row: I, col: O };
    else
      for (let k = s.firstRow; k <= s.lastRow; k++)
        for (const [F, I] of [
          [d.firstCol, s.firstCol - 1],
          [s.lastCol + 1, d.lastCol],
        ])
          for (let O = F; O <= I; O++) yield { row: k, col: O };
  }
  if (p)
    for (const { row, col } of A()) {
      const I =
          u === "series" ? s.firstRow : s.firstRow + E(row - s.firstRow, l),
        O = u === "series" ? s.firstCol : s.firstCol + E(col - s.firstCol, c),
        j = buildFillCellPatches(a, formatCellReference(I, O), b, s),
        z = formatCellReference(row, col);
      if (u === "series") {
        const D = y + (h ? row - s.firstRow : col - s.firstCol) * w;
        (Number.isFinite(D) ||
          throwFillError("The handle series would overflow."),
          (j.raw = D),
          delete j.literal);
      } else
        typeof j.raw == "string" &&
          j.raw.startsWith("=") &&
          !j.literal &&
          (j.raw = shiftFormulaReferences(j.raw, row - I, col - O));
      const V = normalizeCellRecordInput(withCellDecorations(j, a.cells[z]));
      (g.push(z),
        cellContentsDiffer(a.cells[z], V) &&
          v.push({ sheetId: r, address: z, cell: V }));
    }
  return (
    assertFillTargetUnmerged(e, r, g),
    buildCellChangePreview(
      e,
      v,
      {
        targetSheetId: r,
        targetRange: formatRangeBoundsAddress(n),
        targetCount: p,
        generatedCount: p,
        direction: m,
        summary: `Fill handle ${u} from ${formatRangeBoundsAddress(t)} into ${formatRangeBoundsAddress(n)} (${p} extension cells).`,
      },
      o.readOnly,
    )
  );
}
function parseDateCellEntry(e, r, t = "1900") {
  if (typeof e != "string") throw new Error("Cell entry must be text.");
  const n = r ? structuredClone(r) : { raw: null };
  if (((n.literal = e.startsWith("'")), n.literal)) n.raw = e.slice(1);
  else if (/^\d{4}-\d{2}-\d{2}$/.test(e)) {
    try {
      n.raw = parseSeriesSeedDate(e, t);
    } catch {
      throw new Error(
        `Enter a valid YYYY-MM-DD date between ${t === "1904" ? "1904-01-01" : "1899-12-31"} and 9999-12-31, or begin with an apostrophe to keep the entry as text.${t === "1900" ? " Excel's compatibility date 1900-02-29 represents serial 60." : ""}`,
      );
    }
    ((n.format = "date"),
      (n.numFmt = "yyyy-mm-dd"),
      (n.numberFormat = {
        category: "date",
        locale: "en-US",
        pattern: "yyyy-mm-dd",
      }));
  } else n.raw = parseRawCellEntry(e);
  return normalizeCellRecordInput(n);
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
function resolveWorkbookLocale(e) {
  if (e === undefined) return "en-US";
  if (!supportedWorkbookLocales.includes(e))
    throw new Error("Choose a supported workbook interface locale.");
  return e;
}
function formatNumberForLocale(e, r, t = true) {
  return new Intl.NumberFormat(r, {
    useGrouping: t,
    maximumSignificantDigits: 15,
  }).format(e);
}
function parseLocalizedCellEntry(e, r, t = "1900", n) {
  if (
    (resolveWorkbookLocale(n),
    typeof r?.raw == "number" && e === localizeDecimalPoint(r.raw, n))
  )
    return normalizeCellRecordInput(structuredClone(r));
  if (!n || /^[=']/.test(e) || /^\d{4}-\d{2}-\d{2}$/.test(e))
    return parseDateCellEntry(e, r, t);
  const o = n === "de-DE" || n === "fr-FR" ? "," : ".",
    a = n === "de-DE" ? "." : n === "fr-FR" ? "[ \\u00a0\\u202f]" : ",",
    s = (l) => l.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"),
    d = `(?:[1-9]\\d{0,2}(?:${n === "fr-FR" ? a : s(a)}\\d{3})+)`;
  if (new RegExp(`^-?(?:0|[1-9]\\d*|${d})(?:${s(o)}\\d+)?$`).test(e)) {
    const l = e
      .replace(n === "fr-FR" ? /[ \u00a0\u202f]/g : new RegExp(s(a), "g"), "")
      .replace(o, ".");
    return parseDateCellEntry(l, r, t);
  }
  return n === "de-DE" && /^(WAHR|FALSCH)$/i.test(e)
    ? parseDateCellEntry(e.toUpperCase() === "WAHR" ? "TRUE" : "FALSE", r, t)
    : n === "fr-FR" && /^(VRAI|FAUX)$/i.test(e)
      ? parseDateCellEntry(e.toUpperCase() === "VRAI" ? "TRUE" : "FALSE", r, t)
      : /^-?[\d., \u00a0\u202f]+$/.test(e) && /[.,]/.test(e)
        ? parseDateCellEntry("'" + e, r, t)
        : parseDateCellEntry(e, r, t);
}
function formatCellEntryForEdit(e, r) {
  return typeof e?.raw == "number" && r
    ? localizeDecimalPoint(e.raw, r)
    : `${e != null && e.literal && typeof e.raw == "string" ? "'" : ""}${String(e?.raw ?? "")}`;
}
function localizeDecimalPoint(e, r) {
  return Number.isFinite(e)
    ? r === "de-DE" || r === "fr-FR"
      ? String(e).replace(".", ",")
      : String(e)
    : "";
}
function validateLocalizedNumberEntry(e, r, t) {
  let n = e.trim();
  {
    resolveWorkbookLocale(t);
    const o = t === "de-DE" || t === "fr-FR" ? "," : ".",
      a = t === "de-DE" ? "\\." : t === "fr-FR" ? "[ \\u00a0\\u202f]" : ",";
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
function formatLocalizedNumberDisplay(e, r, t = "1900", n) {
  return !n ||
    typeof e != "number" ||
    (r != null && r.numberFormat) ||
    (r != null && r.numFmt) ||
    r?.format === "date"
    ? formatNumberWithPattern(e, r, t)
    : Number.isFinite(e)
      ? r?.format === "currency"
        ? new Intl.NumberFormat(n, {
            style: "currency",
            currency: "USD",
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          }).format(e)
        : r?.format === "percent"
          ? new Intl.NumberFormat(n, {
              style: "percent",
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            }).format(e)
          : r?.format === "number"
            ? new Intl.NumberFormat(n, {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2,
              }).format(e)
            : formatNumberForLocale(e, n, false)
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
function resolveWorkbookMessages(e = {}) {
  const r = resolveWorkbookLocale(e.locale),
    t = {
      ...(r === "de-DE" ? germanWorkbookStatusMessages : {}),
      ...e.messages,
    },
    n = (d) => (Object.hasOwn(t, d) ? d : d.trim()),
    o = (d) => [...new Set(d.match(/\{\w+\}/g) ?? [])].sort().join("|"),
    a = (d) =>
      Object.hasOwn(t, n(d)) &&
      typeof t[n(d)] == "string" &&
      t[n(d)].trim().length > 0 &&
      o(d) === o(t[n(d)]),
    s = (d) => {
      return a(d)
        ? n(d) === d
          ? t[d]
          : `${d.match(/^\s*/)?.[0] ?? ""}${t[n(d)]}${d.match(/\s*$/)?.[0] ?? ""}`
        : d;
    };
  return {
    locale: r,
    entryLocale: e.locale,
    text: s,
    language: (d) => (a(d) ? r : "en"),
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
  const r = useWorkbookMessages(),
    t = ne.Children.toArray(children);
  if (!t.every((o) => typeof o == "string"))
    return i.jsx(i.Fragment, { children });
  const n = t.join("");
  return r.locale.startsWith("en") && r.text(n) === n
    ? i.jsx(i.Fragment, { children })
    : i.jsx("span", { lang: r.language(n), children: r.text(n) });
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
  const h = useWorkbookMessages(),
    f = ne.useId(),
    [p, m] = ne.useState(() =>
      value === "" || value === undefined
        ? ""
        : localizeDecimalPoint(Number(value), h.entryLocale),
    ),
    [v, g] = ne.useState(false),
    [b, y] = ne.useState(false),
    w = ne.useRef(h.entryLocale),
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
  } catch (V) {
    O =
      V instanceof Error && V.message === "bounds"
        ? h.message("Enter a number from {min} to {max}.", {
            min:
              A === undefined ? "−∞" : localizeDecimalPoint(A, h.entryLocale),
            max: k === undefined ? "∞" : localizeDecimalPoint(k, h.entryLocale),
          })
        : h.text("Enter a valid number using the selected locale.");
  }
  (ne.useLayoutEffect(() => {
    const V = value === "" || value === undefined ? "" : String(Number(value));
    (E.current !== undefined &&
      (V === E.current || (V === "" && E.current === "NaN"))) ||
      ((w.current = h.entryLocale),
      m(
        value === "" || value === undefined
          ? ""
          : localizeDecimalPoint(Number(value), h.entryLocale),
      ),
      y(false));
  }, [value]),
    ne.useLayoutEffect(() => {
      v ||
        O ||
        ((w.current = h.entryLocale),
        m(I === undefined ? "" : localizeDecimalPoint(I, h.entryLocale)));
    }, [h.entryLocale, v]));
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
        v && w.current !== h.entryLocale ? `${f}-locale` : null,
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
        onChange: (V) => j(V.target.value),
        onFocus: (V) => {
          (O || (w.current = h.entryLocale),
            g(true),
            onFocus == null || onFocus(V));
        },
        onBlur: (V) => {
          (y(true), g(false), onBlur == null || onBlur(V));
        },
        onKeyDown: (V) => {
          if (
            (onKeyDown == null || onKeyDown(V),
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
        w.current !== h.entryLocale &&
        i.jsx("span", {
          id: `${f}-locale`,
          className: "lg-number-error",
          children: h.text(
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
function findOpenFormulaQuote(e, r) {
  let t = "";
  for (let n = 1; n < r; n++) {
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
function findFormulaReferenceSpan(e, r, t = r) {
  if (!(!e.startsWith("=") || r < 1 || t < r || t > e.length)) {
    for (const n of e.matchAll(formulaReferencePattern)) {
      const o = n.index,
        a = o + n[0].length;
      if (
        !(
          findOpenFormulaQuote(e, o) ||
          /[\w.\[!]/.test(e[o - 1] ?? "") ||
          /[\w.\](]/.test(e[a] ?? "")
        ) &&
        r >= o &&
        t <= a
      )
        return { start: o, end: a };
    }
    if (!findOpenFormulaQuote(e, r)) {
      if (t > r) return { start: r, end: t };
      if (/[=(:,;+\-*/^&<>]$/.test(e.slice(0, r).trimEnd()))
        return { start: r, end: t };
    }
  }
}
function replaceFormulaReferenceSpan(e, r, t, n) {
  const o = n === undefined ? t : `'${n.replaceAll("'", "''")}'!${t}`;
  return {
    value: e.slice(0, r.start) + o + e.slice(r.end),
    span: { start: r.start, end: r.start + o.length },
  };
}
function $ae(e, r, t = r) {
  const n = findFormulaReferenceSpan(e, r, t);
  if (!n || n.start === n.end) return;
  const o = e.slice(n.start, n.end),
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
      value: e.slice(0, n.start) + p + e.slice(n.end),
      span: { start: n.start, end: n.start + p.length },
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
    value: e.slice(0, n.start) + h + e.slice(n.end),
    span: { start: n.start, end: n.start + h.length },
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
function getFormulaScanList(e, r) {
  const t = isDeepFrozenValue(e);
  let n = t ? formulaScanCache.get(e) : undefined;
  if (!n) {
    n = { calls: [], external: [] };
    for (const o in e) {
      const a = e[o];
      typeof a.raw != "string" ||
        !a.raw.startsWith("=") ||
        (a.raw.includes("(") && n.calls.push([o, a]),
        /[\[|\\]|https?:|file:/i.test(a.raw) && n.external.push([o, a]));
    }
    t && formulaScanCache.set(e, n);
  }
  return n[r];
}
const objectIdRegistry = new WeakMap();
let objectIdCounter = 0;
function getStableObjectId(e) {
  let r = objectIdRegistry.get(e);
  return (r || ((r = ++objectIdCounter), objectIdRegistry.set(e, r)), r);
}
function validateExtensionHostPolicy(e, r) {
  const t = r.extensions ?? [],
    n = r.allowedSheetIds ?? e.sheets.map((a) => a.id);
  if (r.readOnly !== undefined && typeof r.readOnly != "boolean")
    throw Error("Read-only policy must be boolean.");
  if (
    !Array.isArray(t) ||
    t.length > extensionRuntimeLimits.extensions ||
    !Array.isArray(n) ||
    n.some((a) => !e.sheets.some((s) => s.id === a))
  )
    throw Error("Invalid extension host scope.");
  const o = new Set();
  for (const a of t) {
    const s = validateExtensionManifest(a.manifest);
    if (
      o.has(s.id) ||
      typeof a.invoke != "function" ||
      !Array.isArray(a.allowedSheetIds) ||
      a.allowedSheetIds.some((d) => !e.sheets.some((u) => u.id === d)) ||
      !Array.isArray(a.grantedPermissions) ||
      a.grantedPermissions.some((d) => !s.permissions.includes(d)) ||
      (a.includeHidden !== undefined && typeof a.includeHidden != "boolean")
    )
      throw Error("Invalid host extension registration.");
    o.add(s.id);
  }
  return {
    extensions: t,
    allowed: n,
    hash: () =>
      fingerprintCanonicalValue([
        !!r.readOnly,
        n,
        t.map((a) => [
          hashExtensionManifest(a.manifest),
          a.grantedPermissions,
          a.allowedSheetIds,
          !!a.includeHidden,
          getStableObjectId(a.invoke),
        ]),
      ]),
  };
}
const resolveExtensionSheetScope = (e, r) =>
  fingerprintCanonicalValue([
    r.grantedPermissions,
    [...r.allowedSheetIds].sort().map((t) => {
      const n = e.sheets.find((o) => o.id === t);
      return [t, !!n.hidden, !!n.veryHidden];
    }),
    !!r.includeHidden,
  ]);
function buildExtensionHostScope(e, r) {
  return validateExtensionHostPolicy(e, r)
    .extensions.filter((t) =>
      t.manifest.permissions.every((n) => t.grantedPermissions.includes(n)),
    )
    .map((t) => ({
      id: t.manifest.id,
      manifestHash: hashExtensionManifest(t.manifest),
      scopeHash: resolveExtensionSheetScope(e, t),
      sheetIds: t.allowedSheetIds.filter((n) => {
        const o = e.sheets.find((a) => a.id === n);
        return !(o.hidden || o.veryHidden) || !!t.includeHidden;
      }),
    }));
}
function withExtensionHostScope(e, r = []) {
  const t = { ...e };
  return (
    setExtensionCacheEntries(t, buildExtensionHostScope(e, { extensions: r })),
    isDeepFrozenValue(e) ? deepFreezeValue(t) : t
  );
}
function assertExtensionWriteScope(e, r, t) {
  var o;
  if (t.readOnly) throw Error("The workbook is read-only.");
  if ((o = e.structureProtection) != null && o.enabled)
    throw Error("Workbook structure is protected.");
  const n = validateExtensionHostPolicy(e, t);
  for (const a of r) {
    if (!n.allowed.includes(a))
      throw Error("An extension destination is outside the host write scope.");
    assertSheetUnprotected(e.sheets.find((s) => s.id === a));
  }
}
function reviewExtensionFormulas(e) {
  const r = new Set(builtinFunctionNames),
    t = new Set((e.definedNames ?? []).map((o) => o.name.toUpperCase())),
    n = [];
  for (const o of e.sheets)
    for (const [a, s] of getFormulaScanList(o.cells, "calls")) {
      if (typeof s.raw != "string" || !s.raw.startsWith("=")) continue;
      const d = [],
        u = (l, c = false) => {
          for (const h of l) {
            if ((h.kind === "group" && u(h.body, c), h.kind !== "call"))
              continue;
            const f = h.callee;
            let p = false;
            if (f.kind === "symbol" && !f.qualifier) {
              const m = f.name.replace(/^_xlfn\./i, "").toUpperCase();
              ((p = !r.has(m) && !t.has(m) && !["LET", "LAMBDA"].includes(m)),
                p &&
                  d.push({
                    name: m,
                    start: h.start,
                    end: h.end,
                    arguments: h.args.map((v) =>
                      v.length
                        ? s.raw.toString().slice(v[0].start, v.at(-1).end)
                        : "",
                    ),
                    nested: c,
                  }));
            }
            h.args.forEach((m) => u(m, c || p));
          }
        };
      try {
        u(tokenizeFormulaText(s.raw));
      } catch {
        continue;
      }
      if (
        (d.length &&
          n.push({
            sheetId: o.id,
            sheetName: o.name,
            address: a,
            formula: s.raw,
            literal: !!s.literal,
            calls: d,
          }),
        n.length > extensionRuntimeLimits.calls)
      )
        throw Error("Extension review exceeds 100 formula cells.");
    }
  return n;
}
function requireApprovedExtension(e, r, t) {
  const n = e.extensions?.installed.find((s) => s.manifest.id === r);
  if (!n || !n.enabled) throw Error("The extension is missing or disabled.");
  const o = validateExtensionHostPolicy(e, t).extensions.find(
    (s) => s.manifest.id === r,
  );
  if (
    !o ||
    hashExtensionManifest(o.manifest) !== hashExtensionManifest(n.manifest)
  )
    throw Error(
      "An approved adapter with the exact extension version and manifest is unavailable.",
    );
  if (!n.manifest.permissions.every((s) => o.grantedPermissions.includes(s)))
    throw Error("The host has not granted all declared extension permissions.");
  return o;
}
function evaluateExtensionFunctionCall(e, r, t, n) {
  const o = n.manifest.functions.find(
    (s) => `${n.manifest.namespace}.${s.id}` === t.name,
  );
  if (!o || o.parameters.length !== t.arguments.length)
    throw Error(
      "Custom function arguments do not match the declared signature.",
    );
  if (t.nested || /\b(?:LET|LAMBDA)\s*\(/i.test(r.formula))
    throw Error(
      "Nested custom calls and custom calls inside LET/LAMBDA require separate worksheet cells.",
    );
  const a = (s) => {
    const d = e.sheets.find((u) => u.id === s);
    return (
      !!d &&
      n.allowedSheetIds.includes(s) &&
      (!(d.hidden || d.veryHidden) || !!n.includeHidden)
    );
  };
  if (!a(r.sheetId))
    throw Error(
      "The extension formula is outside its approved worksheet scope.",
    );
  return t.arguments.map((s, d) => {
    if (!s.trim()) throw Error("Extension arguments cannot be omitted.");
    const u = o.parameters[d].schema,
      l =
        /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*)(?::(\$?[A-Za-z]{1,3}\$?[1-9]\d*))?$/.exec(
          s.trim(),
        );
    if (l) {
      const g = l[1]?.replace(/''/g, "'") ?? l[2],
        b = g
          ? e.sheets.find((F) => F.name.toLowerCase() === g.toLowerCase())
          : e.sheets.find((F) => F.id === r.sheetId);
      if (!b || !a(b.id))
        throw Error(
          "An extension input worksheet is unavailable or unauthorized.",
        );
      const y = parseRangeAddress(l[3] + (l[4] ? ":" + l[4] : "")),
        w = y.lastRow - y.firstRow + 1,
        E = y.lastCol - y.firstCol + 1;
      if (
        w * E > extensionRuntimeLimits.inputCells ||
        (u.type !== "matrix" && w * E !== 1)
      )
        throw Error(
          "The extension input range exceeds its declared scalar/cell budget.",
        );
      const A = evaluateWorkbookFormulas(
          e,
          {},
          { allowRead: a, allowName: () => true, step: () => {} },
        ),
        k = Array.from({ length: w }, (F, I) =>
          Array.from({ length: E }, (O, j) => {
            const z = A.getTypedValue(
              b.id,
              `${columnIndexToLetters(y.firstCol + j)}${y.firstRow + I + 1}`,
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
      return validateExtensionValue(u.type === "matrix" ? k : k[0][0], u);
    }
    if (u.type === "matrix")
      throw Error("Matrix parameters require a static worksheet range.");
    const c = {
      ...e,
      sheets: e.sheets.map((g) =>
        g.id === r.sheetId
          ? { ...g, cells: { ...g.cells, [r.address]: { raw: "=" + s } } }
          : g,
      ),
    };
    setExtensionCacheEntries(c, getExtensionCacheEntries(e));
    const h = evaluateWorkbookFormulas(
        c,
        {},
        { allowRead: a, allowName: () => true, step: () => {} },
      ),
      f = h.getTypedValue(r.sheetId, r.address);
    if (f.type === "error" || h.getOperationalErrors().length)
      throw Error(
        "An extension input is erroneous, stale, unauthorized or over its calculation limit.",
      );
    let p = f.value;
    if (h.getSpillRange(r.sheetId, r.address))
      throw Error("Scalar extension arguments cannot contain arrays.");
    return validateExtensionValue(p, u);
  });
}
function resolveExtensionCachedResult(e, r = {}) {
  const t = validateExtensionHostPolicy(e, r),
    n = reviewExtensionFormulas(e),
    o = n.length ? withExtensionHostScope(e, t.extensions) : e,
    a = n.length ? evaluateWorkbookFormulas(o) : undefined,
    s = n.map((u) => {
      const l = a.getTypedValue(u.sheetId, u.address),
        c = (e.extensions?.caches ?? []).filter(
          (v) => v.sheetId === u.sheetId && v.address === u.address,
        ),
        h = u.calls.map((v) => {
          const g = listInstalledExtensionFunctions(e).find(
            (y) => y.name === v.name,
          );
          let b = "unavailable";
          if (
            g &&
            ((b = g.entry.enabled
              ? buildExtensionHostScope(e, r).some(
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
              const y = requireApprovedExtension(e, g.entry.manifest.id, r),
                w = evaluateExtensionFunctionCall(o, u, v, y),
                E = fingerprintExtensionCall(
                  u.sheetId,
                  u.address,
                  u.formula,
                  v.name,
                  w,
                  hashExtensionManifest(g.entry.manifest),
                );
              c.some(
                (A) =>
                  A.key === E &&
                  A.scopeHash === resolveExtensionSheetScope(e, y),
              ) && (b = "reviewed-cache");
            } catch {}
          return { name: v.name, state: b };
        });
      return {
        ...u,
        importedCache: e.importedCalculationCaches?.entries.find(
          (v) =>
            v.sheetId === u.sheetId &&
            v.address === u.address &&
            v.formula === u.formula,
        )?.cache,
        functions: h,
        result: l,
        caches: c,
        cacheLabel: c.length
          ? l.type === "error" || h.some((v) => v.state !== "reviewed-cache")
            ? "Historical cached result — stale or unavailable"
            : "Reviewed cached result — manual refresh"
          : "No retained result",
      };
    });
  return {
    installed: e.extensions?.installed ?? [],
    cells: s,
    available: t.extensions.map((u) => ({
      manifest: u.manifest,
      fullyGranted: u.manifest.permissions.every((l) =>
        u.grantedPermissions.includes(l),
      ),
    })),
    policy:
      "Only reviewed argument values reach a host adapter. No code, add-in URL, COM, XLL or ActiveX is loaded from workbook content.",
  };
}
const tT = new WeakMap();
function runExtensionOperation(e, r, t, n = {}) {
  const o = validateExtensionHostPolicy(e, n),
    a = e.extensions?.installed.find((m) => m.manifest.id === r),
    s = o.extensions.find((m) => m.manifest.id === r);
  if (!["install", "enable", "disable", "remove"].includes(t))
    throw Error("Unknown extension operation.");
  const d = t === "install" ? s?.manifest : a?.manifest;
  if (!d) throw Error("The requested extension is unavailable.");
  if (
    (validateExtensionManifest(d),
    ["install", "enable"].includes(t) &&
      (!s ||
        hashExtensionManifest(s.manifest) !== hashExtensionManifest(d) ||
        !d.permissions.every((m) => s.grantedPermissions.includes(m))))
  )
    throw Error(
      "The host has not approved this exact manifest and all of its permissions.",
    );
  if (
    d.functions.some(
      (m) =>
        builtinFunctionNames.includes(`${d.namespace}.${m.id}`) ||
        (e.definedNames ?? []).some(
          (v) => v.name.toUpperCase() === `${d.namespace}.${m.id}`,
        ),
    )
  )
    throw Error(
      "Extension functions cannot shadow built-ins or defined names.",
    );
  const u = reviewExtensionFormulas(e).filter((m) =>
    m.calls.some((v) => v.name.startsWith(d.namespace + ".")),
  );
  assertExtensionWriteScope(
    e,
    u.map((m) => m.sheetId),
    n,
  );
  const l = normalizeWorkbookInput(e),
    c = structuredClone(
      l.extensions ?? { schemaVersion: 1, installed: [], caches: [] },
    );
  ((c.installed = c.installed.filter((m) => m.manifest.id !== r)),
    t !== "remove" &&
      c.installed.push({
        manifest: validateExtensionManifest(d),
        enabled: t !== "disable",
      }),
    (t === "remove" || t === "install") &&
      (c.caches = c.caches.filter(
        (m) => !m.functionName.startsWith(d.namespace + "."),
      )),
    (l.extensions = c));
  const h = normalizeWorkbookInput(l),
    f = {
      extensionId: r,
      action: t,
      manifest: structuredClone(d),
      warning:
        "Raw formulas are preserved. Disabling retains labeled historical results; installing/upgrading or removing clears this extension’s caches. Execution requires a separate argument review.",
    };
  return (
    tT.set(f, {
      bookHash: fingerprintCanonicalValue(e),
      policyHash: o.hash(),
      planHash: fingerprintCanonicalValue(f),
      candidate: h,
    }),
    f
  );
}
function applyExtensionPreview(e, r, t = {}) {
  const n = tT.get(r);
  if (!n) throw Error("Unknown or already applied extension preview.");
  if (
    t.readOnly ||
    n.bookHash !== fingerprintCanonicalValue(e) ||
    n.policyHash !== validateExtensionHostPolicy(e, t).hash() ||
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
  applyFillPreview,
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
  applyPastePreview,
  formatLocalizedNumberDisplay,
  validateExtensionHostPolicy,
  reviewExtensionFormulas,
  validateLocalizedNumberEntry,
  applyCellChangePreview,
  assertExtensionWriteScope,
  getFormulaScanList,
  autoCloseFormulaParens,
};
