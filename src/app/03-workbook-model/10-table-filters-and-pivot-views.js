import {
  failWithMessage,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import { isNullOrEmptyValue } from "./03-defined-names.js";
import {
  fingerprintCanonicalValue,
  normalizeDatePolicy,
} from "../01-schemas-date-functions-and-validation.js";
import { number, array, unknown, object } from "zod";
import { workbookLimits } from "../02-text-and-financial-functions.js";
import {
  cloneStyleValue,
  normalizeHiddenIntervals,
  validateWorkbookStyleCatalog,
  columnWidthAt,
  reviewThreadLimits,
  cloneCellRecord,
  hasBorderEdgeChanges,
  canonicalStyleKey,
  validateBoundedStyleText,
  validateNamedStyleDefinition,
  assertNamedStyleBindingCurrent,
  cn,
  resolveReviewTimestamp,
  createDefaultStyleCatalog,
  rowHeightAt,
  throwWithMessage,
  assertSheetUnprotected,
  normalizeBoundedText,
  validateWorkbookTheme,
  validateCellStyle,
  createAxisVisibilityIndex,
  normalizeReviewAuthor,
  resolveNamedStyleBinding,
  resolveEffectiveCellStyle,
} from "./01-cell-styles-and-themes.js";
import {
  $v,
  stripFormulaStringLiterals,
  coerceImportedValue,
} from "./03-defined-names.js";
import {
  buildSlicerItemKey,
  bucketPivotDateValue,
  tableMutationLimits,
  computePivotTable,
  findPivotDefinition,
  findRowOrigin,
  assertUniqueTableNames,
  computePivotRefreshFingerprint,
  findSourceFieldIndex,
  filterPivotSourceRows,
  buildPivotStatusRecord,
  validatePivotFieldReferences,
} from "./04-pivot-tables-and-pictures.js";
import {
  assertCalculatedBeforeFilter,
  normalizeNameForComparison,
  filterEvaluationLimits,
  savedFilterViewHelpText,
  matchesTextFilter,
  normalizeSavedFilterViewName,
  assertSavedFilterViewsEditable,
  isValidFilterViewText,
  resolveDatePolicyContext,
  textFilterOperators,
  normalizeComparableText,
  findSavedFilterView,
  comparableCellValueKey,
  isStorableCellValue,
  validateSavedFilterViewMetadata,
  savedFilterViewLimits,
} from "./05-formula-engine-and-name-operations.js";
import {
  verifyPrintSnapshotGuard,
  nonEmptyLabelSchema,
  validatePrintPageBreaks,
  headerFooterTokenPattern,
  resolvePageDimensions,
  reportOutputLimits,
  hashRecordExcludingHash,
  exportHeaderFooterTemplate,
  columnIndexToLetters,
  printRenderLimits,
  shiftFormulaReferences,
} from "./06-data-model-and-analysis.js";
import {
  doesRangeOverlap,
  parsePrintTitles,
  formatAxisIntervalRef,
  enumerateSelectionCells,
  attachReviewThreadsToWorkbook,
  validateHyperlinkRecord,
  applyChartDrawingLayout,
  _t,
  parseRangeAddress,
  normalizeCellAddress,
} from "./07-conditional-rules-worksheet-features.js";
import {
  measureRangeArea,
  normalizeTableDefinitions,
  Cn,
  Dp,
  resolveSourceRange,
  rangesIntersect,
  toA1RangeAddress,
  collectRangeValues,
  normalizeSourceReference,
  buildTableTotalCell,
  filterSourceRows,
  validateColumnName,
  normalizePivotDefinition,
  validateDrawingObjects,
} from "./08-table-pivot-query-definitions.js";
import {
  validateObjectReportRequest,
  importedCacheEvidenceSchema,
  calculationReportSchema,
  calculateWorkbookWithSettings,
} from "./09-calculation-engine.js";

function parseReportSnapshot(e) {
  if (new TextEncoder().encode(JSON.stringify(e)).length > 4e6)
    throw Error("A report snapshot exceeds 4 MB.");
  const r = calculationReportSchema.parse(e),
    rangeAddress = parseRangeAddress(r.request.range);
  if (
    (validateObjectReportRequest(r.request),
    r.hash !== hashRecordExcludingHash(r))
  )
    throw Error("Report snapshot integrity check failed.");
  const n = Array.from(
    { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
    (s, d) => rangeAddress.firstCol + d,
  );
  if (
    fingerprintCanonicalValue(n) !== fingerprintCanonicalValue(r.columns) ||
    r.widths.length !== n.length
  )
    throw Error("Invalid report column layout.");
  const o = r.rows.map((row) => row.row),
    a = [...o, ...r.excludedRows].sort((s, d) => s - d);
  if (
    r.rows.length * n.length > 2e3 ||
    a.length !== rangeAddress.lastRow - rangeAddress.firstRow + 1 ||
    a.some((s, d) => s !== rangeAddress.firstRow + d + 1) ||
    o.some((s, d) => d > 0 && s <= o[d - 1])
  )
    throw Error("Invalid report row accounting.");
  for (const row of r.rows) {
    if (row.cells.length !== n.length)
      throw Error("Invalid report cell layout.");
    row.cells.forEach((cell, u) => {
      if (cell.address !== `${columnIndexToLetters(n[u])}${row.row}`)
        throw Error("Invalid report cell address.");
      cell.style &&= validateCellStyle(cell.style);
    });
  }
  if (r.object) {
    const s = {
      id: r.request.sheetId,
      name: r.sheetName,
      rowCount: 1048576,
      colCount: 16384,
    };
    (validateDrawingObjects([r.object], s), (r.object = s.objects[0]));
  }
  if (
    (r.request.kind === "object") != !!r.object ||
    (r.object && r.object.id !== r.request.objectId)
  )
    throw Error("Invalid report drawing identity.");
  if (r.hash !== hashRecordExcludingHash(r))
    throw Error("Report metadata is not canonical.");
  return r;
}
function buildSavedReportEnvelope(workbook, r) {
  if (workbook.reportOutputs !== undefined) {
    const t = array(
      object({
        id: nonEmptyLabelSchema,
        version: number().int().min(1).max(1e6),
        snapshot: unknown(),
      }).strict(),
    )
      .max(20)
      .parse(workbook.reportOutputs);
    if (new Set(t.map((n) => n.id)).size !== t.length)
      throw Error("Duplicate report output IDs.");
    if (
      new TextEncoder().encode(JSON.stringify(t)).length >
      reportOutputLimits.metadataBytes
    )
      throw Error("Saved reports exceed 16 MB.");
    r.reportOutputs = t.map((n) => ({
      ...n,
      snapshot: parseReportSnapshot(n.snapshot),
    }));
  }
  if (workbook.embeddedObjectInventories !== undefined) {
    const t = array(importedCacheEvidenceSchema)
      .max(20)
      .parse(workbook.embeddedObjectInventories);
    if (
      new Set(t.map((n) => n.id)).size !== t.length ||
      JSON.stringify(t).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    r.embeddedObjectInventories = t;
  }
}
function validatePageSetup(e, sheet) {
  if (e === undefined) return;
  if (!e || typeof e != "object" || Array.isArray(e))
    throw new Error("Invalid page layout.");
  const t = e,
    margins = t.margins,
    o = (b, y, w) => {
      if (typeof b != "number" || !Number.isFinite(b) || b < y || b > w)
        throw new Error("Page setup contains an out-of-range number.");
      return b;
    };
  if (
    !["portrait", "landscape"].includes(String(t.orientation)) ||
    !["letter", "a4", "legal"].includes(String(t.paper)) ||
    !margins ||
    typeof margins != "object"
  )
    throw new Error("Invalid paper, orientation, or margins.");
  if (
    t.pageOrder !== undefined &&
    t.pageOrder !== "overThenDown" &&
    t.pageOrder !== "downThenOver"
  )
    throw Error(
      "Unsupported print page order. Choose Over, then down or Down, then over.",
    );
  const a = Object.fromEntries(
      ["top", "bottom", "left", "right", "header", "footer"].map((b) => [
        b,
        o(margins[b], 0, 3),
      ]),
    ),
    s =
      t.paper === "a4"
        ? [8.2677, 11.6929]
        : t.paper === "legal"
          ? [8.5, 14]
          : [8.5, 11],
    [d, u] = t.orientation === "landscape" ? s.reverse() : s;
  if (a.left + a.right >= d - 0.25 || a.top + a.bottom >= u - 0.25)
    throw new Error("Margins leave no printable page area.");
  if (typeof t.printArea != "string") throw new Error("Enter a print area.");
  const rangeAddress = parseRangeAddress(t.printArea);
  if (
    rangeAddress.lastRow >= sheet.rowCount ||
    rangeAddress.lastCol >= sheet.colCount
  )
    throw new Error("Print area exceeds the worksheet.");
  if (
    ["fitToPage", "gridlines", "headings"].some((b) => typeof t[b] != "boolean")
  )
    throw new Error("Invalid print option.");
  if (
    ["header", "footer"].some(
      (b) => typeof t[b] != "string" || String(t[b]).length > 255,
    )
  )
    throw new Error("Headers and footers must be at most 255 characters.");
  const fitToWidth = o(t.fitToWidth, 0, 10),
    fitToHeight = o(t.fitToHeight, 0, 10),
    scale = o(t.scale, 10, 400);
  if (![fitToWidth, fitToHeight, scale].every(Number.isInteger))
    throw new Error("Page fitting and scale must be whole numbers.");
  const printTitles = parsePrintTitles(t.repeatRows, "row", sheet),
    m = parsePrintTitles(t.repeatColumns, "column", sheet),
    rowBreaks = validatePrintPageBreaks(t.rowBreaks, sheet.rowCount),
    columnBreaks = validatePrintPageBreaks(t.columnBreaks, sheet.colCount);
  sheet.pageLayout = {
    orientation: t.orientation,
    paper: t.paper,
    margins: a,
    printArea: t.printArea,
    fitToPage: t.fitToPage,
    fitToWidth: fitToWidth,
    fitToHeight: fitToHeight,
    scale: scale,
    gridlines: t.gridlines,
    headings: t.headings,
    header: t.header,
    footer: t.footer,
    ...(t.pageOrder !== undefined ? { pageOrder: t.pageOrder } : {}),
    ...(printTitles
      ? { repeatRows: formatAxisIntervalRef(printTitles, "row") }
      : {}),
    ...(m ? { repeatColumns: formatAxisIntervalRef(m, "column") } : {}),
    ...(rowBreaks ? { rowBreaks: rowBreaks } : {}),
    ...(columnBreaks ? { columnBreaks: columnBreaks } : {}),
  };
}
function zU(sheet, pageLayout) {
  const t = { ...sheet };
  (validatePageSetup(pageLayout, t), (pageLayout = t.pageLayout));
  const rangeAddress = parseRangeAddress(pageLayout.printArea);
  if (
    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
      (rangeAddress.lastCol - rangeAddress.firstCol + 1) >
    printRenderLimits.areaCells
  )
    throw Error(
      "Print preview supports at most 5,000 cells per print area. Choose a smaller print area.",
    );
  const pageDimensions = resolvePageDimensions(pageLayout),
    s =
      pageDimensions.width -
      (pageLayout.margins.left + pageLayout.margins.right) * 96,
    d =
      pageDimensions.height -
      (pageLayout.margins.top + pageLayout.margins.bottom) * 96,
    printTitles = parsePrintTitles(pageLayout.repeatRows, "row", sheet),
    l = parsePrintTitles(pageLayout.repeatColumns, "column", sheet);
  for (const j of sheet.merges ?? []) {
    const z = parseRangeAddress(j);
    for (const [V, D, B] of [
      [printTitles, z.firstRow, z.lastRow],
      [l, z.firstCol, z.lastCol],
    ])
      if (V && D <= V.end && B >= V.start && (D < V.start || B > V.end))
        throw Error(
          "A merged area crosses a repeated-title boundary. Unmerge it or adjust the print titles first.",
        );
  }
  for (const j of sheet.objects ?? []) {
    const z = _t(j.anchor),
      V = applyChartDrawingLayout(sheet, j);
    for (const [D, B, _, C, T] of [
      [printTitles, "row", z.row, V.offsetY, V.height],
      [l, "column", z.col, V.offsetX, V.width],
    ]) {
      if (!D) continue;
      const axisVisibilityIndex = createAxisVisibilityIndex(
        B === "row" ? sheet.rowCount : sheet.colCount,
        cn(sheet, B),
      );
      if (axisVisibilityIndex.isHidden(_)) continue;
      const R = (G) =>
          B === "row" ? rowHeightAt(sheet, G) : columnWidthAt(sheet, G),
        L = (G) => {
          let H = _,
            Q = 0;
          for (; G < 0;) {
            const X = axisVisibilityIndex.next(H, -1);
            if (X === null) return H;
            if (((H = X), (G += R(H)), ++Q > 1e3))
              throw Error(
                "Drawing print inspection exceeds its geometry budget.",
              );
          }
          for (; G >= R(H);) {
            G -= R(H);
            const X = axisVisibilityIndex.next(H, 1);
            if (X === null) return H;
            if (((H = X), ++Q > 1e3))
              throw Error(
                "Drawing print inspection exceeds its geometry budget.",
              );
          }
          return H;
        },
        $ = L(C),
        U = L(C + T - 1e-4);
      if ($ <= D.end && U >= D.start && ($ < D.start || U > D.end))
        throw Error(
          "A drawing crosses a repeated-title boundary. Move or resize the drawing, or adjust the print titles first.",
        );
    }
  }
  const c = (j, z, V) => {
      const axisVisibilityIndex = createAxisVisibilityIndex(
          j === "row" ? sheet.rowCount : sheet.colCount,
          cn(sheet, j),
        ),
        B = [];
      for (let _ = z; _ <= V; _++) axisVisibilityIndex.isHidden(_) || B.push(_);
      return B;
    },
    h = c("row", rangeAddress.firstRow, rangeAddress.lastRow),
    f = c("column", rangeAddress.firstCol, rangeAddress.lastCol);
  if (!h.length || !f.length)
    throw Error(
      "The print area has no visible cells. Unhide its rows or columns, or choose another print area.",
    );
  const p = printTitles ? c("row", printTitles.start, printTitles.end) : [],
    m = l ? c("column", l.start, l.end) : [],
    v = (j, z, V, D, B) => {
      const _ = new Set(z),
        C = j.filter((U) => !_.has(U)),
        T = z.reduce((U, G) => U + D(G), 0);
      if (T > V + 1e-7 || (C.length && T >= V - 1e-7))
        throw Error(
          "Repeated titles leave no space for worksheet data. Reduce the titles, margins or scale.",
        );
      if (!C.length) return [z];
      const N = [];
      let R = [],
        L = T,
        $;
      for (const U of C) {
        const G = D(U),
          H = $ !== undefined && B.some((Q) => Q > $ && Q <= U);
        if (
          (R.length &&
            (L + G > V + 1e-7 || H) &&
            (N.push([...z, ...R]), (R = []), (L = T)),
          z.length && L + G > V + 1e-7 && !R.length)
        )
          throw Error(
            "A data row or column cannot fit beside the repeated titles. Reduce the scale or title dimensions.",
          );
        (R.push(U), (L += G), ($ = U));
      }
      return (R.length && N.push([...z, ...R]), N);
    },
    g = (j) => ({
      rows: v(
        h,
        p,
        d / j - (pageLayout.headings ? 22 : 0),
        (z) => rowHeightAt(sheet, z),
        pageLayout.fitToPage ? [] : (pageLayout.rowBreaks ?? []),
      ),
      columns: v(
        f,
        m,
        s / j - (pageLayout.headings ? 32 : 0),
        (z) => columnWidthAt(sheet, z),
        pageLayout.fitToPage ? [] : (pageLayout.columnBreaks ?? []),
      ),
    });
  let scale = pageLayout.fitToPage ? 1 : pageLayout.scale / 100;
  if (
    pageLayout.fitToPage &&
    (pageLayout.fitToWidth || pageLayout.fitToHeight)
  ) {
    const j = (z) => {
      try {
        const V = g(z);
        return (
          (!pageLayout.fitToWidth ||
            V.columns.length <= pageLayout.fitToWidth) &&
          (!pageLayout.fitToHeight || V.rows.length <= pageLayout.fitToHeight)
        );
      } catch {
        return false;
      }
    };
    if (!j(1)) {
      let z = 0,
        V = 1;
      for (let D = 0; D < 40; D++) {
        const B = (z + V) / 2;
        j(B) ? (z = B) : (V = B);
      }
      scale = z;
    }
  }
  if (!Number.isFinite(scale) || scale <= 0)
    throw Error(
      "The print content cannot fit the selected paper. Reduce the print area or titles.",
    );
  const y = g(scale);
  if (y.rows.length * y.columns.length > printRenderLimits.pages)
    throw Error(
      "This layout exceeds 100 pages. Reduce the print area or fit it to fewer pages.",
    );
  if (
    y.rows.reduce((j, z) => j + z.length, 0) *
      y.columns.reduce((j, z) => j + z.length, 0) >
    printRenderLimits.renderedCells
  )
    throw Error(
      "Repeated print content exceeds 100,000 rendered cells. Reduce titles, pages or the print area.",
    );
  const E = !!(printTitles || l),
    A = (rows, columns) => ({
      firstRow: Math.min(...rows),
      lastRow: Math.max(...rows),
      firstCol: Math.min(...columns),
      lastCol: Math.max(...columns),
      ...(E ? { rows: rows, columns: columns } : {}),
    }),
    pages =
      pageLayout.pageOrder === "downThenOver"
        ? y.columns.flatMap((column) => y.rows.map((row) => A(row, column)))
        : y.rows.flatMap((row) => y.columns.map((column) => A(row, column))),
    notices =
      pageLayout.fitToPage &&
      (pageLayout.rowBreaks?.length ?? 0) +
        (pageLayout.columnBreaks?.length ?? 0)
        ? [
            "Fit to pages ignores saved manual page breaks. Turn it off to use the breaks.",
          ]
        : [];
  return {
    pages: pages,
    scale: scale,
    ...(notices.length ? { notices: notices } : {}),
  };
}
function updateSheetPageSetup(e, sheet, workbookInput) {
  if (!sheet.pageLayout) return [];
  const n = { ...sheet };
  validatePageSetup(sheet.pageLayout, n);
  const pageLayout = n.pageLayout,
    oddHeader = exportHeaderFooterTemplate(pageLayout.header, workbookInput),
    oddFooter = exportHeaderFooterTemplate(pageLayout.footer, workbookInput),
    d = [pageLayout.header, pageLayout.footer].some((u) =>
      [...u.matchAll(headerFooterTokenPattern)].some((l) => !!l[1]),
    );
  return (
    (e.pageSetup = {
      ...e.pageSetup,
      orientation: pageLayout.orientation,
      paperSize:
        pageLayout.paper === "a4" ? 9 : pageLayout.paper === "legal" ? 5 : 1,
      margins: pageLayout.margins,
      printArea: pageLayout.printArea,
      fitToPage: pageLayout.fitToPage,
      fitToWidth: pageLayout.fitToWidth,
      fitToHeight: pageLayout.fitToHeight,
      scale: pageLayout.scale,
      showGridLines: pageLayout.gridlines,
      showRowColHeaders: pageLayout.headings,
      pageOrder: pageLayout.pageOrder ?? "overThenDown",
      ...(pageLayout.repeatRows
        ? { printTitlesRow: pageLayout.repeatRows }
        : {}),
      ...(pageLayout.repeatColumns
        ? { printTitlesColumn: pageLayout.repeatColumns }
        : {}),
    }),
    (e.rowBreaks = (pageLayout.rowBreaks ?? []).map((u) => ({
      id: u,
      max: 16383,
      man: 1,
    }))),
    (e.headerFooter = { oddHeader: oddHeader, oddFooter: oddFooter }),
    d
      ? [
          "Print workbook/engagement/period/revision fields are exported as literal text from this workbook snapshot. Native JSON retains the live identity tokens; XLSX does not update them after later edits.",
        ]
      : []
  );
}
function findOccupiedRangeBounds(e, r = true) {
  let t = null;
  for (const [n, cell] of Object.entries(e.cells)) {
    const a = cell.raw !== null && cell.raw !== "",
      s =
        r &&
        (cell.numberFormat ||
          cell.hyperlink ||
          cell.note ||
          cell.tickmark ||
          cell.numFmt ||
          cell.format ||
          Object.keys(cell.style ?? {}).length);
    if (!a && !s) continue;
    const { row, col } = _t(n);
    t
      ? ((t.firstRow = Math.min(t.firstRow, row)),
        (t.lastRow = Math.max(t.lastRow, row)),
        (t.firstCol = Math.min(t.firstCol, col)),
        (t.lastCol = Math.max(t.lastCol, col)))
      : (t = { firstRow: row, lastRow: row, firstCol: col, lastCol: col });
  }
  return t;
}
function formatOccupiedRangeAddress(e, r = true) {
  const occupiedRangeBounds = findOccupiedRangeBounds(e, r);
  if (!occupiedRangeBounds) return "A1";
  const n = `${columnIndexToLetters(occupiedRangeBounds.firstCol)}${occupiedRangeBounds.firstRow + 1}`,
    o = `${columnIndexToLetters(occupiedRangeBounds.lastCol)}${occupiedRangeBounds.lastRow + 1}`;
  return n === o ? n : `${n}:${o}`;
}
function createDefaultPageSetup(sheetOrThrow) {
  return {
    orientation: "portrait",
    paper: "letter",
    margins: {
      top: 0.75,
      bottom: 0.75,
      left: 0.7,
      right: 0.7,
      header: 0.3,
      footer: 0.3,
    },
    printArea: sheetOrThrow ? formatOccupiedRangeAddress(sheetOrThrow) : "A1",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    scale: 100,
    gridlines: true,
    headings: false,
    pageOrder: "overThenDown",
    header: "",
    footer: "",
  };
}
function assertRangeCellBudget(e) {
  const { firstRow, lastRow, firstCol, lastCol } = parseRangeAddress(e);
  if (
    (lastRow - firstRow + 1) * (lastCol - firstCol + 1) >
    workbookLimits.maxCells
  )
    return throwWithMessage(
      "Range is too large. Select at most 100,000 cells.",
    );
  const a = [];
  for (let s = firstRow; s <= lastRow; s++)
    for (let d = firstCol; d <= lastCol; d++)
      a.push(`${columnIndexToLetters(d)}${s + 1}`);
  return a;
}
function validateSelectionBounds(sheetOrThrow, activeCellAddress) {
  const t = assertRangeCellBudget(activeCellAddress);
  if (
    t.some((n) => {
      const { row, col } = _t(n);
      return row >= sheetOrThrow.rowCount || col >= sheetOrThrow.colCount;
    })
  )
    throw new Error("Selection exceeds worksheet dimensions.");
  return t;
}
function validateRemoveDuplicatesRequest(
  workbook,
  sheetId,
  t,
  n,
  checked = true,
) {
  const a = workbook.sheets.find((sheet) => sheet.id === sheetId);
  assertSheetUnprotected(a);
  const s = assertRangeCellBudget(t);
  s.length > 500 &&
    failWithMessage("Select at most 500 cells to remove duplicates.");
  const rangeAddress = parseRangeAddress(t);
  (!n.length ||
    n.some((m) => m < rangeAddress.firstCol || m > rangeAddress.lastCol)) &&
    failWithMessage("Choose one or more columns inside the selected range.");
  for (const m of a.tables ?? []) {
    const v = parseRangeAddress(m.range);
    doesRangeOverlap(t, m.range) &&
      ((rangeAddress.firstCol > v.firstCol ||
        rangeAddress.lastCol < v.lastCol) &&
        failWithMessage(
          "Include all table columns when removing duplicate table rows.",
        ),
      m.showTotals &&
        rangeAddress.firstRow <= v.lastRow &&
        rangeAddress.lastRow >= v.lastRow &&
        failWithMessage("Select data rows above the table totals row."),
      !checked &&
        rangeAddress.firstRow <= v.firstRow &&
        rangeAddress.lastRow >= v.firstRow &&
        failWithMessage("Keep the table header row selected as headers."));
  }
  (a.merges ?? []).some((m) => doesRangeOverlap(t, m)) &&
    failWithMessage("Unmerge selected cells before removing duplicates.");
  const workbookWithSettings = calculateWorkbookWithSettings(workbook),
    l = new Set(),
    c = [];
  let removed = 0;
  for (
    let firstRow = rangeAddress.firstRow;
    firstRow <= rangeAddress.lastRow;
    firstRow++
  ) {
    const v = JSON.stringify(
      n.map((g) =>
        workbookWithSettings(
          sheetId,
          `${columnIndexToLetters(g)}${firstRow + 1}`,
        ),
      ),
    );
    if (firstRow !== rangeAddress.firstRow || !checked) {
      if (l.has(v)) {
        removed++;
        continue;
      }
      l.add(v);
    }
    c.push(
      Array.from(
        { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
        (g, b) => {
          const cell = structuredClone(
            a.cells[
              `${columnIndexToLetters(rangeAddress.firstCol + b)}${firstRow + 1}`
            ] ?? {
              raw: null,
            },
          );
          return (
            typeof cell.raw == "string" &&
              cell.raw.startsWith("=") &&
              !cell.literal &&
              (cell.raw = shiftFormulaReferences(
                cell.raw,
                rangeAddress.firstRow + c.length - firstRow,
                0,
              )),
            cell
          );
        },
      ),
    );
  }
  const cells = { ...a.cells };
  for (const m of s) {
    const v = _t(m);
    cells[m] = c[v.row - rangeAddress.firstRow]?.[
      v.col - rangeAddress.firstCol
    ] ?? {
      ...cells[m],
      raw: null,
    };
  }
  return {
    workbook: {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === sheetId ? { ...sheet, cells: cells } : sheet,
      ),
    },
    removed: removed,
  };
}
function findSheetOrThrow(workbook, r) {
  return (
    workbook.sheets.find((sheet) => sheet.id === r) ??
    throwWithMessage(`Sheet not found: ${r}`)
  );
}
function applyNamedStyleChange(workbook, request, t = {}) {
  t.readOnly && failWithMessage("This workbook is read-only.");
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells))
      assertNamedStyleBindingCurrent(workbook, cell.style);
  const styleCatalog = workbook.styleCatalog,
    o = cloneStyleValue(styleCatalog ?? createDefaultStyleCatalog(workbook.id));
  let a = cloneStyleValue(request);
  const changes = [],
    warnings = [],
    u = (v) =>
      o.styles.find((style) => style.id === v) ??
      failWithMessage("This named style no longer exists.");
  function l(sheetId, address, b) {
    const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
      w = sheetOrThrow.cells[address];
    if (canonicalStyleKey(w) !== canonicalStyleKey(b)) {
      if (
        (assertSheetUnprotected(sheetOrThrow),
        hasBorderEdgeChanges(w?.style, b.style))
      ) {
        const A = _t(address);
        sheetOrThrow.merges?.some((merge) => {
          const rangeAddress = parseRangeAddress(merge);
          return (
            A.row >= rangeAddress.firstRow &&
            A.row <= rangeAddress.lastRow &&
            A.col >= rangeAddress.firstCol &&
            A.col <= rangeAddress.lastCol
          );
        }) &&
          failWithMessage(
            "Unmerge the intersecting cells before changing named-style border edges.",
          );
      }
      (changes.push({
        sheetId: sheetId,
        address: address,
        ...(w ? { before: cloneStyleValue(w) } : {}),
        after: cloneStyleValue(b),
      }),
        changes.length > 500 &&
          failWithMessage(
            "This action changes more than 500 cells. Use smaller selections or fewer bound cells.",
          ));
    }
  }
  function c(v) {
    for (const sheet of workbook.sheets)
      for (const [y, w] of Object.entries(sheet.cells))
        w.style?.binding &&
          (!v || w.style.binding.styleId === v) &&
          (assertSheetUnprotected(sheet),
          l(sheet.id, y, {
            ...w,
            style: {
              ...w.style,
              binding: resolveNamedStyleBinding(o, w.style.binding),
            },
          }));
  }
  switch (a.kind) {
    case "setTheme":
      ((o.theme = validateWorkbookTheme(a.theme)),
        delete o.missingTheme,
        (a.theme = cloneStyleValue(o.theme)),
        c());
      break;
    case "createStyle": {
      const g = validateBoundedStyleText(a.name, "Style name"),
        b =
          a.styleId === undefined
            ? `style-${crypto.randomUUID()}`
            : validateBoundedStyleText(a.styleId, "Style ID", 100),
        definition = validateNamedStyleDefinition(a.definition);
      (o.styles.push({ id: b, name: g, definition: definition }),
        (a = { ...a, name: g, styleId: b, definition: definition }),
        c(b));
      break;
    }
    case "updateStyle": {
      const g = u(a.styleId);
      ((g.definition = validateNamedStyleDefinition(a.definition)),
        (a.definition = cloneStyleValue(g.definition)),
        c(g.id));
      break;
    }
    case "renameStyle": {
      const g = u(a.styleId);
      ((g.name = validateBoundedStyleText(a.name, "Style name")),
        (a.name = g.name),
        c(g.id));
      break;
    }
    case "deleteStyle":
      u(a.styleId);
      const styleId = a.styleId;
      ((o.styles = o.styles.filter((style) => style.id !== styleId)),
        c(a.styleId),
        warnings.push(
          "Cells using this style will show default formatting plus their direct overrides. Detach the style first to retain its current appearance.",
        ));
      break;
    case "applyStyle":
    case "reapplyStyle":
    case "detachStyle": {
      const sheetOrThrow = findSheetOrThrow(workbook, a.sheetId);
      (assertSheetUnprotected(sheetOrThrow),
        a.kind !== "detachStyle" && u(a.styleId));
      for (const b of enumerateSelectionCells(a.ranges)) {
        const y = sheetOrThrow.cells[b] ?? { raw: null };
        if (a.kind === "detachStyle") {
          if (!y.style?.binding) continue;
          const effectiveCellStyle = resolveEffectiveCellStyle(y.style),
            E = { ...y, style: effectiveCellStyle };
          (Object.keys(effectiveCellStyle).length || delete E.style,
            l(sheetOrThrow.id, b, E));
        } else {
          const namedStyleBinding = resolveNamedStyleBinding(o, {
            namespace: o.namespace,
            styleId: a.styleId,
          });
          l(sheetOrThrow.id, b, {
            ...y,
            style: {
              ...(a.kind === "applyStyle" ? y.style : {}),
              binding: namedStyleBinding,
            },
          });
        }
      }
      break;
    }
    default:
      failWithMessage("Choose a supported theme or named style action.");
  }
  const h = validateWorkbookStyleCatalog(o),
    catalogAfter = a.kind === "detachStyle" && !styleCatalog ? undefined : h,
    p = {
      setTheme: "Change workbook theme",
      createStyle: "Create named style",
      updateStyle: "Update named style",
      renameStyle: "Rename named style",
      deleteStyle: "Delete named style",
      applyStyle: "Apply named style",
      reapplyStyle: "Reapply named style and clear direct style overrides",
      detachStyle: "Detach named style and keep appearance",
    };
  return {
    version: 1,
    snapshot: JSON.stringify(workbook),
    request: a,
    ...(styleCatalog ? { catalogBefore: cloneStyleValue(styleCatalog) } : {}),
    ...(catalogAfter ? { catalogAfter: catalogAfter } : {}),
    changes: changes,
    count: changes.length,
    summary: p[a.kind],
    warnings: warnings,
  };
}
function assertSheetWritable(workbook, sheetId, t) {
  (t.readOnly && throwWithMessage("This workbook is read-only."),
    assertSheetUnprotected(findSheetOrThrow(workbook, sheetId)));
}
function nQ(
  workbook,
  sheetId,
  selectionAnchorAddress,
  n,
  o = "Reviewer",
  a = {},
) {
  assertSheetWritable(workbook, sheetId, a);
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    reviewTimestamp = resolveReviewTimestamp(a),
    reviewAuthor = normalizeReviewAuthor(o),
    l = {
      id: crypto.randomUUID(),
      anchor: {
        sheetId: sheetId,
        sheetName: sheetOrThrow.name,
        address: normalizeCellAddress(selectionAnchorAddress),
      },
      createdAt: reviewTimestamp,
      updatedAt: reviewTimestamp,
      messages: [
        {
          id: crypto.randomUUID(),
          author: reviewAuthor,
          text: normalizeBoundedText(
            n,
            "Comment",
            reviewThreadLimits.textLength,
            true,
          ),
          createdAt: reviewTimestamp,
        },
      ],
    };
  return attachReviewThreadsToWorkbook(workbook, [
    ...(workbook.reviewThreads ?? []),
    l,
  ]);
}
function findReviewThread(workbook, r, t) {
  t.readOnly && throwWithMessage("This workbook is read-only.");
  const n = workbook.reviewThreads?.find(
    (reviewThread) => reviewThread.id === r,
  );
  return n
    ? (n.orphaned &&
        throwWithMessage(
          "This discussion is retained for review because its target was deleted.",
        ),
      assertSheetWritable(workbook, n.anchor.sheetId, t),
      n)
    : throwWithMessage(
        "The discussion could not be found. Reopen the review tools.",
      );
}
function addReviewThreadReply(workbook, r, t, n = "Reviewer", o = {}) {
  const reviewThread = findReviewThread(workbook, r, o);
  reviewThread.resolved &&
    throwWithMessage("Reopen the discussion before adding a reply.");
  const reviewTimestamp = resolveReviewTimestamp(o, reviewThread.updatedAt),
    d = {
      ...reviewThread,
      updatedAt: reviewTimestamp,
      messages: [
        ...reviewThread.messages,
        {
          id: crypto.randomUUID(),
          author: normalizeReviewAuthor(n),
          text: normalizeBoundedText(
            t,
            "Reply",
            reviewThreadLimits.textLength,
            true,
          ),
          createdAt: reviewTimestamp,
        },
      ],
    };
  return attachReviewThreadsToWorkbook(
    workbook,
    workbook.reviewThreads.map((u) => (u.id === r ? d : u)),
  );
}
function setReviewThreadResolved(workbook, r, t, n = "Reviewer", o = {}) {
  typeof t != "boolean" &&
    throwWithMessage("Choose whether to resolve or reopen the discussion.");
  const reviewThread = findReviewThread(workbook, r, o);
  if (!!reviewThread.resolved === t) return workbook;
  const reviewTimestamp = resolveReviewTimestamp(o, reviewThread.updatedAt),
    d = { ...reviewThread, updatedAt: reviewTimestamp };
  return (
    t
      ? (d.resolved = {
          author: normalizeReviewAuthor(n),
          createdAt: reviewTimestamp,
        })
      : (normalizeReviewAuthor(n), delete d.resolved),
    attachReviewThreadsToWorkbook(
      workbook,
      workbook.reviewThreads.map((u) => (u.id === r ? d : u)),
    )
  );
}
function setCellHyperlink(
  workbook,
  sheetId,
  selectionAnchorAddress,
  hyperlink,
  o = {},
) {
  assertSheetWritable(workbook, sheetId, o);
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    cellAddress = normalizeCellAddress(selectionAnchorAddress),
    d = sheetOrThrow.cells[cellAddress],
    u = _t(cellAddress);
  sheetOrThrow.merges?.some((merge) => {
    const rangeAddress = parseRangeAddress(merge);
    return (
      u.row >= rangeAddress.firstRow &&
      u.row <= rangeAddress.lastRow &&
      u.col >= rangeAddress.firstCol &&
      u.col <= rangeAddress.lastCol &&
      (u.row !== rangeAddress.firstRow || u.col !== rangeAddress.firstCol)
    );
  }) &&
    throwWithMessage(
      "Choose the merged area's first cell or unmerge it before changing its link.",
    );
  const l = hyperlink === null ? undefined : validateHyperlinkRecord(hyperlink);
  if (
    (l?.kind === "internal" && findSheetOrThrow(workbook, l.sheetId),
    JSON.stringify(d?.hyperlink) === JSON.stringify(l))
  )
    return workbook;
  !d &&
    workbook.sheets.reduce((p, m) => p + Object.keys(m.cells).length, 0) >=
      workbookLimits.maxCells &&
    throwWithMessage(
      "Adding this link would exceed the workbook populated-cell limit.",
    );
  const cell = { ...(d ?? { raw: null }) };
  l ? (cell.hyperlink = l) : delete cell.hyperlink;
  const cells = { ...sheetOrThrow.cells, [cellAddress]: cell };
  return (
    cell.raw === null &&
      Object.keys(cell).length === 1 &&
      delete cells[cellAddress],
    {
      ...workbook,
      sheets: workbook.sheets.map((sheet) =>
        sheet.id === sheetId ? { ...sheet, cells: cells } : sheet,
      ),
    }
  );
}
function collectHyperlinkEntries(workbook, r, sheetId) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  let o = false;
  const cells = Object.fromEntries(
    Object.entries(sheetOrThrow.cells).map(([s, d]) =>
      d.hyperlink?.kind !== "internal" || d.hyperlink.sheetId !== r
        ? [s, d]
        : ((o = true),
          [s, { ...d, hyperlink: { ...d.hyperlink, sheetId: sheetId } }]),
    ),
  );
  return o
    ? {
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          sheet.id === sheetId ? { ...sheet, cells: cells } : sheet,
        ),
      }
    : workbook;
}
function applyAutoSum(activeWorkbook, sheetId, activeCellAddress) {
  const sheetOrThrow = findSheetOrThrow(activeWorkbook, sheetId),
    o = validateSelectionBounds(sheetOrThrow, activeCellAddress);
  if (o.length !== 1) throw new Error("Select one empty total cell.");
  activeCellAddress = o[0];
  const a = _t(activeCellAddress),
    cell = sheetOrThrow.cells[activeCellAddress];
  if (cell?.raw !== undefined && cell.raw !== null && cell.raw !== "")
    throw new Error("Select an empty cell for AutoSum.");
  const d = (h, f) => {
    if (h < 0 || f < 0) return false;
    const p = sheetOrThrow.cells[`${columnIndexToLetters(f)}${h + 1}`];
    return (
      typeof p?.raw == "number" ||
      (typeof p?.raw == "string" && p.raw.startsWith("=") && !p.literal)
    );
  };
  let u, l;
  if (d(a.row - 1, a.col)) {
    let h = a.row - 1;
    for (; d(h - 1, a.col);) h--;
    ((u = `${columnIndexToLetters(a.col)}${h + 1}`),
      (l = `${columnIndexToLetters(a.col)}${a.row}`));
  } else if (d(a.row, a.col - 1)) {
    let h = a.col - 1;
    for (; d(a.row, h - 1);) h--;
    ((u = `${columnIndexToLetters(h)}${a.row + 1}`),
      (l = `${columnIndexToLetters(a.col - 1)}${a.row + 1}`));
  } else
    throw new Error(
      "AutoSum needs adjacent numbers or formulas above or to the left.",
    );
  const c = cloneCellRecord(cell);
  return (
    (c.raw = `=SUM(${u}:${l})`),
    delete c.literal,
    { sheetId: sheetId, address: activeCellAddress, cell: c }
  );
}
function setTableColumnDataType(workbook, sheetId, tableId, columnId) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    a = sheetOrThrow.tables?.find((table) => table.id === tableId);
  if (!a) throw Error("Table not found.");
  const s = a.columns.findIndex((column) => column.id === columnId);
  if (s < 0) throw Error("Select an existing table column.");
  const rangeAddress = parseRangeAddress(a.range),
    total = Math.max(
      0,
      rangeAddress.lastRow - rangeAddress.firstRow - Number(a.showTotals),
    ),
    workbookWithSettings = calculateWorkbookWithSettings(workbook),
    exceptions = [];
  for (let f = 1; f <= Math.min(total, 500); f++) {
    const address = `${columnIndexToLetters(rangeAddress.firstCol + s)}${rangeAddress.firstRow + f + 1}`,
      m = workbookWithSettings(sheetId, address);
    coerceImportedValue(m, a.columns[s].dataType) ||
      exceptions.push({ address: address, value: m });
  }
  return {
    dataType: a.columns[s].dataType ?? "any",
    checked: Math.min(total, 500),
    total: total,
    truncated: total > 500,
    exceptions: exceptions,
  };
}
function findTableForRange(workbook, sheetId, range) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(range);
  if (measureRangeArea(rangeAddress) !== 1) return toA1RangeAddress(range);
  const a = sheetOrThrow.tables?.find((table) =>
    Dp(
      parseRangeAddress(table.range),
      rangeAddress.firstRow,
      rangeAddress.firstCol,
    ),
  );
  return a ? a.range : formatOccupiedRangeAddress(sheetOrThrow, false);
}
function suggestTableHeaderName(workbook, sheetId, range) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    sourceRange = resolveSourceRange(range, sheetOrThrow),
    a = new Set(),
    s = [];
  for (
    let firstCol = sourceRange.firstCol;
    firstCol <= sourceRange.lastCol;
    firstCol++
  ) {
    const address = Cn(sourceRange.firstRow, firstCol),
      cell = sheetOrThrow.cells[address],
      original = cell?.raw ?? null;
    let h = `Column${firstCol - sourceRange.firstCol + 1}`;
    try {
      isNullOrEmptyValue(original) ||
        (h = validateColumnName(String(original)));
    } catch {}
    let suggested = h,
      p = 2;
    for (; a.has(suggested.toLowerCase());)
      suggested = `${h.slice(0, 250)}${p++}`;
    (a.add(suggested.toLowerCase()),
      s.push({
        address: address,
        original: original,
        suggested: suggested,
        requiresResolution:
          original !== suggested ||
          (suggested.startsWith("=") && !cell?.literal),
      }));
  }
  return s;
}
function validateReviewedTableHeaders(workbook, sheetId, range, headers) {
  const o = suggestTableHeaderName(workbook, sheetId, range);
  if (headers === undefined)
    return (
      o.some((s) => s.requiresResolution) &&
        throwWorkbookError(
          "Review table headers: resolve blank, duplicate, non-text or formula headers explicitly. Original header values will be retained in native source evidence.",
        ),
      o.map((s) => s.suggested)
    );
  (!Array.isArray(headers) || headers.length !== o.length) &&
    throwWorkbookError("Provide one reviewed header for every table column.");
  const a = headers.map(validateColumnName);
  return (
    new Set(a.map((s) => s.toLowerCase())).size !== a.length &&
      throwWorkbookError(
        "Reviewed table headers must be unique, nonblank text.",
      ),
    a
  );
}
function renameTableHeaderField(workbook, r) {
  const t = [];
  for (const sheet of workbook.sheets) {
    const d = r.sheets.find((u) => u.id === sheet.id);
    if (d)
      for (const old of sheet.tables ?? []) {
        const l = d.tables?.find((table) => table.id === old.id);
        if (!l) continue;
        const rangeAddress = parseRangeAddress(old.range),
          h = parseRangeAddress(l.range),
          names = {},
          columns = l.columns.map((column, v) => {
            const g = old.columns.findIndex((k) => k.id === column.id);
            if (g < 0) return column;
            const b = old.columns[g],
              cell =
                sheet.cells[
                  Cn(rangeAddress.firstRow, rangeAddress.firstCol + g)
                ],
              w = d.cells[Cn(h.firstRow, h.firstCol + v)],
              A =
                cell?.raw !== w?.raw || cell?.literal !== w?.literal
                  ? w && !isNullOrEmptyValue(w.raw)
                    ? validateColumnName(String(w.raw))
                    : throwWorkbookError(
                        "Table headers must be nonblank text. The header edit was not applied.",
                      )
                  : column.name;
            return (
              A !== b.name &&
                Object.defineProperty(names, b.name, {
                  value: A,
                  enumerable: true,
                }),
              A === column.name ? column : { ...column, name: A }
            );
          });
        (old.name !== l.name || Object.keys(names).length) &&
          (new Set(columns.map((column) => column.name.toLowerCase())).size !==
            columns.length &&
            throwWorkbookError(
              "Table headers must be unique, nonblank text. The header edit was not applied.",
            ),
          t.push({
            sheetId: d.id,
            old: old,
            current: l,
            columns: columns,
            names: names,
          }));
      }
  }
  if (!t.length) return r;
  const n = structuredClone(r);
  for (const s of t) {
    const sheetOrThrow = findSheetOrThrow(n, s.sheetId),
      u = sheetOrThrow.tables.find((table) => table.id === s.current.id),
      l = (f) => (Object.hasOwn(s.names, f) ? s.names[f] : f);
    ((u.columns = s.columns),
      (u.totals = Object.fromEntries(
        Object.entries(u.totals).map(([f, p]) => [l(f), p]),
      )));
    const rangeAddress = parseRangeAddress(u.range);
    u.columns.forEach((column, p) => {
      const m = Cn(rangeAddress.firstRow, rangeAddress.firstCol + p);
      sheetOrThrow.cells[m] = {
        ...(sheetOrThrow.cells[m] ?? { raw: column.name }),
        raw: column.name,
        ...(column.name.startsWith("=")
          ? { literal: true }
          : { literal: undefined }),
      };
    });
    const h = (f) =>
      f.tableId === u.id ||
      (!f.tableId &&
        f.sheetId === sheetOrThrow.id &&
        toA1RangeAddress(f.range) === s.old.range);
    for (const sheet of n.sheets) {
      for (const p of sheet.tables ?? [])
        for (const column of p.columns)
          column.calculatedFormula &&= $v(
            column.calculatedFormula,
            s.old.name,
            u.name,
            s.names,
            p.id === u.id && sheet.id === sheetOrThrow.id,
          );
      for (const [p, cell] of Object.entries(sheet.cells)) {
        if (
          typeof cell.raw != "string" ||
          !cell.raw.startsWith("=") ||
          cell.literal
        )
          continue;
        const v = _t(p),
          g = sheet.id === sheetOrThrow.id && Dp(rangeAddress, v.row, v.col);
        cell.raw = $v(cell.raw, s.old.name, u.name, s.names, g);
      }
      for (const p of sheet.conditionalFormats ?? [])
        if (p.kind === "formula" && p.formula) {
          const m = parseRangeAddress(p.range),
            v = $v(p.formula, s.old.name, u.name, s.names, false),
            g = $v(p.formula, s.old.name, u.name, s.names, true),
            b =
              sheet.id === sheetOrThrow.id && rangesIntersect(rangeAddress, m),
            y =
              b &&
              Dp(rangeAddress, m.firstRow, m.firstCol) &&
              Dp(rangeAddress, m.lastRow, m.lastCol);
          (v !== g &&
            b &&
            !y &&
            throwWorkbookError(
              "This conditional formula uses unqualified table fields across a table boundary. Split or remove its applied range before renaming the field.",
            ),
            (p.formula = y ? g : v));
        }
      for (const p of [...(sheet.pivots ?? []), ...(sheet.queries ?? [])])
        if (h(p.source))
          if (
            ((p.source = { ...p.source, tableId: u.id, range: u.range }),
            (p.filters &&= p.filters.map((filter) => ({
              ...filter,
              field: l(filter.field),
            }))),
            "rowFields" in p)
          ) {
            ((p.rowFields = p.rowFields.map(l)),
              (p.columnFields = p.columnFields.map(l)),
              (p.values = p.values.map((value) => ({
                ...value,
                field: l(value.field),
              }))),
              (p.dateGroups &&= p.dateGroups.map((dateGroup) => ({
                ...dateGroup,
                field: l(dateGroup.field),
              }))));
            for (const m of ["slicers", "timelines", "sorts"])
              p[m] &&= p[m].map((v) => ({ ...v, field: l(v.field) }));
            p.calculatedFields &&= p.calculatedFields.map(
              (calculatedField) => ({
                ...calculatedField,
                field: l(calculatedField.field),
                ...(calculatedField.otherField
                  ? { otherField: l(calculatedField.otherField) }
                  : {}),
              }),
            );
          } else
            ((p.fields = p.fields.map(l)),
              (p.rename &&= Object.fromEntries(
                Object.entries(p.rename).map(([m, v]) => [l(m), v]),
              )),
              (p.sort &&= { ...p.sort, field: l(p.sort.field) }),
              (p.derived &&= p.derived.map((m) => ({
                ...m,
                field: l(m.field),
                ...(m.otherField ? { otherField: l(m.otherField) } : {}),
              }))));
    }
    for (const f of n.definedNames ?? [])
      f.refersTo = $v(f.refersTo, s.old.name, u.name, s.names);
    normalizeTableDefinitions(sheetOrThrow, sheetOrThrow);
  }
  assertUniqueTableNames(n);
  for (const sheet of workbook.sheets) {
    if (!sheet.protection?.enabled) continue;
    const d = n.sheets.find((u) => u.id === sheet.id);
    JSON.stringify(sheet) !== JSON.stringify(d) &&
      throwWorkbookError(
        `This change would update dependent cells or analysis on protected worksheet ${sheet.name}. Unprotect it before renaming the table field.`,
      );
  }
  return n;
}
function buildPivotSource(workbook, source) {
  const sourceReference = normalizeSourceReference(source),
    sheetOrThrow = findSheetOrThrow(workbook, sourceReference.sheetId),
    o = sourceReference.tableId
      ? (sheetOrThrow.tables?.find(
          (table) => table.id === sourceReference.tableId,
        ) ?? throwWorkbookError("The source table no longer exists."))
      : sheetOrThrow.tables?.find(
          (table) => table.range === toA1RangeAddress(sourceReference.range),
        ),
    sourceRange = resolveSourceRange(
      o?.range ?? sourceReference.range,
      sheetOrThrow,
    ),
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  o?.showTotals && sourceRange.lastRow--;
  const fields = [];
  for (
    let firstCol = sourceRange.firstCol;
    firstCol <= sourceRange.lastCol;
    firstCol++
  )
    fields.push(
      validateColumnName(
        workbookWithSettings(
          sheetOrThrow.id,
          Cn(sourceRange.firstRow, firstCol),
        ),
      ),
    );
  new Set(fields.map((field) => field.toLowerCase())).size !== fields.length &&
    throwWorkbookError(
      "Source headers must be nonblank, unique text. Create a table to normalize blank or duplicate headers.",
    );
  const rows = [];
  for (let h = sourceRange.firstRow + 1; h <= sourceRange.lastRow; h++) {
    const values = fields.map((field, m) =>
      workbookWithSettings(sheetOrThrow.id, Cn(h, sourceRange.firstCol + m)),
    );
    (values.forEach((value, m) => {
      const v = Cn(h, sourceRange.firstCol + m);
      typeof value == "string" &&
        value.startsWith("#") &&
        workbookWithSettings.evaluateFormula(
          sheetOrThrow.id,
          v,
          `=ISERROR(${v})`,
        ) === true &&
        throwWorkbookError(
          `Source calculation error ${value} at ${sheetOrThrow.name}!${v}. Repair the source calculation before building or refreshing analysis; no totals were written.`,
        );
    }),
      (values.some((value) => !isNullOrEmptyValue(value)) ||
        fields.some(
          (field, m) =>
            !isNullOrEmptyValue(
              sheetOrThrow.cells[Cn(h, sourceRange.firstCol + m)]?.raw ?? null,
            ),
        )) &&
        rows.push({ index: h, values: values }));
  }
  return (
    workbookWithSettings.getOperationalErrors().length &&
      throwWorkbookError(
        `Source calculation is incomplete (${workbookWithSettings.getOperationalErrors().join(", ")}). Repair blocked or over-limit dependencies before building analysis.`,
      ),
    {
      fields: fields,
      rows: rows,
      source: {
        ...sourceReference,
        range: o?.range ?? toA1RangeAddress(sourceReference.range),
        ...(o ? { tableId: o.id } : {}),
      },
    }
  );
}
function getPivotSourceFields(workbook, sheetId, range) {
  return buildPivotSource(workbook, { sheetId: sheetId, range: range }).fields;
}
function aggregatePivotRows(workbook, r, pivotSource) {
  const sheetOrThrow = findSheetOrThrow(workbook, pivotSource.source.sheetId),
    rangeAddress = parseRangeAddress(pivotSource.source.range),
    cells = Object.entries(sheetOrThrow.cells)
      .filter(([u]) => {
        const l = _t(u);
        return Dp(rangeAddress, l.row, l.col);
      })
      .sort(([u], [l]) => u.localeCompare(l))
      .map(([u, cell]) => [u, cell.raw, !!cell.literal]),
    { destination, ...settings } = normalizePivotDefinition(r);
  return computePivotRefreshFingerprint({
    source: pivotSource.source,
    fields: pivotSource.fields,
    rows: pivotSource.rows,
    cells: cells,
    settings: settings,
  });
}
function getPivotRefreshState(workbook, pivotId) {
  const t = workbook.sheets.find((sheet) =>
      sheet.pivots?.some((pivot) => pivot.id === pivotId),
    ),
    definition = t?.pivots?.find((pivot) => pivot.id === pivotId);
  if (
    ((!t || !definition) && throwWorkbookError("Pivot definition not found."),
    !definition.refresh)
  )
    return {
      state: "unverified",
      reason:
        "This older result has no refresh evidence. Refresh it before relying on its totals.",
      outputModified: false,
      definition: definition,
    };
  const outputModified =
    collectRangeValues(t, definition.outputRange) !==
    definition.refresh.outputFingerprint;
  try {
    const pivotSource = buildPivotSource(workbook, definition.source);
    validatePivotFieldReferences(definition, pivotSource.fields);
    const d =
      aggregatePivotRows(workbook, definition, pivotSource) !==
      definition.refresh.sourceFingerprint;
    return {
      state: d || outputModified ? "stale" : "current",
      reason: outputModified
        ? "Output cells were edited. Refresh into a new result before relying on these totals."
        : d
          ? "Source content, calculated values, range or settings changed. Refresh into a new result."
          : "Source content and output match the saved refresh snapshot.",
      outputModified: outputModified,
      definition: definition,
    };
  } catch (error) {
    return {
      state: "unavailable",
      reason: error.message,
      outputModified: outputModified,
      definition: definition,
    };
  }
}
function describePivotRefresh(workbook, pivotId) {
  const { definition } = findPivotDefinition(workbook, pivotId),
    pivotRefreshState = getPivotRefreshState(workbook, pivotId);
  return buildPivotStatusRecord(workbook, definition, pivotRefreshState);
}
function inspectPivotControlFields(workbook, pivotId, t) {
  t.length > 30 &&
    throwWorkbookError("Inspect at most 30 pivot control fields together.");
  const { definition } = findPivotDefinition(workbook, pivotId),
    pivotSource = buildPivotSource(workbook, definition.source);
  return new Map(
    [...new Set(t)].map((a) => {
      const s = pivotSource.fields.indexOf(a),
        d = new Map();
      if (s < 0)
        return [
          a,
          {
            items: [],
            error: `Broken pivot field “${a}”: source field was not found. Restore or reselect it.`,
          },
        ];
      for (const row of pivotSource.rows)
        if (
          (d.set(buildSlicerItemKey(row.values[s]), row.values[s]),
          d.size > 500)
        )
          return [
            a,
            {
              items: [],
              error:
                "This slicer exceeds 500 distinct items. Choose another field or reduce the source.",
            },
          ];
      return [a, { items: [...d.values()], error: "" }];
    }),
  );
}
function getPivotSlicerItems(workbook, pivotId, field) {
  const n = inspectPivotControlFields(workbook, pivotId, [field]).get(field);
  return (n.error && throwWorkbookError(n.error), n.items);
}
function buildPivotModel(workbook, pivotId) {
  const { owner, definition } = findPivotDefinition(workbook, pivotId),
    pivotDefinition = normalizePivotDefinition(definition),
    pivotSource = buildPivotSource(workbook, definition.source);
  validatePivotFieldReferences(pivotDefinition, pivotSource.fields);
  const rows = filterPivotSourceRows(
    pivotDefinition,
    pivotSource.fields,
    filterSourceRows(pivotSource, pivotDefinition.filters ?? []),
  );
  return (
    (!definition.refresh ||
      aggregatePivotRows(workbook, pivotDefinition, pivotSource) !==
        definition.refresh.sourceFingerprint ||
      collectRangeValues(owner, definition.outputRange) !==
        definition.refresh.outputFingerprint) &&
      throwWorkbookError(
        "This PivotTable is stale or unverified. Refresh it before drilling or charting; historical totals cannot be matched to current source rows.",
      ),
    {
      owner,
      definition,
      config: pivotDefinition,
      data: pivotSource,
      rows: rows,
      built: computePivotTable(pivotDefinition, pivotSource.fields, rows),
    }
  );
}
function buildPivotChartModel(book, pivotId) {
  const { definition, built } = buildPivotModel(book, pivotId),
    o = definition.chart?.valueIndex ?? 0;
  built.chartRows.length > 200 &&
    throwWorkbookError(
      "PivotCharts are limited to 200 leaf groups. Filter or regroup before charting.",
    );
  const points = built.chartRows.map((chartRow) => ({
    label: chartRow.label,
    value: chartRow.values[o],
  }));
  return (
    points.some(
      (point) => point.value !== null && typeof point.value != "number",
    ) &&
      throwWorkbookError(
        "The PivotChart contains a calculation error. Resolve zero denominators or invalid source calculations.",
      ),
    {
      kind: definition.chart?.kind ?? "column",
      measure: built.measures[o].name,
      showAs: definition.values[o].showAs ?? "value",
      points: points,
      sourceRevision: definition.sourceRevision,
    }
  );
}
function buildPivotDrillDown(workbook, pivotId, t) {
  const { definition, config, data, rows, built } = buildPivotModel(
      workbook,
      pivotId,
    ),
    u = _t(t),
    rangeAddress = parseRangeAddress(definition.outputRange),
    c = u.row - rangeAddress.firstRow - 1,
    h = u.col - rangeAddress.firstCol - Math.max(1, config.rowFields.length);
  (c < 0 ||
    c >= built.rowAxes.length ||
    h < 0 ||
    u.col > rangeAddress.lastCol ||
    u.row > rangeAddress.lastRow) &&
    throwWorkbookError(
      "Choose a PivotTable value cell, subtotal or grand total; headers and group labels cannot be drilled.",
    );
  const f = built.rowAxes[c],
    p = built.columnAxes[Math.floor(h / config.values.length)],
    valueIndex = h % config.values.length,
    v = (E, A, k) =>
      !A ||
      A.parts.every((part, I) => {
        const O = E.values[findSourceFieldIndex(data.fields, k[I])],
          j = config.dateGroups?.find((dateGroup) => dateGroup.field === k[I]);
        return (
          JSON.stringify(
            j
              ? bucketPivotDateValue(O, j, E.index)
              : isNullOrEmptyValue(O)
                ? null
                : O,
          ) === JSON.stringify(part)
        );
      }),
    g = rows.filter(
      (row) => v(row, f, config.rowFields) && v(row, p, config.columnFields),
    ),
    sheetOrThrow = findSheetOrThrow(workbook, data.source.sheetId),
    y = parseRangeAddress(data.source.range),
    records = g.map((E) => ({
      locator: {
        workbookId: workbook.id,
        sheetId: sheetOrThrow.id,
        sheetName: sheetOrThrow.name,
        row: E.index + 1,
        range: `${columnIndexToLetters(y.firstCol)}${E.index + 1}:${columnIndexToLetters(y.lastCol)}${E.index + 1}`,
        original: findRowOrigin(workbook, sheetOrThrow.id, E.index),
      },
      values: [...E.values],
      cells: data.fields.map((field, k) => ({
        address: Cn(E.index, y.firstCol + k),
        cell: structuredClone(
          sheetOrThrow.cells[Cn(E.index, y.firstCol + k)] ?? { raw: null },
        ),
      })),
    }));
  return {
    context: buildPivotStatusRecord(workbook, definition, {
      state: "current",
      reason: "Source and output fingerprints match the saved calculation.",
    }),
    address: `${columnIndexToLetters(u.col)}${u.row + 1}`,
    valueIndex: valueIndex,
    measure: config.values[valueIndex],
    underlyingValue: built.baseMatrix[c + 1][u.col - rangeAddress.firstCol],
    displayedValue: built.matrix[c + 1][u.col - rangeAddress.firstCol],
    rowGroup: f?.parts ?? null,
    columnGroup: p?.parts ?? null,
    fields: [...data.fields],
    records: records,
  };
}
function validateTotalsRowForAppend(workbook, sheetId, t) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(t.range),
    a = (d, u) => {
      const l = stripFormulaStringLiterals(d.replace(/"(?:[^"]|"")*"/g, '""'));
      /\b(?:INDIRECT|OFFSET|EVALUATE|LET|LAMBDA)\s*\(/i.test(l) &&
        throwWorkbookError(
          "A dynamic formula could refer to the moving totals row. Preserve or remove that ambiguity before appending.",
        );
      for (const f of l.matchAll(
        new RegExp(
          "(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*))\\s*!\\s*)?(\\$?[1-9]\\d*)\\s*:\\s*(\\$?[1-9]\\d*)(?![A-Za-z0-9_.])",
          "g",
        ),
      )) {
        const p = f[1]?.replace(/''/g, "'") ?? f[2];
        if (
          p
            ? p.toLowerCase() !== sheetOrThrow.name.toLowerCase()
            : u !== undefined && u !== sheetId
        )
          continue;
        const m = Number(f[3].replace("$", "")) - 1,
          v = Number(f[4].replace("$", "")) - 1;
        rangeAddress.lastRow >= Math.min(m, v) &&
          rangeAddress.lastRow <= Math.max(m, v) &&
          throwWorkbookError(
            "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
          );
      }
      for (const f of l.matchAll(
        new RegExp(
          "(?<![A-Za-z0-9_.])(?:(?:'((?:[^']|'')+)'|([A-Za-z_\\\\][A-Za-z0-9_.\\\\]*))\\s*!\\s*)?(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*)(?:\\s*:\\s*(\\$?[A-Za-z]{1,3}\\$?[1-9]\\d*))?(?![A-Za-z0-9_.])",
          "g",
        ),
      )) {
        const p = f[1]?.replace(/''/g, "'") ?? f[2];
        if (
          p
            ? p.toLowerCase() !== sheetOrThrow.name.toLowerCase()
            : u !== undefined && u !== sheetId
        )
          continue;
        const m = parseRangeAddress(`${f[3]}:${f[4] ?? f[3]}`);
        rangesIntersect(m, {
          ...rangeAddress,
          firstRow: rangeAddress.lastRow,
        }) &&
          throwWorkbookError(
            "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
          );
      }
    };
  for (const sheet of workbook.sheets) {
    for (const [u, cell] of Object.entries(sheet.cells)) {
      const c = _t(u);
      if (
        !(
          sheet.id === sheetId &&
          c.row === rangeAddress.lastRow &&
          c.col >= rangeAddress.firstCol &&
          c.col <= rangeAddress.lastCol
        ) &&
        (typeof cell.raw == "string" &&
          cell.raw.startsWith("=") &&
          !cell.literal &&
          a(cell.raw, sheet.id),
        cell.hyperlink?.kind === "internal" &&
          cell.hyperlink.sheetId === sheetId)
      ) {
        const h = _t(cell.hyperlink.address);
        h.row === rangeAddress.lastRow &&
          h.col >= rangeAddress.firstCol &&
          h.col <= rangeAddress.lastCol &&
          throwWorkbookError(
            "An internal hyperlink targets the moving totals row. Retarget or remove it before appending.",
          );
      }
    }
    for (const u of sheet.conditionalFormats ?? [])
      u.kind === "formula" && u.formula && a(u.formula, sheet.id);
    for (const u of sheet.tables ?? [])
      for (const column of u.columns)
        column.calculatedFormula && a(column.calculatedFormula, sheet.id);
  }
  for (const d of workbook.definedNames ?? []) a(d.refersTo, d.scopeSheetId);
  for (
    let firstCol = rangeAddress.firstCol;
    firstCol <= rangeAddress.lastCol;
    firstCol++
  ) {
    const cell = sheetOrThrow.cells[Cn(rangeAddress.lastRow, firstCol)],
      tableTotalCell = buildTableTotalCell(t, firstCol - rangeAddress.firstCol);
    cell &&
      (cell.raw !== tableTotalCell.raw || cell.literal) &&
      throwWorkbookError(
        "The totals row was edited. Preserve its values before appending.",
      );
  }
}
function resizeTableRange(workbook, sheetId, t, sourceRange, showTotals) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(t.range);
  (showTotals !== t.showTotals &&
    throwWorkbookError(
      "Clear calculated column rules before changing totals-row visibility.",
    ),
    (sourceRange.firstRow !== rangeAddress.firstRow ||
      sourceRange.firstCol !== rangeAddress.firstCol ||
      sourceRange.lastCol !== rangeAddress.lastCol ||
      sourceRange.lastRow < rangeAddress.lastRow) &&
      throwWorkbookError(
        "Calculated tables support downward resize only; keep the existing header and columns. Clear rules before shrinking or changing columns.",
      ),
    (sourceRange.lastRow - rangeAddress.lastRow) *
      (sourceRange.lastCol - sourceRange.firstCol + 1) +
      (t.showTotals ? t.columns.length : 0) >
      tableMutationLimits.changedCells &&
      throwWorkbookError("Table append/resize may change at most 500 cells."));
  for (let u = rangeAddress.lastRow + 1; u <= sourceRange.lastRow; u++)
    for (
      let firstCol = sourceRange.firstCol;
      firstCol <= sourceRange.lastCol;
      firstCol++
    )
      sheetOrThrow.cells[Cn(u, firstCol)] &&
        throwWorkbookError(
          "Table append/resize would overwrite adjacent cells or their formatting. Move that content first.",
        );
  if (t.showTotals && sourceRange.lastRow !== rangeAddress.lastRow) {
    validateTotalsRowForAppend(workbook, sheetId, t);
    for (const u of [
      ...(workbook.formulaWatches ?? []),
      ...(workbook.reviewThreads ?? []),
    ]) {
      if (u.orphaned || u.anchor.sheetId !== sheetId) continue;
      const l = _t(u.anchor.address);
      l.row === sourceRange.lastRow &&
        l.col >= sourceRange.firstCol &&
        l.col <= sourceRange.lastCol &&
        throwWorkbookError(
          "The new totals row already has a saved watch or review discussion. Preserve or remove that destination anchor before appending.",
        );
    }
  }
}
function moveTotalsRowCells(workbook, next, sheetId, n, o) {
  const rangeAddress = parseRangeAddress(n.range),
    s = parseRangeAddress(o.range),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    u = findSheetOrThrow(next, sheetId);
  for (
    let firstCol = rangeAddress.firstCol;
    firstCol <= rangeAddress.lastCol;
    firstCol++
  ) {
    const h = Cn(rangeAddress.lastRow, firstCol),
      f = Cn(s.lastRow, firstCol),
      p = sheetOrThrow.cells[h];
    if (p) {
      const m = {
        ...structuredClone(p),
        ...buildTableTotalCell(o, firstCol - rangeAddress.firstCol),
      };
      (delete m.literal, (u.cells[f] = m));
    }
  }
  const l = (c) => {
    if (c.sheetId !== sheetId) return;
    const h = _t(c.address);
    h.row === rangeAddress.lastRow &&
      h.col >= rangeAddress.firstCol &&
      h.col <= rangeAddress.lastCol &&
      (c.address = Cn(s.lastRow, h.col));
  };
  for (const c of next.formulaWatches ?? []) c.orphaned || l(c.anchor);
  for (const c of next.reviewThreads ?? []) c.orphaned || l(c.anchor);
}
function resolveTableDataRange(workbook, sheetId, range = "A1") {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(range);
  if (
    rangeAddress.firstRow !== rangeAddress.lastRow ||
    rangeAddress.firstCol !== rangeAddress.lastCol
  )
    return range;
  const a = sheetOrThrow.tables?.find((table) => {
    const l = parseRangeAddress(table.range);
    return (
      rangeAddress.firstRow >= l.firstRow &&
      rangeAddress.firstRow <= l.lastRow &&
      rangeAddress.firstCol >= l.firstCol &&
      rangeAddress.firstCol <= l.lastCol
    );
  });
  if (!a) return formatOccupiedRangeAddress(sheetOrThrow, false);
  const s = parseRangeAddress(a.range);
  return `${columnIndexToLetters(s.firstCol)}${s.firstRow + 1}:${columnIndexToLetters(s.lastCol)}${s.lastRow + 1 - (a.showTotals ? 1 : 0)}`;
}
function validateFilterRequest(workbook, sheetId, t) {
  if (
    !t ||
    typeof t.range != "string" ||
    !Array.isArray(t.columns) ||
    t.columns.length > filterEvaluationLimits.columns ||
    [t.caseSensitive, t.unique].some(
      (f) => f !== undefined && typeof f != "boolean",
    )
  )
    throw Error("Choose a header range and up to 64 column filters.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(t.range),
    a = rangeAddress.lastRow - rangeAddress.firstRow,
    s = rangeAddress.lastCol - rangeAddress.firstCol + 1;
  if (
    rangeAddress.lastRow >= sheetOrThrow.rowCount ||
    rangeAddress.lastCol >= sheetOrThrow.colCount ||
    a < 1
  )
    throw Error(
      "Filter range needs one header row and at least one data row within the worksheet.",
    );
  if (
    a > filterEvaluationLimits.rows ||
    (a + 1) * s > filterEvaluationLimits.cells ||
    s > filterEvaluationLimits.columns
  )
    throw Error(
      "Filtering is limited to 10,000 data rows, 100,000 cells including headers, and 64 columns. Choose a smaller range.",
    );
  if (
    sheetOrThrow.merges?.some((merge) => {
      const p = parseRangeAddress(merge);
      return (
        p.firstRow <= rangeAddress.lastRow &&
        p.lastRow >= rangeAddress.firstRow &&
        p.firstCol <= rangeAddress.lastCol &&
        p.lastCol >= rangeAddress.firstCol
      );
    })
  )
    throw Error(
      "Unmerge cells in the filter range first; records cannot be split by filtering.",
    );
  const d = (f) =>
      Number.isInteger(f) &&
      f >= rangeAddress.firstCol &&
      f <= rangeAddress.lastCol,
    u = (f) => {
      if (
        !f ||
        !d(f.column) ||
        !["value", "text", "number", "date"].includes(f.type) ||
        !textFilterOperators.includes(f.operator) ||
        (!["blank", "notBlank"].includes(f.operator) &&
          !isStorableCellValue(f.value))
      )
        throw Error("Invalid filter condition, column or value.");
      if (
        !["blank", "notBlank"].includes(f.operator) &&
        f.type === "text" &&
        typeof f.value != "string"
      )
        throw Error("Text criteria require a text value.");
      if (
        typeof f.value == "string" &&
        f.value.length > filterEvaluationLimits.text
      )
        throw Error("Filter criteria text is limited to 1,024 characters.");
      if (
        [
          "contains",
          "notContains",
          "startsWith",
          "endsWith",
          "wildcard",
        ].includes(f.operator) &&
        f.type !== "text"
      )
        throw Error("Text matching requires a Text condition.");
      if (["gt", "gte", "lt", "lte"].includes(f.operator) && f.type === "value")
        throw Error("Choose Number, Date or Text for ordered comparisons.");
      if (
        !["blank", "notBlank"].includes(f.operator) &&
        ["number", "date"].includes(f.type) &&
        resolveDatePolicyContext(f.value, f.type, workbook) === null
      )
        throw Error(
          `Invalid ${f.type} filter value. Use a finite number or a date in the workbook's date policy.`,
        );
    };
  let l = 0;
  const c = new Set();
  for (const column of t.columns) {
    if (
      !column ||
      !d(column.column) ||
      c.has(column.column) ||
      (column.join !== undefined && !["and", "or"].includes(column.join))
    )
      throw Error("Filter columns must be valid and distinct.");
    if (
      (c.add(column.column),
      column.values !== undefined &&
        (!Array.isArray(column.values) ||
          column.values.length > filterEvaluationLimits.values ||
          !column.values.every(isStorableCellValue)))
    )
      throw Error("Value filters allow up to 1,000 finite typed values.");
    if (
      column.conditions !== undefined &&
      (!Array.isArray(column.conditions) || column.conditions.length > 2)
    )
      throw Error("Each column supports up to two custom conditions.");
    for (const p of column.conditions ?? []) {
      if (p.column !== column.column)
        throw Error("A column condition must use its own column.");
      (u(p), l++);
    }
  }
  if (
    t.advanced !== undefined &&
    (!Array.isArray(t.advanced) ||
      t.advanced.length > filterEvaluationLimits.conditions)
  )
    throw Error("Invalid advanced filter groups.");
  for (const f of t.advanced ?? []) {
    if (
      !Array.isArray(f) ||
      !f.length ||
      f.length > filterEvaluationLimits.conditions
    )
      throw Error(
        "Each advanced OR group must contain at least one AND condition.",
      );
    for (const p of f) (u(p), l++);
  }
  if (l > filterEvaluationLimits.conditions)
    throw Error("At most 128 combined filter conditions are allowed.");
  return structuredClone({
    ...t,
    range: `${columnIndexToLetters(rangeAddress.firstCol)}${rangeAddress.firstRow + 1}:${columnIndexToLetters(rangeAddress.lastCol)}${rangeAddress.lastRow + 1}`,
  });
}
function applyFilterConditions(
  workbook,
  sheetId,
  t,
  n = calculateWorkbookWithSettings(workbook),
) {
  const definition = validateFilterRequest(workbook, sheetId, t),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    rangeAddress = parseRangeAddress(definition.range),
    d = !!definition.caseSensitive,
    axisVisibilityIndex = createAxisVisibilityIndex(
      sheetOrThrow.rowCount,
      cn(sheetOrThrow, "row"),
    ),
    sourceRows = [],
    matchingSourceRows = [],
    visibleSourceRows = [],
    excludedRows = [],
    p = new Set(),
    originalSourceIds = {},
    v = definition.columns.map((column) =>
      column.values === undefined
        ? null
        : new Set(
            column.values.map((value) => comparableCellValueKey(value, d)),
          ),
    );
  let manuallyHiddenRows = 0,
    hiddenMatchingRows = 0,
    y = 0;
  const w = (O, j) => {
      if (
        ((y +=
          j.operator === "wildcard"
            ? Math.max(1, String(O ?? "").length) *
              Math.max(1, String(j.value ?? "").length)
            : 1 + String(O ?? "").length + String(j.value ?? "").length),
        y > filterEvaluationLimits.comparisonWork)
      )
        throw Error(
          "Filter comparison work exceeds its 5,000,000-unit budget. Narrow the range or simplify criteria.",
        );
      return matchesTextFilter(O, j, workbook, d);
    },
    fields = Array.from(
      { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
      (O, j) => ({
        column: rangeAddress.firstCol + j,
        label:
          String(
            assertCalculatedBeforeFilter(
              n,
              sheetId,
              `${columnIndexToLetters(rangeAddress.firstCol + j)}${rangeAddress.firstRow + 1}`,
              workbook,
            ) ?? "",
          ) || columnIndexToLetters(rangeAddress.firstCol + j),
      }),
    );
  for (let O = rangeAddress.firstRow + 1; O <= rangeAddress.lastRow; O++) {
    const j = fields.map((field) =>
        assertCalculatedBeforeFilter(
          n,
          sheetId,
          `${columnIndexToLetters(field.column)}${O + 1}`,
          workbook,
        ),
      ),
      z = (B) => j[B - rangeAddress.firstCol];
    if (
      ((y += j.reduce(
        (B, _) => B + 1 + (typeof _ == "string" ? _.length : 0),
        0,
      )),
      y > filterEvaluationLimits.comparisonWork)
    )
      throw Error(
        "Filter comparison work exceeds its 5,000,000-unit budget. Narrow the range or simplify criteria.",
      );
    (sourceRows.push(O + 1),
      sheetOrThrow.rowOriginIds?.[O] &&
        (originalSourceIds[O + 1] = sheetOrThrow.rowOriginIds[O]));
    const V = axisVisibilityIndex.isHidden(O);
    V && manuallyHiddenRows++;
    let D =
      definition.columns.every(
        (column, _) =>
          (v[_] === null ||
            v[_].has(comparableCellValueKey(z(column.column), d))) &&
          (!column.conditions?.length ||
            (column.join === "or"
              ? column.conditions.some((condition) =>
                  w(z(column.column), condition),
                )
              : column.conditions.every((condition) =>
                  w(z(column.column), condition),
                ))),
      ) &&
      (!definition.advanced?.length ||
        definition.advanced.some((B) => B.every((_) => w(z(_.column), _))));
    if (D && definition.unique) {
      const B = JSON.stringify(j.map((_) => comparableCellValueKey(_, d)));
      p.has(B) ? (D = false) : p.add(B);
    }
    if (D)
      (matchingSourceRows.push(O + 1),
        V ? hiddenMatchingRows++ : visibleSourceRows.push(O + 1));
    else {
      const B = excludedRows.at(-1);
      B?.end === O - 1 ? (B.end = O) : excludedRows.push({ start: O, end: O });
    }
  }
  const A = [];
  for (const O of [...cn(sheetOrThrow, "row"), ...excludedRows].sort(
    (j, z) => j.start - z.start,
  )) {
    const j = A.at(-1);
    j && O.start <= j.end + 1
      ? (j.end = Math.max(j.end, O.end))
      : A.push({ ...O });
  }
  return {
    hiddenRows: normalizeHiddenIntervals(A, sheetOrThrow.rowCount),
    fields: fields,
    excludedRows: excludedRows,
    context: {
      workbookId: workbook.id,
      revision: workbook.revision,
      sheetId: sheetId,
      sheetName: sheetOrThrow.name,
      range: definition.range,
      definition: definition,
      totalRows: sourceRows.length,
      matchedRows: matchingSourceRows.length,
      visibleRows: visibleSourceRows.length,
      filteredRows: sourceRows.length - matchingSourceRows.length,
      manuallyHiddenRows: manuallyHiddenRows,
      hiddenMatchingRows: hiddenMatchingRows,
      sourceRows: sourceRows,
      matchingSourceRows: matchingSourceRows,
      visibleSourceRows: visibleSourceRows,
      originalSourceIds: originalSourceIds,
      policy: savedFilterViewHelpText,
    },
  };
}
function buildFilterFacetValues(workbook, sheetId, t, n, o = "") {
  const a = validateFilterRequest(workbook, sheetId, t),
    rangeAddress = parseRangeAddress(a.range);
  if (
    !Number.isInteger(n) ||
    n < rangeAddress.firstCol ||
    n > rangeAddress.lastCol
  )
    throw Error("Facet column is outside the filter range.");
  const workbookWithSettings = calculateWorkbookWithSettings(workbook),
    u = applyFilterConditions(
      workbook,
      sheetId,
      {
        ...a,
        unique: false,
        columns: a.columns.filter((column) => column.column !== n),
      },
      workbookWithSettings,
    ),
    l = new Map();
  for (const matchingSourceRow of u.context.matchingSourceRows) {
    const h = assertCalculatedBeforeFilter(
        workbookWithSettings,
        sheetId,
        `${columnIndexToLetters(n)}${matchingSourceRow}`,
        workbook,
      ),
      f = comparableCellValueKey(h, a.caseSensitive);
    if (
      !normalizeComparableText(String(h ?? "(Blanks)")).includes(
        normalizeComparableText(o),
      )
    )
      continue;
    const p = l.get(f);
    p ? p.count++ : l.set(f, { value: h, count: 1, key: f });
  }
  return {
    values: [...l.values()].slice(0, filterEvaluationLimits.values),
    distinct: l.size,
    truncated: l.size > filterEvaluationLimits.values,
  };
}
function validateSavedFilterViews(savedFilterViews, r) {
  if (savedFilterViews === undefined) return;
  if (
    !Array.isArray(savedFilterViews) ||
    savedFilterViews.length > savedFilterViewLimits.views ||
    JSON.stringify(savedFilterViews).length > savedFilterViewLimits.characters
  )
    throw Error(
      "Saved filter views are limited to 20 views and 512,000 metadata characters per workbook.",
    );
  const t = new Set(),
    n = new Set();
  r.savedFilterViews = savedFilterViews.map((savedFilterView) => {
    if (
      (validateSavedFilterViewMetadata(savedFilterView, [
        "id",
        "name",
        "sheetId",
        "sheetName",
        "sourceWorkbookId",
        "sourceRevision",
        "savedAt",
        "datePolicy",
        "definition",
        "population",
      ]),
      !isValidFilterViewText(savedFilterView.id, 100) ||
        !isValidFilterViewText(savedFilterView.sheetId, 100) ||
        !isValidFilterViewText(savedFilterView.sheetName, 31) ||
        !isValidFilterViewText(savedFilterView.sourceWorkbookId, 100) ||
        !Number.isSafeInteger(savedFilterView.sourceRevision) ||
        savedFilterView.sourceRevision < 0 ||
        savedFilterView.sourceRevision > Number.MAX_SAFE_INTEGER - 1 ||
        typeof savedFilterView.savedAt != "string" ||
        !Number.isFinite(Date.parse(savedFilterView.savedAt)) ||
        new Date(savedFilterView.savedAt).toISOString() !==
          savedFilterView.savedAt)
    )
      throw Error("Invalid saved filter view identity, revision or timestamp.");
    const savedFilterViewName = normalizeSavedFilterViewName(
      savedFilterView.name,
    );
    if (
      t.has(savedFilterView.id) ||
      n.has(normalizeNameForComparison(savedFilterViewName))
    )
      throw Error("Saved filter view names and IDs must be unique.");
    (t.add(savedFilterView.id),
      n.add(normalizeNameForComparison(savedFilterViewName)),
      validateSavedFilterViewMetadata(savedFilterView.datePolicy, [
        "calendar",
        "dateSystem",
        "textDateOrder",
        "timeZone",
      ]));
    const datePolicy = normalizeDatePolicy(savedFilterView.datePolicy);
    validateSavedFilterViewMetadata(savedFilterView.definition, [
      "range",
      "columns",
      "advanced",
      "caseSensitive",
      "unique",
    ]);
    const d = (h) =>
      validateSavedFilterViewMetadata(h, [
        "column",
        "type",
        "operator",
        "value",
      ]);
    if (!Array.isArray(savedFilterView.definition.columns))
      throw Error("Invalid saved filter columns.");
    for (const column of savedFilterView.definition.columns)
      if (
        (validateSavedFilterViewMetadata(column, [
          "column",
          "values",
          "conditions",
          "join",
        ]),
        column.conditions !== undefined)
      ) {
        if (!Array.isArray(column.conditions))
          throw Error("Invalid saved filter conditions.");
        column.conditions.forEach(d);
      }
    if (savedFilterView.definition.advanced !== undefined) {
      if (!Array.isArray(savedFilterView.definition.advanced))
        throw Error("Invalid saved advanced criteria.");
      for (const h of savedFilterView.definition.advanced) {
        if (!Array.isArray(h)) throw Error("Invalid saved advanced criteria.");
        h.forEach(d);
      }
    }
    const workbook = {
        ...r,
        datePolicy: datePolicy,
        sheets: [
          {
            id: savedFilterView.sheetId,
            name: savedFilterView.sheetName,
            rowCount: workbookLimits.maxRows,
            colCount: workbookLimits.maxCols,
            cells: {},
          },
        ],
      },
      definition = validateFilterRequest(
        workbook,
        savedFilterView.sheetId,
        savedFilterView.definition,
      );
    validateSavedFilterViewMetadata(savedFilterView.population, [
      "totalRows",
      "matchedRows",
      "visibleRows",
      "hiddenMatchingRows",
    ]);
    const population = savedFilterView.population;
    if (
      Object.keys(population).length !== 4 ||
      Object.values(population).some(
        (h) => !Number.isInteger(h) || Number(h) < 0,
      ) ||
      population.totalRows !==
        parseRangeAddress(definition.range).lastRow -
          parseRangeAddress(definition.range).firstRow ||
      population.totalRows > 1e4 ||
      population.totalRows < 1 ||
      population.matchedRows > population.totalRows ||
      population.visibleRows > population.matchedRows ||
      population.hiddenMatchingRows !==
        population.matchedRows - population.visibleRows
    )
      throw Error("Invalid saved population counts.");
    return {
      id: savedFilterView.id,
      name: savedFilterViewName,
      sheetId: savedFilterView.sheetId,
      sheetName: savedFilterView.sheetName,
      sourceWorkbookId: savedFilterView.sourceWorkbookId,
      sourceRevision: savedFilterView.sourceRevision,
      savedAt: savedFilterView.savedAt,
      datePolicy: datePolicy,
      definition: definition,
      population: { ...population },
    };
  });
}
function copyWithSavedFilterViews(workbook, savedFilterViews) {
  const t = { ...workbook };
  return (validateSavedFilterViews(savedFilterViews, t), t);
}
function createSavedFilterView(workbook, sheetId, definition, n, o) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  if (
    (assertSavedFilterViewsEditable(workbook, sheetId),
    o && findSavedFilterView(workbook, o).sheetId !== sheetId)
  )
    throw Error("A saved view must keep its source worksheet.");
  const context = applyFilterConditions(workbook, sheetId, definition).context,
    d = {
      id: o ?? crypto.randomUUID(),
      name: normalizeSavedFilterViewName(n),
      sheetId: sheetId,
      sheetName: sheetOrThrow.name,
      sourceWorkbookId: workbook.id,
      sourceRevision: workbook.revision,
      savedAt: new Date().toISOString(),
      datePolicy: normalizeDatePolicy(workbook.datePolicy),
      definition: context.definition,
      population: {
        totalRows: context.totalRows,
        matchedRows: context.matchedRows,
        visibleRows: context.visibleRows,
        hiddenMatchingRows: context.hiddenMatchingRows,
      },
    };
  return copyWithSavedFilterViews(workbook, [
    ...(workbook.savedFilterViews ?? []).filter((u) => u.id !== o),
    d,
  ]);
}
function addSavedFilterView(workbook, r, t) {
  return (
    assertSavedFilterViewsEditable(
      workbook,
      findSavedFilterView(workbook, r).sheetId,
    ),
    copyWithSavedFilterViews(
      workbook,
      workbook.savedFilterViews.map((savedFilterView) =>
        savedFilterView.id === r
          ? { ...savedFilterView, name: normalizeSavedFilterViewName(t) }
          : savedFilterView,
      ),
    )
  );
}
function removeSavedFilterView(workbook, r) {
  return (
    assertSavedFilterViewsEditable(
      workbook,
      findSavedFilterView(workbook, r).sheetId,
    ),
    copyWithSavedFilterViews(
      workbook,
      workbook.savedFilterViews.filter(
        (savedFilterView) => savedFilterView.id !== r,
      ),
    )
  );
}
function applyPrintLayoutChange(workbook, sheetId, layout, n = {}) {
  if (n.readOnly) throw Error("This workbook is read-only.");
  n.expectedSnapshot && verifyPrintSnapshotGuard(workbook, n.expectedSnapshot);
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId);
  assertSheetUnprotected(sheetOrThrow);
  const a = { ...sheetOrThrow };
  return (
    validatePageSetup(layout, a),
    zU(a, a.pageLayout),
    JSON.stringify(a.pageLayout) === JSON.stringify(sheetOrThrow.pageLayout)
      ? workbook
      : {
          ...workbook,
          sheets: workbook.sheets.map((sheet) =>
            sheet.id === sheetId ? a : sheet,
          ),
        }
  );
}

export {
  suggestTableHeaderName,
  buildFilterFacetValues,
  setTableColumnDataType,
  getPivotRefreshState,
  assertRangeCellBudget,
  moveTotalsRowCells,
  resolveTableDataRange,
  findSheetOrThrow,
  buildPivotDrillDown,
  createSavedFilterView,
  findOccupiedRangeBounds,
  buildSavedReportEnvelope,
  validatePageSetup,
  formatOccupiedRangeAddress,
  buildPivotSource,
  validateRemoveDuplicatesRequest,
  parseReportSnapshot,
  inspectPivotControlFields,
  resizeTableRange,
  createDefaultPageSetup,
  updateSheetPageSetup,
  findTableForRange,
  addSavedFilterView,
  applyFilterConditions,
  validateSavedFilterViews,
  describePivotRefresh,
  renameTableHeaderField,
  addReviewThreadReply,
  applyAutoSum,
  collectHyperlinkEntries,
  nQ,
  getPivotSlicerItems,
  setCellHyperlink,
  setReviewThreadResolved,
  applyPrintLayoutChange,
  getPivotSourceFields,
  removeSavedFilterView,
  aggregatePivotRows,
  buildPivotChartModel,
  applyNamedStyleChange,
  validateReviewedTableHeaders,
  zU,
};

/* Also used by the workers. */
export { validateFilterRequest };
