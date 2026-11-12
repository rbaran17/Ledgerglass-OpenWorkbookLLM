import {
  failWithMessage,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import { isNullOrEmptyValue } from "./03-defined-names.js";
import {
  fingerprintCanonicalValue,
  normalizeDatePolicy,
} from "../01-core-schemas-and-hashing.js";
import { number, array, unknown, object } from "zod";
import { workbookLimits } from "../02-financial-functions.js";
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
} from "./05-consolidation-and-sources.js";
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
} from "./07-conditional-formatting.js";
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
} from "./08-chart-model.js";
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
    t = parseRangeAddress(r.request.range);
  if (
    (validateObjectReportRequest(r.request),
    r.hash !== hashRecordExcludingHash(r))
  )
    throw Error("Report snapshot integrity check failed.");
  const n = Array.from(
    { length: t.lastCol - t.firstCol + 1 },
    (s, d) => t.firstCol + d,
  );
  if (
    fingerprintCanonicalValue(n) !== fingerprintCanonicalValue(r.columns) ||
    r.widths.length !== n.length
  )
    throw Error("Invalid report column layout.");
  const o = r.rows.map((s) => s.row),
    a = [...o, ...r.excludedRows].sort((s, d) => s - d);
  if (
    r.rows.length * n.length > 2e3 ||
    a.length !== t.lastRow - t.firstRow + 1 ||
    a.some((s, d) => s !== t.firstRow + d + 1) ||
    o.some((s, d) => d > 0 && s <= o[d - 1])
  )
    throw Error("Invalid report row accounting.");
  for (const s of r.rows) {
    if (s.cells.length !== n.length) throw Error("Invalid report cell layout.");
    s.cells.forEach((d, u) => {
      if (d.address !== `${columnIndexToLetters(n[u])}${s.row}`)
        throw Error("Invalid report cell address.");
      d.style &&= validateCellStyle(d.style);
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
function buildSavedReportEnvelope(e, r) {
  if (e.reportOutputs !== undefined) {
    const t = array(
      object({
        id: nonEmptyLabelSchema,
        version: number().int().min(1).max(1e6),
        snapshot: unknown(),
      }).strict(),
    )
      .max(20)
      .parse(e.reportOutputs);
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
  if (e.embeddedObjectInventories !== undefined) {
    const t = array(importedCacheEvidenceSchema)
      .max(20)
      .parse(e.embeddedObjectInventories);
    if (
      new Set(t.map((n) => n.id)).size !== t.length ||
      JSON.stringify(t).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    r.embeddedObjectInventories = t;
  }
}
function validatePageSetup(e, r) {
  if (e === undefined) return;
  if (!e || typeof e != "object" || Array.isArray(e))
    throw new Error("Invalid page layout.");
  const t = e,
    n = t.margins,
    o = (b, y, w) => {
      if (typeof b != "number" || !Number.isFinite(b) || b < y || b > w)
        throw new Error("Page setup contains an out-of-range number.");
      return b;
    };
  if (
    !["portrait", "landscape"].includes(String(t.orientation)) ||
    !["letter", "a4", "legal"].includes(String(t.paper)) ||
    !n ||
    typeof n != "object"
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
        o(n[b], 0, 3),
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
  const l = parseRangeAddress(t.printArea);
  if (l.lastRow >= r.rowCount || l.lastCol >= r.colCount)
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
  const c = o(t.fitToWidth, 0, 10),
    h = o(t.fitToHeight, 0, 10),
    f = o(t.scale, 10, 400);
  if (![c, h, f].every(Number.isInteger))
    throw new Error("Page fitting and scale must be whole numbers.");
  const p = parsePrintTitles(t.repeatRows, "row", r),
    m = parsePrintTitles(t.repeatColumns, "column", r),
    v = validatePrintPageBreaks(t.rowBreaks, r.rowCount),
    g = validatePrintPageBreaks(t.columnBreaks, r.colCount);
  r.pageLayout = {
    orientation: t.orientation,
    paper: t.paper,
    margins: a,
    printArea: t.printArea,
    fitToPage: t.fitToPage,
    fitToWidth: c,
    fitToHeight: h,
    scale: f,
    gridlines: t.gridlines,
    headings: t.headings,
    header: t.header,
    footer: t.footer,
    ...(t.pageOrder !== undefined ? { pageOrder: t.pageOrder } : {}),
    ...(p ? { repeatRows: formatAxisIntervalRef(p, "row") } : {}),
    ...(m ? { repeatColumns: formatAxisIntervalRef(m, "column") } : {}),
    ...(v ? { rowBreaks: v } : {}),
    ...(g ? { columnBreaks: g } : {}),
  };
}
function zU(e, r) {
  const t = { ...e };
  (validatePageSetup(r, t), (r = t.pageLayout));
  const n = parseRangeAddress(r.printArea);
  if (
    (n.lastRow - n.firstRow + 1) * (n.lastCol - n.firstCol + 1) >
    printRenderLimits.areaCells
  )
    throw Error(
      "Print preview supports at most 5,000 cells per print area. Choose a smaller print area.",
    );
  const a = resolvePageDimensions(r),
    s = a.width - (r.margins.left + r.margins.right) * 96,
    d = a.height - (r.margins.top + r.margins.bottom) * 96,
    u = parsePrintTitles(r.repeatRows, "row", e),
    l = parsePrintTitles(r.repeatColumns, "column", e);
  for (const j of e.merges ?? []) {
    const z = parseRangeAddress(j);
    for (const [V, D, B] of [
      [u, z.firstRow, z.lastRow],
      [l, z.firstCol, z.lastCol],
    ])
      if (V && D <= V.end && B >= V.start && (D < V.start || B > V.end))
        throw Error(
          "A merged area crosses a repeated-title boundary. Unmerge it or adjust the print titles first.",
        );
  }
  for (const j of e.objects ?? []) {
    const z = _t(j.anchor),
      V = applyChartDrawingLayout(e, j);
    for (const [D, B, _, C, T] of [
      [u, "row", z.row, V.offsetY, V.height],
      [l, "column", z.col, V.offsetX, V.width],
    ]) {
      if (!D) continue;
      const N = createAxisVisibilityIndex(
        B === "row" ? e.rowCount : e.colCount,
        cn(e, B),
      );
      if (N.isHidden(_)) continue;
      const R = (G) => (B === "row" ? rowHeightAt(e, G) : columnWidthAt(e, G)),
        L = (G) => {
          let H = _,
            Q = 0;
          for (; G < 0;) {
            const X = N.next(H, -1);
            if (X === null) return H;
            if (((H = X), (G += R(H)), ++Q > 1e3))
              throw Error(
                "Drawing print inspection exceeds its geometry budget.",
              );
          }
          for (; G >= R(H);) {
            G -= R(H);
            const X = N.next(H, 1);
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
      const D = createAxisVisibilityIndex(
          j === "row" ? e.rowCount : e.colCount,
          cn(e, j),
        ),
        B = [];
      for (let _ = z; _ <= V; _++) D.isHidden(_) || B.push(_);
      return B;
    },
    h = c("row", n.firstRow, n.lastRow),
    f = c("column", n.firstCol, n.lastCol);
  if (!h.length || !f.length)
    throw Error(
      "The print area has no visible cells. Unhide its rows or columns, or choose another print area.",
    );
  const p = u ? c("row", u.start, u.end) : [],
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
        d / j - (r.headings ? 22 : 0),
        (z) => rowHeightAt(e, z),
        r.fitToPage ? [] : (r.rowBreaks ?? []),
      ),
      columns: v(
        f,
        m,
        s / j - (r.headings ? 32 : 0),
        (z) => columnWidthAt(e, z),
        r.fitToPage ? [] : (r.columnBreaks ?? []),
      ),
    });
  let b = r.fitToPage ? 1 : r.scale / 100;
  if (r.fitToPage && (r.fitToWidth || r.fitToHeight)) {
    const j = (z) => {
      try {
        const V = g(z);
        return (
          (!r.fitToWidth || V.columns.length <= r.fitToWidth) &&
          (!r.fitToHeight || V.rows.length <= r.fitToHeight)
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
      b = z;
    }
  }
  if (!Number.isFinite(b) || b <= 0)
    throw Error(
      "The print content cannot fit the selected paper. Reduce the print area or titles.",
    );
  const y = g(b);
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
  const E = !!(u || l),
    A = (j, z) => ({
      firstRow: Math.min(...j),
      lastRow: Math.max(...j),
      firstCol: Math.min(...z),
      lastCol: Math.max(...z),
      ...(E ? { rows: j, columns: z } : {}),
    }),
    k =
      r.pageOrder === "downThenOver"
        ? y.columns.flatMap((j) => y.rows.map((z) => A(z, j)))
        : y.rows.flatMap((j) => y.columns.map((z) => A(j, z))),
    F =
      r.fitToPage && (r.rowBreaks?.length ?? 0) + (r.columnBreaks?.length ?? 0)
        ? [
            "Fit to pages ignores saved manual page breaks. Turn it off to use the breaks.",
          ]
        : [];
  return { pages: k, scale: b, ...(F.length ? { notices: F } : {}) };
}
function updateSheetPageSetup(e, r, t) {
  if (!r.pageLayout) return [];
  const n = { ...r };
  validatePageSetup(r.pageLayout, n);
  const o = n.pageLayout,
    a = exportHeaderFooterTemplate(o.header, t),
    s = exportHeaderFooterTemplate(o.footer, t),
    d = [o.header, o.footer].some((u) =>
      [...u.matchAll(headerFooterTokenPattern)].some((l) => !!l[1]),
    );
  return (
    (e.pageSetup = {
      ...e.pageSetup,
      orientation: o.orientation,
      paperSize: o.paper === "a4" ? 9 : o.paper === "legal" ? 5 : 1,
      margins: o.margins,
      printArea: o.printArea,
      fitToPage: o.fitToPage,
      fitToWidth: o.fitToWidth,
      fitToHeight: o.fitToHeight,
      scale: o.scale,
      showGridLines: o.gridlines,
      showRowColHeaders: o.headings,
      pageOrder: o.pageOrder ?? "overThenDown",
      ...(o.repeatRows ? { printTitlesRow: o.repeatRows } : {}),
      ...(o.repeatColumns ? { printTitlesColumn: o.repeatColumns } : {}),
    }),
    (e.rowBreaks = (o.rowBreaks ?? []).map((u) => ({
      id: u,
      max: 16383,
      man: 1,
    }))),
    (e.headerFooter = { oddHeader: a, oddFooter: s }),
    d
      ? [
          "Print workbook/engagement/period/revision fields are exported as literal text from this workbook snapshot. Native JSON retains the live identity tokens; XLSX does not update them after later edits.",
        ]
      : []
  );
}
function findOccupiedRangeBounds(e, r = true) {
  let t = null;
  for (const [n, o] of Object.entries(e.cells)) {
    const a = o.raw !== null && o.raw !== "",
      s =
        r &&
        (o.numberFormat ||
          o.hyperlink ||
          o.note ||
          o.tickmark ||
          o.numFmt ||
          o.format ||
          Object.keys(o.style ?? {}).length);
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
  const t = findOccupiedRangeBounds(e, r);
  if (!t) return "A1";
  const n = `${columnIndexToLetters(t.firstCol)}${t.firstRow + 1}`,
    o = `${columnIndexToLetters(t.lastCol)}${t.lastRow + 1}`;
  return n === o ? n : `${n}:${o}`;
}
function createDefaultPageSetup(e) {
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
    printArea: e ? formatOccupiedRangeAddress(e) : "A1",
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
function validateSelectionBounds(e, r) {
  const t = assertRangeCellBudget(r);
  if (
    t.some((n) => {
      const { row, col } = _t(n);
      return row >= e.rowCount || col >= e.colCount;
    })
  )
    throw new Error("Selection exceeds worksheet dimensions.");
  return t;
}
function validateRemoveDuplicatesRequest(e, r, t, n, o = true) {
  const a = e.sheets.find((m) => m.id === r);
  assertSheetUnprotected(a);
  const s = assertRangeCellBudget(t);
  s.length > 500 &&
    failWithMessage("Select at most 500 cells to remove duplicates.");
  const d = parseRangeAddress(t);
  (!n.length || n.some((m) => m < d.firstCol || m > d.lastCol)) &&
    failWithMessage("Choose one or more columns inside the selected range.");
  for (const m of a.tables ?? []) {
    const v = parseRangeAddress(m.range);
    doesRangeOverlap(t, m.range) &&
      ((d.firstCol > v.firstCol || d.lastCol < v.lastCol) &&
        failWithMessage(
          "Include all table columns when removing duplicate table rows.",
        ),
      m.showTotals &&
        d.firstRow <= v.lastRow &&
        d.lastRow >= v.lastRow &&
        failWithMessage("Select data rows above the table totals row."),
      !o &&
        d.firstRow <= v.firstRow &&
        d.lastRow >= v.firstRow &&
        failWithMessage("Keep the table header row selected as headers."));
  }
  (a.merges ?? []).some((m) => doesRangeOverlap(t, m)) &&
    failWithMessage("Unmerge selected cells before removing duplicates.");
  const u = calculateWorkbookWithSettings(e),
    l = new Set(),
    c = [];
  let h = 0;
  for (let m = d.firstRow; m <= d.lastRow; m++) {
    const v = JSON.stringify(
      n.map((g) => u(r, `${columnIndexToLetters(g)}${m + 1}`)),
    );
    if (m !== d.firstRow || !o) {
      if (l.has(v)) {
        h++;
        continue;
      }
      l.add(v);
    }
    c.push(
      Array.from({ length: d.lastCol - d.firstCol + 1 }, (g, b) => {
        const y = structuredClone(
          a.cells[`${columnIndexToLetters(d.firstCol + b)}${m + 1}`] ?? {
            raw: null,
          },
        );
        return (
          typeof y.raw == "string" &&
            y.raw.startsWith("=") &&
            !y.literal &&
            (y.raw = shiftFormulaReferences(
              y.raw,
              d.firstRow + c.length - m,
              0,
            )),
          y
        );
      }),
    );
  }
  const f = { ...a.cells };
  for (const m of s) {
    const v = _t(m);
    f[m] = c[v.row - d.firstRow]?.[v.col - d.firstCol] ?? {
      ...f[m],
      raw: null,
    };
  }
  return {
    workbook: {
      ...e,
      sheets: e.sheets.map((m) => (m.id === r ? { ...m, cells: f } : m)),
    },
    removed: h,
  };
}
function findSheetOrThrow(e, r) {
  return (
    e.sheets.find((t) => t.id === r) ??
    throwWithMessage(`Sheet not found: ${r}`)
  );
}
function applyNamedStyleChange(e, r, t = {}) {
  t.readOnly && failWithMessage("This workbook is read-only.");
  for (const v of e.sheets)
    for (const g of Object.values(v.cells))
      assertNamedStyleBindingCurrent(e, g.style);
  const n = e.styleCatalog,
    o = cloneStyleValue(n ?? createDefaultStyleCatalog(e.id));
  let a = cloneStyleValue(r);
  const s = [],
    d = [],
    u = (v) =>
      o.styles.find((g) => g.id === v) ??
      failWithMessage("This named style no longer exists.");
  function l(v, g, b) {
    const y = findSheetOrThrow(e, v),
      w = y.cells[g];
    if (canonicalStyleKey(w) !== canonicalStyleKey(b)) {
      if (
        (assertSheetUnprotected(y), hasBorderEdgeChanges(w?.style, b.style))
      ) {
        const A = _t(g);
        y.merges?.some((k) => {
          const F = parseRangeAddress(k);
          return (
            A.row >= F.firstRow &&
            A.row <= F.lastRow &&
            A.col >= F.firstCol &&
            A.col <= F.lastCol
          );
        }) &&
          failWithMessage(
            "Unmerge the intersecting cells before changing named-style border edges.",
          );
      }
      (s.push({
        sheetId: v,
        address: g,
        ...(w ? { before: cloneStyleValue(w) } : {}),
        after: cloneStyleValue(b),
      }),
        s.length > 500 &&
          failWithMessage(
            "This action changes more than 500 cells. Use smaller selections or fewer bound cells.",
          ));
    }
  }
  function c(v) {
    for (const b of e.sheets)
      for (const [y, w] of Object.entries(b.cells))
        w.style?.binding &&
          (!v || w.style.binding.styleId === v) &&
          (assertSheetUnprotected(b),
          l(b.id, y, {
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
        y = validateNamedStyleDefinition(a.definition);
      (o.styles.push({ id: b, name: g, definition: y }),
        (a = { ...a, name: g, styleId: b, definition: y }),
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
      const v = a.styleId;
      ((o.styles = o.styles.filter((g) => g.id !== v)),
        c(a.styleId),
        d.push(
          "Cells using this style will show default formatting plus their direct overrides. Detach the style first to retain its current appearance.",
        ));
      break;
    case "applyStyle":
    case "reapplyStyle":
    case "detachStyle": {
      const g = findSheetOrThrow(e, a.sheetId);
      (assertSheetUnprotected(g), a.kind !== "detachStyle" && u(a.styleId));
      for (const b of enumerateSelectionCells(a.ranges)) {
        const y = g.cells[b] ?? { raw: null };
        if (a.kind === "detachStyle") {
          if (!y.style?.binding) continue;
          const w = resolveEffectiveCellStyle(y.style),
            E = { ...y, style: w };
          (Object.keys(w).length || delete E.style, l(g.id, b, E));
        } else {
          const w = resolveNamedStyleBinding(o, {
            namespace: o.namespace,
            styleId: a.styleId,
          });
          l(g.id, b, {
            ...y,
            style: { ...(a.kind === "applyStyle" ? y.style : {}), binding: w },
          });
        }
      }
      break;
    }
    default:
      failWithMessage("Choose a supported theme or named style action.");
  }
  const h = validateWorkbookStyleCatalog(o),
    f = a.kind === "detachStyle" && !n ? undefined : h,
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
    snapshot: JSON.stringify(e),
    request: a,
    ...(n ? { catalogBefore: cloneStyleValue(n) } : {}),
    ...(f ? { catalogAfter: f } : {}),
    changes: s,
    count: s.length,
    summary: p[a.kind],
    warnings: d,
  };
}
function assertSheetWritable(e, r, t) {
  (t.readOnly && throwWithMessage("This workbook is read-only."),
    assertSheetUnprotected(findSheetOrThrow(e, r)));
}
function nQ(e, r, t, n, o = "Reviewer", a = {}) {
  assertSheetWritable(e, r, a);
  const s = findSheetOrThrow(e, r),
    d = resolveReviewTimestamp(a),
    u = normalizeReviewAuthor(o),
    l = {
      id: crypto.randomUUID(),
      anchor: {
        sheetId: r,
        sheetName: s.name,
        address: normalizeCellAddress(t),
      },
      createdAt: d,
      updatedAt: d,
      messages: [
        {
          id: crypto.randomUUID(),
          author: u,
          text: normalizeBoundedText(
            n,
            "Comment",
            reviewThreadLimits.textLength,
            true,
          ),
          createdAt: d,
        },
      ],
    };
  return attachReviewThreadsToWorkbook(e, [...(e.reviewThreads ?? []), l]);
}
function findReviewThread(e, r, t) {
  t.readOnly && throwWithMessage("This workbook is read-only.");
  const n = e.reviewThreads?.find((a) => a.id === r);
  return n
    ? (n.orphaned &&
        throwWithMessage(
          "This discussion is retained for review because its target was deleted.",
        ),
      assertSheetWritable(e, n.anchor.sheetId, t),
      n)
    : throwWithMessage(
        "The discussion could not be found. Reopen the review tools.",
      );
}
function addReviewThreadReply(e, r, t, n = "Reviewer", o = {}) {
  const a = findReviewThread(e, r, o);
  a.resolved &&
    throwWithMessage("Reopen the discussion before adding a reply.");
  const s = resolveReviewTimestamp(o, a.updatedAt),
    d = {
      ...a,
      updatedAt: s,
      messages: [
        ...a.messages,
        {
          id: crypto.randomUUID(),
          author: normalizeReviewAuthor(n),
          text: normalizeBoundedText(
            t,
            "Reply",
            reviewThreadLimits.textLength,
            true,
          ),
          createdAt: s,
        },
      ],
    };
  return attachReviewThreadsToWorkbook(
    e,
    e.reviewThreads.map((u) => (u.id === r ? d : u)),
  );
}
function setReviewThreadResolved(e, r, t, n = "Reviewer", o = {}) {
  typeof t != "boolean" &&
    throwWithMessage("Choose whether to resolve or reopen the discussion.");
  const a = findReviewThread(e, r, o);
  if (!!a.resolved === t) return e;
  const s = resolveReviewTimestamp(o, a.updatedAt),
    d = { ...a, updatedAt: s };
  return (
    t
      ? (d.resolved = { author: normalizeReviewAuthor(n), createdAt: s })
      : (normalizeReviewAuthor(n), delete d.resolved),
    attachReviewThreadsToWorkbook(
      e,
      e.reviewThreads.map((u) => (u.id === r ? d : u)),
    )
  );
}
function setCellHyperlink(e, r, t, n, o = {}) {
  assertSheetWritable(e, r, o);
  const a = findSheetOrThrow(e, r),
    s = normalizeCellAddress(t),
    d = a.cells[s],
    u = _t(s);
  a.merges?.some((p) => {
    const m = parseRangeAddress(p);
    return (
      u.row >= m.firstRow &&
      u.row <= m.lastRow &&
      u.col >= m.firstCol &&
      u.col <= m.lastCol &&
      (u.row !== m.firstRow || u.col !== m.firstCol)
    );
  }) &&
    throwWithMessage(
      "Choose the merged area's first cell or unmerge it before changing its link.",
    );
  const l = n === null ? undefined : validateHyperlinkRecord(n);
  if (
    (l?.kind === "internal" && findSheetOrThrow(e, l.sheetId),
    JSON.stringify(d?.hyperlink) === JSON.stringify(l))
  )
    return e;
  !d &&
    e.sheets.reduce((p, m) => p + Object.keys(m.cells).length, 0) >=
      workbookLimits.maxCells &&
    throwWithMessage(
      "Adding this link would exceed the workbook populated-cell limit.",
    );
  const c = { ...(d ?? { raw: null }) };
  l ? (c.hyperlink = l) : delete c.hyperlink;
  const h = { ...a.cells, [s]: c };
  return (
    c.raw === null && Object.keys(c).length === 1 && delete h[s],
    { ...e, sheets: e.sheets.map((p) => (p.id === r ? { ...p, cells: h } : p)) }
  );
}
function collectHyperlinkEntries(e, r, t) {
  const n = findSheetOrThrow(e, t);
  let o = false;
  const a = Object.fromEntries(
    Object.entries(n.cells).map(([s, d]) =>
      d.hyperlink?.kind !== "internal" || d.hyperlink.sheetId !== r
        ? [s, d]
        : ((o = true),
          [s, { ...d, hyperlink: { ...d.hyperlink, sheetId: t } }]),
    ),
  );
  return o
    ? {
        ...e,
        sheets: e.sheets.map((s) => (s.id === t ? { ...s, cells: a } : s)),
      }
    : e;
}
function applyAutoSum(e, r, t) {
  const n = findSheetOrThrow(e, r),
    o = validateSelectionBounds(n, t);
  if (o.length !== 1) throw new Error("Select one empty total cell.");
  t = o[0];
  const a = _t(t),
    s = n.cells[t];
  if (s?.raw !== undefined && s.raw !== null && s.raw !== "")
    throw new Error("Select an empty cell for AutoSum.");
  const d = (h, f) => {
    if (h < 0 || f < 0) return false;
    const p = n.cells[`${columnIndexToLetters(f)}${h + 1}`];
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
  const c = cloneCellRecord(s);
  return (
    (c.raw = `=SUM(${u}:${l})`),
    delete c.literal,
    { sheetId: r, address: t, cell: c }
  );
}
function setTableColumnDataType(e, r, t, n) {
  const o = findSheetOrThrow(e, r),
    a = o.tables?.find((f) => f.id === t);
  if (!a) throw Error("Table not found.");
  const s = a.columns.findIndex((f) => f.id === n);
  if (s < 0) throw Error("Select an existing table column.");
  const d = parseRangeAddress(a.range),
    u = Math.max(0, d.lastRow - d.firstRow - Number(a.showTotals)),
    l = calculateWorkbookWithSettings(e),
    c = [];
  for (let f = 1; f <= Math.min(u, 500); f++) {
    const p = `${columnIndexToLetters(d.firstCol + s)}${d.firstRow + f + 1}`,
      m = l(r, p);
    coerceImportedValue(m, a.columns[s].dataType) ||
      c.push({ address: p, value: m });
  }
  return {
    dataType: a.columns[s].dataType ?? "any",
    checked: Math.min(u, 500),
    total: u,
    truncated: u > 500,
    exceptions: c,
  };
}
function findTableForRange(e, r, t) {
  const n = findSheetOrThrow(e, r),
    o = parseRangeAddress(t);
  if (measureRangeArea(o) !== 1) return toA1RangeAddress(t);
  const a = n.tables?.find((d) =>
    Dp(parseRangeAddress(d.range), o.firstRow, o.firstCol),
  );
  return a ? a.range : formatOccupiedRangeAddress(n, false);
}
function suggestTableHeaderName(e, r, t) {
  const n = findSheetOrThrow(e, r),
    o = resolveSourceRange(t, n),
    a = new Set(),
    s = [];
  for (let d = o.firstCol; d <= o.lastCol; d++) {
    const u = Cn(o.firstRow, d),
      l = n.cells[u],
      c = l?.raw ?? null;
    let h = `Column${d - o.firstCol + 1}`;
    try {
      isNullOrEmptyValue(c) || (h = validateColumnName(String(c)));
    } catch {}
    let f = h,
      p = 2;
    for (; a.has(f.toLowerCase());) f = `${h.slice(0, 250)}${p++}`;
    (a.add(f.toLowerCase()),
      s.push({
        address: u,
        original: c,
        suggested: f,
        requiresResolution: c !== f || (f.startsWith("=") && !l?.literal),
      }));
  }
  return s;
}
function validateReviewedTableHeaders(e, r, t, n) {
  const o = suggestTableHeaderName(e, r, t);
  if (n === undefined)
    return (
      o.some((s) => s.requiresResolution) &&
        throwWorkbookError(
          "Review table headers: resolve blank, duplicate, non-text or formula headers explicitly. Original header values will be retained in native source evidence.",
        ),
      o.map((s) => s.suggested)
    );
  (!Array.isArray(n) || n.length !== o.length) &&
    throwWorkbookError("Provide one reviewed header for every table column.");
  const a = n.map(validateColumnName);
  return (
    new Set(a.map((s) => s.toLowerCase())).size !== a.length &&
      throwWorkbookError(
        "Reviewed table headers must be unique, nonblank text.",
      ),
    a
  );
}
function renameTableHeaderField(e, r) {
  const t = [];
  for (const s of e.sheets) {
    const d = r.sheets.find((u) => u.id === s.id);
    if (d)
      for (const u of s.tables ?? []) {
        const l = d.tables?.find((m) => m.id === u.id);
        if (!l) continue;
        const c = parseRangeAddress(u.range),
          h = parseRangeAddress(l.range),
          f = {},
          p = l.columns.map((m, v) => {
            const g = u.columns.findIndex((k) => k.id === m.id);
            if (g < 0) return m;
            const b = u.columns[g],
              y = s.cells[Cn(c.firstRow, c.firstCol + g)],
              w = d.cells[Cn(h.firstRow, h.firstCol + v)],
              A =
                y?.raw !== w?.raw || y?.literal !== w?.literal
                  ? w && !isNullOrEmptyValue(w.raw)
                    ? validateColumnName(String(w.raw))
                    : throwWorkbookError(
                        "Table headers must be nonblank text. The header edit was not applied.",
                      )
                  : m.name;
            return (
              A !== b.name &&
                Object.defineProperty(f, b.name, {
                  value: A,
                  enumerable: true,
                }),
              A === m.name ? m : { ...m, name: A }
            );
          });
        (u.name !== l.name || Object.keys(f).length) &&
          (new Set(p.map((m) => m.name.toLowerCase())).size !== p.length &&
            throwWorkbookError(
              "Table headers must be unique, nonblank text. The header edit was not applied.",
            ),
          t.push({ sheetId: d.id, old: u, current: l, columns: p, names: f }));
      }
  }
  if (!t.length) return r;
  const n = structuredClone(r);
  for (const s of t) {
    const d = findSheetOrThrow(n, s.sheetId),
      u = d.tables.find((f) => f.id === s.current.id),
      l = (f) => (Object.hasOwn(s.names, f) ? s.names[f] : f);
    ((u.columns = s.columns),
      (u.totals = Object.fromEntries(
        Object.entries(u.totals).map(([f, p]) => [l(f), p]),
      )));
    const c = parseRangeAddress(u.range);
    u.columns.forEach((f, p) => {
      const m = Cn(c.firstRow, c.firstCol + p);
      d.cells[m] = {
        ...(d.cells[m] ?? { raw: f.name }),
        raw: f.name,
        ...(f.name.startsWith("=")
          ? { literal: true }
          : { literal: undefined }),
      };
    });
    const h = (f) =>
      f.tableId === u.id ||
      (!f.tableId &&
        f.sheetId === d.id &&
        toA1RangeAddress(f.range) === s.old.range);
    for (const f of n.sheets) {
      for (const p of f.tables ?? [])
        for (const m of p.columns)
          m.calculatedFormula &&= $v(
            m.calculatedFormula,
            s.old.name,
            u.name,
            s.names,
            p.id === u.id && f.id === d.id,
          );
      for (const [p, m] of Object.entries(f.cells)) {
        if (typeof m.raw != "string" || !m.raw.startsWith("=") || m.literal)
          continue;
        const v = _t(p),
          g = f.id === d.id && Dp(c, v.row, v.col);
        m.raw = $v(m.raw, s.old.name, u.name, s.names, g);
      }
      for (const p of f.conditionalFormats ?? [])
        if (p.kind === "formula" && p.formula) {
          const m = parseRangeAddress(p.range),
            v = $v(p.formula, s.old.name, u.name, s.names, false),
            g = $v(p.formula, s.old.name, u.name, s.names, true),
            b = f.id === d.id && rangesIntersect(c, m),
            y =
              b && Dp(c, m.firstRow, m.firstCol) && Dp(c, m.lastRow, m.lastCol);
          (v !== g &&
            b &&
            !y &&
            throwWorkbookError(
              "This conditional formula uses unqualified table fields across a table boundary. Split or remove its applied range before renaming the field.",
            ),
            (p.formula = y ? g : v));
        }
      for (const p of [...(f.pivots ?? []), ...(f.queries ?? [])])
        if (h(p.source))
          if (
            ((p.source = { ...p.source, tableId: u.id, range: u.range }),
            (p.filters &&= p.filters.map((m) => ({ ...m, field: l(m.field) }))),
            "rowFields" in p)
          ) {
            ((p.rowFields = p.rowFields.map(l)),
              (p.columnFields = p.columnFields.map(l)),
              (p.values = p.values.map((m) => ({ ...m, field: l(m.field) }))),
              (p.dateGroups &&= p.dateGroups.map((m) => ({
                ...m,
                field: l(m.field),
              }))));
            for (const m of ["slicers", "timelines", "sorts"])
              p[m] &&= p[m].map((v) => ({ ...v, field: l(v.field) }));
            p.calculatedFields &&= p.calculatedFields.map((m) => ({
              ...m,
              field: l(m.field),
              ...(m.otherField ? { otherField: l(m.otherField) } : {}),
            }));
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
    normalizeTableDefinitions(d, d);
  }
  assertUniqueTableNames(n);
  for (const s of e.sheets) {
    if (!s.protection?.enabled) continue;
    const d = n.sheets.find((u) => u.id === s.id);
    JSON.stringify(s) !== JSON.stringify(d) &&
      throwWorkbookError(
        `This change would update dependent cells or analysis on protected worksheet ${s.name}. Unprotect it before renaming the table field.`,
      );
  }
  return n;
}
function buildPivotSource(e, r) {
  const t = normalizeSourceReference(r),
    n = findSheetOrThrow(e, t.sheetId),
    o = t.tableId
      ? (n.tables?.find((h) => h.id === t.tableId) ??
        throwWorkbookError("The source table no longer exists."))
      : n.tables?.find((h) => h.range === toA1RangeAddress(t.range)),
    a = resolveSourceRange(o?.range ?? t.range, n),
    s = calculateWorkbookWithSettings(e);
  o?.showTotals && a.lastRow--;
  const d = [];
  for (let h = a.firstCol; h <= a.lastCol; h++)
    d.push(validateColumnName(s(n.id, Cn(a.firstRow, h))));
  new Set(d.map((h) => h.toLowerCase())).size !== d.length &&
    throwWorkbookError(
      "Source headers must be nonblank, unique text. Create a table to normalize blank or duplicate headers.",
    );
  const u = [];
  for (let h = a.firstRow + 1; h <= a.lastRow; h++) {
    const f = d.map((p, m) => s(n.id, Cn(h, a.firstCol + m)));
    (f.forEach((p, m) => {
      const v = Cn(h, a.firstCol + m);
      typeof p == "string" &&
        p.startsWith("#") &&
        s.evaluateFormula(n.id, v, `=ISERROR(${v})`) === true &&
        throwWorkbookError(
          `Source calculation error ${p} at ${n.name}!${v}. Repair the source calculation before building or refreshing analysis; no totals were written.`,
        );
    }),
      (f.some((p) => !isNullOrEmptyValue(p)) ||
        d.some(
          (p, m) =>
            !isNullOrEmptyValue(n.cells[Cn(h, a.firstCol + m)]?.raw ?? null),
        )) &&
        u.push({ index: h, values: f }));
  }
  return (
    s.getOperationalErrors().length &&
      throwWorkbookError(
        `Source calculation is incomplete (${s.getOperationalErrors().join(", ")}). Repair blocked or over-limit dependencies before building analysis.`,
      ),
    {
      fields: d,
      rows: u,
      source: {
        ...t,
        range: o?.range ?? toA1RangeAddress(t.range),
        ...(o ? { tableId: o.id } : {}),
      },
    }
  );
}
function getPivotSourceFields(e, r, t) {
  return buildPivotSource(e, { sheetId: r, range: t }).fields;
}
function aggregatePivotRows(e, r, t) {
  const n = findSheetOrThrow(e, t.source.sheetId),
    o = parseRangeAddress(t.source.range),
    a = Object.entries(n.cells)
      .filter(([u]) => {
        const l = _t(u);
        return Dp(o, l.row, l.col);
      })
      .sort(([u], [l]) => u.localeCompare(l))
      .map(([u, l]) => [u, l.raw, !!l.literal]),
    { destination, ...d } = normalizePivotDefinition(r);
  return computePivotRefreshFingerprint({
    source: t.source,
    fields: t.fields,
    rows: t.rows,
    cells: a,
    settings: d,
  });
}
function getPivotRefreshState(e, r) {
  const t = e.sheets.find((s) => s.pivots?.some((u) => u.id === r)),
    n = t?.pivots?.find((s) => s.id === r);
  if (
    ((!t || !n) && throwWorkbookError("Pivot definition not found."),
    !n.refresh)
  )
    return {
      state: "unverified",
      reason:
        "This older result has no refresh evidence. Refresh it before relying on its totals.",
      outputModified: false,
      definition: n,
    };
  const o =
    collectRangeValues(t, n.outputRange) !== n.refresh.outputFingerprint;
  try {
    const s = buildPivotSource(e, n.source);
    validatePivotFieldReferences(n, s.fields);
    const d = aggregatePivotRows(e, n, s) !== n.refresh.sourceFingerprint;
    return {
      state: d || o ? "stale" : "current",
      reason: o
        ? "Output cells were edited. Refresh into a new result before relying on these totals."
        : d
          ? "Source content, calculated values, range or settings changed. Refresh into a new result."
          : "Source content and output match the saved refresh snapshot.",
      outputModified: o,
      definition: n,
    };
  } catch (s) {
    return {
      state: "unavailable",
      reason: s.message,
      outputModified: o,
      definition: n,
    };
  }
}
function describePivotRefresh(e, r) {
  const { definition } = findPivotDefinition(e, r),
    n = getPivotRefreshState(e, r);
  return buildPivotStatusRecord(e, definition, n);
}
function inspectPivotControlFields(e, r, t) {
  t.length > 30 &&
    throwWorkbookError("Inspect at most 30 pivot control fields together.");
  const { definition } = findPivotDefinition(e, r),
    o = buildPivotSource(e, definition.source);
  return new Map(
    [...new Set(t)].map((a) => {
      const s = o.fields.indexOf(a),
        d = new Map();
      if (s < 0)
        return [
          a,
          {
            items: [],
            error: `Broken pivot field “${a}”: source field was not found. Restore or reselect it.`,
          },
        ];
      for (const u of o.rows)
        if ((d.set(buildSlicerItemKey(u.values[s]), u.values[s]), d.size > 500))
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
function getPivotSlicerItems(e, r, t) {
  const n = inspectPivotControlFields(e, r, [t]).get(t);
  return (n.error && throwWorkbookError(n.error), n.items);
}
function buildPivotModel(e, r) {
  const { owner, definition } = findPivotDefinition(e, r),
    o = normalizePivotDefinition(definition),
    a = buildPivotSource(e, definition.source);
  validatePivotFieldReferences(o, a.fields);
  const s = filterPivotSourceRows(
    o,
    a.fields,
    filterSourceRows(a, o.filters ?? []),
  );
  return (
    (!definition.refresh ||
      aggregatePivotRows(e, o, a) !== definition.refresh.sourceFingerprint ||
      collectRangeValues(owner, definition.outputRange) !==
        definition.refresh.outputFingerprint) &&
      throwWorkbookError(
        "This PivotTable is stale or unverified. Refresh it before drilling or charting; historical totals cannot be matched to current source rows.",
      ),
    {
      owner,
      definition,
      config: o,
      data: a,
      rows: s,
      built: computePivotTable(o, a.fields, s),
    }
  );
}
function buildPivotChartModel(e, r) {
  const { definition, built } = buildPivotModel(e, r),
    o = definition.chart?.valueIndex ?? 0;
  built.chartRows.length > 200 &&
    throwWorkbookError(
      "PivotCharts are limited to 200 leaf groups. Filter or regroup before charting.",
    );
  const a = built.chartRows.map((u) => ({
    label: u.label,
    value: u.values[o],
  }));
  return (
    a.some((u) => u.value !== null && typeof u.value != "number") &&
      throwWorkbookError(
        "The PivotChart contains a calculation error. Resolve zero denominators or invalid source calculations.",
      ),
    {
      kind: definition.chart?.kind ?? "column",
      measure: built.measures[o].name,
      showAs: definition.values[o].showAs ?? "value",
      points: a,
      sourceRevision: definition.sourceRevision,
    }
  );
}
function buildPivotDrillDown(e, r, t) {
  const { definition, config, data, rows, built } = buildPivotModel(e, r),
    u = _t(t),
    l = parseRangeAddress(definition.outputRange),
    c = u.row - l.firstRow - 1,
    h = u.col - l.firstCol - Math.max(1, config.rowFields.length);
  (c < 0 ||
    c >= built.rowAxes.length ||
    h < 0 ||
    u.col > l.lastCol ||
    u.row > l.lastRow) &&
    throwWorkbookError(
      "Choose a PivotTable value cell, subtotal or grand total; headers and group labels cannot be drilled.",
    );
  const f = built.rowAxes[c],
    p = built.columnAxes[Math.floor(h / config.values.length)],
    m = h % config.values.length,
    v = (E, A, k) =>
      !A ||
      A.parts.every((F, I) => {
        const O = E.values[findSourceFieldIndex(data.fields, k[I])],
          j = config.dateGroups?.find((V) => V.field === k[I]);
        return (
          JSON.stringify(
            j
              ? bucketPivotDateValue(O, j, E.index)
              : isNullOrEmptyValue(O)
                ? null
                : O,
          ) === JSON.stringify(F)
        );
      }),
    g = rows.filter(
      (E) => v(E, f, config.rowFields) && v(E, p, config.columnFields),
    ),
    b = findSheetOrThrow(e, data.source.sheetId),
    y = parseRangeAddress(data.source.range),
    w = g.map((E) => ({
      locator: {
        workbookId: e.id,
        sheetId: b.id,
        sheetName: b.name,
        row: E.index + 1,
        range: `${columnIndexToLetters(y.firstCol)}${E.index + 1}:${columnIndexToLetters(y.lastCol)}${E.index + 1}`,
        original: findRowOrigin(e, b.id, E.index),
      },
      values: [...E.values],
      cells: data.fields.map((A, k) => ({
        address: Cn(E.index, y.firstCol + k),
        cell: structuredClone(
          b.cells[Cn(E.index, y.firstCol + k)] ?? { raw: null },
        ),
      })),
    }));
  return {
    context: buildPivotStatusRecord(e, definition, {
      state: "current",
      reason: "Source and output fingerprints match the saved calculation.",
    }),
    address: `${columnIndexToLetters(u.col)}${u.row + 1}`,
    valueIndex: m,
    measure: config.values[m],
    underlyingValue: built.baseMatrix[c + 1][u.col - l.firstCol],
    displayedValue: built.matrix[c + 1][u.col - l.firstCol],
    rowGroup: f?.parts ?? null,
    columnGroup: p?.parts ?? null,
    fields: [...data.fields],
    records: w,
  };
}
function validateTotalsRowForAppend(e, r, t) {
  const n = findSheetOrThrow(e, r),
    o = parseRangeAddress(t.range),
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
            ? p.toLowerCase() !== n.name.toLowerCase()
            : u !== undefined && u !== r
        )
          continue;
        const m = Number(f[3].replace("$", "")) - 1,
          v = Number(f[4].replace("$", "")) - 1;
        o.lastRow >= Math.min(m, v) &&
          o.lastRow <= Math.max(m, v) &&
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
            ? p.toLowerCase() !== n.name.toLowerCase()
            : u !== undefined && u !== r
        )
          continue;
        const m = parseRangeAddress(`${f[3]}:${f[4] ?? f[3]}`);
        rangesIntersect(m, { ...o, firstRow: o.lastRow }) &&
          throwWorkbookError(
            "A formula refers directly to the moving totals row. Use a structured table totals reference or move those references before appending.",
          );
      }
    };
  for (const d of e.sheets) {
    for (const [u, l] of Object.entries(d.cells)) {
      const c = _t(u);
      if (
        !(
          d.id === r &&
          c.row === o.lastRow &&
          c.col >= o.firstCol &&
          c.col <= o.lastCol
        ) &&
        (typeof l.raw == "string" &&
          l.raw.startsWith("=") &&
          !l.literal &&
          a(l.raw, d.id),
        l.hyperlink?.kind === "internal" && l.hyperlink.sheetId === r)
      ) {
        const h = _t(l.hyperlink.address);
        h.row === o.lastRow &&
          h.col >= o.firstCol &&
          h.col <= o.lastCol &&
          throwWorkbookError(
            "An internal hyperlink targets the moving totals row. Retarget or remove it before appending.",
          );
      }
    }
    for (const u of d.conditionalFormats ?? [])
      u.kind === "formula" && u.formula && a(u.formula, d.id);
    for (const u of d.tables ?? [])
      for (const l of u.columns)
        l.calculatedFormula && a(l.calculatedFormula, d.id);
  }
  for (const d of e.definedNames ?? []) a(d.refersTo, d.scopeSheetId);
  for (let d = o.firstCol; d <= o.lastCol; d++) {
    const u = n.cells[Cn(o.lastRow, d)],
      l = buildTableTotalCell(t, d - o.firstCol);
    u &&
      (u.raw !== l.raw || u.literal) &&
      throwWorkbookError(
        "The totals row was edited. Preserve its values before appending.",
      );
  }
}
function resizeTableRange(e, r, t, n, o) {
  const a = findSheetOrThrow(e, r),
    s = parseRangeAddress(t.range);
  (o !== t.showTotals &&
    throwWorkbookError(
      "Clear calculated column rules before changing totals-row visibility.",
    ),
    (n.firstRow !== s.firstRow ||
      n.firstCol !== s.firstCol ||
      n.lastCol !== s.lastCol ||
      n.lastRow < s.lastRow) &&
      throwWorkbookError(
        "Calculated tables support downward resize only; keep the existing header and columns. Clear rules before shrinking or changing columns.",
      ),
    (n.lastRow - s.lastRow) * (n.lastCol - n.firstCol + 1) +
      (t.showTotals ? t.columns.length : 0) >
      tableMutationLimits.changedCells &&
      throwWorkbookError("Table append/resize may change at most 500 cells."));
  for (let u = s.lastRow + 1; u <= n.lastRow; u++)
    for (let l = n.firstCol; l <= n.lastCol; l++)
      a.cells[Cn(u, l)] &&
        throwWorkbookError(
          "Table append/resize would overwrite adjacent cells or their formatting. Move that content first.",
        );
  if (t.showTotals && n.lastRow !== s.lastRow) {
    validateTotalsRowForAppend(e, r, t);
    for (const u of [...(e.formulaWatches ?? []), ...(e.reviewThreads ?? [])]) {
      if (u.orphaned || u.anchor.sheetId !== r) continue;
      const l = _t(u.anchor.address);
      l.row === n.lastRow &&
        l.col >= n.firstCol &&
        l.col <= n.lastCol &&
        throwWorkbookError(
          "The new totals row already has a saved watch or review discussion. Preserve or remove that destination anchor before appending.",
        );
    }
  }
}
function moveTotalsRowCells(e, r, t, n, o) {
  const a = parseRangeAddress(n.range),
    s = parseRangeAddress(o.range),
    d = findSheetOrThrow(e, t),
    u = findSheetOrThrow(r, t);
  for (let c = a.firstCol; c <= a.lastCol; c++) {
    const h = Cn(a.lastRow, c),
      f = Cn(s.lastRow, c),
      p = d.cells[h];
    if (p) {
      const m = {
        ...structuredClone(p),
        ...buildTableTotalCell(o, c - a.firstCol),
      };
      (delete m.literal, (u.cells[f] = m));
    }
  }
  const l = (c) => {
    if (c.sheetId !== t) return;
    const h = _t(c.address);
    h.row === a.lastRow &&
      h.col >= a.firstCol &&
      h.col <= a.lastCol &&
      (c.address = Cn(s.lastRow, h.col));
  };
  for (const c of r.formulaWatches ?? []) c.orphaned || l(c.anchor);
  for (const c of r.reviewThreads ?? []) c.orphaned || l(c.anchor);
}
function resolveTableDataRange(e, r, t = "A1") {
  const n = findSheetOrThrow(e, r),
    o = parseRangeAddress(t);
  if (o.firstRow !== o.lastRow || o.firstCol !== o.lastCol) return t;
  const a = n.tables?.find((u) => {
    const l = parseRangeAddress(u.range);
    return (
      o.firstRow >= l.firstRow &&
      o.firstRow <= l.lastRow &&
      o.firstCol >= l.firstCol &&
      o.firstCol <= l.lastCol
    );
  });
  if (!a) return formatOccupiedRangeAddress(n, false);
  const s = parseRangeAddress(a.range);
  return `${columnIndexToLetters(s.firstCol)}${s.firstRow + 1}:${columnIndexToLetters(s.lastCol)}${s.lastRow + 1 - (a.showTotals ? 1 : 0)}`;
}
function validateFilterRequest(e, r, t) {
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
  const n = findSheetOrThrow(e, r),
    o = parseRangeAddress(t.range),
    a = o.lastRow - o.firstRow,
    s = o.lastCol - o.firstCol + 1;
  if (o.lastRow >= n.rowCount || o.lastCol >= n.colCount || a < 1)
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
    n.merges?.some((f) => {
      const p = parseRangeAddress(f);
      return (
        p.firstRow <= o.lastRow &&
        p.lastRow >= o.firstRow &&
        p.firstCol <= o.lastCol &&
        p.lastCol >= o.firstCol
      );
    })
  )
    throw Error(
      "Unmerge cells in the filter range first; records cannot be split by filtering.",
    );
  const d = (f) => Number.isInteger(f) && f >= o.firstCol && f <= o.lastCol,
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
        resolveDatePolicyContext(f.value, f.type, e) === null
      )
        throw Error(
          `Invalid ${f.type} filter value. Use a finite number or a date in the workbook's date policy.`,
        );
    };
  let l = 0;
  const c = new Set();
  for (const f of t.columns) {
    if (
      !f ||
      !d(f.column) ||
      c.has(f.column) ||
      (f.join !== undefined && !["and", "or"].includes(f.join))
    )
      throw Error("Filter columns must be valid and distinct.");
    if (
      (c.add(f.column),
      f.values !== undefined &&
        (!Array.isArray(f.values) ||
          f.values.length > filterEvaluationLimits.values ||
          !f.values.every(isStorableCellValue)))
    )
      throw Error("Value filters allow up to 1,000 finite typed values.");
    if (
      f.conditions !== undefined &&
      (!Array.isArray(f.conditions) || f.conditions.length > 2)
    )
      throw Error("Each column supports up to two custom conditions.");
    for (const p of f.conditions ?? []) {
      if (p.column !== f.column)
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
    range: `${columnIndexToLetters(o.firstCol)}${o.firstRow + 1}:${columnIndexToLetters(o.lastCol)}${o.lastRow + 1}`,
  });
}
function applyFilterConditions(e, r, t, n = calculateWorkbookWithSettings(e)) {
  const o = validateFilterRequest(e, r, t),
    a = findSheetOrThrow(e, r),
    s = parseRangeAddress(o.range),
    d = !!o.caseSensitive,
    u = createAxisVisibilityIndex(a.rowCount, cn(a, "row")),
    l = [],
    c = [],
    h = [],
    f = [],
    p = new Set(),
    m = {},
    v = o.columns.map((O) =>
      O.values === undefined
        ? null
        : new Set(O.values.map((j) => comparableCellValueKey(j, d))),
    );
  let g = 0,
    b = 0,
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
      return matchesTextFilter(O, j, e, d);
    },
    E = Array.from({ length: s.lastCol - s.firstCol + 1 }, (O, j) => ({
      column: s.firstCol + j,
      label:
        String(
          assertCalculatedBeforeFilter(
            n,
            r,
            `${columnIndexToLetters(s.firstCol + j)}${s.firstRow + 1}`,
            e,
          ) ?? "",
        ) || columnIndexToLetters(s.firstCol + j),
    }));
  for (let O = s.firstRow + 1; O <= s.lastRow; O++) {
    const j = E.map((B) =>
        assertCalculatedBeforeFilter(
          n,
          r,
          `${columnIndexToLetters(B.column)}${O + 1}`,
          e,
        ),
      ),
      z = (B) => j[B - s.firstCol];
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
    (l.push(O + 1), a.rowOriginIds?.[O] && (m[O + 1] = a.rowOriginIds[O]));
    const V = u.isHidden(O);
    V && g++;
    let D =
      o.columns.every(
        (B, _) =>
          (v[_] === null || v[_].has(comparableCellValueKey(z(B.column), d))) &&
          (!B.conditions?.length ||
            (B.join === "or"
              ? B.conditions.some((T) => w(z(B.column), T))
              : B.conditions.every((T) => w(z(B.column), T)))),
      ) &&
      (!o.advanced?.length ||
        o.advanced.some((B) => B.every((_) => w(z(_.column), _))));
    if (D && o.unique) {
      const B = JSON.stringify(j.map((_) => comparableCellValueKey(_, d)));
      p.has(B) ? (D = false) : p.add(B);
    }
    if (D) (c.push(O + 1), V ? b++ : h.push(O + 1));
    else {
      const B = f.at(-1);
      B?.end === O - 1 ? (B.end = O) : f.push({ start: O, end: O });
    }
  }
  const A = [];
  for (const O of [...cn(a, "row"), ...f].sort((j, z) => j.start - z.start)) {
    const j = A.at(-1);
    j && O.start <= j.end + 1
      ? (j.end = Math.max(j.end, O.end))
      : A.push({ ...O });
  }
  return {
    hiddenRows: normalizeHiddenIntervals(A, a.rowCount),
    fields: E,
    excludedRows: f,
    context: {
      workbookId: e.id,
      revision: e.revision,
      sheetId: r,
      sheetName: a.name,
      range: o.range,
      definition: o,
      totalRows: l.length,
      matchedRows: c.length,
      visibleRows: h.length,
      filteredRows: l.length - c.length,
      manuallyHiddenRows: g,
      hiddenMatchingRows: b,
      sourceRows: l,
      matchingSourceRows: c,
      visibleSourceRows: h,
      originalSourceIds: m,
      policy: savedFilterViewHelpText,
    },
  };
}
function buildFilterFacetValues(e, r, t, n, o = "") {
  const a = validateFilterRequest(e, r, t),
    s = parseRangeAddress(a.range);
  if (!Number.isInteger(n) || n < s.firstCol || n > s.lastCol)
    throw Error("Facet column is outside the filter range.");
  const d = calculateWorkbookWithSettings(e),
    u = applyFilterConditions(
      e,
      r,
      { ...a, unique: false, columns: a.columns.filter((c) => c.column !== n) },
      d,
    ),
    l = new Map();
  for (const c of u.context.matchingSourceRows) {
    const h = assertCalculatedBeforeFilter(
        d,
        r,
        `${columnIndexToLetters(n)}${c}`,
        e,
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
function validateSavedFilterViews(e, r) {
  if (e === undefined) return;
  if (
    !Array.isArray(e) ||
    e.length > savedFilterViewLimits.views ||
    JSON.stringify(e).length > savedFilterViewLimits.characters
  )
    throw Error(
      "Saved filter views are limited to 20 views and 512,000 metadata characters per workbook.",
    );
  const t = new Set(),
    n = new Set();
  r.savedFilterViews = e.map((o) => {
    if (
      (validateSavedFilterViewMetadata(o, [
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
      !isValidFilterViewText(o.id, 100) ||
        !isValidFilterViewText(o.sheetId, 100) ||
        !isValidFilterViewText(o.sheetName, 31) ||
        !isValidFilterViewText(o.sourceWorkbookId, 100) ||
        !Number.isSafeInteger(o.sourceRevision) ||
        o.sourceRevision < 0 ||
        o.sourceRevision > Number.MAX_SAFE_INTEGER - 1 ||
        typeof o.savedAt != "string" ||
        !Number.isFinite(Date.parse(o.savedAt)) ||
        new Date(o.savedAt).toISOString() !== o.savedAt)
    )
      throw Error("Invalid saved filter view identity, revision or timestamp.");
    const a = normalizeSavedFilterViewName(o.name);
    if (t.has(o.id) || n.has(normalizeNameForComparison(a)))
      throw Error("Saved filter view names and IDs must be unique.");
    (t.add(o.id),
      n.add(normalizeNameForComparison(a)),
      validateSavedFilterViewMetadata(o.datePolicy, [
        "calendar",
        "dateSystem",
        "textDateOrder",
        "timeZone",
      ]));
    const s = normalizeDatePolicy(o.datePolicy);
    validateSavedFilterViewMetadata(o.definition, [
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
    if (!Array.isArray(o.definition.columns))
      throw Error("Invalid saved filter columns.");
    for (const h of o.definition.columns)
      if (
        (validateSavedFilterViewMetadata(h, [
          "column",
          "values",
          "conditions",
          "join",
        ]),
        h.conditions !== undefined)
      ) {
        if (!Array.isArray(h.conditions))
          throw Error("Invalid saved filter conditions.");
        h.conditions.forEach(d);
      }
    if (o.definition.advanced !== undefined) {
      if (!Array.isArray(o.definition.advanced))
        throw Error("Invalid saved advanced criteria.");
      for (const h of o.definition.advanced) {
        if (!Array.isArray(h)) throw Error("Invalid saved advanced criteria.");
        h.forEach(d);
      }
    }
    const u = {
        ...r,
        datePolicy: s,
        sheets: [
          {
            id: o.sheetId,
            name: o.sheetName,
            rowCount: workbookLimits.maxRows,
            colCount: workbookLimits.maxCols,
            cells: {},
          },
        ],
      },
      l = validateFilterRequest(u, o.sheetId, o.definition);
    validateSavedFilterViewMetadata(o.population, [
      "totalRows",
      "matchedRows",
      "visibleRows",
      "hiddenMatchingRows",
    ]);
    const c = o.population;
    if (
      Object.keys(c).length !== 4 ||
      Object.values(c).some((h) => !Number.isInteger(h) || Number(h) < 0) ||
      c.totalRows !==
        parseRangeAddress(l.range).lastRow -
          parseRangeAddress(l.range).firstRow ||
      c.totalRows > 1e4 ||
      c.totalRows < 1 ||
      c.matchedRows > c.totalRows ||
      c.visibleRows > c.matchedRows ||
      c.hiddenMatchingRows !== c.matchedRows - c.visibleRows
    )
      throw Error("Invalid saved population counts.");
    return {
      id: o.id,
      name: a,
      sheetId: o.sheetId,
      sheetName: o.sheetName,
      sourceWorkbookId: o.sourceWorkbookId,
      sourceRevision: o.sourceRevision,
      savedAt: o.savedAt,
      datePolicy: s,
      definition: l,
      population: { ...c },
    };
  });
}
function copyWithSavedFilterViews(e, r) {
  const t = { ...e };
  return (validateSavedFilterViews(r, t), t);
}
function createSavedFilterView(e, r, t, n, o) {
  const a = findSheetOrThrow(e, r);
  if (
    (assertSavedFilterViewsEditable(e, r),
    o && findSavedFilterView(e, o).sheetId !== r)
  )
    throw Error("A saved view must keep its source worksheet.");
  const s = applyFilterConditions(e, r, t).context,
    d = {
      id: o ?? crypto.randomUUID(),
      name: normalizeSavedFilterViewName(n),
      sheetId: r,
      sheetName: a.name,
      sourceWorkbookId: e.id,
      sourceRevision: e.revision,
      savedAt: new Date().toISOString(),
      datePolicy: normalizeDatePolicy(e.datePolicy),
      definition: s.definition,
      population: {
        totalRows: s.totalRows,
        matchedRows: s.matchedRows,
        visibleRows: s.visibleRows,
        hiddenMatchingRows: s.hiddenMatchingRows,
      },
    };
  return copyWithSavedFilterViews(e, [
    ...(e.savedFilterViews ?? []).filter((u) => u.id !== o),
    d,
  ]);
}
function addSavedFilterView(e, r, t) {
  return (
    assertSavedFilterViewsEditable(e, findSavedFilterView(e, r).sheetId),
    copyWithSavedFilterViews(
      e,
      e.savedFilterViews.map((n) =>
        n.id === r ? { ...n, name: normalizeSavedFilterViewName(t) } : n,
      ),
    )
  );
}
function removeSavedFilterView(e, r) {
  return (
    assertSavedFilterViewsEditable(e, findSavedFilterView(e, r).sheetId),
    copyWithSavedFilterViews(
      e,
      e.savedFilterViews.filter((t) => t.id !== r),
    )
  );
}
function applyPrintLayoutChange(e, r, t, n = {}) {
  if (n.readOnly) throw Error("This workbook is read-only.");
  n.expectedSnapshot && verifyPrintSnapshotGuard(e, n.expectedSnapshot);
  const o = findSheetOrThrow(e, r);
  assertSheetUnprotected(o);
  const a = { ...o };
  return (
    validatePageSetup(t, a),
    zU(a, a.pageLayout),
    JSON.stringify(a.pageLayout) === JSON.stringify(o.pageLayout)
      ? e
      : { ...e, sheets: e.sheets.map((s) => (s.id === r ? a : s)) }
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
