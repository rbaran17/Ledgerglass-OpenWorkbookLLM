import {
  strFromU8 as Ao,
  strToU8 as za,
  zipSync as Rd,
  unzipSync as Nd,
  Unzip as gc,
  UnzipInflate as mc,
} from "fflate";
import * as X from "zod";
import __lib_saxes from "saxes";

import __lib_exceljs from "exceljs";
import __lib_papaparse from "papaparse";
import {
  copyExtensionCacheEntries,
  dateSystemMaxSerial,
  dateSystemSerialOffset,
  dateToSerial,
  deepFrozenValueSet as $V,
  defaultDatePolicy,
  fingerprintCanonicalValue,
  hasLinkedOrBoundCells,
  isDeepFrozenValue,
  linkedOrBoundCellCache as VV,
  linkedOrBoundCellCache as bN,
  linkedOrBoundCellCache as eq,
  linkedOrBoundCellCache as r8,
  millisecondsPerDay,
  millisecondsPerDay as yq,
  normalizeDatePolicy,
  objectDefineProperty,
  serialZeroUtcMs as vq,
  sha256HexDigest,
  validateExtensionState,
  validateSavedRecipeMetadata,
} from "../app/01-core-schemas-and-hashing.js";
import {
  fileImportLimits,
  supportedFontFamilies,
  workbookLimits,
} from "../app/02-financial-functions.js";
import {
  assertKnownStyleKeys,
  borderLineStyles,
  canonicalStyleKey,
  cellBorderEdges,
  cloneStyleValue,
  cn,
  collectStyleFallbackWarnings,
  createAxisVisibilityIndex,
  createWorkbookMetadata,
  failWithMessage as Fr,
  isNonArrayObject as Gs,
  isNonArrayObject as yg,
  isPlainObjectLiteral as bo,
  normalizeHiddenIntervals,
  officeThemeDefinition,
  resolveNamedStyleBinding,
  throwWithMessage as jt,
  throwWorkbookError as Tk,
  validateBoundedStyleText,
  validateCellStyle,
  validateNamedStyleBinding,
  validateNamedStyleDefinition,
  validateOutlineGroups,
  validateSplitPaneState,
  validateWorkbookMetadata,
  validateWorkbookTheme,
  worksheetCapacityLimits,
} from "../app/03-workbook-model/01-cell-styles-and-themes.js";

import {
  Kne as _N,
  createUuid,
  createUuid as GN,
  insertDefinedNamesIntoWorkbookXml,
  maxPictureHashBytes,
  spreadsheetMlNamespace as QV,
  stripFormulaStringLiterals,
  throwXlsxDefinedNameError,
  validateDefinedName,
  validateNameSheetReferences,
  validateWorkbookStructureProtection,
  validateWorksheetVisibility,
} from "../app/03-workbook-model/03-defined-names.js";
import {
  assertUniqueTableNames,
  attachWorkbookLinkMetadata,
  basicAggregateOperations,
  createPictureEvidence,
  inspectPictureDataUrl,
  isBoundedString,
  normalizeRowProvenance,
} from "../app/03-workbook-model/04-pivot-tables-and-pictures.js";
import {
  defaultCalculationSettings,
  externalSourceErrorMessages,
  externalSourceLimits,
  importFormulaCacheEvidence,
  sourceRefreshStateSchema,
  verifyConsolidationEvidence,
  zie as dR,
} from "../app/03-workbook-model/05-consolidation-and-sources.js";
import {
  P2,
  alternateFormatExtensions,
  automationLimits,
  columnIndexToLetters,
  hashRecordExcludingHash,
  nonEmptyLabelSchema,
  normalizeAlternateFormatSource,
  readRetainedXlsxSource,
  reportOutputLimits,
  requireAutomationFileName,
  requireBoundedText,
  requireRetainedXlsxSource,
  retainAlternateFormatSource,
  retainOriginalXlsx,
  shiftFormulaReferences,
  sourceImportLimits,
  validateAlternateSource,
  validateAnalysisReport,
  validateAutomationArtifacts,
  validateDataModelReport,
  validateDelimitedSourceFootprint,
  validateExtractedSourceTables,
  validateFieldTransformRecipe,
  validatePrintPageBreaks,
  validateRetainedSourceFile,
  validateSourceImports,
  xlsxInspectionLimits,
} from "../app/03-workbook-model/06-data-model-and-analysis.js";
import {
  CE as Vm,
  _t,
  aP,
  analysisOutputEvidenceSchema,
  createAnchoredDrawingObject,
  normalizeReviewMetadata,
  parseRangeAddress,
  validateFormulaWatches,
  validateSavedWhatIfAnalyses,
  validateWorksheetFeatures,
} from "../app/03-workbook-model/07-conditional-formatting.js";
import {
  Cn as Ks,
  normalizeTableDefinitions,
  resolveSourceRange,
  tablePaletteStyleNames,
  validateColumnName,
  validateDrawingObjects,
  validateTableName,
} from "../app/03-workbook-model/08-chart-model.js";
import {
  calculateWorkbookWithSettings,
  calculationReportSchema,
  importedCacheEvidenceSchema,
  summarizeValidationInspection,
  validateObjectReportRequest,
} from "../app/03-workbook-model/09-calculation-engine.js";
import {
  createDefaultPageSetup,
  validatePageSetup,
  validateSavedFilterViews,
} from "../app/03-workbook-model/10-print-layout.js";
import {
  normalizeCellRecord,
  normalizeCellRecordInput,
  normalizeTextCleanups,
  validateTableSourceEvidence,
} from "../app/03-workbook-model/11-table-model.js";
import {
  excelMessagePatterns,
  fileProcessingErrorMessages,
  unsupportedXlsxMarkers,
} from "../app/06-analysis-tools.js";

var T8 = Object.create;

var A8 = Object.getOwnPropertyDescriptor;
var I8 = Object.getOwnPropertyNames;
var R8 = Object.getPrototypeOf,
  N8 = Object.prototype.hasOwnProperty;
var Mc = ((e) =>
  typeof require < "u"
    ? require
    : typeof Proxy < "u"
      ? new Proxy(e, { get: (t, n) => (typeof require < "u" ? require : t)[n] })
      : e)(function (e) {
  if (typeof require < "u") return require.apply(this, arguments);
  throw Error('Dynamic require of "' + e + '" is not supported');
});
var O8 = (e, t, n, i) => {
  if ((t && typeof t == "object") || typeof t == "function")
    for (let r of I8(t))
      !N8.call(e, r) &&
        r !== n &&
        objectDefineProperty(e, r, {
          get: () => t[r],
          enumerable: !(i = A8(t, r)) || i.enumerable,
        });
  return e;
};
var Bi = (e, t, n) => (
  (n = e != null ? T8(R8(e)) : {}),
  O8(
    t || !e || !e.__esModule
      ? objectDefineProperty(n, "default", { value: e, enumerable: true })
      : n,
    e,
  )
);

var Mo = () => __lib_saxes;

var u8 = () => __lib_exceljs;
var d8 = () => __lib_papaparse;
var AN = Bi(Mo(), 1);

function _p(e, t = "1900") {
  return (
    dateToSerial(e.getUTCFullYear(), e.getUTCMonth() + 1, e.getUTCDate(), t) +
    (e.getUTCHours() * 36e5 +
      e.getUTCMinutes() * 6e4 +
      e.getUTCSeconds() * 1e3 +
      e.getUTCMilliseconds()) /
      millisecondsPerDay
  );
}

var BP = { maxStyles: 50, maxCells: 500, maxName: 80 };

function VP(e) {
  (bo(e) || Fr("Workbook styles must be an object."),
    assertKnownStyleKeys(
      e,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    e.missingTheme !== undefined &&
      e.missingTheme !== true &&
      Fr("Invalid missing theme flag."),
    (!Array.isArray(e.styles) || e.styles.length > BP.maxStyles) &&
      Fr("A workbook supports at most 50 named styles."));
  let t = new Set(),
    n = new Set(),
    i = e.styles.map((a) => {
      (bo(a) || Fr("Invalid named style."),
        assertKnownStyleKeys(a, ["id", "name", "definition"], "named style"));
      let s = validateBoundedStyleText(a.id, "Style ID", 100),
        l = validateBoundedStyleText(a.name, "Style name");
      return (
        (t.has(s) || n.has(l.toLocaleLowerCase("en-US"))) &&
          Fr("Named style IDs and names must be unique within this workbook."),
        t.add(s),
        n.add(l.toLocaleLowerCase("en-US")),
        {
          id: s,
          name: l,
          definition: validateNamedStyleDefinition(a.definition),
        }
      );
    }),
    r = e.theme === undefined || e.missingTheme === true,
    o =
      e.theme === undefined
        ? cloneStyleValue(officeThemeDefinition)
        : validateWorkbookTheme(e.theme);
  return (
    r &&
      canonicalStyleKey(o) !== canonicalStyleKey(officeThemeDefinition) &&
      Fr("A missing theme must use the Office fallback."),
    {
      namespace: validateBoundedStyleText(e.namespace, "Style namespace", 100),
      theme: o,
      styles: i,
      ...(r ? { missingTheme: true } : {}),
    }
  );
}

function z6(e, t) {
  let n = e === undefined ? undefined : VP(e);
  n ? (t.styleCatalog = n) : delete t.styleCatalog;
  let i = bo(e) && e.theme === undefined;
  for (let r of t.sheets)
    for (let o of hasLinkedOrBoundCells(r.cells)
      ? Object.values(r.cells)
      : []) {
      if (!o.style?.binding) continue;
      let a = validateNamedStyleBinding(o.style.binding),
        s = resolveNamedStyleBinding(n, a);
      (!s.missing &&
        !i &&
        canonicalStyleKey(a) !== canonicalStyleKey(s) &&
        Fr(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (o.style.binding = s));
    }
}

function X6(e, t) {
  let n = t === "row" ? 21 : 40,
    i = t === "row" ? 546 : 600;
  if (typeof e != "number" || !Number.isFinite(e) || e < n || e > i)
    throw Error(
      (t === "row" ? "Default row height" : "Default column width") +
        " must be " +
        n +
        "\u2013" +
        i +
        " pixels.",
    );
  return e;
}
function H6(e, t) {
  let n =
      e.defaultRowHeight === undefined
        ? undefined
        : X6(e.defaultRowHeight, "row"),
    i =
      e.defaultColumnWidth === undefined
        ? undefined
        : X6(e.defaultColumnWidth, "column");
  (n !== undefined && (t.defaultRowHeight = n),
    i !== undefined && (t.defaultColumnWidth = i));
}

function Dw(e, t, n = { rows: t.rowCount, cols: t.colCount }) {
  H6(e, t);
  let i = validateSplitPaneState(e.splitPanes, n);
  if (i && (e.frozenRows || e.frozenCols))
    throw Error("Split panes and frozen panes cannot be enabled together.");
  i && (t.splitPanes = i);
  let r = normalizeHiddenIntervals(e.hiddenRows, n.rows),
    o = normalizeHiddenIntervals(e.hiddenCols, n.cols),
    a = { ...t, hiddenRows: r, hiddenCols: o },
    s = createAxisVisibilityIndex(t.rowCount, cn(a, "row")),
    l = createAxisVisibilityIndex(t.colCount, cn(a, "column"));
  if (!s.count || !l.count)
    throw Error("At least one row and one column must remain visible.");
  let c = e.frozenRows === undefined ? 0 : e.frozenRows,
    u = e.frozenCols === undefined ? 0 : e.frozenCols;
  if (
    !Number.isInteger(c) ||
    !Number.isInteger(u) ||
    c < 0 ||
    u < 0 ||
    c > n.rows ||
    u > n.cols
  )
    throw Error(
      "Frozen boundaries must be valid leading row and column counts.",
    );
  let d = s.countBefore(c),
    f = l.countBefore(u);
  if (
    d > worksheetCapacityLimits.maxFrozenVisibleRows ||
    f > worksheetCapacityLimits.maxFrozenVisibleCols
  )
    throw Error(
      "Freeze panes supports at most 100 visible rows and 50 visible columns.",
    );
  if (d === s.count || f === l.count)
    throw Error("Freeze panes must leave an unfrozen visible row and column.");
  for (let p of t.merges ?? []) {
    let h = parseRangeAddress(p);
    if (
      (h.firstRow < c && h.lastRow >= c) ||
      (h.firstCol < u && h.lastCol >= u)
    )
      throw Error(
        "Freeze panes cannot cross a merged range. Unmerge those cells first.",
      );
  }
  (e.hiddenRows !== undefined && (t.hiddenRows = r),
    e.hiddenCols !== undefined && (t.hiddenCols = o),
    e.frozenRows !== undefined && (t.frozenRows = c),
    e.frozenCols !== undefined && (t.frozenCols = u));
}

function Q6(e = "Book1") {
  return gi({
    schemaVersion: 1,
    id: crypto.randomUUID(),
    name: e.trim(),
    revision: 0,
    metadata: createWorkbookMetadata(),
    sheets: [
      {
        id: crypto.randomUUID(),
        name: "Sheet1",
        cells: {},
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        frozenRows: 0,
      },
    ],
    events: [],
  });
}

function Bw(e) {
  if (!e.sheets.some((n) => n.validations?.length)) return [];
  let t = summarizeValidationInspection(e);
  return t.invalidCells || !t.complete
    ? [
        `Data validation: ${t.message} Source values were retained; use Data Validation to review.`,
      ]
    : [];
}

function s5(e, t, n) {
  if (new Set(n.map((r) => r.toLowerCase())).size !== n.length)
    return throwXlsxDefinedNameError(
      "Duplicate worksheet names make name scopes ambiguous.",
    );
  let i = e.map((r) => {
    let o = r.localSheetId === undefined ? undefined : n[r.localSheetId],
      a = o === undefined ? undefined : t.sheets.find((s) => s.name === o)?.id;
    return r.localSheetId !== undefined && !a
      ? throwXlsxDefinedNameError(
          `Name "${r.name}" has an invalid or unsupported worksheet scope.`,
        )
      : {
          name: r.name,
          refersTo: r.refersTo,
          ...(a ? { scopeSheetId: a } : {}),
          ...(r.comment !== undefined ? { comment: r.comment } : {}),
        };
  });
  for (let r of i) validateNameSheetReferences(r.refersTo, t);
  try {
    return validateDefinedName(i, t);
  } catch (r) {
    return throwXlsxDefinedNameError(r.message);
  }
}

function Wm(e, t) {
  if (
    !isBoundedString(t.id, 100) ||
    !isBoundedString(t.name, 255) ||
    !["file", "snapshot"].includes(t.kind)
  )
    throw Error("Provide a valid source identity, name and origin kind.");
  let n = [...(e.rowOrigins ?? [])],
    i = e.sheets.map((r) => {
      let o = { ...r.rowOriginIds },
        a = new Set(
          Object.keys(r.cells).map((s) => Number(/\d+$/.exec(s)[0]) - 1),
        );
      for (let s of a)
        if (!o[s]) {
          if (n.length >= workbookLimits.maxCells)
            throw Error(
              "Row provenance has reached its 100,000 source-row budget.",
            );
          let l = crypto.randomUUID();
          ((o[s] = l),
            n.push({
              id: l,
              sourceId: t.id,
              sourceName: t.name.slice(0, 255),
              sourceSheet: r.name,
              originalRow: s + 1,
              kind: t.kind,
            }));
        }
      return { ...r, rowOriginIds: o };
    });
  return { ...e, sheets: i, rowOrigins: n };
}

function j5(e, t) {
  let n = [],
    i = [],
    r = e.model.tables ?? [];
  for (let o of r) {
    if (o.headerRow === false) {
      n.push(
        `Table ${o.name} has no header row; its cells were retained without table metadata.`,
      );
      continue;
    }
    try {
      let a = Vm(o.tableRef);
      resolveSourceRange(a, t);
      let s =
          Object.entries(tablePaletteStyleNames).find(
            ([, c]) => c === o.style?.theme,
          )?.[0] ?? "blue",
        l = {};
      for (let c of o.columns) {
        let u = c.totalsRowFunction?.toLowerCase();
        if (basicAggregateOperations.includes(u))
          Object.defineProperty(l, c.name, {
            value: u,
            enumerable: true,
            writable: true,
            configurable: true,
          });
        else if (u === "custom") {
          let d = c.totalsRowFormula ?? "",
            f = /^=?\s*(SUM|COUNTA|AVERAGE|MIN|MAX)\(/i
              .exec(d)?.[1]
              .toLowerCase();
          f &&
            Object.defineProperty(l, c.name, {
              value: f === "counta" ? "count" : f,
              enumerable: true,
              writable: true,
              configurable: true,
            });
        }
      }
      (i.push({
        id: createUuid(),
        name: validateTableName(o.name),
        range: a,
        columns: o.columns.map((c) => ({
          id: createUuid(),
          name: validateColumnName(c.name),
        })),
        style: s,
        showTotals: !!o.totalsRow,
        totals: l,
      }),
        Object.values(tablePaletteStyleNames).includes(o.style?.theme) ||
          n.push(
            `Table ${o.name}: source cells were preserved; unsupported table theme maps to the blue table style for future edits.`,
          ));
    } catch (a) {
      n.push(
        `Table ${o.name}: ${a.message} Cells remain available without this table definition.`,
      );
    }
  }
  return (i.length && ((t.tables = i), normalizeTableDefinitions(t, t)), n);
}

var Ud = class extends Error {
  constructor(n) {
    super(externalSourceErrorMessages[n]);
    this.code = n;
    this.name = "ExternalSourceError";
  }
};

function _R(e, t) {
  let n = 0,
    i = 0;
  for (let r of t.sheets) {
    let o = e.sheets.find((l) => l.id === r.id)?.externalConnection;
    if (o === undefined) continue;
    if (
      ++i > externalSourceLimits.sources ||
      (n += JSON.stringify(o).length) > externalSourceLimits.metadataCharacters
    )
      throw new Ud("limit");
    let a = sourceRefreshStateSchema.safeParse(o);
    if (!a.success) throw new Ud("invalid");
    let s = a.data;
    if (
      s.lastSuccess.controls.length !== s.source.controlFields.length ||
      s.lastSuccess.controls.some(
        (l, c) =>
          l.field !== s.source.controlFields[c] ||
          l.numeric + l.blanks !== s.lastSuccess.rows,
      ) ||
      (s.lastSuccess.rows + 1) * s.source.columns.length >
        externalSourceLimits.cells ||
      Date.parse(s.lastAttempt.at) < Date.parse(s.lastSuccess.at) ||
      (s.lastAttempt.outcome === "succeeded") == !!s.lastAttempt.code
    )
      throw new Ud("invalid");
    r.externalConnection = s;
  }
}

function XW(e) {
  if (new TextEncoder().encode(JSON.stringify(e)).length > 4e6)
    throw Error("A report snapshot exceeds 4 MB.");
  let t = calculationReportSchema.parse(e),
    n = parseRangeAddress(t.request.range);
  if (
    (validateObjectReportRequest(t.request),
    t.hash !== hashRecordExcludingHash(t))
  )
    throw Error("Report snapshot integrity check failed.");
  let i = Array.from(
    { length: n.lastCol - n.firstCol + 1 },
    (a, s) => n.firstCol + s,
  );
  if (
    fingerprintCanonicalValue(i) !== fingerprintCanonicalValue(t.columns) ||
    t.widths.length !== i.length
  )
    throw Error("Invalid report column layout.");
  let r = t.rows.map((a) => a.row),
    o = [...r, ...t.excludedRows].sort((a, s) => a - s);
  if (
    t.rows.length * i.length > 2e3 ||
    o.length !== n.lastRow - n.firstRow + 1 ||
    o.some((a, s) => a !== n.firstRow + s + 1) ||
    r.some((a, s) => s > 0 && a <= r[s - 1])
  )
    throw Error("Invalid report row accounting.");
  for (let a of t.rows) {
    if (a.cells.length !== i.length) throw Error("Invalid report cell layout.");
    a.cells.forEach((s, l) => {
      if (s.address !== `${columnIndexToLetters(i[l])}${a.row}`)
        throw Error("Invalid report cell address.");
      s.style &&= validateCellStyle(s.style);
    });
  }
  if (t.object) {
    let a = {
      id: t.request.sheetId,
      name: t.sheetName,
      cells: {},
      rowCount: 1048576,
      colCount: 16384,
    };
    (validateDrawingObjects([t.object], a), (t.object = a.objects[0]));
  }
  if (
    (t.request.kind === "object") != !!t.object ||
    (t.object && t.object.id !== t.request.objectId)
  )
    throw Error("Invalid report drawing identity.");
  if (t.hash !== hashRecordExcludingHash(t))
    throw Error("Report metadata is not canonical.");
  return t;
}
function OR(e, t) {
  if (e.reportOutputs !== undefined) {
    let n = X.array(
      X.object({
        id: nonEmptyLabelSchema,
        version: X.number().int().min(1).max(1e6),
        snapshot: X.unknown(),
      }).strict(),
    )
      .max(20)
      .parse(e.reportOutputs);
    if (new Set(n.map((i) => i.id)).size !== n.length)
      throw Error("Duplicate report output IDs.");
    if (
      new TextEncoder().encode(JSON.stringify(n)).length >
      reportOutputLimits.metadataBytes
    )
      throw Error("Saved reports exceed 16 MB.");
    t.reportOutputs = n.map((i) => ({ ...i, snapshot: XW(i.snapshot) }));
  }
  if (e.embeddedObjectInventories !== undefined) {
    let n = X.array(importedCacheEvidenceSchema)
      .max(20)
      .parse(e.embeddedObjectInventories);
    if (
      new Set(n.map((i) => i.id)).size !== n.length ||
      JSON.stringify(n).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    t.embeddedObjectInventories = n;
  }
}

function lN(e, t) {
  let n = e.pageSetup,
    i = createDefaultPageSetup(t);
  if (
    n.pageOrder !== undefined &&
    n.pageOrder !== "overThenDown" &&
    n.pageOrder !== "downThenOver"
  )
    throw Error(
      "The Excel worksheet has an unsupported pageOrder value. Import stopped without changing the current workbook.",
    );
  let r = {
    ...i,
    orientation: n.orientation === "landscape" ? "landscape" : "portrait",
    paper: n.paperSize === 9 ? "a4" : n.paperSize === 5 ? "legal" : "letter",
    margins: { ...i.margins, ...n.margins },
    printArea: n.printArea?.split("&&")[0] || i.printArea,
    fitToPage: !!n.fitToPage,
    fitToWidth: Math.max(0, Math.min(10, n.fitToWidth ?? 1)),
    fitToHeight: Math.max(0, Math.min(10, n.fitToHeight ?? 0)),
    scale: Math.max(10, Math.min(400, n.scale ?? 100)),
    gridlines: !!n.showGridLines,
    headings: !!n.showRowColHeaders,
    header: String(e.headerFooter?.oddHeader ?? "").slice(0, 255),
    footer: String(e.headerFooter?.oddFooter ?? "").slice(0, 255),
    pageOrder: n.pageOrder ?? "downThenOver",
    ...(n.printTitlesRow ? { repeatRows: n.printTitlesRow } : {}),
    ...(n.printTitlesColumn ? { repeatColumns: n.printTitlesColumn } : {}),
  };
  validatePageSetup(r, t);
}

var fN = {
  sourceCells: 500,
  sourceCharacters: 2e5,
  fields: 50,
  steps: 12,
  outputCells: 5e3,
  outputCharacters: 1e6,
  inferenceColumns: 8,
  inferencePatterns: 8e3,
  inferenceCandidates: 30,
  records: 100,
  evidenceCharacters: 5e6,
};

function gN(e, t) {
  if (e === undefined) return;
  if (
    !Array.isArray(e) ||
    e.length > 100 ||
    JSON.stringify(e).length > fN.evidenceCharacters
  )
    throw Error(
      "Field transformation evidence exceeds its native history budget. Keep the original backup.",
    );
  let n = new Set(),
    i = 0;
  t.fieldTransforms = e.map((r) => {
    let o = analysisOutputEvidenceSchema.safeParse(r);
    if (!o.success)
      throw Error(
        `Invalid field transformation evidence: ${o.error.issues[0]?.message}`,
      );
    let a = o.data,
      s = validateFieldTransformRecipe(a.request),
      l = parseRangeAddress(s.sourceRange);
    if (
      new Date(a.createdAt).toISOString() !== a.createdAt ||
      a.sources.length !==
        (l.lastRow - l.firstRow + 1) * (l.lastCol - l.firstCol + 1)
    )
      throw Error(
        "Invalid field evidence timestamp or incomplete original source population.",
      );
    if (n.has(a.id) || (i += a.outputs.length) > 2e4)
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (n.add(a.id),
      a.sourceSheetId !== s.sourceSheetId || a.outputSheetName !== s.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    let c = new Set(),
      u = a.sources.map((f) => {
        let p = _t(f.address);
        if (
          p.row < l.firstRow ||
          p.row > l.lastRow ||
          p.col < l.firstCol ||
          p.col > l.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (c.has(f.address)) throw Error("Duplicate original field source.");
        c.add(f.address);
        let h =
          f.original === undefined
            ? undefined
            : normalizeCellRecordInput(f.original);
        if (
          h &&
          typeof h.raw == "string" &&
          h.raw.startsWith("=") &&
          !h.literal
        )
          throw Error("Field source evidence must contain reviewed values.");
        return { ...f, ...(h ? { original: h } : {}) };
      }),
      d = new Set();
    for (let f of a.outputs) {
      if (d.has(f.address))
        throw Error("Duplicate transformed output address.");
      if (
        (d.add(f.address),
        ((f.type === "number" || f.type === "date") &&
          typeof f.value != "number") ||
          (f.type === "boolean" && typeof f.value != "boolean") ||
          (f.type === "text" && typeof f.value != "string") ||
          (f.type === "blank" && f.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (let f of [...a.outputs, ...a.exceptions])
      if (
        f.row < l.firstRow + 1 + Number(s.header) ||
        f.row > l.lastRow + 1 ||
        f.sources.some((p) => !c.has(p)) ||
        new Set(f.sources).size !== f.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...a, request: s, sources: u };
  });
}

function gi(e) {
  return LV(e, false);
}
function LV(e, t) {
  if (t && isDeepFrozenValue(e) && $V.has(e)) return e;
  if (!Gs(e) || e.schemaVersion !== 1)
    return jt("Unsupported workbook schema. Expected schemaVersion 1.");
  let n = requireBoundedText(e.id, "Workbook ID", 100),
    i = requireBoundedText(e.name, "Workbook name", 200);
  if (!n || !i.trim()) return jt("Workbook ID and name are required.");
  let r = P2(e.revision, 0, Number.MAX_SAFE_INTEGER - 1, "Revision");
  if (!Array.isArray(e.sheets) || !e.sheets.length)
    return jt("A workbook must contain at least one sheet.");
  let o = 0,
    a = new Set(),
    s = new Set(),
    l = e.sheets.map((p) => {
      if (!Gs(p)) return jt("Invalid sheet.");
      let h = requireBoundedText(p.id, "Sheet ID", 100),
        m = requireBoundedText(p.name, "Sheet name", 31);
      if (
        !h ||
        !m.trim() ||
        /[\\/\[\]*?:]/.test(m) ||
        a.has(h) ||
        s.has(m.toLowerCase())
      )
        return jt(
          "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
        );
      (a.add(h), s.add(m.toLowerCase()));
      let g = P2(p.rowCount, 1, workbookLimits.maxRows, "Row count"),
        b = P2(p.colCount, 1, workbookLimits.maxCols, "Column count");
      if (!Gs(p.cells)) return jt("Sheet cells must be an address map.");
      let w = t && isDeepFrozenValue(p) ? bN.get(p) : undefined;
      if (w !== undefined) {
        if (((o += w), o > workbookLimits.maxCells))
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        let O = { ...p };
        return (bN.set(O, w), O);
      }
      let _ = {};
      for (let [O, E] of Object.entries(p.cells)) {
        let R = _t(O);
        if (
          O !== `${columnIndexToLetters(R.col)}${R.row + 1}` ||
          R.row >= g ||
          R.col >= b
        )
          return jt(`Invalid or out-of-bounds cell: ${O}`);
        if (++o > workbookLimits.maxCells)
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        _[O] = normalizeCellRecord(E, t);
      }
      let S = {
        id: h,
        name: m,
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        cells: _,
      };
      if (p.columnWidths !== undefined) {
        if (!Gs(p.columnWidths))
          return jt("Column widths must be an index map.");
        S.columnWidths = {};
        for (let [O, E] of Object.entries(p.columnWidths)) {
          if (
            !/^(0|[1-9]\d*)$/.test(O) ||
            Number(O) >= b ||
            typeof E != "number" ||
            !Number.isFinite(E) ||
            E < 40 ||
            E > 600
          )
            return jt(
              "Column widths must be 40\u2013600 pixels for an existing column.",
            );
          S.columnWidths[O] = E;
        }
      }
      if (p.rowHeights !== undefined) {
        if (
          !Gs(p.rowHeights) ||
          (Object.getPrototypeOf(p.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(p.rowHeights) !== null)
        )
          return jt("Row heights must be a plain index map.");
        S.rowHeights = {};
        for (let [O, E] of Object.entries(p.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(O) ||
            Number(O) >= g ||
            typeof E != "number" ||
            !Number.isFinite(E) ||
            E < 21 ||
            E > 546
          )
            return jt(
              `Row heights must be ${21}\u2013${546} pixels for an existing row.`,
            );
          S.rowHeights[O] = E;
        }
      }
      if (p.hidden !== undefined) {
        if (typeof p.hidden != "boolean")
          return jt("Sheet hidden flag must be boolean.");
        S.hidden = p.hidden;
      }
      return (
        validateWorksheetFeatures(p, S),
        validateOutlineGroups(p, S, { rows: g, cols: b }),
        Dw(p, S, { rows: g, cols: b }),
        validateWorksheetVisibility(p, S),
        validateDelimitedSourceFootprint(p.delimitedSource, S),
        normalizeTableDefinitions(p, S),
        p.objects !== undefined && validateDrawingObjects(p.objects, S),
        p.pageLayout !== undefined && validatePageSetup(p.pageLayout, S),
        S
      );
    });
  if (l.every((p) => p.hidden))
    return jt("At least one sheet must be visible.");
  if (!Array.isArray(e.events) || e.events.length > 2e3)
    return jt("Workbook history must be an array of at most 2,000 events.");
  let c = new Set(),
    u = 0,
    d = e.events.map((p) => {
      if (!Gs(p)) return jt("Invalid history event.");
      let h = requireBoundedText(p.id, "Event ID", 100),
        m = requireBoundedText(p.timestamp, "Event timestamp", 40);
      return !h || c.has(h) || !Number.isFinite(Date.parse(m))
        ? jt("History events require unique IDs and valid timestamps.")
        : (c.add(h),
          !Array.isArray(p.changes) ||
          p.changes.length > workbookLimits.maxPatches ||
          (u += p.changes.length) > workbookLimits.maxCells
            ? jt("Invalid history change count.")
            : {
                id: h,
                timestamp: m,
                actor: requireBoundedText(p.actor, "Actor", 100),
                action: requireBoundedText(p.action, "Action", 200),
                changes: p.changes.map((g) => {
                  if (
                    !Gs(g) ||
                    typeof g.sheetId != "string" ||
                    !g.sheetId.length ||
                    g.sheetId.length > 100 ||
                    typeof g.address != "string"
                  )
                    return jt("Invalid history cell reference.");
                  let b = _t(g.address);
                  return g.address !==
                    `${columnIndexToLetters(b.col)}${b.row + 1}`
                    ? jt("History cell address is not canonical.")
                    : g.before === undefined && g.after === undefined
                      ? jt("History change must have a before or after value.")
                      : {
                          sheetId: g.sheetId,
                          address: g.address,
                          ...(g.before !== undefined
                            ? { before: normalizeCellRecord(g.before, t) }
                            : {}),
                          ...(g.after !== undefined
                            ? { after: normalizeCellRecord(g.after, t) }
                            : {}),
                        };
                }),
              });
    }),
    f = { schemaVersion: 1, id: n, name: i, revision: r, sheets: l, events: d };
  return (
    importFormulaCacheEvidence(e, f),
    e.datePolicy !== undefined &&
      (f.datePolicy = normalizeDatePolicy(e.datePolicy)),
    validateSavedFilterViews(e.savedFilterViews, f),
    validateSavedRecipeMetadata(e, f),
    verifyConsolidationEvidence(e, f),
    _R(e, f),
    validateDataModelReport(e, f),
    validateAnalysisReport(e, f),
    OR(e, f),
    normalizeRowProvenance(e, f),
    validateAutomationArtifacts(e, f),
    readRetainedXlsxSource(e, f),
    normalizeAlternateFormatSource(e, f),
    validateSourceImports(e, f),
    attachWorkbookLinkMetadata(e, f),
    normalizeTextCleanups(e.textCleanups, f),
    gN(e.fieldTransforms, f),
    validateTableSourceEvidence(e.tableEvidence, f),
    e.whatIfAnalyses !== undefined &&
      (f.whatIfAnalyses = validateSavedWhatIfAnalyses(e.whatIfAnalyses)),
    e.definedNames !== undefined &&
      (f.definedNames = validateDefinedName(e.definedNames, f)),
    validateExtensionState(e, f),
    validateWorkbookStructureProtection(e.structureProtection, f),
    validateWorkbookMetadata(e.metadata, f),
    assertUniqueTableNames(f),
    normalizeReviewMetadata(e, f),
    z6(e.styleCatalog, f),
    e.formulaWatches !== undefined &&
      (f.formulaWatches = validateFormulaWatches(e.formulaWatches, f)),
    dR(e, f),
    copyExtensionCacheEntries(e, f),
    f
  );
}

var WV = (e) => {
  let t = "";
  for (let n of e) t += String.fromCharCode(n);
  return btoa(t);
};
function SN(e, t, n, i = {}) {
  let r = [],
    o = [],
    a = 0,
    s = 0;
  for (let l of t.getImages()) {
    if (++s > 1e3 || r.length >= _N) {
      o.push(
        "Additional pictures beyond the 100-object / 1,000-source inspection limit were omitted.",
      );
      break;
    }
    try {
      let c = e.getImage(Number(l.imageId)),
        u =
          c.extension === "jpeg"
            ? "jpeg"
            : c.extension === "png"
              ? "png"
              : null;
      if (!u) throw new Error("Only PNG/JPEG images are supported.");
      if (c.buffer && c.buffer.byteLength > maxPictureHashBytes)
        throw Error("Embedded image exceeds 2 MB.");
      let d = c.base64?.startsWith("data:")
          ? c.base64.slice(c.base64.indexOf(",") + 1)
          : c.base64,
        f =
          d !== undefined
            ? Math.floor((d.length * 3) / 4) -
              (d.endsWith("==") ? 2 : d.endsWith("=") ? 1 : 0)
            : (c.buffer?.byteLength ?? 0);
      if (f > maxPictureHashBytes) throw Error("Embedded image exceeds 2 MB.");
      if (a + f > 10 * 1024 * 1024)
        throw Error(
          "Additional picture omitted: pictures in one worksheet may use at most 10 MB.",
        );
      let p = c.base64?.startsWith("data:")
          ? c.base64
          : c.base64
            ? `data:image/${u};base64,${c.base64}`
            : `data:image/${u};base64,${WV(new Uint8Array(c.buffer))}`,
        h = inspectPictureDataUrl(p),
        m = Math.floor(l.range.tl.col),
        g = Math.floor(l.range.tl.row),
        b = createAnchoredDrawingObject(
          "picture",
          `${columnIndexToLetters(m)}${g + 1}`,
        );
      ((b.offsetX = Math.max(
        0,
        (l.range.tl.col - m) * (n.columnWidths?.[m] ?? 84),
      )),
        (b.offsetY = Math.max(
          0,
          (l.range.tl.row - g) * (n.rowHeights?.[g] ?? 21),
        )),
        (b.width = Math.max(24, Math.min(4e3, l.range.ext?.width ?? 240))),
        (b.height = Math.max(24, Math.min(4e3, l.range.ext?.height ?? 140))),
        (b.dataURL = p),
        (b.evidence = createPictureEvidence(p, {
          origin: "xlsx-media",
          ...(i.containerFilename !== undefined
            ? { containerFilename: i.containerFilename }
            : {}),
          mediaReference: `imageId:${l.imageId}`,
        })),
        (b.label = `Picture ${r.length + 1}`),
        r.push(b),
        (a += h.bytes.length));
    } catch (c) {
      o.push(
        c instanceof Error ? c.message : "An image could not be imported.",
      );
    }
  }
  return (validateDrawingObjects(r, n), o);
}

function CN(e, t, n) {
  if (
    (validateExtractedSourceTables(e),
    e.fileName !== n ||
      e.byteLength !== t.length ||
      e.sha256 !== sha256HexDigest(t))
  )
    throw Error("Source inspection does not match its original.");
  VV.set(e, { signature: JSON.stringify(e), bytes: t.slice() });
}
var TN = (e) => e.replaceAll("~", "~0").replaceAll("/", "~1");
function qV(e) {
  let t = 0,
    n = 0,
    i = () => {
      for (; /[\t\n\r ]/.test(e[t] ?? "x");) t++;
    },
    r = () => {
      let s = t++;
      for (; t < e.length;) {
        let l = e[t++];
        if (l === '"') return JSON.parse(e.slice(s, t));
        l === "\\" && t++;
      }
      throw Error("Invalid JSON string.");
    };
  function o(s) {
    if (s > sourceImportLimits.depth || ++n > sourceImportLimits.nodes)
      throw Error("JSON nesting or node budget exceeded.");
    i();
    let l = e[t];
    if (l === '"') return r();
    if (l === "[") {
      (t++, i());
      let d = [];
      if (e[t] === "]") return (t++, d);
      for (;;) {
        if ((d.push(o(s + 1)), i(), e[t++] === "]")) return d;
        if (e[t - 1] !== ",") throw Error("Invalid JSON array.");
      }
    }
    if (l === "{") {
      (t++, i());
      let d = Object.create(null);
      if (e[t] === "}") return (t++, d);
      for (;;) {
        if ((i(), e[t] !== '"')) throw Error("Invalid JSON object.");
        let f = r();
        if (
          ["__proto__", "prototype", "constructor"].includes(f) ||
          Object.hasOwn(d, f)
        )
          throw Error("Unsafe or duplicate JSON field.");
        if ((i(), e[t++] !== ":")) throw Error("Invalid JSON field.");
        if (((d[f] = o(s + 1)), i(), e[t++] === "}")) return d;
        if (e[t - 1] !== ",") throw Error("Invalid JSON object.");
      }
    }
    for (let [d, f] of [
      ["true", true],
      ["false", false],
      ["null", null],
    ])
      if (e.startsWith(d, t)) return ((t += d.length), f);
    let c = /-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(e.slice(t));
    if (!c || c.index !== 0) throw Error("Invalid JSON value.");
    t += c[0].length;
    let u = Number(c[0]);
    if (
      !Number.isFinite(u) ||
      (Number.isInteger(u) && !Number.isSafeInteger(u))
    )
      throw Error("Unsafe JSON number; encode large identifiers as strings.");
    return u;
  }
  let a = o(0);
  if ((i(), t !== e.length)) throw Error("Trailing JSON content.");
  return a;
}
function XV(e, t) {
  return {
    value: typeof e == "object" && e !== null ? JSON.stringify(e) : e,
    type:
      typeof e == "object" && e !== null
        ? Array.isArray(e)
          ? "array"
          : "object"
        : e === null
          ? "null"
          : typeof e,
    path: t,
    confidence: 1,
  };
}
function HV(e) {
  let t = [],
    n = (i, r) => {
      if (Array.isArray(i)) i.length && t.push({ value: i, path: r });
      else if (yg(i))
        for (let [o, a] of Object.entries(i)) n(a, r + "/" + TN(o));
    };
  if (
    (n(e, ""),
    !t.length && yg(e) && t.push({ value: [e], path: "" }),
    !t.length)
  )
    throw Error("Choose a JSON object or a nonempty array.");
  if (t.length > sourceImportLimits.tables)
    throw Error("Source table count exceeded.");
  return t.map((i, r) => {
    let o = yg(e) && i.value[0] === e,
      a = new Set(),
      s = i.value.map((c, u) => {
        let d = new Map(),
          f = o ? i.path : i.path + "/" + u,
          p = (h, m) => {
            if (yg(h) && Object.keys(h).length)
              for (let [g, b] of Object.entries(h)) p(b, m + "/" + TN(g));
            else (a.add(m), d.set(m, XV(h, f + m)));
          };
        return (p(c, ""), d);
      }),
      l = [...a];
    return {
      id: "json-" + r,
      name: i.path || "Root",
      path: i.path,
      ambiguous: t.length > 1,
      columns: l.map((c) => ({ name: c || "Value", path: c })),
      rows: s.map((c, u) =>
        l.map(
          (d) =>
            c.get(d) ?? {
              value: null,
              type: "null",
              path: i.path + (o ? "" : "/" + u) + d,
              confidence: 1,
              missing: true,
            },
        ),
      ),
    };
  });
}
function GV(e) {
  if (/<!DOCTYPE|<!ENTITY/i.test(e))
    throw Error("XML document types and entities are not supported.");
  if (/<\?xml[^>]*encoding\s*=\s*['"](?!utf-8['"])/i.test(e))
    throw Error("Use UTF-8 XML.");
  let t = [],
    n = [],
    i = 0,
    r = new AN.SaxesParser({ xmlns: true }),
    o = (u, d) => (u ? `{${u}}${d}` : d);
  r.on("opentag", (u) => {
    if (++i > sourceImportLimits.nodes || t.length >= sourceImportLimits.depth)
      throw Error("XML nesting or node budget exceeded.");
    let d = o(u.uri, u.local),
      f = t.at(-1),
      p = f?.children ?? n,
      h = {
        namespaces: { ...(f?.namespaces ?? {}), ...u.ns },
        name: d,
        path:
          (f?.path ?? "") +
          "/" +
          d +
          "[" +
          (p.filter((m) => m.name === d).length + 1) +
          "]",
        attributes: [],
        text: "",
        children: [],
      };
    for (let m of Object.values(u.attributes))
      if (m.uri !== "http://www.w3.org/2000/xmlns/")
        if (
          m.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          m.local === "type"
        ) {
          let [g, b] = m.value.includes(":")
            ? m.value.split(":")
            : ["", m.value];
          if (g && !h.namespaces[g])
            throw Error("XML type prefix is undeclared.");
          h.type = o(h.namespaces[g] ?? "", b);
        } else
          m.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          m.local === "nil"
            ? (h.nil = m.value)
            : h.attributes.push({
                name: o(m.uri, m.local),
                value: m.value,
                uri: m.uri,
                local: m.local,
              });
    (p.push(h), t.push(h));
  });
  let a = (u) => {
    let d = t.at(-1);
    if (d && ((d.text += u), d.text.length > 32767))
      throw Error("XML cell text budget exceeded.");
  };
  if (
    (r.on("text", a),
    r.on("cdata", a),
    r.on("closetag", () => {
      t.pop();
    }),
    r.on("error", () => {
      throw Error("Malformed XML source.");
    }),
    r.write(e).close(),
    n.length !== 1)
  )
    throw Error("XML needs one root element.");
  let s = [],
    l = (u) => {
      let d = new Map();
      for (let f of u.children) d.set(f.name, [...(d.get(f.name) ?? []), f]);
      for (let [f, p] of d)
        p.length > 1 ? s.push({ path: u.path + "/" + f, nodes: p }) : l(p[0]);
    };
  if (
    (l(n[0]),
    s.length || s.push({ path: n[0].path, nodes: [n[0]] }),
    s.length > sourceImportLimits.tables)
  )
    throw Error("XML table count exceeded.");
  let c = (u) => {
    let d = u.text,
      f = "string",
      p = u.text,
      h = u.type?.replace("{http://www.w3.org/2001/XMLSchema}", ""),
      m = u.text.trim();
    if (u.nil !== undefined) {
      if (!["true", "1", "false", "0"].includes(u.nil))
        throw Error("Invalid XML nil value.");
      if (["true", "1"].includes(u.nil)) {
        if (m || u.children.length)
          throw Error("A nil XML field cannot contain a value.");
        ((d = null), (f = "null"));
      }
    }
    if (
      f !== "null" &&
      u.type?.startsWith("{http://www.w3.org/2001/XMLSchema}")
    ) {
      if (h === "boolean") {
        if (!["true", "1", "false", "0"].includes(m))
          throw Error("Invalid typed XML boolean.");
        ((d = m === "true" || m === "1"), (f = "boolean"));
      }
      if (
        [
          "decimal",
          "double",
          "float",
          "integer",
          "int",
          "long",
          "short",
          "byte",
          "nonNegativeInteger",
          "positiveInteger",
        ].includes(h)
      ) {
        let g = !["decimal", "double", "float"].includes(h),
          b = g
            ? /^[+-]?\d+$/
            : h === "decimal"
              ? /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/
              : /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
        if (
          ((d = Number(m)),
          !b.test(m) ||
            !Number.isFinite(d) ||
            (Number.isInteger(d) && !Number.isSafeInteger(d)) ||
            (h === "nonNegativeInteger" && d < 0) ||
            (h === "positiveInteger" && d <= 0))
        )
          throw Error("Invalid or unsafe typed XML number.");
        let w = {
          int: [-2147483648, 2147483647],
          short: [-32768, 32767],
          byte: [-128, 127],
        };
        if (w[h] && (d < w[h][0] || d > w[h][1]))
          throw Error("Typed XML integer is out of range.");
        f = "number";
      }
    }
    return {
      value: d,
      type: f,
      path: u.path,
      lexical: p,
      ...(u.type ? { declaredType: u.type } : {}),
      confidence: 1,
    };
  };
  return s.map((u, d) => {
    let f = new Set(),
      p = u.nodes.map((m) => {
        let g = new Map(),
          b = (_, S) => {
            (f.add(_), g.set(_, S));
          },
          w = (_, S) => {
            for (let O of _.attributes)
              b(S + "/@" + O.name, {
                value: O.value,
                type: "string",
                path: _.path + "/@" + O.name,
                confidence: 1,
              });
            if (_.children.length) {
              if (_.text.trim())
                throw Error(
                  "Mixed-content XML tables need explicit extraction.",
                );
              let O = new Map();
              for (let E of _.children) {
                let R = (O.get(E.name) ?? 0) + 1;
                (O.set(E.name, R),
                  w(
                    E,
                    S +
                      "/" +
                      E.name +
                      (_.children.filter((I) => I.name === E.name).length > 1
                        ? "[" + R + "]"
                        : ""),
                  ));
              }
            } else b(S, c(_));
          };
        return (w(m, ""), g);
      }),
      h = [...f];
    return {
      id: "xml-" + d,
      name: u.path.slice(-200),
      path: u.path,
      ambiguous: s.length > 1,
      columns: h.map((m) => ({ name: (m || "Value").slice(-200), path: m })),
      rows: p.map((m, g) =>
        h.map(
          (b) =>
            m.get(b) ?? {
              value: null,
              type: "null",
              path: u.nodes[g].path + b,
              confidence: 1,
              missing: true,
            },
        ),
      ),
    };
  });
}
function gk(e, t) {
  if (e.length < 1 || e.length > sourceImportLimits.bytes)
    throw Error("Source byte budget exceeded.");
  let n = t.split(".").at(-1)?.toLowerCase(),
    i = n === "jpg" ? "jpeg" : n === "htm" ? "html" : n,
    r = [],
    o = "document-adapter",
    a = "Host extraction required",
    s = [];
  if (i === "json" || i === "xml") {
    let c = new TextDecoder("utf-8", { fatal: true })
      .decode(e)
      .replace(/^\uFEFF/, "");
    ((r = i === "json" ? HV(qV(c)) : GV(c)),
      (o = i === "json" ? "native-json" : "native-xml"),
      (a = "Ledgerglass structured source v1"),
      s.push(
        "String identifiers and untyped XML remain text. Nested JSON arrays/empty objects are retained as JSON values. Imported formula-like strings remain literal.",
      ));
  } else if (i === "pdf") {
    if (
      new TextDecoder()
        .decode(e.subarray(0, 8))
        .match(/^%PDF-1\.[0-7]|^%PDF-2\.0/) == null
    )
      throw Error("Source extension does not match PDF content.");
    s.push(
      "PDF bytes are retained without rendering or executing embedded content. A host extractor or matching extraction JSON is required.",
    );
  } else if (i === "png" || i === "jpeg") {
    let c = "";
    for (let u = 0; u < e.length; u += 8192)
      c += String.fromCharCode(...e.subarray(u, u + 8192));
    (inspectPictureDataUrl(`data:image/${i};base64,${btoa(c)}`),
      s.push(
        "Image extraction requires a host OCR adapter or matching extraction JSON.",
      ));
  } else if (i === "html") {
    let c = new TextDecoder("utf-8", { fatal: true }).decode(e);
    if (!/^\s*(?:<!doctype\s+html[^>]*>\s*)?<(?:html|table)\b/i.test(c))
      throw Error("Source extension does not match an HTML document.");
    s.push(
      "HTML is retained as inert bytes. No markup, scripts or resource links are rendered or fetched. A host table extractor or matching extraction JSON is required.",
    );
  } else throw Error("Choose JSON, XML, PDF, PNG, JPEG or HTML.");
  let l = validateExtractedSourceTables({
    schemaVersion: 1,
    fileName: t,
    format: i,
    byteLength: e.length,
    sha256: sha256HexDigest(e),
    tables: r,
    method: o,
    extractor: a,
    warnings: s,
  });
  return (CN(l, e, t), l);
}
var LN = Bi(Mo(), 1);
var NN = Bi(Mo(), 1);
function IN(e, t, n) {
  if (
    n > fileImportLimits.maxXlsxExpandedBytes ||
    (t > fileImportLimits.xlsxRatioThresholdBytes &&
      t / Math.max(e, 1) > fileImportLimits.maxXlsxExpansionRatio)
  )
    throw new Error(
      "The XLSX archive exceeds the safe expansion limit (50 MB or 200:1 ratio).",
    );
}
function RN(e, t, n = 0) {
  let i = (u, d) =>
      new RegExp(`(?:^|\\s)${d}\\s*=\\s*(["'])(.*?)\\1`).exec(u)?.[2],
    r = e.replace(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g, ""),
    o = 0,
    a = 0,
    s = -1,
    l = n,
    c = (u) => {
      throw new Error(
        `Worksheet ${t}: ${u} Import stopped; no cells were truncated.`,
      );
    };
  for (let u of r.matchAll(/<(\/?)(?:[\w.-]+:)?(dimension|row|c)\b([^>]*)>/g)) {
    let [, d, f, p] = u;
    if (d) {
      f === "row" && (o = 0);
      continue;
    }
    if (f === "dimension") {
      let h = i(p, "ref");
      h || c("The declared dimension has no valid cell range.");
      try {
        parseRangeAddress(h);
      } catch {
        c(
          `Declared dimension ${h} exceeds or violates the supported worksheet bounds.`,
        );
      }
    } else if (f === "row") {
      let h = i(p, "r"),
        m = Number(h);
      ((!h || !/^[1-9]\d*$/.test(h) || m > fileImportLimits.maxRows) &&
        c(
          `Row ${h ?? "(missing)"} must be between 1 and ${fileImportLimits.maxRows}.`,
        ),
        (o || m <= a) && c("Rows must have distinct, ascending row numbers."),
        (a = m),
        (o = /\/\s*$/.test(p) ? 0 : m),
        (s = -1));
    } else {
      let h = i(p, "r");
      (!h || !o) &&
        c("Each cell must have an explicit address inside its source row.");
      let m;
      try {
        m = _t(h);
      } catch {
        c(`Cell ${h} exceeds or violates the supported worksheet bounds.`);
      }
      ((m.row + 1 !== o || m.col <= s) &&
        c(`Cell ${h} has a duplicate, unordered or mismatched row address.`),
        (s = m.col),
        ++l > fileImportLimits.maxXlsxCellRecords &&
          c(
            `The XLSX exceeds ${fileImportLimits.maxXlsxCellRecords.toLocaleString("en-US")} source cell records, including explicit blank cells.`,
          ));
    }
  }
  return l;
}
var ZV =
    "Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  St = (e) => {
    throw Error(`${e} ${ZV}`);
  },
  KV = new Uint32Array(256).map((e, t) => {
    for (let n = 0; n < 8; n++) t = t & 1 ? 3988292384 ^ (t >>> 1) : t >>> 1;
    return t >>> 0;
  });
function Tc(e, t) {
  for (let n of t) e = KV[(e ^ n) & 255] ^ (e >>> 8);
  return e;
}
function vg(e, t = {}) {
  let n = yk(e);
  return (
    (!n.has("[Content_Types].xml") ||
      (!n.has("xl/workbook.xml") &&
        !(t.allowBinaryWorkbookInventory && n.has("xl/workbook.bin")))) &&
      St("This archive is not a supported Excel workbook."),
    n
  );
}
function yk(e) {
  let t = new DataView(e.buffer, e.byteOffset, e.byteLength);
  (e.length >= 8 &&
    t.getUint32(0, true) === 3759263696 &&
    t.getUint32(4, true) === 3776623009 &&
    St(
      "Encrypted Office or legacy binary workbooks are unsupported. Remove the open password in the source application.",
    ),
    (e.length < 22 || t.getUint32(0, true) !== 67324752) &&
      St("This file is not a supported XLSX archive."));
  let n = -1;
  for (let f = e.length - 22; f >= Math.max(0, e.length - 65557); f--)
    if (
      t.getUint32(f, true) === 101010256 &&
      f + 22 + t.getUint16(f + 20, true) === e.length
    ) {
      n = f;
      break;
    }
  n < 0 && St("The XLSX archive is incomplete or has trailing data.");
  let i = t.getUint16(n + 10, true),
    r = t.getUint32(n + 12, true),
    o = t.getUint32(n + 16, true);
  (t.getUint16(n + 4, true) ||
    t.getUint16(n + 6, true) ||
    i !== t.getUint16(n + 8, true) ||
    i === 65535 ||
    i > fileImportLimits.maxXlsxEntries ||
    o + r !== n) &&
    St("Multi-volume, ZIP64 or oversized XLSX archives are not supported.");
  let a = new Map(),
    s = new Set(),
    l = [],
    c = (f) => {
      try {
        return new TextDecoder("utf-8", { fatal: true }).decode(f);
      } catch {
        return St("Unsupported XLSX path encoding.");
      }
    },
    u = o,
    d = 0;
  for (let f = 0; f < i; f++) {
    (u + 46 > n || t.getUint32(u, true) !== 33639248) &&
      St("Invalid XLSX archive directory.");
    let p = t.getUint16(u + 8, true),
      h = t.getUint16(u + 10, true),
      m = t.getUint32(u + 16, true),
      g = t.getUint32(u + 20, true),
      b = t.getUint32(u + 24, true),
      w = t.getUint16(u + 28, true),
      _ = t.getUint16(u + 30, true),
      S = t.getUint16(u + 32, true),
      O = t.getUint32(u + 42, true);
    (p & 8257 && St("Encrypted XLSX archive entries are unsupported."),
      (p & -2063 ||
        ![0, 8].includes(h) ||
        [g, b, O].includes(4294967295) ||
        t.getUint16(u + 34, true)) &&
        St("Unsupported XLSX compression or ZIP64 metadata."),
      (u + 46 + w + _ + S > n ||
        O + 30 > o ||
        t.getUint32(O, true) !== 67324752) &&
        St("Invalid XLSX archive offsets."));
    let E = c(e.subarray(u + 46, u + 46 + w));
    ((!E ||
      s.has(E.toLowerCase()) ||
      /^[\/]/.test(E) ||
      /[\\\x00-\x1f\x7f%:]/.test(E) ||
      E.split("/").some(
        (A, L, y) => A === ".." || A === "." || (!A && L < y.length - 1),
      )) &&
      St("Unsafe or duplicate XLSX archive path."),
      s.add(E.toLowerCase()));
    let R = t.getUint16(O + 26, true),
      I = t.getUint16(O + 28, true),
      U = O + 30 + R + I,
      P = U + g;
    (P > o ||
      t.getUint16(O + 6, true) !== p ||
      t.getUint16(O + 8, true) !== h ||
      c(e.subarray(O + 30, O + 30 + R)) !== E) &&
      St("XLSX local headers do not match the archive directory.");
    let T = P;
    if (p & 8) {
      let A = P + (P + 4 <= o && t.getUint32(P, true) === 134695760 ? 4 : 0);
      ((A + 12 > o ||
        t.getUint32(A, true) !== m ||
        t.getUint32(A + 4, true) !== g ||
        t.getUint32(A + 8, true) !== b) &&
        St("Invalid XLSX streaming data descriptor."),
        (T = A + 12));
    } else
      (t.getUint32(O + 14, true) !== m ||
        t.getUint32(O + 18, true) !== g ||
        t.getUint32(O + 22, true) !== b) &&
        St("XLSX local sizes or checksum do not match the archive directory.");
    (h === 0 && g !== b && St("Invalid stored XLSX entry size."),
      l.push({ start: O, end: T }),
      (d += b),
      IN(g, b, d),
      a.set(E, { size: b, crc: m }),
      (u += 46 + w + _ + S));
  }
  return (
    l.sort((f, p) => f.start - p.start),
    (u !== n ||
      l.some((f, p) => f.start !== (p ? l[p - 1].end : 0)) ||
      l.at(-1)?.end !== o) &&
      St("Overlapping or unlisted XLSX archive entries."),
    a
  );
}
var JV =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  YV = "http://schemas.openxmlformats.org/package/2006/relationships",
  bk = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/",
  ON = () => ({
    warnings: new Set(),
    hasMainType: false,
    hasOfficeDocument: false,
    strings: 0,
    maxString: -1,
    styles: 0,
    maxStyle: -1,
    emptyStringCaches: new Map(),
  });
function Zs(e, t) {
  let n;
  try {
    n = decodeURIComponent(t);
  } catch {
    return St(`Invalid relationship target in ${e}.`);
  }
  /[\\\x00-\x1f\x7f:?#]/.test(n) && St(`Unsafe internal relationship in ${e}.`);
  let i =
      e === "_rels/.rels"
        ? []
        : e
            .replace(/_rels\/[^/]+\.rels$/, "")
            .split("/")
            .filter(Boolean),
    r = n.startsWith("/") ? [] : i;
  for (let o of n.split("/"))
    o === ".."
      ? (r.length || St(`Relationship escapes the XLSX package in ${e}.`),
        r.pop())
      : o && o !== "." && r.push(o);
  return r.join("/");
}
function MN(e, t, n, i) {
  let r = new NN.SaxesParser({ xmlns: true }),
    o = 0,
    a = 0,
    s = new Set(),
    l = [],
    c,
    u = /^xl\/worksheets\/[^/]+\.xml$/.test(t);
  (r.on("error", () => St(`Malformed XML in ${t}.`)),
    r.on("doctype", () =>
      St(`DTD/entity declarations are unsupported in ${t}.`),
    ),
    r.on("cdata", () =>
      St(
        `CDATA content is unsupported in ${t}. Resave as standard XLSX text encoding.`,
      ),
    ),
    r.on("opentag", (d) => {
      (++o > 128 || ++a > 1e6) &&
        St(`XML complexity exceeds the supported budget in ${t}.`);
      let f = (h) => d.attributes[h]?.value,
        p = l.at(-1);
      if (
        (l.push(d.local),
        t === "xl/sharedStrings.xml" && d.local === "si" && i.strings++,
        t === "xl/styles.xml" &&
          d.local === "xf" &&
          p === "cellXfs" &&
          i.styles++,
        u && d.local === "c" && p === "row")
      ) {
        let h = f("t") ?? "n";
        (["n", "s", "str", "inlineStr", "b", "e"].includes(h) ||
          St(`Unsupported cell type ${h} in ${t}.`),
          (c = {
            address: f("r") ?? "unknown cell",
            type: h,
            value: "",
            inValue: false,
            formula: false,
          }));
        let m = f("s");
        m !== undefined &&
          ((!/^\d+$/.test(m) || !Number.isSafeInteger(Number(m))) &&
            St(`Invalid cell style index in ${t}.`),
          (i.maxStyle = Math.max(i.maxStyle, Number(m))));
      }
      if (
        (c &&
          p === "c" &&
          d.local === "v" &&
          ((c.value = ""), (c.inValue = true)),
        c &&
          p === "c" &&
          d.local === "f" &&
          ((c.formula = true),
          f("t") &&
            !["normal", "shared", "array"].includes(f("t")) &&
            St(
              `Unsupported formula kind at ${c.address} in ${t}. Convert data-table formulas to reviewed ordinary formulas or values first.`,
            )),
        u &&
          ["sheetProtection", "autoFilter", "extLst", "hyperlinks"].includes(
            d.local,
          ) &&
          i.warnings.add(
            `Worksheet ${d.local} metadata was detected. Protection, filter criteria, interactive links and unsupported extensions are not preserved as editable Excel features.`,
          ),
        o === 1 &&
          ((t === "[Content_Types].xml" &&
            (d.local !== "Types" ||
              d.uri !==
                "http://schemas.openxmlformats.org/package/2006/content-types")) ||
            (t.endsWith(".rels") &&
              (d.local !== "Relationships" || d.uri !== YV)) ||
            (t === "xl/workbook.xml" &&
              (d.local !== "workbook" || d.uri !== QV))) &&
          St(
            `Unsupported XLSX package type or Strict XML namespace in ${t}. Save as standard Excel Workbook (.xlsx).`,
          ),
        t === "[Content_Types].xml" && d.local === "Override")
      ) {
        let h = f("PartName"),
          m = f("ContentType");
        ((!h || !m || s.has(h)) &&
          St("Invalid or duplicate XLSX content type declaration."),
          s.add(h),
          /model\+data|dataModel/i.test(m) &&
            i.warnings.add(
              "Imported Excel Data Model and DAX content is not executed or translated; model caches are historical and require independent review. Exact bytes remain only in the retained original.",
            ),
          /macroEnabled|vbaProject|activeX|oleObject/i.test(m) &&
            St(
              "Macro-enabled or executable workbook content is unsupported; renaming a file does not convert its format.",
            ),
          h === "/xl/workbook.xml" &&
            (m !== JV &&
              St(
                "Unsupported workbook content type. Templates and alternate formats must be saved as XLSX.",
              ),
            (i.hasMainType = true)),
          n.has(Zs("_rels/.rels", h)) || St(`Missing XLSX part ${h}.`));
      }
      if (t.endsWith(".rels") && d.local === "Relationship") {
        let h = f("Id"),
          m = f("Target"),
          g = f("Type"),
          b = f("TargetMode");
        if (
          ((!h ||
            !m ||
            !g ||
            s.has(h) ||
            (b && !["External", "Internal"].includes(b))) &&
            St(`Invalid or duplicate relationship in ${t}.`),
          s.add(h),
          /vbaProject|activeX|oleObject|attachedTemplate/i.test(g) &&
            St("Executable or embedded-object relationships are unsupported."),
          b === "External")
        )
          (g !== `${bk}hyperlink` &&
            g !== `${bk}externalLinkPath` &&
            St(`Unsupported external relationship in ${t}.`),
            i.warnings.add(
              "External hyperlinks/workbook links were found. No external target is fetched; interactive links and refresh connections are omitted.",
            ));
        else {
          let w = Zs(t, m);
          (n.has(w) || St(`Missing relationship target ${w} in ${t}.`),
            t === "_rels/.rels" &&
              g === `${bk}officeDocument` &&
              ((i.hasOfficeDocument || w !== "xl/workbook.xml") &&
                St("Unsupported or duplicate workbook relationship."),
              (i.hasOfficeDocument = true)));
        }
      }
    }),
    r.on("text", (d) => {
      c?.inValue && (c.value += d);
    }),
    r.on("closetag", (d) => {
      if (c && d.local === "v") {
        let f = c.value;
        if (
          (c.type === "n" &&
            f !== "" &&
            (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(f) ||
              !Number.isFinite(Number(f))) &&
            St(`Invalid numeric value at ${c.address} in ${t}.`),
          c.type === "b" &&
            !["0", "1"].includes(f) &&
            St(
              `Unsupported boolean value at ${c.address} in ${t}; save numeric Boolean encoding in a standard XLSX copy.`,
            ),
          c.type === "s" &&
            ((!/^\d+$/.test(f) || !Number.isSafeInteger(Number(f))) &&
              St(`Invalid shared-string index at ${c.address} in ${t}.`),
            (i.maxString = Math.max(i.maxString, Number(f)))),
          c.type === "e" &&
            !/^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA|CONNECT!|BLOCKED!|FIELD!|UNKNOWN!|BUSY!)$/.test(
              f,
            ) &&
            St(`Unsupported error value at ${c.address} in ${t}.`),
          c.type === "e" &&
            !c.formula &&
            i.warnings.add(
              "Source constant error cells are flattened to displayed text. Error propagation from those cells is not preserved; review dependent calculations.",
            ),
          c.type === "str" && c.formula && f === "")
        ) {
          let p = i.emptyStringCaches.get(t) ?? new Set();
          (p.add(c.address), i.emptyStringCaches.set(t, p));
        }
        c.inValue = false;
      }
      (d.local === "c" && (c = undefined), l.pop(), o--);
    }),
    r.write(e).close());
}
function $N(e, t) {
  ((!e.hasMainType || !e.hasOfficeDocument) &&
    St("Missing standard XLSX workbook content type or relationship."),
    e.maxString >= e.strings &&
      St("A worksheet references a missing shared string."),
    e.maxStyle >= e.styles &&
      St("A worksheet references a missing cell style."));
  let n = [
    [
      /^xl\/(?:model|dataModel)\//i,
      "Excel Data Model / Power Pivot content is retained only in the original source and is not executed or translated. Relationships, DAX measures and model caches are unavailable; worksheet results are historical snapshots. Rebuild an explicit local Data Model from reviewed worksheet tables before relying on recalculated results.",
    ],
    [/^xl\/charts\//i, "Charts are not imported as editable native charts."],
    [
      /^xl\/(?:pivotTables|pivotCache)\//i,
      "Pivot definitions and caches are omitted; supported worksheet result cells remain.",
    ],
    [
      /^xl\/(?:queryTables|externalLinks)\/|^xl\/connections\.xml$/i,
      "Queries, external workbook links and connections are omitted and never refreshed.",
    ],
    [
      /^customXml\/|^docProps\/custom\.xml$/i,
      "Custom XML and custom document properties are not editable; the part inventory identifies passive metadata preserved unchanged and content retained only in the original.",
    ],
    [
      /^xl\/(?:threadedComments|persons|slicers|slicerCaches|timelines)\//i,
      "Threaded discussions, people, slicers and timelines are omitted.",
    ],
  ];
  for (let i of t) for (let [r, o] of n) r.test(i) && e.warnings.add(o);
  return [...e.warnings];
}
function PN(e) {
  let t = () => {
      throw Error(
        "Unsupported, encrypted or corrupt legacy workbook. Open the original in its source application and save an unencrypted XLSX copy.",
      );
    },
    n = new DataView(e.buffer, e.byteOffset, e.byteLength),
    i = 4294967294,
    r = 4294967295;
  if (
    e.length < 512 ||
    n.getUint32(0, true) !== 3759263696 ||
    n.getUint32(4, true) !== 3776623009 ||
    n.getUint16(28, true) !== 65534
  )
    return t();
  let o = n.getUint16(26, true),
    a = n.getUint16(30, true),
    s = 2 ** a;
  if (
    !((o === 3 && a === 9) || (o === 4 && a === 12)) ||
    n.getUint16(32, true) !== 6 ||
    n.getUint32(56, true) !== 4096 ||
    e.length % s
  )
    return t();
  let l = e.length / s - 1,
    c = (A) =>
      !Number.isInteger(A) || A < 0 || A >= l
        ? t()
        : e.subarray((A + 1) * s, (A + 2) * s),
    u = [],
    d = new Set();
  for (let A = 0; A < 109; A++) {
    let L = n.getUint32(76 + A * 4, true);
    L !== r && u.push(L);
  }
  let f = n.getUint32(68, true),
    p = n.getUint32(72, true);
  if (p > l) return t();
  for (let A = 0; A < p; A++) {
    if (d.has(f)) return t();
    d.add(f);
    let L = c(f),
      y = new DataView(L.buffer, L.byteOffset, L.byteLength);
    for (let v = 0; v < s - 4; v += 4) {
      let k = y.getUint32(v, true);
      k !== r && u.push(k);
    }
    f = y.getUint32(s - 4, true);
  }
  if (
    (p && f !== i) ||
    u.length !== n.getUint32(44, true) ||
    !u.length ||
    u.length > l ||
    new Set(u).size !== u.length
  )
    return t();
  let h = [];
  for (let A of u) {
    let L = c(A),
      y = new DataView(L.buffer, L.byteOffset, L.byteLength);
    for (let v = 0; v < s; v += 4) h.push(y.getUint32(v, true));
  }
  for (let A of u) if (h[A] !== 4294967293) return t();
  for (let A of d) if (h[A] !== 4294967292) return t();
  let m = (A, L, y, v, k) => {
      let x = new Set(),
        C = [],
        N = A,
        F = 0;
      for (; N !== i;) {
        if (
          N >= L.length ||
          x.has(N) ||
          x.size >= v ||
          (u.includes(N) && L === h) ||
          (d.has(N) && L === h)
        )
          return t();
        x.add(N);
        let q = y(N);
        if ((C.push(q), (F += q.length), (N = L[N]), F > e.length)) return t();
      }
      if (k !== undefined && (k > F || F - k >= (C[0]?.length ?? 1)))
        return t();
      let z = new Uint8Array(k ?? F),
        j = 0;
      for (let q of C)
        (z.set(q.subarray(0, z.length - j), j),
          (j = Math.min(j + q.length, z.length)));
      return z;
    },
    g = m(n.getUint32(48, true), h, c, l),
    b = new DataView(g.buffer);
  if (!g.length || g.length % 128) return t();
  let w = [];
  for (let A = 0; A < g.length; A += 128) {
    let L = b.getUint8(A + 66),
      y = b.getUint16(A + 64, true),
      v = b.getUint32(A + 124, true);
    if (
      ![0, 1, 2, 5].includes(L) ||
      (L && (y < 2 || y > 64 || y % 2 || b.getUint16(A + y - 2, true)))
    )
      return t();
    let k = L
      ? new TextDecoder("utf-16le", { fatal: true }).decode(
          g.subarray(A, A + y - 2),
        )
      : "";
    if (
      L &&
      (!k || /[\x00-\x1f]/.test(k.replace(/^\x05/, "")) || (v && o === 4))
    )
      return t();
    w.push({
      name: k,
      type: L,
      left: b.getUint32(A + 68, true),
      right: b.getUint32(A + 72, true),
      child: b.getUint32(A + 76, true),
      start: b.getUint32(A + 116, true),
      size: b.getUint32(A + 120, true),
    });
  }
  if (
    w[0].type !== 5 ||
    w.some((A) => /^(EncryptedPackage|EncryptionInfo)$/i.test(A.name))
  )
    return t();
  let _ = [],
    S = new Set(),
    O = [w[0].child];
  for (; O.length;) {
    let A = O.pop();
    if (A === r) continue;
    if (!A || A >= w.length || S.has(A)) return t();
    S.add(A);
    let L = w[A];
    if (![1, 2].includes(L.type)) return t();
    (_.push(L), O.push(L.left, L.right));
  }
  let E = _.filter((A) => A.type === 2 && /^(Workbook|Book)$/i.test(A.name));
  if (E.length !== 1) return t();
  let R = E[0];
  if (R.size < 8 || R.size > e.length) return t();
  let I;
  if (R.size >= 4096) I = m(R.start, h, c, l, R.size);
  else {
    let A = n.getUint32(64, true);
    if (!A || A > l) return t();
    let L = m(n.getUint32(60, true), h, c, A, A * s),
      y = new DataView(L.buffer),
      v = [];
    for (let x = 0; x < L.length; x += 4) v.push(y.getUint32(x, true));
    let k = m(w[0].start, h, c, l, w[0].size);
    I = m(
      R.start,
      v,
      (x) => ((x + 1) * 64 > k.length ? t() : k.subarray(x * 64, (x + 1) * 64)),
      Math.ceil(k.length / 64),
      R.size,
    );
  }
  let U = new DataView(I.buffer),
    P = U.getUint16(4, true);
  if (
    U.getUint16(0, true) !== 2057 ||
    U.getUint16(2, true) < 4 ||
    ![1280, 1536].includes(P) ||
    U.getUint16(6, true) !== 5
  )
    return t();
  let T = false;
  for (let A = 0; A + 4 <= I.length;) {
    let L = U.getUint16(A, true),
      y = U.getUint16(A + 2, true);
    if (A + 4 + y > I.length || L === 47) return t();
    if (((A += 4 + y), L === 10)) {
      T = true;
      break;
    }
  }
  return T
    ? {
        paths: w.filter((A) => A.type).map((A) => A.name),
        workbookBytes: I.length,
        version: P === 1536 ? "BIFF8" : "BIFF5",
      }
    : t();
}

function DN(e, t, n) {
  if (
    !e ||
    e.schemaVersion !== 1 ||
    e.fileName !== n ||
    e.fileBytes !== t.length ||
    e.sha256 !== sha256HexDigest(t) ||
    !alternateFormatExtensions.includes(e.format) ||
    !n.toLowerCase().endsWith("." + e.format) ||
    e.execution !== "never" ||
    !Array.isArray(e.parts) ||
    e.parts.length > 2e3
  )
    throw Error("Invalid isolated format inspection.");
  eq.set(e, { signature: JSON.stringify(e), bytes: t.slice() });
}
var _ie = Object.freeze(
    alternateFormatExtensions.map((e) =>
      Object.freeze({
        format: e,
        read:
          e === "xltx"
            ? "Native template adapter: supported XLSX subset"
            : "External conversion to XLSX, then reviewed import",
        write:
          e === "xltx"
            ? "Derived XLTX template with loss review"
            : "Export reviewed XLSX, then Save As in the source application",
        preserve:
          "Exact original bytes in native JSON; derived output is not lossless",
        formulas:
          e === "xltx"
            ? "Supported formulas and cache comparison"
            : "Verify formulas and caches after external conversion",
        dates: "Verify date system, serials and displayed dates",
        styles: "Supported formatting only; compare with original",
        specialParts:
          "Original only unless supported by the XLSX codec; never execute macros or objects",
      }),
    ),
  ),
  wg = Object.freeze({
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    xltm: "application/vnd.ms-excel.template.macroEnabled.main+xml",
    xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
  });
function tq(e) {
  return [
    "Keep the original file. Inspection and conversion here never execute macros, objects or external refreshes.",
    e === "ods"
      ? "In a trusted source application that supports ODS, open the original and Save As an unencrypted Excel Workbook (.xlsx)."
      : "In a trusted source application that supports this format, open the original with macros and external refresh disabled. Save As an unencrypted Excel Workbook (.xlsx).",
    "Review the source application's compatibility warnings. Keep any needed automation, objects and unsupported metadata in the original.",
    "Choose that converted XLSX here. Compare formulas, date systems, dates, styles and special parts with the original; confirm the loss acknowledgement and review the XLSX compatibility report before replacing the workbook.",
    "For output in the original format, export a reviewed derived XLSX, open it in the source application and Save As the target format. Reopen there and verify the result. Renaming an extension does not convert a workbook.",
  ];
}
function vk(e) {
  if (e.length > workbookLimits.maxFileBytes)
    throw Error("Files must be 10 MB or smaller.");
  let t = yk(e),
    n = Object.create(null),
    i = new Set(),
    r = 0,
    o,
    a = new gc((s) => {
      if (!t.has(s.name) || i.has(s.name)) {
        o = Error("Alternate workbook archive directory mismatch.");
        return;
      }
      i.add(s.name);
      let l = 0,
        c = -1,
        u = [];
      ((s.ondata = (d, f, p) => {
        if (!o) {
          if (d) {
            o = Error("Corrupt alternate workbook archive.");
            return;
          }
          if (
            ((l += f.length),
            (r += f.length),
            l > t.get(s.name).size || r > 50 * 1024 * 1024)
          ) {
            ((o = Error(
              "Alternate workbook archive exceeds its expansion budget.",
            )),
              s.terminate());
            return;
          }
          if (((c = Tc(c, f)), u.push(f), p)) {
            if (
              l !== t.get(s.name).size ||
              (c ^ -1) >>> 0 !== t.get(s.name).crc
            ) {
              o = Error("Alternate workbook archive checksum mismatch.");
              return;
            }
            let h = new Uint8Array(l),
              m = 0;
            for (let g of u) (h.set(g, m), (m += g.length));
            n[s.name] = h;
          }
        }
      }),
        s.start());
    });
  a.register(mc);
  try {
    a.push(e, true);
  } catch {
    throw Error("Unable to inspect corrupt alternate workbook archive.");
  }
  if (o) throw o;
  if (Object.keys(n).length !== t.size)
    throw Error("Alternate workbook archive is incomplete.");
  return n;
}
function Ac(e, t) {
  let n = new TextDecoder("utf-8", { fatal: true }).decode(e);
  if (
    n.includes("\0") ||
    /<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(n)
  )
    throw Error("Alternate workbook XML must use UTF-8.");
  let i = new LN.SaxesParser({ xmlns: true }),
    r = 0,
    o = 0;
  (i.on("doctype", () => {
    throw Error("Alternate workbook XML declarations are unsupported.");
  }),
    i.on("opentag", (a) => {
      if (++r > 128 || ++o > 1e6)
        throw Error("Alternate workbook XML exceeds its inspection budget.");
      t(a);
    }),
    i.on("closetag", () => {
      r--;
    }),
    i.write(n).close());
}
function xg(e) {
  let t = (s) =>
    Object.fromEntries(
      Object.values(s.attributes).map((l) => [l.local, l.value]),
    );
  if (e.mimetype) {
    if (
      new TextDecoder().decode(e.mimetype) !==
        "application/vnd.oasis.opendocument.spreadsheet" ||
      !e["META-INF/manifest.xml"] ||
      !e["content.xml"] ||
      e["[Content_Types].xml"]
    )
      throw Error("Unsupported or mismatched OpenDocument workbook identity.");
    let s = false,
      l = false;
    if (
      (Ac(e["META-INF/manifest.xml"], (c) => {
        let u = t(c);
        if (c.uri === "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0") {
          if (c.local === "encryption-data")
            throw Error(
              "Encrypted OpenDocument workbooks require an unencrypted XLSX conversion.",
            );
          c.local === "file-entry" &&
            u["full-path"] === "/" &&
            u["media-type"] ===
              "application/vnd.oasis.opendocument.spreadsheet" &&
            (s = true);
        }
      }),
      Ac(e["content.xml"], (c) => {
        c.local === "spreadsheet" &&
          c.uri === "urn:oasis:names:tc:opendocument:xmlns:office:1.0" &&
          (l = true);
      }),
      !s || !l)
    )
      throw Error("Unsupported or mismatched OpenDocument workbook identity.");
    return "ods";
  }
  if (!e["[Content_Types].xml"] || !e["_rels/.rels"])
    throw Error("Unsupported alternate workbook package identity.");
  let n,
    i = "",
    r = 0,
    o = new Set();
  Ac(e["[Content_Types].xml"], (s) => {
    if (
      s.uri !== "http://schemas.openxmlformats.org/package/2006/content-types"
    )
      throw Error("Unsupported alternate workbook content-type namespace.");
    let l = t(s);
    if (s.local === "Override") {
      if (!l.PartName || o.has(l.PartName))
        throw Error("Duplicate or missing alternate workbook content type.");
      if (
        (o.add(l.PartName),
        l.PartName === "/xl/workbook.xml" || l.PartName === "/xl/workbook.bin")
      ) {
        if (n) throw Error("Ambiguous alternate workbook identity.");
        ((n = l.ContentType), (i = l.PartName.slice(1)));
      }
    }
  });
  let a = Object.entries(wg).find(([, s]) => s === n)?.[0];
  if (
    !a ||
    !e[i] ||
    i !== (a === "xlsb" ? "xl/workbook.bin" : "xl/workbook.xml")
  )
    throw Error("Unsupported alternate workbook content type.");
  if (
    (Ac(e["_rels/.rels"], (s) => {
      let l = t(s);
      if (
        s.uri ===
          "http://schemas.openxmlformats.org/package/2006/relationships" &&
        s.local === "Relationship" &&
        l.Type ===
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      ) {
        if (l.TargetMode || Zs("_rels/.rels", l.Target) !== i)
          throw Error("Invalid alternate workbook main relationship.");
        r++;
      }
    }),
    a === "xlsb")
  ) {
    let s = e[i];
    if (s.length < 3 || s[0] !== 131 || s[1] !== 1 || s[2] !== 0)
      throw Error(
        "Invalid binary workbook beginning record. Convert the original in its source application.",
      );
  } else {
    let s = true;
    Ac(e[i], (l) => {
      if (
        s &&
        ((s = false),
        l.local !== "workbook" ||
          ![
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
          ].includes(l.uri))
      )
        throw Error("Invalid alternate workbook XML identity.");
    });
  }
  if (r !== 1) throw Error("Invalid alternate workbook main relationship.");
  return a;
}
function Vd(e, t) {
  if (
    !t ||
    t.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(t) ||
    !e.length ||
    e.length > workbookLimits.maxFileBytes
  )
    throw Error(
      "Invalid alternate workbook filename or 10 MB file limit exceeded.",
    );
  let n = t.split(".").pop().toLowerCase();
  if (!alternateFormatExtensions.includes(n))
    throw Error("Choose an XLS, XLSB, XLSM, XLTX, XLTM or ODS source.");
  let i, r, o;
  if (e[0] === 208 && e[1] === 207) {
    let s = PN(e);
    ((i = "xls"),
      (r =
        s.version + " workbook in a CFB container; binary cells not decoded"),
      (o = s.paths.map((l) => ({
        path: l,
        bytes: /^(Workbook|Book)$/i.test(l) ? s.workbookBytes : 0,
        disposition: "original-only",
      }))));
  } else {
    let s = vk(e);
    ((i = xg(s)),
      (r =
        i === "ods"
          ? "ODF spreadsheet MIME, manifest and content"
          : "OOXML main content type and package relationship"));
    for (let [l, c] of Object.entries(s))
      /\.(xml|rels)$/i.test(l) && Ac(c, () => {});
    o = Object.entries(s).map(([l, c]) => ({
      path: l,
      bytes: c.length,
      sha256: sha256HexDigest(c),
      disposition: "original-only",
    }));
  }
  if (i !== n)
    throw Error(
      "The filename extension does not match the workbook content. Restore its original extension or convert it in the source application; renaming is not conversion.",
    );
  let a = {
    schemaVersion: 1,
    fileName: t,
    format: n,
    fileBytes: e.length,
    sha256: sha256HexDigest(e),
    identity: r,
    execution: "never",
    parts: o,
    steps: tq(n),
    warnings: [
      "Identity inspection is not workbook conversion, malware scanning or a guarantee of formula, date or style equivalence.",
      "All source parts are retained only in the exact original. No macros, binary formulas, embedded objects or external links execute during inspection.",
      n === "xltx"
        ? "The native template adapter uses the supported XLSX subset. Active or unsupported content may refuse import; review conversion losses."
        : "An external source application is required for editable import and target-format export. No external conversion has been performed here.",
      ...(n === "xls"
        ? [
            "CFB inventory lists directory names; non-workbook stream sizes, nested paths and cell contents are not resolved.",
          ]
        : []),
    ],
  };
  return (DN(a, e, t), a);
}
var UN = Bi(Mo(), 1);
function FN(e, t) {
  let n = (a, s) => {
      if (!e[a]) return;
      let l = new UN.SaxesParser({ xmlns: true });
      (l.on("opentag", s),
        l.on("doctype", () => {
          throw Error("External link metadata cannot contain DTDs.");
        }),
        l.write(Ao(e[a])).close());
    },
    i = new Map();
  n("xl/_rels/workbook.xml.rels", (a) => {
    a.local === "Relationship" &&
      a.attributes.Type?.value.endsWith("/externalLink") &&
      a.attributes.TargetMode?.value !== "External" &&
      i.set(
        a.attributes.Id.value,
        Zs("xl/_rels/workbook.xml.rels", a.attributes.Target.value),
      );
  });
  let r = [];
  if (
    (n("xl/workbook.xml", (a) => {
      if (a.local === "externalReference") {
        let s = Object.values(a.attributes).find(
          (l) => l.local === "id",
        )?.value;
        s && r.push(s);
      }
    }),
    r.length > 100)
  )
    throw Error(
      "More than 100 external workbook targets require an explicit source conversion.",
    );
  let o = [];
  return (
    r.forEach((a, s) => {
      let l = i.get(a);
      if (!l) return;
      let c = l.lastIndexOf("/"),
        u = l.slice(0, c + 1) + "_rels/" + l.slice(c + 1) + ".rels",
        d;
      (n(u, (f) => {
        if (
          f.local === "Relationship" &&
          f.attributes.Type?.value.endsWith("/externalLinkPath") &&
          f.attributes.TargetMode?.value === "External"
        ) {
          if (d !== undefined)
            throw Error("Ambiguous external workbook target relationship.");
          d = f.attributes.Target?.value;
        }
      }),
        d &&
          o.push({ key: `[${s + 1}]`, target: d, part: l, sourceSha256: t }));
    }),
    o
  );
}
function rq(e, t, n) {
  let i = vk(e);
  if (xg(i) !== t)
    throw Error(
      "The workbook content does not match the template conversion route.",
    );
  if (
    ((i["[Content_Types].xml"] = za(
      Ao(i["[Content_Types].xml"]).replace(wg[t], wg[n]),
    )),
    xg(i) !== n)
  )
    throw Error("The template content type could not be converted safely.");
  return Rd(i);
}
async function zN(e) {
  if (e.size > 10 * 1024 * 1024) throw Error("Files must be 10 MB or smaller.");
  let t = new Uint8Array(await e.arrayBuffer()),
    n = Vd(t, e.name);
  if (n.format !== "xltx")
    throw Error("Only XLTX templates use the native template adapter.");
  let i = rq(t, "xltx", "xlsx"),
    r = e.name.replace(/\.xltx$/i, "").slice(0, 236) + ".template-copy.xlsx",
    o = await kg({
      name: r,
      size: i.length,
      arrayBuffer: async () => i.buffer,
    });
  ((o.workbook.name =
    e.name.replace(/\.xltx$/i, "").slice(0, 200) || "Template workbook"),
    (o.workbook.alternateSource = retainAlternateFormatSource(
      t,
      e.name,
      "xltx",
      {
        method: "template-adapter",
        convertedFileName: r,
        convertedSha256: sha256HexDigest(i),
        reviewedAt: new Date().toISOString(),
        equivalence: "not-verified",
      },
    )));
  let a =
    "A new workbook was created from the XLTX template using the supported XLSX subset. The exact template is retained under File > Workbook formats. The XLSX part inventory refers to the converted intermediate, not the original template. Template fidelity requires comparison in the source application.";
  return (
    (o.warnings = [...new Set([...o.warnings, a])]),
    (o.report.warnings = o.warnings),
    (o.report.alternateOriginal = {
      fileName: n.fileName,
      byteLength: n.fileBytes,
      sha256: n.sha256,
      format: n.format,
    }),
    (o.report.fileName = e.name),
    (o.report.fileBytes = t.length),
    (o.report.format = "xltx"),
    (o.report.workbookName = o.workbook.name),
    o
  );
}
var jN = Bi(Mo(), 1);
var _g = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
var nq = "application/vnd.openxmlformats-officedocument.custom-properties+xml",
  wk = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
  iq = { scope: "workbook" };
function Ic(e, t, n, i) {
  let r = new jN.SaxesParser({ xmlns: true });
  (r.on("opentag", t),
    n && r.on("closetag", n),
    i && r.on("text", i),
    r.on("doctype", () => {
      throw Error("DTD is not permitted in preservation metadata.");
    }),
    r.write(Ao(e)).close());
}
function WN(e) {
  return e === "_rels/.rels"
    ? ""
    : e.replace(/(^|\/)\_rels\/([^/]+)\.rels$/, "$1$2");
}
function BN(e) {
  if (!e) return "_rels/.rels";
  let t = e.lastIndexOf("/");
  return `${e.slice(0, t + 1)}_rels/${e.slice(t + 1)}.rels`;
}
function oq(e) {
  let t = new Map(),
    n = new Map(),
    i = [],
    r = new Map();
  Ic(e["[Content_Types].xml"], (o) => {
    (o.local === "Override" &&
      t.set(
        (o.attributes.PartName?.value ?? "").slice(1),
        o.attributes.ContentType?.value ?? "",
      ),
      o.local === "Default" &&
        n.set(
          o.attributes.Extension?.value ?? "",
          o.attributes.ContentType?.value ?? "",
        ));
  });
  for (let [o, a] of Object.entries(e))
    o.endsWith(".rels") &&
      Ic(a, (s) => {
        if (s.local !== "Relationship") return;
        let l = s.attributes.TargetMode?.value === "External",
          c = s.attributes.Target?.value ?? "";
        i.push({
          owner: WN(o),
          rels: o,
          id: s.attributes.Id?.value ?? "",
          type: s.attributes.Type?.value ?? "",
          external: l,
          target: l ? c : Zs(o, c),
        });
      });
  return (
    Ic(e["xl/workbook.xml"], (o) => {
      if (o.local === "sheet") {
        let a = i.find(
          (s) =>
            s.owner === "xl/workbook.xml" &&
            s.id ===
              Object.values(o.attributes).find(
                (l) => l.local === "id" && l.uri === _g.slice(0, -1),
              )?.value,
        );
        a && r.set(a.target, o.attributes.name?.value ?? "Unknown worksheet");
      }
    }),
    {
      relationships: i,
      sheets: r,
      contentType: (o) =>
        t.get(o) ?? n.get(o.split(".").at(-1)) ?? "undeclared",
      types: t,
    }
  );
}
function sq(e, t) {
  return e.endsWith("/")
    ? "directory"
    : e === "[Content_Types].xml" || e.endsWith(".rels")
      ? "package relationships"
      : /vba|macrosheet|activeX|embedding/i.test(e + t)
        ? "active content"
        : /^customXml\//.test(e)
          ? "custom XML"
          : e === "docProps/custom.xml"
            ? "custom properties"
            : /^xl\/charts\//.test(e)
              ? "charts"
              : /^xl\/(?:model|dataModel)\//i.test(e) ||
                  /model\+data|dataModel/i.test(t)
                ? "Excel Data Model"
                : /^xl\/(pivotTables|pivotCache)\//.test(e)
                  ? "pivots"
                  : /connections|queryTables|externalLinks/.test(e)
                    ? "connections and external links"
                    : /slicers|slicerCaches|timelines/.test(e)
                      ? "slicers and timelines"
                      : /threadedComments|persons/.test(e)
                        ? "threaded comments"
                        : /^_xmlsignatures\//.test(e)
                          ? "signatures"
                          : /^xl\/worksheets\/[^/]+\.xml$/.test(e)
                            ? "worksheet"
                            : e === "xl/workbook.xml"
                              ? "workbook and names"
                              : /^xl\/tables\//.test(e)
                                ? "tables"
                                : /^xl\/drawings\//.test(e)
                                  ? "drawings"
                                  : /^xl\/media\//.test(e)
                                    ? "media"
                                    : /^xl\/comments/.test(e)
                                      ? "notes"
                                      : e === "xl/sharedStrings.xml"
                                        ? "shared strings"
                                        : e === "xl/styles.xml" ||
                                            /^xl\/theme\//.test(e)
                                          ? "formatting"
                                          : /^docProps\//.test(e)
                                            ? "document properties"
                                            : "unknown part";
}
function xk(e, t) {
  let n = oq(e),
    i = new Map(),
    r = new Map();
  for (let p of n.relationships)
    (i.set(p.owner, [...(i.get(p.owner) ?? []), p]),
      p.external || r.set(p.target, [...(r.get(p.target) ?? []), p]));
  let o = new Map();
  for (let [p, h] of n.sheets) {
    let m = [p],
      g = new Set();
    for (; m.length;) {
      let b = m.pop();
      if (g.has(b)) continue;
      g.add(b);
      let w = o.get(b) ?? new Set();
      (w.add(h), o.set(b, w));
      for (let _ of i.get(b) ?? []) _.external || m.push(_.target);
    }
  }
  let a = new Set(),
    s = new Map(),
    l = Object.keys(e).filter(
      (p) => p.startsWith("customXml/") && !p.endsWith("/"),
    );
  if (l.length) {
    let p = new Set(l),
      h = n.relationships.filter(
        (_) => p.has(_.owner) || (!_.external && p.has(_.target)),
      ),
      m = !Object.keys(e).some((_) => /xmlMaps|singleXmlCells/i.test(_));
    ((m &&= l.every(
      (_) =>
        _.endsWith(".rels") ||
        (/\.xml$/.test(_) &&
          ["application/xml", "text/xml", wk].includes(n.contentType(_))),
    )),
      (m &&= h.every(
        (_) =>
          !_.external &&
          ((_.owner === "xl/workbook.xml" &&
            _.type === _g + "customXml" &&
            p.has(_.target)) ||
            (p.has(_.owner) &&
              p.has(_.target) &&
              _.type === _g + "customXmlProps")),
      )),
      (m &&= h.every((_) =>
        _.owner === "xl/workbook.xml"
          ? ["application/xml", "text/xml"].includes(n.contentType(_.target))
          : n.contentType(_.owner) !== wk && n.contentType(_.target) === wk,
      )));
    let g = h.filter((_) => _.owner === "xl/workbook.xml"),
      b = new Set(),
      w = g.map((_) => _.target);
    for (; w.length;) {
      let _ = w.pop();
      if (!b.has(_)) {
        (b.add(_), e[BN(_)] && b.add(BN(_)));
        for (let S of i.get(_) ?? []) S.external || w.push(S.target);
      }
    }
    m &&= g.length > 0 && l.every((_) => b.has(_));
    for (let _ of l)
      m
        ? a.add(_)
        : s.set(
            _,
            "Custom XML has bindings, unrecognized types, orphaned parts or relationships outside the passive metadata island. Retained only in the original.",
          );
  }
  let c = "docProps/custom.xml";
  if (e[c]) {
    let p = r.get(c) ?? [],
      h =
        n.contentType(c) === nq &&
        p.length === 1 &&
        p[0].owner === "" &&
        p[0].type === _g + "custom-properties" &&
        !i.get(c)?.length;
    (Ic(e[c], (m) => {
      (m.local === "property" && m.attributes.linkTarget && (h = false),
        (m.uri ===
          "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" &&
          ["Properties", "property"].includes(m.local)) ||
          (m.uri ===
            "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" &&
            [
              "lpwstr",
              "lpstr",
              "bstr",
              "i1",
              "i2",
              "i4",
              "i8",
              "int",
              "ui1",
              "ui2",
              "ui4",
              "ui8",
              "uint",
              "r4",
              "r8",
              "decimal",
              "bool",
              "date",
              "filetime",
              "empty",
              "null",
            ].includes(m.local)) ||
          (h = false));
    }),
      h
        ? a.add(c)
        : s.set(
            c,
            "Custom properties contain links, unsupported value types or relationships; retained only in the original.",
          ));
  }
  let u = new Set();
  if (e["xl/sharedStrings.xml"]) {
    let p = -1;
    Ic(e["xl/sharedStrings.xml"], (h) => {
      (h.local === "si" && p++, h.local === "r" && u.add(p));
    });
  }
  let d = 0,
    f = Object.entries(e)
      .sort(([p], [h]) => p.localeCompare(h))
      .map(([p, h]) => {
        let m = sq(p, n.contentType(p)),
          g = [...(o.get(p.endsWith(".rels") ? WN(p) : p) ?? [])].map((O) => ({
            sheet: O,
            scope: "sheet",
          }));
        g.length ||
          g.push(
            [
              "charts",
              "pivots",
              "drawings",
              "notes",
              "unknown part",
              "media",
            ].includes(m)
              ? {
                  scope: "unresolved",
                  detail:
                    "No worksheet relationship identifies the affected cells.",
                }
              : iq,
          );
        let b = "dropped",
          w =
            "No safe editable mapping. Exact source bytes remain in the original file.";
        (m === "Excel Data Model" &&
          (w =
            "Imported Excel Data Model, relationships and DAX are not executed or translated. Worksheet results are historical snapshots. Exact model bytes remain only in the original source; rebuild reviewed local typed tables and relationships to calculate new results."),
          ["worksheet", "workbook and names", "tables", "notes"].includes(m) &&
            ((b = "editable"),
            (w =
              "Supported content is editable and regenerated; inspect the feature details and conversion notices for losses. XML bytes are not preserved.")),
          [
            "package relationships",
            "shared strings",
            "formatting",
            "document properties",
            "drawings",
          ].includes(m) &&
            ((b = "flattened"),
            (w =
              "Supported content is regenerated or flattened. Unsupported settings and original XML are retained only in the source.")),
          m === "media" &&
            /^image\/(png|jpeg)$/.test(n.contentType(p)) &&
            ((b = "flattened"),
            (w =
              "Supported worksheet pictures are re-encoded with their supported placement; orphaned images and chart assets are not carried over.")),
          a.has(p) &&
            ((b = "preserved-opaque"),
            (w =
              "Passive metadata snapshot copied byte-for-byte, without interpreting or refreshing its contents. Its original meaning may be stale after edits.")),
          s.has(p) && (w = s.get(p)),
          m === "signatures" &&
            (w =
              "Signatures cannot authenticate an edited workbook; retained only in the original file."));
        let _ = [],
          S = (O, E, R, I, U) => {
            if (++d > xlsxInspectionLimits.maxDetails)
              throw Error(
                "XLSX compatibility details exceed the 100,000-entry budget. Split the source workbook before importing.",
              );
            _.push({
              feature: O,
              disposition: E,
              location: {
                ...g[0],
                ...(R ? { range: R, scope: "cell" } : {}),
                ...(I ? { detail: I } : {}),
              },
              reason:
                U ??
                (E === "editable"
                  ? "Supported subset is editable; review conversion notices for unsupported attributes."
                  : "This source feature is not preserved as an editable feature in the derived workbook."),
            });
          };
        if (
          m === "worksheet" ||
          m === "workbook and names" ||
          m === "tables" ||
          m === "drawings" ||
          m === "pivots" ||
          m === "charts"
        ) {
          let O = "",
            E = "",
            R,
            I = "",
            U = false,
            P = false,
            T = "";
          Ic(
            h,
            (A) => {
              let L = (y) => A.attributes[y]?.value;
              (A.local === "c" &&
                ((O = L("r") ?? ""),
                (E = L("t") ?? ""),
                E === "e" && S("error cell", "flattened", O)),
                A.local === "v" && ((P = true), (T = "")),
                A.local === "f" && ((U = true), (I = "")),
                A.local === "definedName" &&
                  ((R = L("name")),
                  S("defined name", "editable", undefined, R)),
                A.local === "dataValidation" &&
                  S("data validation", "editable", L("sqref"), L("type")),
                A.local === "conditionalFormatting" &&
                  S("conditional formatting", "editable", L("sqref")),
                A.local === "mergeCell" &&
                  S("merged cells", "editable", L("ref")),
                A.local === "table" &&
                  S("structured table", "editable", L("ref"), L("name")),
                A.local === "location" &&
                  S("pivot output", "dropped", L("ref")),
                A.local === "worksheetSource" &&
                  S("pivot source", "dropped", L("ref"), L("sheet")),
                A.local === "r" &&
                  O &&
                  E === "inlineStr" &&
                  S("rich text", "flattened", O),
                [
                  "sheetProtection",
                  "workbookProtection",
                  "autoFilter",
                  "extLst",
                  "hyperlink",
                  "legacyDrawingHF",
                  "webPublishItems",
                  "controls",
                  "oleObjects",
                  "pivotTable",
                  "pivotTablePart",
                  "slicer",
                  "timeline",
                  "sp",
                  "graphicFrame",
                ].includes(A.local) &&
                  S(A.local, "dropped", L("ref"), L("name")));
            },
            (A) => {
              (A.local === "v" &&
                ((P = false),
                O &&
                  E === "s" &&
                  u.has(Number(T)) &&
                  S("rich text", "flattened", O)),
                A.local === "c" && (O = ""),
                A.local === "f" &&
                  ((U = false),
                  m === "charts"
                    ? S(
                        "chart source reference",
                        "dropped",
                        undefined,
                        I.slice(0, 1024),
                      )
                    : /(?:\[|https?:|file:|WEBSERVICE|HYPERLINK|RTD|DDE|CALL\s*\(|EXEC\s*\()/i.test(
                        I,
                      ) &&
                      S(
                        "external/action formula",
                        "flattened",
                        O,
                        undefined,
                        "External/action references may be neutralized to literal text; the formula cache report provides the actual conversion outcome.",
                      )));
            },
            (A) => {
              (U && (I += A), P && (T += A));
            },
          );
        }
        for (let O of i.get(p) ?? [])
          O.external &&
            S(
              "external relationship",
              "dropped",
              undefined,
              O.type.split("/").at(-1),
              "External target is never fetched or carried into the derived copy; its URI remains only in the original.",
            );
        return {
          path: p,
          contentType: p.endsWith("/") ? "directory" : n.contentType(p),
          bytes: h.length,
          sha256: sha256HexDigest(h),
          family: m,
          disposition: b,
          reason: w,
          locations: g,
          features: _,
          retainedInOriginal: true,
        };
      });
  return aq({
    schemaVersion: 1,
    policy: "passive-metadata-v1",
    complete: true,
    phase: "eligible-for-export",
    sourceSha256: t,
    parts: f,
    counts: { editable: 0, "preserved-opaque": 0, flattened: 0, dropped: 0 },
    opaqueBytes: 0,
    note: "All package entries are inventoried. Locations come from source relationships and XML; unresolved locations are explicit. Dispositions describe the derived copy, while all parts remain byte-exact in the retained original. Editable parts can also contain flattened or dropped features. No source part is executed or fetched.",
  });
}
function aq(e) {
  ((e.counts = {
    editable: 0,
    "preserved-opaque": 0,
    flattened: 0,
    dropped: 0,
  }),
    (e.opaqueBytes = 0));
  for (let t of e.parts)
    (e.counts[t.disposition]++,
      t.disposition === "preserved-opaque" && (e.opaqueBytes += t.bytes));
  return e;
}
var zie = Object.freeze({
  storage: "IEEE-754 binary64",
  excelSignificantDigits: 15,
  rounding:
    "ROUND uses decimal ties away from zero; monetary rounding is explicit, never automatic.",
  display:
    "General displays 15 significant digits. Display precision does not round stored values or calculations.",
});
function VN(e) {
  let t = /^[+-]?(\d+\.?\d*|\.\d+)(?:e([+-]?\d+))?$/i.exec(e.trim());
  if (!t) return;
  let n = t[1].replace(".", "").replace(/^0+/, "").replace(/0+$/, ""),
    i = Number(e);
  if (!Number.isFinite(i)) return "outside the finite numeric range";
  if (n && (i === 0 || Math.abs(i) < 22250738585072014e-324))
    return "below Excel's normal floating-point range";
  if (n.length > 15) return "more than 15 significant digits";
  if (Number.isInteger(i) && !Number.isSafeInteger(i))
    return "outside the exact integer range";
}
var qN = Bi(Mo(), 1);
function XN(e) {
  let t = new Map(),
    n = new qN.SaxesParser({ xmlns: true }),
    i = "",
    r = "",
    o = "",
    a = false;
  n.on("opentag", (l) => {
    l.local === "c"
      ? ((i = String(l.attributes.r?.value ?? "")),
        (r = String(l.attributes.t?.value ?? "n")))
      : l.local === "v" && i && ((a = true), (o = ""));
  });
  let s = (l) => {
    if (a && ((o += l), o.length > 1024))
      throw Error("An XLSX numeric token exceeds 1,024 characters.");
  };
  return (
    n.on("text", s),
    n.on("cdata", s),
    n.on("closetag", (l) => {
      if (l.local === "v") {
        if (r === "n" && o.trim()) {
          let c = o.trim();
          if (
            !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(c) ||
            !Number.isFinite(Number(c))
          )
            throw Error(`Invalid finite numeric value at ${i}.`);
          t.set(i, c);
        }
        a = false;
      } else l.local === "c" && (i = "");
    }),
    n.write(e).close(),
    t
  );
}
var Xd = (e) => {
  let t = parseRangeAddress(e);
  return `${Ks(t.firstRow, t.firstCol)}:${Ks(t.lastRow, t.lastCol)}`;
};
var qd = (e, t) =>
  e?.argb && /^[a-f\d]{8}$/i.test(e.argb)
    ? `#${e.argb.slice(-6).toLowerCase()}`
    : t;
function lq(e) {
  let t = e.map(_t).sort((o, a) => o.row - a.row || o.col - a.col),
    n = [];
  for (let o of t) {
    let a = n.at(-1);
    a && a.lastRow === o.row && a.lastCol + 1 === o.col
      ? (a.lastCol = o.col)
      : n.push({
          firstRow: o.row,
          lastRow: o.row,
          firstCol: o.col,
          lastCol: o.col,
        });
  }
  let i = [],
    r = new Map();
  for (let o of n) {
    let a = `${o.firstCol}:${o.lastCol}`,
      s = r.get(a);
    s && s.lastRow + 1 === o.firstRow
      ? (s.lastRow = o.lastRow)
      : (i.push(o), r.set(a, o));
  }
  return i.map(
    (o) => `${Ks(o.firstRow, o.firstCol)}:${Ks(o.lastRow, o.lastCol)}`,
  );
}
function ZN(e) {
  let t = [];
  for (let n of e.matchAll(
    /<(?:\w+:)?dataValidation\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?dataValidation>)/g,
  )) {
    let i = (s) =>
        new RegExp("\\b" + s + `=["\\']([^"\\']*)["\\']`).exec(n[1])?.[1],
      r = (i("sqref") ?? "").trim().split(/\s+/).filter(Boolean);
    if (!r.length) continue;
    let o = Xd(r[0]).split(":")[0];
    if (t.length + r.length > 1e3)
      throw Error(
        "XLSX has more than 1,000 validation areas; narrow the source rules before import.",
      );
    let a = [
      ...(n[2] ?? "").matchAll(
        /<(?:\w+:)?formula[12]\b[^>]*>([\s\S]*?)<\/(?:\w+:)?formula[12]>/g,
      ),
    ].map((s) => s[1].trim());
    for (let s of r)
      t.push({
        range: Xd(s),
        anchor: o,
        operands: a,
        dropdown: !["1", "true"].includes(i("showDropDown") ?? "0"),
        alert: ["1", "true"].includes(i("showErrorMessage") ?? "0"),
      });
  }
  return t;
}
var cq = {
  between: "between",
  notBetween: "notBetween",
  greater: "greaterThan",
  greaterEqual: "greaterThanOrEqual",
  less: "lessThan",
  lessEqual: "lessThanOrEqual",
  equal: "equal",
  notEqual: "notEqual",
};
function uq(e) {
  let t = {
      list: "list",
      decimal: "number",
      whole: "integer",
      textLength: "textLength",
      date: "date",
      time: "time",
      custom: "custom",
    }[e.type],
    n = Object.entries(cq).find(([, u]) => u === e.operator)?.[0];
  if (!t || (!n && !["list", "custom"].includes(t))) return null;
  let i = e.formulae ?? [],
    r = (u) =>
      u instanceof Date
        ? (u.getTime() - Date.UTC(1899, 11, 30)) / 864e5
        : typeof u == "number"
          ? u
          : t === "time" &&
              typeof u == "string" &&
              /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(u)
            ? Number(u)
            : NaN,
    o = ["list", "custom"].includes(t) ? 0 : r(i[0]),
    a = ["list", "custom"].includes(t)
      ? 0
      : ["between", "notBetween"].includes(n)
        ? r(i[1])
        : o;
  if (!Number.isFinite(o) || !Number.isFinite(a)) return null;
  let s = [],
    l;
  if (
    (t === "list" &&
      (typeof i[0] != "string" ||
        !/^"[^"]*"$/.test(i[0]) ||
        ((s = i[0].slice(1, -1).split(",")),
        !s.length ||
          s.length > 200 ||
          s.some((u) => !u.length || u.length > 1e3)))) ||
    (t === "custom" &&
      (typeof i[0] != "string" ||
        ((l = "=" + i[0].replace(/^=/, "")),
        l.length > 8192 || l.length < 2 || kk(l))))
  )
    return null;
  let c = e.errorStyle?.toLowerCase() ?? "stop";
  return ["stop", "warning", "information"].includes(c)
    ? {
        kind: t,
        operator: n ?? "equal",
        min: o,
        max: a,
        list: s,
        allowBlank: !!e.allowBlank,
        message: String(e.error ?? "Choose a valid value.").slice(0, 500),
        errorStyle: c,
        errorTitle: String(e.errorTitle ?? "Invalid value").slice(0, 100),
        showInputMessage: !!e.showInputMessage,
        inputTitle: String(e.promptTitle ?? "").slice(0, 100),
        inputMessage: String(e.prompt ?? "").slice(0, 500),
        ...(l ? { formula: l } : {}),
      }
    : null;
}
var kk = (e) => {
    let t = stripFormulaStringLiterals(e).replace(/"(?:[^"]|"")*"/g, '""');
    return (
      /[\[\]|]/.test(t) ||
      /(?:https?:|file:|javascript:|data:|\\\\)/i.test(t) ||
      /\b(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
        t.replace(/'(?:[^']|'')*'(?=\s*!)/g, "Sheet"),
      )
    );
  },
  KN = (e) => '"' + e.replace(/"/g, '""') + '"';
function dq(e, t) {
  return "ISNUMBER(FIND(LOWER(" + KN(e) + "),LOWER(" + t + '&"")))';
}
function HN(e, t) {
  let n = parseRangeAddress(t),
    i = Ks(n.firstRow, n.firstCol),
    r = Xd(t).replace(/([A-Z]+)(\d+)/g, "$$$1$$$2"),
    o =
      "SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" +
      i +
      ',"~","~~"),"*","~*"),"?","~?")';
  return (
    "AND(LEN(" +
    i +
    ")>0,COUNTIF(" +
    r +
    ',"="&' +
    o +
    ")" +
    (e === "duplicate" ? ">1" : "=1") +
    ")"
  );
}
function fq(e) {
  if (e === null || e === "") return true;
  if (typeof e == "number") return Number.isFinite(e);
  if (typeof e != "string") return false;
  let t = e.trim() !== "" && Number.isFinite(Number(e));
  return !(
    /[^\x20-\x7e]/.test(e) ||
    /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|ERROR!|CYCLE!|SPILL!|CALC!|LIMIT!|BLOCKED!)/.test(
      e,
    ) ||
    e.replace(/[~*?]/g, "~~").length > 254 ||
    (t && String(Number(e)) !== e) ||
    (!t && /^[\s+-]*\d/.test(e))
  );
}
function JN(e) {
  let t = {},
    n = new Set();
  for (let i of e.matchAll(/<(?:\w+:)?cfRule\b([^>]*)>/g)) {
    let r = /\bpriority=["']([^"']*)["']/.exec(i[1])?.[1],
      o = /\bstopIfTrue=["']([^"']+)["']/.exec(i[1])?.[1];
    if (r === undefined && /\bid=["']/.test(i[1])) continue;
    let a = Number(r);
    if (!r || !/^\d+$/.test(r) || !Number.isSafeInteger(a) || a < 1 || n.has(a))
      throw Error("Invalid or duplicate conditional-format priority.");
    if ((n.add(a), r && o !== undefined)) {
      if (!["1", "0", "true", "false"].includes(o))
        throw Error("Invalid conditional-format stop flag.");
      t[String(a)] = o === "1" || o === "true";
    }
  }
  for (let i of e.matchAll(/<(?:\w+:)?cfvo\b([^>]*)>/g)) {
    let r = /\bgte=["']([^"']*)["']/.exec(i[1])?.[1];
    if (r !== undefined && r !== "1" && r !== "true")
      throw Error(
        "Strict or malformed conditional icon thresholds are not supported. Use inclusive thresholds or native JSON.",
      );
  }
  return t;
}
function hq(e, t) {
  let n =
      e.style?.fill?.type === "pattern"
        ? qd(e.style.fill.fgColor, "#fff2cc")
        : "#fff2cc",
    i = qd(e.style?.font?.color, "#000000"),
    r = {
      id: GN(),
      priority: Number.isFinite(e.priority) ? e.priority : 1e4,
      range: Xd(t),
      value: "",
      second: "",
      fill: n,
      color: i,
    },
    o = (c) =>
      typeof c == "number" && Number.isFinite(c)
        ? String(c)
        : typeof c == "string" && c.trim() !== "" && Number.isFinite(Number(c))
          ? c
          : typeof c == "string" && /^"(?:[^"]|"")*"$/.test(c)
            ? c.slice(1, -1).replace(/""/g, '"')
            : null;
  if (
    e.type === "containsText" &&
    (!e.operator || e.operator === "containsText")
  ) {
    let c = parseRangeAddress(t),
      u = Ks(c.firstRow, c.firstCol),
      d = typeof e.text == "string" ? e.text : "",
      f = e.formulae?.[0],
      p =
        typeof f == "string" && f.trim()
          ? f.replace(/^=/, "")
          : "NOT(ISERROR(SEARCH(" + KN(d) + "," + u + ")))";
    return p.length > 4095 || kk("=" + p)
      ? null
      : { ...r, kind: "formula", formula: "=" + p };
  }
  if (
    e.type === "colorScale" &&
    e.cfvo?.length === 2 &&
    e.cfvo[0].type === "min" &&
    e.cfvo[1].type === "max"
  )
    return {
      ...r,
      kind: "colorScale",
      minColor: qd(e.color?.[0], "#f8e99b"),
      maxColor: qd(e.color?.[1], "#63be7b"),
    };
  if (
    e.type === "dataBar" &&
    e.cfvo?.length === 2 &&
    ["min", "autoMin"].includes(e.cfvo[0].type) &&
    ["max", "autoMax"].includes(e.cfvo[1].type) &&
    e.showValue !== false &&
    e.direction !== "rightToLeft"
  )
    return { ...r, kind: "dataBar", fill: qd(e.color, "#4472c4") };
  if (
    e.type === "iconSet" &&
    e.iconSet === "3Arrows" &&
    !e.reverse &&
    !e.custom &&
    e.showValue !== false &&
    e.cfvo?.length === 3 &&
    e.cfvo[1].type === e.cfvo[2].type &&
    ["num", "percent"].includes(e.cfvo[1].type)
  ) {
    let [c, u, d] = e.cfvo,
      f = u.value,
      p = d.value;
    return !(c.type === "min" || (c.type === "percent" && c.value === 0)) ||
      typeof f != "number" ||
      typeof p != "number" ||
      !Number.isFinite(f) ||
      !Number.isFinite(p) ||
      f >= p ||
      (u.type === "percent" && (f < 0 || p > 100))
      ? null
      : {
          ...r,
          kind: "iconSet",
          thresholdType: u.type === "num" ? "number" : "percent",
          iconThresholds: [f, p],
        };
  }
  if (e.type === "expression") {
    let c = String(e.formulae?.[0] ?? "").replace(/^=/, "");
    if (!c || c.length > 4095 || kk("=" + c)) return null;
    if (c.toUpperCase() === HN("duplicate", t).toUpperCase())
      return { ...r, kind: "duplicate" };
    if (c.toUpperCase() === HN("unique", t).toUpperCase())
      return { ...r, kind: "unique" };
    let u =
      /^ISNUMBER\(FIND\(LOWER\("((?:[^"]|"")*)"\),LOWER\(([A-Z]+[1-9]\d*)&""\)\)\)$/i.exec(
        c,
      );
    if (u) {
      let h = u[1].replace(/""/g, '"'),
        m = parseRangeAddress(t),
        g = Ks(m.firstRow, m.firstCol);
      if (h.length <= 1e3 && c.toUpperCase() === dq(h, g).toUpperCase())
        return { ...r, kind: "contains", value: h };
    }
    let d = parseRangeAddress(t),
      f = Ks(d.firstRow, d.firstCol),
      p = {
        [`LEN(${f})=0`]: "blank",
        [`LEN(${f})>0`]: "notBlank",
        [`ISERROR(${f})`]: "error",
        [`NOT(ISERROR(${f}))`]: "notError",
      };
    return p[c]
      ? { ...r, kind: p[c] }
      : { ...r, kind: "formula", formula: "=" + c };
  }
  if (e.type !== "cellIs") return null;
  let a = {
      greaterThan: "greater",
      lessThan: "less",
      equal: "equal",
      between: "between",
    }[e.operator],
    s = o(e.formulae?.[0]),
    l = a === "between" ? o(e.formulae?.[1]) : "";
  return !a || s === null || l === null
    ? null
    : { ...r, kind: a, value: s, second: l };
}
function YN(e, t, n = {}, i = []) {
  let r = [],
    o = {},
    a = e.model.merges ?? [];
  a.length && (o.merges = a.map(Xd));
  let s = e.dataValidations.model,
    l = new Map();
  for (let [h, m] of Object.entries(s)) {
    let g = _t(h),
      b = i.find((_) => {
        let S = parseRangeAddress(_.range);
        return (
          g.row >= S.firstRow &&
          g.row <= S.lastRow &&
          g.col >= S.firstCol &&
          g.col <= S.lastCol
        );
      }),
      w = JSON.stringify({ rule: m, area: b });
    (l.has(w) || l.set(w, { rule: m, addresses: [], area: b }),
      l.get(w).addresses.push(h));
  }
  let c = [];
  for (let { rule, addresses, area } of l.values()) {
    let b = !["list", "custom"].includes(rule.type),
      w = area?.operands.every((S) =>
        /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(S),
      ),
      _ =
        b && area && !w
          ? null
          : uq(
              b && area
                ? { ...rule, formulae: area.operands.map(Number) }
                : rule,
            );
    if (!_) {
      r.push(
        "An unsupported validation formula, operator or external list was omitted; source cells were retained.",
      );
      continue;
    }
    area &&
      !area.alert &&
      r.push(
        "An imported validation has error alerts disabled. Local validation requires Stop or explicit Warning/Information review; source values were retained.",
      );
    for (let S of lq(addresses)) {
      if (c.length >= 100) {
        r.push(
          "Additional validation areas beyond the 100-rule limit were omitted; source cells were retained.",
        );
        break;
      }
      let O = _t(area?.anchor ?? S.split(":")[0]),
        E = _t(S.split(":")[0]);
      c.push({
        ..._,
        id: GN(),
        range: S,
        ...(area ? { showDropdown: area.dropdown } : {}),
        ...(_.formula
          ? {
              formula: shiftFormulaReferences(
                _.formula,
                E.row - O.row,
                E.col - O.col,
              ),
            }
          : {}),
      });
    }
  }
  c.length && (o.validations = c);
  let u = [],
    d = new Set(),
    f = 0;
  for (let h of e.conditionalFormattings ?? []) {
    let m = h.ref.trim().split(/\s+/).filter(Boolean);
    if (!m.length)
      throw Error("A conditional formatting rule requires an applied range.");
    let g = parseRangeAddress(m[0]);
    for (let b of h.rules) {
      if (
        !Number.isSafeInteger(b.priority) ||
        b.priority < 1 ||
        d.has(b.priority)
      )
        throw Error("Invalid or duplicate conditional-format priority.");
      d.add(b.priority);
      for (let w of m) {
        let _ = parseRangeAddress(w),
          S = b;
        (b.type === "expression" || b.type === "containsText") &&
          b.formulae?.length &&
          (S = {
            ...b,
            formulae: b.formulae.map((E) =>
              typeof E == "string"
                ? shiftFormulaReferences(
                    "=" + E.replace(/^=/, ""),
                    _.firstRow - g.firstRow,
                    _.firstCol - g.firstCol,
                  ).slice(1)
                : E,
            ),
          });
        let O = hq(S, w);
        if (O && (O.kind === "duplicate" || O.kind === "unique")) {
          let E = true;
          for (let [R, I] of Object.entries(t.cells)) {
            if (++f > 1e6) {
              E = false;
              break;
            }
            let U = _t(R);
            if (
              U.row < _.firstRow ||
              U.row > _.lastRow ||
              U.col < _.firstCol ||
              U.col > _.lastCol
            )
              continue;
            let P = I.raw;
            if (
              (typeof I.raw == "string" &&
                I.raw.startsWith("=") &&
                !I.literal &&
                ((P = e.getCell(R).result),
                P === undefined && /^=\s*""\s*$/.test(I.raw) && (P = "")),
              !fq(P))
            ) {
              E = false;
              break;
            }
          }
          E ||
            ((O = {
              ...O,
              kind: "formula",
              formula: "=" + String(S.formulae?.[0] ?? "").replace(/^=/, ""),
            }),
            r.push(
              "A duplicate/unique condition was retained as a formula because its population cannot use the native literal comparison; formula-engine limits apply.",
            ));
        }
        O && u.length < 100
          ? u.push({
              ...O,
              ...(n[String(b.priority)] !== undefined
                ? { stopIfTrue: n[String(b.priority)] }
                : {}),
            })
          : r.push(
              "An unsupported conditional-format rule or a rule beyond the 100-rule limit was omitted; base cell formatting was retained.",
            );
      }
    }
  }
  if (u.length) {
    let h = [...u].sort((g, b) => (g.priority ?? 1e4) - (b.priority ?? 1e4)),
      m = new Map(h.map((g, b) => [g.id, b + 1]));
    o.conditionalFormats = u.map((g) => ({ ...g, priority: m.get(g.id) }));
  }
  return (
    e.model.sheetProtection?.sheet &&
      r.push(
        "Source Excel sheet-password protection is not imported. The local protection feature and host authorization are separate controls.",
      ),
    validateWorksheetFeatures(o, t),
    r
  );
}
var QN = Bi(Mo(), 1);
function e8(e) {
  let t = {
      mode: "auto",
      iterative: false,
      maxIterations: null,
      maxChange: null,
    },
    n = new QN.SaxesParser({ xmlns: true }),
    i = false;
  return (
    n.on("opentag", (r) => {
      if (r.local !== "calcPr" || !/spreadsheetml/.test(r.uri)) return;
      if (i)
        throw Error(
          "Duplicate XLSX calculation properties. Repair the source workbook.",
        );
      i = true;
      let o = (d) => {
          let f = r.attributes[d];
          return f && typeof f != "string" ? f.value : f;
        },
        a = o("calcMode") ?? "auto",
        s = o("iterate"),
        l = o("iterateCount"),
        c = o("iterateDelta"),
        u = o("fullPrecision");
      if (
        !["auto", "autoNoTable", "manual"].includes(a) ||
        (u !== undefined && !["true", "false", "1", "0"].includes(u)) ||
        (s !== undefined && !["true", "false", "1", "0"].includes(s)) ||
        (l !== undefined &&
          (!/^\d+$/.test(l) ||
            !Number.isSafeInteger(Number(l)) ||
            Number(l) < 1)) ||
        (c !== undefined &&
          (!c.trim() || !Number.isFinite(Number(c)) || Number(c) < 0))
      )
        throw Error(
          "Invalid XLSX calculation properties. Repair the source workbook.",
        );
      t = {
        mode: a,
        iterative: s === "true" || s === "1",
        maxIterations: l === undefined ? null : Number(l),
        maxChange: c === undefined ? null : Number(c),
        ...(u === undefined
          ? {}
          : { fullPrecision: u === "true" || u === "1" }),
      };
    }),
    n.on("doctype", () => {
      throw Error("XLSX calculation properties cannot contain a DTD.");
    }),
    n.write(e).close(),
    t
  );
}
function t8(e, t, n) {
  ((e.calculation = {
    ...defaultCalculationSettings,
    mode: t.mode === "manual" ? "manual" : "automatic",
    iterative: false,
    maxIterations:
      t.maxIterations !== null && t.maxIterations <= 1e3
        ? t.maxIterations
        : defaultCalculationSettings.maxIterations,
    maxChange:
      t.maxChange !== null && t.maxChange <= 1
        ? t.maxChange
        : defaultCalculationSettings.maxChange,
  }),
    t.fullPrecision === false &&
      n.push(
        "The source used precision as displayed. This application calculates with full stored precision and never rounds cells automatically. Source rounding may already have changed stored values irreversibly; review formula results and use explicit ROUND where required. The original setting is retained in imported calculation evidence.",
      ),
    t.mode === "manual" &&
      n.push(
        "Source Manual calculation mode is retained. Formulas remain #PENDING! until Calculate Now; imported caches are historical evidence only.",
      ),
    t.mode === "autoNoTable" &&
      n.push(
        "Excel Automatic except data tables is imported as Automatic. Native Excel data-table calculation is not supported; source mode is retained as evidence.",
      ),
    t.iterative &&
      n.push(
        "Source iterative calculation was enabled. It is disabled locally until you explicitly enable it in Calculation Options and review numeric convergence.",
      ),
    ((t.maxIterations ?? 100) > 1e3 || (t.maxChange ?? 0.001) > 1) &&
      n.push(
        "Source iteration limits exceed local limits. Local defaults are used; original settings remain in imported cache evidence.",
      ));
}
var i8 = Bi(Mo(), 1);

function n8(e) {
  try {
    let t = new TextDecoder("utf-8", { fatal: true }).decode(e);
    if (t.includes("\0")) throw Error();
    return t;
  } catch {
    throw Error(
      "Automation source must be UTF-8 text without null characters. Export a UTF-8 source copy; binary VBA is inventory-only.",
    );
  }
}
function pq(e) {
  let t = vg(e, { allowBinaryWorkbookInventory: true }),
    n = new Map(),
    i = 0,
    r,
    o = new Set(),
    a = new gc((s) => {
      if (!t.has(s.name) || o.has(s.name)) {
        r = Error("Automation archive directory mismatch.");
        return;
      }
      o.add(s.name);
      let l = [],
        c = 0,
        u = -1;
      ((s.ondata = (d, f, p) => {
        if (!r) {
          if (d) {
            r = Error(`Corrupt automation archive: ${s.name}.`);
            return;
          }
          if (
            ((c += f.length),
            (i += f.length),
            c > t.get(s.name).size || i > 50 * 1024 * 1024)
          ) {
            ((r = Error("Automation archive exceeds its expansion budget.")),
              s.terminate());
            return;
          }
          if (((u = Tc(u, f)), l.push(f), p)) {
            if (
              c !== t.get(s.name).size ||
              (u ^ -1) >>> 0 !== t.get(s.name).crc
            ) {
              r = Error(`Automation archive checksum mismatch: ${s.name}.`);
              return;
            }
            let h = new Uint8Array(c),
              m = 0;
            for (let g of l) (h.set(g, m), (m += g.length));
            n.set(s.name, h);
          }
        }
      }),
        s.start());
    });
  a.register(mc);
  try {
    a.push(e, true);
  } catch {
    throw Error("Unable to inspect corrupt automation archive.");
  }
  if (r) throw r;
  if (n.size !== t.size) throw Error("Automation archive is incomplete.");
  return n;
}
function mq(e, t, n) {
  if (/<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(e))
    throw Error(`Unsupported XML encoding in ${t}.`);
  let i = new i8.SaxesParser({ xmlns: false }),
    r = 0,
    o = 0;
  (i.on("doctype", () => {
    throw Error("Automation inventory does not accept XML document types.");
  }),
    i.on("opentag", (a) => {
      if (++r > 128 || ++o > 1e6)
        throw Error("Automation inventory XML exceeds its bounds.");
      n(a.name.split(":").at(-1), a.attributes);
    }),
    i.on("closetag", () => r--),
    i.on("error", (a) => {
      throw Error(`Malformed automation metadata in ${t}: ${a.message}`);
    }),
    i.write(e).close());
}
function o8(e, t) {
  if (
    (requireAutomationFileName(t),
    !ArrayBuffer.isView(e) ||
      Object.prototype.toString.call(e) !== "[object Uint8Array]" ||
      !e.length ||
      e.length > automationLimits.maxFileBytes)
  )
    throw Error("Automation files must contain data and be 10 MB or smaller.");
  let n = e.slice(),
    i = {
      schemaVersion: 1,
      fileName: t,
      kind: "excel-container",
      fileBytes: n.length,
      sha256: sha256HexDigest(n),
      retainedEligible: n.length <= automationLimits.maxArtifactBytes,
      execution: "inventory-only",
      items: [],
      warnings: [
        "Imported automation never runs on open, preservation or native JSON reopen. Only a separately reviewed migration can issue bounded workbook commands.",
        "Preserved artifacts are inert original bytes in native JSON. VBA/XLM binaries are not executed or decompiled; retain originals and use an authorized source application to export source modules.",
      ],
    },
    r = t.split(".").at(-1)?.toLowerCase();
  if (["bas", "vba", "cls", "ts", "osts"].includes(r ?? "")) {
    if (n.length > automationLimits.maxSourceBytes)
      throw Error("Text automation source is limited to 256 KiB.");
    let o = n8(n),
      a = r === "ts" || r === "osts" ? "office-script" : "vba";
    if (((i.kind = r === "osts" ? "office-script-container" : a), r === "osts"))
      try {
        let s = JSON.parse(o);
        if (
          !s ||
          typeof s != "object" ||
          Array.isArray(s) ||
          typeof s.body != "string"
        )
          throw Error();
        if (
          ((o = s.body),
          o.includes("\0") ||
            new TextEncoder().encode(o).length >
              automationLimits.maxSourceBytes)
        )
          throw Error("Source body is outside the supported text budget.");
      } catch {
        (i.warnings.push(
          "Unsupported .osts envelope. Preserve the original or export its code as UTF-8 .ts. This adapter recognizes only a JSON object with a string body; it makes no universal .osts schema claim.",
        ),
          i.items.push({
            path: t,
            kind: "Office Script container",
            bytes: n.length,
            sha256: i.sha256,
            execution: "unsupported",
            note: "Unknown source envelope; no code inferred.",
          }));
      }
    i.items.length ||
      ((i.script = { language: a, text: o }),
      (i.execution = "reviewed-migration-only"),
      i.items.push({
        path: t,
        kind: a === "vba" ? "VBA source" : "Office Script source",
        bytes: n.length,
        sha256: i.sha256,
        execution: "migration-candidate",
        note: "Compatibility scan required. Only literal range operations are eligible; this is not an equivalent VBA or Office Scripts runtime.",
      }));
  } else if (n[0] === 208 && n[1] === 207 && n[2] === 17 && n[3] === 224)
    ((i.kind = "legacy-container"),
      i.items.push({
        path: t,
        kind: "OLE / encrypted or legacy container",
        bytes: n.length,
        sha256: i.sha256,
        execution: "unsupported",
        note: "VBA/XLM presence cannot be determined without a supported decoder. Export a clean OOXML workbook and UTF-8 source modules in an authorized application.",
      }));
  else if (
    ["xlsm", "xlam", "xltm", "xlsx", "xltx"].includes(r ?? "") ||
    (n[0] === 80 && n[1] === 75)
  ) {
    let o = pq(n),
      a = new Map(),
      s = (l, c) => {
        let u = "";
        (/vbaProject/i.test(c) || /vbaproject(?:signature)?\.bin$/i.test(l)
          ? (u = "VBA binary")
          : /(?:intl)?macrosheet/i.test(c) || /(?:^|\/)macrosheets\//i.test(l)
            ? (u = "XLM macro sheet")
            : /office.?script|scriptLink|scriptReference/i.test(c) ||
                /\.(?:osts|ts)$/i.test(l)
              ? (u = "Office Script artifact/reference")
              : (/activeX|oleObject|relationships\/(?:package|control)$/i.test(
                  c,
                ) ||
                  /\/(?:activeX|embeddings)\//i.test(l)) &&
                (u = "ActiveX / embedded object"),
          u &&
            a.set(l, {
              kind: u,
              note:
                u === "Office Script artifact/reference"
                  ? "Package artifact only. Import its separate .ts/.osts source for migration."
                  : "Inventory and opaque preservation only; no binary execution.",
            }));
      };
    for (let [l, c] of o)
      (s(l, ""),
        /\.(?:xml|rels)$/i.test(l) &&
          mq(n8(c), l, (u, d) => {
            if (
              (u === "Override" &&
                d.PartName &&
                s(d.PartName.replace(/^\//, ""), d.ContentType ?? ""),
              u === "Default" && d.Extension)
            )
              for (let f of o.keys())
                f.toLowerCase().endsWith("." + d.Extension.toLowerCase()) &&
                  s(f, d.ContentType ?? "");
            if (
              u === "Relationship" &&
              /vbaProject|(?:intl)?macrosheet|activeX|oleObject|office.?script|scriptLink|scriptReference|relationships\/(?:package|control)$/i.test(
                d.Type ?? "",
              )
            ) {
              if (d.TargetMode === "External") {
                i.items.push({
                  path: `${l} \u2192 ${d.Target ?? "(missing)"}`,
                  kind: "External automation reference",
                  bytes: 0,
                  execution: "external-unresolved",
                  note: "Referenced content was not fetched, authenticated or executed.",
                });
                return;
              }
              let f =
                  l === "_rels/.rels"
                    ? ""
                    : l.slice(0, l.lastIndexOf("/_rels/")) + "/",
                p = [];
              for (let h of ((d.Target ?? "").startsWith("/")
                ? d.Target.slice(1)
                : f + (d.Target ?? "")
              ).split("/"))
                h === ".." ? p.pop() : h && h !== "." && p.push(h);
              s(p.join("/"), d.Type ?? "");
            }
          }));
    for (let [l, c] of a) {
      let u = o.get(l);
      i.items.push({
        path: l,
        kind: c.kind,
        bytes: u?.length ?? 0,
        ...(u ? { sha256: sha256HexDigest(u) } : {}),
        execution: "unsupported",
        note: u
          ? c.note
          : "Referenced part is missing. Inventory is incomplete; repair the original.",
      });
    }
    i.items.length ||
      i.warnings.push(
        "No VBA/XLM or known automation parts were detected in this package. This does not prove the absence of separately stored scripts or automation.",
      );
  } else
    throw Error(
      "Choose VBA .bas/.vba/.cls, Office Script .ts/.osts, or an Excel container for inventory.",
    );
  return (
    i.warnings.push(
      "Office Scripts may be stored separately in OneDrive/SharePoint. Workbook inventory cannot discover those scripts; import the authorized source separately. No linked content is fetched.",
    ),
    i.retainedEligible ||
      i.warnings.push(
        "This file exceeds the 2 MB retained-artifact limit. The inventory is available; preserve the original externally.",
      ),
    r8.set(i, { signature: JSON.stringify(i), bytes: n }),
    i
  );
}
function s8(e) {
  return (e.embeddedObjectInventories ?? []).flatMap((t) =>
    t.items.map((n) => ({
      sourceFile: t.fileName,
      sourceSha256: t.sourceSha256,
      path: n.path,
      kind: n.kind,
      bytes: n.bytes,
      ...(n.sha256 ? { sha256: n.sha256 } : {}),
      disposition: "dropped",
      execution: "never",
      reason:
        "Inventory retained in native JSON only; unsupported OLE/ActiveX payload was not imported and is not emitted in derived XLSX.",
    })),
  );
}
function _k(e, t, n, i, r) {
  return {
    schemaVersion: 1,
    direction: e,
    fileName: t,
    fileBytes: n,
    workbookName: i.name,
    revision: i.revision,
    compatibility: "derived-subset",
    lossless: false,
    originalFileRetained: !!i.xlsxSource,
    ...(i.alternateSource
      ? {
          alternateOriginal: {
            fileName: i.alternateSource.fileName,
            byteLength: i.alternateSource.byteLength,
            sha256: i.alternateSource.sha256,
            format: i.alternateSource.format,
          },
        }
      : {}),
    ...(i.xlsxSource
      ? {
          originalSource: {
            fileName: i.xlsxSource.fileName,
            byteLength: i.xlsxSource.byteLength,
            sha256: i.xlsxSource.sha256,
          },
        }
      : {}),
    sheets: i.sheets.map((o) => ({
      name: o.name,
      storedCells: Object.keys(o.cells).length,
      hidden: !!(o.hidden || o.veryHidden),
    })),
    warnings: [...new Set(r)],
    ...(e === "export" && i.embeddedObjectInventories?.length
      ? { embeddedObjectLosses: s8(i) }
      : {}),
    caches: {
      total: 0,
      pending: 0,
      matched: 0,
      different: 0,
      missing: 0,
      neutralized: 0,
      written: 0,
      omitted: 0,
      issues: [],
      matches: [],
      detailsOmitted: 0,
    },
  };
}
function gq(e, t) {
  if (e != null) {
    if (e instanceof Date) return { type: "number", value: _p(e, t) };
    if (typeof e == "object" && "error" in e && typeof e.error == "string")
      return { type: "error", value: e.error };
    if (typeof e == "string") return { type: "text", value: e };
    if (typeof e == "boolean") return { type: "boolean", value: e };
    if (typeof e == "number" && Number.isFinite(e))
      return { type: "number", value: e };
  }
}
function bq(e, t) {
  return (
    e.type === t.type &&
    (typeof e.value == "number" && typeof t.value == "number"
      ? Math.abs(e.value - t.value) <=
        1e-10 * Math.max(1, Math.abs(e.value), Math.abs(t.value))
      : e.value === t.value)
  );
}
function a8(e, t, n, i = new Map(), r, o = new Map()) {
  let a = calculateWorkbookWithSettings(t);
  for (let s of t.sheets) {
    let l = e.getWorksheet(s.name);
    for (let [c, u] of Object.entries(s.cells)) {
      let d = l.getCell(c).value;
      if (
        !d ||
        typeof d != "object" ||
        !("formula" in d || "sharedFormula" in d)
      )
        continue;
      let f = i.get(s.name)?.has(c)
          ? { type: "text", value: "" }
          : o.get(s.name)?.has(c)
            ? { type: "number", value: Number(o.get(s.name).get(c)) }
            : gq(l.getCell(c).result, t.datePolicy?.dateSystem ?? "1900"),
        p = a.getTypedValue(s.id, c);
      r &&
        r.entries.push({
          sheetId: s.id,
          address: c,
          formula: String(u.raw),
          ...(f ? { cache: f } : {}),
        });
      let h = u.literal
        ? "neutralized"
        : p.type === "error" && p.value === "#PENDING!"
          ? "pending"
          : f
            ? bq(f, p)
              ? "matched"
              : "different"
            : "missing";
      (n.caches.total++, (n.caches[h] = (n.caches[h] ?? 0) + 1));
      let m = h === "matched" ? n.caches.matches : n.caches.issues;
      m.length < 100
        ? m.push({
            sheet: s.name,
            address: c,
            ...(f ? { source: f } : {}),
            calculated: p,
            status: h,
          })
        : n.caches.detailsOmitted++;
    }
  }
  (n.caches.pending &&
    n.warnings.push(
      `${n.caches.pending} formula cache comparison(s) are pending. Use Calculate Now and download calculation evidence for a current comparison.`,
    ),
    n.caches.different &&
      n.warnings.push(
        `${n.caches.different} source formula cache(s) differ from local recalculation. Review the comparison before relying on results; source caches never override local results.`,
      ),
    n.caches.missing &&
      n.warnings.push(
        `${n.caches.missing} formula(s) have no usable source cache. Local results are not independently verified Excel results.`,
      ),
    n.caches.neutralized &&
      n.warnings.push(
        `${n.caches.neutralized} external/action formula(s) were converted to literal text.`,
      ));
}
function l8(e, t) {
  let n = (r, o) =>
      new RegExp(`(?:^|\\s)${o}\\s*=\\s*(["'])(.*?)\\1`, "i")
        .exec(r)?.[2]
        ?.replace(/&#(x[0-9a-f]+|[0-9]+);/gi, (s, l) =>
          String.fromCodePoint(
            l[0].toLowerCase() === "x" ? parseInt(l.slice(1), 16) : Number(l),
          ),
        ),
    i;
  for (let r of e.matchAll(
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)(?:\w+:)?(c|f)\b([^>]*)>/gi,
  )) {
    if (!r[2]) continue;
    if (r[2].toLowerCase() === "c") {
      i = r[1] || /\/\s*$/.test(r[3]) ? undefined : r[3];
      continue;
    }
    if (i === undefined || r[1] || n(r[3], "t") !== "array") continue;
    let o = n(i, "r") ?? "unknown address",
      a = n(r[3], "ref") ?? o,
      s =
        n(i, "cm") !== undefined
          ? "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)"
          : "Legacy CSE array formula";
    throw new Error(
      `${s} at ${t}!${o}, range ${a}. Import stopped before changing the workbook. Open the original in Excel and convert the array to ordinary supported formulas or an explicitly reviewed values snapshot, then import that copy. Fixed CSE ranges and Excel array metadata are not silently treated as dynamic spills.`,
    );
  }
}
var y8 = Bi(u8(), 1);
var p8 = Bi(d8(), 1);

var wq = 2958465,
  Ak = class extends Error {};
function xq() {
  throw new Ak("Date series must stay between 1899-12-31 and 9999-12-31.");
}
function kq(e, t, n) {
  let i = new Date(0);
  return (i.setUTCHours(0, 0, 0, 0), i.setUTCFullYear(e, t, n), i);
}
function _q(e) {
  let t = (e.getTime() - vq) / yq,
    n = t >= 60 ? t + 1 : t;
  return ((!Number.isSafeInteger(n) || n < 0 || n > wq) && xq(), n);
}
function Ik(e, t = "1900") {
  if (typeof e == "number")
    return (
      (!Number.isSafeInteger(e) || e < 0 || e > dateSystemMaxSerial(t)) &&
        Tk(
          `Date values must be whole Excel ${t} serials between 0 and ${dateSystemMaxSerial(t)}.`,
        ),
      e
    );
  if (
    ((typeof e != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(e)) &&
      Tk("Enter a date as YYYY-MM-DD or a whole Excel date serial."),
    e === "1900-02-29" && t === "1900")
  )
    return 60;
  let [n, i, r] = e.split("-").map(Number),
    o = kq(n, i - 1, r);
  return (
    (o.getUTCFullYear() !== n ||
      o.getUTCMonth() !== i - 1 ||
      o.getUTCDate() !== r) &&
      Tk("The series date is not a valid calendar date."),
    Ik(_q(o) - dateSystemSerialOffset(t), t)
  );
}
var f8 = {
  delimiter: "auto",
  encoding: "auto",
  quote: '"',
  escape: "double",
  newline: "auto",
  hasHeader: true,
  decimal: ".",
  grouping: "none",
  dateOrder: "ymd",
  empty: "blank",
  trim: false,
  columns: {},
};
var Eq = ["auto", "text", "number", "integer", "boolean", "date"],
  m8 = Array.from({ length: 256 }, (e, t) => String.fromCharCode(t));
[
  8364, 129, 8218, 402, 8222, 8230, 8224, 8225, 710, 8240, 352, 8249, 338, 141,
  381, 143, 144, 8216, 8217, 8220, 8221, 8226, 8211, 8212, 732, 8482, 353, 8250,
  339, 157, 382, 376,
].forEach((e, t) => (m8[128 + t] = String.fromCharCode(e)));
var ao = (e, t, n, i, r) => {
    (e.issueCount++,
      e.issues.length < 100 &&
        e.issues.push({
          record: t,
          column: n,
          message: i,
          ...(r !== undefined ? { value: r.slice(0, 200) } : {}),
        }));
  },
  Sq = (e, t, n, i, r, o, a) => {
    (e.changeCount++,
      e.changes.push({
        record: t,
        column: n,
        address: i,
        reason: r,
        before: o,
        after: a,
      }));
  };
function Cq(e = {}) {
  let t = { ...f8, ...e, columns: { ...(e.columns ?? {}) } };
  for (let [n, i] of Object.entries({
    delimiter: ["auto", ",", ";", "	", "|"],
    encoding: ["auto", "utf-8", "utf-16le", "utf-16be", "windows-1252"],
    quote: ['"', "'", "none"],
    escape: ["double", "backslash"],
    newline: [
      "auto",
      `\r
`,
      `
`,
      "\r",
    ],
    decimal: [".", ","],
    grouping: ["none", ".", ",", " "],
    dateOrder: ["ymd", "dmy", "mdy"],
    empty: ["blank", "text"],
  }))
    if (!i.includes(String(t[n])))
      throw Error(`Unsupported delimited ${n} setting.`);
  if (typeof t.hasHeader != "boolean" || typeof t.trim != "boolean")
    throw Error("Header and whitespace controls must be boolean.");
  if (t.decimal === t.grouping)
    throw Error("Decimal and grouping separators must differ.");
  for (let [n, i] of Object.entries(t.columns))
    if (
      !/^(0|[1-9]\d*)$/.test(n) ||
      Number(n) >= workbookLimits.maxCols ||
      !Eq.includes(i)
    )
      throw Error("Invalid column type selection.");
  return t;
}
function Tq(e, t, n, i) {
  return {
    schemaVersion: 1,
    direction: e,
    fileName: t,
    fileBytes: n,
    options: i,
    records: 0,
    columns: 0,
    storedCells: 0,
    blankRecords: 0,
    schema: [],
    preview: [],
    issueCount: 0,
    issues: [],
    changeCount: 0,
    changes: [],
    safeTextChanges: 0,
    complete: true,
    valid: false,
    warnings: [],
    safeExportPolicy: "apostrophe-prefix-formula-like-text",
  };
}
function Aq(e, t) {
  if (e.length > workbookLimits.maxFileBytes)
    throw Error("Delimited files must be 10 MB or smaller.");
  let n =
    e[0] === 239 && e[1] === 187 && e[2] === 191
      ? "utf-8"
      : e[0] === 255 && e[1] === 254
        ? "utf-16le"
        : e[0] === 254 && e[1] === 255
          ? "utf-16be"
          : undefined;
  if (n && t !== "auto" && t !== n)
    throw Error(`File BOM indicates ${n}; select that encoding or Automatic.`);
  let i = t === "auto" ? (n ?? "utf-8") : t,
    r;
  try {
    r = (
      i === "windows-1252"
        ? Array.from(e, (o) => m8[o]).join("")
        : new TextDecoder(i, { fatal: true, ignoreBOM: true }).decode(e)
    ).replace(/^\uFEFF/, "");
  } catch {
    throw Error(
      `Invalid ${i} byte sequence. Choose the source encoding; bytes are never silently replaced.`,
    );
  }
  if (r.includes("\0"))
    throw Error(
      "Delimited files cannot contain null bytes; check the encoding.",
    );
  return { text: r, encoding: i };
}
function Iq(e, t, n, i) {
  let r = [],
    o = "",
    a = "start",
    s = 1,
    l = 0,
    c = false,
    u = () => {
      (r.push(o),
        (o = ""),
        (a = "start"),
        r.length > workbookLimits.maxCols &&
          (ao(
            n,
            s,
            undefined,
            `Delimited data exceeds ${workbookLimits.maxCols} columns.`,
          ),
          (c = true)));
    },
    d = (f) => {
      if ((u(), !c)) {
        if (s > workbookLimits.maxRows) {
          (ao(
            n,
            s,
            undefined,
            `Delimited data exceeds ${workbookLimits.maxRows} rows.`,
          ),
            (c = true));
          return;
        }
        ((c = !i(r, s, f === l)), (r = []), s++);
      }
    };
  for (let f = 0; f < e.length && !c; f++) {
    let p = e[f];
    if (a === "quoted")
      t.escape === "backslash" &&
      p === "\\" &&
      (e[f + 1] === t.quote || e[f + 1] === "\\")
        ? (o += e[++f])
        : p === t.quote
          ? t.escape === "double" && e[f + 1] === t.quote
            ? ((o += p), f++)
            : (a = "closed")
          : (o += p);
    else if (p === t.delimiter) u();
    else if (
      p === "\r" ||
      p ===
        `
`
    ) {
      let h =
        p === "\r" &&
        e[f + 1] ===
          `
`
          ? `\r
`
          : p;
      (t.newline !== "auto" &&
        t.newline !== h &&
        ao(
          n,
          s,
          undefined,
          "Record newline does not match the selected dialect.",
        ),
        d(f),
        h.length === 2 && f++,
        (l = f + 1));
    } else
      a === "start" && t.quote !== "none" && p === t.quote
        ? (a = "quoted")
        : (a === "closed"
            ? (ao(n, s, r.length + 1, "Unexpected text after a closing quote."),
              (a = "plain"))
            : t.quote !== "none" &&
              p === t.quote &&
              ao(
                n,
                s,
                r.length + 1,
                "Quote inside an unquoted field; quote the entire field.",
              ),
          (o += p),
          (a = "plain"));
    o.length > fileImportLimits.maxCellTextLength &&
      (ao(
        n,
        s,
        r.length + 1,
        "A field exceeds the 32,767-character cell limit.",
      ),
      (c = true));
  }
  (!c &&
    a === "quoted" &&
    ao(n, s, r.length + 1, "Quoted field is not closed."),
    !c && l < e.length && d(e.length),
    (n.complete = !c));
}
var Rk = (e) => e.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
function h8(e, t) {
  let n = t.grouping === "none" ? "" : t.grouping,
    i = n ? `(?:\\d+|\\d{1,3}(?:${Rk(n)}\\d{3})+)` : "\\d+";
  if (
    !new RegExp(
      `^[+-]?(?:${i}(?:${Rk(t.decimal)}\\d*)?|${Rk(t.decimal)}\\d+)(?:[eE][+-]?\\d+)?$`,
    ).test(e)
  )
    throw Error("Invalid number for the selected decimal/grouping separators.");
  let o = (n ? e.split(n).join("") : e).replace(t.decimal, "."),
    a = o
      .split(/[eE]/)[0]
      .replace(/[^0-9]/g, "")
      .replace(/^0+/, "")
      .replace(/0+$/, "").length,
    s = Number(o);
  if (
    !Number.isFinite(s) ||
    a > 15 ||
    (s === 0 && /[1-9]/.test(o.split(/[eE]/)[0]))
  )
    throw Error(
      "Number is outside supported precision/range; select Text to retain every digit.",
    );
  return s;
}
function Rq(e, t, n = "1900") {
  let i =
    /^(\d{1,4})([-/])(\d{1,2})\2(\d{1,4})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/.exec(
      e,
    );
  if (!i)
    throw Error(
      "Date must match the selected year/month/day order, optionally HH:mm:ss.",
    );
  let r = [i[1], i[3], i[4]],
    [o, a, s] =
      t === "ymd" ? r : t === "dmy" ? [r[2], r[1], r[0]] : [r[2], r[0], r[1]];
  if (o.length !== 4) throw Error("Dates require a four-digit year.");
  let l = Number(i[5] ?? 0),
    c = Number(i[6] ?? 0),
    u = Number(i[7] ?? 0);
  if (l > 23 || c > 59 || u > 59) throw Error("Invalid date/time component.");
  return (
    Ik(`${o}-${a.padStart(2, "0")}-${s.padStart(2, "0")}`, n) +
    (l * 3600 + c * 60 + u) / 86400
  );
}
function Nq(e, t, n, i = "1900") {
  if (e === "") return { raw: n.empty === "text" ? "" : null };
  let r = n.trim ? e.trim() : e;
  if (t === "auto") {
    if (
      /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(r) &&
      r.split(/[eE]/)[0].replace(/[^0-9]/g, "").length <= 15 &&
      n.decimal === "." &&
      n.grouping === "none"
    )
      try {
        return { raw: h8(r, n) };
      } catch {}
    return { raw: r, ...(r.startsWith("=") ? { literal: true } : {}) };
  }
  if (t === "text")
    return { raw: r, ...(r.startsWith("=") ? { literal: true } : {}) };
  if (t === "date")
    return {
      raw: Rq(r, n.dateOrder, i),
      format: "date",
      numFmt: /[ T]\d\d:/.test(r) ? "yyyy-mm-dd hh:mm:ss" : "yyyy-mm-dd",
    };
  if (t === "boolean") {
    if (/^(true|false)$/i.test(r)) return { raw: r.toLowerCase() === "true" };
    throw Error("Boolean values must be TRUE or FALSE (case-insensitive).");
  }
  let o = h8(r, n);
  if (t === "integer" && !Number.isSafeInteger(o))
    throw Error("Integer column requires a safe whole number.");
  return { raw: o };
}
function g8(e, t, n, i) {
  let r = Cq(n),
    o = Aq(e, r.encoding);
  ((r.encoding = o.encoding),
    r.delimiter === "auto" &&
      (r.delimiter = /\.tsv$/i.test(t)
        ? "	"
        : p8.default.parse(o.text.slice(0, 65536), {
            preview: 10,
            quoteChar: r.quote === "none" ? "" : r.quote,
            escapeChar: r.escape === "double" ? r.quote : "\\",
            delimitersToGuess: [",", ";", "	", "|"],
          }).meta.delimiter || ","));
  let a = Tq("import", t, e.length, r),
    s = {};
  (a.warnings.push(
    "Delimited data contains values only. All imported formula-like content stays literal text and is never treated as a formula. Auto keeps leading-zero/long identifiers and dates as text; select explicit types for localized numbers, dates and booleans.",
  ),
    (n.encoding === undefined || n.encoding === "auto") &&
      a.warnings.push(
        `Encoding: ${r.encoding}; automatic mode uses a BOM when present and UTF-8 otherwise. Legacy encodings require an explicit selection.`,
      ),
    (n.delimiter === undefined || n.delimiter === "auto") &&
      a.warnings.push(
        `Detected delimiter: ${r.delimiter === "	" ? "tab" : r.delimiter}. Review the schema and column count before importing.`,
      ));
  let l = 0;
  if (
    (Iq(o.text, r, a, (c, u, d) => {
      ((a.records = u),
        c.every((f) => f === "") && a.blankRecords++,
        !l && !d && (l = c.length),
        (a.columns = Math.max(a.columns, c.length)),
        !d &&
          l &&
          c.length !== l &&
          ao(
            a,
            u,
            undefined,
            `Malformed row: expected ${l} fields, found ${c.length}.`,
          ));
      for (let f = 0; f < c.length; f++) {
        let p = c[f],
          h = r.hasHeader && u === 1 ? "text" : (r.columns[f] ?? "auto"),
          m = `${columnIndexToLetters(f)}${u}`;
        a.schema[f] ||= {
          index: f,
          label: r.hasHeader && u === 1 ? p : columnIndexToLetters(f),
          type: r.columns[f] ?? "auto",
          observed: [],
        };
        let g;
        try {
          g = Nq(p, h, r);
        } catch (b) {
          (ao(
            a,
            u,
            f + 1,
            b instanceof Error ? b.message : "Type conversion failed.",
            p,
          ),
            (g = { raw: p, ...(p.startsWith("=") ? { literal: true } : {}) }));
        }
        if (!(r.hasHeader && u === 1)) {
          let b =
            g.raw === null
              ? "blank"
              : g.format === "date"
                ? "date"
                : typeof g.raw;
          a.schema[f].observed.includes(b) || a.schema[f].observed.push(b);
        }
        if (g.raw !== null && !(d && r.empty === "blank")) {
          if (++a.storedCells > workbookLimits.maxCells)
            return (
              ao(
                a,
                u,
                f + 1,
                "Delimited data exceeds 100,000 populated cells.",
              ),
              false
            );
          i && (s[m] = g);
        }
        (p !== "" &&
          (g.raw !== p || (r.trim && p !== p.trim())) &&
          Sq(
            a,
            u,
            f + 1,
            m,
            g.format === "date"
              ? "Date converted to native serial"
              : "Selected type or whitespace conversion",
            p,
            String(g.raw ?? ""),
          ),
          a.preview.length < 25 &&
            f === 0 &&
            a.preview.push({ record: u, raw: c.slice(0, 200), values: [] }),
          u <= 25 && f < 200 && a.preview.at(-1)?.values.push(g.raw));
      }
      return true;
    }),
    a.records ||
      ao(
        a,
        0,
        undefined,
        "The file has no records. Select a populated delimited file.",
      ),
    r.hasHeader && a.records)
  ) {
    let c = a.schema.map((d) => d.label),
      u = new Set();
    c.some(
      (d) =>
        !d.trim() ||
        (u.has(d.toLowerCase()) ? true : (u.add(d.toLowerCase()), false)),
    ) &&
      a.warnings.push(
        "Blank or duplicate header labels are retained exactly. Column letters identify them; no fields are overwritten or renamed.",
      );
  }
  return ((a.valid = a.complete && a.issueCount === 0), { r: a, cells: s });
}
function b8(e, t, n = {}) {
  return g8(e, t, n, false).r;
}
function Eg(e, t, n = {}) {
  let { r, cells } = g8(e, t, n, true);
  if (!r.valid)
    throw Error(
      `Unable to read delimited data: ${r.issues[0]?.message ?? "Incomplete preview."} Review the import schema; no rows were imported.`,
    );
  let o = Q6(
      t
        .replace(/\.[^.]+$/, " ")
        .trim()
        .slice(0, 200) || "Imported workbook",
    ),
    a = o.sheets[0];
  return (
    (a.name = "Imported data"),
    (a.cells = cells),
    (a.frozenRows = r.options.hasHeader ? 1 : 0),
    (a.delimitedSource = {
      records: r.records,
      columns: r.columns,
      hasHeader: r.options.hasHeader,
    }),
    {
      workbook: gi(
        Wm(o, { id: crypto.randomUUID(), name: t.slice(0, 255), kind: "file" }),
      ),
      warnings: r.warnings,
      delimitedReport: r,
    }
  );
}
var Oq = fileImportLimits.maxXlsxExpandedBytes,
  Mq =
    "XLSX exchange supports cell values, formulas, common formats, supported fonts/strike/alignment/indent/per-edge borders, notes, hidden sheets, column widths, explicit row heights, worksheet default dimensions, bounded hidden rows/columns, frozen rows and columns, basic tables, safe merges, supported validation/conditional-format rules, PNG/JPEG pictures, repeated print titles, full-axis manual print breaks and page setup. XLSX processes at most 100,000 hidden rows per workbook; native JSON keeps larger visibility intervals compactly. Workbook and worksheet-scoped defined names, constant/formula definitions and supported name comments are preserved. Unsupported name attributes are reported; unsafe or malformed definitions stop import/export. Enhanced number formats export supported masks; locale, scaling labels, duration/date-system and accounting alignment fidelity may differ, and enhanced format metadata remains native JSON only. Reusable calculated-column rules are native JSON only; XLSX imports retain supported cell formulas but do not infer rules for future appends. Editable outline groups, threaded discussions and interactive hyperlink metadata remain native-JSON features; XLSX retains only collapsed outline visibility. Native pivot/query refresh definitions, charts/shapes, filters, split panes, default-hidden-row settings, password protection and other unsupported Excel metadata are not preserved. External or executable formulas are converted to literal text on import and export. Keep the original workbook; use native JSON for a lossless application backup.",
  Nk = () => globalThis.crypto.randomUUID();
var $q = (e, t) => ({
    schemaVersion: 1,
    id: Nk(),
    name: e.replace(/\.[^.]+$/, "").slice(0, 200) || "Imported workbook",
    revision: 0,
    metadata: createWorkbookMetadata(),
    sheets: t,
    events: [],
  }),
  Pq = (e) =>
    /[\[\]|]/.test(
      stripFormulaStringLiterals(e).replace(/"(?:[^"]|"")*"/g, '""'),
    ) ||
    /(?:https?:|file:|javascript:|data:|\\\\)/i.test(e) ||
    /\b(?:_xlfn\.)?(?:WEBSERVICE|HYPERLINK|IMAGE|RTD|DDE|CALL|EXEC|REGISTER(?:\.ID)?|RUN|SQL\.REQUEST|STOCKHISTORY)\s*\(/i.test(
      e,
    );
function Wn(e, t) {
  return new RegExp("(?:^|\\s)" + t + `\\s*=\\s*(["'])([\\s\\S]*?)\\1`).exec(
    e,
  )?.[2];
}
function Dq(e) {
  let t = [];
  for (let n of (e.startsWith("/") ? e.slice(1) : "xl/" + e).split("/"))
    n === ".." ? t.pop() : n !== "." && t.push(n);
  return t.join("/");
}
var Rc = (e) =>
  e.replace(/&(?:amp|lt|gt|quot|apos|#\d+|#x[a-f\d]+);/gi, (t) =>
    t === "&amp;"
      ? "&"
      : t === "&lt;"
        ? "<"
        : t === "&gt;"
          ? ">"
          : t === "&quot;"
            ? '"'
            : t === "&apos;"
              ? "'"
              : String.fromCodePoint(
                  Number.parseInt(
                    t.slice(t[2]?.toLowerCase() === "x" ? 3 : 2, -1),
                    t[2]?.toLowerCase() === "x" ? 16 : 10,
                  ),
                ),
  );
function v8(e, t) {
  let n = vg(e),
    i = new Set(n.keys()),
    r = n.size,
    o = ON();
  if (
    [...i].some((_) =>
      /(?:vbaProject\.bin|xl\/macrosheets\/|xl\/activeX\/|xl\/embeddings\/)/i.test(
        _,
      ),
    )
  )
    throw new Error(
      "Workbooks containing macros, ActiveX or embedded objects are not accepted. Save a clean XLSX copy first.",
    );
  let a = 0,
    s,
    l = 0,
    c = 0,
    u = 0,
    d = 0,
    f = 0,
    p = 0,
    h = 0,
    m = false,
    g = new Set(),
    b = new gc((_) => {
      if ((l++, !i.has(_.name) || g.has(_.name))) {
        s = new Error("Archive entries do not match its directory.");
        return;
      }
      g.add(_.name);
      let S = 0,
        O = /\.(?:xml|rels|vml)$/i.test(_.name),
        E = O ? new TextDecoder("utf-8", { fatal: true }) : undefined,
        R = -1,
        I = "";
      ((_.ondata = (U, P, T) => {
        if (U) {
          s = new Error(
            `Corrupt XLSX compressed content in ${_.name}. Repair the original workbook and save a clean XLSX copy.`,
          );
          return;
        }
        if (((a += P.length), (S += P.length), a > Oq)) {
          (_.terminate(),
            (s = new Error("The XLSX archive expands beyond 50 MB.")));
          return;
        }
        if (S > n.get(_.name).size) {
          (_.terminate(),
            (s = new Error(
              "The XLSX archive expansion does not match its directory.",
            )));
          return;
        }
        R = Tc(R, P);
        try {
          E && (I += E.decode(P, { stream: !T }));
        } catch {
          s = new Error(
            `Unsupported XML encoding in ${_.name}. Save a standard UTF-8 XLSX copy.`,
          );
          return;
        }
        if (T) {
          if (S !== n.get(_.name).size) {
            s = new Error(
              "The XLSX archive expansion does not match its directory.",
            );
            return;
          }
          if ((R ^ -1) >>> 0 !== n.get(_.name).crc) {
            s = new Error(
              `XLSX checksum mismatch in ${_.name}. The archive is corrupt; repair the original and save a clean XLSX copy.`,
            );
            return;
          }
          try {
            if ((O && MN(I, _.name, n, o), /^xl\/styles\.xml$/i.test(_.name)))
              for (let A of I.matchAll(/<(?:\w+:)?sz\b([^>]*)>/g)) {
                let L = Number(Wn(A[1], "val"));
                (!Number.isInteger(L) || L < 8 || L > 72) && (m = true);
              }
            if (/^xl\/worksheets\//i.test(_.name)) {
              (l8(I, _.name),
                (h = RN(I, _.name, h)),
                /^xl\/worksheets\/[^/]+\.xml$/i.test(_.name) &&
                  t.numericTokens.set(_.name, XN(I)));
              for (let v of I.matchAll(/<(?:\w+:)?pageSetup\b([^>]*)>/gi)) {
                let k = Wn(v[1], "pageOrder");
                if (
                  k !== undefined &&
                  !["overThenDown", "downThenOver"].includes(Rc(k))
                )
                  throw new Error(
                    "Unsupported XLSX page order. Import stopped to preserve the source layout.",
                  );
              }
              let A = {};
              for (let [v, k, x, C] of [
                [
                  "rowBreaks",
                  "rowBreaks",
                  workbookLimits.maxRows,
                  workbookLimits.maxCols,
                ],
                [
                  "colBreaks",
                  "columnBreaks",
                  workbookLimits.maxCols,
                  workbookLimits.maxRows,
                ],
              ]) {
                let N = [
                  ...I.matchAll(
                    new RegExp(
                      "<(?:\\w+:)?" +
                        v +
                        "\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?" +
                        v +
                        ">",
                      "gi",
                    ),
                  ),
                ];
                if (N.length > 1)
                  throw Error(
                    "Duplicate XLSX page-break sections are not supported.",
                  );
                let F = [],
                  z = 0;
                for (let q of (N[0]?.[1] ?? "").matchAll(
                  /<(?:\w+:)?brk\b([^>]*)\/?\s*>/gi,
                )) {
                  if (++z > 100)
                    throw Error(
                      "XLSX page breaks exceed 100 boundaries per axis.",
                    );
                  let D = Wn(q[1], "id"),
                    W = Wn(q[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(D ?? "") ||
                    !Number.isSafeInteger(Number(D)) ||
                    Number(D) >= x
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (W !== "1" && W !== "true") continue;
                  let $ = Number(Wn(q[1], "min") ?? 0),
                    M = Number(Wn(q[1], "max") ?? C - 1);
                  if ($ !== 0 || !Number.isSafeInteger(M) || M < C - 1)
                    throw Error(
                      "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
                    );
                  F.push(Number(D));
                }
                F.sort((q, D) => q - D);
                let j = validatePrintPageBreaks(F, x);
                j && (A[k] = j);
              }
              (t.printBreaks.set(_.name, A),
                t.conditionalStops.set(_.name, JN(I)),
                t.validationAreas.set(_.name, ZN(I)));
              let L = [],
                y = [];
              for (let v of I.matchAll(/<(?:\w+:)?col\b([^>]*)>/gi)) {
                let k = Wn(v[1], "min"),
                  x = Wn(v[1], "max"),
                  C = Wn(v[1], "width"),
                  N = Number(k) - 1,
                  F = Number(x) - 1;
                if (
                  !/^[1-9]\d*$/.test(k ?? "") ||
                  !/^[1-9]\d*$/.test(x ?? "") ||
                  !Number.isSafeInteger(N) ||
                  !Number.isSafeInteger(F) ||
                  N < 0 ||
                  F < N ||
                  F >= workbookLimits.maxCols ||
                  y.length >= workbookLimits.maxCols
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
                let z = C === undefined ? undefined : Number(C);
                if (
                  z !== undefined &&
                  (!C?.trim() || !Number.isFinite(z) || z < 0)
                )
                  throw Error(
                    "XLSX column widths must be finite nonnegative values.",
                  );
                y.push({
                  start: N,
                  end: F,
                  ...(z === undefined ? {} : { width: z }),
                });
              }
              y.sort((v, k) => v.start - k.start);
              for (let v = 1; v < y.length; v++)
                if (y[v].start <= y[v - 1].end)
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
              if (
                (t.columns.set(_.name, y),
                /<(?:\w+:)?sheetFormatPr\b[^>]*\bzeroHeight\s*=\s*["'](?:1|true)["']/i.test(
                  I,
                ))
              )
                throw new Error(
                  "Default-hidden worksheet rows cannot be imported safely. Replace the default with explicit hidden rows before importing.",
                );
              for (let v of I.matchAll(/<(?:\w+:)?row\b([^>]*)>/gi)) {
                if (!/\bhidden\s*=\s*["'](?:1|true)["']/i.test(v[1])) continue;
                let k = Number(Wn(v[1], "r")) - 1;
                if (
                  !Number.isInteger(k) ||
                  k < 0 ||
                  k >= workbookLimits.maxRows
                )
                  throw new Error(
                    "Hidden XLSX rows must have valid row numbers.",
                  );
                let x = L.at(-1);
                if (x && k <= x.end)
                  throw new Error(
                    "Hidden XLSX row numbers must be distinct and ascending.",
                  );
                if (
                  (x && k === x.end + 1
                    ? (x.end = k)
                    : L.push({ start: k, end: k }),
                  L.length > worksheetCapacityLimits.maxIntervalsPerAxis)
                )
                  throw new Error(
                    "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
                  );
                if (++p > worksheetCapacityLimits.maxXlsxHiddenRows)
                  throw new Error(
                    "XLSX hidden rows exceed the 100,000-row processing budget. Reduce hidden rows in the source; native JSON preserves larger visibility intervals.",
                  );
              }
              L.length && t.hiddenRows.set(_.name, L);
            }
            if (_.name === "xl/workbook.xml")
              for (let A of I.matchAll(/<(?:\w+:)?sheet\b([^>]*)>/g)) {
                let L = Wn(A[1], "name"),
                  y = Wn(A[1], "r:id");
                L && y && t.sheetNames.set(Rc(L), Rc(y));
              }
            if (_.name === "xl/_rels/workbook.xml.rels")
              for (let A of I.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)) {
                let L = Wn(A[1], "Id"),
                  y = Wn(A[1], "Target");
                L &&
                  y &&
                  Wn(A[1], "TargetMode") !== "External" &&
                  t.sheetRelationships.set(Rc(L), Dq(Rc(y)));
              }
            for (let A of I.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (let L of Rc(A[3]).trim().split(/\s+/)) {
                let y = parseRangeAddress(L),
                  v =
                    (y.lastRow - y.firstRow + 1) * (y.lastCol - y.firstCol + 1);
                if (A[1] === "mergeCell") {
                  if (
                    v > 500 ||
                    ++f > 1e3 ||
                    (u += v) > workbookLimits.maxCells
                  )
                    throw new Error(
                      "Merged ranges exceed the supported import budget (500 cells per merge and 100,000 cells overall).",
                    );
                } else if ((d += v) > workbookLimits.maxCells)
                  throw new Error(
                    "XLSX validation ranges exceed the 100,000-cell import budget. Narrow the validation area in the source workbook.",
                  );
              }
            if (_.name === "xl/workbook.xml") {
              t.calculation = e8(I);
              let A = aP(I);
              ((t.names = A.names),
                (t.nameWarnings = A.warnings),
                (t.builtinNameXml = A.builtinXml));
            }
          } catch (A) {
            s =
              A instanceof Error
                ? A
                : new Error("Invalid XLSX worksheet metadata.");
            return;
          }
          c++;
        }
      }),
        _.start());
    });
  b.register(mc);
  for (let _ = 0; _ < e.length; _ += 2048) {
    if (s) throw s;
    try {
      b.push(e.subarray(_, Math.min(_ + 2048, e.length)), _ + 2048 >= e.length);
    } catch {
      throw new Error(
        "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
      );
    }
  }
  if (s) throw s;
  if (l !== r || c !== r)
    throw new Error("The XLSX archive contains incomplete entries.");
  let w = [Mq, ...t.nameWarnings, ...$N(o, i)];
  return (
    (t.emptyStringCaches = o.emptyStringCaches),
    m &&
      w.push(
        "Unsupported fractional or out-of-range font sizes were found in the source. The XLSX reader may normalize them; the editor supports integer font sizes from 8 to 72 points.",
      ),
    [...i].some((_) => /externalLinks|connections\.xml/i.test(_)) &&
      w.push(
        "External workbook links or connections were detected. They will not be refreshed or included in the derived copy; the original retains their source parts.",
      ),
    w
  );
}
function Lq(e) {
  if (!(!e || e === "General"))
    return /%/.test(e)
      ? "percent"
      : /[$€£¥]/.test(e)
        ? "currency"
        : /[ymdhis]/i.test(e.replace(/"[^"]*"/g, ""))
          ? "date"
          : "number";
}
function Ga(e) {
  return e && /^(?:[0-9a-f]{8}|[0-9a-f]{6})$/i.test(e)
    ? `#${e.slice(-6).toLowerCase()}`
    : undefined;
}
function Uq(e, t = [], n, i = false) {
  let r = (m) => {
      t.includes(m) || t.push(m);
    },
    o = e.value;
  if (o === null && !e.note && !Object.keys(e.style ?? {}).length) return;
  let a = null,
    s;
  typeof o == "string"
    ? ((a = o), a.startsWith("=") && (s = true))
    : typeof o == "number" || typeof o == "boolean"
      ? (a = o)
      : o instanceof Date
        ? (a = n !== undefined ? Number(n) : _p(o, i ? "1904" : "1900"))
        : o &&
          typeof o == "object" &&
          ("formula" in o || "sharedFormula" in o
            ? ((a = `=${e.formula}`), Pq(a) && (s = true))
            : "richText" in o
              ? (r(
                  "Rich-text runs are flattened to cell text; per-run formatting is not preserved.",
                ),
                (a = o.richText.map((m) => m.text).join("")),
                a.startsWith("=") && (s = true))
              : "hyperlink" in o
                ? ((a = o.text), a.startsWith("=") && (s = true))
                : "error" in o && (a = o.error));
  let l = {};
  (e.font?.bold && (l.bold = true),
    e.font?.italic && (l.italic = true),
    e.font?.strike && (l.strikethrough = true),
    e.alignment?.vertical &&
      ["top", "middle", "bottom"].includes(e.alignment.vertical) &&
      (l.verticalAlign = e.alignment.vertical),
    e.alignment?.indent !== undefined &&
      (Number.isInteger(e.alignment.indent) &&
      e.alignment.indent >= 0 &&
      e.alignment.indent <= 15
        ? (l.indent = e.alignment.indent)
        : r(
            "Unsupported indentation was omitted; the editor supports integer levels 0\u201315.",
          )),
    e.font?.underline && (l.underline = true),
    supportedFontFamilies.includes(e.font?.name) &&
      (l.fontFamily = e.font.name),
    e.font?.size !== undefined &&
      Number.isInteger(e.font.size) &&
      e.font.size >= 8 &&
      e.font.size <= 72 &&
      (l.fontSize = e.font.size),
    e.alignment?.wrapText && (l.wrap = true));
  let c = {};
  for (let m of cellBorderEdges) {
    let g = e.border?.[m];
    if (!g?.style) continue;
    if (!borderLineStyles.includes(g.style)) {
      r(
        "An unsupported border line was omitted; supported edges and base values were retained.",
      );
      continue;
    }
    let b = Ga(g.color?.argb);
    (g.color &&
      !b &&
      r(
        "A theme/indexed border color used a black fallback; theme bindings are not preserved.",
      ),
      (c[m] = { style: g.style, color: b ?? "#000000" }));
  }
  let u = Object.keys(c),
    d = u.every(
      (m) => c[m]?.style === "thin" && c[m]?.color?.toLowerCase() === "#000000",
    );
  (d && u.length === 4
    ? (l.border = "all")
    : d && u.length === 1 && c.bottom
      ? (l.border = "bottom")
      : u.length && (l.borders = c),
    e.border?.diagonal &&
      r("Diagonal borders are unsupported and were omitted."),
    e.font?.name &&
      !supportedFontFamilies.includes(e.font.name) &&
      r(
        "An unsupported font family uses the editor default; exact font fidelity is not preserved.",
      ),
    e.font?.size !== undefined &&
      (!Number.isInteger(e.font.size) || e.font.size < 8 || e.font.size > 72) &&
      r(
        "An unsupported font size was omitted; supported sizes are integers from 8 to 72 points.",
      ),
    e.font?.underline &&
      e.font.underline !== true &&
      e.font.underline !== "single" &&
      r("An unsupported underline variant was simplified to single underline."),
    (e.font?.vertAlign ||
      e.alignment?.textRotation ||
      e.alignment?.shrinkToFit ||
      (e.alignment?.horizontal &&
        !["left", "center", "right"].includes(e.alignment.horizontal)) ||
      (e.alignment?.vertical &&
        !["top", "middle", "bottom"].includes(e.alignment.vertical))) &&
      r(
        "Unsupported text positioning/alignment settings were omitted; raw values were retained.",
      ),
    e.font?.color &&
      !Ga(e.font.color.argb) &&
      r(
        "A theme/indexed font color uses the editor default; theme bindings are not preserved.",
      ),
    e.fill &&
      (e.fill.type !== "pattern" ||
        !["solid", "none"].includes(e.fill.pattern)) &&
      r(
        "An unsupported gradient or pattern fill was omitted; base cell values were retained.",
      ),
    e.fill?.type === "pattern" &&
      e.fill.pattern === "solid" &&
      e.fill.fgColor &&
      !Ga(e.fill.fgColor.argb) &&
      r(
        "A theme/indexed fill color uses the editor default; theme bindings are not preserved.",
      ));
  let f = Ga(e.font?.color?.argb);
  if (
    (f && (l.color = f),
    e.fill?.type === "pattern" && e.fill.pattern === "solid")
  ) {
    let m = Ga(e.fill.fgColor?.argb);
    m && (l.fill = m);
  }
  e.alignment &&
    ["left", "center", "right"].includes(e.alignment.horizontal ?? "") &&
    (l.align = e.alignment.horizontal);
  let p =
      typeof e.note == "string"
        ? e.note
        : e.note?.texts?.map((m) => m.text).join(""),
    h = e.numFmt && e.numFmt !== "General" ? e.numFmt : undefined;
  return {
    raw: a,
    ...(s ? { literal: s } : {}),
    ...(Object.keys(l).length ? { style: l } : {}),
    ...(h ? { numFmt: h, format: Lq(h) } : {}),
    ...(p ? { note: p } : {}),
  };
}
async function kg(e) {
  if (/\.xltx$/i.test(e.name)) return zN(e);
  if (/\.json$/i.test(e.name)) {
    if (e.size > workbookLimits.maxNativeBytes)
      throw new Error("Native workbook exceeds the 50 MB import limit.");
    let S = Fq(await e.text());
    return {
      workbook: S,
      warnings: [...collectStyleFallbackWarnings(S), ...Bw(S)],
    };
  }
  if (e.size > workbookLimits.maxFileBytes)
    throw new Error("Files must be 10 MB or smaller.");
  if (/\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(e.name))
    throw new Error(
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (/\.(csv|tsv)$/i.test(e.name))
    return Eg(new Uint8Array(await e.arrayBuffer()), e.name, {
      hasHeader: false,
    });
  if (!/\.xlsx$/i.test(e.name))
    throw new Error("Choose an XLSX, CSV, TSV or JSON file.");
  let t = new Uint8Array(await e.arrayBuffer()),
    n = {
      names: [],
      nameWarnings: [],
      numericTokens: new Map(),
      builtinNameXml: "",
      hiddenRows: new Map(),
      conditionalStops: new Map(),
      validationAreas: new Map(),
      printBreaks: new Map(),
      columns: new Map(),
      sheetRelationships: new Map(),
      sheetNames: new Map(),
    };
  if (t.byteLength > workbookLimits.maxFileBytes)
    throw new Error("XLSX files must be 10 MB or smaller.");
  let i = v8(t, n),
    r = retainOriginalXlsx(t, e.name),
    o = xk(Nd(t), r.sha256),
    a = new y8.default.Workbook(),
    s = t;
  if (n.names.length || n.nameWarnings.length) {
    let S = Nd(t);
    ((S["xl/workbook.xml"] = za(
      insertDefinedNamesIntoWorkbookXml(
        Ao(S["xl/workbook.xml"]),
        n.builtinNameXml,
      ),
    )),
      (s = Rd(S)));
  }
  try {
    await a.xlsx.load(s);
  } catch {
    throw new Error(
      "The XLSX workbook contains corrupt or unsupported records. Open the original in Excel, repair it if needed, and save a clean XLSX copy.",
    );
  }
  if (
    a.worksheets.length !== n.sheetNames.size ||
    a.worksheets.some((S) => !n.sheetNames.has(S.name))
  )
    throw new Error(
      "Unsupported or missing worksheet records. Save a standard XLSX copy with ordinary worksheets.",
    );
  if (!a.worksheets.length)
    throw new Error("A workbook must have at least one sheet.");
  let l = 0,
    c = 0,
    u = a.worksheets.map((S) => {
      let O = Math.max(S.columnCount, S.columns?.length ?? 0);
      if (S.rowCount > workbookLimits.maxRows || O > workbookLimits.maxCols)
        throw new Error(
          `Worksheet ${S.name} exceeds ${workbookLimits.maxRows} rows or ${workbookLimits.maxCols} columns.`,
        );
      let E = {},
        R = S.model,
        I = {},
        U = n.sheetRelationships.get(n.sheetNames.get(S.name) ?? ""),
        P = n.hiddenRows.get(U ?? "") ?? [],
        T = (N, F) => {
          let z = N.at(-1);
          if (z && F === z.end + 1) z.end = F;
          else {
            if (N.length >= worksheetCapacityLimits.maxIntervalsPerAxis)
              throw new Error(
                "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
              );
            N.push({ start: F, end: F });
          }
        };
      for (let N of R.rows) {
        if (typeof N.height == "number" && Number.isFinite(N.height)) {
          let F = (N.height * 4) / 3;
          ((I[String(N.number - 1)] = Math.max(21, Math.min(546, F))),
            (F < 21 || F > 546) && c++);
        }
        for (let F of N.cells) {
          let z = S.getCell(F.address);
          if (z.isMerged && z.master.address !== z.address) continue;
          let j = Uq(
            z,
            i,
            n.numericTokens.get(U ?? "")?.get(F.address),
            !!a.properties.date1904,
          );
          if (j) {
            if (++l > workbookLimits.maxCells)
              throw new Error("Workbook exceeds 100,000 populated cells.");
            E[z.address] = j;
          }
        }
      }
      let A = S.views?.find((N) => N.state === "frozen"),
        L = [];
      (S.columns ?? []).forEach((N, F) => {
        N.hidden && T(L, F);
      });
      let y = {};
      for (let N of n.columns.get(U ?? "") ?? []) {
        if (N.width === undefined) continue;
        let F = N.width * 7;
        (F < 40 || F > 600) &&
          i.push(
            "Worksheet " +
              S.name +
              " column widths were clamped to the supported pixel bounds.",
          );
        let z = Math.max(40, Math.min(600, F));
        for (let j = N.start; j <= N.end; j++) y[j] = z;
      }
      let v = (N, F, z, j, q) => {
          if (N === undefined || N === 0) return;
          if (typeof N != "number" || !Number.isFinite(N) || N <= 0)
            throw Error("Invalid worksheet default " + q + ".");
          let D = N * F;
          return (
            (D < z || D > j) &&
              i.push(
                "Worksheet " +
                  S.name +
                  " default " +
                  q +
                  " was clamped to the supported pixel bounds.",
              ),
            Math.max(z, Math.min(j, D))
          );
        },
        k = v(S.properties.defaultRowHeight, 4 / 3, 21, 546, "row height"),
        x = v(S.properties.defaultColWidth, 7, 40, 600, "column width"),
        C = {
          id: Nk(),
          name: S.name,
          ...(k !== undefined ? { defaultRowHeight: k } : {}),
          ...(x !== undefined ? { defaultColumnWidth: x } : {}),
          cells: E,
          rowCount: workbookLimits.maxRows,
          colCount: workbookLimits.maxCols,
          frozenRows: A?.ySplit ?? 0,
          ...(A?.xSplit ? { frozenCols: A.xSplit } : {}),
          ...(P.length ? { hiddenRows: P } : {}),
          ...(L.length ? { hiddenCols: L } : {}),
          ...(Object.keys(y).length ? { columnWidths: y } : {}),
          ...(Object.keys(I).length ? { rowHeights: I } : {}),
          ...(S.state !== "visible" ? { hidden: true } : {}),
          ...(S.state === "veryHidden" ? { veryHidden: true } : {}),
          ...(Ga(S.properties.tabColor?.argb)
            ? { tabColor: Ga(S.properties.tabColor?.argb) }
            : {}),
        };
      (i.push(...j5(S, C)),
        i.push(
          ...YN(
            S,
            C,
            n.conditionalStops.get(U ?? ""),
            n.validationAreas.get(U ?? ""),
          ),
        ),
        i.push(...SN(a, S, C, { containerFilename: e.name })));
      try {
        lN(S, C);
        let N = n.printBreaks.get(U ?? "");
        N && validatePageSetup({ ...C.pageLayout, ...N }, C);
      } catch (N) {
        i.push(`Page setup on ${C.name} could not be retained: ${N.message}`);
      }
      return C;
    });
  (c &&
    i.push(
      `${c} explicit row height(s) were adjusted to the supported ${21}\u2013${546} CSS-pixel range.`,
    ),
    a.properties.date1904 &&
      i.push(
        "The source uses the 1904 date system. Its date system, numeric serials and formulas are preserved; calendar dates are not silently converted to 1900.",
      ),
    i.push(
      "Formula results are recalculated by the supported local formula engine. Unsupported Excel functions may display errors. Cell error values are imported as displayed text. Audit history and application tickmarks are available only in native JSON backups.",
    ));
  let d = $q(e.name, u);
  d.xlsxSource = r;
  let f = FN(Nd(t), r.sha256);
  (f.length &&
    (d.workbookLinks = { schemaVersion: 1, sources: [], packageTargets: f }),
    a.properties.date1904 &&
      (d.datePolicy = { ...defaultDatePolicy, dateSystem: "1904" }));
  let p = 0,
    h = [];
  for (let [S, O] of n.sheetNames)
    for (let [E, R] of n.numericTokens.get(n.sheetRelationships.get(O) ?? "") ??
      []) {
      let I = VN(R);
      I && (p++, h.length < 8 && h.push(`${S}!${E}: ${R.slice(0, 80)} (${I})`));
    }
  (p &&
    i.push(
      `Numeric precision risk: ${p} numeric value/cache(s) may lose significant digits in Excel or binary conversion. Numeric cells retain their declared numeric type; store identifiers as Text and keep the source file. ${h.join("; ")}`,
    ),
    a.created instanceof Date &&
      Number.isFinite(a.created.getTime()) &&
      a.modified instanceof Date &&
      Number.isFinite(a.modified.getTime()) &&
      a.modified >= a.created &&
      (d.metadata = {
        createdAt: a.created.toISOString(),
        modifiedAt: a.modified.toISOString(),
      }));
  let m = s5(n.names, d, [...n.sheetNames.keys()]);
  m.length && (d.definedNames = m);
  let g = n.calculation ?? {
    mode: "auto",
    iterative: false,
    maxIterations: null,
    maxChange: null,
  };
  t8(d, g, i);
  let b = gi(Wm(d, { id: Nk(), name: e.name, kind: "file" })),
    w = _k("import", e.name, t.byteLength, b, i);
  ((w.partInventory = o),
    w.warnings.push(
      "The exact source XLSX is retained in this workbook and native JSON. Passive custom metadata can be copied unchanged; all other unsupported parts remain accessible in the original. Save native JSON to retain that source across sessions.",
    ));
  let _ = {
    schemaVersion: 1,
    fileName: e.name,
    sourceSha256: Array.from(
      new Uint8Array(
        await crypto.subtle.digest("SHA-256", new Uint8Array(t).buffer),
      ),
    )
      .map((S) => S.toString(16).padStart(2, "0"))
      .join(""),
    importedAt: new Date().toISOString(),
    sourceSettings: g,
    entries: [],
  };
  return (
    a8(
      a,
      b,
      w,
      new Map(
        [...n.sheetNames].map(([S, O]) => [
          S,
          n.emptyStringCaches?.get(n.sheetRelationships.get(O) ?? "") ??
            new Set(),
        ]),
      ),
      _,
      new Map(
        [...n.sheetNames].map(([S, O]) => [
          S,
          n.numericTokens.get(n.sheetRelationships.get(O) ?? "") ?? new Map(),
        ]),
      ),
    ),
    (b.importedCalculationCaches = _),
    w.warnings.push(...Bw(b)),
    { workbook: gi(b), warnings: w.warnings, report: w }
  );
}
function Fq(e) {
  if (new TextEncoder().encode(e).length > workbookLimits.maxNativeBytes)
    throw new Error("Native workbook exceeds the 50 MB import limit.");
  let t;
  try {
    t = JSON.parse(e);
  } catch {
    throw new Error("The native workbook is not valid JSON.");
  }
  return gi(t);
}
function zq() {
  return {
    names: [],
    nameWarnings: [],
    numericTokens: new Map(),
    builtinNameXml: "",
    hiddenRows: new Map(),
    conditionalStops: new Map(),
    validationAreas: new Map(),
    printBreaks: new Map(),
    columns: new Map(),
    sheetRelationships: new Map(),
    sheetNames: new Map(),
  };
}
function w8(e) {
  let t = requireRetainedXlsxSource(e),
    n = v8(t.bytes, zq()),
    i = _k("import", t.fileName, t.bytes.length, e, n);
  return ((i.partInventory = xk(Nd(t.bytes), t.sha256)), i);
}

var An = class extends Error {
    constructor(n, i = fileProcessingErrorMessages[n]) {
      super(i);
      this.code = n;
      this.name = "FileProcessingError";
    }
  },
  Bq = new Set([
    ...unsupportedXlsxMarkers,
    "Legacy CSE array formula",
    "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)",
  ]),
  Ok =
    " Open the original in Excel, repair it if needed, and save a clean, unencrypted XLSX copy.",
  jq = excelMessagePatterns.map((e) => ({
    parts: e,
    expression: new RegExp(
      "^" +
        e
          .map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
          .join("([\\s\\S]*?)") +
        "$",
    ),
  }));
function Sg(e, t = 0) {
  if (!(e.length > 4096 || t > 3)) {
    if (Bq.has(e)) return e;
    if (e.endsWith(Ok)) {
      let n = Sg(e.slice(0, -Ok.length), t + 1);
      if (n) return n + Ok;
    }
    for (let { parts, expression } of jq) {
      let r = expression.exec(e);
      if (r)
        return parts
          .map(
            (o, a) =>
              o +
              (a < parts.length - 1
                ? (Sg(r[a + 1], t + 1) ?? "[source detail omitted]")
                : ""),
          )
          .join("");
    }
  }
}
function Wq(e) {
  if (e instanceof An)
    return new An(e.code, Sg(e.message) ?? fileProcessingErrorMessages[e.code]);
  let t = e instanceof Error ? e.message : "",
    n =
      /exceed|budget|limit|too (?:large|many)|at most|smaller|depth|nesting/i.test(
        t,
      )
        ? "FILE_LIMIT"
        : /macro|active.?content|activex|executable|unsafe|document type|doctype|entity|dangerous|DDE/i.test(
              t,
            )
          ? "FILE_ACTIVE"
          : "FILE_INVALID";
  return new An(n, Sg(t) ?? fileProcessingErrorMessages[n]);
}
function _8(e) {
  let { code, message } = Wq(e);
  return { code, message };
}
var Nc = Object.freeze({
  maxFileBytes: workbookLimits.maxFileBytes,
  maxNativeBytes: workbookLimits.maxNativeBytes,
  maxConfigurationBytes: 2e6,
  maxPictureBytes: 2 * 1024 * 1024,
  maxConcurrent: 2,
  defaultTimeoutMs: 3e4,
  maxTimeoutMs: 6e4,
  maxJsonDepth: 128,
  maxJsonNodes: 2e6,
  nodeHeapMiB: 256,
  maxResultBytes: 100 * 1024 * 1024,
});
function $k(e, t) {
  return e === "source"
    ? 4 * 1024 * 1024
    : e === "json"
      ? Nc.maxConfigurationBytes
      : e === "picture"
        ? Nc.maxPictureBytes
        : e === "workbook" && /\.json$/i.test(t)
          ? workbookLimits.maxNativeBytes
          : workbookLimits.maxFileBytes;
}
function E8(e, t, n) {
  if (
    typeof t != "string" ||
    !t.length ||
    t.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(t)
  )
    throw new An("FILE_TYPE");
  if (e === "workbook" && /\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(t))
    throw new An(
      "FILE_TYPE",
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (
    !{
      workbook: /\.(xlsx|xltx|csv|tsv|json)$/i,
      source: /\.(json|xml|pdf|png|jpe?g|html?)$/i,
      formats: /\.(xls|xlsb|xlsm|xltx|xltm|ods)$/i,
      delimited: /\.(csv|tsv)$/i,
      automation: /\.(bas|vba|cls|ts|osts|xlsm|xlam|xltm|xlsx|xltx|xls|xlsb)$/i,
      json: /\.json$/i,
      picture: /\.(png|jpe?g)$/i,
    }[e]?.test(e === "picture" ? t.trim() : t)
  )
    throw new An("FILE_TYPE");
  if (!Number.isSafeInteger(n) || n < 0 || n > $k(e, t))
    throw new An(
      "FILE_LIMIT",
      $k(e, t) === workbookLimits.maxNativeBytes
        ? "Native workbook exceeds the 50 MB import limit."
        : $k(e, t) === workbookLimits.maxFileBytes
          ? "Files must be 10 MB or smaller."
          : undefined,
    );
}
for (let e of [
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "EventSource",
  "importScripts",
  "Worker",
  "SharedWorker",
  "BroadcastChannel",
  "indexedDB",
  "caches",
])
  Object.defineProperty(globalThis, e, {
    value: undefined,
    writable: false,
    configurable: false,
  });
function S8(e) {
  let t = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(e)),
    n = [{ value: t, depth: 0 }],
    i = 0;
  for (; n.length;) {
    let r = n.pop();
    if (++i > Nc.maxJsonNodes || r.depth > Nc.maxJsonDepth)
      throw new An("FILE_LIMIT");
    if (r.value && typeof r.value == "object")
      for (let [o, a] of Object.entries(r.value)) {
        if (["__proto__", "prototype", "constructor"].includes(o))
          throw new An("FILE_ACTIVE");
        n.push({ value: a, depth: r.depth + 1 });
      }
  }
  return t;
}
var Pk = self,
  C8 = false;
Pk.onmessage = async ({ data }) => {
  if (C8) return;
  C8 = true;
  let t = data?.id;
  try {
    if (
      typeof t != "string" ||
      t.length > 100 ||
      !(data.bytes instanceof Uint8Array)
    )
      throw new An("FILE_INVALID");
    E8(data.operation, data.fileName, data.bytes.length);
    let n;
    switch (data.operation) {
      case "source":
        n = gk(data.bytes, data.fileName);
        break;
      case "formats":
        n = Vd(data.bytes, data.fileName);
        break;
      case "workbook": {
        /\.json$/i.test(data.fileName) && S8(data.bytes);
        let i = data.bytes,
          r = await kg({
            name: data.fileName,
            size: i.length,
            arrayBuffer: async () => i.buffer,
            text: async () =>
              new TextDecoder("utf-8", { fatal: true }).decode(i),
          });
        if (
          (/\.json$/i.test(data.fileName) &&
            r.workbook.xlsxSource &&
            w8(r.workbook),
          /\.json$/i.test(data.fileName) && r.workbook.alternateSource)
        ) {
          let o = r.workbook.alternateSource;
          Vd(validateAlternateSource(o), o.fileName);
        }
        if (/\.json$/i.test(data.fileName))
          for (let o of r.workbook.sourceImports ?? []) {
            let a = gk(
              validateRetainedSourceFile(o.original),
              o.original.fileName,
            );
            if (
              a.method !== "document-adapter" &&
              JSON.stringify(a.tables.find((s) => s.id === o.table.id)) !==
                JSON.stringify(o.table)
            )
              throw Error(
                "Retained structured evidence does not match its original.",
              );
          }
        n = r;
        break;
      }
      case "delimited": {
        let i = b8(data.bytes, data.fileName, data.options);
        n = {
          report: i,
          ...(i.valid
            ? { workbook: Eg(data.bytes, data.fileName, data.options).workbook }
            : {}),
        };
        break;
      }
      case "automation":
        n = o8(data.bytes, data.fileName);
        break;
      case "json":
        n = S8(data.bytes);
        break;
      case "picture": {
        let i = "";
        for (let a of data.bytes) i += String.fromCharCode(a);
        let r = `data:image/${/\.png$/i.test(data.fileName.trim()) ? "png" : "jpeg"};base64,${btoa(i)}`,
          o = inspectPictureDataUrl(r);
        n = { dataURL: r, width: o.width, height: o.height };
        break;
      }
      default:
        throw new An("FILE_TYPE");
    }
    if (new TextEncoder().encode(JSON.stringify(n)).length > Nc.maxResultBytes)
      throw new An("FILE_LIMIT");
    Pk.postMessage({ id: t, ok: true, value: n });
  } catch (n) {
    Pk.postMessage({ id: t, ok: false, error: _8(n) });
  }
};
