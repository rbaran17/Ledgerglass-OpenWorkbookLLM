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
  dateToSerial,
  deepFrozenValueSet as $V,
  defaultDatePolicy,
  fingerprintCanonicalValue,
  hasLinkedOrBoundCells,
  isDeepFrozenValue,
  linkedOrBoundCellCache as bN,
  linkedOrBoundCellCache as r8,
  millisecondsPerDay,
  normalizeDatePolicy,
  objectDefineProperty,
  sha256HexDigest,
  validateExtensionState,
  validateSavedRecipeMetadata,
} from "../app/01-schemas-date-functions-and-validation.js";
import {
  fileImportLimits,
  supportedFontFamilies,
  workbookLimits,
} from "../app/02-text-and-financial-functions.js";
import {
  assertKnownStyleKeys,
  borderLineStyles,
  canonicalStyleKey,
  cellBorderEdges,
  cloneStyleValue,
  collectStyleFallbackWarnings,
  createWorkbookMetadata,
  failWithMessage as Fr,
  isNonArrayObject as Gs,
  isNonArrayObject as yg,
  isPlainObjectLiteral as bo,
  officeThemeDefinition,
  resolveNamedStyleBinding,
  throwWithMessage as jt,
  validateBoundedStyleText,
  validateCellStyle,
  validateNamedStyleBinding,
  validateNamedStyleDefinition,
  validateOutlineGroups,
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
  importFormulaCacheEvidence,
  verifyConsolidationEvidence,
  zie as dR,
} from "../app/03-workbook-model/05-formula-engine-and-name-operations.js";
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
} from "../app/03-workbook-model/07-conditional-rules-worksheet-features.js";
import {
  Cn as Ks,
  normalizeTableDefinitions,
  resolveSourceRange,
  tablePaletteStyleNames,
  validateColumnName,
  validateDrawingObjects,
  validateTableName,
} from "../app/03-workbook-model/08-table-pivot-query-definitions.js";
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
} from "../app/03-workbook-model/10-table-filters-and-pivot-views.js";
import {
  normalizeCellRecord,
  normalizeCellRecordInput,
  normalizeTextCleanups,
  validateTableSourceEvidence,
} from "../app/03-workbook-model/11-table-model.js";

import {
  F$,
  Hle,
  Wle,
  buildContainsTextFormula,
  buildDelimitedImportReport,
  buildExchangeSchemaReport,
  byteCharacterTable,
  computeCrc32,
  convertTemplateToXlsx,
  createEmptyExportLedger,
  createWorkbookParseState,
  dataValidationOperators,
  decodeXmlEntities,
  formatRangeCornerCells,
  getXmlAttribute,
  inspectAlternateWorkbookArchive,
  inspectXlsxPackage,
  isExternalActionFormula,
  isExternalReferenceFormula,
  maxXlsxExpandedByteLimit,
  normalizeDelimitedOptions,
  parseAlternateWorkbookXml,
  parseCalculationProperties,
  parseWorksheetCells,
  parseWorksheetXml,
  parseXlsxDataValidations,
  quoteFormulaString,
  readOpenDocumentPackage,
  recordDelimitedChange,
  recordDelimitedIssue,
  recordSourceInspection,
  resolvePackagePartPath,
  resolveRelationshipTarget,
  scanXlsxNumericTokens,
  u1,
  validateConditionalFormatRules,
  validateIsolatedFormatInspection,
  validateWorkbookPackage,
  validateXlsxNumericValue,
  xlsxExchangeSupportNotice,
} from "../app/06c-package-and-worksheet-parsing.js";

import { buildExternalSourceRefreshState } from "../app/03-workbook-model/05-formula-engine-and-name-operations.js";
import { validateFreezePanes } from "../app/03-workbook-model/07-conditional-rules-worksheet-features.js";
import {
  FileProcessingError,
  assertImportFileAllowed,
  parseDelimitedCellValue,
  toFileProcessingError,
} from "../app/06c-package-and-worksheet-parsing.js";

var createPrototypeObject = Object.create;

var readPropertyDescriptor = Object.getOwnPropertyDescriptor;
var listOwnPropertyNames = Object.getOwnPropertyNames;
var readPrototypeOf = Object.getPrototypeOf,
  hasOwnPropertyCheck = Object.prototype.hasOwnProperty;
var dynamicRequire = ((e) =>
  typeof require < "u"
    ? require
    : typeof Proxy < "u"
      ? new Proxy(e, { get: (t, n) => (typeof require < "u" ? require : t)[n] })
      : e)(function (e) {
  if (typeof require < "u") return require.apply(this, arguments);
  throw Error('Dynamic require of "' + e + '" is not supported');
});
var copyModuleExports = (e, t, n, i) => {
  if ((t && typeof t == "object") || typeof t == "function")
    for (let r of listOwnPropertyNames(t))
      !hasOwnPropertyCheck.call(e, r) &&
        r !== n &&
        objectDefineProperty(e, r, {
          get: () => t[r],
          enumerable: !(i = readPropertyDescriptor(t, r)) || i.enumerable,
        });
  return e;
};
var toEsModule = (e, t, n) => (
  (n = e != null ? createPrototypeObject(readPrototypeOf(e)) : {}),
  copyModuleExports(
    t || !e || !e.__esModule
      ? objectDefineProperty(n, "default", { value: e, enumerable: true })
      : n,
    e,
  )
);

var getSaxesModule = () => __lib_saxes;

var getExceljsModule = () => __lib_exceljs;
var getPapaparseModule = () => __lib_papaparse;
var saxesEsModuleForXml = toEsModule(getSaxesModule(), 1);

function dateToExcelSerial(e, t = "1900") {
  return (
    dateToSerial(e.getUTCFullYear(), e.getUTCMonth() + 1, e.getUTCDate(), t) +
    (e.getUTCHours() * 36e5 +
      e.getUTCMinutes() * 6e4 +
      e.getUTCSeconds() * 1e3 +
      e.getUTCMilliseconds()) /
      millisecondsPerDay
  );
}

var namedStyleLimits = { maxStyles: 50, maxCells: 500, maxName: 80 };

function normalizeWorkbookStyleCatalog(e) {
  (bo(e) || Fr("Workbook styles must be an object."),
    assertKnownStyleKeys(
      e,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    e.missingTheme !== undefined &&
      e.missingTheme !== true &&
      Fr("Invalid missing theme flag."),
    (!Array.isArray(e.styles) ||
      e.styles.length > namedStyleLimits.maxStyles) &&
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

function applyNamedStyleCatalog(e, t) {
  let n = e === undefined ? undefined : normalizeWorkbookStyleCatalog(e);
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

function createBlankWorkbook(e = "Book1") {
  return toValidWorkbook({
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

function collectDataValidationWarnings(e) {
  if (!e.sheets.some((n) => n.validations?.length)) return [];
  let t = summarizeValidationInspection(e);
  return t.invalidCells || !t.complete
    ? [
        `Data validation: ${t.message} Source values were retained; use Data Validation to review.`,
      ]
    : [];
}

function resolveDefinedNameScopes(e, t, n) {
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

function createSourceRowProvenance(e, t) {
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

function convertImportedTables(e, t) {
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

function validateReportSnapshot(e) {
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
function validateReportOutputs(e, t) {
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
    t.reportOutputs = n.map((i) => ({
      ...i,
      snapshot: validateReportSnapshot(i.snapshot),
    }));
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

function importPageSetup(e, t) {
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

var fieldTransformLimits = {
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

function validateFieldTransformEvidence(e, t) {
  if (e === undefined) return;
  if (
    !Array.isArray(e) ||
    e.length > 100 ||
    JSON.stringify(e).length > fieldTransformLimits.evidenceCharacters
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

function toValidWorkbook(e) {
  return normalizeWorkbookDocument(e, false);
}
function normalizeWorkbookDocument(e, t) {
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
        validateFreezePanes(p, S, { rows: g, cols: b }),
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
    buildExternalSourceRefreshState(e, f),
    validateDataModelReport(e, f),
    validateAnalysisReport(e, f),
    validateReportOutputs(e, f),
    normalizeRowProvenance(e, f),
    validateAutomationArtifacts(e, f),
    readRetainedXlsxSource(e, f),
    normalizeAlternateFormatSource(e, f),
    validateSourceImports(e, f),
    attachWorkbookLinkMetadata(e, f),
    normalizeTextCleanups(e.textCleanups, f),
    validateFieldTransformEvidence(e.fieldTransforms, f),
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
    applyNamedStyleCatalog(e.styleCatalog, f),
    e.formulaWatches !== undefined &&
      (f.formulaWatches = validateFormulaWatches(e.formulaWatches, f)),
    dR(e, f),
    copyExtensionCacheEntries(e, f),
    f
  );
}

var bytesToBase64 = (e) => {
  let t = "";
  for (let n of e) t += String.fromCharCode(n);
  return btoa(t);
};
function collectImportedPictures(e, t, n, i = {}) {
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
            : `data:image/${u};base64,${bytesToBase64(new Uint8Array(c.buffer))}`,
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

var escapeJsonPointerSegment = (e) =>
  e.replaceAll("~", "~0").replaceAll("/", "~1");
function parseStrictJsonText(e) {
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
function describeStructuredValue(e, t) {
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
function collectJsonSourceTables(e) {
  let t = [],
    n = (i, r) => {
      if (Array.isArray(i)) i.length && t.push({ value: i, path: r });
      else if (yg(i))
        for (let [o, a] of Object.entries(i))
          n(a, r + "/" + escapeJsonPointerSegment(o));
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
              for (let [g, b] of Object.entries(h))
                p(b, m + "/" + escapeJsonPointerSegment(g));
            else (a.add(m), d.set(m, describeStructuredValue(h, f + m)));
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
function parseXmlSourceTables(e) {
  if (/<!DOCTYPE|<!ENTITY/i.test(e))
    throw Error("XML document types and entities are not supported.");
  if (/<\?xml[^>]*encoding\s*=\s*['"](?!utf-8['"])/i.test(e))
    throw Error("Use UTF-8 XML.");
  let t = [],
    n = [],
    i = 0,
    r = new saxesEsModuleForXml.SaxesParser({ xmlns: true }),
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
function inspectSourceFile(e, t) {
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
    ((r =
      i === "json"
        ? collectJsonSourceTables(parseStrictJsonText(c))
        : parseXmlSourceTables(c)),
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
  return (recordSourceInspection(l, e, t), l);
}

function inspectCompoundFileContainer(e) {
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

function buildAlternateFormatGuidance(e) {
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

function inspectAlternateWorkbookFile(e, t) {
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
    let s = inspectCompoundFileContainer(e);
    ((i = "xls"),
      (r =
        s.version + " workbook in a CFB container; binary cells not decoded"),
      (o = s.paths.map((l) => ({
        path: l,
        bytes: /^(Workbook|Book)$/i.test(l) ? s.workbookBytes : 0,
        disposition: "original-only",
      }))));
  } else {
    let s = inspectAlternateWorkbookArchive(e);
    ((i = readOpenDocumentPackage(s)),
      (r =
        i === "ods"
          ? "ODF spreadsheet MIME, manifest and content"
          : "OOXML main content type and package relationship"));
    for (let [l, c] of Object.entries(s))
      /\.(xml|rels)$/i.test(l) && parseAlternateWorkbookXml(c, () => {});
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
    steps: buildAlternateFormatGuidance(n),
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
  return (validateIsolatedFormatInspection(a, e, t), a);
}
var saxesEsModuleForLinks = toEsModule(getSaxesModule(), 1);
function parseExternalLinkTargets(e, t) {
  let n = (a, s) => {
      if (!e[a]) return;
      let l = new saxesEsModuleForLinks.SaxesParser({ xmlns: true });
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
        resolveRelationshipTarget(
          "xl/_rels/workbook.xml.rels",
          a.attributes.Target.value,
        ),
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

async function convertXltxTemplateFile(e) {
  if (e.size > 10 * 1024 * 1024) throw Error("Files must be 10 MB or smaller.");
  let t = new Uint8Array(await e.arrayBuffer()),
    n = inspectAlternateWorkbookFile(t, e.name);
  if (n.format !== "xltx")
    throw Error("Only XLTX templates use the native template adapter.");
  let i = convertTemplateToXlsx(t, "xltx", "xlsx"),
    r = e.name.replace(/\.xltx$/i, "").slice(0, 236) + ".template-copy.xlsx",
    o = await importWorkbookFile({
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

var argbColorOrFallback = (e, t) =>
  e?.argb && /^[a-f\d]{8}$/i.test(e.argb)
    ? `#${e.argb.slice(-6).toLowerCase()}`
    : t;
function collapseCellAddressesToRanges(e) {
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

function buildDataValidationRule(e) {
  let t = {
      list: "list",
      decimal: "number",
      whole: "integer",
      textLength: "textLength",
      date: "date",
      time: "time",
      custom: "custom",
    }[e.type],
    n = Object.entries(dataValidationOperators).find(
      ([, u]) => u === e.operator,
    )?.[0];
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
        l.length > 8192 || l.length < 2 || isExternalReferenceFormula(l))))
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

function buildConditionalFormatRule(e, t) {
  let n =
      e.style?.fill?.type === "pattern"
        ? argbColorOrFallback(e.style.fill.fgColor, "#fff2cc")
        : "#fff2cc",
    i = argbColorOrFallback(e.style?.font?.color, "#000000"),
    r = {
      id: GN(),
      priority: Number.isFinite(e.priority) ? e.priority : 1e4,
      range: formatRangeCornerCells(t),
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
          : "NOT(ISERROR(SEARCH(" + quoteFormulaString(d) + "," + u + ")))";
    return p.length > 4095 || isExternalReferenceFormula("=" + p)
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
      minColor: argbColorOrFallback(e.color?.[0], "#f8e99b"),
      maxColor: argbColorOrFallback(e.color?.[1], "#63be7b"),
    };
  if (
    e.type === "dataBar" &&
    e.cfvo?.length === 2 &&
    ["min", "autoMin"].includes(e.cfvo[0].type) &&
    ["max", "autoMax"].includes(e.cfvo[1].type) &&
    e.showValue !== false &&
    e.direction !== "rightToLeft"
  )
    return {
      ...r,
      kind: "dataBar",
      fill: argbColorOrFallback(e.color, "#4472c4"),
    };
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
    if (!c || c.length > 4095 || isExternalReferenceFormula("=" + c))
      return null;
    if (c.toUpperCase() === Wle("duplicate", t).toUpperCase())
      return { ...r, kind: "duplicate" };
    if (c.toUpperCase() === Wle("unique", t).toUpperCase())
      return { ...r, kind: "unique" };
    let u =
      /^ISNUMBER\(FIND\(LOWER\("((?:[^"]|"")*)"\),LOWER\(([A-Z]+[1-9]\d*)&""\)\)\)$/i.exec(
        c,
      );
    if (u) {
      let h = u[1].replace(/""/g, '"'),
        m = parseRangeAddress(t),
        g = Ks(m.firstRow, m.firstCol);
      if (
        h.length <= 1e3 &&
        c.toUpperCase() === buildContainsTextFormula(h, g).toUpperCase()
      )
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
function convertWorksheetFeatures(e, t, n = {}, i = []) {
  let r = [],
    o = {},
    a = e.model.merges ?? [];
  a.length && (o.merges = a.map(formatRangeCornerCells));
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
          : buildDataValidationRule(
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
    for (let S of collapseCellAddressesToRanges(addresses)) {
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
        let O = buildConditionalFormatRule(S, w);
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
              !Hle(P))
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

function applyCalculationProperties(e, t, n) {
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
var saxesEsModuleForAutomation = toEsModule(getSaxesModule(), 1);

function decodeUtf8Strict(e) {
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
function inspectAutomationArchive(e) {
  let t = inspectXlsxPackage(e, { allowBinaryWorkbookInventory: true }),
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
          if (((u = computeCrc32(u, f)), l.push(f), p)) {
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
function scanAutomationXmlMetadata(e, t, n) {
  if (/<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(e))
    throw Error(`Unsupported XML encoding in ${t}.`);
  let i = new saxesEsModuleForAutomation.SaxesParser({ xmlns: false }),
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
function inspectAutomationFile(e, t) {
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
    let o = decodeUtf8Strict(n),
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
    let o = inspectAutomationArchive(n),
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
          scanAutomationXmlMetadata(decodeUtf8Strict(c), l, (u, d) => {
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

function toTypedCellValue(e, t) {
  if (e != null) {
    if (e instanceof Date)
      return { type: "number", value: dateToExcelSerial(e, t) };
    if (typeof e == "object" && "error" in e && typeof e.error == "string")
      return { type: "error", value: e.error };
    if (typeof e == "string") return { type: "text", value: e };
    if (typeof e == "boolean") return { type: "boolean", value: e };
    if (typeof e == "number" && Number.isFinite(e))
      return { type: "number", value: e };
  }
}
function typedValuesMatch(e, t) {
  return (
    e.type === t.type &&
    (typeof e.value == "number" && typeof t.value == "number"
      ? Math.abs(e.value - t.value) <=
        1e-10 * Math.max(1, Math.abs(e.value), Math.abs(t.value))
      : e.value === t.value)
  );
}
function compareFormulaCachesWithRecalculation(
  e,
  t,
  n,
  i = new Map(),
  r,
  o = new Map(),
) {
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
            : toTypedCellValue(
                l.getCell(c).result,
                t.datePolicy?.dateSystem ?? "1900",
              ),
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
            ? typedValuesMatch(f, p)
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
function assertNoLegacyArrayFormulas(e, t) {
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
var exceljsEsModule = toEsModule(getExceljsModule(), 1);
var papaparseEsModule = toEsModule(getPapaparseModule(), 1);

function decodeDelimitedBytes(e, t) {
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
        ? Array.from(e, (o) => byteCharacterTable[o]).join("")
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
function tokenizeDelimitedText(e, t, n, i) {
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
          (recordDelimitedIssue(
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
          (recordDelimitedIssue(
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
        recordDelimitedIssue(
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
            ? (recordDelimitedIssue(
                n,
                s,
                r.length + 1,
                "Unexpected text after a closing quote.",
              ),
              (a = "plain"))
            : t.quote !== "none" &&
              p === t.quote &&
              recordDelimitedIssue(
                n,
                s,
                r.length + 1,
                "Quote inside an unquoted field; quote the entire field.",
              ),
          (o += p),
          (a = "plain"));
    o.length > fileImportLimits.maxCellTextLength &&
      (recordDelimitedIssue(
        n,
        s,
        r.length + 1,
        "A field exceeds the 32,767-character cell limit.",
      ),
      (c = true));
  }
  (!c &&
    a === "quoted" &&
    recordDelimitedIssue(n, s, r.length + 1, "Quoted field is not closed."),
    !c && l < e.length && d(e.length),
    (n.complete = !c));
}

function parseDelimitedRecords(e, t, n, i) {
  let r = normalizeDelimitedOptions(n),
    o = decodeDelimitedBytes(e, r.encoding);
  ((r.encoding = o.encoding),
    r.delimiter === "auto" &&
      (r.delimiter = /\.tsv$/i.test(t)
        ? "	"
        : papaparseEsModule.default.parse(o.text.slice(0, 65536), {
            preview: 10,
            quoteChar: r.quote === "none" ? "" : r.quote,
            escapeChar: r.escape === "double" ? r.quote : "\\",
            delimitersToGuess: [",", ";", "	", "|"],
          }).meta.delimiter || ","));
  let a = buildDelimitedImportReport("import", t, e.length, r),
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
    (tokenizeDelimitedText(o.text, r, a, (c, u, d) => {
      ((a.records = u),
        c.every((f) => f === "") && a.blankRecords++,
        !l && !d && (l = c.length),
        (a.columns = Math.max(a.columns, c.length)),
        !d &&
          l &&
          c.length !== l &&
          recordDelimitedIssue(
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
          g = parseDelimitedCellValue(p, h, r);
        } catch (b) {
          (recordDelimitedIssue(
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
              recordDelimitedIssue(
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
          recordDelimitedChange(
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
      recordDelimitedIssue(
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
function validateDelimitedImport(e, t, n = {}) {
  return parseDelimitedRecords(e, t, n, false).r;
}
function importDelimitedWorkbook(e, t, n = {}) {
  let { r, cells } = parseDelimitedRecords(e, t, n, true);
  if (!r.valid)
    throw Error(
      `Unable to read delimited data: ${r.issues[0]?.message ?? "Incomplete preview."} Review the import schema; no rows were imported.`,
    );
  let o = createBlankWorkbook(
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
      workbook: toValidWorkbook(
        createSourceRowProvenance(o, {
          id: crypto.randomUUID(),
          name: t.slice(0, 255),
          kind: "file",
        }),
      ),
      warnings: r.warnings,
      delimitedReport: r,
    }
  );
}
var createRandomUuid = () => globalThis.crypto.randomUUID();
var createImportedWorkbook = (e, t) => ({
  schemaVersion: 1,
  id: createRandomUuid(),
  name: e.replace(/\.[^.]+$/, "").slice(0, 200) || "Imported workbook",
  revision: 0,
  metadata: createWorkbookMetadata(),
  sheets: t,
  events: [],
});

function parseXlsxWorkbookPackage(e, t) {
  let n = inspectXlsxPackage(e),
    i = new Set(n.keys()),
    r = n.size,
    o = createWorkbookParseState();
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
        if (((a += P.length), (S += P.length), a > maxXlsxExpandedByteLimit)) {
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
        R = computeCrc32(R, P);
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
            if (
              (O && parseWorksheetXml(I, _.name, n, o),
              /^xl\/styles\.xml$/i.test(_.name))
            )
              for (let A of I.matchAll(/<(?:\w+:)?sz\b([^>]*)>/g)) {
                let L = Number(getXmlAttribute(A[1], "val"));
                (!Number.isInteger(L) || L < 8 || L > 72) && (m = true);
              }
            if (/^xl\/worksheets\//i.test(_.name)) {
              (assertNoLegacyArrayFormulas(I, _.name),
                (h = parseWorksheetCells(I, _.name, h)),
                /^xl\/worksheets\/[^/]+\.xml$/i.test(_.name) &&
                  t.numericTokens.set(_.name, scanXlsxNumericTokens(I)));
              for (let v of I.matchAll(/<(?:\w+:)?pageSetup\b([^>]*)>/gi)) {
                let k = getXmlAttribute(v[1], "pageOrder");
                if (
                  k !== undefined &&
                  !["overThenDown", "downThenOver"].includes(
                    decodeXmlEntities(k),
                  )
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
                  let D = getXmlAttribute(q[1], "id"),
                    W = getXmlAttribute(q[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(D ?? "") ||
                    !Number.isSafeInteger(Number(D)) ||
                    Number(D) >= x
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (W !== "1" && W !== "true") continue;
                  let $ = Number(getXmlAttribute(q[1], "min") ?? 0),
                    M = Number(getXmlAttribute(q[1], "max") ?? C - 1);
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
                t.conditionalStops.set(
                  _.name,
                  validateConditionalFormatRules(I),
                ),
                t.validationAreas.set(_.name, parseXlsxDataValidations(I)));
              let L = [],
                y = [];
              for (let v of I.matchAll(/<(?:\w+:)?col\b([^>]*)>/gi)) {
                let k = getXmlAttribute(v[1], "min"),
                  x = getXmlAttribute(v[1], "max"),
                  C = getXmlAttribute(v[1], "width"),
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
                let k = Number(getXmlAttribute(v[1], "r")) - 1;
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
                let L = getXmlAttribute(A[1], "name"),
                  y = getXmlAttribute(A[1], "r:id");
                L &&
                  y &&
                  t.sheetNames.set(decodeXmlEntities(L), decodeXmlEntities(y));
              }
            if (_.name === "xl/_rels/workbook.xml.rels")
              for (let A of I.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)) {
                let L = getXmlAttribute(A[1], "Id"),
                  y = getXmlAttribute(A[1], "Target");
                L &&
                  y &&
                  getXmlAttribute(A[1], "TargetMode") !== "External" &&
                  t.sheetRelationships.set(
                    decodeXmlEntities(L),
                    resolvePackagePartPath(decodeXmlEntities(y)),
                  );
              }
            for (let A of I.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (let L of decodeXmlEntities(A[3]).trim().split(/\s+/)) {
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
              t.calculation = parseCalculationProperties(I);
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
  let w = [
    xlsxExchangeSupportNotice,
    ...t.nameWarnings,
    ...validateWorkbookPackage(o, i),
  ];
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
function classifyNumberFormat(e) {
  if (!(!e || e === "General"))
    return /%/.test(e)
      ? "percent"
      : /[$€£¥]/.test(e)
        ? "currency"
        : /[ymdhis]/i.test(e.replace(/"[^"]*"/g, ""))
          ? "date"
          : "number";
}
function normalizeArgbHexColor(e) {
  return e && /^(?:[0-9a-f]{8}|[0-9a-f]{6})$/i.test(e)
    ? `#${e.slice(-6).toLowerCase()}`
    : undefined;
}
function convertXlsxCellToRecord(e, t = [], n, i = false) {
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
        ? (a =
            n !== undefined
              ? Number(n)
              : dateToExcelSerial(o, i ? "1904" : "1900"))
        : o &&
          typeof o == "object" &&
          ("formula" in o || "sharedFormula" in o
            ? ((a = `=${e.formula}`), isExternalActionFormula(a) && (s = true))
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
    let b = normalizeArgbHexColor(g.color?.argb);
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
      !normalizeArgbHexColor(e.font.color.argb) &&
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
      !normalizeArgbHexColor(e.fill.fgColor.argb) &&
      r(
        "A theme/indexed fill color uses the editor default; theme bindings are not preserved.",
      ));
  let f = normalizeArgbHexColor(e.font?.color?.argb);
  if (
    (f && (l.color = f),
    e.fill?.type === "pattern" && e.fill.pattern === "solid")
  ) {
    let m = normalizeArgbHexColor(e.fill.fgColor?.argb);
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
    ...(h ? { numFmt: h, format: classifyNumberFormat(h) } : {}),
    ...(p ? { note: p } : {}),
  };
}
async function importWorkbookFile(e) {
  if (/\.xltx$/i.test(e.name)) return convertXltxTemplateFile(e);
  if (/\.json$/i.test(e.name)) {
    if (e.size > workbookLimits.maxNativeBytes)
      throw new Error("Native workbook exceeds the 50 MB import limit.");
    let S = parseNativeWorkbookJson(await e.text());
    return {
      workbook: S,
      warnings: [
        ...collectStyleFallbackWarnings(S),
        ...collectDataValidationWarnings(S),
      ],
    };
  }
  if (e.size > workbookLimits.maxFileBytes)
    throw new Error("Files must be 10 MB or smaller.");
  if (/\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(e.name))
    throw new Error(
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (/\.(csv|tsv)$/i.test(e.name))
    return importDelimitedWorkbook(
      new Uint8Array(await e.arrayBuffer()),
      e.name,
      {
        hasHeader: false,
      },
    );
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
  let i = parseXlsxWorkbookPackage(t, n),
    r = retainOriginalXlsx(t, e.name),
    o = F$(Nd(t), r.sha256),
    a = new exceljsEsModule.default.Workbook(),
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
          let j = convertXlsxCellToRecord(
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
          id: createRandomUuid(),
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
          ...(normalizeArgbHexColor(S.properties.tabColor?.argb)
            ? { tabColor: normalizeArgbHexColor(S.properties.tabColor?.argb) }
            : {}),
        };
      (i.push(...convertImportedTables(S, C)),
        i.push(
          ...convertWorksheetFeatures(
            S,
            C,
            n.conditionalStops.get(U ?? ""),
            n.validationAreas.get(U ?? ""),
          ),
        ),
        i.push(
          ...collectImportedPictures(a, S, C, { containerFilename: e.name }),
        ));
      try {
        importPageSetup(S, C);
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
  let d = createImportedWorkbook(e.name, u);
  d.xlsxSource = r;
  let f = parseExternalLinkTargets(Nd(t), r.sha256);
  (f.length &&
    (d.workbookLinks = { schemaVersion: 1, sources: [], packageTargets: f }),
    a.properties.date1904 &&
      (d.datePolicy = { ...defaultDatePolicy, dateSystem: "1904" }));
  let p = 0,
    h = [];
  for (let [S, O] of n.sheetNames)
    for (let [E, R] of n.numericTokens.get(n.sheetRelationships.get(O) ?? "") ??
      []) {
      let I = validateXlsxNumericValue(R);
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
  let m = resolveDefinedNameScopes(n.names, d, [...n.sheetNames.keys()]);
  m.length && (d.definedNames = m);
  let g = n.calculation ?? {
    mode: "auto",
    iterative: false,
    maxIterations: null,
    maxChange: null,
  };
  applyCalculationProperties(d, g, i);
  let b = toValidWorkbook(
      createSourceRowProvenance(d, {
        id: createRandomUuid(),
        name: e.name,
        kind: "file",
      }),
    ),
    w = buildExchangeSchemaReport("import", e.name, t.byteLength, b, i);
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
    compareFormulaCachesWithRecalculation(
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
    w.warnings.push(...collectDataValidationWarnings(b)),
    { workbook: toValidWorkbook(b), warnings: w.warnings, report: w }
  );
}
function parseNativeWorkbookJson(e) {
  if (new TextEncoder().encode(e).length > workbookLimits.maxNativeBytes)
    throw new Error("Native workbook exceeds the 50 MB import limit.");
  let t;
  try {
    t = JSON.parse(e);
  } catch {
    throw new Error("The native workbook is not valid JSON.");
  }
  return toValidWorkbook(t);
}

function buildRetainedXlsxSchemaReport(e) {
  let t = requireRetainedXlsxSource(e),
    n = parseXlsxWorkbookPackage(t.bytes, createEmptyExportLedger()),
    i = buildExchangeSchemaReport("import", t.fileName, t.bytes.length, e, n);
  return ((i.partInventory = F$(Nd(t.bytes), t.sha256)), i);
}

function toErrorResponse(e) {
  let { code, message } = toFileProcessingError(e);
  return { code, message };
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
function parseJsonWithinBudget(e) {
  let t = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(e)),
    n = [{ value: t, depth: 0 }],
    i = 0;
  for (; n.length;) {
    let r = n.pop();
    if (++i > u1.maxJsonNodes || r.depth > u1.maxJsonDepth)
      throw new FileProcessingError("FILE_LIMIT");
    if (r.value && typeof r.value == "object")
      for (let [o, a] of Object.entries(r.value)) {
        if (["__proto__", "prototype", "constructor"].includes(o))
          throw new FileProcessingError("FILE_ACTIVE");
        n.push({ value: a, depth: r.depth + 1 });
      }
  }
  return t;
}
var workerScope = self,
  messageHandled = false;
workerScope.onmessage = async ({ data }) => {
  if (messageHandled) return;
  messageHandled = true;
  let t = data?.id;
  try {
    if (
      typeof t != "string" ||
      t.length > 100 ||
      !(data.bytes instanceof Uint8Array)
    )
      throw new FileProcessingError("FILE_INVALID");
    assertImportFileAllowed(data.operation, data.fileName, data.bytes.length);
    let n;
    switch (data.operation) {
      case "source":
        n = inspectSourceFile(data.bytes, data.fileName);
        break;
      case "formats":
        n = inspectAlternateWorkbookFile(data.bytes, data.fileName);
        break;
      case "workbook": {
        /\.json$/i.test(data.fileName) && parseJsonWithinBudget(data.bytes);
        let i = data.bytes,
          r = await importWorkbookFile({
            name: data.fileName,
            size: i.length,
            arrayBuffer: async () => i.buffer,
            text: async () =>
              new TextDecoder("utf-8", { fatal: true }).decode(i),
          });
        if (
          (/\.json$/i.test(data.fileName) &&
            r.workbook.xlsxSource &&
            buildRetainedXlsxSchemaReport(r.workbook),
          /\.json$/i.test(data.fileName) && r.workbook.alternateSource)
        ) {
          let o = r.workbook.alternateSource;
          inspectAlternateWorkbookFile(validateAlternateSource(o), o.fileName);
        }
        if (/\.json$/i.test(data.fileName))
          for (let o of r.workbook.sourceImports ?? []) {
            let a = inspectSourceFile(
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
        let i = validateDelimitedImport(
          data.bytes,
          data.fileName,
          data.options,
        );
        n = {
          report: i,
          ...(i.valid
            ? {
                workbook: importDelimitedWorkbook(
                  data.bytes,
                  data.fileName,
                  data.options,
                ).workbook,
              }
            : {}),
        };
        break;
      }
      case "automation":
        n = inspectAutomationFile(data.bytes, data.fileName);
        break;
      case "json":
        n = parseJsonWithinBudget(data.bytes);
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
        throw new FileProcessingError("FILE_TYPE");
    }
    if (new TextEncoder().encode(JSON.stringify(n)).length > u1.maxResultBytes)
      throw new FileProcessingError("FILE_LIMIT");
    workerScope.postMessage({ id: t, ok: true, value: n });
  } catch (n) {
    workerScope.postMessage({ id: t, ok: false, error: toErrorResponse(n) });
  }
};
