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

function dateToExcelSerial(e, dateSystem = "1900") {
  return (
    dateToSerial(
      e.getUTCFullYear(),
      e.getUTCMonth() + 1,
      e.getUTCDate(),
      dateSystem,
    ) +
    (e.getUTCHours() * 36e5 +
      e.getUTCMinutes() * 6e4 +
      e.getUTCSeconds() * 1e3 +
      e.getUTCMilliseconds()) /
      millisecondsPerDay
  );
}

var namedStyleLimits = { maxStyles: 50, maxCells: 500, maxName: 80 };

function normalizeWorkbookStyleCatalog(styleCatalog) {
  (bo(styleCatalog) || Fr("Workbook styles must be an object."),
    assertKnownStyleKeys(
      styleCatalog,
      ["namespace", "theme", "styles", "missingTheme"],
      "workbook style catalog",
    ),
    styleCatalog.missingTheme !== undefined &&
      styleCatalog.missingTheme !== true &&
      Fr("Invalid missing theme flag."),
    (!Array.isArray(styleCatalog.styles) ||
      styleCatalog.styles.length > namedStyleLimits.maxStyles) &&
      Fr("A workbook supports at most 50 named styles."));
  let t = new Set(),
    n = new Set(),
    styles = styleCatalog.styles.map((style) => {
      (bo(style) || Fr("Invalid named style."),
        assertKnownStyleKeys(
          style,
          ["id", "name", "definition"],
          "named style",
        ));
      let s = validateBoundedStyleText(style.id, "Style ID", 100),
        l = validateBoundedStyleText(style.name, "Style name");
      return (
        (t.has(s) || n.has(l.toLocaleLowerCase("en-US"))) &&
          Fr("Named style IDs and names must be unique within this workbook."),
        t.add(s),
        n.add(l.toLocaleLowerCase("en-US")),
        {
          id: s,
          name: l,
          definition: validateNamedStyleDefinition(style.definition),
        }
      );
    }),
    r = styleCatalog.theme === undefined || styleCatalog.missingTheme === true,
    theme =
      styleCatalog.theme === undefined
        ? cloneStyleValue(officeThemeDefinition)
        : validateWorkbookTheme(styleCatalog.theme);
  return (
    r &&
      canonicalStyleKey(theme) !== canonicalStyleKey(officeThemeDefinition) &&
      Fr("A missing theme must use the Office fallback."),
    {
      namespace: validateBoundedStyleText(
        styleCatalog.namespace,
        "Style namespace",
        100,
      ),
      theme: theme,
      styles: styles,
      ...(r ? { missingTheme: true } : {}),
    }
  );
}

function applyNamedStyleCatalog(styleCatalog, workbook) {
  let n =
    styleCatalog === undefined
      ? undefined
      : normalizeWorkbookStyleCatalog(styleCatalog);
  n ? (workbook.styleCatalog = n) : delete workbook.styleCatalog;
  let i = bo(styleCatalog) && styleCatalog.theme === undefined;
  for (let sheet of workbook.sheets)
    for (let o of hasLinkedOrBoundCells(sheet.cells)
      ? Object.values(sheet.cells)
      : []) {
      if (!o.style?.binding) continue;
      let a = validateNamedStyleBinding(o.style.binding),
        namedStyleBinding = resolveNamedStyleBinding(n, a);
      (!namedStyleBinding.missing &&
        !i &&
        canonicalStyleKey(a) !== canonicalStyleKey(namedStyleBinding) &&
        Fr(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (o.style.binding = namedStyleBinding));
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

function collectDataValidationWarnings(workbook) {
  if (!workbook.sheets.some((sheet) => sheet.validations?.length)) return [];
  let t = summarizeValidationInspection(workbook);
  return t.invalidCells || !t.complete
    ? [
        `Data validation: ${t.message} Source values were retained; use Data Validation to review.`,
      ]
    : [];
}

function resolveDefinedNameScopes(names, importedWorkbook, n) {
  if (new Set(n.map((r) => r.toLowerCase())).size !== n.length)
    return throwXlsxDefinedNameError(
      "Duplicate worksheet names make name scopes ambiguous.",
    );
  let i = names.map((r) => {
    let o = r.localSheetId === undefined ? undefined : n[r.localSheetId],
      scopeSheetId =
        o === undefined
          ? undefined
          : importedWorkbook.sheets.find((sheet) => sheet.name === o)?.id;
    return r.localSheetId !== undefined && !scopeSheetId
      ? throwXlsxDefinedNameError(
          `Name "${r.name}" has an invalid or unsupported worksheet scope.`,
        )
      : {
          name: r.name,
          refersTo: r.refersTo,
          ...(scopeSheetId ? { scopeSheetId: scopeSheetId } : {}),
          ...(r.comment !== undefined ? { comment: r.comment } : {}),
        };
  });
  for (let r of i) validateNameSheetReferences(r.refersTo, importedWorkbook);
  try {
    return validateDefinedName(i, importedWorkbook);
  } catch (error) {
    return throwXlsxDefinedNameError(error.message);
  }
}

function createSourceRowProvenance(workbook, t) {
  if (
    !isBoundedString(t.id, 100) ||
    !isBoundedString(t.name, 255) ||
    !["file", "snapshot"].includes(t.kind)
  )
    throw Error("Provide a valid source identity, name and origin kind.");
  let rowOrigins = [...(workbook.rowOrigins ?? [])],
    sheets = workbook.sheets.map((sheet) => {
      let rowOriginIds = { ...sheet.rowOriginIds },
        a = new Set(
          Object.keys(sheet.cells).map(
            (cell) => Number(/\d+$/.exec(cell)[0]) - 1,
          ),
        );
      for (let s of a)
        if (!rowOriginIds[s]) {
          if (rowOrigins.length >= workbookLimits.maxCells)
            throw Error(
              "Row provenance has reached its 100,000 source-row budget.",
            );
          let l = crypto.randomUUID();
          ((rowOriginIds[s] = l),
            rowOrigins.push({
              id: l,
              sourceId: t.id,
              sourceName: t.name.slice(0, 255),
              sourceSheet: sheet.name,
              originalRow: s + 1,
              kind: t.kind,
            }));
        }
      return { ...sheet, rowOriginIds: rowOriginIds };
    });
  return { ...workbook, sheets: sheets, rowOrigins: rowOrigins };
}

function convertImportedTables(worksheet, sheetOrThrow) {
  let n = [],
    i = [],
    r = worksheet.model.tables ?? [];
  for (let o of r) {
    if (o.headerRow === false) {
      n.push(
        `Table ${o.name} has no header row; its cells were retained without table metadata.`,
      );
      continue;
    }
    try {
      let range = Vm(o.tableRef);
      resolveSourceRange(range, sheetOrThrow);
      let s =
          Object.entries(tablePaletteStyleNames).find(
            ([, c]) => c === o.style?.theme,
          )?.[0] ?? "blue",
        totals = {};
      for (let column of o.columns) {
        let u = column.totalsRowFunction?.toLowerCase();
        if (basicAggregateOperations.includes(u))
          Object.defineProperty(totals, column.name, {
            value: u,
            enumerable: true,
            writable: true,
            configurable: true,
          });
        else if (u === "custom") {
          let d = column.totalsRowFormula ?? "",
            f = /^=?\s*(SUM|COUNTA|AVERAGE|MIN|MAX)\(/i
              .exec(d)?.[1]
              .toLowerCase();
          f &&
            Object.defineProperty(totals, column.name, {
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
        range: range,
        columns: o.columns.map((column) => ({
          id: createUuid(),
          name: validateColumnName(column.name),
        })),
        style: s,
        showTotals: !!o.totalsRow,
        totals: totals,
      }),
        Object.values(tablePaletteStyleNames).includes(o.style?.theme) ||
          n.push(
            `Table ${o.name}: source cells were preserved; unsupported table theme maps to the blue table style for future edits.`,
          ));
    } catch (error) {
      n.push(
        `Table ${o.name}: ${error.message} Cells remain available without this table definition.`,
      );
    }
  }
  return (
    i.length &&
      ((sheetOrThrow.tables = i),
      normalizeTableDefinitions(sheetOrThrow, sheetOrThrow)),
    n
  );
}

function validateReportSnapshot(snapshot) {
  if (new TextEncoder().encode(JSON.stringify(snapshot)).length > 4e6)
    throw Error("A report snapshot exceeds 4 MB.");
  let t = calculationReportSchema.parse(snapshot),
    rangeAddress = parseRangeAddress(t.request.range);
  if (
    (validateObjectReportRequest(t.request),
    t.hash !== hashRecordExcludingHash(t))
  )
    throw Error("Report snapshot integrity check failed.");
  let i = Array.from(
    { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
    (a, s) => rangeAddress.firstCol + s,
  );
  if (
    fingerprintCanonicalValue(i) !== fingerprintCanonicalValue(t.columns) ||
    t.widths.length !== i.length
  )
    throw Error("Invalid report column layout.");
  let r = t.rows.map((row) => row.row),
    o = [...r, ...t.excludedRows].sort((a, s) => a - s);
  if (
    t.rows.length * i.length > 2e3 ||
    o.length !== rangeAddress.lastRow - rangeAddress.firstRow + 1 ||
    o.some((a, s) => a !== rangeAddress.firstRow + s + 1) ||
    r.some((a, s) => s > 0 && a <= r[s - 1])
  )
    throw Error("Invalid report row accounting.");
  for (let row of t.rows) {
    if (row.cells.length !== i.length)
      throw Error("Invalid report cell layout.");
    row.cells.forEach((cell, l) => {
      if (cell.address !== `${columnIndexToLetters(i[l])}${row.row}`)
        throw Error("Invalid report cell address.");
      cell.style &&= validateCellStyle(cell.style);
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
function validateReportOutputs(workbook, t) {
  if (workbook.reportOutputs !== undefined) {
    let n = X.array(
      X.object({
        id: nonEmptyLabelSchema,
        version: X.number().int().min(1).max(1e6),
        snapshot: X.unknown(),
      }).strict(),
    )
      .max(20)
      .parse(workbook.reportOutputs);
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
  if (workbook.embeddedObjectInventories !== undefined) {
    let n = X.array(importedCacheEvidenceSchema)
      .max(20)
      .parse(workbook.embeddedObjectInventories);
    if (
      new Set(n.map((i) => i.id)).size !== n.length ||
      JSON.stringify(n).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    t.embeddedObjectInventories = n;
  }
}

function importPageSetup(worksheet, sheetOrThrow) {
  let pageSetup = worksheet.pageSetup,
    defaultPageSetup = createDefaultPageSetup(sheetOrThrow);
  if (
    pageSetup.pageOrder !== undefined &&
    pageSetup.pageOrder !== "overThenDown" &&
    pageSetup.pageOrder !== "downThenOver"
  )
    throw Error(
      "The Excel worksheet has an unsupported pageOrder value. Import stopped without changing the current workbook.",
    );
  let r = {
    ...defaultPageSetup,
    orientation:
      pageSetup.orientation === "landscape" ? "landscape" : "portrait",
    paper:
      pageSetup.paperSize === 9
        ? "a4"
        : pageSetup.paperSize === 5
          ? "legal"
          : "letter",
    margins: { ...defaultPageSetup.margins, ...pageSetup.margins },
    printArea:
      pageSetup.printArea?.split("&&")[0] || defaultPageSetup.printArea,
    fitToPage: !!pageSetup.fitToPage,
    fitToWidth: Math.max(0, Math.min(10, pageSetup.fitToWidth ?? 1)),
    fitToHeight: Math.max(0, Math.min(10, pageSetup.fitToHeight ?? 0)),
    scale: Math.max(10, Math.min(400, pageSetup.scale ?? 100)),
    gridlines: !!pageSetup.showGridLines,
    headings: !!pageSetup.showRowColHeaders,
    header: String(worksheet.headerFooter?.oddHeader ?? "").slice(0, 255),
    footer: String(worksheet.headerFooter?.oddFooter ?? "").slice(0, 255),
    pageOrder: pageSetup.pageOrder ?? "downThenOver",
    ...(pageSetup.printTitlesRow
      ? { repeatRows: pageSetup.printTitlesRow }
      : {}),
    ...(pageSetup.printTitlesColumn
      ? { repeatColumns: pageSetup.printTitlesColumn }
      : {}),
  };
  validatePageSetup(r, sheetOrThrow);
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

function validateFieldTransformEvidence(fieldTransforms, t) {
  if (fieldTransforms === undefined) return;
  if (
    !Array.isArray(fieldTransforms) ||
    fieldTransforms.length > 100 ||
    JSON.stringify(fieldTransforms).length >
      fieldTransformLimits.evidenceCharacters
  )
    throw Error(
      "Field transformation evidence exceeds its native history budget. Keep the original backup.",
    );
  let n = new Set(),
    i = 0;
  t.fieldTransforms = fieldTransforms.map((fieldTransform) => {
    let o = analysisOutputEvidenceSchema.safeParse(fieldTransform);
    if (!o.success)
      throw Error(
        `Invalid field transformation evidence: ${o.error.issues[0]?.message}`,
      );
    let data = o.data,
      request = validateFieldTransformRecipe(data.request),
      rangeAddress = parseRangeAddress(request.sourceRange);
    if (
      new Date(data.createdAt).toISOString() !== data.createdAt ||
      data.sources.length !==
        (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
          (rangeAddress.lastCol - rangeAddress.firstCol + 1)
    )
      throw Error(
        "Invalid field evidence timestamp or incomplete original source population.",
      );
    if (n.has(data.id) || (i += data.outputs.length) > 2e4)
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (n.add(data.id),
      data.sourceSheetId !== request.sourceSheetId ||
        data.outputSheetName !== request.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    let c = new Set(),
      sources = data.sources.map((source) => {
        let p = _t(source.address);
        if (
          p.row < rangeAddress.firstRow ||
          p.row > rangeAddress.lastRow ||
          p.col < rangeAddress.firstCol ||
          p.col > rangeAddress.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (c.has(source.address))
          throw Error("Duplicate original field source.");
        c.add(source.address);
        let original =
          source.original === undefined
            ? undefined
            : normalizeCellRecordInput(source.original);
        if (
          original &&
          typeof original.raw == "string" &&
          original.raw.startsWith("=") &&
          !original.literal
        )
          throw Error("Field source evidence must contain reviewed values.");
        return { ...source, ...(original ? { original: original } : {}) };
      }),
      d = new Set();
    for (let output of data.outputs) {
      if (d.has(output.address))
        throw Error("Duplicate transformed output address.");
      if (
        (d.add(output.address),
        ((output.type === "number" || output.type === "date") &&
          typeof output.value != "number") ||
          (output.type === "boolean" && typeof output.value != "boolean") ||
          (output.type === "text" && typeof output.value != "string") ||
          (output.type === "blank" && output.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (let f of [...data.outputs, ...data.exceptions])
      if (
        f.row < rangeAddress.firstRow + 1 + Number(request.header) ||
        f.row > rangeAddress.lastRow + 1 ||
        f.sources.some((source) => !c.has(source)) ||
        new Set(f.sources).size !== f.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...data, request: request, sources: sources };
  });
}

function toValidWorkbook(workbook) {
  return normalizeWorkbookDocument(workbook, false);
}
function normalizeWorkbookDocument(workbook, t) {
  if (t && isDeepFrozenValue(workbook) && $V.has(workbook)) return workbook;
  if (!Gs(workbook) || workbook.schemaVersion !== 1)
    return jt("Unsupported workbook schema. Expected schemaVersion 1.");
  let n = requireBoundedText(workbook.id, "Workbook ID", 100),
    i = requireBoundedText(workbook.name, "Workbook name", 200);
  if (!n || !i.trim()) return jt("Workbook ID and name are required.");
  let revision = P2(
    workbook.revision,
    0,
    Number.MAX_SAFE_INTEGER - 1,
    "Revision",
  );
  if (!Array.isArray(workbook.sheets) || !workbook.sheets.length)
    return jt("A workbook must contain at least one sheet.");
  let o = 0,
    a = new Set(),
    s = new Set(),
    sheets = workbook.sheets.map((sheet) => {
      if (!Gs(sheet)) return jt("Invalid sheet.");
      let h = requireBoundedText(sheet.id, "Sheet ID", 100),
        m = requireBoundedText(sheet.name, "Sheet name", 31);
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
      let rows = P2(sheet.rowCount, 1, workbookLimits.maxRows, "Row count"),
        cols = P2(sheet.colCount, 1, workbookLimits.maxCols, "Column count");
      if (!Gs(sheet.cells)) return jt("Sheet cells must be an address map.");
      let w = t && isDeepFrozenValue(sheet) ? bN.get(sheet) : undefined;
      if (w !== undefined) {
        if (((o += w), o > workbookLimits.maxCells))
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        let O = { ...sheet };
        return (bN.set(O, w), O);
      }
      let cells = {};
      for (let [O, cell] of Object.entries(sheet.cells)) {
        let R = _t(O);
        if (
          O !== `${columnIndexToLetters(R.col)}${R.row + 1}` ||
          R.row >= rows ||
          R.col >= cols
        )
          return jt(`Invalid or out-of-bounds cell: ${O}`);
        if (++o > workbookLimits.maxCells)
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        cells[O] = normalizeCellRecord(cell, t);
      }
      let S = {
        id: h,
        name: m,
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        cells: cells,
      };
      if (sheet.columnWidths !== undefined) {
        if (!Gs(sheet.columnWidths))
          return jt("Column widths must be an index map.");
        S.columnWidths = {};
        for (let [O, E] of Object.entries(sheet.columnWidths)) {
          if (
            !/^(0|[1-9]\d*)$/.test(O) ||
            Number(O) >= cols ||
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
      if (sheet.rowHeights !== undefined) {
        if (
          !Gs(sheet.rowHeights) ||
          (Object.getPrototypeOf(sheet.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(sheet.rowHeights) !== null)
        )
          return jt("Row heights must be a plain index map.");
        S.rowHeights = {};
        for (let [O, E] of Object.entries(sheet.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(O) ||
            Number(O) >= rows ||
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
      if (sheet.hidden !== undefined) {
        if (typeof sheet.hidden != "boolean")
          return jt("Sheet hidden flag must be boolean.");
        S.hidden = sheet.hidden;
      }
      return (
        validateWorksheetFeatures(sheet, S),
        validateOutlineGroups(sheet, S, { rows: rows, cols: cols }),
        validateFreezePanes(sheet, S, { rows: rows, cols: cols }),
        validateWorksheetVisibility(sheet, S),
        validateDelimitedSourceFootprint(sheet.delimitedSource, S),
        normalizeTableDefinitions(sheet, S),
        sheet.objects !== undefined && validateDrawingObjects(sheet.objects, S),
        sheet.pageLayout !== undefined &&
          validatePageSetup(sheet.pageLayout, S),
        S
      );
    });
  if (sheets.every((sheet) => sheet.hidden))
    return jt("At least one sheet must be visible.");
  if (!Array.isArray(workbook.events) || workbook.events.length > 2e3)
    return jt("Workbook history must be an array of at most 2,000 events.");
  let c = new Set(),
    u = 0,
    events = workbook.events.map((p) => {
      if (!Gs(p)) return jt("Invalid history event.");
      let h = requireBoundedText(p.id, "Event ID", 100),
        timestamp = requireBoundedText(p.timestamp, "Event timestamp", 40);
      return !h || c.has(h) || !Number.isFinite(Date.parse(timestamp))
        ? jt("History events require unique IDs and valid timestamps.")
        : (c.add(h),
          !Array.isArray(p.changes) ||
          p.changes.length > workbookLimits.maxPatches ||
          (u += p.changes.length) > workbookLimits.maxCells
            ? jt("Invalid history change count.")
            : {
                id: h,
                timestamp: timestamp,
                actor: requireBoundedText(p.actor, "Actor", 100),
                action: requireBoundedText(p.action, "Action", 200),
                changes: p.changes.map((change) => {
                  if (
                    !Gs(change) ||
                    typeof change.sheetId != "string" ||
                    !change.sheetId.length ||
                    change.sheetId.length > 100 ||
                    typeof change.address != "string"
                  )
                    return jt("Invalid history cell reference.");
                  let b = _t(change.address);
                  return change.address !==
                    `${columnIndexToLetters(b.col)}${b.row + 1}`
                    ? jt("History cell address is not canonical.")
                    : change.before === undefined && change.after === undefined
                      ? jt("History change must have a before or after value.")
                      : {
                          sheetId: change.sheetId,
                          address: change.address,
                          ...(change.before !== undefined
                            ? { before: normalizeCellRecord(change.before, t) }
                            : {}),
                          ...(change.after !== undefined
                            ? { after: normalizeCellRecord(change.after, t) }
                            : {}),
                        };
                }),
              });
    }),
    f = {
      schemaVersion: 1,
      id: n,
      name: i,
      revision: revision,
      sheets: sheets,
      events: events,
    };
  return (
    importFormulaCacheEvidence(workbook, f),
    workbook.datePolicy !== undefined &&
      (f.datePolicy = normalizeDatePolicy(workbook.datePolicy)),
    validateSavedFilterViews(workbook.savedFilterViews, f),
    validateSavedRecipeMetadata(workbook, f),
    verifyConsolidationEvidence(workbook, f),
    buildExternalSourceRefreshState(workbook, f),
    validateDataModelReport(workbook, f),
    validateAnalysisReport(workbook, f),
    validateReportOutputs(workbook, f),
    normalizeRowProvenance(workbook, f),
    validateAutomationArtifacts(workbook, f),
    readRetainedXlsxSource(workbook, f),
    normalizeAlternateFormatSource(workbook, f),
    validateSourceImports(workbook, f),
    attachWorkbookLinkMetadata(workbook, f),
    normalizeTextCleanups(workbook.textCleanups, f),
    validateFieldTransformEvidence(workbook.fieldTransforms, f),
    validateTableSourceEvidence(workbook.tableEvidence, f),
    workbook.whatIfAnalyses !== undefined &&
      (f.whatIfAnalyses = validateSavedWhatIfAnalyses(workbook.whatIfAnalyses)),
    workbook.definedNames !== undefined &&
      (f.definedNames = validateDefinedName(workbook.definedNames, f)),
    validateExtensionState(workbook, f),
    validateWorkbookStructureProtection(workbook.structureProtection, f),
    validateWorkbookMetadata(workbook.metadata, f),
    assertUniqueTableNames(f),
    normalizeReviewMetadata(workbook, f),
    applyNamedStyleCatalog(workbook.styleCatalog, f),
    workbook.formulaWatches !== undefined &&
      (f.formulaWatches = validateFormulaWatches(workbook.formulaWatches, f)),
    dR(workbook, f),
    copyExtensionCacheEntries(workbook, f),
    f
  );
}

var bytesToBase64 = (e) => {
  let t = "";
  for (let n of e) t += String.fromCharCode(n);
  return btoa(t);
};
function collectImportedPictures(e, worksheet, n, i = {}) {
  let r = [],
    o = [],
    a = 0,
    s = 0;
  for (let l of worksheet.getImages()) {
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
      let dataURL = c.base64?.startsWith("data:")
          ? c.base64
          : c.base64
            ? `data:image/${u};base64,${c.base64}`
            : `data:image/${u};base64,${bytesToBase64(new Uint8Array(c.buffer))}`,
        h = inspectPictureDataUrl(dataURL),
        m = Math.floor(l.range.tl.col),
        g = Math.floor(l.range.tl.row),
        anchoredDrawingObject = createAnchoredDrawingObject(
          "picture",
          `${columnIndexToLetters(m)}${g + 1}`,
        );
      ((anchoredDrawingObject.offsetX = Math.max(
        0,
        (l.range.tl.col - m) * (n.columnWidths?.[m] ?? 84),
      )),
        (anchoredDrawingObject.offsetY = Math.max(
          0,
          (l.range.tl.row - g) * (n.rowHeights?.[g] ?? 21),
        )),
        (anchoredDrawingObject.width = Math.max(
          24,
          Math.min(4e3, l.range.ext?.width ?? 240),
        )),
        (anchoredDrawingObject.height = Math.max(
          24,
          Math.min(4e3, l.range.ext?.height ?? 140),
        )),
        (anchoredDrawingObject.dataURL = dataURL),
        (anchoredDrawingObject.evidence = createPictureEvidence(dataURL, {
          origin: "xlsx-media",
          ...(i.containerFilename !== undefined
            ? { containerFilename: i.containerFilename }
            : {}),
          mediaReference: `imageId:${l.imageId}`,
        })),
        (anchoredDrawingObject.label = `Picture ${r.length + 1}`),
        r.push(anchoredDrawingObject),
        (a += h.bytes.length));
    } catch (error) {
      o.push(
        error instanceof Error
          ? error.message
          : "An image could not be imported.",
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
function describeStructuredValue(e, path) {
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
    path: path,
    confidence: 1,
  };
}
function collectJsonSourceTables(e) {
  let t = [],
    n = (i, path) => {
      if (Array.isArray(i)) i.length && t.push({ value: i, path: path });
      else if (yg(i))
        for (let [o, a] of Object.entries(i))
          n(a, path + "/" + escapeJsonPointerSegment(o));
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
      columns: l.map((path) => ({ name: path || "Value", path: path })),
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
    for (let attribute of Object.values(u.attributes))
      if (attribute.uri !== "http://www.w3.org/2000/xmlns/")
        if (
          attribute.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          attribute.local === "type"
        ) {
          let [g, b] = attribute.value.includes(":")
            ? attribute.value.split(":")
            : ["", attribute.value];
          if (g && !h.namespaces[g])
            throw Error("XML type prefix is undeclared.");
          h.type = o(h.namespaces[g] ?? "", b);
        } else
          attribute.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          attribute.local === "nil"
            ? (h.nil = attribute.value)
            : h.attributes.push({
                name: o(attribute.uri, attribute.local),
                value: attribute.value,
                uri: attribute.uri,
                local: attribute.local,
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
      for (let [f, nodes] of d)
        nodes.length > 1
          ? s.push({ path: u.path + "/" + f, nodes: nodes })
          : l(nodes[0]);
    };
  if (
    (l(n[0]),
    s.length || s.push({ path: n[0].path, nodes: [n[0]] }),
    s.length > sourceImportLimits.tables)
  )
    throw Error("XML table count exceeded.");
  let c = (u) => {
    let text = u.text,
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
        ((text = null), (f = "null"));
      }
    }
    if (
      f !== "null" &&
      u.type?.startsWith("{http://www.w3.org/2001/XMLSchema}")
    ) {
      if (h === "boolean") {
        if (!["true", "1", "false", "0"].includes(m))
          throw Error("Invalid typed XML boolean.");
        ((text = m === "true" || m === "1"), (f = "boolean"));
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
          ((text = Number(m)),
          !b.test(m) ||
            !Number.isFinite(text) ||
            (Number.isInteger(text) && !Number.isSafeInteger(text)) ||
            (h === "nonNegativeInteger" && text < 0) ||
            (h === "positiveInteger" && text <= 0))
        )
          throw Error("Invalid or unsafe typed XML number.");
        let w = {
          int: [-2147483648, 2147483647],
          short: [-32768, 32767],
          byte: [-128, 127],
        };
        if (w[h] && (text < w[h][0] || text > w[h][1]))
          throw Error("Typed XML integer is out of range.");
        f = "number";
      }
    }
    return {
      value: text,
      type: f,
      path: u.path,
      lexical: p,
      ...(u.type ? { declaredType: u.type } : {}),
      confidence: 1,
    };
  };
  return s.map((u, d) => {
    let f = new Set(),
      p = u.nodes.map((node) => {
        let g = new Map(),
          b = (_, S) => {
            (f.add(_), g.set(_, S));
          },
          w = (_, S) => {
            for (let attribute of _.attributes)
              b(S + "/@" + attribute.name, {
                value: attribute.value,
                type: "string",
                path: _.path + "/@" + attribute.name,
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
        return (w(node, ""), g);
      }),
      h = [...f];
    return {
      id: "xml-" + d,
      name: u.path.slice(-200),
      path: u.path,
      ambiguous: s.length > 1,
      columns: h.map((path) => ({
        name: (path || "Value").slice(-200),
        path: path,
      })),
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
function inspectSourceFile(bytes, fileName) {
  if (bytes.length < 1 || bytes.length > sourceImportLimits.bytes)
    throw Error("Source byte budget exceeded.");
  let n = fileName.split(".").at(-1)?.toLowerCase(),
    format = n === "jpg" ? "jpeg" : n === "htm" ? "html" : n,
    tables = [],
    method = "document-adapter",
    extractor = "Host extraction required",
    warnings = [];
  if (format === "json" || format === "xml") {
    let c = new TextDecoder("utf-8", { fatal: true })
      .decode(bytes)
      .replace(/^\uFEFF/, "");
    ((tables =
      format === "json"
        ? collectJsonSourceTables(parseStrictJsonText(c))
        : parseXmlSourceTables(c)),
      (method = format === "json" ? "native-json" : "native-xml"),
      (extractor = "Ledgerglass structured source v1"),
      warnings.push(
        "String identifiers and untyped XML remain text. Nested JSON arrays/empty objects are retained as JSON values. Imported formula-like strings remain literal.",
      ));
  } else if (format === "pdf") {
    if (
      new TextDecoder()
        .decode(bytes.subarray(0, 8))
        .match(/^%PDF-1\.[0-7]|^%PDF-2\.0/) == null
    )
      throw Error("Source extension does not match PDF content.");
    warnings.push(
      "PDF bytes are retained without rendering or executing embedded content. A host extractor or matching extraction JSON is required.",
    );
  } else if (format === "png" || format === "jpeg") {
    let c = "";
    for (let u = 0; u < bytes.length; u += 8192)
      c += String.fromCharCode(...bytes.subarray(u, u + 8192));
    (inspectPictureDataUrl(`data:image/${format};base64,${btoa(c)}`),
      warnings.push(
        "Image extraction requires a host OCR adapter or matching extraction JSON.",
      ));
  } else if (format === "html") {
    let c = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    if (!/^\s*(?:<!doctype\s+html[^>]*>\s*)?<(?:html|table)\b/i.test(c))
      throw Error("Source extension does not match an HTML document.");
    warnings.push(
      "HTML is retained as inert bytes. No markup, scripts or resource links are rendered or fetched. A host table extractor or matching extraction JSON is required.",
    );
  } else throw Error("Choose JSON, XML, PDF, PNG, JPEG or HTML.");
  let l = validateExtractedSourceTables({
    schemaVersion: 1,
    fileName: fileName,
    format: format,
    byteLength: bytes.length,
    sha256: sha256HexDigest(bytes),
    tables: tables,
    method: method,
    extractor: extractor,
    warnings: warnings,
  });
  return (recordSourceInspection(l, bytes, fileName), l);
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

function buildAlternateFormatGuidance(format) {
  return [
    "Keep the original file. Inspection and conversion here never execute macros, objects or external refreshes.",
    format === "ods"
      ? "In a trusted source application that supports ODS, open the original and Save As an unencrypted Excel Workbook (.xlsx)."
      : "In a trusted source application that supports this format, open the original with macros and external refresh disabled. Save As an unencrypted Excel Workbook (.xlsx).",
    "Review the source application's compatibility warnings. Keep any needed automation, objects and unsupported metadata in the original.",
    "Choose that converted XLSX here. Compare formulas, date systems, dates, styles and special parts with the original; confirm the loss acknowledgement and review the XLSX compatibility report before replacing the workbook.",
    "For output in the original format, export a reviewed derived XLSX, open it in the source application and Save As the target format. Reopen there and verify the result. Renaming an extension does not convert a workbook.",
  ];
}

function inspectAlternateWorkbookFile(e, fileName) {
  if (
    !fileName ||
    fileName.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(fileName) ||
    !e.length ||
    e.length > workbookLimits.maxFileBytes
  )
    throw Error(
      "Invalid alternate workbook filename or 10 MB file limit exceeded.",
    );
  let format = fileName.split(".").pop().toLowerCase();
  if (!alternateFormatExtensions.includes(format))
    throw Error("Choose an XLS, XLSB, XLSM, XLTX, XLTM or ODS source.");
  let i, identity, parts;
  if (e[0] === 208 && e[1] === 207) {
    let s = inspectCompoundFileContainer(e);
    ((i = "xls"),
      (identity =
        s.version + " workbook in a CFB container; binary cells not decoded"),
      (parts = s.paths.map((path) => ({
        path: path,
        bytes: /^(Workbook|Book)$/i.test(path) ? s.workbookBytes : 0,
        disposition: "original-only",
      }))));
  } else {
    let s = inspectAlternateWorkbookArchive(e);
    ((i = readOpenDocumentPackage(s)),
      (identity =
        i === "ods"
          ? "ODF spreadsheet MIME, manifest and content"
          : "OOXML main content type and package relationship"));
    for (let [l, c] of Object.entries(s))
      /\.(xml|rels)$/i.test(l) && parseAlternateWorkbookXml(c, () => {});
    parts = Object.entries(s).map(([path, c]) => ({
      path: path,
      bytes: c.length,
      sha256: sha256HexDigest(c),
      disposition: "original-only",
    }));
  }
  if (i !== format)
    throw Error(
      "The filename extension does not match the workbook content. Restore its original extension or convert it in the source application; renaming is not conversion.",
    );
  let a = {
    schemaVersion: 1,
    fileName: fileName,
    format: format,
    fileBytes: e.length,
    sha256: sha256HexDigest(e),
    identity: identity,
    execution: "never",
    parts: parts,
    steps: buildAlternateFormatGuidance(format),
    warnings: [
      "Identity inspection is not workbook conversion, malware scanning or a guarantee of formula, date or style equivalence.",
      "All source parts are retained only in the exact original. No macros, binary formulas, embedded objects or external links execute during inspection.",
      format === "xltx"
        ? "The native template adapter uses the supported XLSX subset. Active or unsupported content may refuse import; review conversion losses."
        : "An external source application is required for editable import and target-format export. No external conversion has been performed here.",
      ...(format === "xls"
        ? [
            "CFB inventory lists directory names; non-workbook stream sizes, nested paths and cell contents are not resolved.",
          ]
        : []),
    ],
  };
  return (validateIsolatedFormatInspection(a, e, fileName), a);
}
var saxesEsModuleForLinks = toEsModule(getSaxesModule(), 1);
function parseExternalLinkTargets(e, sha256) {
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
          (attribute) => attribute.local === "id",
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
      let part = i.get(a);
      if (!part) return;
      let c = part.lastIndexOf("/"),
        u = part.slice(0, c + 1) + "_rels/" + part.slice(c + 1) + ".rels",
        target;
      (n(u, (f) => {
        if (
          f.local === "Relationship" &&
          f.attributes.Type?.value.endsWith("/externalLinkPath") &&
          f.attributes.TargetMode?.value === "External"
        ) {
          if (target !== undefined)
            throw Error("Ambiguous external workbook target relationship.");
          target = f.attributes.Target?.value;
        }
      }),
        target &&
          o.push({
            key: `[${s + 1}]`,
            target: target,
            part: part,
            sourceSha256: sha256,
          }));
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
    convertedFileName =
      e.name.replace(/\.xltx$/i, "").slice(0, 236) + ".template-copy.xlsx",
    o = await importWorkbookFile({
      name: convertedFileName,
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
        convertedFileName: convertedFileName,
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
function collapseCellAddressesToRanges(addresses) {
  let t = addresses.map(_t).sort((o, a) => o.row - a.row || o.col - a.col),
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
    min = ["list", "custom"].includes(t) ? 0 : r(i[0]),
    max = ["list", "custom"].includes(t)
      ? 0
      : ["between", "notBetween"].includes(n)
        ? r(i[1])
        : min;
  if (!Number.isFinite(min) || !Number.isFinite(max)) return null;
  let list = [],
    formula;
  if (
    (t === "list" &&
      (typeof i[0] != "string" ||
        !/^"[^"]*"$/.test(i[0]) ||
        ((list = i[0].slice(1, -1).split(",")),
        !list.length ||
          list.length > 200 ||
          list.some((u) => !u.length || u.length > 1e3)))) ||
    (t === "custom" &&
      (typeof i[0] != "string" ||
        ((formula = "=" + i[0].replace(/^=/, "")),
        formula.length > 8192 ||
          formula.length < 2 ||
          isExternalReferenceFormula(formula))))
  )
    return null;
  let errorStyle = e.errorStyle?.toLowerCase() ?? "stop";
  return ["stop", "warning", "information"].includes(errorStyle)
    ? {
        kind: t,
        operator: n ?? "equal",
        min: min,
        max: max,
        list: list,
        allowBlank: !!e.allowBlank,
        message: String(e.error ?? "Choose a valid value.").slice(0, 500),
        errorStyle: errorStyle,
        errorTitle: String(e.errorTitle ?? "Invalid value").slice(0, 100),
        showInputMessage: !!e.showInputMessage,
        inputTitle: String(e.promptTitle ?? "").slice(0, 100),
        inputMessage: String(e.prompt ?? "").slice(0, 500),
        ...(formula ? { formula: formula } : {}),
      }
    : null;
}

function buildConditionalFormatRule(e, t) {
  let fill =
      e.style?.fill?.type === "pattern"
        ? argbColorOrFallback(e.style.fill.fgColor, "#fff2cc")
        : "#fff2cc",
    color = argbColorOrFallback(e.style?.font?.color, "#000000"),
    r = {
      id: GN(),
      priority: Number.isFinite(e.priority) ? e.priority : 1e4,
      range: formatRangeCornerCells(t),
      value: "",
      second: "",
      fill: fill,
      color: color,
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
    let rangeAddress = parseRangeAddress(t),
      u = Ks(rangeAddress.firstRow, rangeAddress.firstCol),
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
      value = u.value,
      p = d.value;
    return !(c.type === "min" || (c.type === "percent" && c.value === 0)) ||
      typeof value != "number" ||
      typeof p != "number" ||
      !Number.isFinite(value) ||
      !Number.isFinite(p) ||
      value >= p ||
      (u.type === "percent" && (value < 0 || p > 100))
      ? null
      : {
          ...r,
          kind: "iconSet",
          thresholdType: u.type === "num" ? "number" : "percent",
          iconThresholds: [value, p],
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
        rangeAddress = parseRangeAddress(t),
        cellAddress = Ks(rangeAddress.firstRow, rangeAddress.firstCol);
      if (
        h.length <= 1e3 &&
        c.toUpperCase() ===
          buildContainsTextFormula(h, cellAddress).toUpperCase()
      )
        return { ...r, kind: "contains", value: h };
    }
    let rangeAddress = parseRangeAddress(t),
      f = Ks(rangeAddress.firstRow, rangeAddress.firstCol),
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
    second = a === "between" ? o(e.formulae?.[1]) : "";
  return !a || s === null || second === null
    ? null
    : { ...r, kind: a, value: s, second: second };
}
function convertWorksheetFeatures(worksheet, t, n = {}, i = []) {
  let r = [],
    sheet = {},
    a = worksheet.model.merges ?? [];
  a.length && (sheet.merges = a.map(formatRangeCornerCells));
  let model = worksheet.dataValidations.model,
    l = new Map();
  for (let [h, rule] of Object.entries(model)) {
    let g = _t(h),
      area = i.find((_) => {
        let rangeAddress = parseRangeAddress(_.range);
        return (
          g.row >= rangeAddress.firstRow &&
          g.row <= rangeAddress.lastRow &&
          g.col >= rangeAddress.firstCol &&
          g.col <= rangeAddress.lastCol
        );
      }),
      w = JSON.stringify({ rule: rule, area: area });
    (l.has(w) || l.set(w, { rule: rule, addresses: [], area: area }),
      l.get(w).addresses.push(h));
  }
  let c = [];
  for (let { rule, addresses, area } of l.values()) {
    let b = !["list", "custom"].includes(rule.type),
      w = area?.operands.every((operand) =>
        /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(operand),
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
    for (let range of collapseCellAddressesToRanges(addresses)) {
      if (c.length >= 100) {
        r.push(
          "Additional validation areas beyond the 100-rule limit were omitted; source cells were retained.",
        );
        break;
      }
      let O = _t(area?.anchor ?? range.split(":")[0]),
        E = _t(range.split(":")[0]);
      c.push({
        ..._,
        id: GN(),
        range: range,
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
  c.length && (sheet.validations = c);
  let u = [],
    d = new Set(),
    f = 0;
  for (let h of worksheet.conditionalFormattings ?? []) {
    let m = h.ref.trim().split(/\s+/).filter(Boolean);
    if (!m.length)
      throw Error("A conditional formatting rule requires an applied range.");
    let rangeAddress = parseRangeAddress(m[0]);
    for (let rule of h.rules) {
      if (
        !Number.isSafeInteger(rule.priority) ||
        rule.priority < 1 ||
        d.has(rule.priority)
      )
        throw Error("Invalid or duplicate conditional-format priority.");
      d.add(rule.priority);
      for (let w of m) {
        let _ = parseRangeAddress(w),
          S = rule;
        (rule.type === "expression" || rule.type === "containsText") &&
          rule.formulae?.length &&
          (S = {
            ...rule,
            formulae: rule.formulae.map((E) =>
              typeof E == "string"
                ? shiftFormulaReferences(
                    "=" + E.replace(/^=/, ""),
                    _.firstRow - rangeAddress.firstRow,
                    _.firstCol - rangeAddress.firstCol,
                  ).slice(1)
                : E,
            ),
          });
        let conditionalFormatRule = buildConditionalFormatRule(S, w);
        if (
          conditionalFormatRule &&
          (conditionalFormatRule.kind === "duplicate" ||
            conditionalFormatRule.kind === "unique")
        ) {
          let E = true;
          for (let [R, cell] of Object.entries(t.cells)) {
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
            let raw = cell.raw;
            if (
              (typeof cell.raw == "string" &&
                cell.raw.startsWith("=") &&
                !cell.literal &&
                ((raw = worksheet.getCell(R).result),
                raw === undefined &&
                  /^=\s*""\s*$/.test(cell.raw) &&
                  (raw = "")),
              !Hle(raw))
            ) {
              E = false;
              break;
            }
          }
          E ||
            ((conditionalFormatRule = {
              ...conditionalFormatRule,
              kind: "formula",
              formula: "=" + String(S.formulae?.[0] ?? "").replace(/^=/, ""),
            }),
            r.push(
              "A duplicate/unique condition was retained as a formula because its population cannot use the native literal comparison; formula-engine limits apply.",
            ));
        }
        conditionalFormatRule && u.length < 100
          ? u.push({
              ...conditionalFormatRule,
              ...(n[String(rule.priority)] !== undefined
                ? { stopIfTrue: n[String(rule.priority)] }
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
    sheet.conditionalFormats = u.map((g) => ({ ...g, priority: m.get(g.id) }));
  }
  return (
    worksheet.model.sheetProtection?.sheet &&
      r.push(
        "Source Excel sheet-password protection is not imported. The local protection feature and host authorization are separate controls.",
      ),
    validateWorksheetFeatures(sheet, t),
    r
  );
}

function applyCalculationProperties(
  importedWorkbook,
  sourceSettings,
  xlsxWorkbookPackage,
) {
  ((importedWorkbook.calculation = {
    ...defaultCalculationSettings,
    mode: sourceSettings.mode === "manual" ? "manual" : "automatic",
    iterative: false,
    maxIterations:
      sourceSettings.maxIterations !== null &&
      sourceSettings.maxIterations <= 1e3
        ? sourceSettings.maxIterations
        : defaultCalculationSettings.maxIterations,
    maxChange:
      sourceSettings.maxChange !== null && sourceSettings.maxChange <= 1
        ? sourceSettings.maxChange
        : defaultCalculationSettings.maxChange,
  }),
    sourceSettings.fullPrecision === false &&
      xlsxWorkbookPackage.push(
        "The source used precision as displayed. This application calculates with full stored precision and never rounds cells automatically. Source rounding may already have changed stored values irreversibly; review formula results and use explicit ROUND where required. The original setting is retained in imported calculation evidence.",
      ),
    sourceSettings.mode === "manual" &&
      xlsxWorkbookPackage.push(
        "Source Manual calculation mode is retained. Formulas remain #PENDING! until Calculate Now; imported caches are historical evidence only.",
      ),
    sourceSettings.mode === "autoNoTable" &&
      xlsxWorkbookPackage.push(
        "Excel Automatic except data tables is imported as Automatic. Native Excel data-table calculation is not supported; source mode is retained as evidence.",
      ),
    sourceSettings.iterative &&
      xlsxWorkbookPackage.push(
        "Source iterative calculation was enabled. It is disabled locally until you explicitly enable it in Calculation Options and review numeric convergence.",
      ),
    ((sourceSettings.maxIterations ?? 100) > 1e3 ||
      (sourceSettings.maxChange ?? 0.001) > 1) &&
      xlsxWorkbookPackage.push(
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
function inspectAutomationArchive(bytes) {
  let t = inspectXlsxPackage(bytes, { allowBinaryWorkbookInventory: true }),
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
    a.push(bytes, true);
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
function inspectAutomationFile(bytes, fileName) {
  if (
    (requireAutomationFileName(fileName),
    !ArrayBuffer.isView(bytes) ||
      Object.prototype.toString.call(bytes) !== "[object Uint8Array]" ||
      !bytes.length ||
      bytes.length > automationLimits.maxFileBytes)
  )
    throw Error("Automation files must contain data and be 10 MB or smaller.");
  let n = bytes.slice(),
    i = {
      schemaVersion: 1,
      fileName: fileName,
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
    r = fileName.split(".").at(-1)?.toLowerCase();
  if (["bas", "vba", "cls", "ts", "osts"].includes(r ?? "")) {
    if (n.length > automationLimits.maxSourceBytes)
      throw Error("Text automation source is limited to 256 KiB.");
    let o = decodeUtf8Strict(n),
      language = r === "ts" || r === "osts" ? "office-script" : "vba";
    if (
      ((i.kind = r === "osts" ? "office-script-container" : language),
      r === "osts")
    )
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
            path: fileName,
            kind: "Office Script container",
            bytes: n.length,
            sha256: i.sha256,
            execution: "unsupported",
            note: "Unknown source envelope; no code inferred.",
          }));
      }
    i.items.length ||
      ((i.script = { language: language, text: o }),
      (i.execution = "reviewed-migration-only"),
      i.items.push({
        path: fileName,
        kind: language === "vba" ? "VBA source" : "Office Script source",
        bytes: n.length,
        sha256: i.sha256,
        execution: "migration-candidate",
        note: "Compatibility scan required. Only literal range operations are eligible; this is not an equivalent VBA or Office Scripts runtime.",
      }));
  } else if (n[0] === 208 && n[1] === 207 && n[2] === 17 && n[3] === 224)
    ((i.kind = "legacy-container"),
      i.items.push({
        path: fileName,
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
    for (let [path, c] of a) {
      let u = o.get(path);
      i.items.push({
        path: path,
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

function toTypedCellValue(e, dateSystem) {
  if (e != null) {
    if (e instanceof Date)
      return { type: "number", value: dateToExcelSerial(e, dateSystem) };
    if (typeof e == "object" && "error" in e && typeof e.error == "string")
      return { type: "error", value: e.error };
    if (typeof e == "string") return { type: "text", value: e };
    if (typeof e == "boolean") return { type: "boolean", value: e };
    if (typeof e == "number" && Number.isFinite(e))
      return { type: "number", value: e };
  }
}
function typedValuesMatch(e, calculated) {
  return (
    e.type === calculated.type &&
    (typeof e.value == "number" && typeof calculated.value == "number"
      ? Math.abs(e.value - calculated.value) <=
        1e-10 * Math.max(1, Math.abs(e.value), Math.abs(calculated.value))
      : e.value === calculated.value)
  );
}
function compareFormulaCachesWithRecalculation(
  e,
  validWorkbook,
  exchangeSchemaReport,
  i = new Map(),
  r,
  o = new Map(),
) {
  let workbookWithSettings = calculateWorkbookWithSettings(validWorkbook);
  for (let sheet of validWorkbook.sheets) {
    let l = e.getWorksheet(sheet.name);
    for (let [address, cell] of Object.entries(sheet.cells)) {
      let value = l.getCell(address).value;
      if (
        !value ||
        typeof value != "object" ||
        !("formula" in value || "sharedFormula" in value)
      )
        continue;
      let f = i.get(sheet.name)?.has(address)
          ? { type: "text", value: "" }
          : o.get(sheet.name)?.has(address)
            ? { type: "number", value: Number(o.get(sheet.name).get(address)) }
            : toTypedCellValue(
                l.getCell(address).result,
                validWorkbook.datePolicy?.dateSystem ?? "1900",
              ),
        calculated = workbookWithSettings.getTypedValue(sheet.id, address);
      r &&
        r.entries.push({
          sheetId: sheet.id,
          address: address,
          formula: String(cell.raw),
          ...(f ? { cache: f } : {}),
        });
      let h = cell.literal
        ? "neutralized"
        : calculated.type === "error" && calculated.value === "#PENDING!"
          ? "pending"
          : f
            ? typedValuesMatch(f, calculated)
              ? "matched"
              : "different"
            : "missing";
      (exchangeSchemaReport.caches.total++,
        (exchangeSchemaReport.caches[h] =
          (exchangeSchemaReport.caches[h] ?? 0) + 1));
      let m =
        h === "matched"
          ? exchangeSchemaReport.caches.matches
          : exchangeSchemaReport.caches.issues;
      m.length < 100
        ? m.push({
            sheet: sheet.name,
            address: address,
            ...(f ? { source: f } : {}),
            calculated: calculated,
            status: h,
          })
        : exchangeSchemaReport.caches.detailsOmitted++;
    }
  }
  (exchangeSchemaReport.caches.pending &&
    exchangeSchemaReport.warnings.push(
      `${exchangeSchemaReport.caches.pending} formula cache comparison(s) are pending. Use Calculate Now and download calculation evidence for a current comparison.`,
    ),
    exchangeSchemaReport.caches.different &&
      exchangeSchemaReport.warnings.push(
        `${exchangeSchemaReport.caches.different} source formula cache(s) differ from local recalculation. Review the comparison before relying on results; source caches never override local results.`,
      ),
    exchangeSchemaReport.caches.missing &&
      exchangeSchemaReport.warnings.push(
        `${exchangeSchemaReport.caches.missing} formula(s) have no usable source cache. Local results are not independently verified Excel results.`,
      ),
    exchangeSchemaReport.caches.neutralized &&
      exchangeSchemaReport.warnings.push(
        `${exchangeSchemaReport.caches.neutralized} external/action formula(s) were converted to literal text.`,
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

function decodeDelimitedBytes(bytes, encoding) {
  if (bytes.length > workbookLimits.maxFileBytes)
    throw Error("Delimited files must be 10 MB or smaller.");
  let n =
    bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191
      ? "utf-8"
      : bytes[0] === 255 && bytes[1] === 254
        ? "utf-16le"
        : bytes[0] === 254 && bytes[1] === 255
          ? "utf-16be"
          : undefined;
  if (n && encoding !== "auto" && encoding !== n)
    throw Error(`File BOM indicates ${n}; select that encoding or Automatic.`);
  let i = encoding === "auto" ? (n ?? "utf-8") : encoding,
    r;
  try {
    r = (
      i === "windows-1252"
        ? Array.from(bytes, (o) => byteCharacterTable[o]).join("")
        : new TextDecoder(i, { fatal: true, ignoreBOM: true }).decode(bytes)
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
function tokenizeDelimitedText(e, delimitedOptions, delimitedImportReport, i) {
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
            delimitedImportReport,
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
            delimitedImportReport,
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
      delimitedOptions.escape === "backslash" &&
      p === "\\" &&
      (e[f + 1] === delimitedOptions.quote || e[f + 1] === "\\")
        ? (o += e[++f])
        : p === delimitedOptions.quote
          ? delimitedOptions.escape === "double" &&
            e[f + 1] === delimitedOptions.quote
            ? ((o += p), f++)
            : (a = "closed")
          : (o += p);
    else if (p === delimitedOptions.delimiter) u();
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
      (delimitedOptions.newline !== "auto" &&
        delimitedOptions.newline !== h &&
        recordDelimitedIssue(
          delimitedImportReport,
          s,
          undefined,
          "Record newline does not match the selected dialect.",
        ),
        d(f),
        h.length === 2 && f++,
        (l = f + 1));
    } else
      a === "start" &&
      delimitedOptions.quote !== "none" &&
      p === delimitedOptions.quote
        ? (a = "quoted")
        : (a === "closed"
            ? (recordDelimitedIssue(
                delimitedImportReport,
                s,
                r.length + 1,
                "Unexpected text after a closing quote.",
              ),
              (a = "plain"))
            : delimitedOptions.quote !== "none" &&
              p === delimitedOptions.quote &&
              recordDelimitedIssue(
                delimitedImportReport,
                s,
                r.length + 1,
                "Quote inside an unquoted field; quote the entire field.",
              ),
          (o += p),
          (a = "plain"));
    o.length > fileImportLimits.maxCellTextLength &&
      (recordDelimitedIssue(
        delimitedImportReport,
        s,
        r.length + 1,
        "A field exceeds the 32,767-character cell limit.",
      ),
      (c = true));
  }
  (!c &&
    a === "quoted" &&
    recordDelimitedIssue(
      delimitedImportReport,
      s,
      r.length + 1,
      "Quoted field is not closed.",
    ),
    !c && l < e.length && d(e.length),
    (delimitedImportReport.complete = !c));
}

function parseDelimitedRecords(bytes, fileName, n, i) {
  let delimitedOptions = normalizeDelimitedOptions(n),
    o = decodeDelimitedBytes(bytes, delimitedOptions.encoding);
  ((delimitedOptions.encoding = o.encoding),
    delimitedOptions.delimiter === "auto" &&
      (delimitedOptions.delimiter = /\.tsv$/i.test(fileName)
        ? "	"
        : papaparseEsModule.default.parse(o.text.slice(0, 65536), {
            preview: 10,
            quoteChar:
              delimitedOptions.quote === "none" ? "" : delimitedOptions.quote,
            escapeChar:
              delimitedOptions.escape === "double"
                ? delimitedOptions.quote
                : "\\",
            delimitersToGuess: [",", ";", "	", "|"],
          }).meta.delimiter || ","));
  let delimitedImportReport = buildDelimitedImportReport(
      "import",
      fileName,
      bytes.length,
      delimitedOptions,
    ),
    cells = {};
  (delimitedImportReport.warnings.push(
    "Delimited data contains values only. All imported formula-like content stays literal text and is never treated as a formula. Auto keeps leading-zero/long identifiers and dates as text; select explicit types for localized numbers, dates and booleans.",
  ),
    (n.encoding === undefined || n.encoding === "auto") &&
      delimitedImportReport.warnings.push(
        `Encoding: ${delimitedOptions.encoding}; automatic mode uses a BOM when present and UTF-8 otherwise. Legacy encodings require an explicit selection.`,
      ),
    (n.delimiter === undefined || n.delimiter === "auto") &&
      delimitedImportReport.warnings.push(
        `Detected delimiter: ${delimitedOptions.delimiter === "	" ? "tab" : delimitedOptions.delimiter}. Review the schema and column count before importing.`,
      ));
  let l = 0;
  if (
    (tokenizeDelimitedText(
      o.text,
      delimitedOptions,
      delimitedImportReport,
      (c, record, d) => {
        ((delimitedImportReport.records = record),
          c.every((f) => f === "") && delimitedImportReport.blankRecords++,
          !l && !d && (l = c.length),
          (delimitedImportReport.columns = Math.max(
            delimitedImportReport.columns,
            c.length,
          )),
          !d &&
            l &&
            c.length !== l &&
            recordDelimitedIssue(
              delimitedImportReport,
              record,
              undefined,
              `Malformed row: expected ${l} fields, found ${c.length}.`,
            ));
        for (let f = 0; f < c.length; f++) {
          let raw = c[f],
            h =
              delimitedOptions.hasHeader && record === 1
                ? "text"
                : (delimitedOptions.columns[f] ?? "auto"),
            m = `${columnIndexToLetters(f)}${record}`;
          delimitedImportReport.schema[f] ||= {
            index: f,
            label:
              delimitedOptions.hasHeader && record === 1
                ? raw
                : columnIndexToLetters(f),
            type: delimitedOptions.columns[f] ?? "auto",
            observed: [],
          };
          let cell;
          try {
            cell = parseDelimitedCellValue(raw, h, delimitedOptions);
          } catch (error) {
            (recordDelimitedIssue(
              delimitedImportReport,
              record,
              f + 1,
              error instanceof Error
                ? error.message
                : "Type conversion failed.",
              raw,
            ),
              (cell = {
                raw: raw,
                ...(raw.startsWith("=") ? { literal: true } : {}),
              }));
          }
          if (!(delimitedOptions.hasHeader && record === 1)) {
            let b =
              cell.raw === null
                ? "blank"
                : cell.format === "date"
                  ? "date"
                  : typeof cell.raw;
            delimitedImportReport.schema[f].observed.includes(b) ||
              delimitedImportReport.schema[f].observed.push(b);
          }
          if (cell.raw !== null && !(d && delimitedOptions.empty === "blank")) {
            if (++delimitedImportReport.storedCells > workbookLimits.maxCells)
              return (
                recordDelimitedIssue(
                  delimitedImportReport,
                  record,
                  f + 1,
                  "Delimited data exceeds 100,000 populated cells.",
                ),
                false
              );
            i && (cells[m] = cell);
          }
          (raw !== "" &&
            (cell.raw !== raw ||
              (delimitedOptions.trim && raw !== raw.trim())) &&
            recordDelimitedChange(
              delimitedImportReport,
              record,
              f + 1,
              m,
              cell.format === "date"
                ? "Date converted to native serial"
                : "Selected type or whitespace conversion",
              raw,
              String(cell.raw ?? ""),
            ),
            delimitedImportReport.preview.length < 25 &&
              f === 0 &&
              delimitedImportReport.preview.push({
                record: record,
                raw: c.slice(0, 200),
                values: [],
              }),
            record <= 25 &&
              f < 200 &&
              delimitedImportReport.preview.at(-1)?.values.push(cell.raw));
        }
        return true;
      },
    ),
    delimitedImportReport.records ||
      recordDelimitedIssue(
        delimitedImportReport,
        0,
        undefined,
        "The file has no records. Select a populated delimited file.",
      ),
    delimitedOptions.hasHeader && delimitedImportReport.records)
  ) {
    let c = delimitedImportReport.schema.map((d) => d.label),
      u = new Set();
    c.some(
      (d) =>
        !d.trim() ||
        (u.has(d.toLowerCase()) ? true : (u.add(d.toLowerCase()), false)),
    ) &&
      delimitedImportReport.warnings.push(
        "Blank or duplicate header labels are retained exactly. Column letters identify them; no fields are overwritten or renamed.",
      );
  }
  return (
    (delimitedImportReport.valid =
      delimitedImportReport.complete && delimitedImportReport.issueCount === 0),
    { r: delimitedImportReport, cells: cells }
  );
}
function validateDelimitedImport(bytes, fileName, n = {}) {
  return parseDelimitedRecords(bytes, fileName, n, false).r;
}
function importDelimitedWorkbook(bytes, fileName, n = {}) {
  let { r: delimitedReport, cells } = parseDelimitedRecords(
    bytes,
    fileName,
    n,
    true,
  );
  if (!delimitedReport.valid)
    throw Error(
      `Unable to read delimited data: ${delimitedReport.issues[0]?.message ?? "Incomplete preview."} Review the import schema; no rows were imported.`,
    );
  let blankWorkbook = createBlankWorkbook(
      fileName
        .replace(/\.[^.]+$/, " ")
        .trim()
        .slice(0, 200) || "Imported workbook",
    ),
    sheet = blankWorkbook.sheets[0];
  return (
    (sheet.name = "Imported data"),
    (sheet.cells = cells),
    (sheet.frozenRows = delimitedReport.options.hasHeader ? 1 : 0),
    (sheet.delimitedSource = {
      records: delimitedReport.records,
      columns: delimitedReport.columns,
      hasHeader: delimitedReport.options.hasHeader,
    }),
    {
      workbook: toValidWorkbook(
        createSourceRowProvenance(blankWorkbook, {
          id: crypto.randomUUID(),
          name: fileName.slice(0, 255),
          kind: "file",
        }),
      ),
      warnings: delimitedReport.warnings,
      delimitedReport: delimitedReport,
    }
  );
}
var createRandomUuid = () => globalThis.crypto.randomUUID();
var createImportedWorkbook = (e, sheets) => ({
  schemaVersion: 1,
  id: createRandomUuid(),
  name: e.replace(/\.[^.]+$/, "").slice(0, 200) || "Imported workbook",
  revision: 0,
  metadata: createWorkbookMetadata(),
  sheets: sheets,
  events: [],
});

function parseXlsxWorkbookPackage(bytes, t) {
  let n = inspectXlsxPackage(bytes),
    i = new Set(n.keys()),
    size = n.size,
    workbookParseState = createWorkbookParseState();
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
              (O && parseWorksheetXml(I, _.name, n, workbookParseState),
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
                let xmlAttribute = getXmlAttribute(v[1], "pageOrder");
                if (
                  xmlAttribute !== undefined &&
                  !["overThenDown", "downThenOver"].includes(
                    decodeXmlEntities(xmlAttribute),
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
                  let xmlAttribute = getXmlAttribute(q[1], "id"),
                    W = getXmlAttribute(q[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                    !Number.isSafeInteger(Number(xmlAttribute)) ||
                    Number(xmlAttribute) >= x
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (W !== "1" && W !== "true") continue;
                  let $ = Number(getXmlAttribute(q[1], "min") ?? 0),
                    M = Number(getXmlAttribute(q[1], "max") ?? C - 1);
                  if ($ !== 0 || !Number.isSafeInteger(M) || M < C - 1)
                    throw Error(
                      "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
                    );
                  F.push(Number(xmlAttribute));
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
                let xmlAttribute = getXmlAttribute(v[1], "min"),
                  x = getXmlAttribute(v[1], "max"),
                  C = getXmlAttribute(v[1], "width"),
                  start = Number(xmlAttribute) - 1,
                  end = Number(x) - 1;
                if (
                  !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                  !/^[1-9]\d*$/.test(x ?? "") ||
                  !Number.isSafeInteger(start) ||
                  !Number.isSafeInteger(end) ||
                  start < 0 ||
                  end < start ||
                  end >= workbookLimits.maxCols ||
                  y.length >= workbookLimits.maxCols
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
                let width = C === undefined ? undefined : Number(C);
                if (
                  width !== undefined &&
                  (!C?.trim() || !Number.isFinite(width) || width < 0)
                )
                  throw Error(
                    "XLSX column widths must be finite nonnegative values.",
                  );
                y.push({
                  start: start,
                  end: end,
                  ...(width === undefined ? {} : { width: width }),
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
                let xmlAttribute = getXmlAttribute(A[1], "name"),
                  y = getXmlAttribute(A[1], "r:id");
                xmlAttribute &&
                  y &&
                  t.sheetNames.set(
                    decodeXmlEntities(xmlAttribute),
                    decodeXmlEntities(y),
                  );
              }
            if (_.name === "xl/_rels/workbook.xml.rels")
              for (let A of I.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)) {
                let xmlAttribute = getXmlAttribute(A[1], "Id"),
                  y = getXmlAttribute(A[1], "Target");
                xmlAttribute &&
                  y &&
                  getXmlAttribute(A[1], "TargetMode") !== "External" &&
                  t.sheetRelationships.set(
                    decodeXmlEntities(xmlAttribute),
                    resolvePackagePartPath(decodeXmlEntities(y)),
                  );
              }
            for (let A of I.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (let L of decodeXmlEntities(A[3]).trim().split(/\s+/)) {
                let rangeAddress = parseRangeAddress(L),
                  v =
                    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
                    (rangeAddress.lastCol - rangeAddress.firstCol + 1);
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
          } catch (error) {
            s =
              error instanceof Error
                ? error
                : new Error("Invalid XLSX worksheet metadata.");
            return;
          }
          c++;
        }
      }),
        _.start());
    });
  b.register(mc);
  for (let _ = 0; _ < bytes.length; _ += 2048) {
    if (s) throw s;
    try {
      b.push(
        bytes.subarray(_, Math.min(_ + 2048, bytes.length)),
        _ + 2048 >= bytes.length,
      );
    } catch {
      throw new Error(
        "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
      );
    }
  }
  if (s) throw s;
  if (l !== size || c !== size)
    throw new Error("The XLSX archive contains incomplete entries.");
  let w = [
    xlsxExchangeSupportNotice,
    ...t.nameWarnings,
    ...validateWorkbookPackage(workbookParseState, i),
  ];
  return (
    (t.emptyStringCaches = workbookParseState.emptyStringCaches),
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
function classifyNumberFormat(numFmt) {
  if (!(!numFmt || numFmt === "General"))
    return /%/.test(numFmt)
      ? "percent"
      : /[$€£¥]/.test(numFmt)
        ? "currency"
        : /[ymdhis]/i.test(numFmt.replace(/"[^"]*"/g, ""))
          ? "date"
          : "number";
}
function normalizeArgbHexColor(e) {
  return e && /^(?:[0-9a-f]{8}|[0-9a-f]{6})$/i.test(e)
    ? `#${e.slice(-6).toLowerCase()}`
    : undefined;
}
function convertXlsxCellToRecord(e, xlsxWorkbookPackage = [], n, i = false) {
  let r = (m) => {
      xlsxWorkbookPackage.includes(m) || xlsxWorkbookPackage.push(m);
    },
    value = e.value;
  if (value === null && !e.note && !Object.keys(e.style ?? {}).length) return;
  let raw = null,
    literal;
  typeof value == "string"
    ? ((raw = value), raw.startsWith("=") && (literal = true))
    : typeof value == "number" || typeof value == "boolean"
      ? (raw = value)
      : value instanceof Date
        ? (raw =
            n !== undefined
              ? Number(n)
              : dateToExcelSerial(value, i ? "1904" : "1900"))
        : value &&
          typeof value == "object" &&
          ("formula" in value || "sharedFormula" in value
            ? ((raw = `=${e.formula}`),
              isExternalActionFormula(raw) && (literal = true))
            : "richText" in value
              ? (r(
                  "Rich-text runs are flattened to cell text; per-run formatting is not preserved.",
                ),
                (raw = value.richText.map((m) => m.text).join("")),
                raw.startsWith("=") && (literal = true))
              : "hyperlink" in value
                ? ((raw = value.text), raw.startsWith("=") && (literal = true))
                : "error" in value && (raw = value.error));
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
  for (let cellBorderEdge of cellBorderEdges) {
    let g = e.border?.[cellBorderEdge];
    if (!g?.style) continue;
    if (!borderLineStyles.includes(g.style)) {
      r(
        "An unsupported border line was omitted; supported edges and base values were retained.",
      );
      continue;
    }
    let argbHexColor = normalizeArgbHexColor(g.color?.argb);
    (g.color &&
      !argbHexColor &&
      r(
        "A theme/indexed border color used a black fallback; theme bindings are not preserved.",
      ),
      (c[cellBorderEdge] = {
        style: g.style,
        color: argbHexColor ?? "#000000",
      }));
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
  let argbHexColor = normalizeArgbHexColor(e.font?.color?.argb);
  if (
    (argbHexColor && (l.color = argbHexColor),
    e.fill?.type === "pattern" && e.fill.pattern === "solid")
  ) {
    let argbHexColor = normalizeArgbHexColor(e.fill.fgColor?.argb);
    argbHexColor && (l.fill = argbHexColor);
  }
  e.alignment &&
    ["left", "center", "right"].includes(e.alignment.horizontal ?? "") &&
    (l.align = e.alignment.horizontal);
  let note =
      typeof e.note == "string"
        ? e.note
        : e.note?.texts?.map((text) => text.text).join(""),
    numFmt = e.numFmt && e.numFmt !== "General" ? e.numFmt : undefined;
  return {
    raw: raw,
    ...(literal ? { literal: literal } : {}),
    ...(Object.keys(l).length ? { style: l } : {}),
    ...(numFmt ? { numFmt: numFmt, format: classifyNumberFormat(numFmt) } : {}),
    ...(note ? { note: note } : {}),
  };
}
async function importWorkbookFile(e) {
  if (/\.xltx$/i.test(e.name)) return convertXltxTemplateFile(e);
  if (/\.json$/i.test(e.name)) {
    if (e.size > workbookLimits.maxNativeBytes)
      throw new Error("Native workbook exceeds the 50 MB import limit.");
    let nativeWorkbookJson = parseNativeWorkbookJson(await e.text());
    return {
      workbook: nativeWorkbookJson,
      warnings: [
        ...collectStyleFallbackWarnings(nativeWorkbookJson),
        ...collectDataValidationWarnings(nativeWorkbookJson),
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
  let xlsxWorkbookPackage = parseXlsxWorkbookPackage(t, n),
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
    a.worksheets.some((worksheet) => !n.sheetNames.has(worksheet.name))
  )
    throw new Error(
      "Unsupported or missing worksheet records. Save a standard XLSX copy with ordinary worksheets.",
    );
  if (!a.worksheets.length)
    throw new Error("A workbook must have at least one sheet.");
  let l = 0,
    c = 0,
    u = a.worksheets.map((worksheet) => {
      let O = Math.max(worksheet.columnCount, worksheet.columns?.length ?? 0);
      if (
        worksheet.rowCount > workbookLimits.maxRows ||
        O > workbookLimits.maxCols
      )
        throw new Error(
          `Worksheet ${worksheet.name} exceeds ${workbookLimits.maxRows} rows or ${workbookLimits.maxCols} columns.`,
        );
      let cells = {},
        model = worksheet.model,
        rowHeights = {},
        U = n.sheetRelationships.get(n.sheetNames.get(worksheet.name) ?? ""),
        hiddenRows = n.hiddenRows.get(U ?? "") ?? [],
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
      for (let row of model.rows) {
        if (typeof row.height == "number" && Number.isFinite(row.height)) {
          let F = (row.height * 4) / 3;
          ((rowHeights[String(row.number - 1)] = Math.max(
            21,
            Math.min(546, F),
          )),
            (F < 21 || F > 546) && c++);
        }
        for (let cell of row.cells) {
          let z = worksheet.getCell(cell.address);
          if (z.isMerged && z.master.address !== z.address) continue;
          let j = convertXlsxCellToRecord(
            z,
            xlsxWorkbookPackage,
            n.numericTokens.get(U ?? "")?.get(cell.address),
            !!a.properties.date1904,
          );
          if (j) {
            if (++l > workbookLimits.maxCells)
              throw new Error("Workbook exceeds 100,000 populated cells.");
            cells[z.address] = j;
          }
        }
      }
      let A = worksheet.views?.find((view) => view.state === "frozen"),
        hiddenCols = [];
      (worksheet.columns ?? []).forEach((N, F) => {
        N.hidden && T(hiddenCols, F);
      });
      let columnWidths = {};
      for (let N of n.columns.get(U ?? "") ?? []) {
        if (N.width === undefined) continue;
        let F = N.width * 7;
        (F < 40 || F > 600) &&
          xlsxWorkbookPackage.push(
            "Worksheet " +
              worksheet.name +
              " column widths were clamped to the supported pixel bounds.",
          );
        let z = Math.max(40, Math.min(600, F));
        for (let start = N.start; start <= N.end; start++)
          columnWidths[start] = z;
      }
      let v = (N, F, z, j, q) => {
          if (N === undefined || N === 0) return;
          if (typeof N != "number" || !Number.isFinite(N) || N <= 0)
            throw Error("Invalid worksheet default " + q + ".");
          let D = N * F;
          return (
            (D < z || D > j) &&
              xlsxWorkbookPackage.push(
                "Worksheet " +
                  worksheet.name +
                  " default " +
                  q +
                  " was clamped to the supported pixel bounds.",
              ),
            Math.max(z, Math.min(j, D))
          );
        },
        defaultRowHeight = v(
          worksheet.properties.defaultRowHeight,
          4 / 3,
          21,
          546,
          "row height",
        ),
        defaultColumnWidth = v(
          worksheet.properties.defaultColWidth,
          7,
          40,
          600,
          "column width",
        ),
        C = {
          id: createRandomUuid(),
          name: worksheet.name,
          ...(defaultRowHeight !== undefined
            ? { defaultRowHeight: defaultRowHeight }
            : {}),
          ...(defaultColumnWidth !== undefined
            ? { defaultColumnWidth: defaultColumnWidth }
            : {}),
          cells: cells,
          rowCount: workbookLimits.maxRows,
          colCount: workbookLimits.maxCols,
          frozenRows: A?.ySplit ?? 0,
          ...(A?.xSplit ? { frozenCols: A.xSplit } : {}),
          ...(hiddenRows.length ? { hiddenRows: hiddenRows } : {}),
          ...(hiddenCols.length ? { hiddenCols: hiddenCols } : {}),
          ...(Object.keys(columnWidths).length
            ? { columnWidths: columnWidths }
            : {}),
          ...(Object.keys(rowHeights).length ? { rowHeights: rowHeights } : {}),
          ...(worksheet.state !== "visible" ? { hidden: true } : {}),
          ...(worksheet.state === "veryHidden" ? { veryHidden: true } : {}),
          ...(normalizeArgbHexColor(worksheet.properties.tabColor?.argb)
            ? {
                tabColor: normalizeArgbHexColor(
                  worksheet.properties.tabColor?.argb,
                ),
              }
            : {}),
        };
      (xlsxWorkbookPackage.push(...convertImportedTables(worksheet, C)),
        xlsxWorkbookPackage.push(
          ...convertWorksheetFeatures(
            worksheet,
            C,
            n.conditionalStops.get(U ?? ""),
            n.validationAreas.get(U ?? ""),
          ),
        ),
        xlsxWorkbookPackage.push(
          ...collectImportedPictures(a, worksheet, C, {
            containerFilename: e.name,
          }),
        ));
      try {
        importPageSetup(worksheet, C);
        let N = n.printBreaks.get(U ?? "");
        N && validatePageSetup({ ...C.pageLayout, ...N }, C);
      } catch (error) {
        xlsxWorkbookPackage.push(
          `Page setup on ${C.name} could not be retained: ${error.message}`,
        );
      }
      return C;
    });
  (c &&
    xlsxWorkbookPackage.push(
      `${c} explicit row height(s) were adjusted to the supported ${21}\u2013${546} CSS-pixel range.`,
    ),
    a.properties.date1904 &&
      xlsxWorkbookPackage.push(
        "The source uses the 1904 date system. Its date system, numeric serials and formulas are preserved; calendar dates are not silently converted to 1900.",
      ),
    xlsxWorkbookPackage.push(
      "Formula results are recalculated by the supported local formula engine. Unsupported Excel functions may display errors. Cell error values are imported as displayed text. Audit history and application tickmarks are available only in native JSON backups.",
    ));
  let importedWorkbook = createImportedWorkbook(e.name, u);
  importedWorkbook.xlsxSource = r;
  let externalLinkTargets = parseExternalLinkTargets(Nd(t), r.sha256);
  (externalLinkTargets.length &&
    (importedWorkbook.workbookLinks = {
      schemaVersion: 1,
      sources: [],
      packageTargets: externalLinkTargets,
    }),
    a.properties.date1904 &&
      (importedWorkbook.datePolicy = {
        ...defaultDatePolicy,
        dateSystem: "1904",
      }));
  let p = 0,
    h = [];
  for (let [S, O] of n.sheetNames)
    for (let [E, R] of n.numericTokens.get(n.sheetRelationships.get(O) ?? "") ??
      []) {
      let I = validateXlsxNumericValue(R);
      I && (p++, h.length < 8 && h.push(`${S}!${E}: ${R.slice(0, 80)} (${I})`));
    }
  (p &&
    xlsxWorkbookPackage.push(
      `Numeric precision risk: ${p} numeric value/cache(s) may lose significant digits in Excel or binary conversion. Numeric cells retain their declared numeric type; store identifiers as Text and keep the source file. ${h.join("; ")}`,
    ),
    a.created instanceof Date &&
      Number.isFinite(a.created.getTime()) &&
      a.modified instanceof Date &&
      Number.isFinite(a.modified.getTime()) &&
      a.modified >= a.created &&
      (importedWorkbook.metadata = {
        createdAt: a.created.toISOString(),
        modifiedAt: a.modified.toISOString(),
      }));
  let definedNameScopes = resolveDefinedNameScopes(n.names, importedWorkbook, [
    ...n.sheetNames.keys(),
  ]);
  definedNameScopes.length &&
    (importedWorkbook.definedNames = definedNameScopes);
  let sourceSettings = n.calculation ?? {
    mode: "auto",
    iterative: false,
    maxIterations: null,
    maxChange: null,
  };
  applyCalculationProperties(
    importedWorkbook,
    sourceSettings,
    xlsxWorkbookPackage,
  );
  let validWorkbook = toValidWorkbook(
      createSourceRowProvenance(importedWorkbook, {
        id: createRandomUuid(),
        name: e.name,
        kind: "file",
      }),
    ),
    exchangeSchemaReport = buildExchangeSchemaReport(
      "import",
      e.name,
      t.byteLength,
      validWorkbook,
      xlsxWorkbookPackage,
    );
  ((exchangeSchemaReport.partInventory = o),
    exchangeSchemaReport.warnings.push(
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
    sourceSettings: sourceSettings,
    entries: [],
  };
  return (
    compareFormulaCachesWithRecalculation(
      a,
      validWorkbook,
      exchangeSchemaReport,
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
    (validWorkbook.importedCalculationCaches = _),
    exchangeSchemaReport.warnings.push(
      ...collectDataValidationWarnings(validWorkbook),
    ),
    {
      workbook: toValidWorkbook(validWorkbook),
      warnings: exchangeSchemaReport.warnings,
      report: exchangeSchemaReport,
    }
  );
}
function parseNativeWorkbookJson(e) {
  if (new TextEncoder().encode(e).length > workbookLimits.maxNativeBytes)
    throw new Error("Native workbook exceeds the 50 MB import limit.");
  let workbook;
  try {
    workbook = JSON.parse(e);
  } catch {
    throw new Error("The native workbook is not valid JSON.");
  }
  return toValidWorkbook(workbook);
}

function buildRetainedXlsxSchemaReport(workbook) {
  let t = requireRetainedXlsxSource(workbook),
    xlsxWorkbookPackage = parseXlsxWorkbookPackage(
      t.bytes,
      createEmptyExportLedger(),
    ),
    exchangeSchemaReport = buildExchangeSchemaReport(
      "import",
      t.fileName,
      t.bytes.length,
      workbook,
      xlsxWorkbookPackage,
    );
  return (
    (exchangeSchemaReport.partInventory = F$(Nd(t.bytes), t.sha256)),
    exchangeSchemaReport
  );
}

function toErrorResponse(error) {
  let { code, message } = toFileProcessingError(error);
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
function parseJsonWithinBudget(bytes) {
  let t = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)),
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
        let bytes = data.bytes,
          r = await importWorkbookFile({
            name: data.fileName,
            size: bytes.length,
            arrayBuffer: async () => bytes.buffer,
            text: async () =>
              new TextDecoder("utf-8", { fatal: true }).decode(bytes),
          });
        if (
          (/\.json$/i.test(data.fileName) &&
            r.workbook.xlsxSource &&
            buildRetainedXlsxSchemaReport(r.workbook),
          /\.json$/i.test(data.fileName) && r.workbook.alternateSource)
        ) {
          let alternateSource = r.workbook.alternateSource;
          inspectAlternateWorkbookFile(
            validateAlternateSource(alternateSource),
            alternateSource.fileName,
          );
        }
        if (/\.json$/i.test(data.fileName))
          for (let o of r.workbook.sourceImports ?? []) {
            let a = inspectSourceFile(
              validateRetainedSourceFile(o.original),
              o.original.fileName,
            );
            if (
              a.method !== "document-adapter" &&
              JSON.stringify(
                a.tables.find((table) => table.id === o.table.id),
              ) !== JSON.stringify(o.table)
            )
              throw Error(
                "Retained structured evidence does not match its original.",
              );
          }
        n = r;
        break;
      }
      case "delimited": {
        let report = validateDelimitedImport(
          data.bytes,
          data.fileName,
          data.options,
        );
        n = {
          report: report,
          ...(report.valid
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
        for (let byte of data.bytes) i += String.fromCharCode(byte);
        let dataURL = `data:image/${/\.png$/i.test(data.fileName.trim()) ? "png" : "jpeg"};base64,${btoa(i)}`,
          o = inspectPictureDataUrl(dataURL);
        n = { dataURL: dataURL, width: o.width, height: o.height };
        break;
      }
      default:
        throw new FileProcessingError("FILE_TYPE");
    }
    if (new TextEncoder().encode(JSON.stringify(n)).length > u1.maxResultBytes)
      throw new FileProcessingError("FILE_LIMIT");
    workerScope.postMessage({ id: t, ok: true, value: n });
  } catch (error) {
    workerScope.postMessage({
      id: t,
      ok: false,
      error: toErrorResponse(error),
    });
  }
};
