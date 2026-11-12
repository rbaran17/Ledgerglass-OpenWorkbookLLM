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
var dynamicRequire = ((fallbackRequire) =>
  typeof require < "u"
    ? require
    : typeof Proxy < "u"
      ? new Proxy(fallbackRequire, {
          get: (proxyTarget, propertyKey) =>
            (typeof require < "u" ? require : proxyTarget)[propertyKey],
        })
      : fallbackRequire)(function (moduleName) {
  if (typeof require < "u") return require.apply(this, arguments);
  throw Error('Dynamic require of "' + moduleName + '" is not supported');
});
var copyModuleExports = (
  targetObject,
  sourceModule,
  excludedName,
  propertyDescriptor,
) => {
  if (
    (sourceModule && typeof sourceModule == "object") ||
    typeof sourceModule == "function"
  )
    for (let exportName of listOwnPropertyNames(sourceModule))
      !hasOwnPropertyCheck.call(targetObject, exportName) &&
        exportName !== excludedName &&
        objectDefineProperty(targetObject, exportName, {
          get: () => sourceModule[exportName],
          enumerable:
            !(propertyDescriptor = readPropertyDescriptor(
              sourceModule,
              exportName,
            )) || propertyDescriptor.enumerable,
        });
  return targetObject;
};
var toEsModule = (moduleValue, forceDefaultExport, esModule) => (
  (esModule =
    moduleValue != null
      ? createPrototypeObject(readPrototypeOf(moduleValue))
      : {}),
  copyModuleExports(
    forceDefaultExport || !moduleValue || !moduleValue.__esModule
      ? objectDefineProperty(esModule, "default", {
          value: moduleValue,
          enumerable: true,
        })
      : esModule,
    moduleValue,
  )
);

var getSaxesModule = () => __lib_saxes;

var getExceljsModule = () => __lib_exceljs;
var getPapaparseModule = () => __lib_papaparse;
var saxesEsModuleForXml = toEsModule(getSaxesModule(), 1);

function dateToExcelSerial(date, dateSystem = "1900") {
  return (
    dateToSerial(
      date.getUTCFullYear(),
      date.getUTCMonth() + 1,
      date.getUTCDate(),
      dateSystem,
    ) +
    (date.getUTCHours() * 36e5 +
      date.getUTCMinutes() * 6e4 +
      date.getUTCSeconds() * 1e3 +
      date.getUTCMilliseconds()) /
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
  let seenStyleIds = new Set(),
    seenStyleNames = new Set(),
    styles = styleCatalog.styles.map((style) => {
      (bo(style) || Fr("Invalid named style."),
        assertKnownStyleKeys(
          style,
          ["id", "name", "definition"],
          "named style",
        ));
      let styleId = validateBoundedStyleText(style.id, "Style ID", 100),
        styleName = validateBoundedStyleText(style.name, "Style name");
      return (
        (seenStyleIds.has(styleId) ||
          seenStyleNames.has(styleName.toLocaleLowerCase("en-US"))) &&
          Fr("Named style IDs and names must be unique within this workbook."),
        seenStyleIds.add(styleId),
        seenStyleNames.add(styleName.toLocaleLowerCase("en-US")),
        {
          id: styleId,
          name: styleName,
          definition: validateNamedStyleDefinition(style.definition),
        }
      );
    }),
    isThemeMissing =
      styleCatalog.theme === undefined || styleCatalog.missingTheme === true,
    theme =
      styleCatalog.theme === undefined
        ? cloneStyleValue(officeThemeDefinition)
        : validateWorkbookTheme(styleCatalog.theme);
  return (
    isThemeMissing &&
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
      ...(isThemeMissing ? { missingTheme: true } : {}),
    }
  );
}

function applyNamedStyleCatalog(styleCatalog, workbook) {
  let normalizedCatalog =
    styleCatalog === undefined
      ? undefined
      : normalizeWorkbookStyleCatalog(styleCatalog);
  normalizedCatalog
    ? (workbook.styleCatalog = normalizedCatalog)
    : delete workbook.styleCatalog;
  let hasOmittedTheme = bo(styleCatalog) && styleCatalog.theme === undefined;
  for (let sheet of workbook.sheets)
    for (let cell of hasLinkedOrBoundCells(sheet.cells)
      ? Object.values(sheet.cells)
      : []) {
      if (!cell.style?.binding) continue;
      let styleBinding = validateNamedStyleBinding(cell.style.binding),
        namedStyleBinding = resolveNamedStyleBinding(
          normalizedCatalog,
          styleBinding,
        );
      (!namedStyleBinding.missing &&
        !hasOmittedTheme &&
        canonicalStyleKey(styleBinding) !==
          canonicalStyleKey(namedStyleBinding) &&
        Fr(
          "Named style appearance does not match the workbook catalog. Reapply the style from a trusted catalog.",
        ),
        (cell.style.binding = namedStyleBinding));
    }
}

function createBlankWorkbook(workbookName = "Book1") {
  return toValidWorkbook({
    schemaVersion: 1,
    id: crypto.randomUUID(),
    name: workbookName.trim(),
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
  let inspection = summarizeValidationInspection(workbook);
  return inspection.invalidCells || !inspection.complete
    ? [
        `Data validation: ${inspection.message} Source values were retained; use Data Validation to review.`,
      ]
    : [];
}

function resolveDefinedNameScopes(names, importedWorkbook, sheetNames) {
  if (
    new Set(sheetNames.map((sheetName) => sheetName.toLowerCase())).size !==
    sheetNames.length
  )
    return throwXlsxDefinedNameError(
      "Duplicate worksheet names make name scopes ambiguous.",
    );
  let scopedNames = names.map((definedName) => {
    let localSheetName =
        definedName.localSheetId === undefined
          ? undefined
          : sheetNames[definedName.localSheetId],
      scopeSheetId =
        localSheetName === undefined
          ? undefined
          : importedWorkbook.sheets.find(
              (sheet) => sheet.name === localSheetName,
            )?.id;
    return definedName.localSheetId !== undefined && !scopeSheetId
      ? throwXlsxDefinedNameError(
          `Name "${definedName.name}" has an invalid or unsupported worksheet scope.`,
        )
      : {
          name: definedName.name,
          refersTo: definedName.refersTo,
          ...(scopeSheetId ? { scopeSheetId: scopeSheetId } : {}),
          ...(definedName.comment !== undefined
            ? { comment: definedName.comment }
            : {}),
        };
  });
  for (let scopedName of scopedNames)
    validateNameSheetReferences(scopedName.refersTo, importedWorkbook);
  try {
    return validateDefinedName(scopedNames, importedWorkbook);
  } catch (error) {
    return throwXlsxDefinedNameError(error.message);
  }
}

function createSourceRowProvenance(workbook, sourceIdentity) {
  if (
    !isBoundedString(sourceIdentity.id, 100) ||
    !isBoundedString(sourceIdentity.name, 255) ||
    !["file", "snapshot"].includes(sourceIdentity.kind)
  )
    throw Error("Provide a valid source identity, name and origin kind.");
  let rowOrigins = [...(workbook.rowOrigins ?? [])],
    sheets = workbook.sheets.map((sheet) => {
      let rowOriginIds = { ...sheet.rowOriginIds },
        populatedRowIndexes = new Set(
          Object.keys(sheet.cells).map(
            (cell) => Number(/\d+$/.exec(cell)[0]) - 1,
          ),
        );
      for (let rowIndex of populatedRowIndexes)
        if (!rowOriginIds[rowIndex]) {
          if (rowOrigins.length >= workbookLimits.maxCells)
            throw Error(
              "Row provenance has reached its 100,000 source-row budget.",
            );
          let rowOriginId = crypto.randomUUID();
          ((rowOriginIds[rowIndex] = rowOriginId),
            rowOrigins.push({
              id: rowOriginId,
              sourceId: sourceIdentity.id,
              sourceName: sourceIdentity.name.slice(0, 255),
              sourceSheet: sheet.name,
              originalRow: rowIndex + 1,
              kind: sourceIdentity.kind,
            }));
        }
      return { ...sheet, rowOriginIds: rowOriginIds };
    });
  return { ...workbook, sheets: sheets, rowOrigins: rowOrigins };
}

function convertImportedTables(worksheet, sheetOrThrow) {
  let warnings = [],
    importedTables = [],
    sourceTables = worksheet.model.tables ?? [];
  for (let sourceTable of sourceTables) {
    if (sourceTable.headerRow === false) {
      warnings.push(
        `Table ${sourceTable.name} has no header row; its cells were retained without table metadata.`,
      );
      continue;
    }
    try {
      let range = Vm(sourceTable.tableRef);
      resolveSourceRange(range, sheetOrThrow);
      let tableStyle =
          Object.entries(tablePaletteStyleNames).find(
            ([, paletteThemeName]) =>
              paletteThemeName === sourceTable.style?.theme,
          )?.[0] ?? "blue",
        totals = {};
      for (let column of sourceTable.columns) {
        let totalsFunction = column.totalsRowFunction?.toLowerCase();
        if (basicAggregateOperations.includes(totalsFunction))
          Object.defineProperty(totals, column.name, {
            value: totalsFunction,
            enumerable: true,
            writable: true,
            configurable: true,
          });
        else if (totalsFunction === "custom") {
          let totalsFormula = column.totalsRowFormula ?? "",
            aggregateFunction = /^=?\s*(SUM|COUNTA|AVERAGE|MIN|MAX)\(/i
              .exec(totalsFormula)?.[1]
              .toLowerCase();
          aggregateFunction &&
            Object.defineProperty(totals, column.name, {
              value:
                aggregateFunction === "counta" ? "count" : aggregateFunction,
              enumerable: true,
              writable: true,
              configurable: true,
            });
        }
      }
      (importedTables.push({
        id: createUuid(),
        name: validateTableName(sourceTable.name),
        range: range,
        columns: sourceTable.columns.map((column) => ({
          id: createUuid(),
          name: validateColumnName(column.name),
        })),
        style: tableStyle,
        showTotals: !!sourceTable.totalsRow,
        totals: totals,
      }),
        Object.values(tablePaletteStyleNames).includes(
          sourceTable.style?.theme,
        ) ||
          warnings.push(
            `Table ${sourceTable.name}: source cells were preserved; unsupported table theme maps to the blue table style for future edits.`,
          ));
    } catch (error) {
      warnings.push(
        `Table ${sourceTable.name}: ${error.message} Cells remain available without this table definition.`,
      );
    }
  }
  return (
    importedTables.length &&
      ((sheetOrThrow.tables = importedTables),
      normalizeTableDefinitions(sheetOrThrow, sheetOrThrow)),
    warnings
  );
}

function validateReportSnapshot(snapshot) {
  if (new TextEncoder().encode(JSON.stringify(snapshot)).length > 4e6)
    throw Error("A report snapshot exceeds 4 MB.");
  let report = calculationReportSchema.parse(snapshot),
    rangeAddress = parseRangeAddress(report.request.range);
  if (
    (validateObjectReportRequest(report.request),
    report.hash !== hashRecordExcludingHash(report))
  )
    throw Error("Report snapshot integrity check failed.");
  let expectedColumns = Array.from(
    { length: rangeAddress.lastCol - rangeAddress.firstCol + 1 },
    (a, columnOffset) => rangeAddress.firstCol + columnOffset,
  );
  if (
    fingerprintCanonicalValue(expectedColumns) !==
      fingerprintCanonicalValue(report.columns) ||
    report.widths.length !== expectedColumns.length
  )
    throw Error("Invalid report column layout.");
  let populatedRowNumbers = report.rows.map((row) => row.row),
    allRowNumbers = [...populatedRowNumbers, ...report.excludedRows].sort(
      (leftRow, rightRow) => leftRow - rightRow,
    );
  if (
    report.rows.length * expectedColumns.length > 2e3 ||
    allRowNumbers.length !== rangeAddress.lastRow - rangeAddress.firstRow + 1 ||
    allRowNumbers.some(
      (rowNumber, rowPosition) =>
        rowNumber !== rangeAddress.firstRow + rowPosition + 1,
    ) ||
    populatedRowNumbers.some(
      (currentRowNumber, populatedIndex) =>
        populatedIndex > 0 &&
        currentRowNumber <= populatedRowNumbers[populatedIndex - 1],
    )
  )
    throw Error("Invalid report row accounting.");
  for (let row of report.rows) {
    if (row.cells.length !== expectedColumns.length)
      throw Error("Invalid report cell layout.");
    row.cells.forEach((cell, columnPosition) => {
      if (
        cell.address !==
        `${columnIndexToLetters(expectedColumns[columnPosition])}${row.row}`
      )
        throw Error("Invalid report cell address.");
      cell.style &&= validateCellStyle(cell.style);
    });
  }
  if (report.object) {
    let objectSheet = {
      id: report.request.sheetId,
      name: report.sheetName,
      cells: {},
      rowCount: 1048576,
      colCount: 16384,
    };
    (validateDrawingObjects([report.object], objectSheet),
      (report.object = objectSheet.objects[0]));
  }
  if (
    (report.request.kind === "object") != !!report.object ||
    (report.object && report.object.id !== report.request.objectId)
  )
    throw Error("Invalid report drawing identity.");
  if (report.hash !== hashRecordExcludingHash(report))
    throw Error("Report metadata is not canonical.");
  return report;
}
function validateReportOutputs(workbook, normalizedWorkbook) {
  if (workbook.reportOutputs !== undefined) {
    let reportOutputs = X.array(
      X.object({
        id: nonEmptyLabelSchema,
        version: X.number().int().min(1).max(1e6),
        snapshot: X.unknown(),
      }).strict(),
    )
      .max(20)
      .parse(workbook.reportOutputs);
    if (
      new Set(reportOutputs.map((reportOutput) => reportOutput.id)).size !==
      reportOutputs.length
    )
      throw Error("Duplicate report output IDs.");
    if (
      new TextEncoder().encode(JSON.stringify(reportOutputs)).length >
      reportOutputLimits.metadataBytes
    )
      throw Error("Saved reports exceed 16 MB.");
    normalizedWorkbook.reportOutputs = reportOutputs.map(
      (reportOutputEntry) => ({
        ...reportOutputEntry,
        snapshot: validateReportSnapshot(reportOutputEntry.snapshot),
      }),
    );
  }
  if (workbook.embeddedObjectInventories !== undefined) {
    let embeddedObjectInventories = X.array(importedCacheEvidenceSchema)
      .max(20)
      .parse(workbook.embeddedObjectInventories);
    if (
      new Set(embeddedObjectInventories.map((inventory) => inventory.id))
        .size !== embeddedObjectInventories.length ||
      JSON.stringify(embeddedObjectInventories).length > 2e6
    )
      throw Error("Invalid embedded-object inventory count or size.");
    normalizedWorkbook.embeddedObjectInventories = embeddedObjectInventories;
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
  let normalizedPageSetup = {
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
  validatePageSetup(normalizedPageSetup, sheetOrThrow);
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

function validateFieldTransformEvidence(fieldTransforms, normalizedWorkbook) {
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
  let seenTransformIds = new Set(),
    outputCount = 0;
  normalizedWorkbook.fieldTransforms = fieldTransforms.map((fieldTransform) => {
    let parseResult = analysisOutputEvidenceSchema.safeParse(fieldTransform);
    if (!parseResult.success)
      throw Error(
        `Invalid field transformation evidence: ${parseResult.error.issues[0]?.message}`,
      );
    let data = parseResult.data,
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
    if (
      seenTransformIds.has(data.id) ||
      (outputCount += data.outputs.length) > 2e4
    )
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (seenTransformIds.add(data.id),
      data.sourceSheetId !== request.sourceSheetId ||
        data.outputSheetName !== request.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    let sourceAddresses = new Set(),
      sources = data.sources.map((source) => {
        let sourcePosition = _t(source.address);
        if (
          sourcePosition.row < rangeAddress.firstRow ||
          sourcePosition.row > rangeAddress.lastRow ||
          sourcePosition.col < rangeAddress.firstCol ||
          sourcePosition.col > rangeAddress.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (sourceAddresses.has(source.address))
          throw Error("Duplicate original field source.");
        sourceAddresses.add(source.address);
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
      outputAddresses = new Set();
    for (let output of data.outputs) {
      if (outputAddresses.has(output.address))
        throw Error("Duplicate transformed output address.");
      if (
        (outputAddresses.add(output.address),
        ((output.type === "number" || output.type === "date") &&
          typeof output.value != "number") ||
          (output.type === "boolean" && typeof output.value != "boolean") ||
          (output.type === "text" && typeof output.value != "string") ||
          (output.type === "blank" && output.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (let transformRecord of [...data.outputs, ...data.exceptions])
      if (
        transformRecord.row <
          rangeAddress.firstRow + 1 + Number(request.header) ||
        transformRecord.row > rangeAddress.lastRow + 1 ||
        transformRecord.sources.some(
          (source) => !sourceAddresses.has(source),
        ) ||
        new Set(transformRecord.sources).size !== transformRecord.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...data, request: request, sources: sources };
  });
}

function toValidWorkbook(workbook) {
  return normalizeWorkbookDocument(workbook, false);
}
function normalizeWorkbookDocument(workbook, canReuseNormalizedValue) {
  if (
    canReuseNormalizedValue &&
    isDeepFrozenValue(workbook) &&
    $V.has(workbook)
  )
    return workbook;
  if (!Gs(workbook) || workbook.schemaVersion !== 1)
    return jt("Unsupported workbook schema. Expected schemaVersion 1.");
  let workbookId = requireBoundedText(workbook.id, "Workbook ID", 100),
    workbookName = requireBoundedText(workbook.name, "Workbook name", 200);
  if (!workbookId || !workbookName.trim())
    return jt("Workbook ID and name are required.");
  let revision = P2(
    workbook.revision,
    0,
    Number.MAX_SAFE_INTEGER - 1,
    "Revision",
  );
  if (!Array.isArray(workbook.sheets) || !workbook.sheets.length)
    return jt("A workbook must contain at least one sheet.");
  let populatedCellCount = 0,
    sheetIds = new Set(),
    lowerSheetNames = new Set(),
    sheets = workbook.sheets.map((sheet) => {
      if (!Gs(sheet)) return jt("Invalid sheet.");
      let sheetId = requireBoundedText(sheet.id, "Sheet ID", 100),
        sheetName = requireBoundedText(sheet.name, "Sheet name", 31);
      if (
        !sheetId ||
        !sheetName.trim() ||
        /[\\/\[\]*?:]/.test(sheetName) ||
        sheetIds.has(sheetId) ||
        lowerSheetNames.has(sheetName.toLowerCase())
      )
        return jt(
          "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
        );
      (sheetIds.add(sheetId), lowerSheetNames.add(sheetName.toLowerCase()));
      let rows = P2(sheet.rowCount, 1, workbookLimits.maxRows, "Row count"),
        cols = P2(sheet.colCount, 1, workbookLimits.maxCols, "Column count");
      if (!Gs(sheet.cells)) return jt("Sheet cells must be an address map.");
      let cachedCellCount =
        canReuseNormalizedValue && isDeepFrozenValue(sheet)
          ? bN.get(sheet)
          : undefined;
      if (cachedCellCount !== undefined) {
        if (
          ((populatedCellCount += cachedCellCount),
          populatedCellCount > workbookLimits.maxCells)
        )
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        let reusedSheet = { ...sheet };
        return (bN.set(reusedSheet, cachedCellCount), reusedSheet);
      }
      let cells = {};
      for (let [cellAddress, cell] of Object.entries(sheet.cells)) {
        let cellPosition = _t(cellAddress);
        if (
          cellAddress !==
            `${columnIndexToLetters(cellPosition.col)}${cellPosition.row + 1}` ||
          cellPosition.row >= rows ||
          cellPosition.col >= cols
        )
          return jt(`Invalid or out-of-bounds cell: ${cellAddress}`);
        if (++populatedCellCount > workbookLimits.maxCells)
          return jt("Workbook exceeds the 100,000 populated cell limit.");
        cells[cellAddress] = normalizeCellRecord(cell, canReuseNormalizedValue);
      }
      let normalizedSheet = {
        id: sheetId,
        name: sheetName,
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        cells: cells,
      };
      if (sheet.columnWidths !== undefined) {
        if (!Gs(sheet.columnWidths))
          return jt("Column widths must be an index map.");
        normalizedSheet.columnWidths = {};
        for (let [columnIndexKey, columnWidth] of Object.entries(
          sheet.columnWidths,
        )) {
          if (
            !/^(0|[1-9]\d*)$/.test(columnIndexKey) ||
            Number(columnIndexKey) >= cols ||
            typeof columnWidth != "number" ||
            !Number.isFinite(columnWidth) ||
            columnWidth < 40 ||
            columnWidth > 600
          )
            return jt(
              "Column widths must be 40\u2013600 pixels for an existing column.",
            );
          normalizedSheet.columnWidths[columnIndexKey] = columnWidth;
        }
      }
      if (sheet.rowHeights !== undefined) {
        if (
          !Gs(sheet.rowHeights) ||
          (Object.getPrototypeOf(sheet.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(sheet.rowHeights) !== null)
        )
          return jt("Row heights must be a plain index map.");
        normalizedSheet.rowHeights = {};
        for (let [rowIndexKey, rowHeight] of Object.entries(sheet.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(rowIndexKey) ||
            Number(rowIndexKey) >= rows ||
            typeof rowHeight != "number" ||
            !Number.isFinite(rowHeight) ||
            rowHeight < 21 ||
            rowHeight > 546
          )
            return jt(
              `Row heights must be ${21}\u2013${546} pixels for an existing row.`,
            );
          normalizedSheet.rowHeights[rowIndexKey] = rowHeight;
        }
      }
      if (sheet.hidden !== undefined) {
        if (typeof sheet.hidden != "boolean")
          return jt("Sheet hidden flag must be boolean.");
        normalizedSheet.hidden = sheet.hidden;
      }
      return (
        validateWorksheetFeatures(sheet, normalizedSheet),
        validateOutlineGroups(sheet, normalizedSheet, {
          rows: rows,
          cols: cols,
        }),
        validateFreezePanes(sheet, normalizedSheet, { rows: rows, cols: cols }),
        validateWorksheetVisibility(sheet, normalizedSheet),
        validateDelimitedSourceFootprint(
          sheet.delimitedSource,
          normalizedSheet,
        ),
        normalizeTableDefinitions(sheet, normalizedSheet),
        sheet.objects !== undefined &&
          validateDrawingObjects(sheet.objects, normalizedSheet),
        sheet.pageLayout !== undefined &&
          validatePageSetup(sheet.pageLayout, normalizedSheet),
        normalizedSheet
      );
    });
  if (sheets.every((sheet) => sheet.hidden))
    return jt("At least one sheet must be visible.");
  if (!Array.isArray(workbook.events) || workbook.events.length > 2e3)
    return jt("Workbook history must be an array of at most 2,000 events.");
  let eventIds = new Set(),
    changeCount = 0,
    events = workbook.events.map((historyEvent) => {
      if (!Gs(historyEvent)) return jt("Invalid history event.");
      let eventId = requireBoundedText(historyEvent.id, "Event ID", 100),
        timestamp = requireBoundedText(
          historyEvent.timestamp,
          "Event timestamp",
          40,
        );
      return !eventId ||
        eventIds.has(eventId) ||
        !Number.isFinite(Date.parse(timestamp))
        ? jt("History events require unique IDs and valid timestamps.")
        : (eventIds.add(eventId),
          !Array.isArray(historyEvent.changes) ||
          historyEvent.changes.length > workbookLimits.maxPatches ||
          (changeCount += historyEvent.changes.length) > workbookLimits.maxCells
            ? jt("Invalid history change count.")
            : {
                id: eventId,
                timestamp: timestamp,
                actor: requireBoundedText(historyEvent.actor, "Actor", 100),
                action: requireBoundedText(historyEvent.action, "Action", 200),
                changes: historyEvent.changes.map((change) => {
                  if (
                    !Gs(change) ||
                    typeof change.sheetId != "string" ||
                    !change.sheetId.length ||
                    change.sheetId.length > 100 ||
                    typeof change.address != "string"
                  )
                    return jt("Invalid history cell reference.");
                  let changePosition = _t(change.address);
                  return change.address !==
                    `${columnIndexToLetters(changePosition.col)}${changePosition.row + 1}`
                    ? jt("History cell address is not canonical.")
                    : change.before === undefined && change.after === undefined
                      ? jt("History change must have a before or after value.")
                      : {
                          sheetId: change.sheetId,
                          address: change.address,
                          ...(change.before !== undefined
                            ? {
                                before: normalizeCellRecord(
                                  change.before,
                                  canReuseNormalizedValue,
                                ),
                              }
                            : {}),
                          ...(change.after !== undefined
                            ? {
                                after: normalizeCellRecord(
                                  change.after,
                                  canReuseNormalizedValue,
                                ),
                              }
                            : {}),
                        };
                }),
              });
    }),
    normalizedWorkbook = {
      schemaVersion: 1,
      id: workbookId,
      name: workbookName,
      revision: revision,
      sheets: sheets,
      events: events,
    };
  return (
    importFormulaCacheEvidence(workbook, normalizedWorkbook),
    workbook.datePolicy !== undefined &&
      (normalizedWorkbook.datePolicy = normalizeDatePolicy(
        workbook.datePolicy,
      )),
    validateSavedFilterViews(workbook.savedFilterViews, normalizedWorkbook),
    validateSavedRecipeMetadata(workbook, normalizedWorkbook),
    verifyConsolidationEvidence(workbook, normalizedWorkbook),
    buildExternalSourceRefreshState(workbook, normalizedWorkbook),
    validateDataModelReport(workbook, normalizedWorkbook),
    validateAnalysisReport(workbook, normalizedWorkbook),
    validateReportOutputs(workbook, normalizedWorkbook),
    normalizeRowProvenance(workbook, normalizedWorkbook),
    validateAutomationArtifacts(workbook, normalizedWorkbook),
    readRetainedXlsxSource(workbook, normalizedWorkbook),
    normalizeAlternateFormatSource(workbook, normalizedWorkbook),
    validateSourceImports(workbook, normalizedWorkbook),
    attachWorkbookLinkMetadata(workbook, normalizedWorkbook),
    normalizeTextCleanups(workbook.textCleanups, normalizedWorkbook),
    validateFieldTransformEvidence(
      workbook.fieldTransforms,
      normalizedWorkbook,
    ),
    validateTableSourceEvidence(workbook.tableEvidence, normalizedWorkbook),
    workbook.whatIfAnalyses !== undefined &&
      (normalizedWorkbook.whatIfAnalyses = validateSavedWhatIfAnalyses(
        workbook.whatIfAnalyses,
      )),
    workbook.definedNames !== undefined &&
      (normalizedWorkbook.definedNames = validateDefinedName(
        workbook.definedNames,
        normalizedWorkbook,
      )),
    validateExtensionState(workbook, normalizedWorkbook),
    validateWorkbookStructureProtection(
      workbook.structureProtection,
      normalizedWorkbook,
    ),
    validateWorkbookMetadata(workbook.metadata, normalizedWorkbook),
    assertUniqueTableNames(normalizedWorkbook),
    normalizeReviewMetadata(workbook, normalizedWorkbook),
    applyNamedStyleCatalog(workbook.styleCatalog, normalizedWorkbook),
    workbook.formulaWatches !== undefined &&
      (normalizedWorkbook.formulaWatches = validateFormulaWatches(
        workbook.formulaWatches,
        normalizedWorkbook,
      )),
    dR(workbook, normalizedWorkbook),
    copyExtensionCacheEntries(workbook, normalizedWorkbook),
    normalizedWorkbook
  );
}

var bytesToBase64 = (bytes) => {
  let binary = "";
  for (let byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};
function collectImportedPictures(workbook, worksheet, sheet, options = {}) {
  let pictures = [],
    warnings = [],
    pictureBytes = 0,
    inspectedCount = 0;
  for (let imageAnchor of worksheet.getImages()) {
    if (++inspectedCount > 1e3 || pictures.length >= _N) {
      warnings.push(
        "Additional pictures beyond the 100-object / 1,000-source inspection limit were omitted.",
      );
      break;
    }
    try {
      let image = workbook.getImage(Number(imageAnchor.imageId)),
        pictureType =
          image.extension === "jpeg"
            ? "jpeg"
            : image.extension === "png"
              ? "png"
              : null;
      if (!pictureType) throw new Error("Only PNG/JPEG images are supported.");
      if (image.buffer && image.buffer.byteLength > maxPictureHashBytes)
        throw Error("Embedded image exceeds 2 MB.");
      let base64Payload = image.base64?.startsWith("data:")
          ? image.base64.slice(image.base64.indexOf(",") + 1)
          : image.base64,
        imageByteSize =
          base64Payload !== undefined
            ? Math.floor((base64Payload.length * 3) / 4) -
              (base64Payload.endsWith("==")
                ? 2
                : base64Payload.endsWith("=")
                  ? 1
                  : 0)
            : (image.buffer?.byteLength ?? 0);
      if (imageByteSize > maxPictureHashBytes)
        throw Error("Embedded image exceeds 2 MB.");
      if (pictureBytes + imageByteSize > 10 * 1024 * 1024)
        throw Error(
          "Additional picture omitted: pictures in one worksheet may use at most 10 MB.",
        );
      let dataURL = image.base64?.startsWith("data:")
          ? image.base64
          : image.base64
            ? `data:image/${pictureType};base64,${image.base64}`
            : `data:image/${pictureType};base64,${bytesToBase64(new Uint8Array(image.buffer))}`,
        pictureInspection = inspectPictureDataUrl(dataURL),
        columnIndex = Math.floor(imageAnchor.range.tl.col),
        rowIndex = Math.floor(imageAnchor.range.tl.row),
        anchoredDrawingObject = createAnchoredDrawingObject(
          "picture",
          `${columnIndexToLetters(columnIndex)}${rowIndex + 1}`,
        );
      ((anchoredDrawingObject.offsetX = Math.max(
        0,
        (imageAnchor.range.tl.col - columnIndex) *
          (sheet.columnWidths?.[columnIndex] ?? 84),
      )),
        (anchoredDrawingObject.offsetY = Math.max(
          0,
          (imageAnchor.range.tl.row - rowIndex) *
            (sheet.rowHeights?.[rowIndex] ?? 21),
        )),
        (anchoredDrawingObject.width = Math.max(
          24,
          Math.min(4e3, imageAnchor.range.ext?.width ?? 240),
        )),
        (anchoredDrawingObject.height = Math.max(
          24,
          Math.min(4e3, imageAnchor.range.ext?.height ?? 140),
        )),
        (anchoredDrawingObject.dataURL = dataURL),
        (anchoredDrawingObject.evidence = createPictureEvidence(dataURL, {
          origin: "xlsx-media",
          ...(options.containerFilename !== undefined
            ? { containerFilename: options.containerFilename }
            : {}),
          mediaReference: `imageId:${imageAnchor.imageId}`,
        })),
        (anchoredDrawingObject.label = `Picture ${pictures.length + 1}`),
        pictures.push(anchoredDrawingObject),
        (pictureBytes += pictureInspection.bytes.length));
    } catch (error) {
      warnings.push(
        error instanceof Error
          ? error.message
          : "An image could not be imported.",
      );
    }
  }
  return (validateDrawingObjects(pictures, sheet), warnings);
}

var escapeJsonPointerSegment = (segment) =>
  segment.replaceAll("~", "~0").replaceAll("/", "~1");
function parseStrictJsonText(text) {
  let position = 0,
    nodeCount = 0,
    skipWhitespace = () => {
      for (; /[\t\n\r ]/.test(text[position] ?? "x");) position++;
    },
    readString = () => {
      let start = position++;
      for (; position < text.length;) {
        let character = text[position++];
        if (character === '"') return JSON.parse(text.slice(start, position));
        character === "\\" && position++;
      }
      throw Error("Invalid JSON string.");
    };
  function parseValue(depth) {
    if (
      depth > sourceImportLimits.depth ||
      ++nodeCount > sourceImportLimits.nodes
    )
      throw Error("JSON nesting or node budget exceeded.");
    skipWhitespace();
    let firstChar = text[position];
    if (firstChar === '"') return readString();
    if (firstChar === "[") {
      (position++, skipWhitespace());
      let parsedArray = [];
      if (text[position] === "]") return (position++, parsedArray);
      for (;;) {
        if (
          (parsedArray.push(parseValue(depth + 1)),
          skipWhitespace(),
          text[position++] === "]")
        )
          return parsedArray;
        if (text[position - 1] !== ",") throw Error("Invalid JSON array.");
      }
    }
    if (firstChar === "{") {
      (position++, skipWhitespace());
      let parsedObject = Object.create(null);
      if (text[position] === "}") return (position++, parsedObject);
      for (;;) {
        if ((skipWhitespace(), text[position] !== '"'))
          throw Error("Invalid JSON object.");
        let fieldName = readString();
        if (
          ["__proto__", "prototype", "constructor"].includes(fieldName) ||
          Object.hasOwn(parsedObject, fieldName)
        )
          throw Error("Unsafe or duplicate JSON field.");
        if ((skipWhitespace(), text[position++] !== ":"))
          throw Error("Invalid JSON field.");
        if (
          ((parsedObject[fieldName] = parseValue(depth + 1)),
          skipWhitespace(),
          text[position++] === "}")
        )
          return parsedObject;
        if (text[position - 1] !== ",") throw Error("Invalid JSON object.");
      }
    }
    for (let [literalText, literalValue] of [
      ["true", true],
      ["false", false],
      ["null", null],
    ])
      if (text.startsWith(literalText, position))
        return ((position += literalText.length), literalValue);
    let numberMatch = /-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(
      text.slice(position),
    );
    if (!numberMatch || numberMatch.index !== 0)
      throw Error("Invalid JSON value.");
    position += numberMatch[0].length;
    let numberValue = Number(numberMatch[0]);
    if (
      !Number.isFinite(numberValue) ||
      (Number.isInteger(numberValue) && !Number.isSafeInteger(numberValue))
    )
      throw Error("Unsafe JSON number; encode large identifiers as strings.");
    return numberValue;
  }
  let parsedValue = parseValue(0);
  if ((skipWhitespace(), position !== text.length))
    throw Error("Trailing JSON content.");
  return parsedValue;
}
function describeStructuredValue(value, path) {
  return {
    value:
      typeof value == "object" && value !== null
        ? JSON.stringify(value)
        : value,
    type:
      typeof value == "object" && value !== null
        ? Array.isArray(value)
          ? "array"
          : "object"
        : value === null
          ? "null"
          : typeof value,
    path: path,
    confidence: 1,
  };
}
function collectJsonSourceTables(jsonValue) {
  let tables = [],
    visitNode = (node, path) => {
      if (Array.isArray(node))
        node.length && tables.push({ value: node, path: path });
      else if (yg(node))
        for (let [key, entryValue] of Object.entries(node))
          visitNode(entryValue, path + "/" + escapeJsonPointerSegment(key));
    };
  if (
    (visitNode(jsonValue, ""),
    !tables.length &&
      yg(jsonValue) &&
      tables.push({ value: [jsonValue], path: "" }),
    !tables.length)
  )
    throw Error("Choose a JSON object or a nonempty array.");
  if (tables.length > sourceImportLimits.tables)
    throw Error("Source table count exceeded.");
  return tables.map((table, tableIndex) => {
    let isRootTable = yg(jsonValue) && table.value[0] === jsonValue,
      columnPathSet = new Set(),
      cellMaps = table.value.map((sourceRecord, recordIndex) => {
        let cellMap = new Map(),
          basePath = isRootTable ? table.path : table.path + "/" + recordIndex,
          flattenValue = (nodeValue, pathPrefix) => {
            if (yg(nodeValue) && Object.keys(nodeValue).length)
              for (let [childKey, childValue] of Object.entries(nodeValue))
                flattenValue(
                  childValue,
                  pathPrefix + "/" + escapeJsonPointerSegment(childKey),
                );
            else
              (columnPathSet.add(pathPrefix),
                cellMap.set(
                  pathPrefix,
                  describeStructuredValue(nodeValue, basePath + pathPrefix),
                ));
          };
        return (flattenValue(sourceRecord, ""), cellMap);
      }),
      columnPaths = [...columnPathSet];
    return {
      id: "json-" + tableIndex,
      name: table.path || "Root",
      path: table.path,
      ambiguous: tables.length > 1,
      columns: columnPaths.map((path) => ({
        name: path || "Value",
        path: path,
      })),
      rows: cellMaps.map((rowCellMap, rowIndex) =>
        columnPaths.map(
          (columnPath) =>
            rowCellMap.get(columnPath) ?? {
              value: null,
              type: "null",
              path:
                table.path + (isRootTable ? "" : "/" + rowIndex) + columnPath,
              confidence: 1,
              missing: true,
            },
        ),
      ),
    };
  });
}
function parseXmlSourceTables(xmlText) {
  if (/<!DOCTYPE|<!ENTITY/i.test(xmlText))
    throw Error("XML document types and entities are not supported.");
  if (/<\?xml[^>]*encoding\s*=\s*['"](?!utf-8['"])/i.test(xmlText))
    throw Error("Use UTF-8 XML.");
  let elementStack = [],
    rootNodes = [],
    nodeCount = 0,
    parser = new saxesEsModuleForXml.SaxesParser({ xmlns: true }),
    qualifyName = (namespaceUri, localName) =>
      namespaceUri ? `{${namespaceUri}}${localName}` : localName;
  parser.on("opentag", (openTag) => {
    if (
      ++nodeCount > sourceImportLimits.nodes ||
      elementStack.length >= sourceImportLimits.depth
    )
      throw Error("XML nesting or node budget exceeded.");
    let elementName = qualifyName(openTag.uri, openTag.local),
      parentElement = elementStack.at(-1),
      siblings = parentElement?.children ?? rootNodes,
      elementRecord = {
        namespaces: { ...(parentElement?.namespaces ?? {}), ...openTag.ns },
        name: elementName,
        path:
          (parentElement?.path ?? "") +
          "/" +
          elementName +
          "[" +
          (siblings.filter((sibling) => sibling.name === elementName).length +
            1) +
          "]",
        attributes: [],
        text: "",
        children: [],
      };
    for (let attribute of Object.values(openTag.attributes))
      if (attribute.uri !== "http://www.w3.org/2000/xmlns/")
        if (
          attribute.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          attribute.local === "type"
        ) {
          let [typePrefix, typeLocalName] = attribute.value.includes(":")
            ? attribute.value.split(":")
            : ["", attribute.value];
          if (typePrefix && !elementRecord.namespaces[typePrefix])
            throw Error("XML type prefix is undeclared.");
          elementRecord.type = qualifyName(
            elementRecord.namespaces[typePrefix] ?? "",
            typeLocalName,
          );
        } else
          attribute.uri === "http://www.w3.org/2001/XMLSchema-instance" &&
          attribute.local === "nil"
            ? (elementRecord.nil = attribute.value)
            : elementRecord.attributes.push({
                name: qualifyName(attribute.uri, attribute.local),
                value: attribute.value,
                uri: attribute.uri,
                local: attribute.local,
              });
    (siblings.push(elementRecord), elementStack.push(elementRecord));
  });
  let appendText = (textChunk) => {
    let currentElement = elementStack.at(-1);
    if (
      currentElement &&
      ((currentElement.text += textChunk), currentElement.text.length > 32767)
    )
      throw Error("XML cell text budget exceeded.");
  };
  if (
    (parser.on("text", appendText),
    parser.on("cdata", appendText),
    parser.on("closetag", () => {
      elementStack.pop();
    }),
    parser.on("error", () => {
      throw Error("Malformed XML source.");
    }),
    parser.write(xmlText).close(),
    rootNodes.length !== 1)
  )
    throw Error("XML needs one root element.");
  let tableGroups = [],
    collectTableGroups = (ownerElement) => {
      let nodesByName = new Map();
      for (let child of ownerElement.children)
        nodesByName.set(child.name, [
          ...(nodesByName.get(child.name) ?? []),
          child,
        ]);
      for (let [childName, nodes] of nodesByName)
        nodes.length > 1
          ? tableGroups.push({
              path: ownerElement.path + "/" + childName,
              nodes: nodes,
            })
          : collectTableGroups(nodes[0]);
    };
  if (
    (collectTableGroups(rootNodes[0]),
    tableGroups.length ||
      tableGroups.push({ path: rootNodes[0].path, nodes: [rootNodes[0]] }),
    tableGroups.length > sourceImportLimits.tables)
  )
    throw Error("XML table count exceeded.");
  let toTypedValue = (xmlElement) => {
    let text = xmlElement.text,
      valueType = "string",
      lexicalText = xmlElement.text,
      xsdType = xmlElement.type?.replace(
        "{http://www.w3.org/2001/XMLSchema}",
        "",
      ),
      trimmedText = xmlElement.text.trim();
    if (xmlElement.nil !== undefined) {
      if (!["true", "1", "false", "0"].includes(xmlElement.nil))
        throw Error("Invalid XML nil value.");
      if (["true", "1"].includes(xmlElement.nil)) {
        if (trimmedText || xmlElement.children.length)
          throw Error("A nil XML field cannot contain a value.");
        ((text = null), (valueType = "null"));
      }
    }
    if (
      valueType !== "null" &&
      xmlElement.type?.startsWith("{http://www.w3.org/2001/XMLSchema}")
    ) {
      if (xsdType === "boolean") {
        if (!["true", "1", "false", "0"].includes(trimmedText))
          throw Error("Invalid typed XML boolean.");
        ((text = trimmedText === "true" || trimmedText === "1"),
          (valueType = "boolean"));
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
        ].includes(xsdType)
      ) {
        let isIntegerType = !["decimal", "double", "float"].includes(xsdType),
          lexicalPattern = isIntegerType
            ? /^[+-]?\d+$/
            : xsdType === "decimal"
              ? /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/
              : /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
        if (
          ((text = Number(trimmedText)),
          !lexicalPattern.test(trimmedText) ||
            !Number.isFinite(text) ||
            (Number.isInteger(text) && !Number.isSafeInteger(text)) ||
            (xsdType === "nonNegativeInteger" && text < 0) ||
            (xsdType === "positiveInteger" && text <= 0))
        )
          throw Error("Invalid or unsafe typed XML number.");
        let integerRanges = {
          int: [-2147483648, 2147483647],
          short: [-32768, 32767],
          byte: [-128, 127],
        };
        if (
          integerRanges[xsdType] &&
          (text < integerRanges[xsdType][0] || text > integerRanges[xsdType][1])
        )
          throw Error("Typed XML integer is out of range.");
        valueType = "number";
      }
    }
    return {
      value: text,
      type: valueType,
      path: xmlElement.path,
      lexical: lexicalText,
      ...(xmlElement.type ? { declaredType: xmlElement.type } : {}),
      confidence: 1,
    };
  };
  return tableGroups.map((tableGroup, tableIndex) => {
    let columnPathSet = new Set(),
      rowValueMaps = tableGroup.nodes.map((node) => {
        let valuesByColumn = new Map(),
          setValue = (columnPath, cellValue) => {
            (columnPathSet.add(columnPath),
              valuesByColumn.set(columnPath, cellValue));
          },
          flattenElement = (flatElement, basePath) => {
            for (let attribute of flatElement.attributes)
              setValue(basePath + "/@" + attribute.name, {
                value: attribute.value,
                type: "string",
                path: flatElement.path + "/@" + attribute.name,
                confidence: 1,
              });
            if (flatElement.children.length) {
              if (flatElement.text.trim())
                throw Error(
                  "Mixed-content XML tables need explicit extraction.",
                );
              let childCounts = new Map();
              for (let childElement of flatElement.children) {
                let childOrdinal =
                  (childCounts.get(childElement.name) ?? 0) + 1;
                (childCounts.set(childElement.name, childOrdinal),
                  flattenElement(
                    childElement,
                    basePath +
                      "/" +
                      childElement.name +
                      (flatElement.children.filter(
                        (sameNameChild) =>
                          sameNameChild.name === childElement.name,
                      ).length > 1
                        ? "[" + childOrdinal + "]"
                        : ""),
                  ));
              }
            } else setValue(basePath, toTypedValue(flatElement));
          };
        return (flattenElement(node, ""), valuesByColumn);
      }),
      columnPaths = [...columnPathSet];
    return {
      id: "xml-" + tableIndex,
      name: tableGroup.path.slice(-200),
      path: tableGroup.path,
      ambiguous: tableGroups.length > 1,
      columns: columnPaths.map((path) => ({
        name: (path || "Value").slice(-200),
        path: path,
      })),
      rows: rowValueMaps.map((rowValues, rowIndex) =>
        columnPaths.map(
          (column) =>
            rowValues.get(column) ?? {
              value: null,
              type: "null",
              path: tableGroup.nodes[rowIndex].path + column,
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
  let extension = fileName.split(".").at(-1)?.toLowerCase(),
    format =
      extension === "jpg" ? "jpeg" : extension === "htm" ? "html" : extension,
    tables = [],
    method = "document-adapter",
    extractor = "Host extraction required",
    warnings = [];
  if (format === "json" || format === "xml") {
    let sourceText = new TextDecoder("utf-8", { fatal: true })
      .decode(bytes)
      .replace(/^\uFEFF/, "");
    ((tables =
      format === "json"
        ? collectJsonSourceTables(parseStrictJsonText(sourceText))
        : parseXmlSourceTables(sourceText)),
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
    let binaryString = "";
    for (let byteOffset = 0; byteOffset < bytes.length; byteOffset += 8192)
      binaryString += String.fromCharCode(
        ...bytes.subarray(byteOffset, byteOffset + 8192),
      );
    (inspectPictureDataUrl(`data:image/${format};base64,${btoa(binaryString)}`),
      warnings.push(
        "Image extraction requires a host OCR adapter or matching extraction JSON.",
      ));
  } else if (format === "html") {
    let htmlText = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    if (!/^\s*(?:<!doctype\s+html[^>]*>\s*)?<(?:html|table)\b/i.test(htmlText))
      throw Error("Source extension does not match an HTML document.");
    warnings.push(
      "HTML is retained as inert bytes. No markup, scripts or resource links are rendered or fetched. A host table extractor or matching extraction JSON is required.",
    );
  } else throw Error("Choose JSON, XML, PDF, PNG, JPEG or HTML.");
  let inspection = validateExtractedSourceTables({
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
  return (recordSourceInspection(inspection, bytes, fileName), inspection);
}

function inspectCompoundFileContainer(bytes) {
  let failUnsupported = () => {
      throw Error(
        "Unsupported, encrypted or corrupt legacy workbook. Open the original in its source application and save an unencrypted XLSX copy.",
      );
    },
    view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength),
    endOfChain = 4294967294,
    freeSector = 4294967295;
  if (
    bytes.length < 512 ||
    view.getUint32(0, true) !== 3759263696 ||
    view.getUint32(4, true) !== 3776623009 ||
    view.getUint16(28, true) !== 65534
  )
    return failUnsupported();
  let majorVersion = view.getUint16(26, true),
    sectorShift = view.getUint16(30, true),
    sectorSize = 2 ** sectorShift;
  if (
    !(
      (majorVersion === 3 && sectorShift === 9) ||
      (majorVersion === 4 && sectorShift === 12)
    ) ||
    view.getUint16(32, true) !== 6 ||
    view.getUint32(56, true) !== 4096 ||
    bytes.length % sectorSize
  )
    return failUnsupported();
  let maxSectorIndex = bytes.length / sectorSize - 1,
    readSector = (sectorIndex) =>
      !Number.isInteger(sectorIndex) ||
      sectorIndex < 0 ||
      sectorIndex >= maxSectorIndex
        ? failUnsupported()
        : bytes.subarray(
            (sectorIndex + 1) * sectorSize,
            (sectorIndex + 2) * sectorSize,
          ),
    fatSectorIds = [],
    visitedDifatSectors = new Set();
  for (let difatIndex = 0; difatIndex < 109; difatIndex++) {
    let headerFatSectorId = view.getUint32(76 + difatIndex * 4, true);
    headerFatSectorId !== freeSector && fatSectorIds.push(headerFatSectorId);
  }
  let difatSectorId = view.getUint32(68, true),
    difatSectorCount = view.getUint32(72, true);
  if (difatSectorCount > maxSectorIndex) return failUnsupported();
  for (
    let difatSectorIndex = 0;
    difatSectorIndex < difatSectorCount;
    difatSectorIndex++
  ) {
    if (visitedDifatSectors.has(difatSectorId)) return failUnsupported();
    visitedDifatSectors.add(difatSectorId);
    let difatSectorBytes = readSector(difatSectorId),
      difatSectorView = new DataView(
        difatSectorBytes.buffer,
        difatSectorBytes.byteOffset,
        difatSectorBytes.byteLength,
      );
    for (
      let difatEntryOffset = 0;
      difatEntryOffset < sectorSize - 4;
      difatEntryOffset += 4
    ) {
      let difatFatSectorId = difatSectorView.getUint32(difatEntryOffset, true);
      difatFatSectorId !== freeSector && fatSectorIds.push(difatFatSectorId);
    }
    difatSectorId = difatSectorView.getUint32(sectorSize - 4, true);
  }
  if (
    (difatSectorCount && difatSectorId !== endOfChain) ||
    fatSectorIds.length !== view.getUint32(44, true) ||
    !fatSectorIds.length ||
    fatSectorIds.length > maxSectorIndex ||
    new Set(fatSectorIds).size !== fatSectorIds.length
  )
    return failUnsupported();
  let fatEntries = [];
  for (let fatSectorId of fatSectorIds) {
    let fatSectorBytes = readSector(fatSectorId),
      fatSectorView = new DataView(
        fatSectorBytes.buffer,
        fatSectorBytes.byteOffset,
        fatSectorBytes.byteLength,
      );
    for (
      let fatEntryOffset = 0;
      fatEntryOffset < sectorSize;
      fatEntryOffset += 4
    )
      fatEntries.push(fatSectorView.getUint32(fatEntryOffset, true));
  }
  for (let checkedFatSectorId of fatSectorIds)
    if (fatEntries[checkedFatSectorId] !== 4294967293) return failUnsupported();
  for (let checkedDifatSectorId of visitedDifatSectors)
    if (fatEntries[checkedDifatSectorId] !== 4294967292)
      return failUnsupported();
  let readChain = (
      startSector,
      sectorTable,
      readChunk,
      maxSectors,
      expectedSize,
    ) => {
      let visitedSectors = new Set(),
        chunks = [],
        currentSector = startSector,
        totalLength = 0;
      for (; currentSector !== endOfChain;) {
        if (
          currentSector >= sectorTable.length ||
          visitedSectors.has(currentSector) ||
          visitedSectors.size >= maxSectors ||
          (fatSectorIds.includes(currentSector) &&
            sectorTable === fatEntries) ||
          (visitedDifatSectors.has(currentSector) && sectorTable === fatEntries)
        )
          return failUnsupported();
        visitedSectors.add(currentSector);
        let chunk = readChunk(currentSector);
        if (
          (chunks.push(chunk),
          (totalLength += chunk.length),
          (currentSector = sectorTable[currentSector]),
          totalLength > bytes.length)
        )
          return failUnsupported();
      }
      if (
        expectedSize !== undefined &&
        (expectedSize > totalLength ||
          totalLength - expectedSize >= (chunks[0]?.length ?? 1))
      )
        return failUnsupported();
      let output = new Uint8Array(expectedSize ?? totalLength),
        writeOffset = 0;
      for (let chunkPart of chunks)
        (output.set(
          chunkPart.subarray(0, output.length - writeOffset),
          writeOffset,
        ),
          (writeOffset = Math.min(
            writeOffset + chunkPart.length,
            output.length,
          )));
      return output;
    },
    directoryBytes = readChain(
      view.getUint32(48, true),
      fatEntries,
      readSector,
      maxSectorIndex,
    ),
    directoryView = new DataView(directoryBytes.buffer);
  if (!directoryBytes.length || directoryBytes.length % 128)
    return failUnsupported();
  let directoryEntries = [];
  for (
    let directoryEntryOffset = 0;
    directoryEntryOffset < directoryBytes.length;
    directoryEntryOffset += 128
  ) {
    let entryType = directoryView.getUint8(directoryEntryOffset + 66),
      nameByteLength = directoryView.getUint16(directoryEntryOffset + 64, true),
      sizeHigh = directoryView.getUint32(directoryEntryOffset + 124, true);
    if (
      ![0, 1, 2, 5].includes(entryType) ||
      (entryType &&
        (nameByteLength < 2 ||
          nameByteLength > 64 ||
          nameByteLength % 2 ||
          directoryView.getUint16(
            directoryEntryOffset + nameByteLength - 2,
            true,
          )))
    )
      return failUnsupported();
    let entryName = entryType
      ? new TextDecoder("utf-16le", { fatal: true }).decode(
          directoryBytes.subarray(
            directoryEntryOffset,
            directoryEntryOffset + nameByteLength - 2,
          ),
        )
      : "";
    if (
      entryType &&
      (!entryName ||
        /[\x00-\x1f]/.test(entryName.replace(/^\x05/, "")) ||
        (sizeHigh && majorVersion === 4))
    )
      return failUnsupported();
    directoryEntries.push({
      name: entryName,
      type: entryType,
      left: directoryView.getUint32(directoryEntryOffset + 68, true),
      right: directoryView.getUint32(directoryEntryOffset + 72, true),
      child: directoryView.getUint32(directoryEntryOffset + 76, true),
      start: directoryView.getUint32(directoryEntryOffset + 116, true),
      size: directoryView.getUint32(directoryEntryOffset + 120, true),
    });
  }
  if (
    directoryEntries[0].type !== 5 ||
    directoryEntries.some((entry) =>
      /^(EncryptedPackage|EncryptionInfo)$/i.test(entry.name),
    )
  )
    return failUnsupported();
  let reachableEntries = [],
    visitedEntryIds = new Set(),
    entryIdStack = [directoryEntries[0].child];
  for (; entryIdStack.length;) {
    let entryId = entryIdStack.pop();
    if (entryId === freeSector) continue;
    if (
      !entryId ||
      entryId >= directoryEntries.length ||
      visitedEntryIds.has(entryId)
    )
      return failUnsupported();
    visitedEntryIds.add(entryId);
    let directoryEntry = directoryEntries[entryId];
    if (![1, 2].includes(directoryEntry.type)) return failUnsupported();
    (reachableEntries.push(directoryEntry),
      entryIdStack.push(directoryEntry.left, directoryEntry.right));
  }
  let workbookEntries = reachableEntries.filter(
    (candidate) =>
      candidate.type === 2 && /^(Workbook|Book)$/i.test(candidate.name),
  );
  if (workbookEntries.length !== 1) return failUnsupported();
  let workbookEntry = workbookEntries[0];
  if (workbookEntry.size < 8 || workbookEntry.size > bytes.length)
    return failUnsupported();
  let workbookStream;
  if (workbookEntry.size >= 4096)
    workbookStream = readChain(
      workbookEntry.start,
      fatEntries,
      readSector,
      maxSectorIndex,
      workbookEntry.size,
    );
  else {
    let miniFatSectorCount = view.getUint32(64, true);
    if (!miniFatSectorCount || miniFatSectorCount > maxSectorIndex)
      return failUnsupported();
    let miniFatBytes = readChain(
        view.getUint32(60, true),
        fatEntries,
        readSector,
        miniFatSectorCount,
        miniFatSectorCount * sectorSize,
      ),
      miniFatView = new DataView(miniFatBytes.buffer),
      miniFatEntries = [];
    for (
      let miniFatOffset = 0;
      miniFatOffset < miniFatBytes.length;
      miniFatOffset += 4
    )
      miniFatEntries.push(miniFatView.getUint32(miniFatOffset, true));
    let rootStreamBytes = readChain(
      directoryEntries[0].start,
      fatEntries,
      readSector,
      maxSectorIndex,
      directoryEntries[0].size,
    );
    workbookStream = readChain(
      workbookEntry.start,
      miniFatEntries,
      (miniSectorIndex) =>
        (miniSectorIndex + 1) * 64 > rootStreamBytes.length
          ? failUnsupported()
          : rootStreamBytes.subarray(
              miniSectorIndex * 64,
              (miniSectorIndex + 1) * 64,
            ),
      Math.ceil(rootStreamBytes.length / 64),
      workbookEntry.size,
    );
  }
  let workbookView = new DataView(workbookStream.buffer),
    biffVersion = workbookView.getUint16(4, true);
  if (
    workbookView.getUint16(0, true) !== 2057 ||
    workbookView.getUint16(2, true) < 4 ||
    ![1280, 1536].includes(biffVersion) ||
    workbookView.getUint16(6, true) !== 5
  )
    return failUnsupported();
  let hasEofRecord = false;
  for (let recordOffset = 0; recordOffset + 4 <= workbookStream.length;) {
    let recordType = workbookView.getUint16(recordOffset, true),
      recordLength = workbookView.getUint16(recordOffset + 2, true);
    if (
      recordOffset + 4 + recordLength > workbookStream.length ||
      recordType === 47
    )
      return failUnsupported();
    if (((recordOffset += 4 + recordLength), recordType === 10)) {
      hasEofRecord = true;
      break;
    }
  }
  return hasEofRecord
    ? {
        paths: directoryEntries
          .filter((allocatedEntry) => allocatedEntry.type)
          .map((pathEntry) => pathEntry.name),
        workbookBytes: workbookStream.length,
        version: biffVersion === 1536 ? "BIFF8" : "BIFF5",
      }
    : failUnsupported();
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

function inspectAlternateWorkbookFile(bytes, fileName) {
  if (
    !fileName ||
    fileName.length > 255 ||
    /[\\/\x00-\x1f\x7f\u202a-\u202e\u2066-\u2069]/.test(fileName) ||
    !bytes.length ||
    bytes.length > workbookLimits.maxFileBytes
  )
    throw Error(
      "Invalid alternate workbook filename or 10 MB file limit exceeded.",
    );
  let format = fileName.split(".").pop().toLowerCase();
  if (!alternateFormatExtensions.includes(format))
    throw Error("Choose an XLS, XLSB, XLSM, XLTX, XLTM or ODS source.");
  let detectedFormat, identity, parts;
  if (bytes[0] === 208 && bytes[1] === 207) {
    let containerInspection = inspectCompoundFileContainer(bytes);
    ((detectedFormat = "xls"),
      (identity =
        containerInspection.version +
        " workbook in a CFB container; binary cells not decoded"),
      (parts = containerInspection.paths.map((path) => ({
        path: path,
        bytes: /^(Workbook|Book)$/i.test(path)
          ? containerInspection.workbookBytes
          : 0,
        disposition: "original-only",
      }))));
  } else {
    let archiveFiles = inspectAlternateWorkbookArchive(bytes);
    ((detectedFormat = readOpenDocumentPackage(archiveFiles)),
      (identity =
        detectedFormat === "ods"
          ? "ODF spreadsheet MIME, manifest and content"
          : "OOXML main content type and package relationship"));
    for (let [archivePath, archiveBytes] of Object.entries(archiveFiles))
      /\.(xml|rels)$/i.test(archivePath) &&
        parseAlternateWorkbookXml(archiveBytes, () => {});
    parts = Object.entries(archiveFiles).map(([path, partBytes]) => ({
      path: path,
      bytes: partBytes.length,
      sha256: sha256HexDigest(partBytes),
      disposition: "original-only",
    }));
  }
  if (detectedFormat !== format)
    throw Error(
      "The filename extension does not match the workbook content. Restore its original extension or convert it in the source application; renaming is not conversion.",
    );
  let inspection = {
    schemaVersion: 1,
    fileName: fileName,
    format: format,
    fileBytes: bytes.length,
    sha256: sha256HexDigest(bytes),
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
  return (
    validateIsolatedFormatInspection(inspection, bytes, fileName),
    inspection
  );
}
var saxesEsModuleForLinks = toEsModule(getSaxesModule(), 1);
function parseExternalLinkTargets(archiveFiles, sha256) {
  let readXmlPart = (partPath, onOpenTag) => {
      if (!archiveFiles[partPath]) return;
      let xmlParser = new saxesEsModuleForLinks.SaxesParser({ xmlns: true });
      (xmlParser.on("opentag", onOpenTag),
        xmlParser.on("doctype", () => {
          throw Error("External link metadata cannot contain DTDs.");
        }),
        xmlParser.write(Ao(archiveFiles[partPath])).close());
    },
    linkTargetsById = new Map();
  readXmlPart("xl/_rels/workbook.xml.rels", (relationshipElement) => {
    relationshipElement.local === "Relationship" &&
      relationshipElement.attributes.Type?.value.endsWith("/externalLink") &&
      relationshipElement.attributes.TargetMode?.value !== "External" &&
      linkTargetsById.set(
        relationshipElement.attributes.Id.value,
        resolveRelationshipTarget(
          "xl/_rels/workbook.xml.rels",
          relationshipElement.attributes.Target.value,
        ),
      );
  });
  let externalReferenceIds = [];
  if (
    (readXmlPart("xl/workbook.xml", (workbookElement) => {
      if (workbookElement.local === "externalReference") {
        let referenceId = Object.values(workbookElement.attributes).find(
          (attribute) => attribute.local === "id",
        )?.value;
        referenceId && externalReferenceIds.push(referenceId);
      }
    }),
    externalReferenceIds.length > 100)
  )
    throw Error(
      "More than 100 external workbook targets require an explicit source conversion.",
    );
  let externalTargets = [];
  return (
    externalReferenceIds.forEach((externalReferenceId, referenceIndex) => {
      let part = linkTargetsById.get(externalReferenceId);
      if (!part) return;
      let lastSlashIndex = part.lastIndexOf("/"),
        partRelsPath =
          part.slice(0, lastSlashIndex + 1) +
          "_rels/" +
          part.slice(lastSlashIndex + 1) +
          ".rels",
        target;
      (readXmlPart(partRelsPath, (relsElement) => {
        if (
          relsElement.local === "Relationship" &&
          relsElement.attributes.Type?.value.endsWith("/externalLinkPath") &&
          relsElement.attributes.TargetMode?.value === "External"
        ) {
          if (target !== undefined)
            throw Error("Ambiguous external workbook target relationship.");
          target = relsElement.attributes.Target?.value;
        }
      }),
        target &&
          externalTargets.push({
            key: `[${referenceIndex + 1}]`,
            target: target,
            part: part,
            sourceSha256: sha256,
          }));
    }),
    externalTargets
  );
}

async function convertXltxTemplateFile(templateFile) {
  if (templateFile.size > 10 * 1024 * 1024)
    throw Error("Files must be 10 MB or smaller.");
  let templateBytes = new Uint8Array(await templateFile.arrayBuffer()),
    alternateInspection = inspectAlternateWorkbookFile(
      templateBytes,
      templateFile.name,
    );
  if (alternateInspection.format !== "xltx")
    throw Error("Only XLTX templates use the native template adapter.");
  let convertedXlsxBytes = convertTemplateToXlsx(templateBytes, "xltx", "xlsx"),
    convertedFileName =
      templateFile.name.replace(/\.xltx$/i, "").slice(0, 236) +
      ".template-copy.xlsx",
    importResult = await importWorkbookFile({
      name: convertedFileName,
      size: convertedXlsxBytes.length,
      arrayBuffer: async () => convertedXlsxBytes.buffer,
    });
  ((importResult.workbook.name =
    templateFile.name.replace(/\.xltx$/i, "").slice(0, 200) ||
    "Template workbook"),
    (importResult.workbook.alternateSource = retainAlternateFormatSource(
      templateBytes,
      templateFile.name,
      "xltx",
      {
        method: "template-adapter",
        convertedFileName: convertedFileName,
        convertedSha256: sha256HexDigest(convertedXlsxBytes),
        reviewedAt: new Date().toISOString(),
        equivalence: "not-verified",
      },
    )));
  let templateWarning =
    "A new workbook was created from the XLTX template using the supported XLSX subset. The exact template is retained under File > Workbook formats. The XLSX part inventory refers to the converted intermediate, not the original template. Template fidelity requires comparison in the source application.";
  return (
    (importResult.warnings = [
      ...new Set([...importResult.warnings, templateWarning]),
    ]),
    (importResult.report.warnings = importResult.warnings),
    (importResult.report.alternateOriginal = {
      fileName: alternateInspection.fileName,
      byteLength: alternateInspection.fileBytes,
      sha256: alternateInspection.sha256,
      format: alternateInspection.format,
    }),
    (importResult.report.fileName = templateFile.name),
    (importResult.report.fileBytes = templateBytes.length),
    (importResult.report.format = "xltx"),
    (importResult.report.workbookName = importResult.workbook.name),
    importResult
  );
}

var argbColorOrFallback = (color, fallbackColor) =>
  color?.argb && /^[a-f\d]{8}$/i.test(color.argb)
    ? `#${color.argb.slice(-6).toLowerCase()}`
    : fallbackColor;
function collapseCellAddressesToRanges(addresses) {
  let sortedCells = addresses
      .map(_t)
      .sort(
        (leftCell, rightCell) =>
          leftCell.row - rightCell.row || leftCell.col - rightCell.col,
      ),
    ranges = [];
  for (let cell of sortedCells) {
    let previousRange = ranges.at(-1);
    previousRange &&
    previousRange.lastRow === cell.row &&
    previousRange.lastCol + 1 === cell.col
      ? (previousRange.lastCol = cell.col)
      : ranges.push({
          firstRow: cell.row,
          lastRow: cell.row,
          firstCol: cell.col,
          lastCol: cell.col,
        });
  }
  let mergedRanges = [],
    lastRangeByColumns = new Map();
  for (let range of ranges) {
    let columnKey = `${range.firstCol}:${range.lastCol}`,
      sameColumnRange = lastRangeByColumns.get(columnKey);
    sameColumnRange && sameColumnRange.lastRow + 1 === range.firstRow
      ? (sameColumnRange.lastRow = range.lastRow)
      : (mergedRanges.push(range), lastRangeByColumns.set(columnKey, range));
  }
  return mergedRanges.map(
    (finalRange) =>
      `${Ks(finalRange.firstRow, finalRange.firstCol)}:${Ks(finalRange.lastRow, finalRange.lastCol)}`,
  );
}

function buildDataValidationRule(validation) {
  let kind = {
      list: "list",
      decimal: "number",
      whole: "integer",
      textLength: "textLength",
      date: "date",
      time: "time",
      custom: "custom",
    }[validation.type],
    operatorName = Object.entries(dataValidationOperators).find(
      ([, operatorValue]) => operatorValue === validation.operator,
    )?.[0];
  if (!kind || (!operatorName && !["list", "custom"].includes(kind)))
    return null;
  let formulae = validation.formulae ?? [],
    toNumericValue = (operand) =>
      operand instanceof Date
        ? (operand.getTime() - Date.UTC(1899, 11, 30)) / 864e5
        : typeof operand == "number"
          ? operand
          : kind === "time" &&
              typeof operand == "string" &&
              /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(operand)
            ? Number(operand)
            : NaN,
    min = ["list", "custom"].includes(kind) ? 0 : toNumericValue(formulae[0]),
    max = ["list", "custom"].includes(kind)
      ? 0
      : ["between", "notBetween"].includes(operatorName)
        ? toNumericValue(formulae[1])
        : min;
  if (!Number.isFinite(min) || !Number.isFinite(max)) return null;
  let list = [],
    formula;
  if (
    (kind === "list" &&
      (typeof formulae[0] != "string" ||
        !/^"[^"]*"$/.test(formulae[0]) ||
        ((list = formulae[0].slice(1, -1).split(",")),
        !list.length ||
          list.length > 200 ||
          list.some(
            (listItem) => !listItem.length || listItem.length > 1e3,
          )))) ||
    (kind === "custom" &&
      (typeof formulae[0] != "string" ||
        ((formula = "=" + formulae[0].replace(/^=/, "")),
        formula.length > 8192 ||
          formula.length < 2 ||
          isExternalReferenceFormula(formula))))
  )
    return null;
  let errorStyle = validation.errorStyle?.toLowerCase() ?? "stop";
  return ["stop", "warning", "information"].includes(errorStyle)
    ? {
        kind: kind,
        operator: operatorName ?? "equal",
        min: min,
        max: max,
        list: list,
        allowBlank: !!validation.allowBlank,
        message: String(validation.error ?? "Choose a valid value.").slice(
          0,
          500,
        ),
        errorStyle: errorStyle,
        errorTitle: String(validation.errorTitle ?? "Invalid value").slice(
          0,
          100,
        ),
        showInputMessage: !!validation.showInputMessage,
        inputTitle: String(validation.promptTitle ?? "").slice(0, 100),
        inputMessage: String(validation.prompt ?? "").slice(0, 500),
        ...(formula ? { formula: formula } : {}),
      }
    : null;
}

function buildConditionalFormatRule(rule, range) {
  let fill =
      rule.style?.fill?.type === "pattern"
        ? argbColorOrFallback(rule.style.fill.fgColor, "#fff2cc")
        : "#fff2cc",
    color = argbColorOrFallback(rule.style?.font?.color, "#000000"),
    baseRule = {
      id: GN(),
      priority: Number.isFinite(rule.priority) ? rule.priority : 1e4,
      range: formatRangeCornerCells(range),
      value: "",
      second: "",
      fill: fill,
      color: color,
    },
    normalizeOperand = (operand) =>
      typeof operand == "number" && Number.isFinite(operand)
        ? String(operand)
        : typeof operand == "string" &&
            operand.trim() !== "" &&
            Number.isFinite(Number(operand))
          ? operand
          : typeof operand == "string" && /^"(?:[^"]|"")*"$/.test(operand)
            ? operand.slice(1, -1).replace(/""/g, '"')
            : null;
  if (
    rule.type === "containsText" &&
    (!rule.operator || rule.operator === "containsText")
  ) {
    let rangeAddress = parseRangeAddress(range),
      topLeftAddress = Ks(rangeAddress.firstRow, rangeAddress.firstCol),
      searchText = typeof rule.text == "string" ? rule.text : "",
      firstFormula = rule.formulae?.[0],
      formulaText =
        typeof firstFormula == "string" && firstFormula.trim()
          ? firstFormula.replace(/^=/, "")
          : "NOT(ISERROR(SEARCH(" +
            quoteFormulaString(searchText) +
            "," +
            topLeftAddress +
            ")))";
    return formulaText.length > 4095 ||
      isExternalReferenceFormula("=" + formulaText)
      ? null
      : { ...baseRule, kind: "formula", formula: "=" + formulaText };
  }
  if (
    rule.type === "colorScale" &&
    rule.cfvo?.length === 2 &&
    rule.cfvo[0].type === "min" &&
    rule.cfvo[1].type === "max"
  )
    return {
      ...baseRule,
      kind: "colorScale",
      minColor: argbColorOrFallback(rule.color?.[0], "#f8e99b"),
      maxColor: argbColorOrFallback(rule.color?.[1], "#63be7b"),
    };
  if (
    rule.type === "dataBar" &&
    rule.cfvo?.length === 2 &&
    ["min", "autoMin"].includes(rule.cfvo[0].type) &&
    ["max", "autoMax"].includes(rule.cfvo[1].type) &&
    rule.showValue !== false &&
    rule.direction !== "rightToLeft"
  )
    return {
      ...baseRule,
      kind: "dataBar",
      fill: argbColorOrFallback(rule.color, "#4472c4"),
    };
  if (
    rule.type === "iconSet" &&
    rule.iconSet === "3Arrows" &&
    !rule.reverse &&
    !rule.custom &&
    rule.showValue !== false &&
    rule.cfvo?.length === 3 &&
    rule.cfvo[1].type === rule.cfvo[2].type &&
    ["num", "percent"].includes(rule.cfvo[1].type)
  ) {
    let [firstThreshold, middleThreshold, thirdThreshold] = rule.cfvo,
      value = middleThreshold.value,
      upperValue = thirdThreshold.value;
    return !(
      firstThreshold.type === "min" ||
      (firstThreshold.type === "percent" && firstThreshold.value === 0)
    ) ||
      typeof value != "number" ||
      typeof upperValue != "number" ||
      !Number.isFinite(value) ||
      !Number.isFinite(upperValue) ||
      value >= upperValue ||
      (middleThreshold.type === "percent" && (value < 0 || upperValue > 100))
      ? null
      : {
          ...baseRule,
          kind: "iconSet",
          thresholdType: middleThreshold.type === "num" ? "number" : "percent",
          iconThresholds: [value, upperValue],
        };
  }
  if (rule.type === "expression") {
    let expressionText = String(rule.formulae?.[0] ?? "").replace(/^=/, "");
    if (
      !expressionText ||
      expressionText.length > 4095 ||
      isExternalReferenceFormula("=" + expressionText)
    )
      return null;
    if (expressionText.toUpperCase() === Wle("duplicate", range).toUpperCase())
      return { ...baseRule, kind: "duplicate" };
    if (expressionText.toUpperCase() === Wle("unique", range).toUpperCase())
      return { ...baseRule, kind: "unique" };
    let containsMatch =
      /^ISNUMBER\(FIND\(LOWER\("((?:[^"]|"")*)"\),LOWER\(([A-Z]+[1-9]\d*)&""\)\)\)$/i.exec(
        expressionText,
      );
    if (containsMatch) {
      let containsValue = containsMatch[1].replace(/""/g, '"'),
        rangeAddress = parseRangeAddress(range),
        cellAddress = Ks(rangeAddress.firstRow, rangeAddress.firstCol);
      if (
        containsValue.length <= 1e3 &&
        expressionText.toUpperCase() ===
          buildContainsTextFormula(containsValue, cellAddress).toUpperCase()
      )
        return { ...baseRule, kind: "contains", value: containsValue };
    }
    let rangeAddress = parseRangeAddress(range),
      anchorAddress = Ks(rangeAddress.firstRow, rangeAddress.firstCol),
      kindByFormula = {
        [`LEN(${anchorAddress})=0`]: "blank",
        [`LEN(${anchorAddress})>0`]: "notBlank",
        [`ISERROR(${anchorAddress})`]: "error",
        [`NOT(ISERROR(${anchorAddress}))`]: "notError",
      };
    return kindByFormula[expressionText]
      ? { ...baseRule, kind: kindByFormula[expressionText] }
      : { ...baseRule, kind: "formula", formula: "=" + expressionText };
  }
  if (rule.type !== "cellIs") return null;
  let comparisonKind = {
      greaterThan: "greater",
      lessThan: "less",
      equal: "equal",
      between: "between",
    }[rule.operator],
    firstOperand = normalizeOperand(rule.formulae?.[0]),
    second =
      comparisonKind === "between" ? normalizeOperand(rule.formulae?.[1]) : "";
  return !comparisonKind || firstOperand === null || second === null
    ? null
    : {
        ...baseRule,
        kind: comparisonKind,
        value: firstOperand,
        second: second,
      };
}
function convertWorksheetFeatures(
  worksheet,
  targetSheet,
  stopIfTrueByPriority = {},
  validationAreas = [],
) {
  let warnings = [],
    sheet = {},
    merges = worksheet.model.merges ?? [];
  merges.length && (sheet.merges = merges.map(formatRangeCornerCells));
  let model = worksheet.dataValidations.model,
    groupedValidations = new Map();
  for (let [address, rule] of Object.entries(model)) {
    let validationCell = _t(address),
      area = validationAreas.find((candidateArea) => {
        let rangeAddress = parseRangeAddress(candidateArea.range);
        return (
          validationCell.row >= rangeAddress.firstRow &&
          validationCell.row <= rangeAddress.lastRow &&
          validationCell.col >= rangeAddress.firstCol &&
          validationCell.col <= rangeAddress.lastCol
        );
      }),
      groupKey = JSON.stringify({ rule: rule, area: area });
    (groupedValidations.has(groupKey) ||
      groupedValidations.set(groupKey, {
        rule: rule,
        addresses: [],
        area: area,
      }),
      groupedValidations.get(groupKey).addresses.push(address));
  }
  let validations = [];
  for (let { rule, addresses, area } of groupedValidations.values()) {
    let shouldUseAreaOperands = !["list", "custom"].includes(rule.type),
      hasNumericOperands = area?.operands.every((operand) =>
        /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(operand),
      ),
      builtRule =
        shouldUseAreaOperands && area && !hasNumericOperands
          ? null
          : buildDataValidationRule(
              shouldUseAreaOperands && area
                ? { ...rule, formulae: area.operands.map(Number) }
                : rule,
            );
    if (!builtRule) {
      warnings.push(
        "An unsupported validation formula, operator or external list was omitted; source cells were retained.",
      );
      continue;
    }
    area &&
      !area.alert &&
      warnings.push(
        "An imported validation has error alerts disabled. Local validation requires Stop or explicit Warning/Information review; source values were retained.",
      );
    for (let range of collapseCellAddressesToRanges(addresses)) {
      if (validations.length >= 100) {
        warnings.push(
          "Additional validation areas beyond the 100-rule limit were omitted; source cells were retained.",
        );
        break;
      }
      let anchorPosition = _t(area?.anchor ?? range.split(":")[0]),
        rangeStartPosition = _t(range.split(":")[0]);
      validations.push({
        ...builtRule,
        id: GN(),
        range: range,
        ...(area ? { showDropdown: area.dropdown } : {}),
        ...(builtRule.formula
          ? {
              formula: shiftFormulaReferences(
                builtRule.formula,
                rangeStartPosition.row - anchorPosition.row,
                rangeStartPosition.col - anchorPosition.col,
              ),
            }
          : {}),
      });
    }
  }
  validations.length && (sheet.validations = validations);
  let conditionalRules = [],
    usedPriorities = new Set(),
    scannedCellCount = 0;
  for (let conditionalFormatting of worksheet.conditionalFormattings ?? []) {
    let appliedRanges = conditionalFormatting.ref
      .trim()
      .split(/\s+/)
      .filter(Boolean);
    if (!appliedRanges.length)
      throw Error("A conditional formatting rule requires an applied range.");
    let rangeAddress = parseRangeAddress(appliedRanges[0]);
    for (let rule of conditionalFormatting.rules) {
      if (
        !Number.isSafeInteger(rule.priority) ||
        rule.priority < 1 ||
        usedPriorities.has(rule.priority)
      )
        throw Error("Invalid or duplicate conditional-format priority.");
      usedPriorities.add(rule.priority);
      for (let appliedRange of appliedRanges) {
        let appliedAddress = parseRangeAddress(appliedRange),
          effectiveRule = rule;
        (rule.type === "expression" || rule.type === "containsText") &&
          rule.formulae?.length &&
          (effectiveRule = {
            ...rule,
            formulae: rule.formulae.map((formulaItem) =>
              typeof formulaItem == "string"
                ? shiftFormulaReferences(
                    "=" + formulaItem.replace(/^=/, ""),
                    appliedAddress.firstRow - rangeAddress.firstRow,
                    appliedAddress.firstCol - rangeAddress.firstCol,
                  ).slice(1)
                : formulaItem,
            ),
          });
        let conditionalFormatRule = buildConditionalFormatRule(
          effectiveRule,
          appliedRange,
        );
        if (
          conditionalFormatRule &&
          (conditionalFormatRule.kind === "duplicate" ||
            conditionalFormatRule.kind === "unique")
        ) {
          let canUseNativeComparison = true;
          for (let [cellAddress, cell] of Object.entries(targetSheet.cells)) {
            if (++scannedCellCount > 1e6) {
              canUseNativeComparison = false;
              break;
            }
            let cellPosition = _t(cellAddress);
            if (
              cellPosition.row < appliedAddress.firstRow ||
              cellPosition.row > appliedAddress.lastRow ||
              cellPosition.col < appliedAddress.firstCol ||
              cellPosition.col > appliedAddress.lastCol
            )
              continue;
            let raw = cell.raw;
            if (
              (typeof cell.raw == "string" &&
                cell.raw.startsWith("=") &&
                !cell.literal &&
                ((raw = worksheet.getCell(cellAddress).result),
                raw === undefined &&
                  /^=\s*""\s*$/.test(cell.raw) &&
                  (raw = "")),
              !Hle(raw))
            ) {
              canUseNativeComparison = false;
              break;
            }
          }
          canUseNativeComparison ||
            ((conditionalFormatRule = {
              ...conditionalFormatRule,
              kind: "formula",
              formula:
                "=" +
                String(effectiveRule.formulae?.[0] ?? "").replace(/^=/, ""),
            }),
            warnings.push(
              "A duplicate/unique condition was retained as a formula because its population cannot use the native literal comparison; formula-engine limits apply.",
            ));
        }
        conditionalFormatRule && conditionalRules.length < 100
          ? conditionalRules.push({
              ...conditionalFormatRule,
              ...(stopIfTrueByPriority[String(rule.priority)] !== undefined
                ? { stopIfTrue: stopIfTrueByPriority[String(rule.priority)] }
                : {}),
            })
          : warnings.push(
              "An unsupported conditional-format rule or a rule beyond the 100-rule limit was omitted; base cell formatting was retained.",
            );
      }
    }
  }
  if (conditionalRules.length) {
    let sortedRules = [...conditionalRules].sort(
        (leftRule, rightRule) =>
          (leftRule.priority ?? 1e4) - (rightRule.priority ?? 1e4),
      ),
      priorityById = new Map(
        sortedRules.map((sortedRule, ruleIndex) => [
          sortedRule.id,
          ruleIndex + 1,
        ]),
      );
    sheet.conditionalFormats = conditionalRules.map((formatRule) => ({
      ...formatRule,
      priority: priorityById.get(formatRule.id),
    }));
  }
  return (
    worksheet.model.sheetProtection?.sheet &&
      warnings.push(
        "Source Excel sheet-password protection is not imported. The local protection feature and host authorization are separate controls.",
      ),
    validateWorksheetFeatures(sheet, targetSheet),
    warnings
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

function decodeUtf8Strict(bytes) {
  try {
    let decodedText = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    if (decodedText.includes("\0")) throw Error();
    return decodedText;
  } catch {
    throw Error(
      "Automation source must be UTF-8 text without null characters. Export a UTF-8 source copy; binary VBA is inventory-only.",
    );
  }
}
function inspectAutomationArchive(bytes) {
  let packageEntries = inspectXlsxPackage(bytes, {
      allowBinaryWorkbookInventory: true,
    }),
    inflatedEntries = new Map(),
    totalInflatedBytes = 0,
    archiveError,
    seenNames = new Set(),
    unzipper = new gc((entry) => {
      if (!packageEntries.has(entry.name) || seenNames.has(entry.name)) {
        archiveError = Error("Automation archive directory mismatch.");
        return;
      }
      seenNames.add(entry.name);
      let chunks = [],
        entryByteCount = 0,
        runningCrc = -1;
      ((entry.ondata = (chunkError, chunk, isFinal) => {
        if (!archiveError) {
          if (chunkError) {
            archiveError = Error(`Corrupt automation archive: ${entry.name}.`);
            return;
          }
          if (
            ((entryByteCount += chunk.length),
            (totalInflatedBytes += chunk.length),
            entryByteCount > packageEntries.get(entry.name).size ||
              totalInflatedBytes > 50 * 1024 * 1024)
          ) {
            ((archiveError = Error(
              "Automation archive exceeds its expansion budget.",
            )),
              entry.terminate());
            return;
          }
          if (
            ((runningCrc = computeCrc32(runningCrc, chunk)),
            chunks.push(chunk),
            isFinal)
          ) {
            if (
              entryByteCount !== packageEntries.get(entry.name).size ||
              (runningCrc ^ -1) >>> 0 !== packageEntries.get(entry.name).crc
            ) {
              archiveError = Error(
                `Automation archive checksum mismatch: ${entry.name}.`,
              );
              return;
            }
            let entryBytes = new Uint8Array(entryByteCount),
              writeOffset = 0;
            for (let chunkPart of chunks)
              (entryBytes.set(chunkPart, writeOffset),
                (writeOffset += chunkPart.length));
            inflatedEntries.set(entry.name, entryBytes);
          }
        }
      }),
        entry.start());
    });
  unzipper.register(mc);
  try {
    unzipper.push(bytes, true);
  } catch {
    throw Error("Unable to inspect corrupt automation archive.");
  }
  if (archiveError) throw archiveError;
  if (inflatedEntries.size !== packageEntries.size)
    throw Error("Automation archive is incomplete.");
  return inflatedEntries;
}
function scanAutomationXmlMetadata(xmlText, partName, onElement) {
  if (/<\?xml[^>]*encoding\s*=\s*["'](?!utf-8|utf8)[^"']+/i.test(xmlText))
    throw Error(`Unsupported XML encoding in ${partName}.`);
  let parser = new saxesEsModuleForAutomation.SaxesParser({ xmlns: false }),
    elementDepth = 0,
    elementCount = 0;
  (parser.on("doctype", () => {
    throw Error("Automation inventory does not accept XML document types.");
  }),
    parser.on("opentag", (tag) => {
      if (++elementDepth > 128 || ++elementCount > 1e6)
        throw Error("Automation inventory XML exceeds its bounds.");
      onElement(tag.name.split(":").at(-1), tag.attributes);
    }),
    parser.on("closetag", () => elementDepth--),
    parser.on("error", (parseError) => {
      throw Error(
        `Malformed automation metadata in ${partName}: ${parseError.message}`,
      );
    }),
    parser.write(xmlText).close());
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
  let fileBytes = bytes.slice(),
    inventory = {
      schemaVersion: 1,
      fileName: fileName,
      kind: "excel-container",
      fileBytes: fileBytes.length,
      sha256: sha256HexDigest(fileBytes),
      retainedEligible: fileBytes.length <= automationLimits.maxArtifactBytes,
      execution: "inventory-only",
      items: [],
      warnings: [
        "Imported automation never runs on open, preservation or native JSON reopen. Only a separately reviewed migration can issue bounded workbook commands.",
        "Preserved artifacts are inert original bytes in native JSON. VBA/XLM binaries are not executed or decompiled; retain originals and use an authorized source application to export source modules.",
      ],
    },
    extension = fileName.split(".").at(-1)?.toLowerCase();
  if (["bas", "vba", "cls", "ts", "osts"].includes(extension ?? "")) {
    if (fileBytes.length > automationLimits.maxSourceBytes)
      throw Error("Text automation source is limited to 256 KiB.");
    let sourceText = decodeUtf8Strict(fileBytes),
      language =
        extension === "ts" || extension === "osts" ? "office-script" : "vba";
    if (
      ((inventory.kind =
        extension === "osts" ? "office-script-container" : language),
      extension === "osts")
    )
      try {
        let envelope = JSON.parse(sourceText);
        if (
          !envelope ||
          typeof envelope != "object" ||
          Array.isArray(envelope) ||
          typeof envelope.body != "string"
        )
          throw Error();
        if (
          ((sourceText = envelope.body),
          sourceText.includes("\0") ||
            new TextEncoder().encode(sourceText).length >
              automationLimits.maxSourceBytes)
        )
          throw Error("Source body is outside the supported text budget.");
      } catch {
        (inventory.warnings.push(
          "Unsupported .osts envelope. Preserve the original or export its code as UTF-8 .ts. This adapter recognizes only a JSON object with a string body; it makes no universal .osts schema claim.",
        ),
          inventory.items.push({
            path: fileName,
            kind: "Office Script container",
            bytes: fileBytes.length,
            sha256: inventory.sha256,
            execution: "unsupported",
            note: "Unknown source envelope; no code inferred.",
          }));
      }
    inventory.items.length ||
      ((inventory.script = { language: language, text: sourceText }),
      (inventory.execution = "reviewed-migration-only"),
      inventory.items.push({
        path: fileName,
        kind: language === "vba" ? "VBA source" : "Office Script source",
        bytes: fileBytes.length,
        sha256: inventory.sha256,
        execution: "migration-candidate",
        note: "Compatibility scan required. Only literal range operations are eligible; this is not an equivalent VBA or Office Scripts runtime.",
      }));
  } else if (
    fileBytes[0] === 208 &&
    fileBytes[1] === 207 &&
    fileBytes[2] === 17 &&
    fileBytes[3] === 224
  )
    ((inventory.kind = "legacy-container"),
      inventory.items.push({
        path: fileName,
        kind: "OLE / encrypted or legacy container",
        bytes: fileBytes.length,
        sha256: inventory.sha256,
        execution: "unsupported",
        note: "VBA/XLM presence cannot be determined without a supported decoder. Export a clean OOXML workbook and UTF-8 source modules in an authorized application.",
      }));
  else if (
    ["xlsm", "xlam", "xltm", "xlsx", "xltx"].includes(extension ?? "") ||
    (fileBytes[0] === 80 && fileBytes[1] === 75)
  ) {
    let archiveParts = inspectAutomationArchive(fileBytes),
      flaggedParts = new Map(),
      markPart = (partPath, contentType) => {
        let partKind = "";
        (/vbaProject/i.test(contentType) ||
        /vbaproject(?:signature)?\.bin$/i.test(partPath)
          ? (partKind = "VBA binary")
          : /(?:intl)?macrosheet/i.test(contentType) ||
              /(?:^|\/)macrosheets\//i.test(partPath)
            ? (partKind = "XLM macro sheet")
            : /office.?script|scriptLink|scriptReference/i.test(contentType) ||
                /\.(?:osts|ts)$/i.test(partPath)
              ? (partKind = "Office Script artifact/reference")
              : (/activeX|oleObject|relationships\/(?:package|control)$/i.test(
                  contentType,
                ) ||
                  /\/(?:activeX|embeddings)\//i.test(partPath)) &&
                (partKind = "ActiveX / embedded object"),
          partKind &&
            flaggedParts.set(partPath, {
              kind: partKind,
              note:
                partKind === "Office Script artifact/reference"
                  ? "Package artifact only. Import its separate .ts/.osts source for migration."
                  : "Inventory and opaque preservation only; no binary execution.",
            }));
      };
    for (let [archivePath, archiveBytes] of archiveParts)
      (markPart(archivePath, ""),
        /\.(?:xml|rels)$/i.test(archivePath) &&
          scanAutomationXmlMetadata(
            decodeUtf8Strict(archiveBytes),
            archivePath,
            (elementName, attributes) => {
              if (
                (elementName === "Override" &&
                  attributes.PartName &&
                  markPart(
                    attributes.PartName.replace(/^\//, ""),
                    attributes.ContentType ?? "",
                  ),
                elementName === "Default" && attributes.Extension)
              )
                for (let entryPath of archiveParts.keys())
                  entryPath
                    .toLowerCase()
                    .endsWith("." + attributes.Extension.toLowerCase()) &&
                    markPart(entryPath, attributes.ContentType ?? "");
              if (
                elementName === "Relationship" &&
                /vbaProject|(?:intl)?macrosheet|activeX|oleObject|office.?script|scriptLink|scriptReference|relationships\/(?:package|control)$/i.test(
                  attributes.Type ?? "",
                )
              ) {
                if (attributes.TargetMode === "External") {
                  inventory.items.push({
                    path: `${archivePath} \u2192 ${attributes.Target ?? "(missing)"}`,
                    kind: "External automation reference",
                    bytes: 0,
                    execution: "external-unresolved",
                    note: "Referenced content was not fetched, authenticated or executed.",
                  });
                  return;
                }
                let baseDirectory =
                    archivePath === "_rels/.rels"
                      ? ""
                      : archivePath.slice(
                          0,
                          archivePath.lastIndexOf("/_rels/"),
                        ) + "/",
                  resolvedSegments = [];
                for (let segment of ((attributes.Target ?? "").startsWith("/")
                  ? attributes.Target.slice(1)
                  : baseDirectory + (attributes.Target ?? "")
                ).split("/"))
                  segment === ".."
                    ? resolvedSegments.pop()
                    : segment &&
                      segment !== "." &&
                      resolvedSegments.push(segment);
                markPart(resolvedSegments.join("/"), attributes.Type ?? "");
              }
            },
          ));
    for (let [path, flaggedPart] of flaggedParts) {
      let partBytes = archiveParts.get(path);
      inventory.items.push({
        path: path,
        kind: flaggedPart.kind,
        bytes: partBytes?.length ?? 0,
        ...(partBytes ? { sha256: sha256HexDigest(partBytes) } : {}),
        execution: "unsupported",
        note: partBytes
          ? flaggedPart.note
          : "Referenced part is missing. Inventory is incomplete; repair the original.",
      });
    }
    inventory.items.length ||
      inventory.warnings.push(
        "No VBA/XLM or known automation parts were detected in this package. This does not prove the absence of separately stored scripts or automation.",
      );
  } else
    throw Error(
      "Choose VBA .bas/.vba/.cls, Office Script .ts/.osts, or an Excel container for inventory.",
    );
  return (
    inventory.warnings.push(
      "Office Scripts may be stored separately in OneDrive/SharePoint. Workbook inventory cannot discover those scripts; import the authorized source separately. No linked content is fetched.",
    ),
    inventory.retainedEligible ||
      inventory.warnings.push(
        "This file exceeds the 2 MB retained-artifact limit. The inventory is available; preserve the original externally.",
      ),
    r8.set(inventory, {
      signature: JSON.stringify(inventory),
      bytes: fileBytes,
    }),
    inventory
  );
}

function toTypedCellValue(rawValue, dateSystem) {
  if (rawValue != null) {
    if (rawValue instanceof Date)
      return { type: "number", value: dateToExcelSerial(rawValue, dateSystem) };
    if (
      typeof rawValue == "object" &&
      "error" in rawValue &&
      typeof rawValue.error == "string"
    )
      return { type: "error", value: rawValue.error };
    if (typeof rawValue == "string") return { type: "text", value: rawValue };
    if (typeof rawValue == "boolean")
      return { type: "boolean", value: rawValue };
    if (typeof rawValue == "number" && Number.isFinite(rawValue))
      return { type: "number", value: rawValue };
  }
}
function typedValuesMatch(cachedValue, calculated) {
  return (
    cachedValue.type === calculated.type &&
    (typeof cachedValue.value == "number" && typeof calculated.value == "number"
      ? Math.abs(cachedValue.value - calculated.value) <=
        1e-10 *
          Math.max(1, Math.abs(cachedValue.value), Math.abs(calculated.value))
      : cachedValue.value === calculated.value)
  );
}
function compareFormulaCachesWithRecalculation(
  exceljsWorkbook,
  validWorkbook,
  exchangeSchemaReport,
  emptyTextCaches = new Map(),
  formulaCacheReport,
  numericCaches = new Map(),
) {
  let workbookWithSettings = calculateWorkbookWithSettings(validWorkbook);
  for (let sheet of validWorkbook.sheets) {
    let worksheet = exceljsWorkbook.getWorksheet(sheet.name);
    for (let [address, cell] of Object.entries(sheet.cells)) {
      let value = worksheet.getCell(address).value;
      if (
        !value ||
        typeof value != "object" ||
        !("formula" in value || "sharedFormula" in value)
      )
        continue;
      let cachedValue = emptyTextCaches.get(sheet.name)?.has(address)
          ? { type: "text", value: "" }
          : numericCaches.get(sheet.name)?.has(address)
            ? {
                type: "number",
                value: Number(numericCaches.get(sheet.name).get(address)),
              }
            : toTypedCellValue(
                worksheet.getCell(address).result,
                validWorkbook.datePolicy?.dateSystem ?? "1900",
              ),
        calculated = workbookWithSettings.getTypedValue(sheet.id, address);
      formulaCacheReport &&
        formulaCacheReport.entries.push({
          sheetId: sheet.id,
          address: address,
          formula: String(cell.raw),
          ...(cachedValue ? { cache: cachedValue } : {}),
        });
      let cacheStatus = cell.literal
        ? "neutralized"
        : calculated.type === "error" && calculated.value === "#PENDING!"
          ? "pending"
          : cachedValue
            ? typedValuesMatch(cachedValue, calculated)
              ? "matched"
              : "different"
            : "missing";
      (exchangeSchemaReport.caches.total++,
        (exchangeSchemaReport.caches[cacheStatus] =
          (exchangeSchemaReport.caches[cacheStatus] ?? 0) + 1));
      let comparisonList =
        cacheStatus === "matched"
          ? exchangeSchemaReport.caches.matches
          : exchangeSchemaReport.caches.issues;
      comparisonList.length < 100
        ? comparisonList.push({
            sheet: sheet.name,
            address: address,
            ...(cachedValue ? { source: cachedValue } : {}),
            calculated: calculated,
            status: cacheStatus,
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
function assertNoLegacyArrayFormulas(sheetXml, sheetName) {
  let readAttribute = (attributeText, attributeName) =>
      new RegExp(`(?:^|\\s)${attributeName}\\s*=\\s*(["'])(.*?)\\1`, "i")
        .exec(attributeText)?.[2]
        ?.replace(/&#(x[0-9a-f]+|[0-9]+);/gi, (entityMatch, entityCode) =>
          String.fromCodePoint(
            entityCode[0].toLowerCase() === "x"
              ? parseInt(entityCode.slice(1), 16)
              : Number(entityCode),
          ),
        ),
    cellAttributes;
  for (let xmlMatch of sheetXml.matchAll(
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<(\/?)(?:\w+:)?(c|f)\b([^>]*)>/gi,
  )) {
    if (!xmlMatch[2]) continue;
    if (xmlMatch[2].toLowerCase() === "c") {
      cellAttributes =
        xmlMatch[1] || /\/\s*$/.test(xmlMatch[3]) ? undefined : xmlMatch[3];
      continue;
    }
    if (
      cellAttributes === undefined ||
      xmlMatch[1] ||
      readAttribute(xmlMatch[3], "t") !== "array"
    )
      continue;
    let cellAddress = readAttribute(cellAttributes, "r") ?? "unknown address",
      rangeRef = readAttribute(xmlMatch[3], "ref") ?? cellAddress,
      arrayFormulaKind =
        readAttribute(cellAttributes, "cm") !== undefined
          ? "Metadata-qualified Excel array formula (dynamic/legacy classification requires Excel metadata)"
          : "Legacy CSE array formula";
    throw new Error(
      `${arrayFormulaKind} at ${sheetName}!${cellAddress}, range ${rangeRef}. Import stopped before changing the workbook. Open the original in Excel and convert the array to ordinary supported formulas or an explicitly reviewed values snapshot, then import that copy. Fixed CSE ranges and Excel array metadata are not silently treated as dynamic spills.`,
    );
  }
}
var exceljsEsModule = toEsModule(getExceljsModule(), 1);
var papaparseEsModule = toEsModule(getPapaparseModule(), 1);

function decodeDelimitedBytes(bytes, encoding) {
  if (bytes.length > workbookLimits.maxFileBytes)
    throw Error("Delimited files must be 10 MB or smaller.");
  let bomEncoding =
    bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191
      ? "utf-8"
      : bytes[0] === 255 && bytes[1] === 254
        ? "utf-16le"
        : bytes[0] === 254 && bytes[1] === 255
          ? "utf-16be"
          : undefined;
  if (bomEncoding && encoding !== "auto" && encoding !== bomEncoding)
    throw Error(
      `File BOM indicates ${bomEncoding}; select that encoding or Automatic.`,
    );
  let effectiveEncoding =
      encoding === "auto" ? (bomEncoding ?? "utf-8") : encoding,
    decodedText;
  try {
    decodedText = (
      effectiveEncoding === "windows-1252"
        ? Array.from(bytes, (byteValue) => byteCharacterTable[byteValue]).join(
            "",
          )
        : new TextDecoder(effectiveEncoding, {
            fatal: true,
            ignoreBOM: true,
          }).decode(bytes)
    ).replace(/^\uFEFF/, "");
  } catch {
    throw Error(
      `Invalid ${effectiveEncoding} byte sequence. Choose the source encoding; bytes are never silently replaced.`,
    );
  }
  if (decodedText.includes("\0"))
    throw Error(
      "Delimited files cannot contain null bytes; check the encoding.",
    );
  return { text: decodedText, encoding: effectiveEncoding };
}
function tokenizeDelimitedText(
  text,
  delimitedOptions,
  delimitedImportReport,
  onRecord,
) {
  let fields = [],
    fieldText = "",
    fieldState = "start",
    recordNumber = 1,
    recordStart = 0,
    isStopped = false,
    finishField = () => {
      (fields.push(fieldText),
        (fieldText = ""),
        (fieldState = "start"),
        fields.length > workbookLimits.maxCols &&
          (recordDelimitedIssue(
            delimitedImportReport,
            recordNumber,
            undefined,
            `Delimited data exceeds ${workbookLimits.maxCols} columns.`,
          ),
          (isStopped = true)));
    },
    finishRecord = (endIndex) => {
      if ((finishField(), !isStopped)) {
        if (recordNumber > workbookLimits.maxRows) {
          (recordDelimitedIssue(
            delimitedImportReport,
            recordNumber,
            undefined,
            `Delimited data exceeds ${workbookLimits.maxRows} rows.`,
          ),
            (isStopped = true));
          return;
        }
        ((isStopped = !onRecord(
          fields,
          recordNumber,
          endIndex === recordStart,
        )),
          (fields = []),
          recordNumber++);
      }
    };
  for (let charIndex = 0; charIndex < text.length && !isStopped; charIndex++) {
    let character = text[charIndex];
    if (fieldState === "quoted")
      delimitedOptions.escape === "backslash" &&
      character === "\\" &&
      (text[charIndex + 1] === delimitedOptions.quote ||
        text[charIndex + 1] === "\\")
        ? (fieldText += text[++charIndex])
        : character === delimitedOptions.quote
          ? delimitedOptions.escape === "double" &&
            text[charIndex + 1] === delimitedOptions.quote
            ? ((fieldText += character), charIndex++)
            : (fieldState = "closed")
          : (fieldText += character);
    else if (character === delimitedOptions.delimiter) finishField();
    else if (
      character === "\r" ||
      character ===
        `
`
    ) {
      let newlineSequence =
        character === "\r" &&
        text[charIndex + 1] ===
          `
`
          ? `\r
`
          : character;
      (delimitedOptions.newline !== "auto" &&
        delimitedOptions.newline !== newlineSequence &&
        recordDelimitedIssue(
          delimitedImportReport,
          recordNumber,
          undefined,
          "Record newline does not match the selected dialect.",
        ),
        finishRecord(charIndex),
        newlineSequence.length === 2 && charIndex++,
        (recordStart = charIndex + 1));
    } else
      fieldState === "start" &&
      delimitedOptions.quote !== "none" &&
      character === delimitedOptions.quote
        ? (fieldState = "quoted")
        : (fieldState === "closed"
            ? (recordDelimitedIssue(
                delimitedImportReport,
                recordNumber,
                fields.length + 1,
                "Unexpected text after a closing quote.",
              ),
              (fieldState = "plain"))
            : delimitedOptions.quote !== "none" &&
              character === delimitedOptions.quote &&
              recordDelimitedIssue(
                delimitedImportReport,
                recordNumber,
                fields.length + 1,
                "Quote inside an unquoted field; quote the entire field.",
              ),
          (fieldText += character),
          (fieldState = "plain"));
    fieldText.length > fileImportLimits.maxCellTextLength &&
      (recordDelimitedIssue(
        delimitedImportReport,
        recordNumber,
        fields.length + 1,
        "A field exceeds the 32,767-character cell limit.",
      ),
      (isStopped = true));
  }
  (!isStopped &&
    fieldState === "quoted" &&
    recordDelimitedIssue(
      delimitedImportReport,
      recordNumber,
      fields.length + 1,
      "Quoted field is not closed.",
    ),
    !isStopped && recordStart < text.length && finishRecord(text.length),
    (delimitedImportReport.complete = !isStopped));
}

function parseDelimitedRecords(
  bytes,
  fileName,
  rawOptions,
  shouldCollectCells,
) {
  let delimitedOptions = normalizeDelimitedOptions(rawOptions),
    decoded = decodeDelimitedBytes(bytes, delimitedOptions.encoding);
  ((delimitedOptions.encoding = decoded.encoding),
    delimitedOptions.delimiter === "auto" &&
      (delimitedOptions.delimiter = /\.tsv$/i.test(fileName)
        ? "	"
        : papaparseEsModule.default.parse(decoded.text.slice(0, 65536), {
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
    (rawOptions.encoding === undefined || rawOptions.encoding === "auto") &&
      delimitedImportReport.warnings.push(
        `Encoding: ${delimitedOptions.encoding}; automatic mode uses a BOM when present and UTF-8 otherwise. Legacy encodings require an explicit selection.`,
      ),
    (rawOptions.delimiter === undefined || rawOptions.delimiter === "auto") &&
      delimitedImportReport.warnings.push(
        `Detected delimiter: ${delimitedOptions.delimiter === "	" ? "tab" : delimitedOptions.delimiter}. Review the schema and column count before importing.`,
      ));
  let expectedColumnCount = 0;
  if (
    (tokenizeDelimitedText(
      decoded.text,
      delimitedOptions,
      delimitedImportReport,
      (fields, record, isEmptyLine) => {
        ((delimitedImportReport.records = record),
          fields.every((fieldValue) => fieldValue === "") &&
            delimitedImportReport.blankRecords++,
          !expectedColumnCount &&
            !isEmptyLine &&
            (expectedColumnCount = fields.length),
          (delimitedImportReport.columns = Math.max(
            delimitedImportReport.columns,
            fields.length,
          )),
          !isEmptyLine &&
            expectedColumnCount &&
            fields.length !== expectedColumnCount &&
            recordDelimitedIssue(
              delimitedImportReport,
              record,
              undefined,
              `Malformed row: expected ${expectedColumnCount} fields, found ${fields.length}.`,
            ));
        for (let fieldIndex = 0; fieldIndex < fields.length; fieldIndex++) {
          let raw = fields[fieldIndex],
            columnType =
              delimitedOptions.hasHeader && record === 1
                ? "text"
                : (delimitedOptions.columns[fieldIndex] ?? "auto"),
            cellAddress = `${columnIndexToLetters(fieldIndex)}${record}`;
          delimitedImportReport.schema[fieldIndex] ||= {
            index: fieldIndex,
            label:
              delimitedOptions.hasHeader && record === 1
                ? raw
                : columnIndexToLetters(fieldIndex),
            type: delimitedOptions.columns[fieldIndex] ?? "auto",
            observed: [],
          };
          let cell;
          try {
            cell = parseDelimitedCellValue(raw, columnType, delimitedOptions);
          } catch (error) {
            (recordDelimitedIssue(
              delimitedImportReport,
              record,
              fieldIndex + 1,
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
            let observedType =
              cell.raw === null
                ? "blank"
                : cell.format === "date"
                  ? "date"
                  : typeof cell.raw;
            delimitedImportReport.schema[fieldIndex].observed.includes(
              observedType,
            ) ||
              delimitedImportReport.schema[fieldIndex].observed.push(
                observedType,
              );
          }
          if (
            cell.raw !== null &&
            !(isEmptyLine && delimitedOptions.empty === "blank")
          ) {
            if (++delimitedImportReport.storedCells > workbookLimits.maxCells)
              return (
                recordDelimitedIssue(
                  delimitedImportReport,
                  record,
                  fieldIndex + 1,
                  "Delimited data exceeds 100,000 populated cells.",
                ),
                false
              );
            shouldCollectCells && (cells[cellAddress] = cell);
          }
          (raw !== "" &&
            (cell.raw !== raw ||
              (delimitedOptions.trim && raw !== raw.trim())) &&
            recordDelimitedChange(
              delimitedImportReport,
              record,
              fieldIndex + 1,
              cellAddress,
              cell.format === "date"
                ? "Date converted to native serial"
                : "Selected type or whitespace conversion",
              raw,
              String(cell.raw ?? ""),
            ),
            delimitedImportReport.preview.length < 25 &&
              fieldIndex === 0 &&
              delimitedImportReport.preview.push({
                record: record,
                raw: fields.slice(0, 200),
                values: [],
              }),
            record <= 25 &&
              fieldIndex < 200 &&
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
    let headerLabels = delimitedImportReport.schema.map(
        (schemaEntry) => schemaEntry.label,
      ),
      seenLabels = new Set();
    headerLabels.some(
      (label) =>
        !label.trim() ||
        (seenLabels.has(label.toLowerCase())
          ? true
          : (seenLabels.add(label.toLowerCase()), false)),
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
function validateDelimitedImport(bytes, fileName, options = {}) {
  return parseDelimitedRecords(bytes, fileName, options, false).r;
}
function importDelimitedWorkbook(bytes, fileName, options = {}) {
  let { r: delimitedReport, cells } = parseDelimitedRecords(
    bytes,
    fileName,
    options,
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
var createImportedWorkbook = (fileName, sheets) => ({
  schemaVersion: 1,
  id: createRandomUuid(),
  name: fileName.replace(/\.[^.]+$/, "").slice(0, 200) || "Imported workbook",
  revision: 0,
  metadata: createWorkbookMetadata(),
  sheets: sheets,
  events: [],
});

function parseXlsxWorkbookPackage(bytes, parsedPackage) {
  let packageEntries = inspectXlsxPackage(bytes),
    entryNames = new Set(packageEntries.keys()),
    size = packageEntries.size,
    workbookParseState = createWorkbookParseState();
  if (
    [...entryNames].some((entryName) =>
      /(?:vbaProject\.bin|xl\/macrosheets\/|xl\/activeX\/|xl\/embeddings\/)/i.test(
        entryName,
      ),
    )
  )
    throw new Error(
      "Workbooks containing macros, ActiveX or embedded objects are not accepted. Save a clean XLSX copy first.",
    );
  let expandedBytes = 0,
    archiveError,
    entriesSeen = 0,
    entriesParsed = 0,
    mergedCellCount = 0,
    validationCellCount = 0,
    mergeRangeCount = 0,
    hiddenRowCount = 0,
    worksheetCellCount = 0,
    hasOutOfRangeFontSize = false,
    seenEntryNames = new Set(),
    unzipStream = new gc((entry) => {
      if (
        (entriesSeen++,
        !entryNames.has(entry.name) || seenEntryNames.has(entry.name))
      ) {
        archiveError = new Error("Archive entries do not match its directory.");
        return;
      }
      seenEntryNames.add(entry.name);
      let entryExpandedBytes = 0,
        isXmlPart = /\.(?:xml|rels|vml)$/i.test(entry.name),
        xmlDecoder = isXmlPart
          ? new TextDecoder("utf-8", { fatal: true })
          : undefined,
        runningCrc = -1,
        xmlText = "";
      ((entry.ondata = (decompressError, chunk, isFinalChunk) => {
        if (decompressError) {
          archiveError = new Error(
            `Corrupt XLSX compressed content in ${entry.name}. Repair the original workbook and save a clean XLSX copy.`,
          );
          return;
        }
        if (
          ((expandedBytes += chunk.length),
          (entryExpandedBytes += chunk.length),
          expandedBytes > maxXlsxExpandedByteLimit)
        ) {
          (entry.terminate(),
            (archiveError = new Error(
              "The XLSX archive expands beyond 50 MB.",
            )));
          return;
        }
        if (entryExpandedBytes > packageEntries.get(entry.name).size) {
          (entry.terminate(),
            (archiveError = new Error(
              "The XLSX archive expansion does not match its directory.",
            )));
          return;
        }
        runningCrc = computeCrc32(runningCrc, chunk);
        try {
          xmlDecoder &&
            (xmlText += xmlDecoder.decode(chunk, { stream: !isFinalChunk }));
        } catch {
          archiveError = new Error(
            `Unsupported XML encoding in ${entry.name}. Save a standard UTF-8 XLSX copy.`,
          );
          return;
        }
        if (isFinalChunk) {
          if (entryExpandedBytes !== packageEntries.get(entry.name).size) {
            archiveError = new Error(
              "The XLSX archive expansion does not match its directory.",
            );
            return;
          }
          if ((runningCrc ^ -1) >>> 0 !== packageEntries.get(entry.name).crc) {
            archiveError = new Error(
              `XLSX checksum mismatch in ${entry.name}. The archive is corrupt; repair the original and save a clean XLSX copy.`,
            );
            return;
          }
          try {
            if (
              (isXmlPart &&
                parseWorksheetXml(
                  xmlText,
                  entry.name,
                  packageEntries,
                  workbookParseState,
                ),
              /^xl\/styles\.xml$/i.test(entry.name))
            )
              for (let fontSizeMatch of xmlText.matchAll(
                /<(?:\w+:)?sz\b([^>]*)>/g,
              )) {
                let fontSize = Number(getXmlAttribute(fontSizeMatch[1], "val"));
                (!Number.isInteger(fontSize) ||
                  fontSize < 8 ||
                  fontSize > 72) &&
                  (hasOutOfRangeFontSize = true);
              }
            if (/^xl\/worksheets\//i.test(entry.name)) {
              (assertNoLegacyArrayFormulas(xmlText, entry.name),
                (worksheetCellCount = parseWorksheetCells(
                  xmlText,
                  entry.name,
                  worksheetCellCount,
                )),
                /^xl\/worksheets\/[^/]+\.xml$/i.test(entry.name) &&
                  parsedPackage.numericTokens.set(
                    entry.name,
                    scanXlsxNumericTokens(xmlText),
                  ));
              for (let pageSetupMatch of xmlText.matchAll(
                /<(?:\w+:)?pageSetup\b([^>]*)>/gi,
              )) {
                let xmlAttribute = getXmlAttribute(
                  pageSetupMatch[1],
                  "pageOrder",
                );
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
              let pageBreaksByAxis = {};
              for (let [breakTagName, breakKey, axisLimit, crossAxisLimit] of [
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
                let breakSectionMatches = [
                  ...xmlText.matchAll(
                    new RegExp(
                      "<(?:\\w+:)?" +
                        breakTagName +
                        "\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?" +
                        breakTagName +
                        ">",
                      "gi",
                    ),
                  ),
                ];
                if (breakSectionMatches.length > 1)
                  throw Error(
                    "Duplicate XLSX page-break sections are not supported.",
                  );
                let breakBoundaries = [],
                  breakCount = 0;
                for (let breakMatch of (
                  breakSectionMatches[0]?.[1] ?? ""
                ).matchAll(/<(?:\w+:)?brk\b([^>]*)\/?\s*>/gi)) {
                  if (++breakCount > 100)
                    throw Error(
                      "XLSX page breaks exceed 100 boundaries per axis.",
                    );
                  let xmlAttribute = getXmlAttribute(breakMatch[1], "id"),
                    manualFlag = getXmlAttribute(breakMatch[1], "man");
                  if (
                    !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                    !Number.isSafeInteger(Number(xmlAttribute)) ||
                    Number(xmlAttribute) >= axisLimit
                  )
                    throw Error("Invalid XLSX page-break boundary.");
                  if (manualFlag !== "1" && manualFlag !== "true") continue;
                  let spanStart = Number(
                      getXmlAttribute(breakMatch[1], "min") ?? 0,
                    ),
                    spanEnd = Number(
                      getXmlAttribute(breakMatch[1], "max") ??
                        crossAxisLimit - 1,
                    );
                  if (
                    spanStart !== 0 ||
                    !Number.isSafeInteger(spanEnd) ||
                    spanEnd < crossAxisLimit - 1
                  )
                    throw Error(
                      "Partial-span XLSX page breaks are not supported. Use full-row or full-column breaks.",
                    );
                  breakBoundaries.push(Number(xmlAttribute));
                }
                breakBoundaries.sort(
                  (leftBoundary, rightBoundary) => leftBoundary - rightBoundary,
                );
                let validatedBreaks = validatePrintPageBreaks(
                  breakBoundaries,
                  axisLimit,
                );
                validatedBreaks &&
                  (pageBreaksByAxis[breakKey] = validatedBreaks);
              }
              (parsedPackage.printBreaks.set(entry.name, pageBreaksByAxis),
                parsedPackage.conditionalStops.set(
                  entry.name,
                  validateConditionalFormatRules(xmlText),
                ),
                parsedPackage.validationAreas.set(
                  entry.name,
                  parseXlsxDataValidations(xmlText),
                ));
              let hiddenRowIntervals = [],
                columnRanges = [];
              for (let columnMatch of xmlText.matchAll(
                /<(?:\w+:)?col\b([^>]*)>/gi,
              )) {
                let xmlAttribute = getXmlAttribute(columnMatch[1], "min"),
                  maxAttribute = getXmlAttribute(columnMatch[1], "max"),
                  widthAttribute = getXmlAttribute(columnMatch[1], "width"),
                  start = Number(xmlAttribute) - 1,
                  end = Number(maxAttribute) - 1;
                if (
                  !/^[1-9]\d*$/.test(xmlAttribute ?? "") ||
                  !/^[1-9]\d*$/.test(maxAttribute ?? "") ||
                  !Number.isSafeInteger(start) ||
                  !Number.isSafeInteger(end) ||
                  start < 0 ||
                  end < start ||
                  end >= workbookLimits.maxCols ||
                  columnRanges.length >= workbookLimits.maxCols
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
                let width =
                  widthAttribute === undefined
                    ? undefined
                    : Number(widthAttribute);
                if (
                  width !== undefined &&
                  (!widthAttribute?.trim() ||
                    !Number.isFinite(width) ||
                    width < 0)
                )
                  throw Error(
                    "XLSX column widths must be finite nonnegative values.",
                  );
                columnRanges.push({
                  start: start,
                  end: end,
                  ...(width === undefined ? {} : { width: width }),
                });
              }
              columnRanges.sort(
                (leftRange, rightRange) => leftRange.start - rightRange.start,
              );
              for (
                let rangeIndex = 1;
                rangeIndex < columnRanges.length;
                rangeIndex++
              )
                if (
                  columnRanges[rangeIndex].start <=
                  columnRanges[rangeIndex - 1].end
                )
                  throw Error(
                    "XLSX column metadata must stay within 16,384 nonoverlapping columns.",
                  );
              if (
                (parsedPackage.columns.set(entry.name, columnRanges),
                /<(?:\w+:)?sheetFormatPr\b[^>]*\bzeroHeight\s*=\s*["'](?:1|true)["']/i.test(
                  xmlText,
                ))
              )
                throw new Error(
                  "Default-hidden worksheet rows cannot be imported safely. Replace the default with explicit hidden rows before importing.",
                );
              for (let rowMatch of xmlText.matchAll(
                /<(?:\w+:)?row\b([^>]*)>/gi,
              )) {
                if (!/\bhidden\s*=\s*["'](?:1|true)["']/i.test(rowMatch[1]))
                  continue;
                let hiddenRowIndex =
                  Number(getXmlAttribute(rowMatch[1], "r")) - 1;
                if (
                  !Number.isInteger(hiddenRowIndex) ||
                  hiddenRowIndex < 0 ||
                  hiddenRowIndex >= workbookLimits.maxRows
                )
                  throw new Error(
                    "Hidden XLSX rows must have valid row numbers.",
                  );
                let previousInterval = hiddenRowIntervals.at(-1);
                if (previousInterval && hiddenRowIndex <= previousInterval.end)
                  throw new Error(
                    "Hidden XLSX row numbers must be distinct and ascending.",
                  );
                if (
                  (previousInterval &&
                  hiddenRowIndex === previousInterval.end + 1
                    ? (previousInterval.end = hiddenRowIndex)
                    : hiddenRowIntervals.push({
                        start: hiddenRowIndex,
                        end: hiddenRowIndex,
                      }),
                  hiddenRowIntervals.length >
                    worksheetCapacityLimits.maxIntervalsPerAxis)
                )
                  throw new Error(
                    "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
                  );
                if (
                  ++hiddenRowCount > worksheetCapacityLimits.maxXlsxHiddenRows
                )
                  throw new Error(
                    "XLSX hidden rows exceed the 100,000-row processing budget. Reduce hidden rows in the source; native JSON preserves larger visibility intervals.",
                  );
              }
              hiddenRowIntervals.length &&
                parsedPackage.hiddenRows.set(entry.name, hiddenRowIntervals);
            }
            if (entry.name === "xl/workbook.xml")
              for (let sheetMatch of xmlText.matchAll(
                /<(?:\w+:)?sheet\b([^>]*)>/g,
              )) {
                let xmlAttribute = getXmlAttribute(sheetMatch[1], "name"),
                  sheetRelationId = getXmlAttribute(sheetMatch[1], "r:id");
                xmlAttribute &&
                  sheetRelationId &&
                  parsedPackage.sheetNames.set(
                    decodeXmlEntities(xmlAttribute),
                    decodeXmlEntities(sheetRelationId),
                  );
              }
            if (entry.name === "xl/_rels/workbook.xml.rels")
              for (let relationshipMatch of xmlText.matchAll(
                /<(?:\w+:)?Relationship\b([^>]*)>/g,
              )) {
                let xmlAttribute = getXmlAttribute(relationshipMatch[1], "Id"),
                  relationshipTarget = getXmlAttribute(
                    relationshipMatch[1],
                    "Target",
                  );
                xmlAttribute &&
                  relationshipTarget &&
                  getXmlAttribute(relationshipMatch[1], "TargetMode") !==
                    "External" &&
                  parsedPackage.sheetRelationships.set(
                    decodeXmlEntities(xmlAttribute),
                    resolvePackagePartPath(
                      decodeXmlEntities(relationshipTarget),
                    ),
                  );
              }
            for (let rangeMatch of xmlText.matchAll(
              /<(?:\w+:)?(mergeCell|dataValidation)\s[^>]*?\b(ref|sqref)\s*=\s*["']([^"']+)["'][^>]*>/g,
            ))
              for (let rangeRef of decodeXmlEntities(rangeMatch[3])
                .trim()
                .split(/\s+/)) {
                let rangeAddress = parseRangeAddress(rangeRef),
                  rangeCellCount =
                    (rangeAddress.lastRow - rangeAddress.firstRow + 1) *
                    (rangeAddress.lastCol - rangeAddress.firstCol + 1);
                if (rangeMatch[1] === "mergeCell") {
                  if (
                    rangeCellCount > 500 ||
                    ++mergeRangeCount > 1e3 ||
                    (mergedCellCount += rangeCellCount) >
                      workbookLimits.maxCells
                  )
                    throw new Error(
                      "Merged ranges exceed the supported import budget (500 cells per merge and 100,000 cells overall).",
                    );
                } else if (
                  (validationCellCount += rangeCellCount) >
                  workbookLimits.maxCells
                )
                  throw new Error(
                    "XLSX validation ranges exceed the 100,000-cell import budget. Narrow the validation area in the source workbook.",
                  );
              }
            if (entry.name === "xl/workbook.xml") {
              parsedPackage.calculation = parseCalculationProperties(xmlText);
              let definedNamesResult = aP(xmlText);
              ((parsedPackage.names = definedNamesResult.names),
                (parsedPackage.nameWarnings = definedNamesResult.warnings),
                (parsedPackage.builtinNameXml = definedNamesResult.builtinXml));
            }
          } catch (error) {
            archiveError =
              error instanceof Error
                ? error
                : new Error("Invalid XLSX worksheet metadata.");
            return;
          }
          entriesParsed++;
        }
      }),
        entry.start());
    });
  unzipStream.register(mc);
  for (let byteOffset = 0; byteOffset < bytes.length; byteOffset += 2048) {
    if (archiveError) throw archiveError;
    try {
      unzipStream.push(
        bytes.subarray(byteOffset, Math.min(byteOffset + 2048, bytes.length)),
        byteOffset + 2048 >= bytes.length,
      );
    } catch {
      throw new Error(
        "Corrupt XLSX compressed content. Repair the original workbook and save a clean XLSX copy.",
      );
    }
  }
  if (archiveError) throw archiveError;
  if (entriesSeen !== size || entriesParsed !== size)
    throw new Error("The XLSX archive contains incomplete entries.");
  let importWarnings = [
    xlsxExchangeSupportNotice,
    ...parsedPackage.nameWarnings,
    ...validateWorkbookPackage(workbookParseState, entryNames),
  ];
  return (
    (parsedPackage.emptyStringCaches = workbookParseState.emptyStringCaches),
    hasOutOfRangeFontSize &&
      importWarnings.push(
        "Unsupported fractional or out-of-range font sizes were found in the source. The XLSX reader may normalize them; the editor supports integer font sizes from 8 to 72 points.",
      ),
    [...entryNames].some((partName) =>
      /externalLinks|connections\.xml/i.test(partName),
    ) &&
      importWarnings.push(
        "External workbook links or connections were detected. They will not be refreshed or included in the derived copy; the original retains their source parts.",
      ),
    importWarnings
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
function normalizeArgbHexColor(argb) {
  return argb && /^(?:[0-9a-f]{8}|[0-9a-f]{6})$/i.test(argb)
    ? `#${argb.slice(-6).toLowerCase()}`
    : undefined;
}
function convertXlsxCellToRecord(
  xlsxCell,
  xlsxWorkbookPackage = [],
  numericToken,
  isDate1904 = false,
) {
  let addWarning = (warningMessage) => {
      xlsxWorkbookPackage.includes(warningMessage) ||
        xlsxWorkbookPackage.push(warningMessage);
    },
    value = xlsxCell.value;
  if (
    value === null &&
    !xlsxCell.note &&
    !Object.keys(xlsxCell.style ?? {}).length
  )
    return;
  let raw = null,
    literal;
  typeof value == "string"
    ? ((raw = value), raw.startsWith("=") && (literal = true))
    : typeof value == "number" || typeof value == "boolean"
      ? (raw = value)
      : value instanceof Date
        ? (raw =
            numericToken !== undefined
              ? Number(numericToken)
              : dateToExcelSerial(value, isDate1904 ? "1904" : "1900"))
        : value &&
          typeof value == "object" &&
          ("formula" in value || "sharedFormula" in value
            ? ((raw = `=${xlsxCell.formula}`),
              isExternalActionFormula(raw) && (literal = true))
            : "richText" in value
              ? (addWarning(
                  "Rich-text runs are flattened to cell text; per-run formatting is not preserved.",
                ),
                (raw = value.richText
                  .map((richTextRun) => richTextRun.text)
                  .join("")),
                raw.startsWith("=") && (literal = true))
              : "hyperlink" in value
                ? ((raw = value.text), raw.startsWith("=") && (literal = true))
                : "error" in value && (raw = value.error));
  let cellStyle = {};
  (xlsxCell.font?.bold && (cellStyle.bold = true),
    xlsxCell.font?.italic && (cellStyle.italic = true),
    xlsxCell.font?.strike && (cellStyle.strikethrough = true),
    xlsxCell.alignment?.vertical &&
      ["top", "middle", "bottom"].includes(xlsxCell.alignment.vertical) &&
      (cellStyle.verticalAlign = xlsxCell.alignment.vertical),
    xlsxCell.alignment?.indent !== undefined &&
      (Number.isInteger(xlsxCell.alignment.indent) &&
      xlsxCell.alignment.indent >= 0 &&
      xlsxCell.alignment.indent <= 15
        ? (cellStyle.indent = xlsxCell.alignment.indent)
        : addWarning(
            "Unsupported indentation was omitted; the editor supports integer levels 0\u201315.",
          )),
    xlsxCell.font?.underline && (cellStyle.underline = true),
    supportedFontFamilies.includes(xlsxCell.font?.name) &&
      (cellStyle.fontFamily = xlsxCell.font.name),
    xlsxCell.font?.size !== undefined &&
      Number.isInteger(xlsxCell.font.size) &&
      xlsxCell.font.size >= 8 &&
      xlsxCell.font.size <= 72 &&
      (cellStyle.fontSize = xlsxCell.font.size),
    xlsxCell.alignment?.wrapText && (cellStyle.wrap = true));
  let borders = {};
  for (let cellBorderEdge of cellBorderEdges) {
    let edgeBorder = xlsxCell.border?.[cellBorderEdge];
    if (!edgeBorder?.style) continue;
    if (!borderLineStyles.includes(edgeBorder.style)) {
      addWarning(
        "An unsupported border line was omitted; supported edges and base values were retained.",
      );
      continue;
    }
    let argbHexColor = normalizeArgbHexColor(edgeBorder.color?.argb);
    (edgeBorder.color &&
      !argbHexColor &&
      addWarning(
        "A theme/indexed border color used a black fallback; theme bindings are not preserved.",
      ),
      (borders[cellBorderEdge] = {
        style: edgeBorder.style,
        color: argbHexColor ?? "#000000",
      }));
  }
  let borderEdges = Object.keys(borders),
    isUniformThinBorder = borderEdges.every(
      (borderEdge) =>
        borders[borderEdge]?.style === "thin" &&
        borders[borderEdge]?.color?.toLowerCase() === "#000000",
    );
  (isUniformThinBorder && borderEdges.length === 4
    ? (cellStyle.border = "all")
    : isUniformThinBorder && borderEdges.length === 1 && borders.bottom
      ? (cellStyle.border = "bottom")
      : borderEdges.length && (cellStyle.borders = borders),
    xlsxCell.border?.diagonal &&
      addWarning("Diagonal borders are unsupported and were omitted."),
    xlsxCell.font?.name &&
      !supportedFontFamilies.includes(xlsxCell.font.name) &&
      addWarning(
        "An unsupported font family uses the editor default; exact font fidelity is not preserved.",
      ),
    xlsxCell.font?.size !== undefined &&
      (!Number.isInteger(xlsxCell.font.size) ||
        xlsxCell.font.size < 8 ||
        xlsxCell.font.size > 72) &&
      addWarning(
        "An unsupported font size was omitted; supported sizes are integers from 8 to 72 points.",
      ),
    xlsxCell.font?.underline &&
      xlsxCell.font.underline !== true &&
      xlsxCell.font.underline !== "single" &&
      addWarning(
        "An unsupported underline variant was simplified to single underline.",
      ),
    (xlsxCell.font?.vertAlign ||
      xlsxCell.alignment?.textRotation ||
      xlsxCell.alignment?.shrinkToFit ||
      (xlsxCell.alignment?.horizontal &&
        !["left", "center", "right"].includes(xlsxCell.alignment.horizontal)) ||
      (xlsxCell.alignment?.vertical &&
        !["top", "middle", "bottom"].includes(xlsxCell.alignment.vertical))) &&
      addWarning(
        "Unsupported text positioning/alignment settings were omitted; raw values were retained.",
      ),
    xlsxCell.font?.color &&
      !normalizeArgbHexColor(xlsxCell.font.color.argb) &&
      addWarning(
        "A theme/indexed font color uses the editor default; theme bindings are not preserved.",
      ),
    xlsxCell.fill &&
      (xlsxCell.fill.type !== "pattern" ||
        !["solid", "none"].includes(xlsxCell.fill.pattern)) &&
      addWarning(
        "An unsupported gradient or pattern fill was omitted; base cell values were retained.",
      ),
    xlsxCell.fill?.type === "pattern" &&
      xlsxCell.fill.pattern === "solid" &&
      xlsxCell.fill.fgColor &&
      !normalizeArgbHexColor(xlsxCell.fill.fgColor.argb) &&
      addWarning(
        "A theme/indexed fill color uses the editor default; theme bindings are not preserved.",
      ));
  let argbHexColor = normalizeArgbHexColor(xlsxCell.font?.color?.argb);
  if (
    (argbHexColor && (cellStyle.color = argbHexColor),
    xlsxCell.fill?.type === "pattern" && xlsxCell.fill.pattern === "solid")
  ) {
    let argbHexColor = normalizeArgbHexColor(xlsxCell.fill.fgColor?.argb);
    argbHexColor && (cellStyle.fill = argbHexColor);
  }
  xlsxCell.alignment &&
    ["left", "center", "right"].includes(xlsxCell.alignment.horizontal ?? "") &&
    (cellStyle.align = xlsxCell.alignment.horizontal);
  let note =
      typeof xlsxCell.note == "string"
        ? xlsxCell.note
        : xlsxCell.note?.texts?.map((text) => text.text).join(""),
    numFmt =
      xlsxCell.numFmt && xlsxCell.numFmt !== "General"
        ? xlsxCell.numFmt
        : undefined;
  return {
    raw: raw,
    ...(literal ? { literal: literal } : {}),
    ...(Object.keys(cellStyle).length ? { style: cellStyle } : {}),
    ...(numFmt ? { numFmt: numFmt, format: classifyNumberFormat(numFmt) } : {}),
    ...(note ? { note: note } : {}),
  };
}
async function importWorkbookFile(file) {
  if (/\.xltx$/i.test(file.name)) return convertXltxTemplateFile(file);
  if (/\.json$/i.test(file.name)) {
    if (file.size > workbookLimits.maxNativeBytes)
      throw new Error("Native workbook exceeds the 50 MB import limit.");
    let nativeWorkbookJson = parseNativeWorkbookJson(await file.text());
    return {
      workbook: nativeWorkbookJson,
      warnings: [
        ...collectStyleFallbackWarnings(nativeWorkbookJson),
        ...collectDataValidationWarnings(nativeWorkbookJson),
      ],
    };
  }
  if (file.size > workbookLimits.maxFileBytes)
    throw new Error("Files must be 10 MB or smaller.");
  if (/\.(xlsm|xlsb|xls|xltm|xlam)$/i.test(file.name))
    throw new Error(
      "Legacy or macro-enabled Excel formats are not supported. Save as XLSX, CSV or native JSON.",
    );
  if (/\.(csv|tsv)$/i.test(file.name))
    return importDelimitedWorkbook(
      new Uint8Array(await file.arrayBuffer()),
      file.name,
      {
        hasHeader: false,
      },
    );
  if (!/\.xlsx$/i.test(file.name))
    throw new Error("Choose an XLSX, CSV, TSV or JSON file.");
  let fileBytes = new Uint8Array(await file.arrayBuffer()),
    xlsxMetadata = {
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
  if (fileBytes.byteLength > workbookLimits.maxFileBytes)
    throw new Error("XLSX files must be 10 MB or smaller.");
  let xlsxWorkbookPackage = parseXlsxWorkbookPackage(fileBytes, xlsxMetadata),
    retainedSource = retainOriginalXlsx(fileBytes, file.name),
    partInventory = F$(Nd(fileBytes), retainedSource.sha256),
    exceljsWorkbook = new exceljsEsModule.default.Workbook(),
    workbookBytes = fileBytes;
  if (xlsxMetadata.names.length || xlsxMetadata.nameWarnings.length) {
    let zipParts = Nd(fileBytes);
    ((zipParts["xl/workbook.xml"] = za(
      insertDefinedNamesIntoWorkbookXml(
        Ao(zipParts["xl/workbook.xml"]),
        xlsxMetadata.builtinNameXml,
      ),
    )),
      (workbookBytes = Rd(zipParts)));
  }
  try {
    await exceljsWorkbook.xlsx.load(workbookBytes);
  } catch {
    throw new Error(
      "The XLSX workbook contains corrupt or unsupported records. Open the original in Excel, repair it if needed, and save a clean XLSX copy.",
    );
  }
  if (
    exceljsWorkbook.worksheets.length !== xlsxMetadata.sheetNames.size ||
    exceljsWorkbook.worksheets.some(
      (worksheet) => !xlsxMetadata.sheetNames.has(worksheet.name),
    )
  )
    throw new Error(
      "Unsupported or missing worksheet records. Save a standard XLSX copy with ordinary worksheets.",
    );
  if (!exceljsWorkbook.worksheets.length)
    throw new Error("A workbook must have at least one sheet.");
  let populatedCellCount = 0,
    adjustedRowHeightCount = 0,
    importedSheets = exceljsWorkbook.worksheets.map((worksheet) => {
      let columnCount = Math.max(
        worksheet.columnCount,
        worksheet.columns?.length ?? 0,
      );
      if (
        worksheet.rowCount > workbookLimits.maxRows ||
        columnCount > workbookLimits.maxCols
      )
        throw new Error(
          `Worksheet ${worksheet.name} exceeds ${workbookLimits.maxRows} rows or ${workbookLimits.maxCols} columns.`,
        );
      let cells = {},
        model = worksheet.model,
        rowHeights = {},
        sheetPath = xlsxMetadata.sheetRelationships.get(
          xlsxMetadata.sheetNames.get(worksheet.name) ?? "",
        ),
        hiddenRows = xlsxMetadata.hiddenRows.get(sheetPath ?? "") ?? [],
        addInterval = (intervals, position) => {
          let lastInterval = intervals.at(-1);
          if (lastInterval && position === lastInterval.end + 1)
            lastInterval.end = position;
          else {
            if (intervals.length >= worksheetCapacityLimits.maxIntervalsPerAxis)
              throw new Error(
                "Hidden metadata exceeds 10,000 intervals per worksheet axis.",
              );
            intervals.push({ start: position, end: position });
          }
        };
      for (let row of model.rows) {
        if (typeof row.height == "number" && Number.isFinite(row.height)) {
          let rowHeightPx = (row.height * 4) / 3;
          ((rowHeights[String(row.number - 1)] = Math.max(
            21,
            Math.min(546, rowHeightPx),
          )),
            (rowHeightPx < 21 || rowHeightPx > 546) &&
              adjustedRowHeightCount++);
        }
        for (let cell of row.cells) {
          let sheetCell = worksheet.getCell(cell.address);
          if (
            sheetCell.isMerged &&
            sheetCell.master.address !== sheetCell.address
          )
            continue;
          let cellRecord = convertXlsxCellToRecord(
            sheetCell,
            xlsxWorkbookPackage,
            xlsxMetadata.numericTokens.get(sheetPath ?? "")?.get(cell.address),
            !!exceljsWorkbook.properties.date1904,
          );
          if (cellRecord) {
            if (++populatedCellCount > workbookLimits.maxCells)
              throw new Error("Workbook exceeds 100,000 populated cells.");
            cells[sheetCell.address] = cellRecord;
          }
        }
      }
      let frozenView = worksheet.views?.find((view) => view.state === "frozen"),
        hiddenCols = [];
      (worksheet.columns ?? []).forEach((column, columnIndex) => {
        column.hidden && addInterval(hiddenCols, columnIndex);
      });
      let columnWidths = {};
      for (let columnEntry of xlsxMetadata.columns.get(sheetPath ?? "") ?? []) {
        if (columnEntry.width === undefined) continue;
        let columnWidthPx = columnEntry.width * 7;
        (columnWidthPx < 40 || columnWidthPx > 600) &&
          xlsxWorkbookPackage.push(
            "Worksheet " +
              worksheet.name +
              " column widths were clamped to the supported pixel bounds.",
          );
        let clampedWidth = Math.max(40, Math.min(600, columnWidthPx));
        for (let start = columnEntry.start; start <= columnEntry.end; start++)
          columnWidths[start] = clampedWidth;
      }
      let resolveDefaultMetric = (
          rawValue,
          pixelScale,
          minPixels,
          maxPixels,
          metricName,
        ) => {
          if (rawValue === undefined || rawValue === 0) return;
          if (
            typeof rawValue != "number" ||
            !Number.isFinite(rawValue) ||
            rawValue <= 0
          )
            throw Error("Invalid worksheet default " + metricName + ".");
          let pixelValue = rawValue * pixelScale;
          return (
            (pixelValue < minPixels || pixelValue > maxPixels) &&
              xlsxWorkbookPackage.push(
                "Worksheet " +
                  worksheet.name +
                  " default " +
                  metricName +
                  " was clamped to the supported pixel bounds.",
              ),
            Math.max(minPixels, Math.min(maxPixels, pixelValue))
          );
        },
        defaultRowHeight = resolveDefaultMetric(
          worksheet.properties.defaultRowHeight,
          4 / 3,
          21,
          546,
          "row height",
        ),
        defaultColumnWidth = resolveDefaultMetric(
          worksheet.properties.defaultColWidth,
          7,
          40,
          600,
          "column width",
        ),
        sheetRecord = {
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
          frozenRows: frozenView?.ySplit ?? 0,
          ...(frozenView?.xSplit ? { frozenCols: frozenView.xSplit } : {}),
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
      (xlsxWorkbookPackage.push(
        ...convertImportedTables(worksheet, sheetRecord),
      ),
        xlsxWorkbookPackage.push(
          ...convertWorksheetFeatures(
            worksheet,
            sheetRecord,
            xlsxMetadata.conditionalStops.get(sheetPath ?? ""),
            xlsxMetadata.validationAreas.get(sheetPath ?? ""),
          ),
        ),
        xlsxWorkbookPackage.push(
          ...collectImportedPictures(exceljsWorkbook, worksheet, sheetRecord, {
            containerFilename: file.name,
          }),
        ));
      try {
        importPageSetup(worksheet, sheetRecord);
        let printBreakSettings = xlsxMetadata.printBreaks.get(sheetPath ?? "");
        printBreakSettings &&
          validatePageSetup(
            { ...sheetRecord.pageLayout, ...printBreakSettings },
            sheetRecord,
          );
      } catch (error) {
        xlsxWorkbookPackage.push(
          `Page setup on ${sheetRecord.name} could not be retained: ${error.message}`,
        );
      }
      return sheetRecord;
    });
  (adjustedRowHeightCount &&
    xlsxWorkbookPackage.push(
      `${adjustedRowHeightCount} explicit row height(s) were adjusted to the supported ${21}\u2013${546} CSS-pixel range.`,
    ),
    exceljsWorkbook.properties.date1904 &&
      xlsxWorkbookPackage.push(
        "The source uses the 1904 date system. Its date system, numeric serials and formulas are preserved; calendar dates are not silently converted to 1900.",
      ),
    xlsxWorkbookPackage.push(
      "Formula results are recalculated by the supported local formula engine. Unsupported Excel functions may display errors. Cell error values are imported as displayed text. Audit history and application tickmarks are available only in native JSON backups.",
    ));
  let importedWorkbook = createImportedWorkbook(file.name, importedSheets);
  importedWorkbook.xlsxSource = retainedSource;
  let externalLinkTargets = parseExternalLinkTargets(
    Nd(fileBytes),
    retainedSource.sha256,
  );
  (externalLinkTargets.length &&
    (importedWorkbook.workbookLinks = {
      schemaVersion: 1,
      sources: [],
      packageTargets: externalLinkTargets,
    }),
    exceljsWorkbook.properties.date1904 &&
      (importedWorkbook.datePolicy = {
        ...defaultDatePolicy,
        dateSystem: "1904",
      }));
  let precisionRiskCount = 0,
    precisionRiskExamples = [];
  for (let [sheetName, sheetRelationshipId] of xlsxMetadata.sheetNames)
    for (let [cellAddress, numericToken] of xlsxMetadata.numericTokens.get(
      xlsxMetadata.sheetRelationships.get(sheetRelationshipId) ?? "",
    ) ?? []) {
      let precisionIssue = validateXlsxNumericValue(numericToken);
      precisionIssue &&
        (precisionRiskCount++,
        precisionRiskExamples.length < 8 &&
          precisionRiskExamples.push(
            `${sheetName}!${cellAddress}: ${numericToken.slice(0, 80)} (${precisionIssue})`,
          ));
    }
  (precisionRiskCount &&
    xlsxWorkbookPackage.push(
      `Numeric precision risk: ${precisionRiskCount} numeric value/cache(s) may lose significant digits in Excel or binary conversion. Numeric cells retain their declared numeric type; store identifiers as Text and keep the source file. ${precisionRiskExamples.join("; ")}`,
    ),
    exceljsWorkbook.created instanceof Date &&
      Number.isFinite(exceljsWorkbook.created.getTime()) &&
      exceljsWorkbook.modified instanceof Date &&
      Number.isFinite(exceljsWorkbook.modified.getTime()) &&
      exceljsWorkbook.modified >= exceljsWorkbook.created &&
      (importedWorkbook.metadata = {
        createdAt: exceljsWorkbook.created.toISOString(),
        modifiedAt: exceljsWorkbook.modified.toISOString(),
      }));
  let definedNameScopes = resolveDefinedNameScopes(
    xlsxMetadata.names,
    importedWorkbook,
    [...xlsxMetadata.sheetNames.keys()],
  );
  definedNameScopes.length &&
    (importedWorkbook.definedNames = definedNameScopes);
  let sourceSettings = xlsxMetadata.calculation ?? {
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
        name: file.name,
        kind: "file",
      }),
    ),
    exchangeSchemaReport = buildExchangeSchemaReport(
      "import",
      file.name,
      fileBytes.byteLength,
      validWorkbook,
      xlsxWorkbookPackage,
    );
  ((exchangeSchemaReport.partInventory = partInventory),
    exchangeSchemaReport.warnings.push(
      "The exact source XLSX is retained in this workbook and native JSON. Passive custom metadata can be copied unchanged; all other unsupported parts remain accessible in the original. Save native JSON to retain that source across sessions.",
    ));
  let calculationCaches = {
    schemaVersion: 1,
    fileName: file.name,
    sourceSha256: Array.from(
      new Uint8Array(
        await crypto.subtle.digest("SHA-256", new Uint8Array(fileBytes).buffer),
      ),
    )
      .map((byte) => byte.toString(16).padStart(2, "0"))
      .join(""),
    importedAt: new Date().toISOString(),
    sourceSettings: sourceSettings,
    entries: [],
  };
  return (
    compareFormulaCachesWithRecalculation(
      exceljsWorkbook,
      validWorkbook,
      exchangeSchemaReport,
      new Map(
        [...xlsxMetadata.sheetNames].map(
          ([cacheSheetName, cacheRelationshipId]) => [
            cacheSheetName,
            xlsxMetadata.emptyStringCaches?.get(
              xlsxMetadata.sheetRelationships.get(cacheRelationshipId) ?? "",
            ) ?? new Set(),
          ],
        ),
      ),
      calculationCaches,
      new Map(
        [...xlsxMetadata.sheetNames].map(
          ([tokenSheetName, tokenRelationshipId]) => [
            tokenSheetName,
            xlsxMetadata.numericTokens.get(
              xlsxMetadata.sheetRelationships.get(tokenRelationshipId) ?? "",
            ) ?? new Map(),
          ],
        ),
      ),
    ),
    (validWorkbook.importedCalculationCaches = calculationCaches),
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
function parseNativeWorkbookJson(jsonText) {
  if (new TextEncoder().encode(jsonText).length > workbookLimits.maxNativeBytes)
    throw new Error("Native workbook exceeds the 50 MB import limit.");
  let workbook;
  try {
    workbook = JSON.parse(jsonText);
  } catch {
    throw new Error("The native workbook is not valid JSON.");
  }
  return toValidWorkbook(workbook);
}

function buildRetainedXlsxSchemaReport(workbook) {
  let retainedSource = requireRetainedXlsxSource(workbook),
    xlsxWorkbookPackage = parseXlsxWorkbookPackage(
      retainedSource.bytes,
      createEmptyExportLedger(),
    ),
    exchangeSchemaReport = buildExchangeSchemaReport(
      "import",
      retainedSource.fileName,
      retainedSource.bytes.length,
      workbook,
      xlsxWorkbookPackage,
    );
  return (
    (exchangeSchemaReport.partInventory = F$(
      Nd(retainedSource.bytes),
      retainedSource.sha256,
    )),
    exchangeSchemaReport
  );
}

function toErrorResponse(error) {
  let { code, message } = toFileProcessingError(error);
  return { code, message };
}

for (let globalName of [
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
  Object.defineProperty(globalThis, globalName, {
    value: undefined,
    writable: false,
    configurable: false,
  });
function parseJsonWithinBudget(bytes) {
  let parsedJson = JSON.parse(
      new TextDecoder("utf-8", { fatal: true }).decode(bytes),
    ),
    nodeStack = [{ value: parsedJson, depth: 0 }],
    nodeCount = 0;
  for (; nodeStack.length;) {
    let entry = nodeStack.pop();
    if (++nodeCount > u1.maxJsonNodes || entry.depth > u1.maxJsonDepth)
      throw new FileProcessingError("FILE_LIMIT");
    if (entry.value && typeof entry.value == "object")
      for (let [key, childValue] of Object.entries(entry.value)) {
        if (["__proto__", "prototype", "constructor"].includes(key))
          throw new FileProcessingError("FILE_ACTIVE");
        nodeStack.push({ value: childValue, depth: entry.depth + 1 });
      }
  }
  return parsedJson;
}
var workerScope = self,
  messageHandled = false;
workerScope.onmessage = async ({ data }) => {
  if (messageHandled) return;
  messageHandled = true;
  let messageId = data?.id;
  try {
    if (
      typeof messageId != "string" ||
      messageId.length > 100 ||
      !(data.bytes instanceof Uint8Array)
    )
      throw new FileProcessingError("FILE_INVALID");
    assertImportFileAllowed(data.operation, data.fileName, data.bytes.length);
    let result;
    switch (data.operation) {
      case "source":
        result = inspectSourceFile(data.bytes, data.fileName);
        break;
      case "formats":
        result = inspectAlternateWorkbookFile(data.bytes, data.fileName);
        break;
      case "workbook": {
        /\.json$/i.test(data.fileName) && parseJsonWithinBudget(data.bytes);
        let bytes = data.bytes,
          importResult = await importWorkbookFile({
            name: data.fileName,
            size: bytes.length,
            arrayBuffer: async () => bytes.buffer,
            text: async () =>
              new TextDecoder("utf-8", { fatal: true }).decode(bytes),
          });
        if (
          (/\.json$/i.test(data.fileName) &&
            importResult.workbook.xlsxSource &&
            buildRetainedXlsxSchemaReport(importResult.workbook),
          /\.json$/i.test(data.fileName) &&
            importResult.workbook.alternateSource)
        ) {
          let alternateSource = importResult.workbook.alternateSource;
          inspectAlternateWorkbookFile(
            validateAlternateSource(alternateSource),
            alternateSource.fileName,
          );
        }
        if (/\.json$/i.test(data.fileName))
          for (let sourceImport of importResult.workbook.sourceImports ?? []) {
            let sourceInspection = inspectSourceFile(
              validateRetainedSourceFile(sourceImport.original),
              sourceImport.original.fileName,
            );
            if (
              sourceInspection.method !== "document-adapter" &&
              JSON.stringify(
                sourceInspection.tables.find(
                  (table) => table.id === sourceImport.table.id,
                ),
              ) !== JSON.stringify(sourceImport.table)
            )
              throw Error(
                "Retained structured evidence does not match its original.",
              );
          }
        result = importResult;
        break;
      }
      case "delimited": {
        let report = validateDelimitedImport(
          data.bytes,
          data.fileName,
          data.options,
        );
        result = {
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
        result = inspectAutomationFile(data.bytes, data.fileName);
        break;
      case "json":
        result = parseJsonWithinBudget(data.bytes);
        break;
      case "picture": {
        let binaryText = "";
        for (let byte of data.bytes) binaryText += String.fromCharCode(byte);
        let dataURL = `data:image/${/\.png$/i.test(data.fileName.trim()) ? "png" : "jpeg"};base64,${btoa(binaryText)}`,
          pictureInfo = inspectPictureDataUrl(dataURL);
        result = {
          dataURL: dataURL,
          width: pictureInfo.width,
          height: pictureInfo.height,
        };
        break;
      }
      default:
        throw new FileProcessingError("FILE_TYPE");
    }
    if (
      new TextEncoder().encode(JSON.stringify(result)).length >
      u1.maxResultBytes
    )
      throw new FileProcessingError("FILE_LIMIT");
    workerScope.postMessage({ id: messageId, ok: true, value: result });
  } catch (error) {
    workerScope.postMessage({
      id: messageId,
      ok: false,
      error: toErrorResponse(error),
    });
  }
};
