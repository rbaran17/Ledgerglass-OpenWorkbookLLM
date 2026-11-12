import {
  isNonArrayObject,
  throwWithMessage,
  failWithMessage,
  throwWorkbookError,
} from "./01-cell-styles-and-themes.js";
import {
  isPlainDataObject,
  isFiniteNumber,
} from "./04-pivot-tables-and-pictures.js";
import {
  workbookLimits,
  fileImportLimits,
  maxRowHeight,
  minRowHeight,
  maxColumnWidth,
  minColumnWidth,
} from "../02-text-and-financial-functions.js";
import {
  isDeepFrozenValue,
  copyExtensionCacheEntries,
  validateSavedRecipeMetadata,
  validateExtensionState,
  normalizeDatePolicy,
  fingerprintCanonicalValue,
  deepFreezeValue,
} from "../01-schemas-date-functions-and-validation.js";
import {
  cloneStyleValue,
  reconcileWorkbookStyleBindings,
  isPlainObjectValue,
  canonicalStyleKey,
  regenerateOutlineGroupIds,
  validateOutlineGroups,
  aQ,
  createReviewThread,
  fQ,
  createFormulaWatch,
  assertSheetUnprotected,
  createWorkbookMetadata,
  validateCellStyle,
  validateWorkbookMetadata,
  renameStyleNamespace,
  trimRequiredMetadataText,
} from "./01-cell-styles-and-themes.js";
import { validateNumberFormatSettings } from "./02-formula-function-evaluators.js";
import {
  assertWorkbookStructureEditable,
  createUniqueWorksheetName,
  $v,
  describeSheetDeletionImpacts,
  visitTableFormulasOutsideSheet,
  isNullOrEmptyValue,
  hashWorkbookStructurePassword,
  coerceImportedValue,
  convertTextForDataType,
  rewriteStructuredTableReference,
  validateDefinedName,
  createUuid,
  validateTableColumnType,
  validateWorkbookPassword,
  validateWorkbookStructureProtection,
  findWorksheetById,
  assertUniqueWorksheetName,
  parseWorksheetQualifiedReference,
  validateWorksheetVisibility,
  validateTabColor,
  isVisibleWorksheet,
} from "./03-defined-names.js";
import {
  tableMutationLimits,
  attachWorkbookLinkMetadata,
  computePivotTable,
  findPivotDefinition,
  assertUniqueTableNames,
  findSourceFieldIndex,
  normalizeRowProvenance,
  filterPivotSourceRows,
  buildUniqueSheetName,
  validateCalculatedColumnFormula,
  assertWithinCellBudget,
} from "./04-pivot-tables-and-pictures.js";
import {
  importFormulaCacheEvidence,
  buildExternalSourceRefreshState,
  verifyConsolidationEvidence,
  zie,
} from "./05-formula-engine-and-name-operations.js";
import {
  validateAnalysisReport,
  validateDataModelReport,
  requireBoundedText,
  readRetainedXlsxSource,
  JU,
  validateDelimitedSourceFootprint,
  validateFieldTransformRecipe,
  assertPlainObjectKeys,
  NC,
  P2,
  QU,
  validateSourceImports,
  WU,
  YE,
  validateTextCleanupSteps,
  validateAutomationArtifacts,
  textInferenceLimits,
  throwTextCleanupEvidenceError,
  jm,
  normalizeAlternateFormatSource,
  columnIndexToLetters,
  hasOwnKey,
  shiftFormulaReferences,
  numberFormatCategories,
} from "./06-data-model-and-analysis.js";
import {
  normalizeReviewMetadata,
  isCanonicalCellAddress,
  validateFreezePanes,
  validateFormulaWatches,
  validateWorksheetFeatures,
  validateHyperlinkRecord,
  _t,
  parseRangeAddress,
  validateSavedWhatIfAnalyses,
  analysisOutputEvidenceSchema,
} from "./07-conditional-rules-worksheet-features.js";
import {
  measureRangeArea,
  normalizeTableDefinitions,
  Cn,
  Dp,
  resolveSourceRange,
  tableDataBodyRange,
  normalizeQueryDefinition,
  rangesIntersect,
  applyTableStyleToCells,
  toA1RangeAddress,
  collectRangeValues,
  validateTableName,
  buildTableTotalCell,
  filterSourceRows,
  normalizePivotDefinition,
  validateDrawingObjects,
} from "./08-table-pivot-query-definitions.js";
import {
  updateConditionalFormatRules,
  calculateWorkbookWithSettings,
  checkDataValidationBatch,
} from "./09-calculation-engine.js";
import {
  assertRangeCellBudget,
  moveTotalsRowCells,
  findSheetOrThrow,
  buildPivotDrillDown,
  buildSavedReportEnvelope,
  validatePageSetup,
  buildPivotSource,
  resizeTableRange,
  validateSavedFilterViews,
  renameTableHeaderField,
  collectHyperlinkEntries,
  aggregatePivotRows,
  applyNamedStyleChange,
  validateReviewedTableHeaders,
} from "./10-table-filters-and-pivot-views.js";

function createWorksheetRecord(sheetName) {
  return !sheetName.trim() ||
    sheetName.length > 31 ||
    /[\\/\[\]*?:]/.test(sheetName)
    ? throwWithMessage(
        "Sheet names must contain 1–31 characters and cannot contain \\ / [ ] * ? : .",
      )
    : {
        id: crypto.randomUUID(),
        name: sheetName.trim(),
        cells: {},
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        frozenRows: 0,
      };
}
function writeAnalysisOutputBlock(
  sourceWorkbook,
  fallbackSheetName,
  matrix,
  destination,
) {
  destination?.sheetId || assertWorkbookStructureEditable(sourceWorkbook);
  const columnCount = matrix[0]?.length ?? 0;
  ((!columnCount || matrix.length * columnCount > workbookLimits.maxCells) &&
    throwWorkbookError("The analysis output exceeds 100,000 cells."),
    matrix.some((row) =>
      row.some(
        (cellValue) => typeof cellValue == "string" && cellValue.length > 32767,
      ),
    ) &&
      throwWorkbookError(
        "An analysis label exceeds the 32,767-character cell limit. Shorten source labels or value-field names.",
      ));
  const workbook = structuredClone(sourceWorkbook),
    sheet = destination?.sheetId
      ? findSheetOrThrow(workbook, destination.sheetId)
      : createWorksheetRecord(
          buildUniqueSheetName(
            workbook,
            destination?.sheetName ?? fallbackSheetName,
          ),
        );
  (destination?.sheetId || workbook.sheets.push(sheet),
    assertSheetUnprotected(sheet));
  const destinationPosition = _t(destination?.address ?? "A1"),
    lastRow = destinationPosition.row + matrix.length - 1,
    lastCol = destinationPosition.col + columnCount - 1;
  (lastRow >= sheet.rowCount || lastCol >= sheet.colCount) &&
    throwWorkbookError("The output does not fit the destination worksheet.");
  const outputRange = {
    firstRow: destinationPosition.row,
    lastRow: lastRow,
    firstCol: destinationPosition.col,
    lastCol: lastCol,
  };
  for (const cell of Object.keys(sheet.cells)) {
    const existingPosition = _t(cell);
    Dp(outputRange, existingPosition.row, existingPosition.col) &&
      throwWorkbookError(
        "The output destination contains existing cells. Choose an empty destination or a new sheet.",
      );
  }
  return (
    sheet.tables?.some((table) =>
      rangesIntersect(outputRange, parseRangeAddress(table.range)),
    ) && throwWorkbookError("The output destination overlaps a table."),
    (sheet.merges?.some((merge) =>
      rangesIntersect(outputRange, parseRangeAddress(merge)),
    ) ||
      sheet.pivots?.some((pivot) =>
        rangesIntersect(outputRange, parseRangeAddress(pivot.outputRange)),
      ) ||
      sheet.queries?.some((query) =>
        rangesIntersect(outputRange, parseRangeAddress(query.outputRange)),
      )) &&
      throwWorkbookError(
        "The output destination overlaps an existing merged area or analysis result. Choose another empty destination.",
      ),
    matrix.forEach((rowValues, rowIndex) =>
      rowValues.forEach((raw, columnIndex) => {
        sheet.cells[
          Cn(
            destinationPosition.row + rowIndex,
            destinationPosition.col + columnIndex,
          )
        ] = {
          raw: raw,
          ...(typeof raw == "string" && raw.startsWith("=")
            ? { literal: true }
            : {}),
          ...(rowIndex === 0
            ? { style: { bold: true, fill: "#174a7e", color: "#ffffff" } }
            : isFiniteNumber(raw)
              ? { format: "number" }
              : {}),
        };
      }),
    ),
    assertWithinCellBudget(workbook),
    {
      workbook: workbook,
      sheet: sheet,
      range: `${Cn(destinationPosition.row, destinationPosition.col)}:${Cn(lastRow, lastCol)}`,
    }
  );
}
function buildPivotTableOutput(workbook, pivotRequest) {
  const pivotDefinition = normalizePivotDefinition(pivotRequest),
    pivotSource = buildPivotSource(workbook, pivotDefinition.source),
    rows = filterPivotSourceRows(
      pivotDefinition,
      pivotSource.fields,
      filterSourceRows(pivotSource, pivotDefinition.filters ?? []),
    ),
    pivotTable = computePivotTable(pivotDefinition, pivotSource.fields, rows),
    outputBlock = writeAnalysisOutputBlock(
      workbook,
      pivotDefinition.name,
      pivotTable.matrix,
      pivotDefinition.destination,
    ),
    { destination, ...pivotSettings } = pivotDefinition,
    destinationPosition = _t(pivotDefinition.destination?.address ?? "A1");
  for (let rowIndex = 1; rowIndex < pivotTable.matrix.length; rowIndex++)
    for (
      let columnIndex = Math.max(1, pivotDefinition.rowFields.length);
      columnIndex < pivotTable.matrix[0].length;
      columnIndex++
    ) {
      const valueField =
        pivotDefinition.values[
          (columnIndex - Math.max(1, pivotDefinition.rowFields.length)) %
            pivotDefinition.values.length
        ];
      valueField.showAs &&
        valueField.showAs !== "value" &&
        (outputBlock.sheet.cells[
          Cn(
            destinationPosition.row + rowIndex,
            destinationPosition.col + columnIndex,
          )
        ].format = "percent");
    }
  if (!pivotDefinition.destination?.sheetId) {
    let headerRowHeight = 26;
    (pivotTable.matrix[0].forEach((headerValue, headerColumnIndex) => {
      const headerText = String(headerValue ?? ""),
        columnWidth = Math.min(320, Math.max(90, headerText.length * 7 + 24));
      outputBlock.sheet.columnWidths = {
        ...outputBlock.sheet.columnWidths,
        [destinationPosition.col + headerColumnIndex]: columnWidth,
      };
      const headerAddress = Cn(
          destinationPosition.row,
          destinationPosition.col + headerColumnIndex,
        ),
        headerCell = outputBlock.sheet.cells[headerAddress];
      ((headerCell.style = { ...headerCell.style, wrap: true }),
        (headerRowHeight = Math.max(
          headerRowHeight,
          Math.min(
            120,
            Math.ceil((headerText.length * 7) / (columnWidth - 16)) * 18 + 8,
          ),
        )));
    }),
      (outputBlock.sheet.rowHeights = {
        ...outputBlock.sheet.rowHeights,
        [destinationPosition.row]: headerRowHeight,
      }));
  }
  for (const subtotalRow of pivotTable.subtotalRows)
    for (
      let subtotalColumnIndex = 0;
      subtotalColumnIndex < pivotTable.matrix[0].length;
      subtotalColumnIndex++
    ) {
      const subtotalAddress = Cn(
          destinationPosition.row + subtotalRow,
          destinationPosition.col + subtotalColumnIndex,
        ),
        subtotalCell = outputBlock.sheet.cells[subtotalAddress];
      subtotalCell.style = {
        ...subtotalCell.style,
        bold: true,
        fill: "#e4edf7",
      };
    }
  return (
    (outputBlock.sheet.pivots = [
      ...(outputBlock.sheet.pivots ?? []),
      {
        ...pivotSettings,
        source: pivotSource.source,
        id: createUuid(),
        outputRange: outputBlock.range,
        sourceRevision: workbook.revision,
        refresh: {
          at: new Date().toISOString(),
          sourceSheetName: findSheetOrThrow(
            workbook,
            pivotSource.source.sheetId,
          ).name,
          sourceRange: pivotSource.source.range,
          sourceFingerprint: aggregatePivotRows(
            workbook,
            { ...pivotDefinition, source: pivotSource.source },
            pivotSource,
          ),
          outputFingerprint: collectRangeValues(
            outputBlock.sheet,
            outputBlock.range,
          ),
          sourceRows: pivotSource.rows.length,
          includedRows: rows.length,
          excludedRows: pivotSource.rows.length - rows.length,
          measures: pivotTable.measures,
        },
      },
    ]),
    outputBlock.workbook
  );
}
function refreshPivotTable(workbook, pivotId, destination) {
  const pivotDefinition =
    workbook.sheets
      .flatMap((sheet) => sheet.pivots ?? [])
      .find((pivot) => pivot.id === pivotId) ??
    throwWorkbookError("Pivot definition not found.");
  return buildPivotTableOutput(workbook, {
    ...pivotDefinition,
    destination: destination,
    name: `${pivotDefinition.name.slice(0, 90)} refreshed`,
  });
}
function applyPivotViewChange(workbook, pivotId, viewChanges, previewContext) {
  (previewContext.readOnly && throwWorkbookError("This workbook is read-only."),
    previewContext.snapshot !== JSON.stringify(workbook) &&
      throwWorkbookError("The workbook changed. Preview pivot filters again."),
    (!isPlainDataObject(viewChanges) ||
      Object.keys(viewChanges).some(
        (changeKey) =>
          !["slicers", "timelines", "sorts", "values", "chart"].includes(
            changeKey,
          ),
      )) &&
      throwWorkbookError(
        "Pivot view changes can only set filters, sorting, value display and chart settings.",
      ));
  const { owner, definition } = findPivotDefinition(workbook, pivotId);
  return (
    assertSheetUnprotected(owner),
    buildPivotTableOutput(workbook, {
      ...definition,
      ...viewChanges,
      name: `${definition.name.slice(0, 88)} filtered`,
      destination: undefined,
    })
  );
}
function extractPivotDrillThrough(
  workbook,
  pivotId,
  drillRequest,
  previewContext,
) {
  (previewContext.readOnly && throwWorkbookError("This workbook is read-only."),
    previewContext.snapshot !== JSON.stringify(workbook) &&
      throwWorkbookError("The workbook changed. Preview drill-through again."));
  const pivotDrillDown = buildPivotDrillDown(workbook, pivotId, drillRequest),
    { owner } = findPivotDefinition(workbook, pivotId);
  assertSheetUnprotected(owner);
  const contextJson = JSON.stringify({ ...pivotDrillDown, records: undefined });
  contextJson.length > 32700 &&
    throwWorkbookError(
      "Drill context exceeds the cell-note budget. Reduce filter selections before extracting.",
    );
  const drillRows = [
      [
        "Source locator",
        "Source revision",
        "Original source",
        "Original sheet",
        "Original row",
        ...pivotDrillDown.fields,
      ],
      ...pivotDrillDown.records.map((record) => [
        `${record.locator.sheetName}!${record.locator.range}`,
        pivotDrillDown.context.sourceRevision,
        record.locator.original?.sourceName ?? null,
        record.locator.original?.sourceSheet ?? null,
        record.locator.original?.originalRow ?? null,
        ...record.values,
      ]),
    ],
    outputBlock = writeAnalysisOutputBlock(
      workbook,
      `${pivotDrillDown.context.name.slice(0, 18)} details`,
      drillRows,
    );
  return (
    (outputBlock.sheet.cells.A1.note = `Pivot drill-through context (snapshot, not signed): ${contextJson}`),
    (outputBlock.sheet.columnWidths = {
      0: 240,
      1: 110,
      2: 160,
      3: 130,
      4: 110,
    }),
    (outputBlock.sheet.rowOriginIds = {}),
    (outputBlock.workbook.rowOrigins = [...(workbook.rowOrigins ?? [])]),
    pivotDrillDown.records.forEach((record, recordIndex) => {
      let original = record.locator.original;
      (original ||
        ((original = {
          id: createUuid(),
          sourceId: workbook.id,
          sourceName: workbook.name,
          sourceSheet: record.locator.sheetName,
          originalRow: record.locator.row,
          kind: "snapshot",
        }),
        outputBlock.workbook.rowOrigins.length >= workbookLimits.maxCells &&
          throwWorkbookError(
            "Row provenance exceeds its 100,000 identity budget.",
          ),
        outputBlock.workbook.rowOrigins.push(original)),
        (outputBlock.sheet.rowOriginIds[recordIndex + 1] = original.id),
        record.cells.forEach(({ cell }, columnIndex) => {
          const noteText = [
            cell.note,
            typeof cell.raw == "string" &&
            cell.raw.startsWith("=") &&
            !cell.literal
              ? `Source formula at extraction: ${cell.raw}`
              : undefined,
          ].filter(Boolean).join(`
`);
          (noteText.length > 32767 &&
            throwWorkbookError(
              "A source note/formula exceeds the drill-output note budget.",
            ),
            noteText &&
              (outputBlock.sheet.cells[
                `${columnIndexToLetters(columnIndex + 5)}${recordIndex + 2}`
              ].note = noteText));
        }));
    }),
    outputBlock.workbook.sheets.reduce(
      (total, sheet) => total + Object.keys(sheet.rowOriginIds ?? {}).length,
      0,
    ) > workbookLimits.maxCells &&
      throwWorkbookError(
        "Drill row associations exceed the 100,000 source-row budget.",
      ),
    outputBlock.workbook
  );
}
function runQueryDefinition(workbook, queryRequest) {
  const queryDefinition = normalizeQueryDefinition(queryRequest),
    pivotSource = buildPivotSource(workbook, queryDefinition.source),
    filteredRows = filterSourceRows(pivotSource, queryDefinition.filters ?? []),
    fieldIndexes = queryDefinition.fields.map((field) =>
      findSourceFieldIndex(pivotSource.fields, field),
    ),
    derivedFields = (queryDefinition.derived ?? []).map(
      (derivedDefinition) => ({
        ...derivedDefinition,
        index: findSourceFieldIndex(
          pivotSource.fields,
          derivedDefinition.field,
        ),
        other: derivedDefinition.otherField
          ? findSourceFieldIndex(
              pivotSource.fields,
              derivedDefinition.otherField,
            )
          : undefined,
      }),
    );
  if (queryDefinition.sort) {
    const sourceFieldIndex = findSourceFieldIndex(
        pivotSource.fields,
        queryDefinition.sort.field,
      ),
      sortMultiplier = queryDefinition.sort.direction === "ascending" ? 1 : -1;
    filteredRows.sort((leftRow, rightRow) => {
      const leftValue = leftRow.values[sourceFieldIndex],
        rightValue = rightRow.values[sourceFieldIndex];
      return (
        (isFiniteNumber(leftValue) && isFiniteNumber(rightValue)
          ? leftValue - rightValue
          : String(leftValue ?? "").localeCompare(String(rightValue ?? ""))) *
          sortMultiplier || leftRow.index - rightRow.index
      );
    });
  }
  const headerRow = [
      ...queryDefinition.fields.map((field) =>
        Object.hasOwn(queryDefinition.rename ?? {}, field)
          ? queryDefinition.rename[field]
          : field,
      ),
      ...derivedFields.map((derivedEntry) => derivedEntry.name),
    ],
    outputMatrix = [headerRow];
  (filteredRows.length + 1) * headerRow.length > workbookLimits.maxCells &&
    throwWorkbookError("The query output exceeds 100,000 cells.");
  for (const sourceRow of filteredRows) {
    const outputRow = fieldIndexes.map(
      (fieldIndex) => sourceRow.values[fieldIndex],
    );
    for (const derivedField of derivedFields) {
      const sourceValue = sourceRow.values[derivedField.index],
        operandValue =
          derivedField.other !== undefined
            ? sourceRow.values[derivedField.other]
            : (derivedField.constant ?? null);
      if (derivedField.operation === "upper")
        outputRow.push(String(sourceValue ?? "").toUpperCase());
      else if (derivedField.operation === "lower")
        outputRow.push(String(sourceValue ?? "").toLowerCase());
      else if (derivedField.operation === "trim")
        outputRow.push(String(sourceValue ?? "").trim());
      else if (derivedField.operation === "concat")
        outputRow.push(String(sourceValue ?? "") + String(operandValue ?? ""));
      else {
        ((!isFiniteNumber(sourceValue) || !isFiniteNumber(operandValue)) &&
          throwWorkbookError(
            `Derived field “${derivedField.name}” requires numbers on source row ${sourceRow.index + 1}.`,
          ),
          derivedField.operation === "divide" &&
            operandValue === 0 &&
            throwWorkbookError(
              `Derived field “${derivedField.name}” divides by zero on source row ${sourceRow.index + 1}.`,
            ));
        const computedValue =
          derivedField.operation === "add"
            ? sourceValue + operandValue
            : derivedField.operation === "subtract"
              ? sourceValue - operandValue
              : derivedField.operation === "multiply"
                ? sourceValue * operandValue
                : sourceValue / operandValue;
        (Number.isFinite(computedValue) ||
          throwWorkbookError(
            "Derived calculation exceeds finite numeric bounds.",
          ),
          outputRow.push(computedValue));
      }
    }
    (outputRow.some(
      (cellValue) => typeof cellValue == "string" && cellValue.length > 32767,
    ) && throwWorkbookError("Derived text exceeds the cell text limit."),
      outputMatrix.push(outputRow));
  }
  const outputBlock = writeAnalysisOutputBlock(
      workbook,
      queryDefinition.name,
      outputMatrix,
      queryDefinition.destination,
    ),
    { destination, ...querySettings } = queryDefinition;
  return (
    (outputBlock.sheet.queries = [
      ...(outputBlock.sheet.queries ?? []),
      {
        ...querySettings,
        source: pivotSource.source,
        id: createUuid(),
        sourceRevision: workbook.revision,
        outputRange: outputBlock.range,
      },
    ]),
    outputBlock.workbook
  );
}
function refreshQueryDefinition(workbook, pivotId, destination) {
  const queryDefinition =
    workbook.sheets
      .flatMap((sheet) => sheet.queries ?? [])
      .find((query) => query.id === pivotId) ??
    throwWorkbookError("Query definition not found.");
  return runQueryDefinition(workbook, {
    ...queryDefinition,
    destination: destination,
    name: `${queryDefinition.name.slice(0, 90)} refreshed`,
  });
}
function normalizeCellRecord(cell, canReuseRecord) {
  if (
    canReuseRecord &&
    isDeepFrozenValue(cell) &&
    JU.has(cell) &&
    !cell.hyperlink &&
    !cell.style?.binding
  )
    return cell;
  if (!isNonArrayObject(cell) || !hasOwnKey(cell, "raw"))
    return throwWithMessage("A cell must have a raw value.");
  const raw = cell.raw;
  if (
    raw !== null &&
    typeof raw != "string" &&
    typeof raw != "boolean" &&
    (typeof raw != "number" || !Number.isFinite(raw))
  )
    return throwWithMessage(
      "Cell values must be text, finite numbers, booleans, or null.",
    );
  typeof raw == "string" &&
    requireBoundedText(
      raw,
      "Cell value",
      raw.startsWith("=") && cell.literal !== true
        ? fileImportLimits.maxFormulaLength
        : fileImportLimits.maxCellTextLength,
    );
  const normalizedCell = { raw: raw };
  if (cell.literal !== undefined) {
    if (typeof cell.literal != "boolean")
      return throwWithMessage("Cell literal flag must be boolean.");
    normalizedCell.literal = cell.literal;
  }
  if (
    (cell.numberFormat !== undefined &&
      (normalizedCell.numberFormat = validateNumberFormatSettings(
        cell.numberFormat,
      )),
    cell.numFmt !== undefined &&
      (normalizedCell.numFmt = requireBoundedText(
        cell.numFmt,
        "Excel number format",
        200,
      )),
    cell.format !== undefined)
  ) {
    if (
      typeof cell.format != "string" ||
      !numberFormatCategories.has(cell.format)
    )
      return throwWithMessage("Unsupported cell format.");
    normalizedCell.format = cell.format;
  }
  return (
    cell.hyperlink !== undefined &&
      (normalizedCell.hyperlink = validateHyperlinkRecord(cell.hyperlink)),
    cell.note !== undefined &&
      (normalizedCell.note = requireBoundedText(cell.note, "Cell note", 1e4)),
    cell.tickmark !== undefined &&
      (normalizedCell.tickmark = requireBoundedText(
        cell.tickmark,
        "Tickmark",
        40,
      )),
    cell.style !== undefined &&
      (normalizedCell.style = validateCellStyle(cell.style)),
    normalizedCell
  );
}
function normalizeCellRecordInput(cell) {
  return normalizeCellRecord(cell, false);
}
function paintCellFormatting(workbook, sheetId, sourceAddress, targetRange) {
  const sheet = workbook.sheets.find((sheet) => sheet.id === sheetId);
  assertSheetUnprotected(sheet);
  const targetAddresses = assertRangeCellBudget(targetRange);
  targetAddresses.length > 500 && failWithMessage("Paint at most 500 cells.");
  const sourceCell = sheet.cells[sourceAddress] ?? {},
    cells = { ...sheet.cells };
  for (const targetAddress of targetAddresses)
    cells[targetAddress] = normalizeCellRecordInput({
      ...cells[targetAddress],
      raw: cells[targetAddress]?.raw ?? null,
      style: structuredClone(sourceCell.style ?? {}),
      format: sourceCell.format ?? "general",
      numFmt: sourceCell.numFmt,
      numberFormat: sourceCell.numberFormat
        ? structuredClone(sourceCell.numberFormat)
        : undefined,
    });
  return {
    ...workbook,
    sheets: workbook.sheets.map((sheet) =>
      sheet.id === sheetId ? { ...sheet, cells: cells } : sheet,
    ),
  };
}
function validateTableSourceEvidence(tableEvidence, targetWorkbook) {
  if (tableEvidence === undefined) return;
  const rejectEvidence = () => {
      throw Error(
        "Invalid or over-capacity table source evidence. Keep the native backup.",
      );
    },
    isObjectRecord = (candidate) =>
      !!candidate && typeof candidate == "object" && !Array.isArray(candidate);
  (!Array.isArray(tableEvidence) ||
    tableEvidence.length > 100 ||
    JSON.stringify(tableEvidence).length > 5e6) &&
    rejectEvidence();
  const seenIds = new Set();
  let totalChangeCount = 0;
  targetWorkbook.tableEvidence = tableEvidence.map((evidenceEntry) => {
    (!isObjectRecord(evidenceEntry) ||
      Object.keys(evidenceEntry).some(
        (fieldName) =>
          ![
            "id",
            "kind",
            "sourceRevision",
            "sheetId",
            "sheetName",
            "tableId",
            "tableName",
            "columnId",
            "dataType",
            "changes",
          ].includes(fieldName),
      )) &&
      rejectEvidence();
    const evidenceRecord = evidenceEntry;
    for (const textFieldName of [
      "id",
      "sheetId",
      "sheetName",
      "tableId",
      "tableName",
    ])
      (typeof evidenceRecord[textFieldName] != "string" ||
        !evidenceRecord[textFieldName] ||
        evidenceRecord[textFieldName].length > 512) &&
        rejectEvidence();
    ((seenIds.has(evidenceRecord.id) ||
      !Number.isSafeInteger(evidenceRecord.sourceRevision) ||
      Number(evidenceRecord.sourceRevision) < 0 ||
      !["headers", "columnType"].includes(String(evidenceRecord.kind)) ||
      !Array.isArray(evidenceRecord.changes) ||
      evidenceRecord.changes.length > 500 ||
      (totalChangeCount += evidenceRecord.changes.length) > 1e4) &&
      rejectEvidence(),
      evidenceRecord.kind === "columnType"
        ? ((typeof evidenceRecord.columnId != "string" ||
            !evidenceRecord.columnId ||
            evidenceRecord.columnId.length > 100) &&
            rejectEvidence(),
          validateTableColumnType(evidenceRecord.dataType))
        : (evidenceRecord.columnId !== undefined ||
            evidenceRecord.dataType !== undefined) &&
          rejectEvidence(),
      seenIds.add(evidenceRecord.id));
    const seenAddresses = new Set(),
      changes = evidenceRecord.changes.map((change) => {
        (!isObjectRecord(change) ||
          Object.keys(change).some(
            (changeFieldName) =>
              !["address", "before", "after"].includes(changeFieldName),
          ) ||
          typeof change.address != "string") &&
          rejectEvidence();
        const changeRecord = change,
          address = changeRecord.address,
          addressPosition = _t(address);
        return (
          (`${columnIndexToLetters(addressPosition.col)}${addressPosition.row + 1}` !==
            address ||
            seenAddresses.has(address)) &&
            rejectEvidence(),
          seenAddresses.add(address),
          {
            address: address,
            ...(changeRecord.before !== undefined
              ? { before: normalizeCellRecordInput(changeRecord.before) }
              : {}),
            after: normalizeCellRecordInput(changeRecord.after),
          }
        );
      });
    return { ...evidenceRecord, changes: changes };
  });
}
function previewColumnTypeChange(
  workbook,
  sheetId,
  tableId,
  request,
  options = {},
) {
  if (options.readOnly) throw Error("This workbook is read-only.");
  if (
    !request ||
    typeof request != "object" ||
    Object.keys(request).some(
      (requestKey) =>
        ![
          "columnId",
          "dataType",
          "mode",
          "invalid",
          "decimal",
          "group",
        ].includes(requestKey),
    )
  )
    throw Error("Invalid column type request.");
  if (
    (validateTableColumnType(request.dataType),
    !["convert", "tag"].includes(request.mode) ||
      !["reject", "keep"].includes(request.invalid) ||
      ![undefined, ".", ","].includes(request.decimal) ||
      ![undefined, "", ".", ",", " "].includes(request.group) ||
      (request.decimal ?? ".") === (request.group ?? ","))
  )
    throw Error(
      "Choose a conversion mode, exception resolution and different decimal/group separators.",
    );
  const columnTypeRequest = structuredClone(request),
    sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    targetTable = sheetOrThrow.tables?.find((table) => table.id === tableId);
  if ((assertSheetUnprotected(sheetOrThrow), !targetTable))
    throw Error("Table not found.");
  const columnIndex = targetTable.columns.findIndex(
    (column) => column.id === columnTypeRequest.columnId,
  );
  if (columnIndex < 0) throw Error("Select an existing table column.");
  const rangeAddress = parseRangeAddress(targetTable.range),
    dataRowCount =
      rangeAddress.lastRow -
      rangeAddress.firstRow -
      Number(targetTable.showTotals);
  if (dataRowCount > 500)
    throw Error(
      "A reviewed column type operation supports at most 500 data rows. Preserve the native source and use bounded extracts for larger conversions.",
    );
  const next = structuredClone(workbook),
    nextSheet = findSheetOrThrow(next, sheetId),
    nextColumn = nextSheet.tables.find((table) => table.id === tableId).columns[
      columnIndex
    ];
  columnTypeRequest.dataType === "any"
    ? delete nextColumn.dataType
    : (nextColumn.dataType = columnTypeRequest.dataType);
  const changes = [],
    exceptions = [],
    workbookWithSettings = calculateWorkbookWithSettings(workbook);
  for (let dataRowIndex = 1; dataRowIndex <= dataRowCount; dataRowIndex++) {
    const address = `${columnIndexToLetters(rangeAddress.firstCol + columnIndex)}${rangeAddress.firstRow + dataRowIndex + 1}`,
      cell = sheetOrThrow.cells[address],
      rawValue = cell?.raw ?? null,
      isFormula =
        typeof rawValue == "string" &&
        rawValue.startsWith("=") &&
        !cell?.literal;
    try {
      if (isFormula || columnTypeRequest.mode === "tag") {
        if (
          !coerceImportedValue(
            workbookWithSettings(sheetId, address),
            columnTypeRequest.dataType,
          )
        )
          throw Error(
            isFormula
              ? "Formula result does not match the selected type; the formula is preserved."
              : "Value does not match the selected type; source values are unchanged.",
          );
        continue;
      }
      const computedValue = workbookWithSettings(sheetId, address);
      if (isNullOrEmptyValue(rawValue) && !isNullOrEmptyValue(computedValue)) {
        if (!coerceImportedValue(computedValue, columnTypeRequest.dataType))
          throw Error(
            "Spill result does not match the selected type; the spill is preserved.",
          );
        continue;
      }
      const raw = convertTextForDataType(rawValue, columnTypeRequest);
      if (raw !== rawValue) {
        const cellRecordInput = normalizeCellRecordInput({
          ...cell,
          raw: raw,
          ...(typeof raw == "string" && raw.startsWith("=")
            ? { literal: true }
            : { literal: undefined }),
        });
        ((nextSheet.cells[address] = cellRecordInput),
          changes.push({
            address: address,
            ...(cell ? { before: structuredClone(cell) } : {}),
            after: structuredClone(cellRecordInput),
          }));
      }
    } catch (error) {
      exceptions.push({
        address: address,
        value: rawValue,
        reason: error.message,
      });
    }
  }
  const beforeType = targetTable.columns[columnIndex].dataType ?? "any";
  return (
    (changes.length || beforeType !== columnTypeRequest.dataType) &&
      ((next.tableEvidence = [
        ...(next.tableEvidence ?? []),
        {
          id: `${tableId}:${nextColumn.id}:${workbook.revision}:${workbook.tableEvidence?.length ?? 0}`,
          kind: "columnType",
          sourceRevision: workbook.revision,
          sheetId: sheetId,
          sheetName: sheetOrThrow.name,
          tableId: tableId,
          tableName: targetTable.name,
          columnId: nextColumn.id,
          dataType: columnTypeRequest.dataType,
          changes: structuredClone(changes),
        },
      ]),
      validateTableSourceEvidence(next.tableEvidence, next)),
    checkDataValidationBatch(
      next,
      changes.map((change) => ({ ...change, sheetId: sheetId })),
    ),
    {
      snapshot: JSON.stringify(workbook),
      sheetId: sheetId,
      tableId: tableId,
      request: columnTypeRequest,
      beforeType: beforeType,
      canApply: !exceptions.length || columnTypeRequest.invalid === "keep",
      changes: changes,
      exceptions: exceptions,
      next: next,
    }
  );
}
function applyColumnTypePreview(workbook, plan, options = {}) {
  if (JSON.stringify(workbook) !== plan.snapshot)
    throw Error("The workbook changed. Preview column types again.");
  const freshPreview = previewColumnTypeChange(
    workbook,
    plan.sheetId,
    plan.tableId,
    plan.request,
    options,
  );
  if (JSON.stringify(freshPreview) !== JSON.stringify(plan))
    throw Error("The column type preview changed. Preview again.");
  if (!freshPreview.canApply)
    throw Error(
      "Resolve the listed exceptions or explicitly choose to retain them unchanged.",
    );
  return JSON.stringify(workbook) === JSON.stringify(freshPreview.next)
    ? workbook
    : freshPreview.next;
}
function createWorksheetTable(workbook, sheetId, options) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    tableName = validateTableName(options.name),
    sourceRange = resolveSourceRange(options.range, sheetOrThrow);
  if (
    (assertSheetUnprotected(sheetOrThrow),
    sheetOrThrow.merges?.some((merge) =>
      rangesIntersect(sourceRange, parseRangeAddress(merge)),
    ) &&
      throwWorkbookError(
        "A table cannot overlap merged cells. Unmerge the source first.",
      ),
    workbook.sheets.some((sheet) =>
      sheet.tables?.some(
        (table) => table.name.toLowerCase() === tableName.toLowerCase(),
      ),
    ) && throwWorkbookError("A table with that name already exists."),
    options.showTotals)
  ) {
    sourceRange.lastRow + 1 >= sheetOrThrow.rowCount &&
      throwWorkbookError("There is no room for a totals row.");
    for (
      let firstCol = sourceRange.firstCol;
      firstCol <= sourceRange.lastCol;
      firstCol++
    )
      sheetOrThrow.cells[Cn(sourceRange.lastRow + 1, firstCol)] &&
        throwWorkbookError("The totals row would overwrite existing cells.");
    sourceRange.lastRow++;
  }
  (measureRangeArea(sourceRange) > workbookLimits.maxCells ||
    sheetOrThrow.tables?.some((table) =>
      rangesIntersect(sourceRange, parseRangeAddress(table.range)),
    )) &&
    throwWorkbookError(
      "The table overlaps another table or exceeds the cell budget.",
    );
  const nextWorkbook = structuredClone(workbook),
    nextSheet = findSheetOrThrow(nextWorkbook, sheetId),
    newTable = {
      id: createUuid(),
      name: tableName,
      range: `${Cn(sourceRange.firstRow, sourceRange.firstCol)}:${Cn(sourceRange.lastRow, sourceRange.lastCol)}`,
      columns: validateReviewedTableHeaders(
        workbook,
        sheetId,
        options.range,
        options.headers,
      ).map((headerName) => ({ id: createUuid(), name: headerName })),
      style: options.style ?? "blue",
      showTotals: !!options.showTotals,
      totals: options.totals ?? {},
    };
  ((nextSheet.tables = [...(nextSheet.tables ?? []), newTable]),
    normalizeTableDefinitions(nextSheet, nextSheet),
    applyTableStyleToCells(
      nextSheet,
      nextSheet.tables[nextSheet.tables.length - 1],
    ));
  const changes = newTable.columns.flatMap((column, columnIndex) => {
    const address = Cn(
        sourceRange.firstRow,
        sourceRange.firstCol + columnIndex,
      ),
      cell = sheetOrThrow.cells[address],
      nextCell = nextSheet.cells[address];
    return cell?.raw === column.name &&
      !(column.name.startsWith("=") && !cell?.literal)
      ? []
      : [
          {
            address: address,
            ...(cell ? { before: structuredClone(cell) } : {}),
            after: structuredClone(nextCell),
          },
        ];
  });
  return (
    changes.length &&
      ((nextWorkbook.tableEvidence = [
        ...(nextWorkbook.tableEvidence ?? []),
        {
          id: `${newTable.id}:headers`,
          kind: "headers",
          sourceRevision: workbook.revision,
          sheetId: sheetId,
          sheetName: sheetOrThrow.name,
          tableId: newTable.id,
          tableName: tableName,
          changes: changes,
        },
      ]),
      validateTableSourceEvidence(nextWorkbook.tableEvidence, nextWorkbook)),
    assertWithinCellBudget(nextWorkbook)
  );
}
function fillCalculatedTableColumns(sheet, table, firstRow, lastRow) {
  const bodyRange = tableDataBodyRange(table);
  for (const [columnIndex, column] of table.columns.entries())
    if (column.calculatedFormula)
      for (let rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
        const address = Cn(rowIndex, bodyRange.firstCol + columnIndex),
          raw = shiftFormulaReferences(
            column.calculatedFormula,
            rowIndex - bodyRange.firstRow,
            0,
          );
        validateCalculatedColumnFormula(raw);
        const cell = { ...(sheet.cells[address] ?? { raw: null }), raw: raw };
        (delete cell.literal,
          (sheet.cells[address] = normalizeCellRecordInput(cell)));
      }
}
function resizeTable(workbook, sheetId, tableId, options) {
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    existingTable =
      sheetOrThrow.tables?.find((table) => table.id === tableId) ??
      throwWorkbookError("Table not found.");
  assertSheetUnprotected(sheetOrThrow);
  const tableName = validateTableName(options.name),
    sourceRange = resolveSourceRange(options.range, sheetOrThrow),
    rangeAddress = parseRangeAddress(existingTable.range);
  ((sourceRange.firstRow !== rangeAddress.firstRow ||
    sourceRange.firstCol !== rangeAddress.firstCol) &&
    throwWorkbookError(
      "Resize a table from its existing top-left header. Moving table source cells is a separate operation.",
    ),
    workbook.sheets.some((sheet) =>
      sheet.tables?.some(
        (table) =>
          table.id !== tableId &&
          table.name.toLowerCase() === tableName.toLowerCase(),
      ),
    ) && throwWorkbookError("A table with that name already exists."));
  const showTotals = options.showTotals ?? existingTable.showTotals;
  if (
    (showTotals !== existingTable.showTotals &&
      toA1RangeAddress(options.range) === existingTable.range &&
      (sourceRange.lastRow += showTotals ? 1 : -1),
    (sourceRange.lastRow >= sheetOrThrow.rowCount ||
      sourceRange.lastRow < sourceRange.firstRow ||
      measureRangeArea(sourceRange) > workbookLimits.maxCells) &&
      throwWorkbookError(
        "The resized table does not fit the worksheet or processing budget.",
      ),
    sheetOrThrow.tables?.some(
      (table) =>
        table.id !== tableId &&
        rangesIntersect(sourceRange, parseRangeAddress(table.range)),
    ) && throwWorkbookError("Tables cannot overlap."),
    showTotals &&
      (!existingTable.showTotals ||
        sourceRange.lastRow !== rangeAddress.lastRow))
  )
    for (
      let firstCol = sourceRange.firstCol;
      firstCol <= sourceRange.lastCol;
      firstCol++
    )
      sheetOrThrow.cells[Cn(sourceRange.lastRow, firstCol)] &&
        throwWorkbookError(
          "The new totals row would overwrite existing cells. Include an empty final row in the resized range.",
        );
  const hasCalculatedColumns = existingTable.columns.some(
    (column) => column.calculatedFormula,
  );
  (hasCalculatedColumns &&
    sourceRange.lastCol !== rangeAddress.lastCol &&
    throwWorkbookError(
      "Clear calculated column rules before changing table columns.",
    ),
    hasCalculatedColumns &&
      (sourceRange.lastRow !== rangeAddress.lastRow ||
        showTotals !== existingTable.showTotals) &&
      resizeTableRange(
        workbook,
        sheetId,
        existingTable,
        sourceRange,
        showTotals,
      ));
  const nextWorkbook = structuredClone(workbook),
    nextSheet = findSheetOrThrow(nextWorkbook, sheetId);
  if (
    existingTable.showTotals &&
    (!showTotals || sourceRange.lastRow !== rangeAddress.lastRow)
  )
    for (
      let firstCol = rangeAddress.firstCol;
      firstCol <= rangeAddress.lastCol;
      firstCol++
    ) {
      const totalsAddress = Cn(rangeAddress.lastRow, firstCol),
        cell = nextSheet.cells[totalsAddress],
        tableTotalCell = buildTableTotalCell(
          existingTable,
          firstCol - rangeAddress.firstCol,
        );
      (cell &&
        cell.raw !== tableTotalCell.raw &&
        throwWorkbookError(
          "The old totals row was edited. Preserve or move those values before resizing it.",
        ),
        delete nextSheet.cells[totalsAddress]);
    }
  const headerNames = validateReviewedTableHeaders(
      nextWorkbook,
      sheetId,
      options.range,
      options.headers,
    ),
    columns = headerNames.map((headerName, columnIndex) => ({
      ...existingTable.columns[columnIndex],
      id:
        existingTable.columns[columnIndex]?.id ??
        `${existingTable.id}:column:${columnIndex}:${workbook.revision}`,
      name: headerName,
    })),
    totals =
      options.totals ??
      Object.fromEntries(
        Object.entries(existingTable.totals).filter(([columnName]) =>
          columns.some((column) => column.name === columnName),
        ),
      ),
    updatedTable = {
      ...existingTable,
      name: tableName,
      range: `${Cn(sourceRange.firstRow, sourceRange.firstCol)}:${Cn(sourceRange.lastRow, sourceRange.lastCol)}`,
      columns: columns,
      style: options.style ?? existingTable.style,
      showTotals: showTotals,
      totals: totals,
    };
  ((nextSheet.tables = nextSheet.tables.map((table) =>
    table.id === tableId ? updatedTable : table,
  )),
    normalizeTableDefinitions(nextSheet, nextSheet),
    applyTableStyleToCells(nextSheet, updatedTable));
  const changes = columns.flatMap((column, columnOffset) => {
    const address = Cn(
        sourceRange.firstRow,
        sourceRange.firstCol + columnOffset,
      ),
      cell = sheetOrThrow.cells[address],
      nextCell = nextSheet.cells[address];
    return cell?.raw === nextCell.raw && !!cell?.literal == !!nextCell.literal
      ? []
      : [
          {
            address: address,
            ...(cell ? { before: structuredClone(cell) } : {}),
            after: structuredClone(nextCell),
          },
        ];
  });
  if (
    (changes.length &&
      ((nextWorkbook.tableEvidence = [
        ...(nextWorkbook.tableEvidence ?? []),
        {
          id: `${tableId}:headers:${workbook.revision}:${nextWorkbook.tableEvidence?.length ?? 0}`,
          kind: "headers",
          sourceRevision: workbook.revision,
          sheetId: sheetId,
          sheetName: sheetOrThrow.name,
          tableId: tableId,
          tableName: tableName,
          changes: changes,
        },
      ]),
      validateTableSourceEvidence(nextWorkbook.tableEvidence, nextWorkbook)),
    hasCalculatedColumns)
  ) {
    const previousBodyRange = tableDataBodyRange(existingTable),
      updatedBodyRange = tableDataBodyRange(updatedTable);
    for (
      let firstRow = rangeAddress.firstRow;
      sourceRange.lastRow !== rangeAddress.lastRow &&
      firstRow <= Math.min(previousBodyRange.lastRow, updatedBodyRange.lastRow);
      firstRow++
    )
      for (
        let firstCol = previousBodyRange.firstCol;
        firstCol <= previousBodyRange.lastCol;
        firstCol++
      ) {
        const bodyCellAddress = Cn(firstRow, firstCol);
        sheetOrThrow.cells[bodyCellAddress] &&
          (nextSheet.cells[bodyCellAddress] = structuredClone(
            sheetOrThrow.cells[bodyCellAddress],
          ));
      }
    (existingTable.showTotals &&
      showTotals &&
      sourceRange.lastRow !== rangeAddress.lastRow &&
      moveTotalsRowCells(
        workbook,
        nextWorkbook,
        sheetId,
        existingTable,
        updatedTable,
      ),
      fillCalculatedTableColumns(
        nextSheet,
        updatedTable,
        Math.max(previousBodyRange.lastRow + 1, updatedBodyRange.firstRow),
        updatedBodyRange.lastRow,
      ));
  }
  return assertWithinCellBudget(renameTableHeaderField(workbook, nextWorkbook));
}
function applyTableOperation(
  workbook,
  sheetId,
  tableId,
  request,
  options = {},
) {
  options.readOnly && throwWorkbookError("This workbook is read-only.");
  const sheetOrThrow = findSheetOrThrow(workbook, sheetId),
    existingTable =
      sheetOrThrow.tables?.find((table) => table.id === tableId) ??
      throwWorkbookError("Select an existing table.");
  (assertSheetUnprotected(sheetOrThrow),
    isPlainDataObject(request) ||
      throwWorkbookError("Choose a table operation."));
  const allowedFieldsByKind = {
    setColumnFormula: ["kind", "columnId", "formula"],
    clearColumnFormula: ["kind", "columnId"],
    appendRows: ["kind", "count", "rows"],
    resize: ["kind", "range"],
    updateTable: ["kind", "options"],
  };
  (!allowedFieldsByKind[request.kind] ||
    Object.keys(request).some(
      (fieldName) => !allowedFieldsByKind[request.kind].includes(fieldName),
    )) &&
    throwWorkbookError("Unknown calculated-table operation field.");
  let next = structuredClone(workbook),
    nextSheet = findSheetOrThrow(next, sheetId),
    table = nextSheet.tables.find((table) => table.id === tableId);
  if (
    request.kind === "setColumnFormula" ||
    request.kind === "clearColumnFormula"
  ) {
    const targetColumn =
      table.columns.find((column) => column.id === request.columnId) ??
      throwWorkbookError("Select an existing table column.");
    if (request.kind === "setColumnFormula") {
      targetColumn.calculatedFormula = validateCalculatedColumnFormula(
        request.formula,
      );
      const formulaBodyRange = tableDataBodyRange(table);
      (formulaBodyRange.lastRow - formulaBodyRange.firstRow + 1 >
        tableMutationLimits.changedCells &&
        throwWorkbookError(
          "A calculated column fill supports at most 500 data-body rows.",
        ),
        fillCalculatedTableColumns(
          nextSheet,
          {
            ...table,
            columns: table.columns.map((column) =>
              column.id === targetColumn.id
                ? column
                : { ...column, calculatedFormula: undefined },
            ),
          },
          formulaBodyRange.firstRow,
          formulaBodyRange.lastRow,
        ));
    } else delete targetColumn.calculatedFormula;
  } else if (
    request.kind === "appendRows" ||
    request.kind === "resize" ||
    request.kind === "updateTable"
  ) {
    const rangeAddress = parseRangeAddress(existingTable.range),
      existingBodyRange = tableDataBodyRange(existingTable);
    let rawRows,
      appendCount = 0,
      resizeOptions;
    request.kind === "appendRows"
      ? (request.rows !== undefined
          ? (request.count !== undefined &&
              throwWorkbookError("Specify raw rows or a row count, not both."),
            (!Array.isArray(request.rows) ||
              request.rows.length > tableMutationLimits.appendedRows ||
              request.rows.some(
                (row) =>
                  !Array.isArray(row) ||
                  row.length !== existingTable.columns.length,
              )) &&
              throwWorkbookError(
                "Each appended row must contain one raw value for every table column.",
              ),
            (rawRows = request.rows),
            (appendCount = rawRows.length))
          : (appendCount = request.count ?? 0),
        (!Number.isInteger(appendCount) ||
          appendCount < 1 ||
          appendCount > tableMutationLimits.appendedRows) &&
          throwWorkbookError("Append 1–100 rows at a time."),
        rangeAddress.lastRow + appendCount >= sheetOrThrow.rowCount &&
          throwWorkbookError(
            "The appended table exceeds the worksheet boundary.",
          ),
        (resizeOptions = {
          ...existingTable,
          range: `${Cn(rangeAddress.firstRow, rangeAddress.firstCol)}:${Cn(rangeAddress.lastRow + appendCount, rangeAddress.lastCol)}`,
        }))
      : (resizeOptions =
          request.kind === "resize"
            ? { ...existingTable, range: request.range }
            : request.options);
    const sourceRange = resolveSourceRange(resizeOptions.range, sheetOrThrow),
      showTotals = resizeOptions.showTotals ?? existingTable.showTotals;
    if (
      ((request.kind !== "updateTable" ||
        existingTable.columns.some((column) => column.calculatedFormula)) &&
        resizeTableRange(
          workbook,
          sheetId,
          existingTable,
          sourceRange,
          showTotals,
        ),
      (next = resizeTable(workbook, sheetId, tableId, resizeOptions)),
      (nextSheet = findSheetOrThrow(next, sheetId)),
      (table = nextSheet.tables.find((table) => table.id === tableId)),
      request.kind === "appendRows" &&
        !existingTable.columns.some((column) => column.calculatedFormula))
    ) {
      for (
        let firstRow = rangeAddress.firstRow;
        firstRow <= existingBodyRange.lastRow;
        firstRow++
      )
        for (
          let firstCol = existingBodyRange.firstCol;
          firstCol <= existingBodyRange.lastCol;
          firstCol++
        ) {
          const cellAddress = Cn(firstRow, firstCol);
          sheetOrThrow.cells[cellAddress] &&
            (nextSheet.cells[cellAddress] = structuredClone(
              sheetOrThrow.cells[cellAddress],
            ));
        }
      existingTable.showTotals &&
        moveTotalsRowCells(workbook, next, sheetId, existingTable, table);
    }
    if (rawRows)
      for (
        let appendRowIndex = 0;
        appendRowIndex < rawRows.length;
        appendRowIndex++
      )
        for (
          let columnIndex = 0;
          columnIndex < existingTable.columns.length;
          columnIndex++
        ) {
          const raw = rawRows[appendRowIndex][columnIndex];
          if (existingTable.columns[columnIndex].calculatedFormula) {
            raw !== null &&
              raw !== "" &&
              throwWorkbookError(
                "Leave calculated-column input slots blank; their saved formulas supply appended values.",
              );
            continue;
          }
          coerceImportedValue(
            raw,
            existingTable.columns[columnIndex].dataType,
          ) ||
            throwWorkbookError(
              `Appended value in ${existingTable.columns[columnIndex].name} does not match its ${existingTable.columns[columnIndex].dataType} type. Supply a typed value or explicitly revise the column type; no rows were appended.`,
            );
          const appendAddress = Cn(
              existingBodyRange.lastRow + 1 + appendRowIndex,
              existingBodyRange.firstCol + columnIndex,
            ),
            cellRecordInput = normalizeCellRecordInput({
              ...nextSheet.cells[appendAddress],
              raw: raw,
              ...(typeof raw == "string" && raw.startsWith("=")
                ? { literal: true }
                : { literal: undefined }),
            });
          nextSheet.cells[appendAddress] = cellRecordInput;
        }
  } else throwWorkbookError("Unknown calculated-table operation.");
  (normalizeTableDefinitions(nextSheet, nextSheet),
    assertWithinCellBudget(next));
  const changes = [];
  for (const sheet of next.sheets) {
    const sheetOrThrow = findSheetOrThrow(workbook, sheet.id);
    sheetOrThrow.protection?.enabled &&
      JSON.stringify(sheetOrThrow) !== JSON.stringify(sheet) &&
      throwWorkbookError(
        "An affected worksheet is protected. Unprotect it before changing dependent table formulas.",
      );
    for (const address of new Set([
      ...Object.keys(sheetOrThrow.cells),
      ...Object.keys(sheet.cells),
    ])) {
      const beforeCell = sheetOrThrow.cells[address],
        afterCell = sheet.cells[address];
      (JSON.stringify(beforeCell) !== JSON.stringify(afterCell) &&
        changes.push({
          sheetId: sheet.id,
          address: address,
          ...(beforeCell ? { before: structuredClone(beforeCell) } : {}),
          ...(afterCell ? { after: structuredClone(afterCell) } : {}),
        }),
        changes.length > tableMutationLimits.changedCells &&
          throwWorkbookError(
            "A reviewed table operation may change at most 500 cells.",
          ));
    }
  }
  checkDataValidationBatch(next, changes);
  const anchorChanges = [];
  for (const [kind, previousItems, nextItems] of [
    ["watch", workbook.formulaWatches, next.formulaWatches],
    ["review", workbook.reviewThreads, next.reviewThreads],
  ])
    for (const nextItem of nextItems ?? []) {
      const previousItem = previousItems?.find(
        (item) => item.id === nextItem.id,
      );
      previousItem &&
        previousItem.anchor.address !== nextItem.anchor.address &&
        anchorChanges.push({
          kind: kind,
          id: nextItem.id,
          before: previousItem.anchor.address,
          after: nextItem.anchor.address,
        });
    }
  return {
    snapshot: JSON.stringify(workbook),
    sheetId: sheetId,
    tableId: tableId,
    request: structuredClone(request),
    beforeRange: existingTable.range,
    afterRange: table.range,
    beforeColumns: structuredClone(existingTable.columns),
    afterColumns: structuredClone(table.columns),
    changes: changes,
    anchorChanges: anchorChanges,
    next: next,
  };
}
function applyTableOperationPreview(workbook, plan, options = {}) {
  plan.snapshot !== JSON.stringify(workbook) &&
    throwWorkbookError(
      "The workbook changed after the table preview. Preview again.",
    );
  const recomputedPlan = applyTableOperation(
    workbook,
    plan.sheetId,
    plan.tableId,
    plan.request,
    options,
  );
  return (
    JSON.stringify(recomputedPlan) !== JSON.stringify(plan) &&
      throwWorkbookError("The table preview was changed. Preview again."),
    JSON.stringify(recomputedPlan.next) === JSON.stringify(workbook)
      ? workbook
      : recomputedPlan.next
  );
}
function normalizeTextCleanups(textCleanups, workbookModel) {
  if (textCleanups === undefined) return;
  (!Array.isArray(textCleanups) ||
    textCleanups.length > 100 ||
    JSON.stringify(textCleanups).length > 5e6) &&
    throwTextCleanupEvidenceError();
  const seenIds = new Set();
  let totalRowCount = 0;
  workbookModel.textCleanups = textCleanups.map((textCleanup) => {
    assertPlainObjectKeys(textCleanup, [
      "id",
      "workbookId",
      "sourceRevision",
      "sourceSheetId",
      "sourceSheetName",
      "outputSheetId",
      "outputSheetName",
      "createdAt",
      "steps",
      "semantics",
      "rows",
    ]);
    for (const idField of [
      "id",
      "workbookId",
      "sourceSheetId",
      "outputSheetId",
    ])
      jm(textCleanup[idField], 100) || throwTextCleanupEvidenceError();
    ((!jm(textCleanup.sourceSheetName, 31) ||
      !jm(textCleanup.outputSheetName, 31) ||
      seenIds.has(textCleanup.id) ||
      !Number.isSafeInteger(textCleanup.sourceRevision) ||
      Number(textCleanup.sourceRevision) < 0 ||
      textCleanup.semantics !== "unicode-v2-en-US-explicit-separators" ||
      !Array.isArray(textCleanup.rows) ||
      !textCleanup.rows.length ||
      textCleanup.rows.length > 500 ||
      (totalRowCount += textCleanup.rows.length) > 1e4 ||
      typeof textCleanup.createdAt != "string" ||
      !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(
        textCleanup.createdAt,
      ) ||
      !Number.isFinite(Date.parse(textCleanup.createdAt)) ||
      new Date(textCleanup.createdAt).toISOString() !==
        textCleanup.createdAt) &&
      throwTextCleanupEvidenceError(),
      seenIds.add(textCleanup.id));
    const steps = validateTextCleanupSteps(textCleanup.steps),
      seenSourceAddresses = new Set(),
      seenOutputAddresses = new Set(),
      rows = textCleanup.rows.map((row) => {
        (assertPlainObjectKeys(row, [
          "sourceAddress",
          "original",
          "result",
          "outputAddress",
          "origin",
        ]),
          (!isCanonicalCellAddress(row.sourceAddress) ||
            !isCanonicalCellAddress(row.outputAddress) ||
            seenSourceAddresses.has(row.sourceAddress) ||
            seenOutputAddresses.has(row.outputAddress) ||
            !WU(row.result)) &&
            throwTextCleanupEvidenceError(),
          seenSourceAddresses.add(row.sourceAddress),
          seenOutputAddresses.add(row.outputAddress));
        const original =
          row.original === undefined
            ? undefined
            : normalizeCellRecordInput(row.original);
        original &&
          typeof original.raw == "string" &&
          original.raw.startsWith("=") &&
          !original.literal &&
          throwTextCleanupEvidenceError();
        let normalizedOrigin;
        if (row.origin !== undefined) {
          assertPlainObjectKeys(row.origin, [
            "id",
            "sourceId",
            "sourceName",
            "sourceSheet",
            "originalRow",
            "kind",
          ]);
          const origin = row.origin;
          ((!jm(origin.id, 100) ||
            !jm(origin.sourceId, 100) ||
            !jm(origin.sourceName, 255) ||
            !jm(origin.sourceSheet, 31) ||
            !Number.isInteger(origin.originalRow) ||
            Number(origin.originalRow) < 1 ||
            Number(origin.originalRow) > 1048576 ||
            !["snapshot", "file"].includes(String(origin.kind))) &&
            throwTextCleanupEvidenceError(),
            (normalizedOrigin = { ...origin }));
        }
        return {
          sourceAddress: row.sourceAddress,
          outputAddress: row.outputAddress,
          result: row.result,
          ...(original ? { original: original } : {}),
          ...(normalizedOrigin ? { origin: normalizedOrigin } : {}),
        };
      });
    return { ...textCleanup, steps: steps, rows: rows };
  });
}
function validateFieldTransformEvidence(fieldTransforms, workbookModel) {
  if (fieldTransforms === undefined) return;
  if (
    !Array.isArray(fieldTransforms) ||
    fieldTransforms.length > 100 ||
    JSON.stringify(fieldTransforms).length >
      textInferenceLimits.evidenceCharacters
  )
    throw Error(
      "Field transformation evidence exceeds its native history budget. Keep the original backup.",
    );
  const seenIds = new Set();
  let totalOutputCount = 0;
  workbookModel.fieldTransforms = fieldTransforms.map((fieldTransform) => {
    const parsed = analysisOutputEvidenceSchema.safeParse(fieldTransform);
    if (!parsed.success)
      throw Error(
        `Invalid field transformation evidence: ${parsed.error.issues[0]?.message}`,
      );
    const data = parsed.data,
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
    if (seenIds.has(data.id) || (totalOutputCount += data.outputs.length) > 2e4)
      throw Error("Duplicate or over-capacity field transformation evidence.");
    if (
      (seenIds.add(data.id),
      data.sourceSheetId !== request.sourceSheetId ||
        data.outputSheetName !== request.outputName)
    )
      throw Error("Field transformation evidence does not match its recipe.");
    const seenSourceAddresses = new Set(),
      sources = data.sources.map((source) => {
        const sourceCell = _t(source.address);
        if (
          sourceCell.row < rangeAddress.firstRow ||
          sourceCell.row > rangeAddress.lastRow ||
          sourceCell.col < rangeAddress.firstCol ||
          sourceCell.col > rangeAddress.lastCol
        )
          throw Error(
            "An original field source is outside the recorded range.",
          );
        if (seenSourceAddresses.has(source.address))
          throw Error("Duplicate original field source.");
        seenSourceAddresses.add(source.address);
        const original =
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
      seenOutputAddresses = new Set();
    for (const output of data.outputs) {
      if (seenOutputAddresses.has(output.address))
        throw Error("Duplicate transformed output address.");
      if (
        (seenOutputAddresses.add(output.address),
        ((output.type === "number" || output.type === "date") &&
          typeof output.value != "number") ||
          (output.type === "boolean" && typeof output.value != "boolean") ||
          (output.type === "text" && typeof output.value != "string") ||
          (output.type === "blank" && output.value !== null))
      )
        throw Error("Transformed output type does not match its value.");
    }
    for (const entry of [...data.outputs, ...data.exceptions])
      if (
        entry.row < rangeAddress.firstRow + 1 + Number(request.header) ||
        entry.row > rangeAddress.lastRow + 1 ||
        entry.sources.some((source) => !seenSourceAddresses.has(source)) ||
        new Set(entry.sources).size !== entry.sources.length
      )
        throw Error("Unknown or repeated original field source locator.");
    return { ...data, request: request, sources: sources };
  });
}
function normalizeWorkbookModel(workbook, canReuseFrozen) {
  if (canReuseFrozen && isDeepFrozenValue(workbook) && NC.has(workbook))
    return workbook;
  if (!isNonArrayObject(workbook) || workbook.schemaVersion !== 1)
    return throwWithMessage(
      "Unsupported workbook schema. Expected schemaVersion 1.",
    );
  const workbookId = requireBoundedText(workbook.id, "Workbook ID", 100),
    workbookName = requireBoundedText(workbook.name, "Workbook name", 200);
  if (!workbookId || !workbookName.trim())
    return throwWithMessage("Workbook ID and name are required.");
  const revision = P2(
    workbook.revision,
    0,
    Number.MAX_SAFE_INTEGER - 1,
    "Revision",
  );
  if (!Array.isArray(workbook.sheets) || !workbook.sheets.length)
    return throwWithMessage("A workbook must contain at least one sheet.");
  let populatedCellCount = 0;
  const seenSheetIds = new Set(),
    seenSheetNames = new Set(),
    sheets = workbook.sheets.map((sheet) => {
      if (!isNonArrayObject(sheet)) return throwWithMessage("Invalid sheet.");
      const sheetId = requireBoundedText(sheet.id, "Sheet ID", 100),
        sheetName = requireBoundedText(sheet.name, "Sheet name", 31);
      if (
        !sheetId ||
        !sheetName.trim() ||
        /[\\/\[\]*?:]/.test(sheetName) ||
        seenSheetIds.has(sheetId) ||
        seenSheetNames.has(sheetName.toLowerCase())
      )
        return throwWithMessage(
          "Sheet IDs and names must be nonempty and unique; sheet names cannot contain Excel-reserved characters.",
        );
      (seenSheetIds.add(sheetId), seenSheetNames.add(sheetName.toLowerCase()));
      const rows = P2(sheet.rowCount, 1, workbookLimits.maxRows, "Row count"),
        cols = P2(sheet.colCount, 1, workbookLimits.maxCols, "Column count");
      if (!isNonArrayObject(sheet.cells))
        return throwWithMessage("Sheet cells must be an address map.");
      const cachedCellCount =
        canReuseFrozen && isDeepFrozenValue(sheet) ? YE.get(sheet) : undefined;
      if (cachedCellCount !== undefined) {
        if (
          ((populatedCellCount += cachedCellCount),
          populatedCellCount > workbookLimits.maxCells)
        )
          return throwWithMessage(
            "Workbook exceeds the 100,000 populated cell limit.",
          );
        const reusedSheet = { ...sheet };
        return (YE.set(reusedSheet, cachedCellCount), reusedSheet);
      }
      const cells = {};
      for (const [cellAddress, cell] of Object.entries(sheet.cells)) {
        const cellLocator = _t(cellAddress);
        if (
          cellAddress !==
            `${columnIndexToLetters(cellLocator.col)}${cellLocator.row + 1}` ||
          cellLocator.row >= rows ||
          cellLocator.col >= cols
        )
          return throwWithMessage(
            `Invalid or out-of-bounds cell: ${cellAddress}`,
          );
        if (++populatedCellCount > workbookLimits.maxCells)
          return throwWithMessage(
            "Workbook exceeds the 100,000 populated cell limit.",
          );
        cells[cellAddress] = normalizeCellRecord(cell, canReuseFrozen);
      }
      const normalizedSheet = {
        id: sheetId,
        name: sheetName,
        rowCount: workbookLimits.maxRows,
        colCount: workbookLimits.maxCols,
        cells: cells,
      };
      if (sheet.columnWidths !== undefined) {
        if (!isNonArrayObject(sheet.columnWidths))
          return throwWithMessage("Column widths must be an index map.");
        normalizedSheet.columnWidths = {};
        for (const [columnKey, columnWidth] of Object.entries(
          sheet.columnWidths,
        )) {
          if (
            !/^(0|[1-9]\d*)$/.test(columnKey) ||
            Number(columnKey) >= cols ||
            typeof columnWidth != "number" ||
            !Number.isFinite(columnWidth) ||
            columnWidth < minColumnWidth ||
            columnWidth > maxColumnWidth
          )
            return throwWithMessage(
              "Column widths must be 40–600 pixels for an existing column.",
            );
          normalizedSheet.columnWidths[columnKey] = columnWidth;
        }
      }
      if (sheet.rowHeights !== undefined) {
        if (
          !isNonArrayObject(sheet.rowHeights) ||
          (Object.getPrototypeOf(sheet.rowHeights) !== Object.prototype &&
            Object.getPrototypeOf(sheet.rowHeights) !== null)
        )
          return throwWithMessage("Row heights must be a plain index map.");
        normalizedSheet.rowHeights = {};
        for (const [rowKey, rowHeight] of Object.entries(sheet.rowHeights)) {
          if (
            !/^(0|[1-9]\d*)$/.test(rowKey) ||
            Number(rowKey) >= rows ||
            typeof rowHeight != "number" ||
            !Number.isFinite(rowHeight) ||
            rowHeight < minRowHeight ||
            rowHeight > maxRowHeight
          )
            return throwWithMessage(
              `Row heights must be ${minRowHeight}–${maxRowHeight} pixels for an existing row.`,
            );
          normalizedSheet.rowHeights[rowKey] = rowHeight;
        }
      }
      if (sheet.hidden !== undefined) {
        if (typeof sheet.hidden != "boolean")
          return throwWithMessage("Sheet hidden flag must be boolean.");
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
    return throwWithMessage("At least one sheet must be visible.");
  if (!Array.isArray(workbook.events) || workbook.events.length > 2e3)
    return throwWithMessage(
      "Workbook history must be an array of at most 2,000 events.",
    );
  const seenEventIds = new Set();
  let totalChangeCount = 0;
  const events = workbook.events.map((historyEvent) => {
      if (!isNonArrayObject(historyEvent))
        return throwWithMessage("Invalid history event.");
      const eventId = requireBoundedText(historyEvent.id, "Event ID", 100),
        timestamp = requireBoundedText(
          historyEvent.timestamp,
          "Event timestamp",
          40,
        );
      return !eventId ||
        seenEventIds.has(eventId) ||
        !Number.isFinite(Date.parse(timestamp))
        ? throwWithMessage(
            "History events require unique IDs and valid timestamps.",
          )
        : (seenEventIds.add(eventId),
          !Array.isArray(historyEvent.changes) ||
          historyEvent.changes.length > workbookLimits.maxPatches ||
          (totalChangeCount += historyEvent.changes.length) >
            workbookLimits.maxCells
            ? throwWithMessage("Invalid history change count.")
            : {
                id: eventId,
                timestamp: timestamp,
                actor: requireBoundedText(historyEvent.actor, "Actor", 100),
                action: requireBoundedText(historyEvent.action, "Action", 200),
                changes: historyEvent.changes.map((change) => {
                  if (
                    !isNonArrayObject(change) ||
                    typeof change.sheetId != "string" ||
                    !change.sheetId.length ||
                    change.sheetId.length > 100 ||
                    typeof change.address != "string"
                  )
                    return throwWithMessage("Invalid history cell reference.");
                  const changeLocator = _t(change.address);
                  return change.address !==
                    `${columnIndexToLetters(changeLocator.col)}${changeLocator.row + 1}`
                    ? throwWithMessage("History cell address is not canonical.")
                    : change.before === undefined && change.after === undefined
                      ? throwWithMessage(
                          "History change must have a before or after value.",
                        )
                      : {
                          sheetId: change.sheetId,
                          address: change.address,
                          ...(change.before !== undefined
                            ? {
                                before: normalizeCellRecord(
                                  change.before,
                                  canReuseFrozen,
                                ),
                              }
                            : {}),
                          ...(change.after !== undefined
                            ? {
                                after: normalizeCellRecord(
                                  change.after,
                                  canReuseFrozen,
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
    buildSavedReportEnvelope(workbook, normalizedWorkbook),
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
    reconcileWorkbookStyleBindings(workbook.styleCatalog, normalizedWorkbook),
    workbook.formulaWatches !== undefined &&
      (normalizedWorkbook.formulaWatches = validateFormulaWatches(
        workbook.formulaWatches,
        normalizedWorkbook,
      )),
    zie(workbook, normalizedWorkbook),
    copyExtensionCacheEntries(workbook, normalizedWorkbook),
    normalizedWorkbook
  );
}
function normalizeWorkbookInput(workbook) {
  return normalizeWorkbookModel(workbook, false);
}
function assertThemePreviewCurrent(workbook, plan, options = {}) {
  (options.readOnly && failWithMessage("This workbook is read-only."),
    (plan.version !== 1 || plan.snapshot !== JSON.stringify(workbook)) &&
      failWithMessage(
        "The workbook changed after this preview. Close and reopen Themes & Cell Styles.",
      ));
  const recomputedPlan = applyNamedStyleChange(workbook, plan.request, options);
  if (
    (canonicalStyleKey(recomputedPlan) !== canonicalStyleKey(plan) &&
      failWithMessage(
        "This theme preview changed. Create a fresh preview before applying it.",
      ),
    !plan.count &&
      canonicalStyleKey(plan.catalogBefore) ===
        canonicalStyleKey(plan.catalogAfter))
  )
    return workbook;
  const nextWorkbook = cloneStyleValue(workbook);
  plan.catalogAfter
    ? (nextWorkbook.styleCatalog = cloneStyleValue(plan.catalogAfter))
    : delete nextWorkbook.styleCatalog;
  for (const change of plan.changes)
    findSheetOrThrow(nextWorkbook, change.sheetId).cells[change.address] =
      cloneStyleValue(change.after);
  return normalizeWorkbookInput(nextWorkbook);
}
function Mw(workbook) {
  const workbookInput = normalizeWorkbookInput(workbook);
  return (
    validateWorkbookMetadata(workbook.metadata, workbookInput),
    workbookInput
  );
}
function updateWorkbookProperties(workbook, propertyUpdates) {
  if (
    !isPlainObjectValue(propertyUpdates) ||
    Object.keys(propertyUpdates).some(
      (fieldName) => !["name", "engagement", "period"].includes(fieldName),
    )
  )
    throw new Error("Invalid workbook property fields.");
  const nextWorkbook = Mw(workbook);
  if (
    (propertyUpdates.name !== undefined &&
      (nextWorkbook.name = trimRequiredMetadataText(
        propertyUpdates.name,
        "Workbook name",
        200,
      )),
    propertyUpdates.engagement !== undefined ||
      propertyUpdates.period !== undefined)
  ) {
    if (
      typeof propertyUpdates.engagement == "string" &&
      propertyUpdates.engagement.length > 200
    )
      throw new Error("Engagement must contain at most 200 characters.");
    if (
      typeof propertyUpdates.period == "string" &&
      propertyUpdates.period.length > 100
    )
      throw new Error("Period must contain at most 100 characters.");
    const metadata = { ...(nextWorkbook.metadata ?? createWorkbookMetadata()) };
    (propertyUpdates.engagement !== undefined &&
      (typeof propertyUpdates.engagement == "string" &&
      !propertyUpdates.engagement.trim()
        ? delete metadata.engagement
        : (metadata.engagement = trimRequiredMetadataText(
            propertyUpdates.engagement,
            "Engagement",
            200,
          ))),
      propertyUpdates.period !== undefined &&
        (typeof propertyUpdates.period == "string" &&
        !propertyUpdates.period.trim()
          ? delete metadata.period
          : (metadata.period = trimRequiredMetadataText(
              propertyUpdates.period,
              "Period",
              100,
            ))),
      validateWorkbookMetadata(metadata, nextWorkbook));
  }
  return nextWorkbook;
}
function renameWorkbook(activeWorkbook, newName) {
  return updateWorkbookProperties(activeWorkbook, { name: newName });
}
function cloneWorkbookWithNewIds(sourceWorkbook) {
  const workbook = structuredClone(Mw(sourceWorkbook)),
    newSheetIds = new Map(
      workbook.sheets.map((sheet) => [sheet.id, crypto.randomUUID()]),
    ),
    newTableIds = new Map(
      workbook.sheets.flatMap((sheet) =>
        (sheet.tables ?? []).map((tableRecord) => [
          tableRecord.id,
          crypto.randomUUID(),
        ]),
      ),
    );
  ((workbook.id = crypto.randomUUID()),
    (workbook.revision = 0),
    (workbook.events = []));
  for (const sheet of workbook.sheets)
    if (sheet.analysisReport) {
      const request = sheet.analysisReport.request;
      if (request.kind === "statistics")
        request.sheetId = newSheetIds.get(request.sheetId) ?? request.sheetId;
      else
        for (const analysisReference of [
          request.objective,
          ...request.variables.map((variable) => variable.cell),
          ...request.constraints.flatMap((constraint) => [
            constraint.left,
            ...(typeof constraint.right == "number" ? [] : [constraint.right]),
          ]),
        ])
          analysisReference.sheetId =
            newSheetIds.get(analysisReference.sheetId) ??
            analysisReference.sheetId;
      sheet.analysisReport.definitionHash = fingerprintCanonicalValue(request);
    }
  const dataModelDefinitions = [
    workbook.dataModel,
    ...workbook.sheets.map((sheet) => sheet.dataModelReport?.definition),
  ].filter(Boolean);
  for (const dataModelDefinition of dataModelDefinitions)
    for (const table of dataModelDefinition.tables)
      ((table.sheetId = newSheetIds.get(table.sheetId) ?? table.sheetId),
        (table.tableId &&= newTableIds.get(table.tableId) ?? table.tableId));
  for (const sheet of workbook.sheets)
    sheet.dataModelReport &&
      (sheet.dataModelReport.definitionHash = fingerprintCanonicalValue(
        sheet.dataModelReport.definition,
      ));
  for (const filterView of workbook.savedFilterViews ?? [])
    ((filterView.id = crypto.randomUUID()),
      (filterView.sheetId =
        newSheetIds.get(filterView.sheetId) ?? filterView.sheetId));
  for (const transformationRecipe of workbook.transformationRecipes ?? []) {
    transformationRecipe.id = crypto.randomUUID();
    for (const version of transformationRecipe.versions) {
      for (const tableSource of [
        version.definition.source,
        ...version.definition.steps.flatMap((step) =>
          step.kind === "append" || step.kind === "merge" ? [step.source] : [],
        ),
      ])
        ((tableSource.sheetId =
          newSheetIds.get(tableSource.sheetId) ?? tableSource.sheetId),
          (tableSource.tableId &&=
            newTableIds.get(tableSource.tableId) ?? tableSource.tableId));
      version.definitionHash = fingerprintCanonicalValue(version.definition);
    }
  }
  for (const sheet of workbook.sheets) {
    ((sheet.id = newSheetIds.get(sheet.id)),
      (sheet.outlineGroups &&= regenerateOutlineGroupIds(sheet)));
    for (const sheetTable of sheet.tables ?? []) {
      sheetTable.id = newTableIds.get(sheetTable.id);
      for (const column of sheetTable.columns) column.id = crypto.randomUUID();
    }
    for (const pivotOrQuery of [
      ...(sheet.pivots ?? []),
      ...(sheet.queries ?? []),
    ])
      ((pivotOrQuery.id = crypto.randomUUID()),
        (pivotOrQuery.source.sheetId =
          newSheetIds.get(pivotOrQuery.source.sheetId) ??
          pivotOrQuery.source.sheetId),
        (pivotOrQuery.source.tableId &&=
          newTableIds.get(pivotOrQuery.source.tableId) ??
          pivotOrQuery.source.tableId),
        (pivotOrQuery.sourceRevision = 0));
    for (const objectOrRule of [
      ...(sheet.objects ?? []),
      ...(sheet.validations ?? []),
      ...(sheet.conditionalFormats ?? []),
    ])
      objectOrRule.id = crypto.randomUUID();
  }
  for (const definedName of workbook.definedNames ?? [])
    definedName.scopeSheetId &&=
      newSheetIds.get(definedName.scopeSheetId) ?? definedName.scopeSheetId;
  return Mw(
    createReviewThread(
      createFormulaWatch(
        renameStyleNamespace(workbook, crypto.randomUUID()),
        newSheetIds,
      ),
      newSheetIds,
    ),
  );
}
function createWorkbookFromTemplate(workbook, workbookDetails) {
  const workbookName = trimRequiredMetadataText(
      workbookDetails.name,
      "Workbook name",
      200,
    ),
    engagement = trimRequiredMetadataText(
      workbookDetails.engagement,
      "Engagement",
      200,
    ),
    period = trimRequiredMetadataText(workbookDetails.period, "Period", 100),
    templateWorkbook = Mw(workbook),
    clonedWorkbook = cloneWorkbookWithNewIds(templateWorkbook);
  return (
    (clonedWorkbook.name = workbookName),
    (clonedWorkbook.metadata = createWorkbookMetadata({
      engagement: engagement,
      period: period,
      template: {
        id: templateWorkbook.id,
        name: templateWorkbook.name,
        version: templateWorkbook.metadata?.template?.version ?? 1,
        sourceWorkbookId: templateWorkbook.id,
        sourceRevision: templateWorkbook.revision,
      },
    })),
    Mw(clonedWorkbook)
  );
}
function createBlankWorkbook(workbookName = "Book1") {
  return normalizeWorkbookInput({
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
function assertConditionalFormatPreviewCurrent(
  workbook,
  preview,
  options = {},
) {
  if (
    preview.workbookId !== workbook.id ||
    preview.revision !== workbook.revision ||
    preview.snapshotGuard !== JSON.stringify(workbook)
  )
    throw Error(
      "Conditional format preview is stale. Reopen the manager to refresh it.",
    );
  const updateResult = updateConditionalFormatRules(
    workbook,
    preview.sheetId,
    preview.rules,
    options,
  );
  if (JSON.stringify(updateResult.changes) !== JSON.stringify(preview.changes))
    throw Error(
      "Conditional format preview was modified. Preview the rules again.",
    );
  return updateResult.changes.length
    ? normalizeWorkbookInput({
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          sheet.id === preview.sheetId
            ? { ...sheet, conditionalFormats: updateResult.rules }
            : sheet,
        ),
      })
    : workbook;
}
function w0(workbook) {
  const workbookInput = normalizeWorkbookInput(aQ(workbook));
  return (
    workbookInput.sheets.forEach((sheet, sheetIndex) =>
      validateWorksheetVisibility(workbook.sheets[sheetIndex], sheet),
    ),
    validateWorkbookStructureProtection(
      workbook.structureProtection,
      workbookInput,
    ),
    workbookInput
  );
}
function addWorksheet(activeWorkbook, worksheetName) {
  return (
    assertWorkbookStructureEditable(activeWorkbook),
    w0({
      ...activeWorkbook,
      sheets: [
        ...activeWorkbook.sheets,
        createWorksheetRecord(
          assertUniqueWorksheetName(activeWorkbook, worksheetName),
        ),
      ],
    })
  );
}
function renameWorksheet(activeWorkbook, sheetId, newName) {
  assertWorkbookStructureEditable(activeWorkbook);
  const worksheetById = findWorksheetById(activeWorkbook, sheetId),
    uniqueName = assertUniqueWorksheetName(activeWorkbook, newName, sheetId);
  if (worksheetById.name === uniqueName) return activeWorkbook;
  const workbook = structuredClone(activeWorkbook);
  for (const sheet of workbook.sheets)
    for (const cell of Object.values(sheet.cells))
      !cell.literal &&
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        (cell.raw = rewriteStructuredTableReference(
          cell.raw,
          activeWorkbook,
          worksheetById,
          uniqueName,
        ));
  for (const sheet of workbook.sheets)
    for (const ruleEntry of [
      ...(sheet.conditionalFormats ?? []),
      ...(sheet.validations ?? []),
    ])
      if (
        (ruleEntry.kind === "formula" || ruleEntry.kind === "custom") &&
        ruleEntry.formula
      ) {
        const rewrittenFormula = rewriteStructuredTableReference(
          ruleEntry.formula,
          activeWorkbook,
          worksheetById,
          uniqueName,
        );
        (rewrittenFormula !== ruleEntry.formula &&
          assertSheetUnprotected(sheet),
          (ruleEntry.formula = rewrittenFormula));
      }
  for (const definedName of workbook.definedNames ?? [])
    definedName.refersTo = rewriteStructuredTableReference(
      definedName.refersTo,
      activeWorkbook,
      worksheetById,
      uniqueName,
    );
  for (const sheet of workbook.sheets)
    for (const sheetTable of sheet.tables ?? [])
      for (const column of sheetTable.columns)
        if (column.calculatedFormula) {
          const rewrittenCalculatedFormula = rewriteStructuredTableReference(
            column.calculatedFormula,
            activeWorkbook,
            worksheetById,
            uniqueName,
          );
          (rewrittenCalculatedFormula !== column.calculatedFormula &&
            assertSheetUnprotected(sheet),
            (column.calculatedFormula = rewrittenCalculatedFormula));
        }
  return (
    (findWorksheetById(workbook, sheetId).name = uniqueName),
    w0(workbook)
  );
}
function moveWorksheet(workbook, sheetId, targetIndex) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, sheetId),
    currentIndex = workbook.sheets.indexOf(worksheetById);
  if (
    ((!Number.isInteger(targetIndex) ||
      targetIndex < 0 ||
      targetIndex >= workbook.sheets.length) &&
      throwWorkbookError("Choose a worksheet position inside the workbook."),
    currentIndex === targetIndex)
  )
    return workbook;
  visitTableFormulasOutsideSheet(workbook, (formula) => {
    parseWorksheetQualifiedReference(formula).some(
      (workbook) => workbook.sheets,
    ) &&
      throwWorkbookError(
        "Reordering worksheets with three-dimensional references requires explicit references first.",
      );
  });
  const sheets = [...workbook.sheets];
  return (
    sheets.splice(currentIndex, 1),
    sheets.splice(targetIndex, 0, worksheetById),
    w0({ ...workbook, sheets: sheets })
  );
}
function setWorksheetVisibility(workbook, sheetId, visibility) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, sheetId);
  (["visible", "hidden", "veryHidden"].includes(visibility) ||
    throwWorkbookError("Choose visible, hidden or very hidden."),
    visibility !== "visible" &&
      isVisibleWorksheet(worksheetById) &&
      workbook.sheets.filter(isVisibleWorksheet).length === 1 &&
      throwWorkbookError("At least one worksheet must remain visible."));
  const currentVisibility = worksheetById.veryHidden
    ? "veryHidden"
    : worksheetById.hidden
      ? "hidden"
      : "visible";
  return visibility === currentVisibility
    ? workbook
    : w0({
        ...workbook,
        sheets: workbook.sheets.map((sheet) =>
          sheet.id === sheetId
            ? {
                ...sheet,
                hidden: visibility !== "visible",
                veryHidden: visibility === "veryHidden",
              }
            : sheet,
        ),
      });
}
function setWorksheetTabColor(workbook, sheetId, tabColor) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, sheetId),
    validatedTabColor =
      tabColor === undefined ? undefined : validateTabColor(tabColor);
  return worksheetById.tabColor === validatedTabColor
    ? workbook
    : w0({
        ...workbook,
        sheets: workbook.sheets.map((sheet) => {
          if (sheet.id !== sheetId) return sheet;
          const updatedSheet = { ...sheet };
          return (
            validatedTabColor
              ? (updatedSheet.tabColor = validatedTabColor)
              : delete updatedSheet.tabColor,
            updatedSheet
          );
        }),
      });
}
function duplicateWorksheet(workbook, sheetId, t) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, sheetId),
    copyName = assertUniqueWorksheetName(
      workbook,
      createUniqueWorksheetName(workbook, worksheetById.name),
    );
  (visitTableFormulasOutsideSheet(workbook, (formula) => {
    parseWorksheetQualifiedReference(formula).some(
      (workbook) => workbook.sheets,
    ) &&
      throwWorkbookError(
        "Duplicating worksheets with three-dimensional references requires explicit references first.",
      );
  }),
    workbook.sheets.reduce(
      (total, workbookSheet) => total + Object.keys(workbookSheet.cells).length,
      0,
    ) +
      Object.keys(worksheetById.cells).length >
      workbookLimits.maxCells &&
      throwWorkbookError(
        "Duplicating this worksheet would exceed the 100,000 populated-cell budget.",
      ));
  const sheet = structuredClone(worksheetById);
  ((sheet.id = createUuid()),
    (sheet.outlineGroups &&= regenerateOutlineGroupIds(sheet)),
    (sheet.name = copyName));
  const takenNames = new Set([
      ...workbook.sheets.flatMap((sheet) =>
        (sheet.tables ?? []).map((existingTable) =>
          existingTable.name.toLowerCase(),
        ),
      ),
      ...(workbook.definedNames ?? []).map((existingDefinedName) =>
        existingDefinedName.name.toLowerCase(),
      ),
    ]),
    tableIdMap = new Map(),
    tableNameMap = new Map();
  for (const copiedTable of sheet.tables ?? []) {
    const originalTableId = copiedTable.id,
      originalTableName = copiedTable.name;
    ((copiedTable.id = createUuid()),
      (copiedTable.columns = copiedTable.columns.map((column) => ({
        ...column,
        id: createUuid(),
      }))));
    for (let copyNumber = 2; ; copyNumber++) {
      const copySuffix = `_Copy${copyNumber}`,
        candidateName =
          originalTableName.slice(0, 255 - copySuffix.length) + copySuffix;
      if (!takenNames.has(candidateName.toLowerCase())) {
        ((copiedTable.name = candidateName),
          takenNames.add(candidateName.toLowerCase()));
        break;
      }
    }
    (tableIdMap.set(originalTableId, copiedTable.id),
      tableNameMap.set(originalTableName, copiedTable.name));
  }
  const rewriteFormula = (formulaText) => {
    let rewrittenText = rewriteStructuredTableReference(
      formulaText,
      workbook,
      worksheetById,
      copyName,
    );
    for (const [originalName, renamedName] of tableNameMap)
      rewrittenText = $v(rewrittenText, originalName, renamedName);
    return rewrittenText;
  };
  for (const sheetTable of sheet.tables ?? [])
    for (const column of sheetTable.columns)
      column.calculatedFormula &&= rewriteFormula(column.calculatedFormula);
  for (const cell of Object.values(sheet.cells))
    !cell.literal &&
      typeof cell.raw == "string" &&
      cell.raw.startsWith("=") &&
      (cell.raw = rewriteFormula(cell.raw));
  (sheet.objects?.forEach((object) => {
    object.id = createUuid();
  }),
    sheet.validations?.forEach((validation) => {
      ((validation.id = createUuid()),
        validation.kind === "custom" &&
          validation.formula &&
          (validation.formula = rewriteFormula(validation.formula)));
    }),
    sheet.conditionalFormats?.forEach((conditionalFormat) => {
      ((conditionalFormat.id = createUuid()),
        conditionalFormat.kind === "formula" &&
          conditionalFormat.formula &&
          (conditionalFormat.formula = rewriteFormula(
            conditionalFormat.formula,
          )));
    }));
  for (const pivotOrQuery of [
    ...(sheet.pivots ?? []),
    ...(sheet.queries ?? []),
  ])
    ((pivotOrQuery.id = createUuid()),
      (pivotOrQuery.sourceRevision = workbook.revision),
      pivotOrQuery.source.sheetId === worksheetById.id &&
        ((pivotOrQuery.source.sheetId = sheet.id),
        (pivotOrQuery.source.tableId &&=
          tableIdMap.get(pivotOrQuery.source.tableId) ??
          throwWorkbookError(
            "A copied analysis refers to a missing source table.",
          ))));
  const copiedDefinedNames = (workbook.definedNames ?? [])
      .filter((definedName) => definedName.scopeSheetId === worksheetById.id)
      .map((sourceDefinedName) => ({
        ...structuredClone(sourceDefinedName),
        scopeSheetId: sheet.id,
        refersTo: rewriteFormula(sourceDefinedName.refersTo),
      })),
    sheets = [...workbook.sheets];
  return (
    sheets.splice(workbook.sheets.indexOf(worksheetById) + 1, 0, sheet),
    w0(
      collectHyperlinkEntries(
        {
          ...workbook,
          sheets: sheets,
          ...(copiedDefinedNames.length
            ? {
                definedNames: [
                  ...(workbook.definedNames ?? []),
                  ...copiedDefinedNames,
                ],
              }
            : {}),
        },
        worksheetById.id,
        sheet.id,
      ),
    )
  );
}
function deleteWorksheet(workbook, sheetId, isDeletionConfirmed = false) {
  assertWorkbookStructureEditable(workbook);
  const worksheetById = findWorksheetById(workbook, sheetId);
  (workbook.sheets.length === 1 &&
    throwWorkbookError("The final worksheet cannot be deleted."),
    isVisibleWorksheet(worksheetById) &&
      workbook.sheets.filter(isVisibleWorksheet).length === 1 &&
      throwWorkbookError("At least one worksheet must remain visible."));
  const deletionImpacts = describeSheetDeletionImpacts(workbook, sheetId);
  (deletionImpacts.some((impact) => impact.kind === "unsupported-reference") &&
    throwWorkbookError(
      "Deletion is blocked by ambiguous or unsupported references. Replace dynamic or three-dimensional references with explicit references first.",
    ),
    deletionImpacts.length &&
      !isDeletionConfirmed &&
      throwWorkbookError(
        "This worksheet has dependencies. Review them and explicitly confirm deletion.",
      ));
  const remainingWorkbook = structuredClone(workbook);
  remainingWorkbook.sheets = remainingWorkbook.sheets.filter(
    (sheet) => sheet.id !== sheetId,
  );
  for (const sheet of remainingWorkbook.sheets) {
    for (const sheetTable of sheet.tables ?? [])
      for (const column of sheetTable.columns)
        if (column.calculatedFormula) {
          const rewrittenCalculatedFormula = rewriteStructuredTableReference(
            column.calculatedFormula,
            workbook,
            worksheetById,
            undefined,
            true,
          );
          (rewrittenCalculatedFormula !== column.calculatedFormula &&
            assertSheetUnprotected(sheet),
            (column.calculatedFormula = rewrittenCalculatedFormula));
        }
    for (const cell of Object.values(sheet.cells))
      !cell.literal &&
        typeof cell.raw == "string" &&
        cell.raw.startsWith("=") &&
        (cell.raw = rewriteStructuredTableReference(
          cell.raw,
          workbook,
          worksheetById,
          undefined,
          true,
        ));
    for (const ruleEntry of [
      ...(sheet.conditionalFormats ?? []),
      ...(sheet.validations ?? []),
    ])
      if (
        (ruleEntry.kind === "formula" || ruleEntry.kind === "custom") &&
        ruleEntry.formula
      ) {
        const rewrittenFormula = rewriteStructuredTableReference(
          ruleEntry.formula,
          workbook,
          worksheetById,
          undefined,
          true,
        );
        (rewrittenFormula !== ruleEntry.formula &&
          assertSheetUnprotected(sheet),
          (ruleEntry.formula = rewrittenFormula));
      }
    ((sheet.pivots &&= sheet.pivots.filter(
      (pivot) => pivot.source.sheetId !== sheetId,
    )),
      (sheet.queries &&= sheet.queries.filter(
        (query) => query.source.sheetId !== sheetId,
      )));
  }
  return (
    (remainingWorkbook.definedNames &&= remainingWorkbook.definedNames
      .filter((definedName) => definedName.scopeSheetId !== sheetId)
      .map((retainedDefinedName) => ({
        ...retainedDefinedName,
        refersTo: rewriteStructuredTableReference(
          retainedDefinedName.refersTo,
          workbook,
          worksheetById,
          undefined,
          true,
        ),
      }))),
    w0(fQ(remainingWorkbook, sheetId))
  );
}
async function protectWorkbookStructure(workbook, password) {
  (assertWorkbookStructureEditable(workbook),
    validateWorkbookPassword(password));
  const uuid = createUuid(),
    structureProtection = password
      ? {
          enabled: true,
          salt: uuid,
          passwordHash: await hashWorkbookStructurePassword(password, uuid),
        }
      : { enabled: true };
  return w0({ ...workbook, structureProtection: structureProtection });
}
async function unprotectWorkbookStructure(workbook, password) {
  validateWorkbookPassword(password);
  const structureProtection = workbook.structureProtection;
  if (!structureProtection) return workbook;
  (validateWorkbookStructureProtection(structureProtection, { ...workbook }),
    structureProtection.passwordHash &&
      (await hashWorkbookStructurePassword(
        password,
        structureProtection.salt,
      )) !== structureProtection.passwordHash.toLowerCase() &&
      throwWorkbookError("The workbook structure password does not match."));
  const unprotectedWorkbook = { ...workbook };
  return (
    delete unprotectedWorkbook.structureProtection,
    w0(unprotectedWorkbook)
  );
}
function t$(workbookModel) {
  deepFreezeValue(workbookModel);
  for (const sheet of workbookModel.sheets) {
    if (YE.has(sheet)) continue;
    const cells = Object.values(sheet.cells);
    for (const cell of cells) JU.add(cell);
    (QU.set(sheet.cells, cells.length),
      cells.some(
        (cellRecord) => cellRecord.hyperlink || cellRecord.style?.binding,
      ) || YE.set(sheet, cells.length));
  }
  return (NC.add(workbookModel), workbookModel);
}

export {
  assertConditionalFormatPreviewCurrent,
  addWorksheet,
  setWorksheetTabColor,
  applyTableOperationPreview,
  normalizeTextCleanups,
  unprotectWorkbookStructure,
  updateWorkbookProperties,
  createWorkbookFromTemplate,
  createWorksheetRecord,
  applyPivotViewChange,
  renameWorkbook,
  renameWorksheet,
  applyTableOperation,
  previewColumnTypeChange,
  setWorksheetVisibility,
  deleteWorksheet,
  protectWorkbookStructure,
  validateFieldTransformEvidence,
  runQueryDefinition,
  applyColumnTypePreview,
  extractPivotDrillThrough,
  normalizeWorkbookModel,
  assertThemePreviewCurrent,
  createWorksheetTable,
  normalizeCellRecordInput,
  refreshPivotTable,
  paintCellFormatting,
  normalizeWorkbookInput,
  t$,
  createBlankWorkbook,
  moveWorksheet,
  buildPivotTableOutput,
  refreshQueryDefinition,
  duplicateWorksheet,
};

/* Also used by the workers. */
export { normalizeCellRecord, validateTableSourceEvidence };
